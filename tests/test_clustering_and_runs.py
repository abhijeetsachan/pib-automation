import contextlib
import io
import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import Mock, patch

import main as pib
import dprcg_automation as cg
import send_email
from pib_scraper import PIBScraper
from run_status import (FetchError, ParseError, NoReleasesError, IncompleteRunError,
                        RunStatus, completed_files)
from topic_clustering import safe_clusters


TARGET = date(2026, 11, 16)


def release(rid, title, body="Unique facts", **extra):
    return {"id": rid, "headline": title, "raw_title": title,
            "date": TARGET.isoformat(), "body_text": body,
            "url": f"https://dprcg.gov.in/post/{rid}/news", **extra}


def calendar(day=16):
    return ''.join(f'<select id="ctl00_ContentPlaceHolder1_{name}">'
                   f'<option value="{value}" selected="selected">{value}</option></select>'
                   for name, value in (("ddlday", day), ("ddlMonth", 11), ("ddlYear", 2026)))


class ClusteringTests(unittest.TestCase):
    def test_community_and_location_do_not_imply_scheme(self):
        stories = [release("1", "Pahadi Korwa health camp at Dharamjaigarh"),
                   release("2", "Mahtari Vandan payment released"),
                   release("3", "Janjatiya Gaurav Divas celebration")]
        self.assertEqual(len(safe_clusters(stories)), 3)

    def test_same_scheme_different_actions_stay_separate(self):
        stories = [release("1", "Paddy procurement begins"),
                   release("2", "Paddy procurement smuggling raids")]
        self.assertEqual(len(safe_clusters(stories)), 2)

    def test_same_subject_action_date_groups_district_facts(self):
        stories = [release("1", "Raipur: Paddy procurement begins", "100 farmers"),
                   release("2", "Durg: Paddy procurement begins", "250 farmers")]
        clusters = safe_clusters(stories)
        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0]["dropped_ids"], [])
        self.assertEqual(len(clusters[0]["releases"]), 2)

    def test_different_publication_dates_stay_separate(self):
        stories = [release("1", "Paddy procurement begins"),
                   release("2", "Paddy procurement begins", date="2026-11-17")]
        self.assertEqual(len(safe_clusters(stories)), 2)

    def test_different_event_dates_stay_separate(self):
        stories = [release("1", "Paddy procurement begins on 15 November 2026"),
                   release("2", "Paddy procurement begins on 20 November 2026")]
        self.assertEqual(len(safe_clusters(stories)), 2)

    def test_incidental_body_mentions_do_not_merge(self):
        stories = [release("1", "Health camp opens", "Mahtari Vandan payment mentioned."),
                   release("2", "Mahtari Vandan payment released")]
        self.assertEqual(len(safe_clusters(stories)), 2)

    def test_only_identical_text_is_dropped(self):
        title = "Paddy procurement begins"
        stories = [release("1", title, "100 farmers"), release("2", title, "100 farmers"),
                   release("3", title, "100 farmers; 20 centres"),
                   release("4", title, "100 farmers", district="Durg")]
        cluster = safe_clusters(stories)[0]
        self.assertEqual(cluster["dropped_ids"], ["2"])
        self.assertEqual(len(cluster["urls"]), 4)

    def test_proposed_overlap_and_unsafe_drops_are_revalidated(self):
        a = release("1", "Pahadi Korwa hospital opens")
        b = release("2", "Mahtari Vandan payment released")
        proposals = [{"releases": [a, b], "dropped_ids": ["1"]}, {"releases": [a]}]
        result = cg.TopicClusterer._consolidate_overlapping_clusters(proposals)
        self.assertEqual(len(result), 2)
        self.assertEqual([r["id"] for c in result for r in c["releases"]], ["1", "2"])
        self.assertTrue(all(not c["dropped_ids"] for c in result))


class ScraperFailureTests(unittest.TestCase):
    def test_missing_viewstate_never_uses_default_page(self):
        scraper = PIBScraper()
        scraper._request_with_retry = Mock(return_value=Mock(text='<div class="content-area"></div>'))
        with self.assertRaises(ParseError):
            scraper.fetch_releases(TARGET)
        self.assertEqual(scraper._request_with_retry.call_count, 1)

    def test_hidden_attribute_order_and_split_viewstate_are_preserved(self):
        initial = '<input value="state" name="__VIEWSTATE" type="hidden">' \
                  '<input type="hidden" value="part2" name="__VIEWSTATE1">'
        final = calendar() + '<div class="content-area">No releases found</div>'
        scraper = PIBScraper()
        scraper._request_with_retry = Mock(side_effect=[Mock(text=initial), Mock(text=final)])
        releases, _ = scraper.fetch_releases(TARGET)
        self.assertEqual(releases, [])
        self.assertEqual(scraper._request_with_retry.call_args.kwargs["data"]["__VIEWSTATE1"], "part2")

    def test_wrong_returned_date_fails(self):
        with self.assertRaises(ParseError):
            PIBScraper._verify_date(calendar(day=17), TARGET)

    def test_missing_container_is_not_an_empty_day(self):
        with self.assertRaises(ParseError):
            PIBScraper()._parse_releases('<html>Maintenance</html>')

    def test_dprcg_network_failure_is_not_no_releases(self):
        scraper = cg.DPRCGScraper()
        scraper._request_with_retry = Mock(side_effect=FetchError("offline"))
        with self.assertRaises(FetchError):
            scraper.fetch_date_releases(TARGET)

    def test_dprcg_layout_change_is_not_no_releases(self):
        scraper = cg.DPRCGScraper()
        scraper._request_with_retry = Mock(return_value=Mock(
            text="<html>Maintenance</html>", url=cg.DPRCG_DATE_NEWS_URL.format(date_str=TARGET.isoformat())))
        with self.assertRaises(ParseError):
            scraper.fetch_date_releases(TARGET)

    def test_dprcg_explicit_empty_day(self):
        scraper = cg.DPRCGScraper()
        scraper._request_with_retry = Mock(return_value=Mock(
            text="<main>No news found</main>", url=cg.DPRCG_DATE_NEWS_URL.format(date_str=TARGET.isoformat())))
        self.assertEqual(scraper.fetch_date_releases(TARGET)[0], [])

    def test_dprcg_redirect_to_home_fails(self):
        scraper = cg.DPRCGScraper()
        scraper._request_with_retry = Mock(return_value=Mock(text="No news found", url=cg.DPRCG_BASE_URL))
        with self.assertRaises(ParseError):
            scraper.fetch_date_releases(TARGET)


class PipelineTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.stack = contextlib.ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(contextlib.redirect_stdout(io.StringIO()))
        self.stack.enter_context(patch.object(pib, "OUTPUT_DIR", self.root))
        self.stack.enter_context(patch.object(cg, "DPRCG_OUTPUT_BASE", self.root / "dprcg"))

    @property
    def pib_manifest(self):
        return self.root / f"PIB_UPSC_Daily_{TARGET.isoformat()}_status.json"

    @property
    def cg_manifest(self):
        return self.root / "dprcg" / TARGET.isoformat() / "run_status.json"

    def mock_pib(self, releases):
        scraper = self.stack.enter_context(patch.object(pib, "PIBScraper")).return_value
        scraper.fetch_releases.return_value = (releases, "16 November 2026")
        scraper.enrich_releases.return_value = releases
        return scraper

    def mock_cg(self, releases):
        scraper = self.stack.enter_context(patch.object(cg, "DPRCGScraper")).return_value
        scraper.fetch_date_releases.return_value = (releases, "16 November 2026")
        scraper.enrich_all.return_value = releases
        return scraper

    def test_no_releases_invalidates_old_complete_manifest(self):
        self.pib_manifest.write_text('{"status":"complete","files":{"docx":"old"}}')
        scraper = self.mock_pib([])
        with self.assertRaises(NoReleasesError):
            pib.run_pipeline(TARGET)
        data = json.loads(self.pib_manifest.read_text())
        self.assertEqual(data["status"], "no_releases")
        self.assertEqual(data["files"], {})
        scraper.fetch_releases.assert_called_once_with(target_date=TARGET)

    def test_missing_body_withholds_full_report(self):
        self.mock_pib([{"prid": "1", "headline": "Scholarship", "article_fetch_status": "unavailable"}])
        with self.assertRaises(IncompleteRunError):
            pib.run_pipeline(TARGET)
        self.assertEqual(json.loads(self.pib_manifest.read_text())["status"], "partial")
        self.assertEqual(list(self.root.glob("*.docx")), [])

    def test_locked_filename_is_used_in_result_and_manifest(self):
        self.mock_pib([{"prid": "1", "headline": "New scholarship", "body_text": "Education policy.",
                        "ministry": "Ministry of Education", "article_fetch_status": "ok",
                        "url": "https://www.pib.gov.in/"}])

        def alternate(callback, path):
            path = str(Path(path).with_name(Path(path).stem + " (1)" + Path(path).suffix))
            callback(path)
            return path

        with patch.object(pib.DocumentBuilder, "_safe_save_file", side_effect=alternate):
            result = pib.run_pipeline(TARGET)
        files = completed_files(self.pib_manifest, TARGET.isoformat())
        self.assertEqual(Path(result[0]), files["xlsx"])
        self.assertEqual(Path(result[1]), files["docx"])
        self.assertIn("(1)", result[1])

    def test_failed_generation_does_not_publish_stale_files(self):
        self.mock_pib([{"prid": "1", "body_text": "Education", "article_fetch_status": "ok"}])
        with patch.object(pib.DocumentBuilder, "generate_excel", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                pib.run_pipeline(TARGET)
        with self.assertRaises(IncompleteRunError):
            completed_files(self.pib_manifest, TARGET.isoformat())

    def test_dprcg_partial_report_is_labelled_and_failure_is_recorded(self):
        stories = [release("1", "Scholarship launched", "New education scholarship."),
                   release("2", "Solar energy policy", "New solar energy policy.")]
        self.mock_cg(stories)
        good = {"tag": "CHHATTISGARH", "title": "Scholarship", "why_in_news": "Scholarship launched.",
                "sections": [{"heading": "Facts", "items": [
                    {"type": "point", "level": 0, "label": "Education", "text": "Scholarship approved."}]}]}
        synth = self.stack.enter_context(patch.object(cg, "GeminiCurrentAffairsSynthesizer")).return_value
        synth.synthesize_clusters.return_value = [good]
        synth.failed_topics = [{"topic": "Solar energy policy", "release_ids": ["2"]}]
        with self.assertRaises(IncompleteRunError):
            cg.run_dprcg_pipeline(TARGET)
        data = json.loads(self.cg_manifest.read_text())
        self.assertEqual(data["status"], "partial")
        self.assertEqual((data["expected_topics"], data["completed_topics"]), (2, 1))
        self.assertEqual(data["failed_topics"][0]["release_ids"], ["2"])
        self.assertIn("PARTIAL", data["files"]["docx"])
        self.assertIn("PARTIAL", Path(data["files"]["txt"]).read_text(encoding="utf-8"))

    def test_synthesis_exhausted_retries_tracks_missing_topic(self):
        synth = cg.GeminiCurrentAffairsSynthesizer(api_key="test")
        cluster = safe_clusters([release("1", "Scholarship launch")])[0]
        with patch.object(synth, "synthesize_cluster", return_value=None), patch.object(cg.time, "sleep"):
            result = synth.synthesize_clusters([cluster])
        self.assertEqual(result, [])
        self.assertEqual(synth.failed_topics[0]["release_ids"], ["1"])

    def test_no_llm_is_explicit_index_only(self):
        self.mock_cg([release("1", "Scholarship launched", "Education scholarship.")])
        result = cg.run_dprcg_pipeline(TARGET, no_llm=True)
        self.assertEqual(result[:2], (None, None))
        self.assertEqual(json.loads(self.cg_manifest.read_text())["status"], "index_only")

    def test_missing_dprcg_body_prevents_generation(self):
        self.mock_cg([release("1", "Scholarship launched", "")])
        with self.assertRaises(IncompleteRunError):
            cg.run_dprcg_pipeline(TARGET)
        data = json.loads(self.cg_manifest.read_text())
        self.assertEqual(data["failed_article_ids"], ["1"])
        self.assertEqual(data["files"], {})

    def test_cli_no_releases_exits_nonzero(self):
        self.mock_pib([])
        with patch("sys.argv", ["main.py", "--date", TARGET.isoformat()]):
            with self.assertRaises(SystemExit) as raised:
                pib.main()
        self.assertNotEqual(raised.exception.code, 0)

    def test_email_rejects_archived_files_without_completed_manifest(self):
        (self.root / f"PIB_UPSC_Daily_{TARGET.isoformat()}.docx").write_bytes(b"old report")
        with patch.object(send_email, "OUTPUT_DIR", self.root), \
             patch.object(send_email, "load_env_file"), \
             patch.dict("os.environ", {"MAIL_USERNAME": "test@example.com", "MAIL_PASSWORD": "test"}), \
             patch.object(send_email.smtplib, "SMTP") as smtp:
            with self.assertRaises(FileNotFoundError):
                send_email.send_daily_brief_email(TARGET.isoformat())
            smtp.assert_not_called()

    def test_complete_manifest_rejects_wrong_date_and_missing_attachment(self):
        with RunStatus(self.pib_manifest, TARGET) as status:
            status.update(status="complete", files={})
        with self.assertRaises(IncompleteRunError):
            completed_files(self.pib_manifest, "2026-11-17")
        with self.assertRaises(IncompleteRunError):
            completed_files(self.pib_manifest, TARGET.isoformat())


if __name__ == "__main__":
    unittest.main()
