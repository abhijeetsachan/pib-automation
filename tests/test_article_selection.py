"""Offline selection regressions; no scraping, generation, or email calls."""

import unittest
from unittest.mock import Mock

from article_selection import contains
from dprcg_automation import CGPSCClassifier
from pib_scraper import PIBScraper
from upsc_classifier import UPSCClassifier


class PIBSelectionTests(unittest.TestCase):
    def setUp(self):
        self.classifier = UPSCClassifier()

    def test_scholarship_is_social_justice(self):
        result = self.classifier.classify_release(
            "Government launches scholarship for schoolchildren", "Ministry of Education")
        self.assertTrue(result["is_relevant"])
        self.assertEqual(result["paper"], "GS-2")

    def test_welcomes_does_not_veto_policy(self):
        result = self.classifier.classify_release(
            "Minister welcomes new biodiversity conservation policy", "Environment Ministry")
        self.assertTrue(result["is_relevant"])
        self.assertEqual(result["paper"], "GS-3")

    def test_body_decides_generic_headline(self):
        result = self.classifier.process_all([{
            "headline": "Government announces new measures", "ministry": "Government",
            "body_text": "A scholarship for schoolchildren will fund access to education.",
        }])[0]
        self.assertTrue(result["is_relevant"])
        self.assertEqual(result["paper"], "GS-2")

    def test_body_noise_does_not_veto_policy(self):
        result = self.classifier.classify_release(
            "Cabinet approves renewable energy policy", "Cabinet",
            "The minister welcomes delegates and congratulates the hosts.")
        self.assertTrue(result["is_relevant"])

    def test_institution_alone_does_not_override_greeting(self):
        result = self.classifier.classify_release("PM greets ISRO staff on festival", "PMO")
        self.assertFalse(result["is_relevant"])

    def test_ministry_alone_is_not_topic_evidence(self):
        result = self.classifier.classify_release("Annual staff picnic", "Ministry of External Affairs")
        self.assertFalse(result["is_relevant"])

    def test_science_congratulations_with_actual_launch(self):
        result = self.classifier.classify_release(
            "PM congratulates ISRO", "PMO", "ISRO launched a new satellite launch mission.")
        self.assertTrue(result["is_relevant"])

    def test_enrichment_preserves_order_and_failed_fetch(self):
        scraper = PIBScraper()
        scraper.fetch_article_text = Mock(side_effect=["Education scholarship announced.", ""])
        releases = [{"prid": "1"}, {"prid": "2"}]
        result = scraper.enrich_releases(releases)
        self.assertEqual([r["prid"] for r in result], ["1", "2"])
        self.assertEqual([r["article_fetch_status"] for r in result], ["ok", "unavailable"])
        self.assertNotIn("body_text", releases[0])

    def test_existing_body_is_not_downloaded_again(self):
        scraper = PIBScraper()
        scraper.fetch_article_text = Mock()
        scraper.enrich_releases([{"prid": "1", "body_text": "Existing text"}])
        scraper.fetch_article_text.assert_not_called()

    def test_extraction_omits_navigation(self):
        scraper = PIBScraper()
        scraper._request_with_retry = Mock(return_value=Mock(text='''
            <p>Courtesy visit congratulations and greetings from navigation.</p>
            <div class="ReleaseLang"><p>The government launches a scholarship for schoolchildren.</p></div>
        '''))
        text = scraper.fetch_article_text("1")
        self.assertIn("scholarship", text)
        self.assertNotIn("navigation", text)


class CGPSCSelectionTests(unittest.TestCase):
    def classify(self, title, body="", generated_headline=""):
        return CGPSCClassifier.classify_release({
            "raw_title": title, "headline": generated_headline or title, "body_text": body})

    def test_generic_official_courtesy_visit(self):
        self.assertFalse(self.classify("Chief minister courtesy visit")["is_relevant"])

    def test_many_officials_do_not_defeat_routine_filter(self):
        self.assertFalse(self.classify("Chief minister and governor met with collector")["is_relevant"])

    def test_hindi_courtesy_visit(self):
        self.assertFalse(self.classify("मुख्यमंत्री से प्रतिनिधिमंडल की शिष्टाचार भेंट")["is_relevant"])

    def test_hindi_policy_approval(self):
        result = self.classify("राज्य सरकार ने नई औद्योगिक नीति को मंजूरी दी")
        self.assertTrue(result["is_relevant"])
        self.assertEqual(result["paper"], "CGPSC-P5")

    def test_hindi_body_of_generic_release(self):
        result = self.classify("महत्वपूर्ण घोषणा", "छात्रवृत्ति के लिए नई स्वीकृति जारी की गई।")
        self.assertTrue(result["is_relevant"])
        self.assertEqual(result["paper"], "CGPSC-P7")

    def test_policy_change_survives_courtesy_headline(self):
        result = self.classify("Chief minister courtesy meeting", "A new industrial policy was approved.")
        self.assertTrue(result["is_relevant"])
        self.assertEqual(result["paper"], "CGPSC-P5")

    def test_generated_label_cannot_create_evidence(self):
        result = self.classify("प्रशासनिक सूचना", generated_headline="Women Empowerment Government Welfare Scheme")
        self.assertFalse(result["is_relevant"])

    def test_no_substring_acronym_match(self):
        self.assertFalse(contains("shared financial information", "fra"))

    def test_hindi_vowel_marks_match(self):
        self.assertTrue(contains("नई औद्योगिक नीति लागू", "औद्योगिक नीति"))


if __name__ == "__main__":
    unittest.main()
