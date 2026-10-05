"""PIB Web Scraper module with robust ASP.NET postback handling and retry mechanisms."""

import re
import time
import logging
from typing import List, Dict, Optional, Tuple
from datetime import datetime, date
import requests
from bs4 import BeautifulSoup

from config import (
    PIB_BASE_URL,
    PIB_ALL_REL_URL,
    PIB_DETAIL_URL,
    DEFAULT_HEADERS,
    REQUEST_TIMEOUT,
    MAX_RETRIES,
    RETRY_BACKOFF,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


class PIBScraper:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)

    def _request_with_retry(self, method: str, url: str, **kwargs) -> requests.Response:
        """Executes HTTP request with exponential backoff retry."""
        kwargs.setdefault("timeout", REQUEST_TIMEOUT)
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self.session.request(method, url, **kwargs)
                if resp.status_code == 200:
                    return resp
                logger.warning("Attempt %d: HTTP %d received from %s", attempt, resp.status_code, url)
            except (requests.RequestException, Exception) as exc:
                logger.warning("Attempt %d: Network error connecting to %s: %s", attempt, url, exc)
            
            if attempt < MAX_RETRIES:
                sleep_time = RETRY_BACKOFF * attempt
                logger.info("Waiting %d seconds before retrying...", sleep_time)
                time.sleep(sleep_time)

        raise RuntimeError(f"Failed to fetch {url} after {MAX_RETRIES} attempts.")

    def fetch_releases(self, target_date: Optional[date] = None) -> Tuple[List[Dict], str]:
        """
        Fetches all press releases from PIB for the given date.
        If target_date is None or today's date, performs a standard GET request.
        If a historical/past date is requested, performs an ASP.NET PostBack.
        
        Returns:
            Tuple of (releases_list, formatted_date_string)
        """
        today = date.today()
        is_today = target_date is None or target_date == today
        
        target = target_date if target_date else today
        formatted_date = target.strftime("%d %B %Y")
        logger.info("Initiating PIB release fetch for date: %s", formatted_date)

        # Step 1: Initial GET to obtain session and ViewState
        resp = self._request_with_retry("GET", PIB_ALL_REL_URL)
        html_content = resp.text

        # Step 2: If a past date is specified, trigger ASP.NET PostBack
        if not is_today:
            logger.info("Performing ASP.NET PostBack for date: day=%d, month=%d, year=%d", target.day, target.month, target.year)
            viewstate_match = re.search(r'id="__VIEWSTATE"\s+value="([^"]+)"', html_content)
            viewstate_gen_match = re.search(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]+)"', html_content)
            event_val_match = re.search(r'id="__EVENTVALIDATION"\s+value="([^"]+)"', html_content)

            if not viewstate_match:
                logger.error("Could not extract __VIEWSTATE from PIB response. Falling back to default response.")
            else:
                post_data = {
                    "__VIEWSTATE": viewstate_match.group(1),
                    "__VIEWSTATEGENERATOR": viewstate_gen_match.group(1) if viewstate_gen_match else "",
                    "__EVENTTARGET": "ctl00$ContentPlaceHolder1$ddlday",
                    "__EVENTARGUMENT": "",
                    "ctl00$Bar1$ddlregion": "48",  # 48 = National (All India)
                    "ctl00$Bar1$ddlLang": "1",     # 1 = English
                    "ctl00$ContentPlaceHolder1$ddlMinistry": "0",  # 0 = All Ministries
                    "ctl00$ContentPlaceHolder1$ddlday": str(target.day),
                    "ctl00$ContentPlaceHolder1$ddlMonth": str(target.month),
                    "ctl00$ContentPlaceHolder1$ddlYear": str(target.year),
                }
                if event_val_match:
                    post_data["__EVENTVALIDATION"] = event_val_match.group(1)

                post_resp = self._request_with_retry("POST", PIB_ALL_REL_URL, data=post_data)
                html_content = post_resp.text

        # Step 3: Parse releases from HTML
        releases = self._parse_releases(html_content)
        logger.info("Successfully extracted %d releases for %s", len(releases), formatted_date)
        return releases, formatted_date

    def _parse_releases(self, html: str) -> List[Dict]:
        """Parses releases categorized by Ministry from the content-area div."""
        soup = BeautifulSoup(html, "html.parser")
        content_area = soup.find("div", class_="content-area")
        if not content_area:
            logger.warning("Could not find div.content-area in PIB HTML response.")
            return []

        releases = []
        current_ministry = "Unknown Ministry"

        for elem in content_area.find_all(["h3", "ul"]):
            if elem.name == "h3":
                current_ministry = elem.get_text(strip=True)
            elif elem.name == "ul" and "num" in elem.get("class", []):
                for li in elem.find_all("li"):
                    a_tag = li.find("a")
                    if a_tag and a_tag.get("href"):
                        href = a_tag["href"]
                        headline = a_tag.get_text(strip=True)
                        title_attr = a_tag.get("title", "").strip()
                        final_headline = title_attr if title_attr and len(title_attr) > len(headline) else headline
                        
                        prid_match = re.search(r'PRID=(\d+)', href)
                        prid = prid_match.group(1) if prid_match else ""
                        
                        full_url = f"{PIB_BASE_URL}{href}" if href.startswith("/") else href
                        direct_content_url = PIB_DETAIL_URL.format(prid=prid) if prid else full_url

                        releases.append({
                            "prid": prid,
                            "ministry": current_ministry,
                            "headline": final_headline,
                            "url": full_url,
                            "direct_content_url": direct_content_url,
                        })

        return releases

    def fetch_article_text(self, prid: str, max_chars: int = 1500) -> str:
        """
        Fetches the article text directly from the reader pane iframe (PressReleasePage.aspx).
        Bypasses the heavy JavaScript detail shell.
        """
        if not prid:
            return ""
        url = PIB_DETAIL_URL.format(prid=prid)
        try:
            resp = self._request_with_retry("GET", url)
            soup = BeautifulSoup(resp.text, "html.parser")
            paras = soup.find_all("p")
            clean_paras = [p.get_text(" ", strip=True) for p in paras if len(p.get_text(strip=True)) > 25]
            full_text = " ".join(clean_paras)
            return full_text[:max_chars].strip()
        except Exception as err:
            logger.warning("Error fetching article text for PRID %s: %s", prid, err)
            return ""
