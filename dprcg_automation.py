"""
DPRCG Daily News Automation Engine with LLM Integration
-------------------------------------------------------
Fetches official daily news articles from the Directorate of Public Relations,
Government of Chhattisgarh (DPRCG: https://dprcg.gov.in), ingests full article
bodies, classifies them for CGPSC Civil Services relevance, and synthesizes them
using Google Gemini 3.8 Flash into an authoritative current affairs daily bulletin.

Key Features:
- Self-contained single-file automation (does NOT touch or alter PIB automation).
- Native LLM Synthesis (Gemini 3.8 Flash): Transforms raw Hindi/English DPRCG releases
  into structured, high-yield English current affairs articles.
- Reverse-Engineered Structure:
    [DATE HEADER] (e.g. 02-11-2026)
    CHHATTISGARH
    <Source URL>
    <Article Title in Title Case>
    Why in News?
    <Context Hook>
    Key Points
    <Label 1>: <Fact>
    <Label 2>: <Fact>
    | Read More: <Static Syllabus Topics> |
- Single Consolidated Output: All high-yield CGPSC articles saved in one master Word (.docx)
  document named `DD-MM-YYYY.docx` (e.g. `02-11-2026.docx`) and matching `.txt` file.
- Explicit Run Status: Missing data and failed topics are reported without date fallback.
- Date-wise Organized Output: `output/dprcg/<YYYY-MM-DD>/`.
"""

import os
import re
import sys
import time
import json
import logging
import argparse
import urllib.parse
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

# Reconfigure stdout for Windows terminal Unicode compatibility
sys.stdout.reconfigure(encoding="utf-8")

import requests
from bs4 import BeautifulSoup
from article_selection import (CGPSC_HINDI_KEYWORDS, CGPSC_ROUTINE_TERMS,
                               contains, has_substantive_evidence, weighted_matches)
from topic_clustering import safe_clusters
from run_status import FetchError, ParseError, NoReleasesError, IncompleteRunError, RunStatus
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Document generation libraries
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE


# ==============================================================================
# ENVIRONMENT & CONFIGURATION
# ==============================================================================

BASE_DIR = Path(__file__).resolve().parent

def load_env_file():
    """Loads key-value pairs from .env into os.environ if not already set."""
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())

load_env_file()

DPRCG_OUTPUT_BASE = BASE_DIR / "output" / "dprcg"
DPRCG_OUTPUT_BASE.mkdir(parents=True, exist_ok=True)

DPRCG_BASE_URL = "https://dprcg.gov.in"
DPRCG_LANG_EN_URL = "https://dprcg.gov.in/lang/en"
DPRCG_DATE_NEWS_URL = "https://dprcg.gov.in/news/date-wise-news/{date_str}"

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive"
}

REQUEST_TIMEOUT = 20
MAX_RETRIES = 3
RETRY_BACKOFF = 2
MAX_CONCURRENT_WORKERS = 6

# Gemini LLM Configuration
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODELS = ["gemini-3.8-flash", "gemini-3.5-flash", "gemini-3.1-flash-lite"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("DPRCG")


# ==============================================================================
# CGPSC SYLLABUS DEFINITIONS & KEYWORD DICTIONARIES
# ==============================================================================

CGPSC_PAPERS = {
    "CGPSC-P3": {
        "title": "Paper III: Polity, Constitution & State Governance",
        "badge_color": "DBEAFE",
        "keywords": [
            "governance", "cabinet", "chief minister", "governor", "panchayat", "panchayati raj",
            "urban local body", "municipal", "nagar nigam", "nagar palika", "collector", "vidhan sabha",
            "assembly", "high court", "bilaspur high court", "ordinance", "bill", "act",
            "public administration", "administrative reform", "e-governance", "transparency",
            "police administration", "law and order", "revenue administration", "civil services",
            "election", "good governance", "inter-state", "bureaucracy", "ombudsman", "lokayukta"
        ]
    },
    "CGPSC-P4": {
        "title": "Paper IV: Science, Technology & Environment of CG",
        "badge_color": "DCFCE7",
        "keywords": [
            "environment", "biodiversity", "wildlife", "national park", "tiger reserve",
            "kanger valley", "guru ghasidas", "indravati", "sanctuary", "forest cover",
            "renewable energy", "solar energy", "green energy", "biofuel", "climate change",
            "pollution control", "water conservation", "groundwater", "jal jeevan", "science and technology",
            "drone technology", "information technology", "creda", "forest conservation", "waste management"
        ]
    },
    "CGPSC-P5": {
        "title": "Paper V: Economy, Agriculture & Geography of CG",
        "badge_color": "FEF3C7",
        "keywords": [
            "economy", "gdp", "state budget", "economic survey", "investment", "industrial policy",
            "industry", "msme", "mining", "mineral", "iron ore", "coal", "bauxite", "tin", "dolomite",
            "bailadila", "korba", "agriculture", "farmer", "kisan", "paddy procurement",
            "minimum support price", "msp", "krishak unnati yojana", "irrigation", "dam", "canal",
            "mahanadi", "shivnath", "indravati", "river basin", "infrastructure", "national highway",
            "pwd", "road connectivity", "railway network", "power generation", "cspgcl", "electricity",
            "minor forest produce", "tendu patta", "mandi", "food processing"
        ]
    },
    "CGPSC-P6": {
        "title": "Paper VI: Tribal Culture, Heritage & Society of CG",
        "badge_color": "F3E8FF",
        "keywords": [
            "tribal", "tribe", "scheduled tribe", "bastar", "surguja", "pvtg", "gond", "baiga",
            "kamar", "birhor", "abhujmaria", "pahadi korwa", "culture", "heritage", "festival",
            "rajyotsav", "bastar dussehra", "hareli", "chherchhera", "pola", "teeja", "folk dance",
            "pandwani", "karma", "panthi", "raut nacha", "suwa dance", "dadariya", "handicraft",
            "bell metal", "dhokra", "terracotta", "kosa silk", "woodcraft", "archaeology",
            "sirpur", "bhoramdeo", "rajim", "tourism", "customs", "tradition", "folk art"
        ]
    },
    "CGPSC-P7": {
        "title": "Paper VII: Welfare Schemes, Development & Social Justice",
        "badge_color": "FCE7F3",
        "keywords": [
            "welfare", "scheme", "mahtari vandan", "mahatari vandan yojana", "kisan unnati",
            "ramlala darshan", "deendayal antyodaya", "housing", "pm awas", "pradhan mantri awas",
            "health", "ayushman", "hospital", "medical college", "education", "school",
            "swami atmanand", "higher education", "women empowerment", "self-help group",
            "shg", "bihan", "nrlm", "nutrition", "poshan", "social security", "pension",
            "scheduled castes", "vulnerable sections", "labour welfare", "skill development",
            "employment", "forest rights act", "fra", "pesa", "food security", "pds"
        ]
    },
    "CGPSC-PRE": {
        "title": "Prelims: State Current Affairs & Milestones",
        "badge_color": "E0F2FE",
        "keywords": [
            "award", "national award", "state award", "record", "sports", "games", "academy",
            "summit", "mou signed", "inaugurated", "foundation stone", "exhibition",
            "air show", "padma shri", "padma bhushan", "governor", "dignitary visit"
        ]
    }
}

ROUTINE_KEYWORDS = [
    "condolence", "greetings on birthday", "met with", "courtesy visit", "wished happy",
    "transfers", "posting", "inspected road", "reviewed arrangements", "meeting chaired",
    "routine inspection", "traffic advisory", "tender notice", "encroachment removed"
]

DISTRICT_NAMES = [
    "Raipur", "Bilaspur", "Durg", "Bastar", "Surguja", "Rajnandgaon", "Korba",
    "Raigarh", "Janjgir-Champa", "Balod", "Baloda Bazar", "Balrampur", "Bemetara",
    "Bijapur", "Dantewada", "Dhamtari", "Gariaband", "Gaurela-Pendra-Marwahi",
    "Jashpur", "Kabirdham", "Kanker", "Kondagaon", "Khairagarh-Chhuikhadan-Gandai",
    "Korea", "Mahasamund", "Manendragarh-Chirmiri-Bharatpur", "Mohla-Manpur-Ambagarh Chowki",
    "Mungeli", "Narayanpur", "Sakti", "Sarangarh-Bilaigarh", "Sukma", "Surajpur"
]

HINDI_TO_ENGLISH_DISTRICTS = {
    "रायपुर": "Raipur", "बिलासपुर": "Bilaspur", "दुर्ग": "Durg", "बस्तर": "Bastar",
    "सरगुजा": "Surguja", "राजनंदगांव": "Rajnandgaon", "कोरबा": "Korba", "रायगढ़": "Raigarh",
    "जांजगीर-चांपा": "Janjgir-Champa", "जांजगीर": "Janjgir-Champa", "चांपा": "Janjgir-Champa",
    "बालोद": "Balod", "बलौदाबाजार-भाटापारा": "Baloda Bazar", "बलौदाबाजार": "Baloda Bazar",
    "बलरामपुर-रामानुजगंज": "Balrampur", "बलरामपुर": "Balrampur", "बेमेतरा": "Bemetara",
    "बीजापुर": "Bijapur", "दंतेवाड़ा": "Dantewada", "दंतेवाड़ा": "Dantewada", "धमतरी": "Dhamtari",
    "गरियाबंद": "Gariaband", "गौरेला-पेंड्रा-मरवाही": "Gaurela-Pendra-Marwahi", "गौरेला": "Gaurela-Pendra-Marwahi",
    "जशपुर": "Jashpur", "कबीरधाम": "Kabirdham", "कवर्धा": "Kabirdham", "कांकेर": "Kanker",
    "कोंडागांव": "Kondagaon", "खैरागढ़-छुईखदान-गंडई": "Khairagarh", "खैरागढ़": "Khairagarh",
    "कोरिया": "Korea", "महासमुंद": "Mahasamund", "मनेंद्रगढ़-चिरमिरी-भरतपुर": "MCB",
    "मोहला-मानपुर-अंबागढ़ चौकी": "Mohla-Manpur", "मुंगेली": "Mungeli", "नारायणपुर": "Narayanpur",
    "सक्ती": "Sakti", "सारंगढ़-बिलाईगढ़": "Sarangarh-Bilaigarh", "सुकमा": "Sukma", "सूरजपुर": "Surajpur"
}

ADMIN_TERMS_MAP = {
    "मुख्यमंत्री": "Chief Minister",
    "राज्यपाल": "Governor",
    "उप मुख्यमंत्री": "Deputy Chief Minister",
    "वित्त मंत्री": "Finance Minister",
    "ओ.पी. चौधरी": "O.P. Choudhary",
    "विष्णु देव साय": "Vishnu Deo Sai",
    "रमेन डेका": "Ramen Deka",
    "अरुण साव": "Arun Sao",
    "विजय शर्मा": "Vijay Sharma",
    "राज्योत्सव": "State Festival (Rajyotsav)",
    "धान खरीदी": "Paddy Procurement",
    "किसान": "Farmers Welfare",
    "महतारी वंदन": "Mahtari Vandan Scheme",
    "स्वच्छ भारत": "Swachh Bharat Mission",
    "जल जीवन": "Jal Jeevan Mission",
    "पेयजल": "Drinking Water Supply",
    "समीक्षा बैठक": "Administrative Review Meeting",
    "बैठक": "Review Meeting",
    "शिविर": "Public Welfare & Grievance Redressal Camp",
    "योजना": "Government Welfare Scheme",
    "शिल्पग्राम": "Crafts & Cultural Village Exhibition",
    "लोककला": "Folk Art and Cultural Tradition",
    "संस्कृति": "Cultural Heritage Programme",
    "वन संरक्षण": "Forest Conservation & MFP Livelihood",
    "पौधरोपण": "Plantation Drive",
    "स्व-सहायता समूह": "Women Self-Help Groups (SHGs)",
    "महिला": "Women Empowerment",
    "रोजगार": "Employment Generation",
    "शिक्षा": "Education Initiatives",
    "अस्पताल": "Healthcare & Medical Services",
    "सड़क": "Road & Infrastructure Development",
    "निरीक्षण": "Administrative Inspection",
    "स्वीकृति": "Sanction & Development Approval",
    "युवा": "Youth Leadership & Development"
}


# ==============================================================================
# DPRCG WEB SCRAPER ENGINE
# ==============================================================================

class DPRCGScraper:
    """Scrapes date-wise press releases from DPRCG with bilingual extraction."""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)
        self.session.verify = False  # Avoid Govt self-signed cert rejections
        # Do not force lang/en: DPRCG portal stores >95% of state releases under default Hindi,
        # and setting English language session filters out all regular releases.
        # self._init_english_session()

    def _init_english_session(self):
        """Initializes session with DPRCG English language cookie."""
        try:
            self.session.get(DPRCG_LANG_EN_URL, timeout=REQUEST_TIMEOUT)
            logger.info("DPRCG English session mode activated successfully.")
        except Exception as exc:
            logger.warning("Could not set English session cookie: %s. Will rely on slug parsing.", exc)

    def _request_with_retry(self, url: str) -> requests.Response:
        """Executes GET request with exponential retry backoff."""
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self.session.get(url, timeout=REQUEST_TIMEOUT)
                if resp.status_code == 200:
                    return resp
                logger.warning("Attempt %d: HTTP %d for %s", attempt, resp.status_code, url)
            except Exception as exc:
                logger.warning("Attempt %d: Request failed for %s: %s", attempt, url, exc)

            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF * attempt)

        raise FetchError(f"Failed to fetch {url} after {MAX_RETRIES} attempts.")

    @staticmethod
    def _clean_slug_to_english_headline(raw_slug: str, raw_title: str = "") -> Tuple[str, str]:
        """
        Converts a URL slug into a clean English headline and extracts the district.
        Handles Hindi slugs by translating administrative terms into English.
        Returns: (District, Clean English Headline)
        """
        unquoted = urllib.parse.unquote(raw_slug)
        parts = [p for p in unquoted.split("-") if p.strip()]

        district = "Chhattisgarh"
        headline_parts = parts

        if parts:
            for d_name in DISTRICT_NAMES:
                clean_d = d_name.replace(" ", "").replace("-", "").lower()
                if parts[0].lower() == clean_d or parts[0].lower() == d_name.lower():
                    district = d_name
                    headline_parts = parts[1:]
                    break

        # Check raw_title for Hindi district prefix if district not yet found
        if district == "Chhattisgarh" and raw_title:
            if ":" in raw_title:
                h_prefix = raw_title.split(":", 1)[0].strip()
                for h_d, e_d in HINDI_TO_ENGLISH_DISTRICTS.items():
                    if h_d in h_prefix:
                        district = e_d
                        break

        headline = " ".join(headline_parts).strip()
        headline = re.sub(r"\s+", " ", headline)

        # If slug contains Devanagari characters, provide a clean English translation
        if re.search(r"[\u0900-\u097F]", headline):
            source_text = f"{raw_title} {headline}"
            matched_terms = []
            for h_term, e_term in ADMIN_TERMS_MAP.items():
                if h_term in source_text and e_term not in matched_terms:
                    matched_terms.append(e_term)

            if matched_terms:
                headline = f"Administrative Focus: {', '.join(matched_terms[:3])}"
            else:
                headline = f"Official {district} District Administration Press Release"

        if not headline or len(headline) < 3:
            headline = f"{district} District Government Press Release"

        return district, headline

    def fetch_date_releases(self, target_date: date) -> Tuple[List[Dict], str]:
        """
        Fetches all releases for a specific date from DPRCG.
        Returns: Tuple of (releases_list, formatted_date_str)
        """
        date_str = target_date.strftime("%Y-%m-%d")
        formatted_date = target_date.strftime("%d %B %Y")
        url = DPRCG_DATE_NEWS_URL.format(date_str=date_str)

        logger.info("Fetching DPRCG releases for date: %s (%s)", formatted_date, url)

        resp = self._request_with_retry(url)
        returned_path = urllib.parse.urlparse(resp.url).path.rstrip("/")
        if returned_path != urllib.parse.urlparse(url).path.rstrip("/"):
            raise ParseError("DPRCG redirected away from the requested date listing.")

        soup = BeautifulSoup(resp.text, "html.parser")
        items = soup.find_all("div", class_="list-news")

        if not items:
            text = soup.get_text(" ", strip=True).lower()
            if any(marker in text for marker in ("no news found", "no records found", "no data found", "??? ?????? ????")):
                return [], formatted_date
            raise ParseError("DPRCG listing structure missing; cannot confirm an empty news day.")

        logger.info("Discovered %d releases on DPRCG date page.", len(items))

        releases = []
        seen_ids = set()

        for item in items:
            a_tag = item.find("a", href=True)
            if not a_tag:
                continue

            href = a_tag["href"]
            raw_title = a_tag.get_text(strip=True)

            # Match post ID and slug from URL: /post/(\d+)/(.*)
            match = re.search(r"/post/(\d+)(?:/(.*))?", href)
            if not match:
                continue

            post_id = match.group(1)
            raw_slug = match.group(2) or ""

            if post_id in seen_ids:
                continue
            seen_ids.add(post_id)

            # Extract District and English Headline
            district, english_headline = self._clean_slug_to_english_headline(raw_slug, raw_title)

            # If English title in listing is already available (from /lang/en session)
            if ":" in raw_title and not any(ord(c) > 127 for c in raw_title):
                parts = raw_title.split(":", 1)
                district = parts[0].strip()
                english_headline = parts[1].strip()
            elif not any(ord(c) > 127 for c in raw_title) and len(raw_title) > 10:
                english_headline = raw_title.strip()

            clean_url = shorten_dprcg_url(href)
            releases.append({
                "id": post_id,
                "url": clean_url,
                "original_url": href,
                "district": district,
                "headline": english_headline,
                "raw_title": raw_title,
                "date": date_str,
                "body_text": "",
                "summary": "",
                "image_url": ""
            })

        if not releases:
            raise ParseError("DPRCG listing contained items but no valid release links.")
        return releases, formatted_date

    def enrich_article_details(self, release: Dict) -> Dict:
        """Fetches full article text and image from the detail page."""
        url = urllib.parse.urljoin(DPRCG_BASE_URL, release.get("original_url") or release.get("url", ""))
        if not url:
            return release

        try:
            resp = self._request_with_retry(url)
            soup = BeautifulSoup(resp.text, "html.parser")

            # Extract Post Title
            title_tag = soup.find(class_="post-title")
            if title_tag:
                full_title = title_tag.get_text(strip=True)
                if not any(ord(c) > 127 for c in full_title):
                    if ":" in full_title:
                        parts = full_title.split(":", 1)
                        release["district"] = parts[0].strip()
                        release["headline"] = parts[1].strip()
                    else:
                        release["headline"] = full_title

            # Extract Full Article Body Text
            text_div = soup.find(class_="article-text-content")
            if text_div:
                paragraphs = [
                    p.get_text(" ", strip=True)
                    for p in text_div.find_all(["p", "div", "li"])
                    if len(p.get_text(strip=True)) > 20
                ]
                full_body = "\n\n".join(paragraphs) if paragraphs else text_div.get_text("\n", strip=True)
                release["body_text"] = full_body

                # Create 2-3 sentence executive summary
                sentences = re.split(r"(?<=[.!?])\s+", full_body.replace("\n", " "))
                english_sentences = [s.strip() for s in sentences if s.strip() and not any(ord(c) > 127 for c in s)]
                if english_sentences:
                    release["summary"] = " ".join(english_sentences[:3])
                else:
                    release["summary"] = f"Official press release regarding {release['headline']}."

            # Extract Featured Image
            img_div = soup.find(class_="article-img")
            if img_div:
                img_tag = img_div.find("img")
                if img_tag and img_tag.get("src"):
                    release["image_url"] = img_tag["src"]

        except Exception as exc:
            logger.debug("Failed detail fetch for %s: %s", url, exc)
            if not release.get("summary"):
                release["summary"] = f"Government initiative announced: {release['headline']}."

        release["article_fetch_status"] = "ok" if release.get("body_text", "").strip() else "unavailable"
        return release

    def enrich_all(self, releases: List[Dict], max_workers: int = MAX_CONCURRENT_WORKERS) -> List[Dict]:
        """Concurrently enriches all releases with body text and summaries."""
        logger.info("Enriching %d releases in parallel with %d workers...", len(releases), max_workers)
        enriched = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_release = {executor.submit(self.enrich_article_details, rel): rel for rel in releases}
            for future in as_completed(future_to_release):
                try:
                    res = future.result()
                    enriched.append(res)
                except Exception as exc:
                    enriched.append(future_to_release[future])

        # Preserve original release ordering by post ID
        id_order = {r["id"]: idx for idx, r in enumerate(releases)}
        enriched.sort(key=lambda r: id_order.get(r["id"], 9999))
        return enriched


# ==============================================================================
# CGPSC SPECIFIC CLASSIFIER ENGINE
# ==============================================================================

class CGPSCClassifier:
    """Classifies DPRCG releases against the CGPSC State Civil Services Syllabus."""

    @staticmethod
    def classify_release(release: Dict) -> Dict:
        """
        Classifies an article into a CGPSC Paper category and calculates relevance.
        Returns the article dict enriched with CGPSC metadata.
        """
        # Prefer the original title: slug-derived administrative labels and
        # generated summaries are not independent evidence of relevance.
        headline = release.get("raw_title") or release.get("headline", "")
        body = release.get("body_text", "")
        text_corpus = f"{headline} {body}"
        is_routine = any(contains(headline, term)
                         for term in ROUTINE_KEYWORDS + CGPSC_ROUTINE_TERMS)
        best_paper = "Routine"
        best_paper_title = "Local Administration"
        best_score = 0
        matched_keywords = []
        all_specific = []
        for paper_id, config in CGPSC_PAPERS.items():
            terms = config["keywords"] + CGPSC_HINDI_KEYWORDS.get(paper_id, [])
            matches, specific, score = weighted_matches(headline, body, terms)
            all_specific.extend(specific)
            if score > best_score:
                best_score = score
                best_paper = paper_id
                best_paper_title = config["title"]
                matched_keywords = specific

        substantive = has_substantive_evidence(text_corpus, all_specific)
        is_relevant = best_score > 0 and (not is_routine or substantive)
        if is_relevant:
            relevance_reason = f"Relevant for {best_paper_title}. Topics: {', '.join(matched_keywords[:4])}"
        elif is_routine:
            relevance_reason = "Routine or ceremonial visit/update without a substantive policy, data, or research change."
        else:
            relevance_reason = "No specific syllabus subject found in the original headline or article body."

        release.update({
            "is_relevant": is_relevant,
            "paper": best_paper,
            "paper_title": best_paper_title,
            "relevance_score": best_score,
            "relevance_reason": relevance_reason,
            "matched_keywords": matched_keywords
        })
        return release

    def process_all(self, releases: List[Dict]) -> List[Dict]:
        """Classifies and sorts all releases by CGPSC relevance."""
        logger.info("Running CGPSC Syllabus Classification on %d releases...", len(releases))
        processed = [self.classify_release(r) for r in releases]

        # Sort: Relevant first (highest score), then routine
        processed.sort(key=lambda r: (1 if r["is_relevant"] else 0, r["relevance_score"]), reverse=True)
        return processed



ADMIN_TERMS_ORDERED = [
    ("जनजातीय गौरव", "janjatiya-gaurav"),
    ("बिरसा", "birsa-munda"),
    ("आदिवासी", "tribal"),
    ("संग्रहालय", "tribal-museum"),
    ("धान खरीदी", "paddy-procurement"),
    ("उपार्जन", "paddy-procurement"),
    ("समर्थन मूल्य", "msp"),
    ("महतारी वंदन", "mahtari-vandan"),
    ("पहाड़ी कोरवा", "pahadi-korwa"),
    ("विष्णु के पाती", "vishnu-ke-paati"),
    ("सूर्य घर", "surya-ghar"),
    ("मुफ्त बिजली", "solar-power"),
    ("ग्राम सभा", "gram-sabha"),
    ("स्वास्थ्य", "health-camp"),
    ("शौर्य दिवस", "shaurya-diwas"),
    ("दीवाली", "diwali-milan"),
    ("निर्वाचन", "election-commission"),
    ("चुनाव", "election"),
    ("स्वनिधि", "svanidhi-yojana"),
    ("शहरी", "urban-development"),
    ("महिला", "women-empowerment"),
    ("किसान", "farmers"),
    ("विकास", "development"),
]

HINDI_DISTRICTS_LOOKUP = [
    (["रायगढ़", "रायगढ़"], "raigarh"),
    (["रायपुर"], "raipur"),
    (["बालोद"], "balod"),
    (["बस्तर"], "bastar"),
    (["धमतरी"], "dhamtari"),
    (["कांकेर"], "kanker"),
    (["कबीरधाम"], "kabirdham"),
    (["जांजगीर"], "janjgir"),
    (["अम्बिकापुर"], "ambikapur"),
    (["गौरेला"], "gpm"),
    (["गरियाबंद"], "gariyaband"),
    (["सक्ती"], "sakti"),
    (["महासमुंद"], "mahasamund"),
    (["बेमेतरा"], "bemetara"),
    (["सूरजपुर"], "surajpur"),
    (["कोरिया"], "koriya"),
    (["दुर्ग"], "durg"),
    (["बिलासपुर"], "bilaspur"),
    (["राजनांदगांव"], "rajnandgaon"),
]


def shorten_dprcg_url(url: str, max_slug_len: int = 35) -> str:
    """
    Truncates and normalizes lengthy DPRCG URLs into concise, readable, and functional links.
    Eliminates bloated percent-encoded Hindi slugs and lengthy English slugs while keeping
    the critical post ID intact for 100% routing fidelity on DPRCG servers.
    Example:
      Input:  https://dprcg.gov.in/post/1794669874/%E0%A4%B0%E0%A4%BE... (880 chars)
      Output: https://dprcg.gov.in/post/1794669874/raipur-paddy-procurement-msp (~65 chars)
    """
    if not url or "dprcg.gov.in/post/" not in url:
        return url

    match = re.search(r"(https?://dprcg\.gov\.in/post/(\d+))(?:/([^/?#\s]+))?", url)
    if not match:
        return url

    base_prefix = match.group(1)
    raw_slug = match.group(3) or ""
    unquoted = urllib.parse.unquote(raw_slug).strip()

    # Replace smart quotes and dashes with standard hyphens
    unquoted_clean = re.sub(r"[\u2018\u2019\u201C\u201D\u2013\u2014]", "-", unquoted)

    # Check if predominantly ASCII English
    ascii_chars = sum(1 for c in unquoted_clean if ord(c) < 128)
    is_mostly_ascii = (ascii_chars / len(unquoted_clean) > 0.8) if unquoted_clean else False

    if is_mostly_ascii and unquoted_clean:
        clean_words = re.findall(r"[A-Za-z0-9]+", unquoted_clean)
        slug_parts = []
        cur_len = 0
        for w in clean_words:
            if cur_len + len(w) + 1 > max_slug_len and slug_parts:
                break
            slug_parts.append(w.lower())
            cur_len += len(w) + 1
        short_slug = "-".join(slug_parts)
    else:
        # Hindi or percent-encoded slug: match district and key terms
        unquoted_spaces = unquoted.replace("-", " ")
        matched_keywords = []
        for h_term, e_term in ADMIN_TERMS_ORDERED:
            if h_term in unquoted_spaces and e_term not in matched_keywords:
                matched_keywords.append(e_term)

        district_found = ""
        for h_patterns, e_d in HINDI_DISTRICTS_LOOKUP:
            if any(p in unquoted_spaces for p in h_patterns):
                district_found = e_d
                break

        slug_tokens = []
        if district_found:
            slug_tokens.append(district_found)
        if matched_keywords:
            slug_tokens.extend(matched_keywords[:2])
        else:
            slug_tokens.append("news")

        short_slug = "-".join(slug_tokens)[:max_slug_len].rstrip("-")

    if not short_slug:
        short_slug = "news"

    return f"{base_prefix}/{short_slug}"


def format_why_in_news(text: str, existing_bullets: Optional[List[str]] = None) -> Tuple[str, List[str]]:
    """
    Ensures the 'Why in News?' section is crisp, concise, and structured.
    - If the context is lengthy, it splits the narrative so the opening lead
      is strictly 1 concise sentence answering Who, What, When, and Where (max 1-2 lines).
    - The lengthier background, objective, significance, or impact is broken out
      into a clean bullet point with an authoritative label.
    """
    clean_text = (text or "").strip()
    bullets = [b.strip() for b in (existing_bullets or []) if b and b.strip()]

    if bullets:
        return clean_text, bullets

    if not clean_text:
        return "", []

    # Protect common honorifics, titles, and abbreviations from being split
    protected = clean_text
    subs = [
        ("Shri.", "Shri#DOT#"), ("Smt.", "Smt#DOT#"), ("Dr.", "Dr#DOT#"),
        ("Prof.", "Prof#DOT#"), ("Mr.", "Mr#DOT#"), ("Mrs.", "Mrs#DOT#"),
        ("Gov.", "Gov#DOT#"), ("Lt.", "Lt#DOT#"), ("Col.", "Col#DOT#"),
        ("e.g.", "eg#DOT#"), ("i.e.", "ie#DOT#")
    ]
    for orig, rep in subs:
        protected = protected.replace(orig, rep)

    parts = re.split(r'\.\s+(?=[A-Z])', protected)
    if len(parts) <= 1:
        return clean_text, []

    lead = parts[0].strip()
    for orig, rep in subs:
        lead = lead.replace(rep, orig)
    if not lead.endswith('.'):
        lead += '.'

    rest = '. '.join(p.strip() for p in parts[1:])
    for orig, rep in subs:
        rest = rest.replace(rep, orig)
    if not rest.endswith('.'):
        rest += '.'

    if ":" in rest and len(rest.split(":", 1)[0].split()) <= 4:
        bullets.append(rest)
    else:
        rest_lower = rest.lower()
        if any(k in rest_lower for k in ["outlay", "₹", "crore", "budget", "subsidy", "fund"]):
            label = "Financial & Strategic Outlay"
        elif any(k in rest_lower for k in ["platform", "interaction", "showcase", "briefing", "meeting"]):
            label = "Key Focus"
        elif any(k in rest_lower for k in ["aims to", "objective", "mandate", "institutionalize", "designed to"]):
            label = "Policy Objective"
        elif any(k in rest_lower for k in ["empower", "benefit", "pillar", "impact", "transform", "eliminat"]):
            label = "Key Impact"
        elif any(k in rest_lower for k in ["honor", "martyr", "commemorat"]):
            label = "Commemoration"
        elif any(k in rest_lower for k in ["initiative highlights", "commitment", "morale", "foster"]):
            label = "Strategic Focus"
        else:
            label = "Key Highlights"

        bullets.append(f"{label}: {rest}")

    return lead, bullets


# ==============================================================================
# SEMANTIC TOPIC CLUSTERER & REDUNDANCY FILTER
# ==============================================================================

class TopicClusterer:
    """Group only releases with an explicit shared subject, action, and date.

    Clustering is deterministic: ambiguous stories stay independent. LLM guesses
    about redundant sources cannot remove facts from the synthesis corpus.
    """

    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key

    def cluster_releases(self, releases: List[Dict]) -> List[Dict]:
        return safe_clusters(releases)

    @classmethod
    def _consolidate_overlapping_clusters(cls, clusters: List[Dict]) -> List[Dict]:
        # Repartition from source evidence, ignoring proposed topics/drop IDs.
        return safe_clusters([r for cluster in clusters for r in cluster.get("releases", [])])

    def _cluster_heuristic(self, releases: List[Dict]) -> List[Dict]:
        return safe_clusters(releases)


# ==============================================================================
# GEMINI 2.5 FLASH CURRENT AFFAIRS SYNTHESIZER
# ==============================================================================

class GeminiCurrentAffairsSynthesizer:
    """Transforms raw DPRCG releases into consolidated, reverse-engineered exam current affairs articles."""

    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or GEMINI_API_KEY
        self.models = GEMINI_MODELS

    def _get_endpoint(self, model_name: str) -> str:
        return f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={self.api_key}"

    def synthesize_cluster(self, cluster: Dict) -> Optional[Dict]:
        """
        Sends combined release texts of a topic cluster to Gemini to generate
        a unified, de-duplicated, hierarchical current affairs master article.
        """
        if not self.api_key:
            logger.warning("No GEMINI_API_KEY configured. Skipping LLM generation.")
            return None

        topic = cluster.get("topic", "")
        paper_info = cluster.get("paper_title") or cluster.get("paper", "CGPSC State Current Affairs")
        releases = cluster.get("releases", [])
        urls = list(dict.fromkeys(cluster.get("urls", [])))
        dropped_ids = set(cluster.get("dropped_ids", []))

        # Filter active releases (omitting dropped redundant releases from synthesis corpus)
        active_releases = [r for r in releases if str(r.get("id")) not in dropped_ids]
        if not active_releases:
            active_releases = releases

        sources_blocks = []
        for idx, r in enumerate(active_releases, 1):
            h = r.get("headline", "")
            d = r.get("district", "Chhattisgarh")
            u = r.get("url", "")
            b = r.get("body_text", "") or r.get("summary", "")
            if not b.strip():
                b = r.get("raw_title", "")
            sources_blocks.append(f"--- SOURCE RELEASE {idx} [{d}] ---\nHeadline: {h}\nURL: {u}\nContent:\n{b}")
        sources_text = "\n\n".join(sources_blocks)

        urls_formatted = "\n".join(f"* {u}" for u in urls) if len(urls) > 1 else (urls[0] if urls else "")

        prompt = f"""You are an expert editorial writer, syllabus director, and faculty for Chhattisgarh Public Service Commission (CGPSC) Civil Services Current Affairs.
Your task is to transform official Chhattisgarh Government (DPRCG) press releases on the topic '{topic}' into ONE authoritative, unified, and comprehensively structured current affairs article for CGPSC aspirants.

SOURCE INFORMATION:
Number of Source Releases: {len(releases)}
Source URLs:
{urls_formatted}

Sources Content:
{sources_text}

MANDATORY EDITORIAL REQUIREMENTS:

1. HIERARCHICAL HEADINGS (STRICT REQUIREMENT):
   Follow this exact heading hierarchy without deviation:
   Line 1: CHHATTISGARH
   Line 2: Source URL(s):
     - If 1 source URL: output the raw single URL on this line.
     - If 2 or more source URLs: output a bulleted list of all source URLs, one per line:
       * <URL 1>
       * <URL 2>
   Line 3: Main Article Title in H2 format, starting with '## ':
     ## <SEO / AEO / GEO Optimized Article Title in Title Case>
   Section 1: Context hook in H3 format:
     ### Why in News?
     MANDATORY CONCISENESS & BULLET FORMAT (STRICT REQUIREMENT):
     - NEVER write a lengthy continuous paragraph (narrative lead MUST be strictly 1-2 lines).
     - Line 1: Exactly ONE concise lead sentence stating the direct trigger: Who, What, When, Where.
     - Line 2+: If there is additional context, objective, significance, or lengthier background, NEVER merge it into the opening paragraph. Break it out into 1 (maximum 2) clean bullet point(s) directly underneath:
       * <Label>: <Concise description of the objective, focus, or significance>
   Subsequent Sections: Thematic Headings in H3 format:
     ### <Thematic Section Title in Title Case>
   Sub-themes / Components: Subheadings in H4 format under H3 sections:
     #### <Sub-theme Subheading in Title Case>
     Followed by multi-level nested bullets:
     * <Label 1>: <Category header or key metric overview>
       * <Sub-label A>: <Specific detail, breakdown, figure, or metric>
       * <Sub-label B>: <Specific detail, breakdown, figure, or metric>
     * <Label 2>: <Factual statement or standalone point>
   Footer Line:
     | Read More:        |

2. SEO / AEO / GEO TITLE OPTIMIZATION:
   - SEO (Search Engine Optimization): Front-load primary entities ('Chhattisgarh', Scheme/Event Name, Year). Formulated to capture high-intent search queries from aspirants and researchers.
   - AEO (Answer Engine Optimization): Formulated to directly answer user queries on AI search engines (Perplexity, ChatGPT, Claude) with high factual clarity.
   - GEO (Generative Engine Optimization): Authoritative tone, concrete numerical markers (e.g. ₹57,759 Cr, 23 Schemes, 15 Nov, 3100 Centers), and statutory/policy anchors that LLMs cite as source truth.
   - Example of strong H2 title:
     ## Janjatiya Gaurav Divas 2026: Chhattisgarh Government Unveils ₹57,759 Crore Multi-Sectoral Outlay and Cultural Commemoration Roadmap

3. NATURAL MERGING & STRICT DE-DUPLICATION (ZERO REDUNDANCY):
   - You must club all related releases naturally into a single cohesive narrative.
   - DO NOT repeat overlapping facts, dates, background paragraphs, or metrics across sections.
   - If any source release contains repetitive information or lacks distinct new facts (e.g. generic routine announcements simply repeating statewide dates), DROP that redundant content completely.
   - Merge complementary facts: e.g. connect the state-level financial outlay with ground-level district implementation, museum initiatives, or infrastructure provisions.

4. CGPSC EXAM RELEVANCE:
   - Provide concrete policy facts, statutory bodies, constitutional provisions (e.g. PESA, FRA 2006, Articles 244/275), and district details.
   - NEVER use the forbidden word 'Static' in any heading or bullet (e.g., write 'Statutory Framework and Institutional Architecture', NOT 'Static Context').
   - Output in 100% pure, formal English. Do NOT wrap output in markdown code fences.

5. OPTIONAL HIGH-IMPORTANCE CALLOUT BOX (SELECTIVE USE ONLY):
   - If (and ONLY if) an article contains exceptionally high-importance, exam-vital information—such as Constitutional Articles/Provisions (e.g. Articles 243K, 243ZA, 244, 275, 5th Schedule), Core Statutory Frameworks (e.g. PESA Act 1996, FRA 2006, CG Panchayat Raj Act 1993), Landmark Policy Caps/Metrics (e.g. ₹3,100/quintal Paddy procurement formula with 21 quintals/acre cap, or ₹57,759.73 Cr 3-year outlay across 23 schemes)—encapsulate that critical takeaway in a dedicated callout box:
     [BOX: <Clear High-Yield Focus Title>]
     * <Label 1>: <Crisp statutory/constitutional provision or essential exam fact>
     * <Label 2>: <Crisp statutory/constitutional provision or essential exam fact>
     [/BOX]
   - CRITICAL SELECTIVITY RULE: DO NOT include a box in every article! Routine news, courtesy meetings, awards, or local camps must NOT have a box. Use it selectively (only for articles with genuine high-yield statutory, constitutional, or key fiscal takeaways).
"""

        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.2,
                "maxOutputTokens": 8192
            }
        }

        for model_name in self.models:
            endpoint = self._get_endpoint(model_name)
            for attempt in range(1, 3):
                try:
                    resp = requests.post(
                        endpoint,
                        json=payload,
                        headers={"Content-Type": "application/json"},
                        timeout=90
                    )
                    if resp.status_code == 200:
                        data = resp.json()
                        candidates = data.get("candidates", [])
                        if candidates and "content" in candidates[0]:
                            raw_result = candidates[0]["content"]["parts"][0]["text"].strip()
                            finish_reason = candidates[0].get("finishReason")
                            if finish_reason == "MAX_TOKENS":
                                logger.warning("Model hit MAX_TOKENS ceiling. Retrying...")
                                continue

                            is_incomplete = raw_result.strip().endswith(("₹", ":", "and", "the", "with", "of", "a", "an", "at", "in"))
                            if is_incomplete:
                                logger.warning("Synthesis output appears incomplete (ends with dangling token). Retrying...")
                                continue

                            parsed = self._parse_block(raw_result, default_urls=urls)
                            parsed["raw_text"] = raw_result
                            parsed["cluster_topic"] = topic
                            parsed["source_count"] = len(releases)
                            parsed["urls"] = list(dict.fromkeys(urls))

                            total_items = sum(len(s.get("items", [])) for s in parsed.get("sections", []))
                            if len(parsed.get("why_in_news", "")) > 10 and total_items >= 2:
                                return parsed
                    elif resp.status_code == 429:
                        logger.warning("Model %s hit 429 rate limit. Trying fallback...", model_name)
                        time.sleep(2)
                        break
                    elif resp.status_code == 503:
                        logger.warning("Model %s 503 overloaded. Waiting 3s...", model_name)
                        time.sleep(3)
                    else:
                        logger.warning("Model %s HTTP %d: %s", model_name, resp.status_code, resp.text[:120])
                except Exception as exc:
                    logger.warning("Request exception on %s (attempt %d): %s", model_name, attempt, exc)
                    time.sleep(2)

        logger.error("Failed LLM synthesis for cluster: %s", topic)
        return None

    def synthesize_article(self, article: Dict) -> Optional[Dict]:
        """Backwards compatibility helper: synthesizes a single release as a 1-item cluster."""
        cluster = {
            "topic": article.get("headline", ""),
            "paper": article.get("paper", "CGPSC-PRE"),
            "paper_title": article.get("paper_title", "CGPSC Current Affairs"),
            "releases": [article],
            "urls": [article.get("url", "")] if article.get("url") else [],
            "dropped_ids": []
        }
        return self.synthesize_cluster(cluster)

    def synthesize_clusters(self, clusters: List[Dict], max_workers: int = 1) -> List[Dict]:
        """Synthesizes all topic clusters into structured master articles with polite pacing."""
        logger.info("Synthesizing %d consolidated current affairs topics via Gemini LLM...", len(clusters))
        synthesized_articles = []
        self.failed_topics = []

        for idx, cluster in enumerate(clusters, start=1):
            topic = cluster.get("topic", "")
            rel_count = len(cluster.get("releases", []))
            logger.info("[%d/%d] Synthesizing: %s (%d releases)", idx, len(clusters), topic[:60], rel_count)

            result = self.synthesize_cluster(cluster)
            if result:
                synthesized_articles.append(result)
            else:
                logger.warning("Retrying topic [%d] with delay...", idx)
                time.sleep(3)
                retry_res = self.synthesize_cluster(cluster)
                if retry_res:
                    synthesized_articles.append(retry_res)
                else:
                    self.failed_topics.append({
                        "topic": topic,
                        "release_ids": [str(r["id"]) for r in cluster.get("releases", [])],
                    })

            if idx < len(clusters):
                time.sleep(2.5)

        logger.info("Successfully synthesized %d master current affairs articles.", len(synthesized_articles))
        return synthesized_articles

    def synthesize_all(self, relevant_releases: List[Dict], max_workers: int = 1) -> List[Dict]:
        """Clusters releases and synthesizes all consolidated current affairs articles."""
        clusterer = TopicClusterer(api_key=self.api_key)
        clusters = clusterer.cluster_releases(relevant_releases)
        return self.synthesize_clusters(clusters, max_workers=max_workers)

    @staticmethod
    def _parse_block(raw_text: str, default_urls: Optional[List[str]] = None) -> Dict:
        """Parses LLM output into structured fields, supporting H2 title, H3/H4 headings, and nested bullets."""
        clean = raw_text.strip()
        clean = re.sub(r"^```[a-zA-Z]*\n", "", clean)
        clean = re.sub(r"\n```$", "", clean).strip()

        raw_lines = clean.split("\n")
        parsed = {
            "tag": "CHHATTISGARH",
            "urls": list(default_urls) if default_urls else [],
            "url": default_urls[0] if default_urls else "",
            "title": "",
            "why_in_news": "",
            "why_in_news_bullets": [],
            "sections": [],
            "dynamic_heading": "Key Highlights:",
            "key_points": [],
            "read_more": "| Read More:        |"
        }

        state = "HEADER"
        curr_section = None
        prev_bullet_lvl = 0
        in_box = False
        curr_box = None

        for raw_line in raw_lines:
            expanded_line = raw_line.expandtabs(2)
            indent_len = len(expanded_line) - len(expanded_line.lstrip())
            clean_line = raw_line.strip()

            if not clean_line:
                continue

            stripped = clean_line.replace("**", "").strip()

            # Check for callout box tags: [BOX: <Title>] and [/BOX]
            box_start_m = re.match(r"^\[BOX(?::\s*([^\]]+))?\]", clean_line, flags=re.IGNORECASE)
            if box_start_m:
                in_box = True
                b_title = box_start_m.group(1).strip() if box_start_m.group(1) else "Key Exam Takeaway"
                curr_box = {"type": "box", "title": b_title, "items": []}
                if curr_section is None:
                    curr_section = {"heading": "Key Highlights", "items": [], "points": []}
                    parsed["sections"].append(curr_section)
                curr_section["items"].append(curr_box)
                continue

            if clean_line.upper().startswith("[/BOX]") or clean_line.upper() == "[/BOX]":
                in_box = False
                curr_box = None
                continue

            if in_box and curr_box is not None:
                b_item = re.sub(r"^[-*•◦▪▫]\s*", "", clean_line).strip().replace("**", "")
                if b_item:
                    curr_box["items"].append(b_item)
                continue

            if stripped.upper() == "CHHATTISGARH":
                continue

            # Skip URL line or Source URL header in HEADER state
            if state == "HEADER":
                if clean_line.lower().startswith("source url") or clean_line.lower().startswith("source:") or clean_line.lower().startswith("urls:"):
                    continue
                url_match = re.search(r"https?://[^\s]+", clean_line)
                if url_match and (clean_line.startswith("http") or clean_line.startswith("* http") or clean_line.startswith("- http") or clean_line.startswith("• http") or clean_line.startswith("URL:")):
                    u = shorten_dprcg_url(url_match.group(0).rstrip(")"))
                    if u not in parsed["urls"]:
                        parsed["urls"].append(u)
                        if not parsed["url"]:
                            parsed["url"] = u
                    continue

            # Check for H2 Title: "## <Title>" or "Title: <Title>" or first text line in HEADER state
            if state == "HEADER":
                if clean_line.startswith("## "):
                    t = clean_line[3:].strip().replace("**", "")
                    if t:
                        parsed["title"] = t
                    continue
                elif clean_line.lower().startswith("title:"):
                    t = re.sub(r"^title:\s*", "", clean_line, flags=re.IGNORECASE).strip().replace("**", "")
                    if t:
                        parsed["title"] = t
                    continue
                elif not parsed["title"] and not clean_line.startswith("#") and not clean_line.startswith(("*", "-", "•")) and not any(k in clean_line.lower() for k in ["source", "http", "www."]):
                    parsed["title"] = stripped
                    continue

            # Check for "Why in News?" (H3)
            if "why in news" in clean_line.lower() and state in ["HEADER", "WHY_IN_NEWS"]:
                state = "WHY_IN_NEWS"
                continue

            # Check for footer
            if clean_line.startswith("|") and "read more" in clean_line.lower():
                parsed["read_more"] = "| Read More:        |"
                state = "DONE"
                break

            if state == "WHY_IN_NEWS":
                # Check if we transitioned to a new section heading
                is_heading = clean_line.startswith("### ") or (
                    indent_len == 0 and not clean_line.startswith(("-", "•", "*", "◦", "▪", "▫")) and len(stripped) < 90 and not stripped.endswith(".")
                )
                if is_heading:
                    state = "SECTIONS"
                    h_text = re.sub(r"^#+\s*", "", clean_line).strip()
                    h_text = re.sub(r"\bstatic\s+", "", h_text, flags=re.IGNORECASE).strip()
                    h_text = re.sub(r"^\s*context\s*:\s*", "", h_text, flags=re.IGNORECASE).rstrip(":").strip()
                    curr_section = {"heading": h_text, "items": [], "points": []}
                    parsed["sections"].append(curr_section)
                    prev_bullet_lvl = 0
                    continue

                # Check if line is a bullet under Why in News
                is_bullet = clean_line.startswith(("-", "•", "*", "◦", "▪", "▫"))
                if is_bullet:
                    b_text = re.sub(r"^[-*•◦▪▫]\s*", "", stripped).strip()
                    if b_text:
                        parsed["why_in_news_bullets"].append(b_text)
                    continue

                # Otherwise it is lead narrative text
                if not parsed["why_in_news"]:
                    parsed["why_in_news"] = stripped
                else:
                    parsed["why_in_news"] = (parsed["why_in_news"] + " " + stripped).strip()
                continue

            if state == "SECTIONS":
                # Check for H3 Section: "### <Heading>"
                if clean_line.startswith("### "):
                    h_text = clean_line[4:].strip().replace("**", "")
                    h_text = re.sub(r"\bstatic\s+", "", h_text, flags=re.IGNORECASE).strip()
                    h_text = re.sub(r"^\s*context\s*:\s*", "", h_text, flags=re.IGNORECASE).rstrip(":").strip()
                    curr_section = {"heading": h_text, "items": [], "points": []}
                    parsed["sections"].append(curr_section)
                    prev_bullet_lvl = 0
                    continue

                # Check for H4 Sub-theme: "#### <Subheading>"
                if clean_line.startswith("#### "):
                    sub_text = clean_line[5:].strip().replace("**", "").rstrip(":")
                    if curr_section is None:
                        curr_section = {"heading": "Key Highlights", "items": [], "points": []}
                        parsed["sections"].append(curr_section)
                    curr_section["items"].append({"type": "subheading", "text": sub_text})
                    continue

                # Check for bullet points
                is_bullet = bool(re.match(r"^[-•*◦▪]\s+", clean_line))
                pt_clean = re.sub(r"^[-•*◦▪]\s*", "", clean_line).strip().replace("**", "")

                # Unindented non-bullet line that is a heading
                if indent_len == 0 and not is_bullet and len(stripped) < 90 and not stripped.endswith("."):
                    colon_match = re.match(r"^([A-Za-z0-9\s/()\-–&]{2,35}):\s+(.+)$", stripped)
                    if not colon_match or stripped.endswith(":"):
                        h_text = re.sub(r"^#+\s*", "", clean_line).strip()
                        h_text = re.sub(r"\bstatic\s+", "", h_text, flags=re.IGNORECASE).rstrip(":").strip()
                        curr_section = {"heading": h_text, "items": [], "points": []}
                        parsed["sections"].append(curr_section)
                        continue

                # Regular bullet point
                if curr_section is None:
                    curr_section = {"heading": "Key Highlights", "items": [], "points": []}
                    parsed["sections"].append(curr_section)

                if indent_len >= 4:
                    raw_lvl = 2
                elif indent_len >= 2:
                    raw_lvl = 1
                else:
                    raw_lvl = 0

                if raw_lvl > prev_bullet_lvl + 1:
                    lvl = prev_bullet_lvl + 1
                else:
                    lvl = raw_lvl
                prev_bullet_lvl = lvl

                if ":" in pt_clean:
                    lbl, val = pt_clean.split(":", 1)
                else:
                    lbl, val = "", pt_clean

                lbl_str = lbl.strip()
                val_str = val.strip()

                if lbl_str.lower().startswith("static "):
                    lbl_str = lbl_str[7:].strip()

                curr_section["items"].append({
                    "type": "point",
                    "level": lvl,
                    "label": lbl_str,
                    "text": val_str
                })
                curr_section["points"].append({"level": lvl, "label": lbl_str, "text": val_str})

        # Backwards compatibility populating dynamic_heading and key_points
        if parsed["sections"]:
            parsed["dynamic_heading"] = parsed["sections"][0].get("heading", "Key Highlights:") + ":"
            all_pts = []
            for s in parsed["sections"]:
                for pt in s.get("points", []):
                    all_pts.append((pt.get("label", ""), pt.get("text", "")))
            parsed["key_points"] = all_pts

        # Format and bulletize why_in_news if lengthy
        lead, bullets = format_why_in_news(parsed["why_in_news"], parsed["why_in_news_bullets"])
        parsed["why_in_news"] = lead
        parsed["why_in_news_bullets"] = bullets

        return parsed


# ==============================================================================
# DOCUMENT BUILDER (MASTER DOCX, TXT & EXCEL GENERATORS)
# ==============================================================================

class DPRCGDocumentBuilder:
    """Generates the consolidated master Word, Text, and Excel reports."""

    @staticmethod
    def _safe_save_file(save_callback, target_path: str) -> str:
        """Collision protection: Handles locked files by saving as `(1).ext`."""
        try:
            save_callback(target_path)
            return target_path
        except PermissionError:
            path_obj = Path(target_path)
            for i in range(1, 10):
                fallback = str(path_obj.parent / f"{path_obj.stem} ({i}){path_obj.suffix}")
                try:
                    save_callback(fallback)
                    logger.warning("Original '%s' locked. Saved to '%s' instead.", target_path, fallback)
                    return fallback
                except PermissionError:
                    continue
            fallback = str(path_obj.parent / f"{path_obj.stem}_alt{path_obj.suffix}")
            save_callback(fallback)
            return fallback

    @staticmethod
    def _setup_bullet_numbering(doc: Document) -> int:
        """Injects default MS Word multi-level bullet numbering (Symbol small disc for lvl 0, Courier New hollow circle for lvl 1, Wingdings small square for lvl 2)."""
        numbering_part = doc.part.numbering_part
        numbering_elm = numbering_part._element

        abstract_nums = numbering_elm.findall(qn("w:abstractNum"))
        highest_abstract_id = 100
        for an in abstract_nums:
            an_id = int(an.get(qn("w:abstractNumId"), 0))
            if an_id > highest_abstract_id:
                highest_abstract_id = an_id

        nums = numbering_elm.findall(qn("w:num"))
        highest_num_id = 100
        for n in nums:
            n_id = int(n.get(qn("w:numId"), 0))
            if n_id > highest_num_id:
                highest_num_id = n_id

        new_abstract_id = highest_abstract_id + 1
        new_num_id = highest_num_id + 1

        abstract_num_xml = f"""
        <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="{new_abstract_id}">
          <w:multiLevelType w:val="hybridMultilevel"/>
          <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val="&#xF0B7;"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
              <w:ind w:left="280" w:hanging="220"/>
            </w:pPr>
            <w:rPr>
              <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
              <w:sz w:val="18"/>
            </w:rPr>
          </w:lvl>
          <w:lvl w:ilvl="1">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val="o"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
              <w:ind w:left="500" w:hanging="220"/>
            </w:pPr>
            <w:rPr>
              <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
              <w:sz w:val="16"/>
            </w:rPr>
          </w:lvl>
          <w:lvl w:ilvl="2">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val="&#xF0A7;"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
              <w:ind w:left="720" w:hanging="220"/>
            </w:pPr>
            <w:rPr>
              <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
              <w:sz w:val="16"/>
            </w:rPr>
          </w:lvl>
        </w:abstractNum>
        """
        abstract_elm = parse_xml(abstract_num_xml)
        numbering_elm.append(abstract_elm)

        num_xml = f"""
        <w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="{new_num_id}">
          <w:abstractNumId w:val="{new_abstract_id}"/>
        </w:num>
        """
        num_elm = parse_xml(num_xml)
        numbering_elm.append(num_elm)

        return new_num_id

    @staticmethod
    def _add_hyperlink(paragraph, url: str, text: str, color_hex: str = "0284C7", font_size_pt: float = 9.5):
        """Adds a clickable OpenXML hyperlink to a Word paragraph."""
        if not url:
            return
        try:
            part = paragraph.part
            r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
            sz_val = str(int(font_size_pt * 2))
            clean_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
            hyperlink_xml = (
                f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                f'r:id="{r_id}" '
                f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f'<w:r>'
                f'<w:rPr>'
                f'<w:color w:val="{color_hex}"/>'
                f'<w:u w:val="single"/>'
                f'<w:sz w:val="{sz_val}"/>'
                f'</w:rPr>'
                f'<w:t>{clean_text}</w:t>'
                f'</w:r>'
                f'</w:hyperlink>'
            )
            paragraph._p.append(parse_xml(hyperlink_xml))
        except Exception:
            run = paragraph.add_run(text)
            run.font.size = Pt(font_size_pt)
            run.font.color.rgb = RGBColor(2, 132, 199)
            run.font.underline = True

    @classmethod
    def _add_callout_box(cls, doc: Document, title: str, items: List[str]):
        """
        Renders high-importance exam takeaways in a single-cell callout box
        with a light background, left accent border, and bold labels.
        """
        if not items and not title:
            return

        tbl = doc.add_table(rows=1, cols=1)
        tbl.autofit = False

        tblPr = tbl._tbl.tblPr
        tblPr.append(parse_xml('<w:tblW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:w="5000" w:type="pct"/>'))
        tblPr.append(parse_xml('<w:tblCellMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tblCellMar>'))

        cell = tbl.cell(0, 0)
        tcPr = cell._tc.get_or_add_tcPr()

        shd_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="clear" w:color="auto" w:fill="F8FAFC"/>'
        borders_xml = '''<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:left w:val="single" w:sz="36" w:space="0" w:color="1E3A8A"/>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
          <w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
        </w:tcBorders>'''
        tcPr.append(parse_xml(shd_xml))
        tcPr.append(parse_xml(borders_xml))

        p_title = cell.paragraphs[0]
        p_title.paragraph_format.space_before = Pt(2)
        p_title.paragraph_format.space_after = Pt(4)
        clean_title = title.strip()
        if not clean_title.startswith("📌"):
            clean_title = f"📌 {clean_title}"
        r_t = p_title.add_run(clean_title)
        r_t.bold = True
        r_t.font.name = "Arial"
        r_t.font.size = Pt(10.0)
        r_t.font.color.rgb = RGBColor(30, 58, 138)

        for itm in items:
            p_itm = cell.add_paragraph()
            p_itm.paragraph_format.space_before = Pt(1)
            p_itm.paragraph_format.space_after = Pt(2.5)
            p_itm.paragraph_format.line_spacing = 1.15

            clean_itm = re.sub(r"^[-*•◦▪▫]\s*", "", str(itm)).strip()
            if ":" in clean_itm:
                parts = clean_itm.split(":", 1)
                lbl = parts[0].strip().replace("**", "")
                val = parts[1].strip()

                r_b = p_itm.add_run("• ")
                r_b.font.name = "Arial"
                r_b.font.size = Pt(9.5)
                r_b.font.color.rgb = RGBColor(30, 58, 138)

                r_lbl = p_itm.add_run(f"{lbl}: ")
                r_lbl.bold = True
                r_lbl.font.name = "Arial"
                r_lbl.font.size = Pt(9.5)
                r_lbl.font.color.rgb = RGBColor(15, 23, 42)

                r_val = p_itm.add_run(val)
                r_val.font.name = "Arial"
                r_val.font.size = Pt(9.5)
                r_val.font.color.rgb = RGBColor(51, 65, 85)
            else:
                r_b = p_itm.add_run("• ")
                r_b.font.name = "Arial"
                r_b.font.size = Pt(9.5)
                r_b.font.color.rgb = RGBColor(30, 58, 138)

                r_t = p_itm.add_run(clean_itm)
                r_t.font.name = "Arial"
                r_t.font.size = Pt(9.5)
                r_t.font.color.rgb = RGBColor(51, 65, 85)

        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    @classmethod
    def generate_master_docx(
        cls,
        articles: List[Dict],
        output_path: str,
        display_date: str
    ) -> str:
        """
        Builds the consolidated master Word document in the reverse-engineered reference format.
        Filename: DD-MM-YYYY.docx (e.g. 02-11-2026.docx)
        """
        doc = Document()
        bullet_num_id = cls._setup_bullet_numbering(doc)

        # Set 1-inch margins
        for section in doc.sections:
            section.top_margin = Inches(1.0)
            section.bottom_margin = Inches(1.0)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)

        # Date Header: DD-MM-YYYY
        hdr_p = doc.add_paragraph()
        hdr_p.paragraph_format.space_before = Pt(0)
        hdr_p.paragraph_format.space_after = Pt(14)
        hdr_run = hdr_p.add_run(display_date)
        hdr_run.font.name = "Arial"
        hdr_run.font.size = Pt(14)
        hdr_run.font.bold = True
        hdr_run.font.color.rgb = RGBColor(30, 41, 59)

        for idx, art in enumerate(articles):
            tag = art.get("tag", "CHHATTISGARH")
            urls = list(dict.fromkeys(art.get("urls") or ([art.get("url")] if art.get("url") else [])))
            title = art.get("title", "")
            why_in_news = art.get("why_in_news", "")
            sections = art.get("sections", [])
            if not sections and art.get("key_points"):
                sections = [{"heading": art.get("dynamic_heading", "Key Highlights:"), "points": art.get("key_points", [])}]
            read_more = art.get("read_more", "| Read More:        |")

            # 1. State / Category Tag
            tag_p = doc.add_paragraph()
            tag_p.paragraph_format.space_before = Pt(12)
            tag_p.paragraph_format.space_after = Pt(2)
            tag_run = tag_p.add_run(tag)
            tag_run.font.name = "Arial"
            tag_run.font.size = Pt(11)
            tag_run.font.bold = True
            tag_run.font.color.rgb = RGBColor(15, 118, 110)

            # 2. Source URLs: single link if 1 URL, bulleted list with hanging indent if multiple URLs
            if len(urls) <= 1:
                url_str = urls[0] if urls else ""
                if url_str:
                    target_url = shorten_dprcg_url(url_str)
                    disp_text = target_url
                    if len(disp_text) > 75:
                        disp_text = disp_text[:72] + "..."
                    url_p = doc.add_paragraph()
                    url_p.paragraph_format.space_after = Pt(4)
                    cls._add_hyperlink(url_p, target_url, disp_text, color_hex="0284C7", font_size_pt=9.5)
            else:
                for u in urls:
                    target_url = shorten_dprcg_url(u)
                    disp_text = target_url
                    if len(disp_text) > 75:
                        disp_text = disp_text[:72] + "..."

                    url_p = doc.add_paragraph()
                    url_p.paragraph_format.left_indent = Inches(0.20)
                    url_p.paragraph_format.first_line_indent = Inches(-0.20)
                    url_p.paragraph_format.space_before = Pt(1)
                    url_p.paragraph_format.space_after = Pt(1.5)

                    b_run = url_p.add_run("•\u00A0")
                    b_run.font.name = "Arial"
                    b_run.font.size = Pt(9.5)
                    b_run.font.color.rgb = RGBColor(100, 116, 139)

                    cls._add_hyperlink(url_p, target_url, disp_text, color_hex="0284C7", font_size_pt=9.5)
                url_p.paragraph_format.space_after = Pt(4)

            # 3. Main Article Title (Heading 2 Format, SEO/AEO/GEO Optimized)
            title_p = doc.add_paragraph(style="Heading 2")
            title_p.paragraph_format.space_before = Pt(6)
            title_p.paragraph_format.space_after = Pt(6)
            title_run = title_p.add_run(title)
            title_run.font.name = "Arial"
            title_run.font.size = Pt(13.5)
            title_run.font.bold = True
            title_run.font.color.rgb = RGBColor(15, 23, 42)

            # 4. "Why in News?" Section (Heading 3 Format)
            win_h = doc.add_paragraph(style="Heading 3")
            win_h.paragraph_format.space_before = Pt(6)
            win_h.paragraph_format.space_after = Pt(2)
            win_h_run = win_h.add_run("Why in News?")
            win_h_run.font.name = "Arial"
            win_h_run.font.size = Pt(11)
            win_h_run.font.bold = True
            win_h_run.font.color.rgb = RGBColor(30, 41, 59)

            win_lead, win_bullets = format_why_in_news(why_in_news, art.get("why_in_news_bullets", []))

            if win_lead:
                win_p = doc.add_paragraph()
                win_p.paragraph_format.space_after = Pt(3 if win_bullets else 8)
                win_p.paragraph_format.line_spacing = 1.15
                win_run = win_p.add_run(win_lead)
                win_run.font.name = "Arial"
                win_run.font.size = Pt(10.5)

            for b_text in win_bullets:
                bp = doc.add_paragraph()
                bp.paragraph_format.space_before = Pt(1.5)
                bp.paragraph_format.space_after = Pt(3)
                bp.paragraph_format.line_spacing = 1.15

                bpPr = bp._p.get_or_add_pPr()
                numPr_xml = f"""
                <w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:ilvl w:val="0"/>
                  <w:numId w:val="{bullet_num_id}"/>
                </w:numPr>
                """
                ind_xml = """
                <w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:left="280" w:hanging="220"/>
                """
                bpPr.append(parse_xml(numPr_xml))
                bpPr.append(parse_xml(ind_xml))

                if ":" in b_text:
                    parts = b_text.split(":", 1)
                    lbl = parts[0].strip().replace("**", "")
                    val = parts[1].strip()

                    lrun = bp.add_run(f"{lbl}: ")
                    lrun.bold = True
                    lrun.font.name = "Arial"
                    lrun.font.size = Pt(10.0)
                    lrun.font.color.rgb = RGBColor(15, 23, 42)

                    vrun = bp.add_run(val)
                    vrun.font.name = "Arial"
                    vrun.font.size = Pt(10.0)
                    vrun.font.color.rgb = RGBColor(51, 65, 85)
                else:
                    trun = bp.add_run(b_text)
                    trun.font.name = "Arial"
                    trun.font.size = Pt(10.0)
                    trun.font.color.rgb = RGBColor(51, 65, 85)

            # 5. Multiple Thematic Headings (H3) with Sub-themes (H4) & Nested Bullets
            for sec in sections:
                raw_heading = sec.get("heading", "Key Highlights")
                sec_heading = re.sub(r"\bstatic\s+", "", raw_heading, flags=re.IGNORECASE).strip()
                sec_heading = re.sub(r"^\s*context\s*:\s*", "", sec_heading, flags=re.IGNORECASE).strip()
                sec_heading = sec_heading.rstrip(":").strip()
                if sec_heading:
                    sec_heading = sec_heading[0].upper() + sec_heading[1:]
                if not sec_heading:
                    sec_heading = "Key Highlights"

                sec_h = doc.add_paragraph(style="Heading 3")
                sec_h.paragraph_format.space_before = Pt(8)
                sec_h.paragraph_format.space_after = Pt(4)
                sec_h_run = sec_h.add_run(sec_heading)
                sec_h_run.font.name = "Arial"
                sec_h_run.font.size = Pt(11)
                sec_h_run.font.bold = True
                sec_h_run.font.color.rgb = RGBColor(30, 41, 59)

                items = sec.get("items", [])
                if not items and sec.get("points"):
                    for pt in sec.get("points"):
                        if isinstance(pt, dict):
                            items.append({"type": "point", "level": pt.get("level", 0), "label": pt.get("label", ""), "text": pt.get("text") or pt.get("val", "")})
                        elif isinstance(pt, (list, tuple)):
                            items.append({"type": "point", "level": pt[2] if len(pt) >= 3 else 0, "label": pt[0], "text": pt[1]})
                        else:
                            items.append({"type": "point", "level": 0, "label": "", "text": str(pt)})

                for it in items:
                    if it.get("type") == "box":
                        cls._add_callout_box(doc, it.get("title", "Key Exam Takeaway"), it.get("items", []))
                        continue
                    elif it.get("type") == "subheading":
                        sub_h = doc.add_paragraph(style="Heading 4")
                        sub_h.paragraph_format.space_before = Pt(5)
                        sub_h.paragraph_format.space_after = Pt(2)
                        sub_run = sub_h.add_run(it.get("text", ""))
                        sub_run.font.name = "Arial"
                        sub_run.font.size = Pt(10.5)
                        sub_run.font.bold = True
                        sub_run.font.color.rgb = RGBColor(51, 65, 85)
                    else:
                        lvl = it.get("level", 0)
                        label = it.get("label", "").strip()
                        val = it.get("text") or it.get("val", "")
                        val = str(val).strip()

                        lbl_lower = label.lower()
                        val_lower = val.lower()
                        if ("static context" in lbl_lower or "static context" in val_lower or
                            "static background" in lbl_lower or "static background" in val_lower) and len(val_lower.split()) < 6:
                            continue
                        if lbl_lower in ["static context", "static background"]:
                            continue
                        if lbl_lower.startswith("static "):
                            label = label[7:].strip()

                        p = doc.add_paragraph()
                        p.paragraph_format.space_before = Pt(1.5)
                        p.paragraph_format.space_after = Pt(2.5 if lvl == 0 else 2.0)

                        left_ind = 280 + lvl * 220
                        pPr = p._p.get_or_add_pPr()
                        numPr_xml = f"""
                        <w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                          <w:ilvl w:val="{lvl}"/>
                          <w:numId w:val="{bullet_num_id}"/>
                        </w:numPr>
                        """
                        pPr.append(parse_xml(numPr_xml))

                        ind_xml = f"""
                        <w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:left="{left_ind}" w:hanging="220"/>
                        """
                        pPr.append(parse_xml(ind_xml))

                        font_size_pt = 10.0 if lvl == 0 else 9.5
                        label_color = RGBColor(15, 23, 42) if lvl == 0 else (RGBColor(30, 41, 59) if lvl == 1 else RGBColor(51, 65, 85))

                        if label:
                            l_run = p.add_run(f"{label}: ")
                            l_run.font.name = "Arial"
                            l_run.font.size = Pt(font_size_pt)
                            l_run.font.bold = True
                            l_run.font.color.rgb = label_color

                        v_run = p.add_run(val)
                        v_run.font.name = "Arial"
                        v_run.font.size = Pt(font_size_pt)

            # 6. "| Read More:        |" Footer (Left blank for user notes)
            rm_p = doc.add_paragraph()
            rm_p.paragraph_format.space_before = Pt(6)
            rm_p.paragraph_format.space_after = Pt(18)
            rm_run = rm_p.add_run("| Read More:        |")
            rm_run.font.name = "Arial"
            rm_run.font.size = Pt(9.5)
            rm_run.font.italic = True
            rm_run.font.color.rgb = RGBColor(100, 116, 139)

            # Divider between articles
            if idx < len(articles) - 1:
                sep_p = doc.add_paragraph()
                sep_p.paragraph_format.space_after = Pt(10)

        saved_path = cls._safe_save_file(doc.save, output_path)
        logger.info("Saved Master Current Affairs Word Report to: %s", saved_path)
        return saved_path

    @classmethod
    def generate_master_txt(
        cls,
        articles: List[Dict],
        output_path: str,
        display_date: str
    ) -> str:
        """Saves exact plain-text representation matching the reference document with strict H2/H3/H4 hierarchy."""
        blocks = [display_date]
        for art in articles:
            tag = art.get("tag", "CHHATTISGARH")
            urls = list(dict.fromkeys(art.get("urls") or ([art.get("url")] if art.get("url") else [])))
            title = art.get("title", "")
            win = art.get("why_in_news", "")
            sections = art.get("sections", [])

            clean_urls = []
            for u in urls:
                su = shorten_dprcg_url(u)
                if len(su) > 75:
                    su = su[:72] + "..."
                if su not in clean_urls:
                    clean_urls.append(su)

            lines = [tag]
            if len(clean_urls) <= 1:
                if clean_urls:
                    lines.append(clean_urls[0])
            else:
                lines.append("\n".join(f"• {u}" for u in clean_urls))

            win_lead, win_bullets = format_why_in_news(win, art.get("why_in_news_bullets", []))
            lines.append(f"## {title}")
            win_block = f"### Why in News?\n{win_lead}"
            if win_bullets:
                win_block += "\n" + "\n".join(f"• {b}" for b in win_bullets)
            lines.append(win_block)

            for sec in sections:
                raw_heading = sec.get("heading", "Key Highlights")
                heading = re.sub(r"\bstatic\s+", "", raw_heading, flags=re.IGNORECASE).strip()
                heading = re.sub(r"^\s*context\s*:\s*", "", heading, flags=re.IGNORECASE).rstrip(":").strip()
                if heading:
                    heading = heading[0].upper() + heading[1:]
                if not heading:
                    heading = "Key Highlights"

                lines.append(f"### {heading}")

                items = sec.get("items", [])
                if not items and sec.get("points"):
                    for pt in sec.get("points"):
                        if isinstance(pt, dict):
                            items.append({"type": "point", "level": pt.get("level", 0), "label": pt.get("label", ""), "text": pt.get("text") or pt.get("val", "")})
                        elif isinstance(pt, (list, tuple)):
                            items.append({"type": "point", "level": pt[2] if len(pt) >= 3 else 0, "label": pt[0], "text": pt[1]})
                        else:
                            items.append({"type": "point", "level": 0, "label": "", "text": str(pt)})

                pts_lines = []
                for it in items:
                    if it.get("type") == "box":
                        if pts_lines:
                            lines.append("\n".join(pts_lines))
                            pts_lines = []
                        b_title = it.get("title", "Key Exam Takeaway")
                        box_block = [f"[BOX: {b_title}]"]
                        for bi in it.get("items", []):
                            clean_bi = re.sub(r"^[-*•◦▪▫]\s*", "", str(bi)).strip()
                            box_block.append(f"• {clean_bi}")
                        box_block.append("[/BOX]")
                        lines.append("\n" + "\n".join(box_block) + "\n")
                        continue
                    elif it.get("type") == "subheading":
                        if pts_lines:
                            lines.append("\n".join(pts_lines))
                            pts_lines = []
                        lines.append(f"#### {it.get('text', '')}")
                    else:
                        lvl = it.get("level", 0)
                        lbl = it.get("label", "").strip()
                        val = str(it.get("text") or it.get("val", "")).strip()

                        lbl_lower = lbl.lower()
                        val_lower = val.lower()
                        if ("static context" in lbl_lower or "static context" in val_lower or
                            "static background" in lbl_lower or "static background" in val_lower) and len(val_lower.split()) < 6:
                            continue
                        if lbl_lower in ["static context", "static background"]:
                            continue
                        if lbl_lower.startswith("static "):
                            lbl = lbl[7:].strip()

                        bullet_char = "•" if lvl == 0 else ("◦" if lvl == 1 else "▪")
                        indent = "  " * lvl
                        if lbl:
                            pts_lines.append(f"{indent}{bullet_char} {lbl}: {val}")
                        else:
                            pts_lines.append(f"{indent}{bullet_char} {val}")

                if pts_lines:
                    lines.append("\n".join(pts_lines))

            lines.append("| Read More:        |")
            blocks.append("\n\n".join(lines))

        txt_output = "\n\n".join(blocks) + "\n"
        def save_txt(p):
            with open(p, "w", encoding="utf-8") as f:
                f.write(txt_output)
        saved_path = cls._safe_save_file(save_txt, output_path)
        logger.info("Saved Master Current Affairs Text file to: %s", saved_path)
        return saved_path

    @classmethod
    def generate_excel(cls, releases: List[Dict], output_path: str) -> str:
        """Generates an executive Excel workbook index of all scraped releases."""
        wb = openpyxl.Workbook()

        header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        border_thin = Border(
            left=Side(style="thin", color="CBD5E1"),
            right=Side(style="thin", color="CBD5E1"),
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="thin", color="CBD5E1")
        )

        relevant_items = [r for r in releases if r.get("is_relevant")]

        # Sheet 1: CGPSC High-Yield Releases
        ws1 = wb.active
        ws1.title = "CGPSC High-Yield Briefing"
        ws1.views.sheetView[0].showGridLines = True

        headers1 = ["S.No", "CGPSC Paper", "English Headline", "Summary & Key Takeaways", "Keywords", "Source Link"]
        ws1.append(headers1)
        for col_idx in range(1, len(headers1) + 1):
            cell = ws1.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = border_thin
        ws1.row_dimensions[1].height = 28

        for idx, item in enumerate(relevant_items, start=1):
            paper_info = item.get("paper", "CGPSC")
            headline = item.get("headline", "")
            summary = item.get("summary") or item.get("relevance_reason", "")
            keywords = ", ".join(item.get("matched_keywords", [])[:4])
            url = item.get("url", "")

            row_data = [idx, paper_info, headline, summary, keywords, "Read DPRCG Article"]
            ws1.append(row_data)
            curr_row = ws1.max_row
            ws1.row_dimensions[curr_row].height = 42

            for c_idx in range(1, len(row_data) + 1):
                cell = ws1.cell(row=curr_row, column=c_idx)
                cell.border = border_thin
                cell.alignment = align_left if c_idx in (3, 4) else align_center

            link_cell = ws1.cell(row=curr_row, column=6)
            link_cell.hyperlink = url
            link_cell.font = Font(name="Calibri", size=10, color="0284C7", underline="single")

        col_widths1 = [8, 18, 48, 60, 25, 20]
        for idx, w in enumerate(col_widths1, start=1):
            ws1.column_dimensions[get_column_letter(idx)].width = w

        # Sheet 2: All Scraped DPRCG Releases
        ws2 = wb.create_sheet(title="All DPRCG Releases")
        ws2.views.sheetView[0].showGridLines = True
        headers2 = ["S.No", "Post ID", "CGPSC Relevance", "English Headline", "Official URL"]
        ws2.append(headers2)
        for col_idx in range(1, len(headers2) + 1):
            cell = ws2.cell(row=1, column=col_idx)
            cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            cell.font = header_font
            cell.alignment = align_center
            cell.border = border_thin
        ws2.row_dimensions[1].height = 28

        for idx, item in enumerate(releases, start=1):
            rel_status = "CGPSC Relevant" if item.get("is_relevant") else "Routine / Local"
            row_data = [idx, item.get("id", ""), rel_status, item.get("headline", ""), "View Release"]
            ws2.append(row_data)
            curr_row = ws2.max_row
            ws2.row_dimensions[curr_row].height = 30

            for c_idx in range(1, len(row_data) + 1):
                cell = ws2.cell(row=curr_row, column=c_idx)
                cell.border = border_thin
                cell.alignment = align_left if c_idx == 4 else align_center

            link_cell = ws2.cell(row=curr_row, column=5)
            link_cell.hyperlink = item.get("url", "")
            link_cell.font = Font(name="Calibri", size=10, color="0284C7", underline="single")

        col_widths2 = [8, 14, 20, 70, 18]
        for idx, w in enumerate(col_widths2, start=1):
            ws2.column_dimensions[get_column_letter(idx)].width = w

        saved_path = cls._safe_save_file(wb.save, output_path)
        logger.info("Saved Excel index to: %s", saved_path)
        return saved_path


# ==============================================================================
# PIPELINE ORCHESTRATOR & CLI INTERFACE
# ==============================================================================

def run_dprcg_pipeline(target_date: Optional[date] = None, no_llm: bool = False):
    target = target_date or date.today()
    manifest = DPRCG_OUTPUT_BASE / target.isoformat() / "run_status.json"
    with RunStatus(manifest, target) as status:
        return _run_dprcg_pipeline(target, no_llm, status)


def _run_dprcg_pipeline(target_date, no_llm, status):
    """Executes the complete fetch -> classify -> LLM synthesize -> master document generation."""
    target = target_date if target_date else date.today()
    date_code = target.strftime("%Y-%m-%d")
    display_date = target.strftime("%d-%m-%Y")

    print("\n" + "=" * 70)
    print(f"  DPRCG CHHATTISGARH STATE CURRENT AFFAIRS (LLM POWERED): {display_date}")
    print("=" * 70)

    scraper = DPRCGScraper()
    releases, formatted_date = scraper.fetch_date_releases(target_date=target)

    if not releases:
        raise NoReleasesError(f"No DPRCG releases published for {formatted_date}; no reports generated.")

    # Step 1: Concurrently enrich articles with body text
    enriched = scraper.enrich_all(releases)
    missing = [str(r["id"]) for r in enriched if not r.get("body_text", "").strip()]
    status.update(total_releases=len(enriched), failed_article_ids=missing)
    if missing:
        raise IncompleteRunError(f"Article extraction failed for {len(missing)} DPRCG releases; withholding full bulletin.")

    # Step 2: CGPSC Syllabus Classification & High-Yield Filtering
    classifier = CGPSCClassifier()
    processed = classifier.process_all(enriched)

    relevant_items = [r for r in processed if r.get("is_relevant")]
    routine_items = [r for r in processed if not r.get("is_relevant")]

    print("\n" + "=" * 70)
    print(f"  EXTRACTED ARTICLES SUMMARY ({len(processed)} Total | {len(relevant_items)} High-Yield CGPSC)")
    print("=" * 70)

    for idx, item in enumerate(relevant_items, start=1):
        paper = item.get("paper", "CGPSC")
        headline = item.get("headline", "")
        url = item.get("url", "")
        print(f"[{idx}] [{paper}] {headline}")
        print(f"    URL: {url}")

    # Step 3: Date-wise Organized Output Directory: output/dprcg/YYYY-MM-DD/
    date_output_dir = DPRCG_OUTPUT_BASE / date_code
    date_output_dir.mkdir(parents=True, exist_ok=True)

    master_docx_path = str(date_output_dir / f"{display_date}.docx")
    master_txt_path = str(date_output_dir / f"{display_date}.txt")
    excel_path = str(date_output_dir / f"DPRCG_CGPSC_Daily_{date_code}.xlsx")

    # Step 4: Semantic Clustering & LLM Synthesis with Gemini 3.8 Flash
    synthesized_articles = []
    failed_topics = []
    clusters = []
    if not no_llm and relevant_items:
        print("\n" + "=" * 70)
        print(f"  CLUSTERING & CLUBBING DPRCG RELEASES ({len(relevant_items)} Relevant Items)")
        print("=" * 70)
        clusterer = TopicClusterer()
        clusters = clusterer.cluster_releases(relevant_items)
        print(f"[*] Grouped into {len(clusters)} consolidated topic clusters.")
        for i, c in enumerate(clusters, 1):
            rel_count = len(c.get("releases", []))
            dropped_count = len(c.get("dropped_ids", []))
            drop_info = f", {dropped_count} redundant dropped" if dropped_count else ""
            print(f"    [{i}] {c.get('topic', 'Topic')} ({rel_count} source releases{drop_info})")

        print("\n" + "=" * 70)
        print(f"  RUNNING GEMINI LLM CURRENT AFFAIRS SYNTHESIS ({len(clusters)} Consolidated Articles)")
        print("=" * 70)
        synthesizer = GeminiCurrentAffairsSynthesizer()
        synthesized_articles = synthesizer.synthesize_clusters(clusters)
        failed_topics = synthesizer.failed_topics
        status.update(expected_topics=len(clusters), completed_topics=len(synthesized_articles),
                      failed_topics=failed_topics)
        if len(synthesized_articles) != len(clusters) and not failed_topics:
            raise IncompleteRunError("Synthesis topic coverage does not match the cluster count.")
        if failed_topics:
            master_docx_path = str(date_output_dir / f"{display_date}_PARTIAL.docx")
            master_txt_path = str(date_output_dir / f"{display_date}_PARTIAL.txt")
            display_date += f" - PARTIAL: {len(synthesized_articles)}/{len(clusters)} topics; see run_status.json for missing topics"

    # Step 5: Generate Master Output Files in Reverse-Engineered Reference Structure
    if synthesized_articles:
        saved_docx = DPRCGDocumentBuilder.generate_master_docx(synthesized_articles, master_docx_path, display_date)
        saved_txt = DPRCGDocumentBuilder.generate_master_txt(synthesized_articles, master_txt_path, display_date)
        print(f"\n[+] Master Word Report: {saved_docx}")
        print(f"[+] Master Text File:   {saved_txt}")
    else:
        saved_docx = None
        saved_txt = None
        print("[!] No synthesized articles generated.")

    # Generate Excel Index
    saved_excel = DPRCGDocumentBuilder.generate_excel(processed, excel_path)
    print(f"[+] Excel Summary Index: {saved_excel}")

    files = {"xlsx": saved_excel}
    if saved_docx:
        files.update(docx=saved_docx, txt=saved_txt)
    if not all(path and Path(path).is_file() and Path(path).stat().st_size > 0 for path in files.values()):
        raise IncompleteRunError("Required DPRCG artifacts were not generated.")
    if failed_topics:
        status.update(status="partial", files=files)
        raise IncompleteRunError(f"DPRCG bulletin incomplete: {len(failed_topics)} topics failed. Partial artifacts are labelled PARTIAL.")
    outcome = "index_only" if no_llm or not relevant_items else "complete"
    status.update(status=outcome, files=files, expected_topics=len(clusters),
                  completed_topics=len(synthesized_articles), failed_topics=[])
    print(f"\n[*] DPRCG run status: {outcome}.")
    return saved_docx, saved_txt, saved_excel


def main():
    parser = argparse.ArgumentParser(description="Fetch, classify, and synthesize DPRCG daily news for CGPSC.")
    parser.add_argument(
        "--date",
        type=str,
        default=None,
        help="Date to scrape in YYYY-MM-DD format (defaults to today; no automatic fallback)."
    )
    parser.add_argument(
        "--no-llm",
        action="store_true",
        help="Skip LLM synthesis and only generate the Excel report."
    )

    args = parser.parse_args()

    target_date = None
    if args.date:
        try:
            target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print(f"[Error] Invalid date format '{args.date}'. Expected YYYY-MM-DD.")
            sys.exit(1)

    try:
        run_dprcg_pipeline(target_date=target_date, no_llm=args.no_llm)
    except RuntimeError as exc:
        logger.error("DPRCG run failed: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
