"""Configuration settings and constants for PIB UPSC Automation."""

import os
from pathlib import Path

# Paths
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# PIB Endpoints
PIB_BASE_URL = "https://www.pib.gov.in"
PIB_ALL_REL_URL = "https://www.pib.gov.in/allRel.aspx?reg=48&lang=1"
PIB_DETAIL_URL = "https://www.pib.gov.in/PressReleasePage.aspx?PRID={prid}"

# Request Headers
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

# Network settings
REQUEST_TIMEOUT = 25  # seconds
MAX_RETRIES = 3
RETRY_BACKOFF = 2  # multiplier in seconds

# UPSC Syllabus Definitions
UPSC_PAPERS = {
    "GS-1": {
        "title": "General Studies 1",
        "subtopics": "Indian Heritage, Culture, History, Society & World Geography",
        "badge_color": "FEF3C7",  # Soft Amber
        "keywords": [
            "heritage", "monument", "archaeology", "excavation", "craft", "temple",
            "tribal", "folk art", "classical dance", "classical language", "unesco",
            "freedom struggle", "commemoration", "tribute", "birth anniversary", "death anniversary",
            "monsoon", "cyclone", "glacier", "earthquake", "geological", "western ghats",
            "himalayas", "river basin", "census", "demographic", "urbanization", "folk"
        ]
    },
    "GS-2": {
        "title": "General Studies 2",
        "subtopics": "Governance, Constitution, Polity, Social Justice & International Relations",
        "badge_color": "DBEAFE",  # Soft Blue
        "keywords": [
            "cabinet approves", "bill", "amendment", "act", "ordinance", "constitution",
            "fundamental rights", "judiciary", "high court", "supreme court", "election commission",
            "parliament", "lok sabha", "rajya sabha", "statutory", "quasi-judicial", "tribunal",
            "governance", "transparency", "e-governance", "citizen charter", "civil service",
            "welfare scheme", "vulnerable sections", "sc/st", "minorities", "women empowerment",
            "health policy", "education policy", "nep", "ayushman bharat", "poverty", "malnutrition",
            "treaty", "mou", "bilateral", "summit", "g20", "asean", "quad", "brics", "sco",
            "unsc", "wto", "who", "extradition", "diplomatic", "external affairs"
        ]
    },
    "GS-3": {
        "title": "General Studies 3",
        "subtopics": "Economy, Agriculture, Science & Tech, Environment & Internal Security",
        "badge_color": "DCFCE7",  # Soft Green
        "keywords": [
            "economy", "gdp", "growth", "inflation", "cpi", "wpi", "rbi", "monetary policy",
            "fiscal", "budget", "direct tax", "indirect tax", "gst", "customs", "export", "import",
            "msme", "industrial", "infrastructure", "railways", "national highway", "port",
            "sagarmala", "bharatmala", "aviation", "logistics", "fdi", "investment",
            "agriculture", "farmer", "crop", "msp", "procurement", "pm-kisan", "irrigation",
            "fertilizer", "organic farming", "food processing", "animal husbandry", "fisheries",
            "science and technology", "isro", "space", "drdo", "defence manufacturing", "missile",
            "satellite", "atomic energy", "nuclear", "ai", "artificial intelligence", "semiconductor",
            "quantum", "biotechnology", "supercomputer", "cyber", "cybersecurity",
            "environment", "climate change", "cop", "renewable energy", "solar", "green hydrogen",
            "biodiversity", "wildlife", "forest", "sanctuary", "national park", "tiger reserve",
            "iucn", "great indian bustard", "pollution", "air quality", "caqm", "water conservation",
            "disaster management", "ndrf", "ndma",
            "internal security", "border management", "dri", "narcotics", "smuggling",
            "money laundering", "ed", "nia", "cbi", "coastal security"
        ]
    },
    "GS-4": {
        "title": "General Studies 4",
        "subtopics": "Ethics, Integrity, Probity & Administrative Reforms",
        "badge_color": "F3E8FF",  # Soft Purple
        "keywords": [
            "ethics", "integrity", "probity", "vigilance", "anti-corruption", "cvc",
            "transparency", "accountability", "whistleblower", "code of conduct",
            "public service values", "compassion", "empathy in governance", "moral"
        ]
    }
}

# Negative noise patterns (routine ceremonies, sports congratulations, condolence messages)
NOISE_PATTERNS = [
    r"\bgreets\b", r"\bgreetings\b", r"\bcondoles\b", r"\bpasses away\b",
    r"\bcongratulates\b", r"\bcongratulated\b", r"\bwon silver\b", r"\bwon gold\b",
    r"\bwon bronze\b", r"\bcongratulated on winning\b", r"\bannual day celebration\b",
    r"\bbook launch\b", r"\bwishes on\b", r"\bwishing on\b", r"\bcondolences\b",
    r"\binaugurates exhibition\b", r"\battends dinner\b", r"\bwelcomes\b"
]
