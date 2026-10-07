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
            "freedom struggle", "geological survey",
            "monsoon", "cyclone", "glacier", "earthquake", "geological", "western ghats",
            "himalayas", "river basin", "census", "demographic", "urbanization", "folk tradition"
        ]
    },
    "GS-2": {
        "title": "General Studies 2",
        "subtopics": "Governance, Constitution, Polity, Social Justice & International Relations",
        "badge_color": "DBEAFE",  # Soft Blue
        "keywords": [
            "cabinet approves", "bill", "amendment", "ordinance", "constitution",
            "fundamental rights", "judiciary", "high court", "supreme court", "election commission",
            "parliament", "lok sabha", "rajya sabha", "statutory body", "quasi-judicial", "tribunal",
            "governance", "transparency", "e-governance", "citizen charter", "civil services",
            "chief justice of india", "supreme court judge", "attorney general", "solicitor general",
            "comptroller and auditor general", "cag", "election commissioner", "chief election commissioner",
            "upsc chairman", "finance commission", "law commission", "central vigilance commission", "cvc",
            "central information commission", "cic", "national human rights commission", "nhrc",
            "cabinet secretary", "foreign secretary", "national security advisor", "nsa", "inter-state council",
            "welfare scheme", "vulnerable sections", "scheduled castes", "scheduled tribes", "women empowerment",
            "national health policy", "ayushman bharat", "national education policy", "poverty alleviation", "malnutrition",
            "bilateral treaty", "mou signed", "bilateral relations", "international summit",
            "g20", "asean", "quad", "brics", "shanghai cooperation", "unsc", "wto", "extradition",
            "ministry of external affairs", "foreign policy"
        ]
    },
    "GS-3": {
        "title": "General Studies 3",
        "subtopics": "Economy, Agriculture, Science & Tech, Environment & Internal Security",
        "badge_color": "DCFCE7",  # Soft Green
        "keywords": [
            "economy", "gdp", "economic growth", "inflation", "cpi", "wpi", "reserve bank", "monetary policy",
            "fiscal deficit", "union budget", "direct tax", "indirect tax", "gst", "customs duty", "export promotion", "import substitution",
            "msme", "industrial corridor", "infrastructure", "railways network", "national highway", "inland waterway", "port modern",
            "sagarmala", "bharatmala", "aviation sector", "logistics policy", "foreign direct investment",
            "agriculture", "farmer", "crop production", "minimum support price", "procurement", "pm-kisan", "micro irrigation",
            "fertilizer subsidy", "organic farming", "food processing", "animal husbandry", "fisheries sector",
            "science and technology", "isro", "space exploration", "drdo", "defence manufacturing", "missile system",
            "satellite launch", "atomic energy", "nuclear reactor", "uranium", "artificial intelligence", "semiconductor mission",
            "quantum mission", "biotechnology", "supercomputer", "cybersecurity",
            "black hole", "binary black hole", "astronomy", "astrophysics", "cosmic", "gravitational waves",
            "telescope", "space telescope", "radio telescope", "giant metrewave radio telescope", "gmrt", "astrosat",
            "exoplanet", "dark matter", "dark energy", "supernova", "galaxy cluster", "deep space", "space science",
            "scientific breakthrough", "department of science and technology", "dst", "csir", "iisc", "tifr",
            "raman research institute", "deep ocean mission", "nanotechnology",
            "environment conservation", "climate change", "renewable energy", "solar energy", "green hydrogen",
            "biodiversity", "wildlife sanctuary", "forest cover", "national park", "tiger reserve", "endangered species",
            "great indian bustard", "pollution control", "air quality", "caqm", "water conservation",
            "disaster management", "ndrf", "ndma",
            "internal security", "border management", "narcotics control", "illicit drug trafficking",
            "money laundering", "enforcement directorate", "national investigation agency", "coastal security",
            "chief of the air staff", "chief of air staff", "chief of army staff", "chief of naval staff",
            "chief of defence staff", "cds", "air marshal", "general", "admiral",
            "indian air force", "indian navy", "indian army", "armed forces", "armed forces tribunal",
            "sarex", "search and rescue", "military exercise", "joint military exercise", "naval exercise",
            "bilateral exercise", "indian coast guard", "coast guard", "maritime security", "anti-piracy",
            "air defence", "indigenous warship", "aircraft carrier", "patrol vessel", "border roads organisation",
            "bro", "tejas", "ins", "iaf", "tri-service", "theatre command"
        ]
    },
    "GS-4": {
        "title": "General Studies 4",
        "subtopics": "Ethics, Integrity, Probity & Administrative Reforms",
        "badge_color": "F3E8FF",  # Soft Purple
        "keywords": [
            "ethics in governance", "integrity pact", "probity in public life", "vigilance awareness",
            "central vigilance commission", "anti-corruption", "whistleblower protection", "code of conduct",
            "public service values", "compassion in administration", "transparency in governance"
        ]
    }
}

# Negative noise patterns (routine ceremonies, quotes, poems, sports, condolences, protocol)
NOISE_PATTERNS = [
    r"\bsubhashitam\b",
    r"\bshares\s+(?:sanskrit\s+)?(?:subhashitam|shloka|quote|poem|glimpses|pictures|video|thoughts|reflections|post)\b",
    r"\bshares\s+a\s+(?:poem|quote|message|glimpse)\b",
    r"\binspiring\s+thoughts\b",
    r"\bgreets\b",
    r"\bgreetings\b",
    r"\bwishes\s+(?:on|the\s+people)\b",
    r"\bcongratulates\b",
    r"\bcongratulated\b",
    r"\bcongratulated\s+on\s+winning\b",
    r"\bwon\s+(?:gold|silver|bronze|medal|match|tournament|championship|trophy)\b",
    r"\basian\s+games\b",
    r"\bolympics\b",
    r"\bcondoles\b",
    r"\bcondolences\b",
    r"\bpasses\s+away\b",
    r"\bexpresses\s+(?:grief|sorrow|sadness)\b",
    r"\bmourns\s+the\s+(?:demise|passing|loss)\b",
    r"\bpays\s+(?:floral\s+)?tributes?\b",
    r"\bpays\s+homage\s+to\s+freedom\s+fighter\b",
    r"\bpays\s+homage\b",
    r"\bcalls\s+on\b",
    r"\bcalls\s+upon\b",
    r"\bannual\s+day\s+celebration\b",
    r"\bbook\s+launch\b",
    r"\binaugurates\s+exhibition\b",
    r"\battends\s+dinner\b",
    r"\bwelcomes\b"
]
