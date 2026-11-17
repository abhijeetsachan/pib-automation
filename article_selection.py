"""Shared, deterministic relevance signals for English and Hindi releases."""

import re
import unicodedata


def normalize(text):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", text or "")).strip().casefold()


def contains(text, term):
    text, term = normalize(text), normalize(term)
    # Devanagari vowel marks are not regex word characters. Word boundaries
    # therefore split Hindi words; use literal phrase matching for Hindi.
    if re.search(r"[\u0900-\u097f]", term):
        return term in text
    return bool(re.search(r"(?<!\w)" + re.escape(term) + r"(?!\w)", text))


GENERIC_TERMS = {
    "chief minister", "governor", "collector", "cabinet", "assembly",
    "government", "governance", "public administration", "bureaucracy",
    "scheme", "initiative", "report", "exhibition", "inaugurated",
    "foundation stone", "dignitary visit", "मुख्यमंत्री", "राज्यपाल",
    "कलेक्टर", "मंत्रिमंडल", "योजना", "शासन",
}

ACTION_PATTERN = re.compile(
    r"\b(?:approv(?:e[sd]?|al)|launch(?:es|ed)?|notifi(?:es|ed|cation)|"
    r"introduc(?:es?|ed)|amend(?:s|ed|ment)?|allocat(?:es?|ed|ion)|"
    r"releas(?:es?|ed)|publish(?:es|ed)?|discover(?:s|ed|y)?|"
    r"develop(?:s|ed)?|adopt(?:s|ed)?|expand(?:s|ed)?|"
    r"new|revised|increas(?:es?|ed)|implement(?:s|ed|ation)?)\b|"
    r"मंजूरी|स्वीकृत|शुभारंभ|शुभारम्भ|शुरू|लागू|अधिसूचित|संशोधन|"
    r"आवंटन|जारी|नई|नयी|नया|वृद्धि|खोज|विकसित|विस्तार",
    re.IGNORECASE,
)


def has_substantive_evidence(text, specific_terms):
    """Require a concrete subject and change in the same sentence."""
    return any(
        ACTION_PATTERN.search(sentence)
        and any(contains(sentence, term) for term in specific_terms)
        for sentence in re.split(r"[.!?।\n]+", text or "")
    )


def weighted_matches(headline, body, terms):
    """Count each distinct subject once; prioritize headline evidence."""
    matches = [term for term in dict.fromkeys(terms)
               if contains(headline, term) or contains(body, term)]
    specific = [term for term in matches if term not in GENERIC_TERMS]
    score = sum(3 if contains(headline, term) else 1 for term in specific)
    return matches, specific, score


CGPSC_HINDI_KEYWORDS = {
    "CGPSC-P3": ["संविधान", "पंचायत", "नगरपालिका", "विधानसभा", "अधिनियम",
                 "विधेयक", "निर्वाचन", "सुशासन", "ई-गवर्नेंस", "पुलिस सुधार"],
    "CGPSC-P4": ["पर्यावरण", "जैव विविधता", "वन्यजीव", "राष्ट्रीय उद्यान",
                 "अभयारण्य", "सौर ऊर्जा", "नवीकरणीय ऊर्जा", "जल संरक्षण",
                 "प्रदूषण", "जलवायु", "विज्ञान", "प्रौद्योगिकी"],
    "CGPSC-P5": ["औद्योगिक नीति", "अर्थव्यवस्था", "बजट", "निवेश", "खनन",
                 "कृषि", "किसान", "धान खरीदी", "धान उपार्जन", "सिंचाई",
                 "न्यूनतम समर्थन मूल्य", "कृषक उन्नति", "खाद्य प्रसंस्करण"],
    "CGPSC-P6": ["जनजातीय", "आदिवासी", "विशेष पिछड़ी जनजाति", "लोक नृत्य",
                 "लोककला", "पुरातत्व", "विरासत", "हस्तशिल्प", "बस्तर दशहरा"],
    "CGPSC-P7": ["छात्रवृत्ति", "शिक्षा", "स्वास्थ्य", "पोषण", "महतारी वंदन",
                 "स्वयं सहायता समूह", "स्व-सहायता समूह", "महिला सशक्तिकरण",
                 "वन अधिकार", "रोजगार", "सामाजिक सुरक्षा", "पेंशन"],
    "CGPSC-PRE": ["राष्ट्रीय पुरस्कार", "राज्य पुरस्कार", "पद्म श्री", "खेल"],
}

CGPSC_ROUTINE_TERMS = [
    "courtesy visit", "courtesy meeting", "met with", "condolence",
    "birthday greetings", "शिष्टाचार", "सौजन्य भेंट", "शोक", "श्रद्धांजलि",
    "जन्मदिन की बधाई", "जन्मदिन पर बधाई", "स्थानांतरण", "यातायात एडवाइजरी",
]
