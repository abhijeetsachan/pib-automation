"""UPSC Syllabus Classifier & Alignment Engine.
Categorizes PIB releases into General Studies Papers (GS-1, GS-2, GS-3, GS-4)
and eliminates ceremonial / administrative noise.
"""

import re
import os
import logging
from typing import Dict, List, Optional
from config import UPSC_PAPERS, NOISE_PATTERNS

logger = logging.getLogger(__name__)


class UPSCClassifier:
    def __init__(self):
        self.noise_regexes = [re.compile(p, re.IGNORECASE) for p in NOISE_PATTERNS]
        self.gemini_api_key = os.environ.get("GEMINI_API_KEY", "").strip()

    def is_noise(self, headline: str) -> bool:
        """Identifies purely ceremonial, sports result, or routine protocol releases."""
        for pattern in self.noise_regexes:
            if pattern.search(headline):
                return True
        return False

    def classify_release(self, headline: str, ministry: str, article_snippet: str = "") -> Dict:
        """
        Classifies a single release into UPSC General Studies papers.
        
        Returns:
            Dict containing:
                - is_relevant (bool)
                - paper (str: GS-1, GS-2, GS-3, GS-4, or Routine)
                - paper_title (str)
                - relevance_reason (str)
                - tags (List[str])
                - score (int: 1-10)
        """
        # Step 1: Filter out obvious ceremonial noise
        if self.is_noise(headline):
            return {
                "is_relevant": False,
                "paper": "Routine/Ceremonial",
                "paper_title": "Routine Announcements / Sports / Greetings",
                "relevance_reason": "Ceremonial announcement, greeting, or sports update with no direct syllabus overlap.",
                "tags": ["Routine/Protocol"],
                "score": 1
            }

        combined_text = f"{headline} {ministry} {article_snippet}".lower()

        # Step 2: Match against UPSC taxonomy keywords
        paper_matches = {}
        for paper_code, paper_info in UPSC_PAPERS.items():
            matched_kw = [kw for kw in paper_info["keywords"] if kw in combined_text]
            if matched_kw:
                paper_matches[paper_code] = matched_kw

        # If keywords matched across papers, pick the paper with the strongest match count
        if paper_matches:
            best_paper = max(paper_matches.keys(), key=lambda p: len(paper_matches[p]))
            matched_keywords = paper_matches[best_paper]
            relevance_score = min(10, 5 + len(matched_keywords) * 2)

            return {
                "is_relevant": True,
                "paper": best_paper,
                "paper_title": UPSC_PAPERS[best_paper]["title"],
                "relevance_reason": f"Aligined with {UPSC_PAPERS[best_paper]['title']} topics: {', '.join(matched_keywords[:3])}",
                "tags": matched_keywords[:4],
                "score": relevance_score
            }

        # Step 3: Nodal Strategic Ministries Fallback
        # If no explicit keyword was found, but release is from a key strategic ministry
        strategic_ministries = {
            "ministry of defence": ("GS-3", "Internal & External Security / Defence Capabilities"),
            "ministry of external affairs": ("GS-2", "Bilateral & Regional International Relations"),
            "ministry of finance": ("GS-3", "Economic Governance & Fiscal Policy"),
            "ministry of environment, forest and climate change": ("GS-3", "Ecology, Climate Action & Conservation"),
            "department of atomic energy": ("GS-3", "Science, Nuclear Energy & Strategic Tech"),
            "ministry of law and justice": ("GS-2", "Constitutional Provisions & Legal Framework"),
            "ministry of personnel, public grievances & pensions": ("GS-2", "Administrative Reforms & Governance Quality"),
            "ministry of science and technology": ("GS-3", "Indigenization of Technology & R&D"),
        }

        min_lower = ministry.lower()
        for strat_min, (paper, subtopic) in strategic_ministries.items():
            if strat_min in min_lower:
                return {
                    "is_relevant": True,
                    "paper": paper,
                    "paper_title": UPSC_PAPERS[paper]["title"],
                    "relevance_reason": f"Strategic policy update from nodal ministry ({ministry}): {subtopic}",
                    "tags": [subtopic.split("/")[0].strip()],
                    "score": 6
                }

        # Step 4: General administrative fallback (Non-core)
        return {
            "is_relevant": False,
            "paper": "General/Administrative",
            "paper_title": "Routine Administrative Release",
            "relevance_reason": "General departmental update with low direct question probability in Civil Services.",
            "tags": ["General"],
            "score": 2
        }

    def process_all(self, releases: List[Dict]) -> List[Dict]:
        """Processes and enriches a list of releases with UPSC classification metadata."""
        enriched = []
        for item in releases:
            classification = self.classify_release(
                headline=item.get("headline", ""),
                ministry=item.get("ministry", ""),
                article_snippet=""
            )
            enriched_item = {**item, **classification}
            enriched.append(enriched_item)
        return enriched
