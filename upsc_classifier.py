"""UPSC Syllabus Classifier & Alignment Engine.
Categorizes PIB releases into General Studies Papers (GS-1, GS-2, GS-3, GS-4)
and eliminates ceremonial, inspirational, social media, and administrative noise.
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

    def is_noise(self, headline: str, combined_text: str = "") -> bool:
        """
        Identifies ceremonial, quotes/poems/subhashitam, sports results,
        condolences, routine protocol releases, or generic stub headlines.
        """
        hl_clean = headline.strip()
        check_text = f"{headline} {combined_text}".strip()
        for pattern in self.noise_regexes:
            if pattern.search(hl_clean) or pattern.search(check_text):
                return True
        return False

    def classify_release(self, headline: str, ministry: str, article_snippet: str = "") -> Dict:
        """
        Classifies a single release into UPSC General Studies papers using word-boundary matching.
        
        Returns:
            Dict containing:
                - is_relevant (bool)
                - paper (str: GS-1, GS-2, GS-3, GS-4, or Routine)
                - paper_title (str)
                - relevance_reason (str)
                - tags (List[str])
                - score (int: 1-10)
        """
        combined_text = f"{headline} {ministry} {article_snippet}".lower()

        # Step 1: Strict noise filtering (Social media quotes, Subhashitam, Sports, Protocol)
        if self.is_noise(headline, combined_text):
            return {
                "is_relevant": False,
                "paper": "Routine/Ceremonial",
                "paper_title": "Routine Announcements / Quotes / Sports / Protocol",
                "relevance_reason": "Ceremonial announcement, quote sharing, greeting, or sports update with no direct syllabus overlap.",
                "tags": ["Filtered Noise"],
                "score": 1
            }

        # Step 2: Strict word-boundary matching against UPSC taxonomy keywords
        paper_matches = {}
        for paper_code, paper_info in UPSC_PAPERS.items():
            matched_kw = []
            for kw in paper_info["keywords"]:
                # Use regex \bword\b boundary to prevent partial substring false positives
                # (e.g. preventing 'ed' matching 'shared' or 'hearted')
                regex_pattern = r"\b" + re.escape(kw) + r"\b"
                if re.search(regex_pattern, combined_text, re.IGNORECASE):
                    matched_kw.append(kw)

            if matched_kw:
                paper_matches[paper_code] = matched_kw

        # If keywords matched across papers, pick the paper with the highest match count
        if paper_matches:
            best_paper = max(paper_matches.keys(), key=lambda p: len(paper_matches[p]))
            matched_keywords = paper_matches[best_paper]
            relevance_score = min(10, 5 + len(matched_keywords) * 2)

            return {
                "is_relevant": True,
                "paper": best_paper,
                "paper_title": UPSC_PAPERS[best_paper]["title"],
                "relevance_reason": f"Aligned with {UPSC_PAPERS[best_paper]['title']} topics: {', '.join(matched_keywords[:3])}",
                "tags": matched_keywords[:4],
                "score": relevance_score
            }

        # Step 3: Nodal Strategic Ministries Policy Fallback
        # Captures policy actions, military exercises/commissioning, and scientific discoveries
        policy_action_words = [
            # Policy, reform & governance initiatives
            r"\bapproves?\b", r"\blaunches?\b", r"\bpolicy\b", r"\bguidelines?\b",
            r"\bschemes?\b", r"\bmission\b", r"\breport\b", r"\bindex\b",
            r"\bagreement\b", r"\binitiative\b", r"\boperations?\b", r"\bmeasures?\b",
            # Military, Maritime & Security Operations / Exercises / Commissioning
            r"\bexercises?\b", r"\bcommissions?\b", r"\bflagged?\s+off\b", r"\binducts?\b",
            r"\bdrills?\b", r"\bdeployment\b", r"\binaugurates?\b", r"\bparticipates?\b",
            # Scientific discoveries, R&D & breakthroughs
            r"\bdiscovers?\b", r"\bdiscovery\b", r"\breveals?\b", r"\bbreakthrough\b",
            r"\bfindings?\b", r"\bidentifies?\b", r"\bdevelops?\b", r"\bunveils?\b",
            r"\bdetects?\b", r"\bresearchers?\b", r"\bastronomers?\b", r"\bstudy\b"
        ]
        has_policy_action = any(re.search(pat, combined_text) for pat in policy_action_words)

        if has_policy_action:
            strategic_ministries = {
                "ministry of defence": ("GS-3", "Internal & External Security / Defence Capabilities"),
                "ministry of external affairs": ("GS-2", "Bilateral & Regional International Relations"),
                "ministry of finance": ("GS-3", "Economic Governance & Fiscal Policy"),
                "ministry of environment, forest and climate change": ("GS-3", "Ecology, Climate Action & Conservation"),
                "department of atomic energy": ("GS-3", "Science, Nuclear Energy & Strategic Tech"),
                "department of space": ("GS-3", "Space Technology & Planetary Exploration"),
                "ministry of earth sciences": ("GS-1", "Oceanography, Meteorology & Earth Systems"),
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
        """
        Processes, enriches, and deduplicates releases.
        Eliminates duplicate entries across Cabinet and Line Ministries.
        """
        enriched = []
        seen_normalized_headlines = {}

        for item in releases:
            classification = self.classify_release(
                headline=item.get("headline", ""),
                ministry=item.get("ministry", ""),
                article_snippet=""
            )
            enriched_item = {**item, **classification}

            # Pillar 1: Deduplication of relevant releases
            if enriched_item.get("is_relevant"):
                raw_hl = item.get("headline", "")
                norm_hl = re.sub(r"[^a-zA-Z0-9]", "", raw_hl.lower())

                if norm_hl in seen_normalized_headlines:
                    prev_item = seen_normalized_headlines[norm_hl]
                    # If current item is from Cabinet, prioritize Cabinet over Line Ministry
                    if "cabinet" in item.get("ministry", "").lower() and "cabinet" not in prev_item.get("ministry", "").lower():
                        prev_item["is_relevant"] = False
                        prev_item["paper"] = "Routine/Duplicate"
                        prev_item["relevance_reason"] = f"Duplicate multi-ministry entry (superseded by Cabinet release PRID: {item.get('prid')})."
                        seen_normalized_headlines[norm_hl] = enriched_item
                    else:
                        enriched_item["is_relevant"] = False
                        enriched_item["paper"] = "Routine/Duplicate"
                        enriched_item["relevance_reason"] = f"Duplicate multi-ministry entry (already captured under {prev_item.get('ministry')}, PRID: {prev_item.get('prid')})."
                else:
                    seen_normalized_headlines[norm_hl] = enriched_item

            enriched.append(enriched_item)
        return enriched
