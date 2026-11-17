"""Conservative event grouping; shared places or communities never imply an event."""

import re
from article_selection import contains, normalize


SUBJECTS = {
    "Janjatiya Gaurav Divas": ["janjatiya gaurav", "जनजातीय गौरव", "tribal pride day"],
    "Paddy Procurement": ["paddy procurement", "धान खरीदी", "धान उपार्जन", "dhan kharidi"],
    "Mahtari Vandan": ["mahtari vandan", "mahatari vandan", "महतारी वंदन"],
    "PM Surya Ghar": ["surya ghar", "सूर्य घर"],
    "Vishnu Ke Paati": ["vishnu ke paati", "विष्णु के पाती"],
}
ACTION_GROUPS = {
    "enforcement": ["seizure", "seizes", "raid", "raids", "smuggling", "जब्त", "छापा", "तस्करी"],
    "policy": ["policy", "approval", "approves", "guidelines", "नीति", "मंजूरी", "दिशानिर्देश"],
    "payment": ["payment", "instalment", "installment", "transfers", "disburses", "भुगतान", "किस्त"],
    "launch": ["launch", "launches", "launched", "begins", "commences", "शुभारंभ", "शुरू"],
    "review": ["reviews", "review", "inspection", "inspects", "समीक्षा", "निरीक्षण"],
    "celebration": ["celebrates", "celebration", "observes", "commemorates", "मनाया", "आयोजन"],
}


def event_key(release):
    # Generated labels and incidental body mentions cannot establish identity.
    title = release.get("raw_title") or release.get("headline", "")
    subjects = [key for key, terms in SUBJECTS.items() if any(contains(title, t) for t in terms)]
    actions = [key for key, terms in ACTION_GROUPS.items() if any(contains(title, t) for t in terms)]
    # Unknown, ambiguous, or undated stories remain independent.
    if len(subjects) != 1 or len(actions) != 1 or not release.get("date"):
        return None
    # Explicit event/campaign dates distinguish stories published on the same day.
    dates = tuple(sorted(set(re.findall(
        r"\b\d{4}-\d{2}-\d{2}\b|\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b|"
        r"\b\d{1,2}\s+(?:Jan\w*|Feb\w*|Mar\w*|Apr\w*|May|Jun\w*|Jul\w*|Aug\w*|Sep\w*|Oct\w*|Nov\w*|Dec\w*)\s+\d{4}\b|"
        r"\d{1,2}\s+(?:जनवरी|फरवरी|मार्च|अप्रैल|मई|जून|जुलाई|अगस्त|सितंबर|अक्टूबर|नवंबर|दिसंबर)\s+\d{4}|"
        r"\b20\d{2}(?:-\d{2,4})?\b", title, flags=re.IGNORECASE))))
    return subjects[0], actions[0], str(release["date"]), dates


def safe_clusters(releases):
    """Partition input IDs once and drop text only for exact source duplicates."""
    groups, seen_ids = {}, set()
    for release in releases:
        rid = str(release.get("id", ""))
        if not rid:
            raise ValueError("Clustering requires a nonempty release ID.")
        if rid in seen_ids:
            continue
        seen_ids.add(rid)
        key = event_key(release) or ("independent", rid)
        groups.setdefault(key, []).append(release)

    result = []
    for key, members in groups.items():
        fingerprints, dropped = set(), []
        for release in members:
            body = normalize(release.get("body_text", ""))
            fingerprint = (normalize(release.get("raw_title") or release.get("headline", "")),
                           body, normalize(release.get("district", "")))
            if body and fingerprint in fingerprints:
                dropped.append(str(release["id"]))
            fingerprints.add(fingerprint)
        first = members[0]
        result.append({
            "topic": (f"{key[0]}: {key[1]} ({key[2]})" if key[0] != "independent"
                      else first.get("headline") or first.get("raw_title", "")),
            "paper": first.get("paper", "CGPSC-PRE"),
            "paper_title": first.get("paper_title", ""),
            "releases": members,
            "urls": list(dict.fromkeys(r["url"] for r in members if r.get("url"))),
            "dropped_ids": dropped,
        })
    return result
