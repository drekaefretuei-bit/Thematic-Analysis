from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


# =========================================================
# USER SETTINGS
# =========================================================
TXT_DIR = Path("SLR Text Articles")

OUTPUT_XLSX = Path("appendix_iii_with_quotes_TEXT.xlsx")
OUTPUT_CSV = Path("appendix_iii_with_quotes_TEXT.csv")

SUMMARY_XLSX = Path("appendix_iii_summary_with_quotes_TEXT.xlsx")
SUMMARY_CSV = Path("appendix_iii_summary_with_quotes_TEXT.csv")


# =========================================================
# CONFIGURATION
# =========================================================
SECTION_HEADERS = {
    "abstract": ["abstract"],
    "introduction": ["introduction", "background", "literature review"],
    "methodology": [
        "methodology", "methods", "method", "research design",
        "data and methodology", "data and methods", "materials and methods",
        "empirical design", "research method",
    ],
    "sample": [
        "sample", "sample characteristics", "data", "dataset", "research context",
    ],
    "results": [
        "results", "findings", "discussion", "results and discussion",
        "empirical results", "key findings", "analysis and results",
    ],
    "conclusion": [
        "conclusion", "conclusions", "implications", "concluding remarks",
    ],
}

STUDY_DESIGN_PATTERNS = {
    "Panel data study": [
        r"\bpanel data\b", r"\bpanel regression\b", r"\bfirm-years?\b",
        r"\bcompany-years?\b", r"\blongitudinal\b",
    ],
    "Regression analysis": [
        r"\bregression analysis\b", r"\bmultiple regression\b",
        r"\bols\b", r"\borderinary least squares\b",
    ],
    "Fixed-effects model": [r"\bfixed effects?\b"],
    "Random-effects model": [r"\brandom effects?\b"],
    "Logistic regression": [r"\blogistic regression\b", r"\blogit\b", r"\bprobit\b"],
    "Difference-in-differences": [r"\bdifference[- ]in[- ]differences\b", r"\bdid\b"],
    "Content analysis": [
        r"\bcontent analysis\b", r"\bmanual coding\b", r"\bqualitative content analysis\b",
    ],
    "Automated text analysis": [
        r"\btext analysis\b", r"\bautomated coding\b", r"\bmachine learning\b",
        r"\btopic model(?:ing)?\b", r"\blda\b", r"\bnatural language processing\b",
    ],
    "Case study": [r"\bcase study\b"],
    "Survey study": [r"\bsurvey\b", r"\bquestionnaire\b"],
    "Interview study": [r"\binterviews?\b", r"\bsemi-structured interviews?\b"],
    "Mixed-methods study": [r"\bmixed methods?\b", r"\bmixed-methods?\b"],
    "Qualitative study": [r"\bqualitative\b"],
    "Quantitative study": [r"\bquantitative\b"],
    "Cross-sectional study": [r"\bcross-sectional\b"],
}

METHOD_TERMS = [
    "panel data", "regression analysis", "fixed effects", "random effects",
    "logistic regression", "ols", "ordinary least squares",
    "difference-in-differences", "did", "content analysis", "text analysis",
    "manual coding", "automated coding", "machine learning", "topic modeling",
    "lda", "case study", "survey", "interview", "mixed methods",
    "qualitative", "quantitative", "event study", "poisson",
    "negative binomial", "structural equation modeling", "sem",
    "hierarchical regression", "multilevel", "longitudinal", "cross-sectional",
]

BOARD_DIVERSITY_TERMS = [
    "gender diversity", "female directors", "women on board", "women directors",
    "board gender diversity", "ethnic diversity", "racial diversity",
    "board ethnic diversity", "board nationality diversity",
    "board demographic diversity", "minority directors",
    "intersectionality", "intersectional",
]

OUTCOME_TERMS = [
    "narrative reporting", "narrative disclosure", "corporate disclosure",
    "integrated reporting", "sustainability reporting", "csr disclosure",
    "esg disclosure", "annual report", "forward-looking disclosure",
    "readability", "tone", "impression management",
    "non-financial disclosure", "voluntary disclosure",
]

RESEARCH_FOCUS_TERMS = BOARD_DIVERSITY_TERMS + OUTCOME_TERMS + [
    "board diversity", "disclosure", "reporting",
]

RESULT_HINTS = [
    "results", "findings", "discussion", "conclusion",
    "empirical results", "key findings",
]

INTERSECTION_A = ["gender", "female", "women", "woman"]
INTERSECTION_B = ["ethnic", "racial", "minority", "nationality", "race", "ethnicity"]

FINDING_TERMS = [
    "significant", "positively", "negative", "negatively", "associated",
    "relationship", "effect", "influence", "impact", "increase", "decrease",
    "improves", "reduces", "higher", "lower", "suggest", "show", "demonstrate",
    "indicate", "evidence", "support", "reveal",
]

THEORY_PATTERNS = {
    "Agency theory": [
        r"\bagency theory\b",
        r"\bagency perspective\b",
        r"\bagency theoretic\b",
    ],
    "Stakeholder theory": [
        r"\bstakeholder theory\b",
        r"\bstakeholder perspective\b",
    ],
    "Legitimacy theory": [
        r"\blegitimacy theory\b",
        r"\blegitimation theory\b",
    ],
    "Institutional theory": [
        r"\binstitutional theory\b",
        r"\binstitutional perspective\b",
        r"\bnew institutionalism\b",
    ],
    "Resource dependence theory": [
        r"\bresource dependence theory\b",
        r"\bresource-dependence theory\b",
        r"\brdt\b",
    ],
    "Upper echelons theory": [
        r"\bupper echelons theory\b",
        r"\bupper-echelon theory\b",
    ],
    "Signaling theory": [
        r"\bsignaling theory\b",
        r"\bsignalling theory\b",
    ],
    "Critical mass theory": [
        r"\bcritical mass theory\b",
    ],
    "Social identity theory": [
        r"\bsocial identity theory\b",
    ],
    "Tokenism theory": [
        r"\btokenism\b",
        r"\btoken theory\b",
    ],
    "Feminist theory": [
        r"\bfeminist theory\b",
        r"\bfeminist perspective\b",
    ],
    "Intersectionality theory": [
        r"\bintersectionality theory\b",
        r"\bintersectionality perspective\b",
        r"\bintersectional framework\b",
        r"\bintersectionality framework\b",
    ],
    "Diversity theory": [
        r"\bdiversity theory\b",
        r"\bcognitive diversity\b",
    ],
    "Human capital theory": [
        r"\bhuman capital theory\b",
    ],
    "Social capital theory": [
        r"\bsocial capital theory\b",
    ],
    "Stewardship theory": [
        r"\bstewardship theory\b",
    ],
    "Institutional logics": [
        r"\binstitutional logics\b",
        r"\blogics perspective\b",
    ],
    "Neo-institutional theory": [
        r"\bneo-institutional theory\b",
        r"\bneoinstitutional theory\b",
    ],
    "Impression management theory": [
        r"\bimpression management theory\b",
        r"\bimpression management perspective\b",
    ],
    "Role congruity theory": [
        r"\brole congruity theory\b",
    ],
    "Social role theory": [
        r"\bsocial role theory\b",
    ],
    "Status characteristics theory": [
        r"\bstatus characteristics theory\b",
    ],
    "Behavioral theory": [
        r"\bbehavioral theory\b",
        r"\bbehavioural theory\b",
    ],
    "Contingency theory": [
        r"\bcontingency theory\b",
    ],
    "Political economy theory": [
        r"\bpolitical economy theory\b",
    ],
    "Organizational theory": [
        r"\borganizational theory\b",
        r"\borganisation theory\b",
    ],
}

THEORY_HINT_TERMS = [
    "theory", "theoretical", "perspective", "framework", "lens", "grounded in",
    "drawing on", "based on", "underpinned by", "guided by", "anchored in",
]

DATE_RANGE_PATTERNS = [
    r"\b(?:19|20)\d{2}\s*(?:to|-|–|—)\s*(?:19|20)\d{2}\b",
    r"\bfrom\s+(?:fiscal\s+years?\s+)?((?:19|20)\d{2})\s+(?:to|until|through)\s+((?:19|20)\d{2})\b",
    r"\bduring\s+((?:19|20)\d{2})\s*(?:to|-|–|—)\s*((?:19|20)\d{2})\b",
    r"\bcovering\s+((?:19|20)\d{2})\s*(?:to|-|–|—)\s*((?:19|20)\d{2})\b",
    r"\bfor the period\s+((?:19|20)\d{2})\s*(?:to|-|–|—)\s*((?:19|20)\d{2})\b",
    r"\bover the period\s+((?:19|20)\d{2})\s*(?:to|-|–|—)\s*((?:19|20)\d{2})\b",
]

SAMPLE_PATTERNS = [
    r"\b(?:sample of|consists of|comprises|includes|using)\s+(\d{2,6})\s+(firms|companies|observations|bank-years|firm-years|reports|annual reports|articles|boards|listed firms|listed companies)\b",
    r"\b(\d{2,6})\s+(firms|companies|observations|bank-years|firm-years|reports|annual reports|articles|boards|listed firms|listed companies)\b",
]

COUNTRY_PATTERNS = [
    r"\bunited kingdom\b", r"\buk\b", r"\bu\.k\.\b",
    r"\bunited states\b", r"\busa\b", r"\bu\.s\.\b", r"\bus\b",
    r"\baustralia\b", r"\bcanada\b", r"\bchina\b", r"\bindia\b",
    r"\bsouth africa\b", r"\bnigeria\b", r"\bmalaysia\b", r"\bindonesia\b",
    r"\bfrance\b", r"\bgermany\b", r"\bitaly\b", r"\bspain\b",
    r"\beurope\b", r"\bglobal\b", r"\binternational\b",
]

INDEX_PATTERNS = [
    r"\bftse\b", r"\bs&p\b", r"\basx\b", r"\btsx\b", r"\bbse\b",
    r"\bnse\b", r"\bnyse\b", r"\bnasdaq\b", r"\bstoxx\b",
]

CURRENT_STUDY_FINDING_PATTERNS = [
    r"\bour findings?\b",
    r"\bour results?\b",
    r"\bwe find\b",
    r"\bwe found\b",
    r"\bwe show\b",
    r"\bwe demonstrate\b",
    r"\bwe document\b",
    r"\bwe observe\b",
    r"\bwe report\b",
    r"\bwe reveal\b",
    r"\bthe results? show\b",
    r"\bthe results? suggest\b",
    r"\bthe results? indicate\b",
    r"\bthe findings? show\b",
    r"\bthe findings? suggest\b",
    r"\bthe findings? indicate\b",
    r"\bthis shows\b",
    r"\bthis study shows\b",
    r"\bthis study finds\b",
    r"\bthis study found\b",
    r"\bthis study suggests\b",
    r"\bthis study indicates\b",
    r"\bthe study shows\b",
    r"\bthe study finds\b",
    r"\bthe study found\b",
    r"\bthe evidence shows\b",
    r"\bthe evidence suggests\b",
    r"\bthe evidence indicates\b",
    r"\bour evidence shows\b",
    r"\bour evidence suggests\b",
    r"\bour analysis shows\b",
    r"\bour analysis suggests\b",
    r"\bthese results show\b",
    r"\bthese findings suggest\b",
    r"\bthe empirical results show\b",
    r"\bthe empirical findings suggest\b",
]

EXCLUDE_SECONDARY_PATTERNS = [
    r"\bprevious studies\b",
    r"\bprior studies\b",
    r"\bprior research\b",
    r"\bexisting literature\b",
    r"\bthe literature\b",
    r"\bpast research\b",
    r"\bearlier studies\b",
    r"\baccording to\b",
    r"\bas noted by\b",
    r"\bas argued by\b",
    r"\bas reported by\b",
    r"\bet al\.\b",
]


# =========================================================
# TEXT UTILITIES
# =========================================================
def normalize_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = text.replace("\uf0b7", " ")
    text = text.replace("ﬁ", "fi").replace("ﬂ", "fl")
    text = re.sub(r"(?<=\w)-\s*\n\s*(?=\w)", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def clean_inline(text: str) -> str:
    text = normalize_text(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def shorten(text: str, max_chars: int = 1400) -> str:
    text = clean_inline(text)
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 3].rstrip() + "..."


def coalesce(*values: Optional[str]) -> str:
    for value in values:
        if value and str(value).strip():
            return str(value).strip()
    return "Not specified"


def unique_preserve_order(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for item in items:
        norm = item.strip().lower()
        if norm and norm not in seen:
            seen.add(norm)
            out.append(item.strip())
    return out


def sentence_tokenize(text: str) -> List[str]:
    text = clean_inline(text)
    if not text:
        return []
    parts = re.split(r"(?<=[.!?])\s+", text)
    return [p.strip() for p in parts if p.strip()]


def remove_references_section(text: str) -> str:
    match = re.search(r"(?im)^\s*references\s*$", text)
    if match:
        return text[:match.start()].strip()
    return text.strip()


# =========================================================
# TXT READING
# =========================================================
def extract_txt_pages(txt_path: Path) -> Tuple[List[Dict[str, object]], str]:
    try:
        raw_text = txt_path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        raw_text = txt_path.read_text(encoding="latin-1")

    raw_text = normalize_text(raw_text)

    if "\f" in raw_text:
        raw_pages = [normalize_text(p) for p in raw_text.split("\f") if p.strip()]
    else:
        parts = re.split(r"(?im)^\s*page\s+\d+\s*$", raw_text)
        raw_pages = [normalize_text(p) for p in parts if p.strip()]

    if not raw_pages:
        raw_pages = [raw_text]

    page_records = [{"page": i + 1, "text": page_text} for i, page_text in enumerate(raw_pages)]

    full_text = "\n\n".join(
        page["text"] for page in page_records if str(page["text"]).strip()
    ).strip()
    full_text = remove_references_section(full_text)

    trimmed_page_records: List[Dict[str, object]] = []
    references_found = False

    for record in page_records:
        page_num = int(record["page"])
        page_text = str(record["text"] or "")

        if references_found:
            break

        match = re.search(r"(?im)^\s*references\s*$", page_text)
        if match:
            kept_text = page_text[:match.start()].strip()
            if kept_text:
                trimmed_page_records.append({"page": page_num, "text": kept_text})
            references_found = True
        else:
            trimmed_page_records.append({"page": page_num, "text": page_text})

    return trimmed_page_records, full_text


def first_nonempty_title_candidate(text: str) -> Optional[str]:
    lines = [ln.strip() for ln in text.splitlines()[:40] if ln.strip()]
    for line in lines:
        if len(line) < 15:
            continue
        if re.match(r"^(doi|http|www\.|received|accepted|published)", line.lower()):
            continue
        if sum(ch.isalpha() for ch in line) < 10:
            continue
        return line
    return None


def get_title(txt_path: Path, full_text: str) -> str:
    return first_nonempty_title_candidate(full_text) or txt_path.stem.replace("_", " ")


# =========================================================
# SECTION SPLITTING
# =========================================================
def split_sections(text: str) -> Dict[str, str]:
    lines = text.splitlines()
    sections: Dict[str, List[str]] = {"front_matter": []}
    current = "front_matter"

    header_lookup = {}
    for canonical, variants in SECTION_HEADERS.items():
        for variant in variants:
            header_lookup[variant.lower()] = canonical

    for line in lines:
        stripped = line.strip().lower()
        canonical = header_lookup.get(stripped)
        if canonical:
            current = canonical
            sections.setdefault(current, [])
        else:
            sections.setdefault(current, []).append(line)

    return {k: normalize_text("\n".join(v)) for k, v in sections.items()}


# =========================================================
# EXTRACTION HELPERS
# =========================================================
def best_sentences(text: str, keywords: List[str], max_sentences: int = 3) -> List[str]:
    sentences = sentence_tokenize(text)
    scored: List[Tuple[int, int, str]] = []

    for sent in sentences:
        s_low = sent.lower()
        score = sum(1 for kw in keywords if kw.lower() in s_low)
        if score > 0:
            scored.append((score, len(sent), sent))

    scored.sort(key=lambda x: (-x[0], x[1]))
    candidates = [s for _, _, s in scored[: max_sentences * 4]]
    return unique_preserve_order(candidates)[:max_sentences]


def find_terms(text: str, terms: List[str]) -> List[str]:
    low = text.lower()
    hits = [t for t in terms if t.lower() in low]
    return unique_preserve_order(hits)


def detect_study_design(text: str) -> str:
    hits = []
    for label, patterns in STUDY_DESIGN_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, text, flags=re.I):
                hits.append(label)
                break
    return "; ".join(unique_preserve_order(hits)) if hits else "Not specified"


def extract_date_ranges(text: str) -> List[str]:
    matches = []

    for pattern in DATE_RANGE_PATTERNS:
        for m in re.finditer(pattern, text, flags=re.I):
            if m.groups():
                groups = [g for g in m.groups() if g]
                if len(groups) >= 2:
                    matches.append(f"{groups[0]} to {groups[1]}")
                else:
                    matches.append(m.group(0))
            else:
                matches.append(m.group(0))

    generic = re.findall(r"\b((?:19|20)\d{2})\s*(?:to|-|–|—)\s*((?:19|20)\d{2})\b", text)
    for a, b in generic:
        matches.append(f"{a} to {b}")

    return unique_preserve_order(matches)


def extract_sample_size(text: str) -> str:
    for pattern in SAMPLE_PATTERNS:
        m = re.search(pattern, text, flags=re.I)
        if m:
            number = m.group(1)
            unit = m.group(2) if len(m.groups()) > 1 else ""
            return f"{number} {unit}".strip()
    return "Not specified"


def normalize_geo_hit(hit: str) -> str:
    low = hit.lower().strip(" .")
    mapping = {
        "uk": "United Kingdom",
        "u.k": "United Kingdom",
        "u.k.": "United Kingdom",
        "united kingdom": "United Kingdom",
        "us": "United States",
        "u.s": "United States",
        "u.s.": "United States",
        "usa": "United States",
        "united states": "United States",
        "ftse": "FTSE",
        "s&p": "S&P",
        "asx": "ASX",
        "tsx": "TSX",
        "bse": "BSE",
        "nse": "NSE",
        "nyse": "NYSE",
        "nasdaq": "NASDAQ",
        "stoxx": "STOXX",
    }
    return mapping.get(low, hit.title())


def extract_geography(text: str) -> str:
    hits = []

    for pattern in COUNTRY_PATTERNS + INDEX_PATTERNS:
        for m in re.finditer(pattern, text, flags=re.I):
            hits.append(normalize_geo_hit(m.group(0)))

    hits = unique_preserve_order(hits)
    return ", ".join(hits[:8]) if hits else "Not specified"


def extract_theory(sections: Dict[str, str], full_text: str) -> str:
    source_text = " ".join([
        sections.get("abstract", ""),
        sections.get("introduction", ""),
        sections.get("front_matter", ""),
        full_text[:8000],
    ]).strip()

    theory_hits = []
    evidence_sentences = []

    for theory_name, patterns in THEORY_PATTERNS.items():
        matched = False
        for pattern in patterns:
            if re.search(pattern, source_text, flags=re.I):
                matched = True
                break
        if matched:
            theory_hits.append(theory_name)

    theory_context_sents = best_sentences(
        source_text,
        list(THEORY_PATTERNS.keys()) + THEORY_HINT_TERMS,
        max_sentences=3,
    )

    if theory_hits:
        evidence_sentences.extend(theory_context_sents[:2])
        pieces = ["; ".join(unique_preserve_order(theory_hits))]
        if evidence_sentences:
            pieces.append("Evidence: " + " ".join(evidence_sentences))
        return shorten(" | ".join(pieces), 1200)

    # fallback for generic theory mentions
    generic_sents = best_sentences(source_text, THEORY_HINT_TERMS, max_sentences=2)
    if generic_sents:
        return shorten("Generic theory mention: " + " ".join(generic_sents), 1200)

    return "Not specified"


# =========================================================
# QUOTE EXTRACTION
# =========================================================
def page_has_result_hint(page_text: str) -> bool:
    low = page_text.lower()
    return any(hint in low for hint in RESULT_HINTS)


def has_current_study_finding_phrase(sentence: str) -> bool:
    return any(re.search(pattern, sentence, flags=re.I) for pattern in CURRENT_STUDY_FINDING_PATTERNS)


def has_secondary_citation_signal(sentence: str) -> bool:
    return any(re.search(pattern, sentence, flags=re.I) for pattern in EXCLUDE_SECONDARY_PATTERNS)


def quote_score(sentence: str, page_text: str) -> int:
    s = sentence.lower()
    score = 0

    has_a = any(term in s for term in INTERSECTION_A)
    has_b = any(term in s for term in INTERSECTION_B)

    if has_current_study_finding_phrase(sentence):
        score += 20
    if has_a and has_b:
        score += 10
    if "intersectional" in s or "intersectionality" in s:
        score += 10
    if any(term in s for term in BOARD_DIVERSITY_TERMS):
        score += 4
    if any(term in s for term in OUTCOME_TERMS):
        score += 4
    if any(term in s for term in FINDING_TERMS):
        score += 4
    if "disclosure" in s or "reporting" in s or "narrative" in s:
        score += 2
    if page_has_result_hint(page_text):
        score += 5
    if has_secondary_citation_signal(sentence):
        score -= 12

    sent_len = len(sentence)
    if 40 <= sent_len <= 350:
        score += 3
    elif sent_len > 500:
        score -= 3
    elif sent_len < 25:
        score -= 2

    return score


def build_section_page_records(
    page_records: List[Dict[str, object]],
    section_text: str,
    fallback_label: str,
) -> List[Dict[str, object]]:
    section_text = section_text.strip()
    if not section_text:
        return []

    section_sentences = set(sentence_tokenize(section_text))
    out: List[Dict[str, object]] = []

    for record in page_records:
        page_num = int(record["page"])
        page_text = str(record["text"] or "")
        matched_sentences = []

        for sent in sentence_tokenize(page_text):
            cleaned = clean_inline(sent)
            if cleaned in section_sentences:
                matched_sentences.append(cleaned)

        if matched_sentences:
            out.append({
                "page": page_num,
                "text": " ".join(matched_sentences),
                "section": fallback_label,
            })

    if not out:
        out = [{"page": 1, "text": section_text, "section": fallback_label}]

    return out


def extract_verbatim_quotes_with_pages(
    page_records: List[Dict[str, object]],
    sections: Dict[str, str],
    full_text: str,
    max_quotes: int = 10,
) -> List[Dict[str, object]]:
    preferred_records: List[Dict[str, object]] = []
    preferred_records.extend(build_section_page_records(page_records, sections.get("results", ""), "results"))
    preferred_records.extend(build_section_page_records(page_records, sections.get("conclusion", ""), "conclusion"))

    dedup_pref = []
    seen_pref = set()
    for rec in preferred_records:
        key = (rec["page"], clean_inline(str(rec["text"])))
        if key not in seen_pref and str(rec["text"]).strip():
            seen_pref.add(key)
            dedup_pref.append(rec)
    preferred_records = dedup_pref

    def collect_candidates(records: List[Dict[str, object]]) -> List[Dict[str, object]]:
        primary_candidates: List[Dict[str, object]] = []
        secondary_candidates: List[Dict[str, object]] = []

        for record in records:
            page_num = int(record["page"])
            page_text = str(record["text"] or "")
            if not page_text.strip():
                continue

            for sent in sentence_tokenize(page_text):
                sent_clean = clean_inline(sent)
                if len(sent_clean) < 25:
                    continue

                score = quote_score(sent_clean, page_text)
                candidate = {"quote": sent_clean, "page": page_num, "score": score}

                if has_current_study_finding_phrase(sent_clean) and score >= 12:
                    primary_candidates.append(candidate)
                elif score >= 10:
                    secondary_candidates.append(candidate)

        primary_candidates.sort(key=lambda x: (-int(x["score"]), len(str(x["quote"]))))
        secondary_candidates.sort(key=lambda x: (-int(x["score"]), len(str(x["quote"]))))
        return primary_candidates + secondary_candidates

    selected: List[Dict[str, object]] = []
    seen = set()

    for item in collect_candidates(preferred_records):
        norm = re.sub(r"\W+", "", str(item["quote"]).lower())
        if norm in seen:
            continue
        seen.add(norm)
        selected.append(item)
        if len(selected) >= max_quotes:
            return selected

    remaining_needed = max_quotes - len(selected)
    if remaining_needed > 0:
        fallback_candidates = collect_candidates(page_records)
        for item in fallback_candidates:
            norm = re.sub(r"\W+", "", str(item["quote"]).lower())
            if norm in seen:
                continue
            seen.add(norm)
            selected.append(item)
            if len(selected) >= max_quotes:
                return selected

    return selected


def format_quotes(quotes: List[Dict[str, object]]) -> str:
    if not quotes:
        return "Not specified"
    return "\n".join(
        f'{i}. "{item["quote"]}" (p. {item["page"]})'
        for i, item in enumerate(quotes, start=1)
    )


def format_quote_pages(quotes: List[Dict[str, object]]) -> str:
    if not quotes:
        return "Not specified"
    pages = unique_preserve_order([f"p. {item['page']}" for item in quotes])
    return "; ".join(pages)


# =========================================================
# FIELD BUILDERS
# =========================================================
def extract_research_focus(sections: Dict[str, str], full_text: str) -> str:
    source_text = " ".join([
        sections.get("abstract", ""),
        sections.get("introduction", ""),
        full_text[:3000],
    ]).strip()
    sents = best_sentences(source_text, RESEARCH_FOCUS_TERMS, max_sentences=2)
    return coalesce(shorten(" ".join(sents), 1000))


def extract_methodology(sections: Dict[str, str], full_text: str) -> str:
    source_text = " ".join([
        sections.get("methodology", ""),
        sections.get("sample", ""),
        full_text[:7000],
    ]).strip()

    design_sents = best_sentences(
        source_text,
        METHOD_TERMS + ["data", "sample", "period", "model", "analysis"],
        max_sentences=3,
    )
    term_hits = find_terms(source_text, METHOD_TERMS)
    period_hits = extract_date_ranges(source_text)

    pieces = []
    if design_sents:
        pieces.append(" ".join(design_sents))
    if term_hits:
        pieces.append("Methodological terms: " + ", ".join(term_hits[:10]))
    if period_hits:
        pieces.append("Detected period(s): " + "; ".join(period_hits[:4]))

    return coalesce(shorten(" | ".join(pieces), 1600))


def extract_sample_characteristics(sections: Dict[str, str], full_text: str) -> str:
    source_text = " ".join([
        sections.get("sample", ""),
        sections.get("methodology", ""),
        full_text[:7000],
    ]).strip()

    sents = best_sentences(
        source_text,
        ["sample", "firms", "companies", "observations", "listed", "country", "dataset"],
        max_sentences=3,
    )

    sample_size = extract_sample_size(source_text)
    geography = extract_geography(source_text)

    pieces = []
    if sample_size != "Not specified":
        pieces.append(f"Sample size: {sample_size}")
    if geography != "Not specified":
        pieces.append(f"Geography: {geography}")
    if sents:
        pieces.append(" ".join(sents))

    return coalesce(shorten(" | ".join(pieces), 1600))


def extract_key_variables(sections: Dict[str, str], full_text: str) -> str:
    source_text = " ".join([
        sections.get("abstract", ""),
        sections.get("introduction", ""),
        sections.get("methodology", ""),
        sections.get("results", ""),
    ]).strip()

    diversity_hits = find_terms(source_text, BOARD_DIVERSITY_TERMS)
    outcome_hits = find_terms(source_text, OUTCOME_TERMS)

    sents = best_sentences(
        source_text,
        BOARD_DIVERSITY_TERMS + OUTCOME_TERMS + [
            "measure", "variable", "dependent", "independent", "proxy", "index"
        ],
        max_sentences=3,
    )

    pieces = []
    if diversity_hits:
        pieces.append("Board-diversity measures: " + ", ".join(diversity_hits[:10]))
    if outcome_hits:
        pieces.append("Outcome definitions: " + ", ".join(outcome_hits[:10]))
    if sents:
        pieces.append(" ".join(sents))

    return coalesce(shorten(" | ".join(pieces), 1600))


def extract_key_findings_summary(sections: Dict[str, str], full_text: str) -> str:
    source_text = " ".join([
        sections.get("results", ""),
        sections.get("conclusion", ""),
        full_text,
    ]).strip()

    sents = best_sentences(
        source_text,
        INTERSECTION_A + INTERSECTION_B + [
            "intersectional", "intersectionality", "disclosure", "reporting"
        ] + FINDING_TERMS,
        max_sentences=3,
    )

    return coalesce(shorten(" ".join(sents), 1600))


def extract_data_collection_period(sections: Dict[str, str], full_text: str) -> str:
    source_text = " ".join([
        sections.get("methodology", ""),
        sections.get("sample", ""),
        full_text[:8000],
    ]).strip()

    periods = extract_date_ranges(source_text)
    return "; ".join(periods[:4]) if periods else "Not specified"


# =========================================================
# ROW BUILDER
# =========================================================
def build_row(txt_path: Path) -> Dict[str, str]:
    page_records, full_text = extract_txt_pages(txt_path)
    article_id = txt_path.stem

    if not full_text:
        return {
            "Article ID": article_id,
            "Study Title": txt_path.stem,
            "Research Focus": "Not specified",
            "Theory": "Not specified",
            "Study Design Type": "Not specified",
            "Data Collection Period": "Not specified",
            "Sample Size": "Not specified",
            "Geography": "Not specified",
            "Methodology": "Not specified",
            "Sample Characteristics": "Not specified",
            "Key Variables": "Not specified",
            "Key Findings Summary": "Not specified",
            "Verbatim Finding Quotes": "Not specified",
            "Quote Pages": "Not specified",
        }

    sections = split_sections(full_text)

    methodology_source = " ".join([
        sections.get("methodology", ""),
        sections.get("sample", ""),
        full_text[:7000],
    ]).strip()

    quotes = extract_verbatim_quotes_with_pages(
        page_records=page_records,
        sections=sections,
        full_text=full_text,
        max_quotes=10,
    )

    row = {
        "Article ID": article_id,
        "Study Title": get_title(txt_path, full_text),
        "Research Focus": extract_research_focus(sections, full_text),
        "Theory": extract_theory(sections, full_text),
        "Study Design Type": detect_study_design(methodology_source),
        "Data Collection Period": extract_data_collection_period(sections, full_text),
        "Sample Size": extract_sample_size(methodology_source),
        "Geography": extract_geography(methodology_source),
        "Methodology": extract_methodology(sections, full_text),
        "Sample Characteristics": extract_sample_characteristics(sections, full_text),
        "Key Variables": extract_key_variables(sections, full_text),
        "Key Findings Summary": extract_key_findings_summary(sections, full_text),
        "Verbatim Finding Quotes": format_quotes(quotes),
        "Quote Pages": format_quote_pages(quotes),
    }

    for key, value in row.items():
        if not str(value).strip():
            row[key] = "Not specified"

    return row


# =========================================================
# SAVE OUTPUTS
# =========================================================
def save_outputs(detailed_df: pd.DataFrame) -> None:
    detailed_df = detailed_df.fillna("Not specified")
    detailed_df = detailed_df.replace(r"^\s*$", "Not specified", regex=True)

    summary_df = detailed_df[
        [
            "Article ID",
            "Study Title",
            "Research Focus",
            "Theory",
            "Methodology",
            "Sample Characteristics",
            "Key Variables",
        ]
    ].copy()

    summary_df["Key Findings"] = detailed_df.apply(
        lambda row: row["Verbatim Finding Quotes"]
        if row["Verbatim Finding Quotes"] != "Not specified"
        else row["Key Findings Summary"],
        axis=1,
    )

    detailed_df.to_excel(OUTPUT_XLSX, index=False)
    detailed_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    summary_df.to_excel(SUMMARY_XLSX, index=False)
    summary_df.to_csv(SUMMARY_CSV, index=False, encoding="utf-8-sig")


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    if not TXT_DIR.exists():
        raise SystemExit(f"Folder not found: {TXT_DIR.resolve()}")

    txt_files = sorted(TXT_DIR.glob("*.txt"))
    if not txt_files:
        raise SystemExit(f"No TXT files found in: {TXT_DIR.resolve()}")

    rows = []
    for txt_file in txt_files:
        print(f"Processing: {txt_file.name}")
        try:
            rows.append(build_row(txt_file))
        except Exception as exc:
            print(f"  Error processing {txt_file.name}: {exc}")
            rows.append({
                "Article ID": txt_file.stem,
                "Study Title": txt_file.stem,
                "Research Focus": "Not specified",
                "Theory": "Not specified",
                "Study Design Type": "Not specified",
                "Data Collection Period": "Not specified",
                "Sample Size": "Not specified",
                "Geography": "Not specified",
                "Methodology": "Not specified",
                "Sample Characteristics": "Not specified",
                "Key Variables": "Not specified",
                "Key Findings Summary": "Not specified",
                "Verbatim Finding Quotes": "Not specified",
                "Quote Pages": "Not specified",
            })

    detailed_df = pd.DataFrame(rows, columns=[
        "Article ID",
        "Study Title",
        "Research Focus",
        "Theory",
        "Study Design Type",
        "Data Collection Period",
        "Sample Size",
        "Geography",
        "Methodology",
        "Sample Characteristics",
        "Key Variables",
        "Key Findings Summary",
        "Verbatim Finding Quotes",
        "Quote Pages",
    ])

    save_outputs(detailed_df)

    print("\nDone.")
    print(f"Detailed Excel saved to: {OUTPUT_XLSX.resolve()}")
    print(f"Detailed CSV saved to:   {OUTPUT_CSV.resolve()}")
    print(f"Summary Excel saved to:  {SUMMARY_XLSX.resolve()}")
    print(f"Summary CSV saved to:    {SUMMARY_CSV.resolve()}")

    print("\nPreview:")
    print(detailed_df.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
