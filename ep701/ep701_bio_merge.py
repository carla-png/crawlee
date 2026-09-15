#!/usr/bin/env python3
"""EP701 bio-fact merge and QA pipeline.

Inputs (same folder as this script unless overridden with --dir):
  EP701_personalised.csv   master file (701 firms)
  bios_v3.csv              fresh bio text (287 firms)

Outputs:
  EP701_personalised_v3.csv
  EP701_personalised_v3.xlsx   (All sheet + one sheet per tier, bold header, frozen header row)
  EP701_v3_report.md           (counts requested in the brief)

Rules implemented (from the brief):
  * Only bios_v3 rows with a blank error and bio_text > 300 characters are candidates.
  * One concrete professional fact is taken from bio_text only. Fact types:
    licensed since year, founded year, board certified, wrote a book, prior career,
    bar role, radio show. Nothing personal (family, kids, hometown, pets, health).
  * first_line: max 18 words, plain wording, no dashes, starts with Saw / Noticed / Read / Came across.
  * bio_fact max 12 words, bio_quote is an exact substring of bio_text, max 20 words.
  * Every bio_quote is re-checked verbatim against bio_text; failures are dropped.
  * Merge on domain. B_events_page_only -> D_bio_fact (old first_line kept in seminar_line).
    C_none -> D_bio_fact. A_dated_event and existing D_bio_fact rows are never touched.
  * QA over the whole file: word count 5..20, no dashes, opener, personal details,
    duplicate first_lines across domains, icp_flag for injury / mass tort / criminal firms.

Nothing is invented: every quote is sliced out of the bio text itself and the first_line
is built from a fixed template around the extracted value.
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

OPENERS = ("Saw", "Noticed", "Read", "Came across")
DASH_RE = re.compile(r"[-‐‑‒–—―]")
YEAR = r"(?:19[3-9]\d|20[0-2]\d)"

PERSONAL_WORDS = [
    "wife", "husband", "spouse", "partner in life", "married", "marriage", "children", "child",
    "kids", "kid", "son", "sons", "daughter", "daughters", "grandchild", "grandchildren",
    "grandkids", "grandson", "granddaughter", "mother", "father", "mom", "dad", "parents",
    "brother", "sister", "family", "grew up", "raised in", "born in", "born and raised",
    "native of", "hometown", "home town", "dog", "dogs", "cat", "cats", "puppy", "pets",
    "pet", "horse", "cancer", "survivor", "illness", "diagnosed", "disease", "surgery",
    "passed away", "hobbies", "hobby", "enjoys", "loves to", "avid", "golf",
    "fishing", "hiking", "skiing", "cycling", "gardening", "cooking", "traveling", "travelling",
    "church", "parish", "faith", "vacation", "his free time", "her free time", "spare time",
    "when not", "when he is not", "when she is not", "outside the office", "outside of work",
]
PERSONAL_RE = re.compile(r"\b(?:" + "|".join(re.escape(w).replace(r"\ ", r"\s+") for w in PERSONAL_WORDS) + r")\b", re.I)

# ---------------------------------------------------------------------------
# Fact patterns. Each entry: (fact_type, compiled regex, priority). Lower priority wins.
# Regex named groups feed the templates.
# ---------------------------------------------------------------------------
FACT_PATTERNS = [
    # licensed / practicing since year
    ("licensed_since", re.compile(
        rf"(?:licensed|admitted)\s+(?:to\s+(?:practice|the\s+bar)|to\s+practice\s+law)[^.;]{{0,80}}?\b(?:in|since)\s+(?P<year>{YEAR})\b",
        re.I), 1),
    ("licensed_since", re.compile(
        rf"(?:practicing|practising|practiced|practised)\s+(?:law\s+|estate\s+planning\s+)?(?:since|beginning\s+in)\s+(?P<year>{YEAR})\b",
        re.I), 1),
    ("licensed_since", re.compile(
        rf"(?:member\s+of\s+the\s+)?(?:[A-Z][\w.]*\s+){{0,3}}(?:State\s+)?Bar(?:\s+Association)?\s+since\s+(?P<year>{YEAR})\b",
        re.I), 1),
    ("licensed_since", re.compile(
        rf"(?:has\s+been\s+)?(?:an?\s+)?(?:attorney|lawyer)\s+since\s+(?P<year>{YEAR})\b",
        re.I), 1),
    # founded year
    ("founded", re.compile(
        rf"(?:founded|established|opened|started|launched|formed)\s+(?:the\s+firm|his\s+(?:own\s+)?(?:firm|practice)|her\s+(?:own\s+)?(?:firm|practice)|their\s+(?:firm|practice)|his\s+own\s+law\s+(?:firm|practice)|her\s+own\s+law\s+(?:firm|practice)|(?:[A-Z][\w&'.,-]*\s+){{1,6}}(?:Law|Legal|Firm|Group|Offices?|PLLC|LLC|P\.?C\.?|LLP|APC|Ltd\.?))[^.;]{{0,80}}?\b(?:in|since)\s+(?P<year>{YEAR})\b",
        re.I), 1),
    ("founded", re.compile(
        rf"(?:firm|practice)\s+(?:was\s+)?(?:founded|established|opened|formed)\s+(?:in|since)\s+(?P<year>{YEAR})\b",
        re.I), 1),
    ("founded", re.compile(
        rf"(?:serving|has\s+served)\s+(?:clients|families|the\s+community|[A-Z][\w\s]{{2,40}}?)\s+since\s+(?P<year>{YEAR})\b",
        re.I), 2),
    # board certified
    ("board_certified", re.compile(
        r"board[\s-]certified\s+(?:as\s+an?\s+|in\s+|by\s+the\s+|specialist\s+in\s+)(?P<spec>[A-Za-z][A-Za-z ,&/]{3,70}?)(?=[.,;:)]|\s+by\s+|\s+and\s+|\s+since|\s+which|\s+who|$)",
        re.I), 1),
    ("board_certified", re.compile(
        r"(?:is|was)\s+board[\s-]certified\b", re.I), 2),
    # wrote a book
    ("book", re.compile(
        r"(?:author|co-author|authored|co-authored|wrote|published)\s+(?:of\s+)?(?:the\s+book|a\s+book|books?|the\s+bestselling\s+book|the\s+best-selling\s+book)\s*[,:]?\s*[\"“‘']?(?P<title>[A-Z][^\"”’.,;]{3,80}?)[\"”’]?(?=[.,;:(]|\s+which|\s+that|\s+and|$)",
        re.I), 1),
    ("book", re.compile(
        r"(?:author|co-author|authored|co-authored|wrote|written|published)\s+(?:of\s+)?(?:the\s+book|a\s+book|several\s+books|two\s+books|three\s+books|books)\b",
        re.I), 2),
    # prior career
    ("prior_career", re.compile(
        r"(?:before|prior\s+to)\s+(?:becoming\s+an?\s+attorney|becoming\s+a\s+lawyer|entering\s+law|attending\s+law\s+school|law\s+school|practicing\s+law|his\s+legal\s+career|her\s+legal\s+career|joining\s+the\s+firm)[^.;]{0,20}?,?\s+(?:he|she|they|[A-Z][a-z]+)\s+(?:was|worked\s+as|served\s+as|spent\s+[^.;]{1,25}\s+as|had\s+a\s+career\s+as)\s+(?:an?\s+)?(?P<career>[A-Za-z][A-Za-z ,&/]{3,60}?)(?=[.,;:(]|\s+for\s+|\s+at\s+|\s+with\s+|\s+in\s+|\s+where|\s+and\s+|$)",
        re.I), 1),
    ("prior_career", re.compile(
        r"(?:worked|served)\s+as\s+(?:an?\s+)?(?P<career>[A-Za-z][A-Za-z ,&/]{3,60}?)\s+(?:before|prior\s+to)\s+(?:becoming|entering|attending|practicing|law\s+school|his\s+legal|her\s+legal|going\s+to\s+law)",
        re.I), 1),
    ("prior_career", re.compile(
        r"\b(?:a\s+)?former\s+(?P<career>(?:certified\s+public\s+accountant|CPA|registered\s+nurse|nurse|teacher|schoolteacher|engineer|police\s+officer|accountant|paralegal|financial\s+(?:advisor|planner)|pilot|banker|trust\s+officer|social\s+worker|firefighter|prosecutor|judge|judicial\s+clerk|law\s+clerk|professor|pharmacist|physician|doctor|stockbroker|insurance\s+agent|real\s+estate\s+(?:agent|broker)|journalist|military\s+officer|naval\s+officer|army\s+officer|marine))\b",
        re.I), 2),
    # bar role
    ("bar_role", re.compile(
        r"(?P<role>(?:past|former|current|immediate\s+past)?\s*(?:president|chair|chairman|chairwoman|chairperson|vice\s+president|vice\s+chair|secretary|treasurer|board\s+member|director|trustee|council\s+member|member\s+of\s+the\s+board))\s+of\s+(?:the\s+)?(?P<org>[A-Z][A-Za-z .&'-]{3,80}?(?:Bar\s+Association|Bar|Estate\s+Planning\s+Council|Law\s+Section|Section|Committee|Inn\s+of\s+Court))\b",
        re.I), 1),
    ("bar_role", re.compile(
        r"served\s+as\s+(?:the\s+)?(?P<role>chair|president|co-chair|vice\s+president|secretary|treasurer)\s+of\s+(?:the\s+)?(?P<org>[A-Z][A-Za-z .&'-]{3,80}?(?:Bar\s+Association|Bar|Estate\s+Planning\s+Council|Section|Committee|Inn\s+of\s+Court))\b",
        re.I), 1),
    ("bar_role", re.compile(
        r"(?P<role>chaired|chairs)\s+(?:the\s+)?(?P<org>[A-Z][A-Za-z .&'-]{3,80}?(?:Bar\s+Association|Bar|Estate\s+Planning\s+Council|Section|Committee|Inn\s+of\s+Court))\b",
        re.I), 1),
    # radio / podcast show
    ("radio", re.compile(
        r"(?:host|hosts|hosted|co-hosts?|co-hosted)\s+(?:of\s+)?(?:the\s+|a\s+|an?\s+)?(?:weekly\s+|popular\s+|long-running\s+|local\s+)?(?:[\"“]?(?P<show>[A-Z][^\"”.,;]{2,60}?)[\"”]?\s*,?\s*)?(?:a\s+)?(?:weekly\s+|monthly\s+|daily\s+)?(?P<kind>radio\s+(?:show|program|programme)|podcast|television\s+(?:show|program)|TV\s+(?:show|program))\b",
        re.I), 1),
    ("radio", re.compile(
        r"(?P<kind>radio\s+(?:show|program|programme)|podcast|television\s+show|TV\s+show)\b[^.;]{0,60}?\b(?:host|hosted|hosts)\b",
        re.I), 2),
]

ICP_TERMS = {
    "injury": ["personal injury", "car accidents?", "auto accidents?", "truck accidents?", "motorcycle accidents?",
               "slip and fall", "wrongful death", "catastrophic injur(?:y|ies)", "injury lawyers?", "injury attorneys?",
               "injured", "accident lawyers?", "accident attorneys?", "medical malpractice", "dog bites?",
               "workers'? comp(?:ensation)?", "premises liability"],
    "mass_tort": ["mass torts?", "class actions?", "product liability", "talcum", "roundup", "camp lejeune",
                  "mesothelioma", "asbestos", "hernia mesh", "zantac", "paraquat", "3m earplugs?",
                  "multidistrict", "mdl", "defective drugs?", "dangerous drugs?"],
    "criminal": ["criminal defen[cs]e", "dui", "dwi", "owi", "drug charges", "felony", "felonies",
                 "misdemeanors?", "expungement", "domestic violence", "sex crimes", "assault charges",
                 "theft charges", "criminal law", "arrested", "bail bonds?", "traffic tickets"],
}
ESTATE_TERMS = ["estate plans?", "estate planning", "trusts", "living trusts?", "revocable trusts?", "irrevocable trusts?",
                "trust administration", "wills", "last will", "probate", "elder law", "medicaid planning",
                "asset protection", "powers? of attorney", "guardianships?", "conservatorships?", "special needs",
                "estate administration", "succession planning", "legacy planning"]
ICP_RE = {k: re.compile(r"\b(?:" + "|".join(v) + r")\b", re.I) for k, v in ICP_TERMS.items()}
ESTATE_RE = re.compile(r"\b(?:" + "|".join(ESTATE_TERMS) + r")\b", re.I)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def norm_domain(v) -> str:
    s = "" if pd.isna(v) else str(v).strip().lower()
    s = re.sub(r"^https?://", "", s)
    s = re.sub(r"^www\.", "", s)
    s = s.split("/")[0].strip()
    return s


def blank(v) -> bool:
    return pd.isna(v) or str(v).strip() == "" or str(v).strip().lower() in {"nan", "none", "null"}


def words(s: str) -> list[str]:
    return [w for w in re.split(r"\s+", str(s).strip()) if w]


def has_dash(s: str) -> bool:
    return bool(DASH_RE.search(str(s)))


def personal_hit(s: str) -> str | None:
    m = PERSONAL_RE.search(str(s))
    return m.group(0).lower() if m else None


def clean_value(v: str) -> str:
    """Plain wording for template insertion: collapse whitespace, replace dashes with spaces."""
    v = DASH_RE.sub(" ", v)
    v = re.sub(r"\s+", " ", v).strip(" ,.;:")
    return v


def sentence_bounds(text: str, start: int, end: int) -> tuple[int, int]:
    """Expand [start, end) to the enclosing sentence."""
    s = start
    while s > 0 and text[s - 1] not in ".!?\n":
        s -= 1
    e = end
    while e < len(text) and text[e] not in ".!?\n":
        e += 1
    if e < len(text):
        e += 1  # keep the terminal punctuation
    return s, e


def make_quote(text: str, m: re.Match, max_words: int = 20) -> str:
    """Exact substring of text, at most max_words words, containing the match when possible."""
    s, e = sentence_bounds(text, m.start(), m.end())
    sent = text[s:e]
    if len(words(sent)) <= max_words:
        q = sent.strip()
        return q if q in text else text[s:e]
    # Sentence too long: take a window of max_words words that contains the match.
    # Work on word start offsets relative to the sentence.
    word_spans = [(mm.start() + s, mm.end() + s) for mm in re.finditer(r"\S+", sent)]
    # index of first word touching match start and last word touching match end
    first = next((i for i, (a, b) in enumerate(word_spans) if b > m.start()), 0)
    last = next((i for i, (a, b) in reversed(list(enumerate(word_spans))) if a < m.end()), len(word_spans) - 1)
    need = last - first + 1
    if need > max_words:
        # match itself is longer than the limit: take the first max_words words of the match
        last = first + max_words - 1
        need = max_words
    spare = max_words - need
    lo = max(0, first - spare // 2)
    hi = min(len(word_spans) - 1, lo + max_words - 1)
    lo = max(0, hi - max_words + 1)
    q = text[word_spans[lo][0]:word_spans[hi][1]]
    return q


def trim_words(s: str, n: int) -> str:
    w = words(s)
    return " ".join(w[:n])


def first_name(name) -> str:
    if blank(name):
        return ""
    n = re.sub(r"\b(?:Mr|Mrs|Ms|Dr|Esq|Attorney|Atty)\.?\s*", "", str(name)).strip()
    parts = n.split()
    return parts[0] if parts else ""


# ---------------------------------------------------------------------------
# fact extraction
# ---------------------------------------------------------------------------
def wrap(opener: str, clause: str, tail: str = "") -> str:
    """Grammatical wrapper for each allowed opener around a clause like 'you founded the firm in 2001'."""
    if opener == "Saw":
        core = f"Saw {clause}"
    elif opener == "Noticed":
        core = f"Noticed {clause}"
    elif opener == "Read":
        core = f"Read that {clause}"
    else:
        core = f"Came across your bio and saw {clause}"
    line = f"{core}, {tail}." if tail else f"{core}."
    if len(words(line)) > 18 or has_dash(line):
        line = f"{core}."
    if len(words(line)) > 18:
        line = f"{opener} {clause}." if opener != "Came across" else f"Came across your bio, {clause}."
    return DASH_RE.sub(" ", line)


def build_line(kind: str, gd: dict, opener_idx: int) -> tuple[str, str]:
    """Return (first_line, bio_fact) for a fact kind. Never more than 18 / 12 words."""
    year = gd.get("year")
    spec = clean_value(gd.get("spec") or "").lower()
    title = clean_value(gd.get("title") or "")
    career = clean_value(gd.get("career") or "").lower()
    role = clean_value(gd.get("role") or "").lower()
    if role in ("chaired", "chairs"):
        role = "chair"
    org = clean_value(gd.get("org") or "")
    show = clean_value(gd.get("show") or "")
    show_kind = clean_value(gd.get("kind") or "radio show").lower()
    if show_kind.startswith(("tv", "television")):
        show_kind = "TV show"
    elif show_kind != "podcast":
        show_kind = "radio show"

    op = OPENERS[opener_idx % len(OPENERS)]
    tails = ["that is a long run", "not many can say that", "that stood out", ""]
    tail = tails[(opener_idx // len(OPENERS)) % len(tails)]

    if kind == "licensed_since":
        clause, fact = f"you have been practicing law since {year}", f"Practicing law since {year}"
    elif kind == "founded":
        clause, fact = f"you founded the firm back in {year}", f"Firm founded in {year}"
        tail = {"that is a long run": "nice to see that longevity"}.get(tail, tail)
    elif kind == "board_certified":
        if spec:
            clause, fact = f"you are board certified in {spec}", f"Board certified in {spec}"
        else:
            clause, fact = "you are board certified", "Board certified attorney"
        tail = {"that is a long run": "not many attorneys can say that"}.get(tail, tail)
    elif kind == "book":
        if title:
            clause, fact = f"you wrote the book {title}", f"Author of {title}"
        else:
            clause, fact = "you have written a book", "Has authored a book"
        tail = {"that is a long run": "that is quite an undertaking"}.get(tail, tail)
    elif kind == "prior_career":
        if career:
            art = "an" if career[0] in "aeiou" and career != "cpa" else "a"
            clause, fact = f"you worked as {art} {career} before law", f"Former {career} before law"
        else:
            clause, fact = "you had a career before law", "Had a career before law"
        tail = {"that is a long run": "that is a useful background"}.get(tail, tail)
    elif kind == "bar_role":
        r = role or "a leadership role"
        if org:
            clause, fact = f"you served as {r} of the {org}", f"{r.capitalize()} of the {org}"
        else:
            clause, fact = f"you served as {r} in the bar", f"{r.capitalize()} in the bar"
        tail = {"that is a long run": "that is real involvement"}.get(tail, tail)
    elif kind == "radio":
        if show:
            clause, fact = f"you host the {show} {show_kind}", f"Hosts the {show} {show_kind}"
        else:
            clause, fact = f"you host a {show_kind}", f"Hosts a {show_kind}"
        tail = {"that is a long run": "a great way to reach people"}.get(tail, tail)
    else:
        raise ValueError(kind)

    return wrap(op, clause, tail), trim_words(fact, 12)


def extract_candidates(text: str) -> list[dict]:
    """All matching facts in priority order (priority, position)."""
    out = []
    for kind, rx, prio in FACT_PATTERNS:
        for m in rx.finditer(text):
            out.append({"kind": kind, "prio": prio, "pos": m.start(), "m": m, "gd": m.groupdict()})
    out.sort(key=lambda d: (d["prio"], d["pos"]))
    # de-duplicate overlapping matches of the same kind
    seen = []
    res = []
    for c in out:
        span = (c["m"].start(), c["m"].end())
        if any(c["kind"] == k and not (span[1] <= a or span[0] >= b) for k, (a, b) in seen):
            continue
        seen.append((c["kind"], span))
        res.append(c)
    return res


def pick_fact(text: str, opener_idx: int) -> tuple[dict | None, int, str]:
    """Return (chosen, personal_swaps, status_note)."""
    swaps = 0
    cands = extract_candidates(text)
    if not cands:
        return None, 0, "NO_FACT_FOUND_v3"
    for c in cands:
        quote = make_quote(text, c["m"])
        if quote not in text:
            continue  # cannot happen for a slice, but keep the guard the brief asks for
        p = personal_hit(quote)
        if p:
            swaps += 1
            continue
        line, fact = build_line(c["kind"], c["gd"], opener_idx)
        p2 = personal_hit(line) or personal_hit(fact)
        if p2:
            swaps += 1
            continue
        c = dict(c, quote=quote, line=line, fact=fact)
        return c, swaps, "OK_v3"
    return None, swaps, "FACT_PERSONAL_ONLY_v3"


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def find_col(df: pd.DataFrame, *names: str) -> str | None:
    low = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower() in low:
            return low[n.lower()]
    return None


def ensure_col(df: pd.DataFrame, name: str) -> str:
    c = find_col(df, name)
    if c is None:
        df[name] = ""
        return name
    return c


def icp_flag(row_text: str) -> str:
    scores = {k: len(rx.findall(row_text)) for k, rx in ICP_RE.items()}
    estate = len(ESTATE_RE.findall(row_text))
    best = max(scores, key=scores.get)
    if scores[best] == 0:
        return ""
    if scores[best] >= estate or scores[best] >= 3:
        return best
    return ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=str(Path(__file__).resolve().parent))
    ap.add_argument("--master", default="EP701_personalised.csv")
    ap.add_argument("--bios", default="bios_v3.csv")
    args = ap.parse_args()
    d = Path(args.dir)

    master_path, bios_path = d / args.master, d / args.bios
    for p in (master_path, bios_path):
        if not p.exists():
            print(f"ERROR: missing input {p}", file=sys.stderr)
            return 2

    master = pd.read_csv(master_path, dtype=str, keep_default_na=False)
    bios = pd.read_csv(bios_path, dtype=str, keep_default_na=False)

    # --- column resolution -------------------------------------------------
    m_domain = find_col(master, "domain", "website", "url")
    if m_domain is None:
        print("ERROR: master has no domain column", file=sys.stderr)
        return 2
    m_first = ensure_col(master, "first_line")
    m_tier = ensure_col(master, "tier")
    m_status = ensure_col(master, "bio_status")
    m_name = ensure_col(master, "attorney_name")
    m_fact = ensure_col(master, "bio_fact")
    m_quote = ensure_col(master, "bio_quote")
    m_url = ensure_col(master, "bio_url")
    m_sem = ensure_col(master, "seminar_line")
    for extra in ("bio_fact_type", "qa_flags", "icp_flag"):
        ensure_col(master, extra)

    b_domain = find_col(bios, "domain", "website", "url")
    b_err = find_col(bios, "error", "bio_error", "err")
    b_text = find_col(bios, "bio_text", "bio", "text")
    b_url = find_col(bios, "bio_url", "url", "source_url", "page_url")
    b_name = find_col(bios, "attorney_name", "name", "attorney")
    if b_domain is None or b_text is None:
        print("ERROR: bios_v3 needs domain and bio_text columns", file=sys.stderr)
        return 2

    bios["_dom"] = bios[b_domain].map(norm_domain)
    master["_dom"] = master[m_domain].map(norm_domain)
    bio_by_dom: dict[str, dict] = {}
    for _, r in bios.iterrows():
        if r["_dom"] and r["_dom"] not in bio_by_dom:
            bio_by_dom[r["_dom"]] = r.to_dict()

    # --- extraction per candidate bio --------------------------------------
    report = Counter()
    still_c_reasons = Counter()
    b_to_d, c_to_d = 0, 0
    personal_swaps = 0
    extracted: dict[str, dict] = {}
    opener_idx = 0
    for dom, r in bio_by_dom.items():
        err = r.get(b_err, "") if b_err else ""
        text = str(r.get(b_text, "") or "")
        if not blank(err):
            extracted[dom] = {"status": f"BIO_ERROR_v3: {str(err).strip()[:80]}"}
            continue
        if len(text) <= 300:
            extracted[dom] = {"status": f"BIO_SHORT_v3 ({len(text)} chars)"}
            continue
        chosen, swaps, status = pick_fact(text, opener_idx)
        personal_swaps += swaps
        if chosen is None:
            extracted[dom] = {"status": status}
            continue
        # programmatic verbatim check, as the brief requires
        if chosen["quote"] not in text or len(words(chosen["quote"])) > 20:
            extracted[dom] = {"status": "QUOTE_NOT_VERBATIM_v3"}
            report["quote_dropped"] += 1
            continue
        opener_idx += 1
        extracted[dom] = {
            "status": "OK_v3",
            "line": chosen["line"],
            "fact": trim_words(chosen["fact"], 12),
            "quote": chosen["quote"],
            "kind": chosen["kind"],
            "url": r.get(b_url, "") if b_url else "",
            "name": r.get(b_name, "") if b_name else "",
        }

    # --- merge ---------------------------------------------------------------
    for i, row in master.iterrows():
        tier = str(row[m_tier]).strip()
        if tier not in ("B_events_page_only", "C_none"):
            continue  # A_dated_event and existing D_bio_fact are never touched
        dom = row["_dom"]
        ex = extracted.get(dom)
        if ex is None:
            master.at[i, m_status] = "NO_BIO_v3" if dom not in bio_by_dom else "NO_BIO_v3"
            if tier == "C_none":
                still_c_reasons["no bio in bios_v3"] += 1
            continue
        if ex["status"] != "OK_v3":
            master.at[i, m_status] = ex["status"]
            if tier == "C_none":
                still_c_reasons[ex["status"].split(" (")[0].split(":")[0]] += 1
            continue
        if tier == "B_events_page_only":
            master.at[i, m_sem] = row[m_first]
            b_to_d += 1
        else:
            c_to_d += 1
        master.at[i, m_first] = ex["line"]
        master.at[i, m_fact] = ex["fact"]
        master.at[i, m_quote] = ex["quote"]
        master.at[i, m_url] = ex["url"]
        master.at[i, "bio_fact_type"] = ex["kind"]
        if not blank(ex["name"]):
            master.at[i, m_name] = ex["name"]
        master.at[i, m_tier] = "D_bio_fact"
        master.at[i, m_status] = "OK_v3"

    # --- QA ----------------------------------------------------------------
    dup_map = defaultdict(list)
    for i, row in master.iterrows():
        fl = str(row[m_first]).strip()
        if fl:
            dup_map[fl.lower()].append(row["_dom"])
    dups = {k: v for k, v in dup_map.items() if len(set(v)) > 1}

    text_cols = [c for c in master.columns if c not in ("_dom",)]
    icp_counts = Counter()
    for i, row in master.iterrows():
        flags = []
        fl = str(row[m_first]).strip()
        n = len(words(fl))
        if fl:
            if n < 5 or n > 20:
                flags.append(f"len_{n}")
            if has_dash(fl):
                flags.append("dash")
            if not fl.startswith(OPENERS):
                flags.append("opener")
            p = personal_hit(fl)
            if p:
                flags.append(f"personal:{p}")
            if fl.lower() in dups:
                flags.append("duplicate")
        elif str(row[m_tier]).strip() != "C_none":
            flags.append("empty_first_line")
        q = str(row[m_quote]).strip()
        if q:
            dom = row["_dom"]
            bt = str(bio_by_dom.get(dom, {}).get(b_text, "") or "")
            if bt and q not in bt and str(row[m_status]).strip() == "OK_v3":
                flags.append("quote_not_in_bio")
        master.at[i, "qa_flags"] = ";".join(flags)
        blob = " ".join(str(row[c]) for c in text_cols)
        blob += " " + str(bio_by_dom.get(row["_dom"], {}).get(b_text, "") or "")
        f = icp_flag(blob)
        master.at[i, "icp_flag"] = f
        if f:
            icp_counts[f] += 1

    master = master.drop(columns=["_dom"])

    # --- outputs -----------------------------------------------------------
    out_csv = d / "EP701_personalised_v3.csv"
    out_xlsx = d / "EP701_personalised_v3.xlsx"
    master.to_csv(out_csv, index=False)

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
        sheets = [("All", master)]
        for t in sorted(master[m_tier].astype(str).str.strip().unique()):
            name = (t or "no_tier")[:31]
            sheets.append((name, master[master[m_tier].astype(str).str.strip() == t]))
        for name, df in sheets:
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for j, col in enumerate(df.columns, start=1):
                width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) if len(df) else len(str(col))
                ws.column_dimensions[get_column_letter(j)].width = min(max(10, width + 2), 60)

    tier_counts = master[m_tier].value_counts().to_dict()
    still_c = int(tier_counts.get("C_none", 0))
    lines = [
        "# EP701 v3 merge report", "",
        f"Master rows: {len(master)}",
        f"bios_v3 rows: {len(bios)} (unique domains {len(bio_by_dom)})",
        f"Candidates (blank error, bio_text > 300 chars): {sum(1 for v in extracted.values() if not v['status'].startswith(('BIO_ERROR', 'BIO_SHORT')))}",
        f"Facts extracted with verified verbatim quote: {sum(1 for v in extracted.values() if v['status'] == 'OK_v3')}",
        f"Quotes dropped for failing verbatim check: {report['quote_dropped']}", "",
        f"B_events_page_only -> D_bio_fact: {b_to_d}",
        f"C_none -> D_bio_fact: {c_to_d}",
        f"Still C_none: {still_c}",
    ]
    for reason, n in still_c_reasons.most_common():
        lines.append(f"  - {reason}: {n}")
    lines += [
        "", f"Personal swaps (facts rejected for personal content): {personal_swaps}",
        f"Duplicate first_lines across domains: {len(dups)} distinct lines, "
        f"{sum(len(set(v)) for v in dups.values())} rows",
    ]
    for k, v in list(dups.items())[:25]:
        lines.append(f"  - \"{k}\" -> {', '.join(sorted(set(v)))}")
    lines += ["", f"ICP flags: {sum(icp_counts.values())}"]
    for k, v in icp_counts.most_common():
        lines.append(f"  - {k}: {v}")
    lines += ["", "Tier counts:"]
    for k, v in sorted(tier_counts.items()):
        lines.append(f"  - {k}: {v}")
    qa = master["qa_flags"].astype(str)
    lines += ["", f"Rows with QA flags: {int((qa != '').sum())}"]
    for flag, n in Counter(f.split(':')[0] for s in qa if s for f in s.split(';')).most_common():
        lines.append(f"  - {flag}: {n}")
    rep = "\n".join(lines) + "\n"
    (d / "EP701_v3_report.md").write_text(rep)
    print(rep)
    return 0


if __name__ == "__main__":
    sys.exit(main())
