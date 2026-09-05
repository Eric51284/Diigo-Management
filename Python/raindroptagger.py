"""
raindroptagger.py
=================
Built by Claude Sonnet-4.6 on 2026-02-26

Fetch each article URL from CSV or Word docx input and enrich each row with:

    • pub_date   — publication date (YYYY-MM-DD or partial)
    • wordcount  — estimated article body word count

For CSV inputs, AI-tagged rows can also receive semantic outline placement
against Output files/Capstone AI outline headings.csv. On a confident match,
the script writes AI outline diagnostics and upserts a primary _outl:ROMAN-LETTER
token into the tags column.

Usage
-----
    # From a CSV file:
    python raindroptagger.py --csv input.csv
    python raindroptagger.py --csv input.csv -o output_tagged.csv

    # From a Word document:
    python raindroptagger.py --docx articles.docx
    python raindroptagger.py --docx articles.docx -o output_tagged.csv

    # Slow down requests (default: 2 s between each):
    python raindroptagger.py --csv input.csv --delay 3.5

    # Retry blocked URLs via browser, then re-fetch with cookies:
    python raindroptagger.py --csv input.csv --manual-browser-retry
    python raindroptagger.py --csv input.csv --manual-browser-retry --browser-cookies edge

    # Fall back to saved local HTML files for failed fetches:
    python raindroptagger.py --csv input.csv --local-html-dir "Source files/saved_html"

Requirements
------------
    pip install pandas requests beautifulsoup4 python-docx

Optional (improves text extraction accuracy):
    pip install trafilatura       # cleaner article-body extraction
    pip install browser_cookie3   # pass live browser cookies on manual retries

Input format: CSV
-----------------
Any CSV containing at minimum a URL column. Column names are auto-detected
(case-insensitive):

    url column   — detected by name: url, URL, link, Link
                   fallback: first column whose values look like URLs
    title column — detected by name: title, Title, headline
                   (optional; rows without a detected title column use "")
    notes column — detected by name: notes, Notes, note, Note, description,
                   Description, annotation, comment
                   If found, wordcount and pub_date are appended to this
                   column. If no notes column is detected, a new "notes"
                   column is created.
    tags column  — detected by name: tags, Tags, tag, Tag
                   Used for AI row detection and routing hints.
                   If AI placement matches, _outl:* is upserted (existing
                   _outl tokens are replaced, not duplicated).
    local_html_path
                 — optional per-row path to saved .html fallback; name is
                   configurable via --local-html-path-column

All original input columns are preserved in output order. Diagnostic columns
are appended after them.

Input format: Word docx
-----------------------
Every paragraph that starts with "-" or contains "http" is treated as an
article entry. The paragraph text (minus leading "- ") becomes title;
embedded hyperlinks are extracted as URL.

Output CSV columns
------------------
For CSV input, original columns are preserved. Notes are annotated with:

    <original notes text> wordcount:nnnn pub:yyyy-mm-dd

Appended diagnostics:

    pub_date                 — publication date string or empty
    date_status              — date extraction path (meta/jsonld/time/local_*/...)
    wordcount                — estimated body word count or empty
    wc_status                — success / no_url / no_text_found / http_NNN / ...
    wc_method                — extraction method used
    local_html_path          — fallback HTML path used (if any)
    ai_outline_tag           — matched outline tag (example: IV-B)
    ai_outline_subsection_id — matched subsection id (example: s4b)
    ai_outline_heading       — matched subsection heading text
    ai_outline_score         — final semantic score (with boosts)
    ai_outline_score_margin  — best score minus second-best score
    ai_outline_status        — matched / ambiguous:* / low_confidence:* /
                               no_shared_terms / no_semantic_text /
                               not_ai_tagged / no_url / ...

AI outline assignment behavior
------------------------------
Assignment runs only for rows matching --ai-tag-pattern
(default matches ai or #ai token).

Base ranking uses semantic overlap between article text and subsection headings.
Existing tags and title cues provide score boosts to break ties and improve
placement. Examples of built-in hints:

    privacy -> IV-E
    regulatory + open call -> IV-B
    dystopian -> VII-F
    how to + personal tools -> II-F
    education + medical -> V-D
    agentic or release-style title cues -> I-G

CLI arguments
-------------
    --csv PATH                  Input CSV file
    --docx PATH                 Input Word docx file
    -o / --output PATH          Output CSV path
                                  default for --csv: <input>_raindroptagged.csv
                                  default for --docx: Output files/raindrop_tagged.csv
    --delay FLOAT               Seconds to wait between requests [default: 2.0]
    --heartbeat-every INT       Log progress every N requests [default: 10]
    --manual-browser-retry      On HTTP/request errors, open URL in browser,
                                  then retry fetch
    --browser-cookies BROWSER   Use browser cookies during manual retries
                                  choices: chrome | edge | firefox
    --manual-wait-seconds INT   Wait seconds for non-interactive manual retry
                                  [default: 20]
    --local-html-dir DIR        Directory of saved .html/.htm fallback files
    --local-html-path-column COL
                                CSV column for per-row local HTML path
                                  [default: local_html_path]
    --outline-headings-csv PATH Outline headings CSV for AI semantic placement
                                  [default: Output files/Capstone AI outline headings.csv]
    --disable-ai-outline-assignment
                                Disable AI semantic outline placement
    --ai-outline-min-score FLOAT
                                Minimum score required to accept placement
    --ai-tag-pattern REGEX      Regex used to detect AI rows
                                  [default: (^|[\s,;])#?ai\b]
"""

import argparse
import csv
import json
import logging
import math
import os
import re
import sys
import time
import webbrowser
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import trafilatura
except Exception:
    trafilatura = None

try:
    import browser_cookie3
except Exception:
    browser_cookie3 = None

from docx import Document

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


CSV_WRITE_ENCODING = "utf-8-sig"
CSV_READ_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")

OUTLINE_DEFAULT_CSV = os.path.join("Output files", "Capstone AI outline headings.csv")

TAG_CANDIDATE_COLUMNS = ["tags", "Tags", "tag", "Tag"]
NOTES_CANDIDATE_COLUMNS = ["notes", "Notes", "note", "Note", "description", "Description", "annotation", "comment"]

STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "has", "have", "how",
    "in", "is", "it", "its", "of", "on", "or", "that", "the", "their", "this", "to", "with",
    "about", "into", "over", "under", "across", "within", "between", "via", "vs", "using",
    "ai", "artificial", "intelligence",
}

TAG_OUTLINE_HINTS = {
    "privacy": "IV-E",
    "regulatory": "IV-B",
    "open call": "IV-B",
    "personal tools": "II-F",
    "education": "V-D",
    "medical": "V-D",
    "dystopian": "VII-F",
    "agentic": "I-G",
    "key release": "I-G",
}

TITLE_RELEASE_PATTERN = re.compile(
    r"\b(is\s+here|released?|launch(?:ed)?|roll(?:ed)?\s*out|version|gpt[-\s]?\d|claude\s*\d|openclaw\s*\d)\b",
    re.IGNORECASE,
)


def build_http_session():
    session = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=1.0,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET", "HEAD", "OPTIONS"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


HTTP_SESSION = build_http_session()


def get_browser_cookie_jar(browser_name, url):
    if browser_cookie3 is None:
        raise RuntimeError("browser_cookie3 is not installed")

    hostname = urlparse(url).hostname
    if not hostname:
        raise ValueError(f"Could not determine hostname for URL: {url}")

    loaders = {
        "chrome": browser_cookie3.chrome,
        "edge": browser_cookie3.edge,
        "firefox": browser_cookie3.firefox,
    }
    loader = loaders.get(browser_name)
    if loader is None:
        raise ValueError(f"Unsupported browser for cookies: {browser_name}")

    return loader(domain_name=hostname)


# ---------- Helpers copied/adapted from existing scripts ----------

def normalize_text(text):
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def count_words(text):
    words = re.findall(r"\b[\w'-]+\b", text)
    return len(words)


def tokenize_for_semantic(text):
    if not text:
        return []
    tokens = re.findall(r"[a-zA-Z][a-zA-Z0-9'-]{1,}", text.lower())
    clean = []
    for tok in tokens:
        tok = tok.strip("'-")
        if len(tok) < 2:
            continue
        if tok in STOPWORDS:
            continue
        clean.append(tok)
    return clean


def sub_id_to_outl_tag(sub_id):
    m = re.fullmatch(r"s(\d+)([a-z])", str(sub_id or ""), re.IGNORECASE)
    if not m:
        return None
    roman_map = {
        1: "I", 2: "II", 3: "III", 4: "IV", 5: "V",
        6: "VI", 7: "VII", 8: "VIII", 9: "IX", 10: "X",
    }
    sec_num = int(m.group(1))
    roman = roman_map.get(sec_num)
    if not roman:
        return None
    return f"{roman}-{m.group(2).upper()}"


def outl_tag_to_sub_id(outl_tag):
    if not outl_tag:
        return None
    m = re.fullmatch(r"([IVX]+)-([A-Z])", str(outl_tag).strip(), re.IGNORECASE)
    if not m:
        return None
    roman_to_num = {
        "I": 1, "II": 2, "III": 3, "IV": 4, "V": 5,
        "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10,
    }
    sec_num = roman_to_num.get(m.group(1).upper())
    if sec_num is None:
        return None
    return f"s{sec_num}{m.group(2).lower()}"


def split_tag_tokens(tag_value):
    if tag_value is None:
        return []
    raw = str(tag_value)
    parts = [p.strip().lower() for p in re.split(r"[,;]", raw) if p and p.strip()]
    return parts


def derive_outline_hints_from_tags(tag_value):
    tokens = split_tag_tokens(tag_value)
    token_set = set(tokens)
    hints = {}

    if "regulatory" in token_set and "dystopian" in token_set:
        sub_id = outl_tag_to_sub_id("IV-B")
        if sub_id:
            hints[sub_id] = 0.35

    if "education" in token_set and "medical" in token_set:
        sub_id = outl_tag_to_sub_id("V-D")
        if sub_id:
            hints[sub_id] = max(hints.get(sub_id, 0.0), 0.30)

    if "how to" in token_set and "personal tools" in token_set:
        sub_id = outl_tag_to_sub_id("II-F")
        if sub_id:
            hints[sub_id] = max(hints.get(sub_id, 0.0), 0.30)

    for tok in tokens:
        outl = TAG_OUTLINE_HINTS.get(tok)
        if not outl:
            continue
        sub_id = outl_tag_to_sub_id(outl)
        if not sub_id:
            continue
        hints[sub_id] = max(hints.get(sub_id, 0.0), 0.22)

    if "how to" in token_set and "privacy" not in token_set:
        sub_id = outl_tag_to_sub_id("II-F")
        if sub_id:
            hints[sub_id] = max(hints.get(sub_id, 0.0), 0.18)

    return hints


def derive_outline_hints_from_title(title):
    if not title:
        return {}
    hints = {}
    if TITLE_RELEASE_PATTERN.search(str(title)):
        sub_id = outl_tag_to_sub_id("I-G")
        if sub_id:
            hints[sub_id] = 0.25
    return hints


def load_outline_matcher(outline_csv_path):
    if not outline_csv_path or not os.path.exists(outline_csv_path):
        logger.warning("Outline headings CSV not found: %s", outline_csv_path)
        return None

    section_titles = {}
    subsection_profiles = []

    with open(outline_csv_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            level = str(row.get("level") or "").strip()
            section_id = str(row.get("section_id") or "").strip()
            title = str(row.get("title") or "").strip()
            if level == "1" and section_id and title:
                section_titles[section_id] = title

    with open(outline_csv_path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            level = str(row.get("level") or "").strip()
            if level != "2":
                continue
            section_id = str(row.get("section_id") or "").strip()
            subsection_id = str(row.get("subsection_id") or "").strip()
            sub_title = str(row.get("title") or "").strip()
            if not subsection_id or not sub_title:
                continue
            section_title = section_titles.get(section_id, "")
            merged_text = f"{sub_title} {section_title}".strip()
            token_set = set(tokenize_for_semantic(merged_text))
            subsection_profiles.append(
                {
                    "section_id": section_id,
                    "subsection_id": subsection_id,
                    "outline_tag": sub_id_to_outl_tag(subsection_id),
                    "title": sub_title,
                    "section_title": section_title,
                    "tokens": token_set,
                }
            )

    if not subsection_profiles:
        logger.warning("No subsection rows found in outline headings CSV: %s", outline_csv_path)
        return None

    doc_count = len(subsection_profiles)
    token_df = {}
    for profile in subsection_profiles:
        for tok in profile["tokens"]:
            token_df[tok] = token_df.get(tok, 0) + 1

    idf = {tok: (math.log((doc_count + 1) / (df + 1)) + 1.0) for tok, df in token_df.items()}

    for profile in subsection_profiles:
        token_weight_sq = sum((idf.get(tok, 1.0) ** 2) for tok in profile["tokens"])
        profile["token_norm"] = math.sqrt(token_weight_sq) if token_weight_sq > 0 else 1.0

    logger.info("Loaded %d outline subsections from %s", len(subsection_profiles), outline_csv_path)
    return {"profiles": subsection_profiles, "idf": idf}


def semantic_match_outline(
    title,
    article_text,
    outline_matcher,
    min_score=0.12,
    min_margin=0.015,
    boost_hints=None,
):
    if not outline_matcher:
        return None, "outline_unavailable"

    combined = f"{title or ''} {article_text or ''}".strip()
    article_tokens = set(tokenize_for_semantic(combined))
    if not article_tokens:
        return None, "no_semantic_text"

    idf = outline_matcher["idf"]
    article_norm_sq = sum((idf.get(tok, 1.0) ** 2) for tok in article_tokens)
    if article_norm_sq <= 0:
        return None, "no_semantic_text"
    article_norm = math.sqrt(article_norm_sq)

    scored = []
    boost_hints = boost_hints or {}

    for profile in outline_matcher["profiles"]:
        shared = article_tokens & profile["tokens"]
        if not shared:
            score = 0.0
        else:
            dot = sum((idf.get(tok, 1.0) ** 2) for tok in shared)
            score = dot / max(article_norm * profile["token_norm"], 1e-9)

        score += float(boost_hints.get(profile["subsection_id"], 0.0))
        if score <= 0:
            continue
        scored.append((score, profile))

    if not scored:
        return None, "no_shared_terms"

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_profile = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0

    if best_score < min_score:
        return None, f"low_confidence:{best_score:.3f}"
    if (best_score - second_score) < min_margin:
        return None, f"ambiguous:{best_score:.3f}/{second_score:.3f}"

    best_profile = dict(best_profile)
    best_profile["score"] = round(best_score, 4)
    best_profile["score_margin"] = round(best_score - second_score, 4)
    return best_profile, "matched"


def is_ai_tagged_article(article, tag_col=None, notes_col=None, ai_tag_pattern=r"(^|[\s,;])#?ai\b"):
    rx = re.compile(ai_tag_pattern, re.IGNORECASE)
    candidate_cols = []
    if tag_col:
        candidate_cols.append(tag_col)
    if notes_col and notes_col not in candidate_cols:
        candidate_cols.append(notes_col)
    for fallback in TAG_CANDIDATE_COLUMNS + NOTES_CANDIDATE_COLUMNS:
        if fallback not in candidate_cols:
            candidate_cols.append(fallback)

    for col in candidate_cols:
        val = article.get(col)
        if val is None:
            continue
        text = str(val)
        if rx.search(text):
            return True
    return False


def append_csv_tag(existing_value, new_tag):
    existing = str(existing_value or "").strip()
    if not existing:
        return new_tag
    parts = [p.strip() for p in existing.split(",") if p.strip()]
    norm = {p.lower() for p in parts}
    if new_tag.lower() in norm:
        return existing
    parts.append(new_tag)
    return ", ".join(parts)


def upsert_outline_tag(existing_value, outl_tag_token):
    existing = str(existing_value or "").strip()
    parts = [p.strip() for p in existing.split(",") if p.strip()]
    kept = [p for p in parts if not p.lower().startswith("_outl:")]
    kept.append(outl_tag_token)
    return ", ".join(kept)


def parse_date_string(date_str):
    if not date_str:
        return None
    date_str = str(date_str).strip()
    formats = [
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%SZ",
        "%Y-%m-%dT%H:%M:%S.%fZ",
        "%B %d, %Y",
        "%b %d, %Y",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%B %d %Y",
        "%b %d %Y",
        "%d %B %Y",
        "%d %b %Y",
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    m = re.search(r"(\d{4}-\d{2}-\d{2})", date_str)
    if m:
        return m.group(1)
    return None


def extract_json_ld_article_text(html_text):
    try:
        soup = BeautifulSoup(html_text, "html.parser")
        scripts = soup.find_all("script", type="application/ld+json")
        candidates = []

        def collect_texts(obj):
            if isinstance(obj, dict):
                for key, value in obj.items():
                    if key in {"articleBody", "text", "description"}:
                        text_val = normalize_text(value)
                        if text_val:
                            candidates.append(text_val)
                    else:
                        collect_texts(value)
            elif isinstance(obj, list):
                for item in obj:
                    collect_texts(item)

        for script in scripts:
            if not script.string:
                continue
            try:
                data = json.loads(script.string)
            except Exception:
                continue
            collect_texts(data)

        if not candidates:
            return "", "jsonld_unavailable"
        best_text = max(candidates, key=count_words)
        return best_text, "jsonld"
    except Exception:
        return "", "jsonld_failed"


BOILERPLATE_CLASS_ID_PATTERN = re.compile(
    r"(^|[-_\s])(ad|ads|advert|advertisement|sponsor|promo|related|newsletter|"
    r"footer|sidebar|share|social|cookie|banner|recommend|trending|outbrain|"
    r"taboola)($|[-_\s])",
    re.IGNORECASE,
)


def extract_main_text_with_bs4(soup):
    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()
    nodes_to_remove = []
    for node in soup.find_all(True):
        class_attr = " ".join(node.get("class", []))
        id_attr = node.get("id", "")
        marker = f"{class_attr} {id_attr}".strip()
        if marker and BOILERPLATE_CLASS_ID_PATTERN.search(marker):
            nodes_to_remove.append(node)
    for node in nodes_to_remove:
        node.decompose()
    article_node = soup.find("article")
    if article_node:
        return normalize_text(article_node.get_text(" ", strip=True)), "article_tag"
    main_node = soup.find("main")
    if main_node:
        return normalize_text(main_node.get_text(" ", strip=True)), "main_tag"
    if soup.body:
        return normalize_text(soup.body.get_text(" ", strip=True)), "body_fallback"
    return normalize_text(soup.get_text(" ", strip=True)), "document_fallback"


def extract_main_text_with_trafilatura(html_text):
    if trafilatura is None:
        return "", "trafilatura_unavailable"
    try:
        extracted = trafilatura.extract(html_text, include_comments=False, include_tables=False, favor_precision=True)
        extracted = normalize_text(extracted)
        if extracted:
            return extracted, "trafilatura"
    except Exception:
        pass
    return "", "trafilatura_failed"


def get_pub_date_from_soup(soup):
    # try meta tags
    meta_selectors = [
        'meta[property="article:published_time"]',
        'meta[property="article:published"]',
        'meta[name="publish-date"]',
        'meta[name="publication-date"]',
        'meta[name="date"]',
        'meta[name="DC.date"]',
        'meta[name="DC.Date"]',
        'meta[property="og:published_time"]',
        'meta[name="publishdate"]',
        'meta[name="pub_date"]',
        'meta[itemprop="datePublished"]',
        'meta[itemprop="publishDate"]',
    ]
    for selector in meta_selectors:
        meta = soup.select_one(selector)
        if meta:
            content = meta.get("content") or meta.get("value")
            if content:
                parsed = parse_date_string(content)
                if parsed:
                    return parsed, "meta"
    # json-ld
    try:
        scripts = soup.find_all("script", type="application/ld+json")
        for script in scripts:
            if not script.string:
                continue
            try:
                data = json.loads(script.string)
            except Exception:
                continue
            items = data if isinstance(data, list) else [data]
            for item in items:
                if isinstance(item, dict):
                    for f in ("datePublished", "publishDate", "dateCreated", "uploadDate"):
                        if f in item:
                            parsed = parse_date_string(item[f])
                            if parsed:
                                return parsed, "jsonld"
    except Exception:
        pass
    # time tags and common selectors
    time_selectors = ["time[datetime]", "time[pubdate]", ".published-date time", ".publish-date time", ".date time"]
    for selector in time_selectors:
        elem = soup.select_one(selector)
        if elem:
            dtattr = elem.get("datetime") or elem.get("pubdate")
            if dtattr:
                parsed = parse_date_string(dtattr)
                if parsed:
                    return parsed, "time_attr"
            txt = elem.get_text().strip()
            if txt:
                parsed = parse_date_string(txt)
                if parsed:
                    return parsed, "time_text"
    # article/date classes
    date_selectors = [
        ".published-date",
        ".publish-date",
        ".publication-date",
        ".date-published",
        ".article-date",
        ".post-date",
        ".entry-date",
        ".timestamp",
        '[class*="date"]',
        '[class*="publish"]',
    ]
    for selector in date_selectors:
        elems = soup.select(selector)
        for e in elems:
            text = e.get_text().strip()
            if text and len(text) < 100:
                parsed = parse_date_string(text)
                if parsed:
                    return parsed, "class_text"
    # raw text patterns
    text = soup.get_text()
    patterns = [r"Published:?\s*([A-Za-z]+ \d{1,2},? \d{4})", r"(\d{1,2}/\d{1,2}/\d{4})", r"(\d{4}-\d{2}-\d{2})", r"([A-Za-z]+ \d{1,2},? \d{4})"]
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for m in matches:
            parsed = parse_date_string(m)
            if parsed:
                return parsed, "text_pattern"
    return None, "no_date_found"


def get_wordcount_from_html(html_text, soup):
    candidates = []
    jsonld_text, jsonld_method = extract_json_ld_article_text(html_text)
    if jsonld_text:
        candidates.append((jsonld_text, jsonld_method))
    tr_text, tr_method = extract_main_text_with_trafilatura(html_text)
    if tr_text:
        candidates.append((tr_text, tr_method))
    if soup:
        article_node = soup.find("article")
        if article_node:
            article_p_text = normalize_text(" ".join(p.get_text(" ", strip=True) for p in article_node.find_all("p")))
            if article_p_text:
                candidates.append((article_p_text, "article_p"))
            article_raw = normalize_text(article_node.get_text(" ", strip=True))
            if article_raw:
                candidates.append((article_raw, "article_tag"))
        main_node = soup.find("main")
        if main_node:
            main_p_text = normalize_text(" ".join(p.get_text(" ", strip=True) for p in main_node.find_all("p")))
            if main_p_text:
                candidates.append((main_p_text, "main_p"))
        all_p_text = normalize_text(" ".join(p.get_text(" ", strip=True) for p in soup.find_all("p")))
        if all_p_text:
            candidates.append((all_p_text, "all_p"))
        bs4_text, bs4_method = extract_main_text_with_bs4(soup)
        if bs4_text:
            candidates.append((bs4_text, bs4_method))
        if soup.body:
            body_text = normalize_text(soup.body.get_text(" ", strip=True))
            if body_text:
                candidates.append((body_text, "body_full"))
    # score candidates
    scored = []
    for candidate_text, candidate_method in candidates:
        wc = count_words(candidate_text)
        if wc > 0:
            scored.append((wc, candidate_method))
    if not scored:
        return None, "no_text_found", "no_candidate_text"
    wc_by_method = {method: wc for wc, method in scored}
    for preferred in ["article_p", "main_p", "jsonld", "trafilatura", "article_tag"]:
        if wc_by_method.get(preferred, 0) >= 120:
            return wc_by_method[preferred], "success", preferred
    non_full = [(wc, method) for wc, method in scored if method not in {"all_p", "body_full"}]
    if non_full:
        best_wc, best_method = max(non_full, key=lambda item: item[0])
        return best_wc, "success", best_method
    ranked = sorted(scored, key=lambda item: item[0], reverse=True)
    if len(ranked) >= 2 and ranked[0][0] > int(ranked[1][0] * 1.6) and (ranked[0][0] - ranked[1][0]) > 500:
        best_wc, best_method = ranked[1]
    else:
        best_wc, best_method = ranked[0]
    return best_wc, "success", best_method


def extract_semantic_text_from_html(html_text, soup):
    candidates = []
    jsonld_text, _ = extract_json_ld_article_text(html_text)
    if jsonld_text:
        candidates.append(jsonld_text)
    tr_text, _ = extract_main_text_with_trafilatura(html_text)
    if tr_text:
        candidates.append(tr_text)
    if soup:
        article_node = soup.find("article")
        if article_node:
            article_p_text = normalize_text(" ".join(p.get_text(" ", strip=True) for p in article_node.find_all("p")))
            if article_p_text:
                candidates.append(article_p_text)
        main_node = soup.find("main")
        if main_node:
            main_p_text = normalize_text(" ".join(p.get_text(" ", strip=True) for p in main_node.find_all("p")))
            if main_p_text:
                candidates.append(main_p_text)
        bs4_text, _ = extract_main_text_with_bs4(soup)
        if bs4_text:
            candidates.append(bs4_text)

    # Prefer the richest extracted candidate as semantic context.
    best_text = ""
    best_wc = 0
    for text in candidates:
        wc = count_words(text)
        if wc > best_wc:
            best_text = text
            best_wc = wc
    return best_text


# ---------- Document hyperlink extraction (from NewArticles) ----------

def extract_hyperlink_method1(paragraph):
    try:
        for elem in paragraph._element.iter():
            if "hyperlink" in str(elem.tag).lower():
                r_id = elem.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if r_id and hasattr(paragraph.part, "rels") and r_id in paragraph.part.rels:
                    url = paragraph.part.rels[r_id].target_ref
                    return url
    except Exception:
        pass
    return None


def extract_hyperlink_method2(paragraph):
    try:
        for run in paragraph.runs:
            if run._element.rPr is not None:
                for elem in run._element.iter():
                    if "hyperlink" in str(elem.tag).lower():
                        r_id = elem.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                        if r_id and hasattr(paragraph.part, "rels") and r_id in paragraph.part.rels:
                            url = paragraph.part.rels[r_id].target_ref
                            return url
    except Exception:
        pass
    return None


def extract_hyperlink_method3(paragraph):
    try:
        xml_str = str(paragraph._element.xml)
        hyperlink_pattern = r'r:id="(rId\d+)"'
        matches = re.findall(hyperlink_pattern, xml_str)
        for r_id in matches:
            if hasattr(paragraph.part, "rels") and r_id in paragraph.part.rels:
                url = paragraph.part.rels[r_id].target_ref
                return url
    except Exception:
        pass
    return None


def extract_articles_and_links_from_docx(docx_path):
    doc = Document(docx_path)
    articles = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        # treat any paragraph with a hyphen or with a URL as an article line
        if text.startswith("-") or "http" in text.lower():
            title = text.lstrip("- ").strip()
            hyperlink_url = (
                extract_hyperlink_method1(paragraph)
                or extract_hyperlink_method2(paragraph)
                or extract_hyperlink_method3(paragraph)
            )
            articles.append({"title": title, "url": hyperlink_url})
    return articles


def fetch_url_once(url, timeout=30, cookie_jar=None):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
    }
    try:
        resp = HTTP_SESSION.get(url, headers=headers, timeout=timeout, cookies=cookie_jar)
        status_code = resp.status_code
        html_text = resp.text
        soup = BeautifulSoup(resp.content, "html.parser")
        if status_code >= 400:
            logger.warning(f"HTTP {status_code} for {url}")
            return html_text, soup, f"http_{status_code}"
        return html_text, soup, "success"
    except requests.exceptions.Timeout:
        return None, None, "timeout"
    except requests.exceptions.RequestException as e:
        logger.warning(f"Request error for {url}: {e}")
        return None, None, "request_error"
    except Exception as e:
        logger.warning(f"Unexpected error fetching {url}: {e}")
        return None, None, "error"


def build_output_path(input_path, suffix="_raindroptagged.csv"):
    root, ext = os.path.splitext(input_path)
    return f"{root}{suffix}"


def read_csv_with_fallbacks(csv_path):
    last_error = None
    for encoding in CSV_READ_ENCODINGS:
        try:
            return pd.read_csv(csv_path, encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    if last_error:
        raise last_error
    return pd.read_csv(csv_path)


def sanitize_url_for_filename(url):
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()
    path = parsed.path or ""
    if path.endswith("/"):
        path = path[:-1]
    candidate = f"{host}{path}"
    candidate = re.sub(r"[^a-zA-Z0-9._-]+", "_", candidate)
    return candidate.strip("_")


def find_local_html_path(article, url, local_html_dir=None, local_html_path_column="local_html_path"):
    csv_path = article.get(local_html_path_column)
    if csv_path and str(csv_path).strip():
        direct_path = Path(str(csv_path).strip().strip('"'))
        if direct_path.exists() and direct_path.is_file():
            return direct_path

    if not local_html_dir:
        return None

    directory = Path(local_html_dir)
    if not directory.exists() or not directory.is_dir():
        return None

    url_key = sanitize_url_for_filename(url)
    if url_key:
        for ext in (".html", ".htm"):
            candidate = directory / f"{url_key}{ext}"
            if candidate.exists() and candidate.is_file():
                return candidate
        for candidate in directory.glob(f"{url_key}*.htm*"):
            if candidate.is_file():
                return candidate

    parsed = urlparse(url)
    slug = Path(parsed.path).name
    if slug:
        slug_pattern = re.sub(r"[^a-zA-Z0-9._-]+", "_", slug).lower()
        for candidate in directory.glob("*.htm*"):
            if slug_pattern in candidate.stem.lower():
                return candidate

    return None


def fetch_local_html(article, url, local_html_dir=None, local_html_path_column="local_html_path"):
    local_path = find_local_html_path(
        article,
        url,
        local_html_dir=local_html_dir,
        local_html_path_column=local_html_path_column,
    )
    if local_path is None:
        return None, None, None, "local_html_not_found"

    try:
        html_text = local_path.read_text(encoding="utf-8", errors="ignore")
        soup = BeautifulSoup(html_text, "html.parser")
        return html_text, soup, str(local_path), "success"
    except Exception as exc:
        logger.warning("Failed reading local HTML %s: %s", local_path, exc)
        return None, None, str(local_path), "local_html_read_error"


def save_failed_response_html(url, html_text, local_html_dir, status_label, attempt_label="initial"):
    if not local_html_dir or not html_text:
        return None

    directory = Path(local_html_dir)
    try:
        directory.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        logger.warning("Could not create local HTML dir %s: %s", directory, exc)
        return None

    url_key = sanitize_url_for_filename(url) or "failed_url"
    safe_status = re.sub(r"[^a-zA-Z0-9._-]+", "_", (status_label or "error"))
    safe_attempt = re.sub(r"[^a-zA-Z0-9._-]+", "_", (attempt_label or "attempt"))
    save_path = directory / f"{url_key}__{safe_attempt}__{safe_status}.html"

    try:
        save_path.write_text(html_text, encoding="utf-8", errors="ignore")
        logger.info("Saved failed response HTML to %s", save_path)
        return str(save_path)
    except Exception as exc:
        logger.warning("Failed to save response HTML to %s: %s", save_path, exc)
        return None


def should_attempt_manual_retry(fetch_status):
    if not fetch_status:
        return False
    if fetch_status.startswith("http_"):
        return True
    return fetch_status in {"request_error", "timeout"}


def process_articles(
    articles,
    delay=2.0,
    heartbeat_every=10,
    manual_browser_retry=False,
    browser_cookies=None,
    manual_wait_seconds=20,
    local_html_dir=None,
    local_html_path_column="local_html_path",
    outline_matcher=None,
    ai_outline_enabled=False,
    ai_outline_min_score=0.12,
    ai_outline_tag_col=None,
    notes_col=None,
    ai_tag_pattern=r"(^|[\s,;])#?ai\b",
):
    total = len(articles)
    processed = 0
    success_count = 0
    start = time.perf_counter()
    for idx, article in enumerate(articles, 1):
        url = article.get("url")
        title = article.get("title") or ""
        ai_tagged = is_ai_tagged_article(article, tag_col=ai_outline_tag_col, notes_col=notes_col, ai_tag_pattern=ai_tag_pattern)
        tag_hint_boosts = derive_outline_hints_from_tags(article.get(ai_outline_tag_col)) if ai_outline_tag_col else {}
        title_hint_boosts = derive_outline_hints_from_title(title)
        boost_hints = {}
        for sub_id, weight in tag_hint_boosts.items():
            boost_hints[sub_id] = max(boost_hints.get(sub_id, 0.0), weight)
        for sub_id, weight in title_hint_boosts.items():
            boost_hints[sub_id] = max(boost_hints.get(sub_id, 0.0), weight)
        if not url:
            article.update(
                {
                    "pub_date": None,
                    "date_status": "no_url",
                    "wordcount": None,
                    "wc_status": "no_url",
                    "wc_method": None,
                    "ai_outline_status": "no_url" if ai_tagged else "not_ai_tagged",
                }
            )
            continue
        processed += 1
        logger.info(f"[{processed}/{total}] Fetching {url[:90]}")
        html_text, soup, fetch_status = fetch_url_once(url)
        failed_html_path = None
        if fetch_status != "success":
            failed_html_path = save_failed_response_html(
                url,
                html_text,
                local_html_dir=local_html_dir,
                status_label=fetch_status,
                attempt_label="initial",
            )
            retried_status = fetch_status
            if manual_browser_retry and should_attempt_manual_retry(fetch_status):
                logger.info("Manual browser retry enabled for %s", url)
                try:
                    opened = webbrowser.open(url, new=2)
                    if opened:
                        logger.info("Opened URL in your default browser for manual unlock.")
                    else:
                        logger.warning("Could not automatically open browser. Open URL manually: %s", url)
                except Exception as open_error:
                    logger.warning("Failed to open browser automatically for %s: %s", url, open_error)
                logger.info("Let the page fully load in your browser, then press Enter here to retry.")
                proceed_with_retry = False
                if sys.stdin and sys.stdin.isatty():
                    try:
                        input("Press Enter to retry this URL now... ")
                        proceed_with_retry = True
                    except EOFError:
                        logger.warning("No interactive input available for %s", url)
                else:
                    wait_seconds = max(0, int(manual_wait_seconds))
                    if wait_seconds > 0:
                        logger.warning(
                            "No interactive stdin; waiting %ds before retry so you can open the URL in your browser.",
                            wait_seconds,
                        )
                        time.sleep(wait_seconds)
                        proceed_with_retry = True
                    else:
                        logger.warning("No interactive stdin and --manual-wait-seconds=0; skipping manual retry for %s", url)

                if proceed_with_retry:
                    cookie_jar = None
                    if browser_cookies:
                        try:
                            cookie_jar = get_browser_cookie_jar(browser_cookies, url)
                            logger.info("Loaded browser cookies from %s for %s", browser_cookies, url)
                        except Exception as cookie_error:
                            logger.warning("Could not load %s cookies for %s: %s", browser_cookies, url, cookie_error)
                    html_text, soup, retry_status = fetch_url_once(url, cookie_jar=cookie_jar)
                    if retry_status == "success":
                        pub_date, date_status = get_pub_date_from_soup(soup)
                        wc, wc_status, wc_method = get_wordcount_from_html(html_text, soup)
                        article.update({"pub_date": pub_date, "date_status": date_status, "wordcount": wc, "wc_status": wc_status, "wc_method": wc_method})
                        if ai_tagged and ai_outline_enabled:
                            sem_text = extract_semantic_text_from_html(html_text, soup)
                            match, sem_status = semantic_match_outline(
                                title,
                                sem_text,
                                outline_matcher,
                                min_score=ai_outline_min_score,
                                boost_hints=boost_hints,
                            )
                            if match:
                                article.update(
                                    {
                                        "ai_outline_tag": match.get("outline_tag"),
                                        "ai_outline_subsection_id": match.get("subsection_id"),
                                        "ai_outline_heading": match.get("title"),
                                        "ai_outline_score": match.get("score"),
                                        "ai_outline_score_margin": match.get("score_margin"),
                                        "ai_outline_status": sem_status,
                                    }
                                )
                                if ai_outline_tag_col and match.get("outline_tag"):
                                    article[ai_outline_tag_col] = upsert_outline_tag(article.get(ai_outline_tag_col), f"_outl:{match['outline_tag']}")
                            else:
                                article.update({"ai_outline_status": sem_status})
                        elif ai_tagged:
                            article.update({"ai_outline_status": "outline_disabled"})
                        else:
                            article.update({"ai_outline_status": "not_ai_tagged"})
                        if wc_status == "success":
                            success_count += 1
                        retried_status = "success"
                    else:
                        manual_failed_html_path = save_failed_response_html(
                            url,
                            html_text,
                            local_html_dir=local_html_dir,
                            status_label=retry_status,
                            attempt_label="manual",
                        )
                        if manual_failed_html_path:
                            failed_html_path = manual_failed_html_path
                        retried_status = f"manual_{retry_status}"

            if retried_status != "success":
                local_html_text, local_html_soup, local_html_source, local_html_status = fetch_local_html(
                    article,
                    url,
                    local_html_dir=local_html_dir,
                    local_html_path_column=local_html_path_column,
                )
                if local_html_status == "success":
                    pub_date, date_status = get_pub_date_from_soup(local_html_soup)
                    wc, wc_status, wc_method = get_wordcount_from_html(local_html_text, local_html_soup)
                    article.update(
                        {
                            "pub_date": pub_date,
                            "date_status": f"local_{date_status}",
                            "wordcount": wc,
                            "wc_status": f"local_{wc_status}",
                            "wc_method": f"local_{wc_method}" if wc_method else "local",
                            "local_html_path": local_html_source or failed_html_path,
                        }
                    )
                    if ai_tagged and ai_outline_enabled:
                        sem_text = extract_semantic_text_from_html(local_html_text, local_html_soup)
                        match, sem_status = semantic_match_outline(
                            title,
                            sem_text,
                            outline_matcher,
                            min_score=ai_outline_min_score,
                            boost_hints=boost_hints,
                        )
                        if match:
                            article.update(
                                {
                                    "ai_outline_tag": match.get("outline_tag"),
                                    "ai_outline_subsection_id": match.get("subsection_id"),
                                    "ai_outline_heading": match.get("title"),
                                    "ai_outline_score": match.get("score"),
                                    "ai_outline_score_margin": match.get("score_margin"),
                                    "ai_outline_status": sem_status,
                                }
                            )
                            if ai_outline_tag_col and match.get("outline_tag"):
                                article[ai_outline_tag_col] = upsert_outline_tag(article.get(ai_outline_tag_col), f"_outl:{match['outline_tag']}")
                        else:
                            article.update({"ai_outline_status": sem_status})
                    elif ai_tagged:
                        article.update({"ai_outline_status": "outline_disabled"})
                    else:
                        article.update({"ai_outline_status": "not_ai_tagged"})
                    if wc_status == "success":
                        success_count += 1
                else:
                    article.update(
                        {
                            "pub_date": None,
                            "date_status": retried_status,
                            "wordcount": None,
                            "wc_status": retried_status,
                            "wc_method": None,
                            "local_html_path": local_html_source or failed_html_path,
                            "ai_outline_status": retried_status if ai_tagged else "not_ai_tagged",
                        }
                    )
        else:
            pub_date, date_status = get_pub_date_from_soup(soup)
            wc, wc_status, wc_method = get_wordcount_from_html(html_text, soup)
            article.update({"pub_date": pub_date, "date_status": date_status, "wordcount": wc, "wc_status": wc_status, "wc_method": wc_method, "local_html_path": None})
            if ai_tagged and ai_outline_enabled:
                sem_text = extract_semantic_text_from_html(html_text, soup)
                match, sem_status = semantic_match_outline(
                    title,
                    sem_text,
                    outline_matcher,
                    min_score=ai_outline_min_score,
                    boost_hints=boost_hints,
                )
                if match:
                    article.update(
                        {
                            "ai_outline_tag": match.get("outline_tag"),
                            "ai_outline_subsection_id": match.get("subsection_id"),
                            "ai_outline_heading": match.get("title"),
                            "ai_outline_score": match.get("score"),
                            "ai_outline_score_margin": match.get("score_margin"),
                            "ai_outline_status": sem_status,
                        }
                    )
                    if ai_outline_tag_col and match.get("outline_tag"):
                        article[ai_outline_tag_col] = upsert_outline_tag(article.get(ai_outline_tag_col), f"_outl:{match['outline_tag']}")
                else:
                    article.update({"ai_outline_status": sem_status})
            elif ai_tagged:
                article.update({"ai_outline_status": "outline_disabled"})
            else:
                article.update({"ai_outline_status": "not_ai_tagged"})
            if date_status == "no_date_found":
                logger.info(f"No date for: {title[:60]}")
            if wc_status == "success":
                success_count += 1

        if heartbeat_every > 0 and processed % heartbeat_every == 0:
            elapsed = time.perf_counter() - start
            avg = elapsed / processed if processed else 0
            remaining = max(total - processed, 0)
            eta = avg * remaining
            logger.info("HEARTBEAT: %d/%d elapsed=%.1fs eta=%.1fs success_wc=%d", processed, total, elapsed, eta, success_count)

        if processed < total:
            time.sleep(delay)

    return articles


def save_results_csv(articles, output_csv_path, original_columns=None, notes_col=None):
    df = pd.DataFrame(articles)

    # ── Annotate the Notes column with wordcount and pub_date ────────────────
    # Pattern written into notes: "wordcount:nnnn pub:yyyy-mm-dd"
    # Skips a tag if an identical tag already exists in the notes text.
    def _build_annotation(row, existing=""):
        parts = []
        wc = row.get("wordcount")
        pub = row.get("pub_date")
        if wc is not None and pd.notna(wc) and not re.search(r"wordcount:\d+", existing, re.IGNORECASE):
            try:
                parts.append(f"wordcount:{int(wc)}")
            except (ValueError, TypeError):
                pass
        if pub and pd.notna(pub) and not re.search(r"pub:\d{4}", existing, re.IGNORECASE):
            parts.append(f"pub:{pub}")
        return " ".join(parts)

    if notes_col:
        if notes_col not in df.columns:
            df[notes_col] = ""
        def _update_notes(row):
            existing = str(row[notes_col]).strip() if pd.notna(row[notes_col]) else ""
            annotation = _build_annotation(row, existing)
            if existing and annotation:
                return f"{existing} {annotation}"
            return existing or annotation
        df[notes_col] = df.apply(_update_notes, axis=1)
    else:
        # No notes column in source — create one to carry the annotation
        notes_col = "notes"
        df[notes_col] = df.apply(lambda row: _build_annotation(row, ""), axis=1)
        if original_columns is not None:
            original_columns = list(original_columns) + [notes_col]

    # ── Build output column order ─────────────────────────────────────────────
    # Base: original input columns (in original order, notes already updated)
    # Appended: diagnostic columns not present in the original input
    extra_cols = [
        "pub_date",
        "date_status",
        "wordcount",
        "wc_status",
        "wc_method",
        "local_html_path",
        "ai_outline_tag",
        "ai_outline_subsection_id",
        "ai_outline_heading",
        "ai_outline_score",
        "ai_outline_score_margin",
        "ai_outline_status",
    ]

    if original_columns:
        base_cols = [c for c in original_columns if c in df.columns]
        appended = [c for c in extra_cols if c in df.columns and c not in base_cols]
        output_cols = base_cols + appended
    else:
        # docx path — sensible default ordering
        fallback = ["title", "url", "notes"] + extra_cols
        output_cols = [c for c in fallback if c in df.columns]
        for c in df.columns:
            if c not in output_cols:
                output_cols.append(c)

    for col in output_cols:
        if col not in df.columns:
            df[col] = None

    df = df[output_cols]
    output_dir = os.path.dirname(output_csv_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # utf-8-sig keeps Unicode punctuation intact and opens cleanly in Excel.
    df.to_csv(output_csv_path, index=False, encoding=CSV_WRITE_ENCODING)
    logger.info(f"Saved results to {output_csv_path}")


def detect_url_column(df):
    for candidate in ["url", "URL", "link", "Link"]:
        if candidate in df.columns:
            return candidate
    # fallback: find first column with http in any row
    for col in df.columns:
        sample = df[col].astype(str).head(50).str.contains(r"https?://")
        if sample.any():
            return col
    return None


def main():
    parser = argparse.ArgumentParser(description="Tag articles with publication date and wordcount")
    parser.add_argument("--docx", help="Path to Word docx with article list")
    parser.add_argument("--csv", help="Path to CSV input with URLs")
    parser.add_argument("-o", "--output", help="Output CSV path")
    parser.add_argument("--delay", type=float, default=2.0, help="Delay between requests")
    parser.add_argument("--heartbeat-every", type=int, default=10, help="Heartbeat frequency")
    parser.add_argument(
        "--manual-browser-retry",
        action="store_true",
        help="On HTTP/request errors, pause and let you open the URL in a browser, then retry.",
    )
    parser.add_argument(
        "--browser-cookies",
        choices=["chrome", "edge", "firefox"],
        help="Optional browser cookie source to use during manual retries.",
    )
    parser.add_argument(
        "--manual-wait-seconds",
        type=int,
        default=20,
        help="When stdin is non-interactive, wait this many seconds before manual retry.",
    )
    parser.add_argument(
        "--local-html-dir",
        help="Directory to search for saved .html/.htm files when URL fetch fails.",
    )
    parser.add_argument(
        "--local-html-path-column",
        default="local_html_path",
        help="CSV column containing per-row local HTML file path fallback.",
    )
    parser.add_argument(
        "--outline-headings-csv",
        default=OUTLINE_DEFAULT_CSV,
        help="CSV containing section/subsection headings for semantic outline assignment.",
    )
    parser.add_argument(
        "--disable-ai-outline-assignment",
        action="store_true",
        help="Disable semantic outline assignment for rows tagged #ai.",
    )
    parser.add_argument(
        "--ai-outline-min-score",
        type=float,
        default=0.12,
        help="Minimum semantic score needed to assign an outline location for #ai rows.",
    )
    parser.add_argument(
        "--ai-tag-pattern",
        default=r"(^|[\s,;])#?ai\b",
        help="Regex used to detect AI-tagged rows (default matches ai or #ai token).",
    )
    args = parser.parse_args()

    articles = []
    original_columns = None
    notes_col = None
    tags_col = None
    if args.docx:
        if not os.path.exists(args.docx):
            raise FileNotFoundError(f"Docx not found: {args.docx}")
        articles = extract_articles_and_links_from_docx(args.docx)
    elif args.csv:
        if not os.path.exists(args.csv):
            raise FileNotFoundError(f"CSV not found: {args.csv}")
        df = read_csv_with_fallbacks(args.csv)
        url_col = detect_url_column(df)
        if url_col is None:
            raise ValueError("Could not detect URL column in CSV. Add a column named 'url' or include links.")
        title_col = None
        for c in ["title", "Title", "headline"]:
            if c in df.columns:
                title_col = c
                break
        notes_col = None
        for c in NOTES_CANDIDATE_COLUMNS:
            if c in df.columns:
                notes_col = c
                break
        tags_col = None
        for c in TAG_CANDIDATE_COLUMNS:
            if c in df.columns:
                tags_col = c
                break
        original_columns = list(df.columns)
        for _, row in df.iterrows():
            article = {k: (None if pd.isna(v) else v) for k, v in row.items()}
            # ensure internal lowercase keys used by process_articles
            article["title"] = row[title_col] if title_col else ""
            article["url"] = row[url_col]
            if args.local_html_path_column not in article:
                article[args.local_html_path_column] = None
            articles.append(article)
    else:
        parser.error("Provide either --docx or --csv input")

    if not articles:
        logger.error("No articles found in input")
        return

    if args.browser_cookies and not args.manual_browser_retry:
        logger.warning("--browser-cookies is set but --manual-browser-retry is not enabled; cookies will not be used.")

    ai_outline_enabled = not args.disable_ai_outline_assignment
    outline_matcher = None
    if ai_outline_enabled:
        outline_matcher = load_outline_matcher(args.outline_headings_csv)
        if outline_matcher is None:
            ai_outline_enabled = False
            logger.warning("AI outline assignment disabled because outline matcher could not be loaded.")

    processed = process_articles(
        articles,
        delay=args.delay,
        heartbeat_every=args.heartbeat_every,
        manual_browser_retry=args.manual_browser_retry,
        browser_cookies=args.browser_cookies,
        manual_wait_seconds=args.manual_wait_seconds,
        local_html_dir=args.local_html_dir,
        local_html_path_column=args.local_html_path_column,
        outline_matcher=outline_matcher,
        ai_outline_enabled=ai_outline_enabled,
        ai_outline_min_score=args.ai_outline_min_score,
        ai_outline_tag_col=tags_col,
        notes_col=notes_col,
        ai_tag_pattern=args.ai_tag_pattern,
    )

    if args.output:
        output_path = args.output
    else:
        if args.csv:
            output_path = build_output_path(args.csv, "_raindroptagged.csv")
        else:
            output_path = os.path.join("Output files", "raindrop_tagged.csv")

    save_results_csv(processed, output_path, original_columns=original_columns, notes_col=notes_col)


if __name__ == "__main__":
    main()
