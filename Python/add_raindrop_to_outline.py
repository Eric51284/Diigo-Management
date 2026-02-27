"""
add_raindrop_to_outline.py
──────────────────────────
Adds articles from a Raindrop export CSV into a structured HTML outline.
By default fetches each article URL to extract title, description, headings
and body text for improved section classification.  Falls back to CSV fields
(title, tags, note, excerpt) when a fetch fails.

Usage examples
  python add_raindrop_to_outline.py               # uses defaults, fetches URLs
  python add_raindrop_to_outline.py --no-fetch    # CSV signals only
  python add_raindrop_to_outline.py --dry-run     # classify + report, no output
  python add_raindrop_to_outline.py --verbose     # show per-article placement
"""

import argparse
import logging
import re
import time
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "can", "did", "do", "for",
    "from", "has", "have", "how", "in", "is", "it", "its", "of", "on", "or",
    "our", "says", "so", "than", "that", "the", "this", "to", "up", "us",
    "was", "we", "what", "when", "where", "which", "who", "why", "will", "with",
    "you", "your", "ai", "artificial", "intelligence",
}

SECTION_HINTS: Dict[str, set] = {
    "s1": {
        "llm", "llms", "model", "models", "hallucination", "hallucinations", "rag",
        "benchmark", "benchmarks", "benchmarking", "architecture", "training",
        "inference", "reasoning", "transformer", "diffusion", "embedding", "token",
        "tokens", "tokenizer", "parameter", "parameters", "pretraining", "finetuning",
        "multimodal", "vision", "image", "video", "generation", "language",
    },
    "s2": {
        "tool", "tools", "workflow", "prompt", "prompting", "coding", "code",
        "productivity", "search", "assistant", "copilot", "plugin", "browser",
        "extension", "summarize", "summarization", "ocr", "document", "spreadsheet",
        "excel", "pdf", "notebooklm", "perplexity", "replit", "writing",
    },
    "s3": {
        "agent", "agents", "agentic", "autonomous", "orchestration", "multiagent",
        "multi-agent", "openclaw", "moltbook", "workflow", "automation", "automate",
        "deploy", "deployment", "mcp", "protocol",
    },
    "s4": {
        "safety", "ethics", "ethical", "alignment", "misuse", "policy", "legal",
        "regulatory", "regulation", "law", "risk", "governance", "security",
        "privacy", "bias", "misinformation", "deepfake", "harm", "danger",
        "threat", "censor", "censorship", "rights",
    },
    "s5": {
        "education", "educational", "student", "students", "teaching", "learning",
        "classroom", "school", "university", "college", "course", "curriculum",
        "academic", "professor", "homework",
    },
    "s6": {
        "cognitive", "cognition", "psychology", "psychological", "neuroscience",
        "brain", "behavior", "behaviour", "human", "mind", "consciousness",
        "emotion", "mental", "perception", "memory",
    },
    "s7": {
        "economy", "economic", "jobs", "job", "labor", "labour", "work", "workforce",
        "society", "societal", "social", "business", "industry", "industries",
        "advertising", "corporate", "company", "companies", "inequality", "wage",
        "employment", "unemployment", "copyright", "art", "creative",
    },
    "s8": {
        "science", "biology", "health", "medicine", "medical", "climate",
        "environment", "environmental", "research", "discovery", "physics",
        "chemistry", "drug", "drugs", "genomics", "protein", "astronomy",
    },
}

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


@dataclass
class ArticleSignals:
    """All text signals gathered for one article."""
    title: str = ""
    tags: str = ""
    note: str = ""
    excerpt: str = ""
    fetched_title: str = ""
    fetched_description: str = ""
    fetched_keywords: str = ""
    fetched_headings: str = ""
    fetched_body: str = ""
    fetch_ok: bool = False

    def combined_text(self) -> str:
        parts = [
            self.fetched_title or self.title,
            self.tags,
            self.note,
            self.excerpt,
            self.fetched_description,
            self.fetched_keywords,
            self.fetched_headings,
        ]
        return " ".join(p for p in parts if p)

    def body_text(self) -> str:
        return self.fetched_body

    def display_title(self) -> str:
        return self.title or self.fetched_title


@dataclass
class SubsectionProfile:
    sub_id: str
    section_id: str
    summary_text: str
    summary_tokens: set
    title_tokens: Counter
    ul_node: object


def tokenize(text: str) -> List[str]:
    words = re.findall(r"[a-z0-9][a-z0-9\-']+", (text or "").lower())
    return [w for w in words if len(w) > 2 and w not in STOPWORDS]


def _meta_content(soup: BeautifulSoup, *selectors: str) -> str:
    for sel in selectors:
        tag = soup.select_one(sel)
        if tag:
            return (tag.get("content") or tag.get("value") or "").strip()
    return ""


def fetch_article_signals(url: str, session: requests.Session, timeout: int = 12) -> ArticleSignals:
    """Fetch a URL and extract classification signals."""
    sig = ArticleSignals()
    try:
        resp = session.get(url, headers=REQUEST_HEADERS, timeout=timeout, allow_redirects=True)
        if resp.status_code != 200:
            logger.debug("HTTP %s for %s", resp.status_code, url)
            return sig
        if "html" not in resp.headers.get("content-type", ""):
            return sig

        soup = BeautifulSoup(resp.text, "html.parser")

        sig.fetched_title = (
            _meta_content(soup, 'meta[property="og:title"]')
            or (soup.title.get_text(" ", strip=True) if soup.title else "")
        )
        sig.fetched_description = _meta_content(
            soup,
            'meta[property="og:description"]',
            'meta[name="description"]',
            'meta[name="Description"]',
        )
        sig.fetched_keywords = " ".join(filter(None, [
            _meta_content(soup, 'meta[name="keywords"]', 'meta[name="Keywords"]'),
            _meta_content(soup, 'meta[property="article:section"]'),
            _meta_content(soup, 'meta[property="article:tag"]'),
        ]))

        headings = []
        for tag in soup.select("h1, h2"):
            txt = tag.get_text(" ", strip=True)
            if txt and len(txt) < 200:
                headings.append(txt)
        sig.fetched_headings = " ".join(headings[:6])

        body_node = soup.find("article") or soup.find("main") or soup.body
        if body_node:
            for dead in body_node.find_all(["script", "style", "nav", "footer", "aside", "header"]):
                dead.decompose()
            raw = re.sub(r"\s+", " ", body_node.get_text(" ", strip=True))
            sig.fetched_body = " ".join(raw.split()[:600])

        sig.fetch_ok = True
    except Exception as exc:
        logger.debug("Fetch failed for %s: %s", url, exc)
    return sig


def parse_pub_date(row: pd.Series) -> str:
    note = str(row.get("note", "") or "")
    m = re.search(r"pub\s*:\s*(\d{4}-\d{2}-\d{2})", note)
    if m:
        return m.group(1)

    for field in ("publication_date", "pub_date", "date"):
        value = row.get(field)
        if pd.notna(value):
            text = str(value)
            m = re.search(r"(\d{4}-\d{2}-\d{2})", text)
            if m:
                return m.group(1)

    created = row.get("created")
    if pd.notna(created):
        text = str(created)
        m = re.search(r"(\d{4}-\d{2}-\d{2})", text)
        if m:
            return m.group(1)

    return "1900-01-01"


def parse_wordcount(row: pd.Series) -> Optional[int]:
    note = str(row.get("note", "") or "")
    m = re.search(r"wordcount\s*:\s*(\d+)", note, flags=re.IGNORECASE)
    if m:
        return int(m.group(1))

    for field in ("wordcount", "word_count"):
        value = row.get(field)
        if pd.notna(value):
            try:
                return int(float(value))
            except Exception:
                continue

    return None


def build_profiles(soup: BeautifulSoup) -> List[SubsectionProfile]:
    profiles: List[SubsectionProfile] = []
    for sub in soup.select("details.sub"):
        sub_id = sub.get("id", "")
        section = sub.find_parent("details", class_="sec")
        section_id = section.get("id", "") if section else ""
        summary = sub.find("summary")
        summary_text = summary.get_text(" ", strip=True) if summary else ""
        summary_tokens = set(tokenize(summary_text))

        title_counter: Counter = Counter()
        for a in sub.select("ul.arts li a"):
            title_counter.update(tokenize(a.get_text(" ", strip=True)))

        ul_node = sub.select_one("div.sub-body ul.arts")
        if ul_node is None:
            continue

        profiles.append(SubsectionProfile(
            sub_id=sub_id,
            section_id=section_id,
            summary_text=summary_text,
            summary_tokens=summary_tokens,
            title_tokens=title_counter,
            ul_node=ul_node,
        ))
    return profiles


def score_subsection(
    profile: SubsectionProfile,
    main_tokens: set,
    tag_tokens: set,
    body_tokens: set,
) -> float:
    section_hints = SECTION_HINTS.get(profile.section_id, set())
    heading_overlap_main = len(main_tokens & profile.summary_tokens)
    heading_overlap_tags = len(tag_tokens & profile.summary_tokens)
    hint_overlap_main = len(main_tokens & section_hints)
    hint_overlap_tags = len(tag_tokens & section_hints)
    hint_overlap_body = len(body_tokens & section_hints)
    title_cooccur = sum(min(profile.title_tokens.get(t, 0), 3) for t in main_tokens)
    body_heading_overlap = len(body_tokens & profile.summary_tokens)
    return (
        5.0 * heading_overlap_tags
        + 3.0 * heading_overlap_main
        + 4.0 * hint_overlap_tags
        + 3.0 * hint_overlap_main
        + 2.5 * hint_overlap_body
        + 1.5 * body_heading_overlap
        + 0.6 * title_cooccur
    )


def best_subsections(
    profiles: List[SubsectionProfile],
    signals: ArticleSignals,
    max_sections: int,
) -> List[SubsectionProfile]:
    main_tokens = set(tokenize(signals.combined_text()))
    tag_tokens = set(tokenize(signals.tags))
    body_tokens = set(tokenize(signals.body_text()))

    if not main_tokens and not tag_tokens and not body_tokens:
        return [profiles[0]] if profiles else []

    scored: List[Tuple[float, SubsectionProfile]] = [
        (score_subsection(p, main_tokens, tag_tokens, body_tokens), p)
        for p in profiles
    ]
    scored.sort(key=lambda x: x[0], reverse=True)
    if not scored:
        return []

    top_score = scored[0][0]
    if top_score <= 0:
        return [scored[0][1]]

    chosen = [scored[0][1]]
    if max_sections > 1:
        for score, profile in scored[1:]:
            if len(chosen) >= max_sections:
                break
            if score >= top_score * 0.88:
                chosen.append(profile)
    return chosen


def create_article_li(
    soup: BeautifulSoup,
    title: str,
    url: str,
    pub_date: str,
    wordcount: Optional[int],
    cross_ref: Optional[str] = None,
):
    li = soup.new_tag("li")

    a = soup.new_tag("a", href=url, target="_blank", rel="noopener")
    a.string = title
    li.append(a)

    meta = soup.new_tag("span", attrs={"class": "meta"})

    date_badge = soup.new_tag("span", attrs={"class": "bd bd-d"})
    date_badge.string = pub_date
    meta.append(date_badge)

    if wordcount is not None:
        wc_badge = soup.new_tag("span", attrs={"class": "bd bd-w"})
        wc_badge.string = f"{wordcount:,} wds"
        meta.append(wc_badge)

    if cross_ref:
        xr = soup.new_tag("span", attrs={"class": "xr"})
        xr.string = cross_ref
        meta.append(xr)

    li.append(meta)
    return li


def parse_li_date(li_node) -> datetime:
    date_span = li_node.select_one("span.bd.bd-d")
    if date_span:
        txt = date_span.get_text(strip=True)
        try:
            return datetime.strptime(txt, "%Y-%m-%d")
        except Exception:
            pass
    return datetime(1900, 1, 1)


def sort_all_subsections_by_date(soup: BeautifulSoup):
    for ul in soup.select("details.sub ul.arts"):
        items = [li for li in ul.find_all("li", recursive=False)]
        items.sort(key=parse_li_date, reverse=True)
        ul.clear()
        for li in items:
            ul.append(li)


def refresh_counts(soup: BeautifulSoup) -> None:
    for sec in soup.select("details.sec"):
        summary = sec.find("summary")
        if not summary:
            continue
        span = summary.select_one("span.art-count")
        if not span:
            continue
        span.string = f"({len(sec.select('details.sub ul.arts > li'))} articles)"

    total = len(soup.select("details.sub ul.arts > li"))
    secs = len(soup.select("details.sec"))
    subs = len(soup.select("details.sub"))
    hp = soup.select_one("header p")
    if hp:
        hp.string = (
            f"{total} articles · {secs} sections · {subs} subsections"
            " · links, publication dates & word counts from source CSV"
        )


def build_session() -> requests.Session:
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    sess = requests.Session()
    retry = Retry(total=2, backoff_factor=0.5, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry)
    sess.mount("http://", adapter)
    sess.mount("https://", adapter)
    return sess


def update_outline(
    html_path: Path,
    csv_path: Path,
    output_path: Path,
    max_sections_per_article: int,
    dry_run: bool,
    fetch_urls: bool,
    fetch_delay: float,
    fetch_timeout: int,
    verbose: bool,
) -> None:
    soup = BeautifulSoup(html_path.read_text(encoding="utf-8"), "html.parser")
    profiles = build_profiles(soup)
    if not profiles:
        raise RuntimeError("No subsection profiles found in outline HTML.")

    df = pd.read_csv(csv_path)
    existing_urls = {
        (a.get("href") or "").strip()
        for a in soup.select("details.sub ul.arts li a[href]")
    }

    session = build_session() if fetch_urls else None
    inserted = skipped = fetch_ok_count = fetch_fail_count = 0

    for _, row in df.iterrows():
        title = str(row.get("title", "") or "").strip()
        url = str(row.get("url", "") or "").strip()
        if not title or not url:
            continue
        if url in existing_urls:
            skipped += 1
            continue

        signals = ArticleSignals(
            title=title,
            tags=str(row.get("tags", "") or ""),
            note=str(row.get("note", "") or ""),
            excerpt=str(row.get("excerpt", "") or ""),
        )

        if fetch_urls and session is not None:
            fetched = fetch_article_signals(url, session, timeout=fetch_timeout)
            signals.fetched_title = fetched.fetched_title
            signals.fetched_description = fetched.fetched_description
            signals.fetched_keywords = fetched.fetched_keywords
            signals.fetched_headings = fetched.fetched_headings
            signals.fetched_body = fetched.fetched_body
            signals.fetch_ok = fetched.fetch_ok
            if fetched.fetch_ok:
                fetch_ok_count += 1
            else:
                fetch_fail_count += 1
            if fetch_delay > 0:
                time.sleep(fetch_delay)

        targets = best_subsections(profiles, signals, max_sections=max_sections_per_article)
        if not targets:
            continue

        pub_date = parse_pub_date(row)
        wc = parse_wordcount(row)
        display_title = signals.display_title()

        if verbose:
            placements = ", ".join(t.sub_id for t in targets)
            fetch_note = "(fetched)" if signals.fetch_ok else "(csv-only)"
            print(f"  [{placements}] {fetch_note} {display_title[:80]}")

        for i, target in enumerate(targets):
            cross_ref = None
            if i == 0 and len(targets) > 1:
                refs = [f"§{t.section_id[1:].upper()}-{t.sub_id[-1].upper()}" for t in targets[1:]]
                cross_ref = f"→ also in {', '.join(refs)}"
            li = create_article_li(soup, display_title, url, pub_date, wc, cross_ref)
            target.ul_node.append(li)

        existing_urls.add(url)
        inserted += 1

    sort_all_subsections_by_date(soup)
    refresh_counts(soup)

    print(f"\nInserted  : {inserted} new articles")
    print(f"Skipped   : {skipped} duplicate URLs")
    if fetch_urls:
        print(f"Fetched   : {fetch_ok_count} OK, {fetch_fail_count} failed")

    if dry_run:
        print("Dry run — no output written.")
        return

    output_path.write_text(str(soup), encoding="utf-8")
    print(f"Output    : {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Add Raindrop CSV articles into a structured HTML outline."
    )
    parser.add_argument(
        "--html",
        default="../Output files/Capstone AI articles.html",
        help="Input outline HTML (default: Capstone AI articles.html)",
    )
    parser.add_argument(
        "--csv",
        default="../Output files/raindrop_for_outline.csv",
        help="Input Raindrop CSV (default: raindrop_for_outline.csv)",
    )
    parser.add_argument(
        "-o", "--output",
        default="../Output files/Capstone AI articles.updated.html",
        help="Output HTML path",
    )
    parser.add_argument(
        "--max-sections-per-article", type=int, default=1,
        help="Max subsection placements per article (default 1)",
    )
    parser.add_argument(
        "--no-fetch", dest="fetch_urls", action="store_false", default=True,
        help="Skip URL fetching; classify using CSV fields only",
    )
    parser.add_argument(
        "--fetch-delay", type=float, default=1.5,
        help="Seconds to wait between fetches (default 1.5)",
    )
    parser.add_argument(
        "--fetch-timeout", type=int, default=12,
        help="HTTP timeout per request in seconds (default 12)",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Classify and report placements without writing output",
    )
    parser.add_argument(
        "--verbose", action="store_true",
        help="Print per-article placement decisions",
    )

    args = parser.parse_args()
    html_path = Path(args.html)
    csv_path = Path(args.csv)
    output_path = Path(args.output)

    if not html_path.exists():
        raise FileNotFoundError(f"HTML not found: {html_path}")
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    update_outline(
        html_path=html_path,
        csv_path=csv_path,
        output_path=output_path,
        max_sections_per_article=max(1, args.max_sections_per_article),
        dry_run=args.dry_run,
        fetch_urls=args.fetch_urls,
        fetch_delay=args.fetch_delay,
        fetch_timeout=args.fetch_timeout,
        verbose=args.verbose,
    )


if __name__ == "__main__":
    main()
