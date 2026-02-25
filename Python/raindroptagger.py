import argparse
import json
import logging
import os
import re
import sys
import time
import webbrowser
from datetime import datetime
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
        if status_code >= 400:
            logger.warning(f"HTTP {status_code} for {url}")
            return None, None, f"http_{status_code}"
        html_text = resp.text
        soup = BeautifulSoup(resp.content, "html.parser")
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
):
    total = len(articles)
    processed = 0
    success_count = 0
    start = time.perf_counter()
    for idx, article in enumerate(articles, 1):
        url = article.get("url")
        title = article.get("title") or ""
        if not url:
            article.update({"pub_date": None, "date_status": "no_url", "wordcount": None, "wc_status": "no_url", "wc_method": None})
            continue
        processed += 1
        logger.info(f"[{processed}/{total}] Fetching {url[:90]}")
        html_text, soup, fetch_status = fetch_url_once(url)
        if fetch_status != "success":
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
                        if wc_status == "success":
                            success_count += 1
                        retried_status = "success"
                    else:
                        retried_status = f"manual_{retry_status}"

            if retried_status != "success":
                article.update({"pub_date": None, "date_status": retried_status, "wordcount": None, "wc_status": retried_status, "wc_method": None})
        else:
            pub_date, date_status = get_pub_date_from_soup(soup)
            wc, wc_status, wc_method = get_wordcount_from_html(html_text, soup)
            article.update({"pub_date": pub_date, "date_status": date_status, "wordcount": wc, "wc_status": wc_status, "wc_method": wc_method})
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


def save_results_csv(articles, output_csv_path):
    df = pd.DataFrame(articles)
    # normalize column names
    desired = ["title", "url", "pub_date", "date_status", "wordcount", "wc_status", "wc_method"]
    for col in desired:
        if col not in df.columns:
            df[col] = None
    df = df[desired]
    output_dir = os.path.dirname(output_csv_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    df.to_csv(output_csv_path, index=False)
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
    args = parser.parse_args()

    articles = []
    if args.docx:
        if not os.path.exists(args.docx):
            raise FileNotFoundError(f"Docx not found: {args.docx}")
        articles = extract_articles_and_links_from_docx(args.docx)
    elif args.csv:
        if not os.path.exists(args.csv):
            raise FileNotFoundError(f"CSV not found: {args.csv}")
        df = pd.read_csv(args.csv)
        url_col = detect_url_column(df)
        if url_col is None:
            raise ValueError("Could not detect URL column in CSV. Add a column named 'url' or include links.")
        title_col = None
        for c in ["title", "Title", "headline"]:
            if c in df.columns:
                title_col = c
                break
        for _, row in df.iterrows():
            url = row[url_col]
            title = row[title_col] if title_col else ""
            articles.append({"title": title, "url": url})
    else:
        parser.error("Provide either --docx or --csv input")

    if not articles:
        logger.error("No articles found in input")
        return

    if args.browser_cookies and not args.manual_browser_retry:
        logger.warning("--browser-cookies is set but --manual-browser-retry is not enabled; cookies will not be used.")

    processed = process_articles(
        articles,
        delay=args.delay,
        heartbeat_every=args.heartbeat_every,
        manual_browser_retry=args.manual_browser_retry,
        browser_cookies=args.browser_cookies,
        manual_wait_seconds=args.manual_wait_seconds,
    )

    if args.output:
        output_path = args.output
    else:
        if args.csv:
            output_path = build_output_path(args.csv, "_raindroptagged.csv")
        else:
            output_path = os.path.join("Output files", "raindrop_tagged.csv")

    save_results_csv(processed, output_path)


if __name__ == "__main__":
    main()
