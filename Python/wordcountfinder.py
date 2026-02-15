import argparse
import json
import logging
import os
import re
import time

import pandas as pd
import requests
from bs4 import BeautifulSoup

try:
    import trafilatura
except Exception:
    trafilatura = None


logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


BOILERPLATE_TAGS = [
    "script",
    "style",
    "noscript",
    "svg",
    "iframe",
    "nav",
    "aside",
    "footer",
    "form",
]

BOILERPLATE_CLASS_ID_PATTERN = re.compile(
    r"(^|[-_\s])(ad|ads|advert|advertisement|sponsor|promo|related|newsletter|"
    r"footer|sidebar|share|social|cookie|banner|recommend|trending|outbrain|"
    r"taboola)($|[-_\s])",
    re.IGNORECASE,
)


def normalize_text(text):
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def count_words(text):
    words = re.findall(r"\b[\w'-]+\b", text)
    return len(words)


def extract_json_ld_article_text(html_text):
    """Extract article-like text from JSON-LD fields such as articleBody."""
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
        extracted = trafilatura.extract(
            html_text,
            include_comments=False,
            include_tables=False,
            favor_precision=True,
        )
        extracted = normalize_text(extracted)
        if extracted:
            return extracted, "trafilatura"
    except Exception:
        pass

    return "", "trafilatura_failed"


def get_article_word_count(url):
    """Visit URL and estimate article word count from page text."""
    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/91.0.4472.124 Safari/537.36"
            )
        }

        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        html_text = response.text
        candidates = []

        jsonld_text, jsonld_method = extract_json_ld_article_text(html_text)
        if jsonld_text:
            candidates.append((jsonld_text, jsonld_method))

        tr_text, tr_method = extract_main_text_with_trafilatura(html_text)
        if tr_text:
            candidates.append((tr_text, tr_method))

        soup = BeautifulSoup(response.content, "html.parser")

        article_node = soup.find("article")
        if article_node:
            article_p_text = normalize_text(
                " ".join(p.get_text(" ", strip=True) for p in article_node.find_all("p"))
            )
            if article_p_text:
                candidates.append((article_p_text, "article_p"))

            article_raw_text = normalize_text(article_node.get_text(" ", strip=True))
            if article_raw_text:
                candidates.append((article_raw_text, "article_tag"))

        main_node = soup.find("main")
        if main_node:
            main_p_text = normalize_text(
                " ".join(p.get_text(" ", strip=True) for p in main_node.find_all("p"))
            )
            if main_p_text:
                candidates.append((main_p_text, "main_p"))

        all_p_text = normalize_text(
            " ".join(p.get_text(" ", strip=True) for p in soup.find_all("p"))
        )
        if all_p_text:
            candidates.append((all_p_text, "all_p"))

        bs4_text, bs4_method = extract_main_text_with_bs4(soup)
        if bs4_text:
            candidates.append((bs4_text, bs4_method))

        if soup.body:
            body_text = normalize_text(soup.body.get_text(" ", strip=True))
            if body_text:
                candidates.append((body_text, "body_full"))

        scored = []
        for candidate_text, candidate_method in candidates:
            wc = count_words(candidate_text)
            if wc > 0:
                scored.append((wc, candidate_method))

        if not scored:
            return None, "no_text_found", "no_candidate_text"

        wc_by_method = {method: wc for wc, method in scored}

        for preferred_method in ["article_p", "main_p", "jsonld", "trafilatura", "article_tag"]:
            preferred_wc = wc_by_method.get(preferred_method, 0)
            if preferred_wc >= 120:
                return preferred_wc, "success", preferred_method

        non_fullpage = [(wc, method) for wc, method in scored if method not in {"all_p", "body_full"}]
        if non_fullpage:
            best_wc, best_method = max(non_fullpage, key=lambda item: item[0])
            return best_wc, "success", best_method

        ranked = sorted(scored, key=lambda item: item[0], reverse=True)
        if len(ranked) >= 2 and ranked[0][0] > int(ranked[1][0] * 1.6) and (ranked[0][0] - ranked[1][0]) > 500:
            best_wc, best_method = ranked[1]
        else:
            best_wc, best_method = ranked[0]

        return best_wc, "success", best_method

    except requests.exceptions.Timeout:
        return None, "timeout", "request_timeout"
    except requests.exceptions.RequestException:
        return None, "request_error", "request_error"
    except Exception as exc:
        logger.warning(f"Unexpected parse error for {url}: {type(exc).__name__}: {exc}")
        return None, "error", f"unexpected_error_{type(exc).__name__}"


def build_output_path(input_csv_path, output_csv_path=None):
    if output_csv_path:
        return output_csv_path

    root, ext = os.path.splitext(input_csv_path)
    return f"{root}_wordcount{ext}"


def apply_word_counts_to_csv(
    input_csv_path, output_csv_path=None, delay=2.0, in_place=False
):
    """Read URLs from column F and write word counts to column D (note)."""
    if not os.path.exists(input_csv_path):
        raise FileNotFoundError(f"Input CSV not found: {input_csv_path}")

    logger.info(f"Reading CSV: {input_csv_path}")
    df = pd.read_csv(input_csv_path)

    if len(df.columns) < 6:
        raise ValueError("CSV must contain at least 6 columns so column F can be read.")

    note_col = df.columns[3]  # Column D
    url_col = df.columns[5]  # Column F
    status_col = "wordcount_status"
    method_col = "wordcount_method"

    df[note_col] = df[note_col].astype("object")
    df[status_col] = "no_url"
    df[method_col] = "not_processed"

    logger.info(f"Using column F for URLs: {url_col}")
    logger.info(f"Writing word counts to column D: {note_col}")
    logger.info(f"Writing processing status to column: {status_col}")
    logger.info(f"Writing extraction method to column: {method_col}")

    total_urls = int(df[url_col].notna().sum())
    processed = 0
    success = 0

    for index, url in df[url_col].items():
        if pd.isna(url):
            df.at[index, status_col] = "no_url"
            df.at[index, method_col] = "no_url"
            continue

        url_text = str(url).strip()
        if not url_text:
            df.at[index, status_col] = "no_url"
            df.at[index, method_col] = "no_url"
            continue

        processed += 1
        logger.info(f"Processing {processed}/{total_urls}: {url_text[:90]}")

        word_count, status, method = get_article_word_count(url_text)
        df.at[index, status_col] = status
        df.at[index, method_col] = method

        if status == "success" and word_count is not None:
            existing_note = df.at[index, note_col]
            if pd.isna(existing_note) or str(existing_note).strip() == "":
                df.at[index, note_col] = str(word_count)
            else:
                df.at[index, note_col] = f"{existing_note} | word_count: {word_count}"
            success += 1

        if processed < total_urls:
            time.sleep(delay)

    final_output = (
        input_csv_path if in_place else build_output_path(input_csv_path, output_csv_path)
    )
    logger.info(f"Saving updated CSV: {final_output}")
    df.to_csv(final_output, index=False)

    logger.info("Finished processing")
    logger.info(f"Rows with URL values: {total_urls}")
    logger.info(f"Rows processed: {processed}")
    logger.info(f"Word counts found and written to column D: {success}")

    return final_output


def parse_args():
    parser = argparse.ArgumentParser(
        description="Find article word counts from URLs in CSV column F and write to column D (note)."
    )
    parser.add_argument(
        "input_csv",
        nargs="?",
        default=r".\Source files\datepicker.csv",
        help="Path to the input CSV file.",
    )
    parser.add_argument(
        "-o",
        "--output",
        dest="output_csv",
        default=None,
        help="Path to output CSV. Default adds _wordcount suffix to input filename.",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Delay in seconds between URL requests.",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite the input CSV instead of creating a new output file.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    output_path = apply_word_counts_to_csv(
        input_csv_path=args.input_csv,
        output_csv_path=args.output_csv,
        delay=args.delay,
        in_place=args.in_place,
    )
    print(f"Updated CSV saved to: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()
