import argparse
import json
import logging
import os
import re
import time
from datetime import datetime

import pandas as pd
import requests
from bs4 import BeautifulSoup


logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def get_publication_date_from_url(url):
    """Visit URL and extract publication date."""
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

        soup = BeautifulSoup(response.content, "html.parser")

        date = (
            find_date_in_meta_tags(soup)
            or find_date_in_json_ld(soup)
            or find_date_in_time_tags(soup)
            or find_date_in_article_tags(soup)
            or find_date_in_text_patterns(soup)
        )

        if date:
            return date, "success"

        return None, "no_date_found"

    except requests.exceptions.Timeout:
        return None, "timeout"
    except requests.exceptions.RequestException:
        return None, "request_error"
    except Exception:
        return None, "error"


def find_date_in_meta_tags(soup):
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
        if not meta:
            continue

        content = meta.get("content") or meta.get("value")
        if not content:
            continue

        parsed_date = parse_date_string(content)
        if parsed_date:
            return parsed_date

    return None


def find_date_in_json_ld(soup):
    try:
        scripts = soup.find_all("script", type="application/ld+json")
        for script in scripts:
            if not script.string:
                continue

            data = json.loads(script.string)

            if isinstance(data, list):
                for item in data:
                    date = extract_date_from_json_object(item)
                    if date:
                        return date
            else:
                date = extract_date_from_json_object(data)
                if date:
                    return date
    except Exception:
        pass

    return None


def extract_date_from_json_object(obj):
    if not isinstance(obj, dict):
        return None

    date_fields = ["datePublished", "publishDate", "dateCreated", "uploadDate"]
    for field in date_fields:
        if field not in obj:
            continue

        parsed_date = parse_date_string(obj[field])
        if parsed_date:
            return parsed_date

    return None


def find_date_in_time_tags(soup):
    time_selectors = [
        "time[datetime]",
        "time[pubdate]",
        ".published-date time",
        ".publish-date time",
        ".date time",
    ]

    for selector in time_selectors:
        time_elem = soup.select_one(selector)
        if not time_elem:
            continue

        datetime_attr = time_elem.get("datetime") or time_elem.get("pubdate")
        if datetime_attr:
            parsed_date = parse_date_string(datetime_attr)
            if parsed_date:
                return parsed_date

        text = time_elem.get_text().strip()
        if text:
            parsed_date = parse_date_string(text)
            if parsed_date:
                return parsed_date

    return None


def find_date_in_article_tags(soup):
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
        elements = soup.select(selector)
        for elem in elements:
            text = elem.get_text().strip()
            if not text or len(text) >= 100:
                continue

            parsed_date = parse_date_string(text)
            if parsed_date:
                return parsed_date

    return None


def find_date_in_text_patterns(soup):
    text = soup.get_text()
    patterns = [
        r"Published:?\s*([A-Za-z]+ \d{1,2},? \d{4})",
        r"Publication Date:?\s*([A-Za-z]+ \d{1,2},? \d{4})",
        r"(\d{1,2}/\d{1,2}/\d{4})",
        r"(\d{4}-\d{2}-\d{2})",
        r"([A-Za-z]+ \d{1,2},? \d{4})",
    ]

    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            parsed_date = parse_date_string(match)
            if parsed_date:
                return parsed_date

    return None


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

    date_match = re.search(r"(\d{4}-\d{2}-\d{2})", date_str)
    if date_match:
        return date_match.group(1)

    return None


def build_output_path(input_csv_path, output_csv_path=None):
    if output_csv_path:
        return output_csv_path

    root, ext = os.path.splitext(input_csv_path)
    return f"{root}_dated{ext}"


def apply_dates_to_csv(input_csv_path, output_csv_path=None, delay=2.0, in_place=False):
    """Read URLs from column F and write publication dates to column D (note)."""
    if not os.path.exists(input_csv_path):
        raise FileNotFoundError(f"Input CSV not found: {input_csv_path}")

    logger.info(f"Reading CSV: {input_csv_path}")
    df = pd.read_csv(input_csv_path)

    if len(df.columns) < 6:
        raise ValueError("CSV must contain at least 6 columns so column F can be read.")

    note_col = df.columns[3]  # Column D
    url_col = df.columns[5]  # Column F
    status_col = "date_status"

    df[note_col] = df[note_col].astype("object")
    df[status_col] = "no_url"

    logger.info(f"Using column F for URLs: {url_col}")
    logger.info(f"Writing publication dates to column D: {note_col}")
    logger.info(f"Writing processing status to column: {status_col}")

    total_urls = int(df[url_col].notna().sum())
    processed = 0
    success = 0

    for index, url in df[url_col].items():
        if pd.isna(url):
            df.at[index, status_col] = "no_url"
            continue

        url_text = str(url).strip()
        if not url_text:
            df.at[index, status_col] = "no_url"
            continue

        processed += 1
        logger.info(f"Processing {processed}/{total_urls}: {url_text[:90]}")

        date_str, status = get_publication_date_from_url(url_text)
        df.at[index, status_col] = status
        if status == "success" and date_str:
            existing_note = df.at[index, note_col]
            if pd.isna(existing_note) or str(existing_note).strip() == "":
                df.at[index, note_col] = date_str
            else:
                df.at[index, note_col] = f"{existing_note} | publication_date: {date_str}"
            success += 1

        if processed < total_urls:
            time.sleep(delay)

    final_output = input_csv_path if in_place else build_output_path(input_csv_path, output_csv_path)
    logger.info(f"Saving updated CSV: {final_output}")
    df.to_csv(final_output, index=False)

    logger.info("Finished processing")
    logger.info(f"Rows with URL values: {total_urls}")
    logger.info(f"Rows processed: {processed}")
    logger.info(f"Dates found and written to column D: {success}")

    return final_output


def parse_args():
    parser = argparse.ArgumentParser(
        description="Find publication dates from URLs in CSV column F and write to column D (note)."
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
        help="Path to output CSV. Default adds _dated suffix to input filename.",
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
    output_path = apply_dates_to_csv(
        input_csv_path=args.input_csv,
        output_csv_path=args.output_csv,
        delay=args.delay,
        in_place=args.in_place,
    )
    print(f"Updated CSV saved to: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()