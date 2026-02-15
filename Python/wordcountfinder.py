import argparse
import logging
import os
import re
import time

import pandas as pd
import requests
from bs4 import BeautifulSoup


logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


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

        soup = BeautifulSoup(response.content, "html.parser")

        for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
            tag.decompose()

        article_node = soup.find("article")
        if article_node:
            text = article_node.get_text(" ", strip=True)
        elif soup.body:
            text = soup.body.get_text(" ", strip=True)
        else:
            text = soup.get_text(" ", strip=True)

        text = re.sub(r"\s+", " ", text).strip()
        words = re.findall(r"\b[\w'-]+\b", text)

        if not words:
            return None, "no_text_found"

        return len(words), "success"

    except requests.exceptions.Timeout:
        return None, "timeout"
    except requests.exceptions.RequestException:
        return None, "request_error"
    except Exception:
        return None, "error"


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

    df[note_col] = df[note_col].astype("object")
    df[status_col] = "no_url"

    logger.info(f"Using column F for URLs: {url_col}")
    logger.info(f"Writing word counts to column D: {note_col}")
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

        word_count, status = get_article_word_count(url_text)
        df.at[index, status_col] = status

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
