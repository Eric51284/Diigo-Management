"""
Extract section/subsection outline headings from the Capstone HTML file.

Outputs a CSV that preserves hierarchy by storing:
- level (1=section, 2=subsection)
- parent/child ids
- display order indexes
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

from bs4 import BeautifulSoup


def clean_section_title(summary_tag) -> tuple[str, str]:
    """Return (section_title_without_count, article_count_text)."""
    if summary_tag is None:
        return "", ""

    count_span = summary_tag.find("span", class_="art-count")
    article_count = ""
    if count_span:
        article_count = count_span.get_text(" ", strip=True).strip("()")
        count_span.extract()

    title = summary_tag.get_text(" ", strip=True)
    return title, article_count


def parse_outline(html_path: Path) -> list[dict[str, str | int]]:
    with html_path.open("r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    rows: list[dict[str, str | int]] = []

    for section_index, sec in enumerate(soup.select("details.sec"), start=1):
        sec_summary = sec.find("summary", recursive=False)
        sec_id = sec.get("id", "")
        sec_title, sec_count = clean_section_title(sec_summary)

        rows.append(
            {
                "level": 1,
                "section_index": section_index,
                "subsection_index": "",
                "section_id": sec_id,
                "subsection_id": "",
                "parent_id": "",
                "title": sec_title,
                "article_count": sec_count,
            }
        )

        for subsection_index, sub in enumerate(sec.select("details.sub"), start=1):
            sub_summary = sub.find("summary", recursive=False)
            sub_id = sub.get("id", "")
            sub_title = sub_summary.get_text(" ", strip=True) if sub_summary else ""

            rows.append(
                {
                    "level": 2,
                    "section_index": section_index,
                    "subsection_index": subsection_index,
                    "section_id": sec_id,
                    "subsection_id": sub_id,
                    "parent_id": sec_id,
                    "title": sub_title,
                    "article_count": "",
                }
            )

    return rows


def write_csv(rows: list[dict[str, str | int]], output_path: Path) -> None:
    fieldnames = [
        "level",
        "section_index",
        "subsection_index",
        "section_id",
        "subsection_id",
        "parent_id",
        "title",
        "article_count",
    ]

    with output_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract outline headings to CSV.")
    parser.add_argument("html_path", type=Path, help="Path to source HTML file")
    parser.add_argument("output_csv", type=Path, help="Path to output CSV file")
    args = parser.parse_args()

    rows = parse_outline(args.html_path)
    write_csv(rows, args.output_csv)
    print(f"Wrote {len(rows)} rows to {args.output_csv}")


if __name__ == "__main__":
    main()
