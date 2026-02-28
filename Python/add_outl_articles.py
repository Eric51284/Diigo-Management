"""
add_outl_articles.py
--------------------
Reads new articles from Source files/outl.csv and inserts them into the
appropriate subsection of Output files/Capstone AI articles.html.

Outline section is determined by the '_outl:' tag in the 'tags' column,
e.g. '_outl:VII-A' → subsection id 's7a'.

Publication date is taken from 'pub:YYYY-MM-DD' in the 'note' column.
Word count is taken from 'wordcount:NNN' in the 'note' column.

Articles are inserted in date-descending order within each subsection.
Duplicate URLs (same URL already in that subsection) are skipped.
Section article counts and the header total are updated automatically.
"""

import csv
import os
import re
from datetime import date

from bs4 import BeautifulSoup

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR = r'c:\Users\evanzant\Dropbox (Personal)\Projects\Diigo Management'
HTML_PATH = os.path.join(BASE_DIR, 'Output files', 'Capstone AI articles.html')
CSV_PATH  = os.path.join(BASE_DIR, 'Source files', 'outl.csv')

# ── Roman-numeral → integer ────────────────────────────────────────────────────
ROMAN = {
    'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5,
    'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10,
}


def roman_to_int(r: str):
    return ROMAN.get(r.upper())


def outl_to_sub_id(outl_tag: str):
    """
    Convert a tag like '_outl:VIII-C' to a subsection HTML id like 's8c'.
    Returns None if the tag doesn't match the expected pattern.
    """
    val = outl_tag.strip()
    if not val.startswith('_outl:'):
        return None
    val = val[6:].strip()                          # remove '_outl:' prefix
    m = re.fullmatch(r'([IVX]+)-([A-Z])', val, re.IGNORECASE)
    if not m:
        return None
    num = roman_to_int(m.group(1))
    if num is None:
        return None
    return f's{num}{m.group(2).lower()}'


# ── Date helpers ───────────────────────────────────────────────────────────────
def parse_date(s: str):
    """Return a date object for YYYY-MM-DD strings, or None for anything else."""
    try:
        return date.fromisoformat(s)
    except (ValueError, TypeError, AttributeError):
        return None


def li_date(li) -> date | None:
    """Extract the publication date from an existing article <li>."""
    badge = li.find('span', class_='bd-d')
    if badge:
        return parse_date(badge.get_text(strip=True))
    return None


# ── HTML element builder ───────────────────────────────────────────────────────
def make_li(soup, title: str, url: str, pub_date: str | None, wordcount: int | None):
    """
    Build an article <li> matching the existing pattern:

      <li>
        <a href="URL" rel="noopener" target="_blank">TITLE</a>
        <span class="meta">
          <span class="bd bd-d">DATE</span>
          <span class="bd bd-w">NNN wds</span>   ← optional
        </span>
      </li>
    """
    li   = soup.new_tag('li')
    a    = soup.new_tag('a', href=url, target='_blank')
    a['rel'] = 'noopener'
    a.string = title
    li.append(a)

    meta = soup.new_tag('span', **{'class': 'meta'})

    date_span = soup.new_tag('span', **{'class': 'bd bd-d'})
    date_span.string = pub_date if pub_date else ''
    meta.append(date_span)

    if wordcount is not None:
        wc_span = soup.new_tag('span', **{'class': 'bd bd-w'})
        wc_span.string = f'{wordcount:,} wds'
        meta.append(wc_span)

    li.append(meta)
    return li


# ── CSV parsing ────────────────────────────────────────────────────────────────
def load_csv(csv_path: str) -> list[dict]:
    """
    Returns a list of dicts with keys:
      title, url, pub_date (str or None), wordcount (int or None), sub_id (str)
    One entry per _outl: tag per row.
    """
    articles = []
    with open(csv_path, newline='', encoding='utf-8') as fh:
        for row in csv.DictReader(fh):
            title    = (row.get('title') or '').strip()
            url      = (row.get('url')   or '').strip()
            note     = (row.get('note')  or '').strip()
            tags_raw = (row.get('tags')  or '').strip()

            if not title or not url:
                continue

            # wordcount from note field
            wc_m  = re.search(r'wordcount:(\d[\d,]*)', note)
            wordcount = int(wc_m.group(1).replace(',', '')) if wc_m else None

            # pub date from note field
            pub_m    = re.search(r'pub:(\d{4}-\d{2}-\d{2})', note)
            pub_date = pub_m.group(1) if pub_m else None

            # collect _outl: tags
            for tag in (t.strip() for t in tags_raw.split(',')):
                sub_id = outl_to_sub_id(tag)
                if sub_id:
                    articles.append(dict(
                        title=title,
                        url=url,
                        pub_date=pub_date,
                        wordcount=wordcount,
                        sub_id=sub_id,
                    ))
    return articles


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    # ── Load HTML ──────────────────────────────────────────────────────────────
    with open(HTML_PATH, encoding='utf-8') as fh:
        soup = BeautifulSoup(fh.read(), 'html.parser')

    # ── Load CSV ───────────────────────────────────────────────────────────────
    new_articles = load_csv(CSV_PATH)
    if not new_articles:
        print('No articles found in CSV.')
        return

    added = 0
    skipped = 0
    modified_sec_ids: set[str] = set()

    for art in new_articles:
        sub_id = art['sub_id']

        # Find subsection
        sub_el = soup.find('details', id=sub_id)
        if not sub_el:
            print(f'  WARNING: subsection "{sub_id}" not found  — skipping "{art["title"]}"')
            skipped += 1
            continue

        # Find article list
        ul = sub_el.find('ul', class_='arts')
        if not ul:
            print(f'  WARNING: no <ul class="arts"> in "{sub_id}" — skipping "{art["title"]}"')
            skipped += 1
            continue

        # Duplicate-URL check
        existing_urls = {a['href'] for a in ul.find_all('a', href=True)}
        if art['url'] in existing_urls:
            print(f'  SKIP (already present in {sub_id}): {art["title"]}')
            skipped += 1
            continue

        # Build new <li>
        new_li   = make_li(soup, art['title'], art['url'], art['pub_date'], art['wordcount'])
        new_date = parse_date(art['pub_date'])

        # Insert at correct position (descending date order)
        inserted = False
        if new_date is not None:
            for existing_li in ul.find_all('li'):
                ex_date = li_date(existing_li)
                # Insert before first item that is undated OR older than new article
                if ex_date is None or new_date > ex_date:
                    existing_li.insert_before(new_li)
                    inserted = True
                    break
        if not inserted:
            ul.append(new_li)

        # Track which top-level section was changed
        sec_el = sub_el.find_parent('details', class_='sec')
        if sec_el:
            modified_sec_ids.add(sec_el.get('id', ''))

        added += 1
        print(f'  Added to {sub_id}: {art["title"]}')

    # ── Update per-section article counts ─────────────────────────────────────
    for sec_id in modified_sec_ids:
        sec_el = soup.find('details', id=sec_id)
        if not sec_el:
            continue
        count = sum(len(ul.find_all('li')) for ul in sec_el.find_all('ul', class_='arts'))
        summary = sec_el.find('summary', recursive=False)
        if summary:
            span = summary.find('span', class_='art-count')
            if span:
                span.string = f'({count} articles)'

    # ── Update header total ────────────────────────────────────────────────────
    total = sum(len(ul.find_all('li')) for ul in soup.find_all('ul', class_='arts'))
    header = soup.find('header')
    if header:
        p = header.find('p')
        if p:
            p.string = re.sub(r'\d+ articles', f'{total} articles', p.get_text())

    # ── Write HTML ─────────────────────────────────────────────────────────────
    with open(HTML_PATH, 'w', encoding='utf-8') as fh:
        fh.write(str(soup))

    print(f'\nDone. Added {added}, skipped {skipped}.')
    print(f'Saved: {HTML_PATH}')


if __name__ == '__main__':
    main()
