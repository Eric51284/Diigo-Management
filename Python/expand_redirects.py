import csv
import requests
import time
from urllib.parse import urlparse

INPUT_FILE = "Source files/raindrop_export_2026_02_14/export.csv"
OUTPUT_FILE = "Source files/raindrop_export_2026_02_14/export_expanded.csv"

# Set to True if you ONLY want to expand flip.it links
ONLY_EXPAND_FLIP = True

# Timeout in seconds for each request
REQUEST_TIMEOUT = 10

def is_flip_url(url):
    try:
        parsed = urlparse(url)
        return "flip.it" in parsed.netloc
    except:
        return False

def resolve_url(url):
    try:
        response = requests.get(
            url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            headers={"User-Agent": "Mozilla/5.0"}
        )
        return response.url
    except Exception as e:
        print(f"Error resolving {url}: {e}")
        return url  # return original if failure

def main():
    with open(INPUT_FILE, newline='', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)
        fieldnames = reader.fieldnames

        with open(OUTPUT_FILE, "w", newline='', encoding='utf-8') as outfile:
            writer = csv.DictWriter(outfile, fieldnames=fieldnames)
            writer.writeheader()

            for i, row in enumerate(reader, start=1):
                original_url = row.get("url") or row.get("URL") or row.get("Url")

                if not original_url:
                    writer.writerow(row)
                    continue

                if ONLY_EXPAND_FLIP and not is_flip_url(original_url):
                    writer.writerow(row)
                    continue

                print(f"[{i}] Resolving: {original_url}")

                final_url = resolve_url(original_url)

                if final_url != original_url:
                    print(f"    → {final_url}")
                    row["url"] = final_url

                writer.writerow(row)

                # Be polite — avoid hammering servers
                time.sleep(0.5)

    print("\nDone. Output written to:", OUTPUT_FILE)

if __name__ == "__main__":
    main()