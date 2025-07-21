#!/usr/bin/env python3
import sys
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError
from consts import SEARCH_URL, BASE_URL, INPUT_FILE, OUTPUT_FILE, IMAGE_COL, RESULT_COL

def fetch_product_link(img_url: str) -> str:
    """
    Goes to searchbyimage.com, submits the image URL, and returns the href
    of the <a> inside <div id="product_img"> (or None if not found).
    """
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        page = browser.new_page()

        # 1) Navigate to the search-by-image page
        page.goto(SEARCH_URL)

        # 2) Enter the image URL and submit
        page.fill("input#search_query", img_url)
        page.click("#searchbyimage_submit")

        # 3) Wait for the product_img container to appear
        try:
            page.wait_for_selector("div#product_img a", timeout=30000)
            href = page.query_selector("div#product_img a").get_attribute("href")
        except TimeoutError:
            href = None

        browser.close()
        return href

def main():
    df = pd.read_excel(INPUT_FILE)
    if IMAGE_COL not in df.columns:
        print(f"Error: Excel must have a column named '{IMAGE_COL}'", file=sys.stderr)
        sys.exit(1)

    results = []
    for img_url in df[IMAGE_COL].dropna():
        link = fetch_product_link(img_url)
        a = [BASE_URL, link]
        final_url = ''.join(a)
        print(f"{img_url} → {final_url}")
        results.append(final_url)

    df[RESULT_COL] = results
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Done: wrote {len(results)} rows to '{OUTPUT_FILE}'")

if __name__ == "__main__":
    main()
