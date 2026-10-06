"""
Amazon Seller Lead Crawler v1
Search keywords -> collect products + seller info -> export Excel
Perfect for B2B lead generation, supplier discovery, competitive analysis.
Edit keywords.txt to change search terms.
"""
import asyncio
import csv
import re
from datetime import datetime
from pathlib import Path

import httpx
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright

# ============ Config ============
BASE_DIR = Path(__file__).parent
KEYWORDS_FILE = BASE_DIR / "keywords.txt"
HISTORY_FILE = BASE_DIR / "历史记录.csv"
OUTPUT_DIR = BASE_DIR / "导出结果"
MAX_RESULTS = 500
HEADLESS = False
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leather bag"]


def load_history():
    seen = set()
    if HISTORY_FILE.exists():
        with open(HISTORY_FILE, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            for row in reader:
                if row:
                    seen.add(row[0])
    return seen


def save_history(results):
    new_file = not HISTORY_FILE.exists()
    with open(HISTORY_FILE, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["product_url", "title", "collect_time"])
        for r in results:
            writer.writerow([r["product_url"], r["title"], r["collect_time"]])


def clean_price(p):
    if not p:
        return ""
    m = re.search(r"[\d,.]+", str(p))
    return m.group(0) if m else ""


def clean_num(v):
    if not v:
        return 0
    s = re.sub(r"[^\d]", "", str(v))
    return int(s) if s else 0


async def search_amazon(page, keyword, seen):
    url = f"https://www.amazon.com/s?k={keyword}"
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(3)

    results = []
    products_seen = set()

    for page_num in range(1, 15):
        print(f"  Page {page_num}...", flush=True)

        items = await page.evaluate("""
            () => {
                const cards = document.querySelectorAll('div[data-component-type="s-search-result"]');
                return [...cards].map(card => {
                    const asin = card.getAttribute('data-asin') || '';
                    const title = card.querySelector('h2 span')?.innerText?.trim() || '';
                    const price = card.querySelector('.a-price .a-offscreen')?.innerText?.trim() || '';
                    const rating = card.querySelector('.a-icon-alt')?.innerText?.trim() || '';
                    const reviews = card.querySelector('span.a-size-base.s-underline-text')?.innerText?.trim() || '';
                    const link = 'https://www.amazon.com/dp/' + asin;
                    const badge = card.querySelector('.a-badge-text')?.innerText?.trim() || '';
                    const sponsored = !!card.querySelector('.a-color-secondary:has-text("Sponsored")');
                    return {asin, title, price, rating, reviews, link, badge, sponsored};
                });
            }
        """)

        for item in items:
            asin = item.get("asin", "")
            if not asin or asin in products_seen:
                continue
            if item["link"] in seen:
                continue
            products_seen.add(asin)
            results.append(item)
            if len(results) >= MAX_RESULTS:
                break

        if len(results) >= MAX_RESULTS:
            break

        # Next page
        try:
            next_btn = await page.query_selector('a.s-pagination-next')
            if not next_btn:
                break
            await next_btn.click()
            await asyncio.sleep(3)
        except:
            break

    return results


async def get_seller_info(page, product_url):
    """Visit product page to get seller/brand info."""
    try:
        await page.goto(product_url, wait_until="domcontentloaded", timeout=20000)
        await asyncio.sleep(2)

        info = await page.evaluate("""
            () => {
                const brand = document.querySelector('#bylineInfo')?.innerText?.trim() || '';
                const seller = document.querySelector('#sellerProfileTriggerId')?.innerText?.trim() || '';
                const soldBy = document.querySelector('#tabular-buybox .tabular-buybox-text')?.innerText?.trim() || '';
                const bullets = [...document.querySelectorAll('#feature-bullets li span')].map(x => x.innerText.trim()).slice(0,3);
                const description = document.querySelector('#productDescription')?.innerText?.trim()?.slice(0, 300) || '';
                return {brand, seller, soldBy, bullets: bullets.join(' | '), description};
            }
        """)
        return info
    except:
        return {"brand": "", "seller": "", "soldBy": "", "bullets": "", "description": ""}


def extract_email_from_text(text):
    if not text:
        return ""
    emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text)
    emails = [e for e in emails if not any(x in e.lower() for x in ['example.', 'sentry', '.png', '.jpg'])]
    return emails[0] if emails else ""


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Amazon Leads"

    headers = ["Score", "Keyword", "Product Title", "Price", "Rating", "Reviews",
                "Brand", "Seller", "Product URL", "Bullet Points", "Description",
                "Email", "Collected At"]
    ws.append(headers)

    for r in all_data:
        ws.append([
            r["score"], r["keyword"], r["title"], r["price"], r["rating"], r["reviews"],
            r["brand"], r["seller"], r["product_url"], r["bullets"], r["description"],
            r["email"], r["collect_time"]
        ])

    # Style
    fill = PatternFill("solid", fgColor="FF9900")
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
    ws.freeze_panes = "A2"

    widths = [8, 15, 50, 10, 8, 10, 20, 20, 45, 50, 50, 30, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    fname = OUTPUT_DIR / f"Amazon_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}", flush=True)


async def main():
    keywords = load_keywords()
    print("=" * 50, flush=True)
    print("  Amazon Seller Lead Crawler v1", flush=True)
    print("=" * 50, flush=True)
    print(f"Keywords: {', '.join(keywords)}", flush=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, channel="chrome")
        ctx = await browser.new_context(
            locale="en-US",
            viewport={"width": 1280, "height": 800},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page = await ctx.new_page()

        seen = load_history()
        if seen:
            print(f"Loaded history: {len(seen)} already-collected", flush=True)

        total = 0
        for kw in keywords:
            print(f"\n=== {kw} ===", flush=True)
            try:
                products = await search_amazon(page, kw, seen)
                print(f"  Found {len(products)} products", flush=True)

                results = []
                for i, prod in enumerate(products, 1):
                    seller_info = await get_seller_info(page, prod["link"])
                    email = extract_email_from_text(seller_info.get("description", ""))

                    reviews = clean_num(prod.get("reviews"))
                    rating = clean_num(prod.get("rating"))

                    # Score: more reviews + higher rating = better lead
                    score = min(100, reviews // 10 + int(rating) * 5)

                    results.append({
                        "keyword": kw,
                        "title": prod.get("title", ""),
                        "price": clean_price(prod.get("price", "")),
                        "rating": prod.get("rating", ""),
                        "reviews": prod.get("reviews", ""),
                        "brand": seller_info.get("brand", ""),
                        "seller": seller_info.get("seller", ""),
                        "product_url": prod["link"],
                        "bullets": seller_info.get("bullets", ""),
                        "description": seller_info.get("description", ""),
                        "email": email,
                        "score": score,
                        "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    })

                    seen.add(prod["link"])
                    if i % 10 == 0:
                        print(f"    [{i}/{len(products)}] {prod['title'][:40]}", flush=True)

                    await asyncio.sleep(0.5)

                results.sort(key=lambda x: x["score"], reverse=True)
                if results:
                    save_to_excel(results, kw)
                    save_history(results)
                total += len(results)
            except Exception as e:
                print(f"  Error: {e}", flush=True)

        print(f"\nDone! Total {total} products.", flush=True)
        print(f"Output: {OUTPUT_DIR.resolve()}", flush=True)
        await browser.close()
        input("Press Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())
