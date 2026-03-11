#!/usr/bin/env python3
"""
Web Scraper for Contact Information
Scrapes phone numbers, email addresses, and social media URLs from a webpage
and its linked menu/submenu pages, then saves results as JSON.

Requirements:
    pip install playwright beautifulsoup4
    playwright install chromium
"""

import asyncio
import json
import os
import re
import sys
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ── Regex patterns ────────────────────────────────────────────────────────────
PHONE_RE = re.compile(
    r"""
    (?:(?:\+?1[\s.\-]?)?          # optional country code
    (?:\(?\d{3}\)?[\s.\-]?)       # area code
    \d{3}[\s.\-]?\d{4})           # local number
    """,
    re.VERBOSE,
)

EMAIL_RE = re.compile(
    r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b"
)

SOCIAL_DOMAINS = [
    "facebook.com", "fb.com",
    "twitter.com", "x.com",
    "instagram.com",
    "linkedin.com",
    "youtube.com",
    "tiktok.com",
    "pinterest.com",
    "snapchat.com",
    "reddit.com",
    "github.com",
    "threads.net",
    "mastodon.social",
    "bluesky.app", "bsky.app",
    "tumblr.com",
    "vimeo.com",
    "discord.gg", "discord.com",
    "telegram.me", "t.me",
    "whatsapp.com",
]

MAX_SOCIAL = 10          # top-N social URLs to keep
NAV_DEPTH  = 2           # how deep to follow nav/menu links
PAGE_TIMEOUT = 30_000    # ms


# ── Helpers ───────────────────────────────────────────────────────────────────

def same_domain(base: str, target: str) -> bool:
    """Return True if *target* shares the registered domain with *base*."""
    base_host   = urlparse(base).netloc.lower().lstrip("www.")
    target_host = urlparse(target).netloc.lower().lstrip("www.")
    return target_host == base_host or target_host.endswith("." + base_host)


def is_social_url(url: str) -> bool:
    host = urlparse(url).netloc.lower().lstrip("www.")
    return any(host == d or host.endswith("." + d) for d in SOCIAL_DOMAINS)


def normalise_phone(raw: str) -> str:
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"+1 ({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    return raw.strip()


def extract_contact(html: str) -> dict:
    """Return phones, emails, social URLs found in *html*."""
    soup  = BeautifulSoup(html, "html.parser")
    text  = soup.get_text(separator=" ")

    # phones
    raw_phones = PHONE_RE.findall(text)
    phones = list({normalise_phone(p) for p in raw_phones})

    # emails – skip image-like filenames accidentally matched
    raw_emails = EMAIL_RE.findall(text)
    emails = list({
        e.lower() for e in raw_emails
        if not re.search(r"\.(png|jpg|jpeg|gif|svg|webp)$", e, re.I)
    })

    # social URLs from <a href>
    social = []
    seen   = set()
    for tag in soup.find_all("a", href=True):
        href = tag["href"].strip()
        if href.startswith("http") and is_social_url(href):
            norm = href.rstrip("/")
            if norm not in seen:
                seen.add(norm)
                social.append(norm)

    return {"phones": phones, "emails": emails, "social_urls": social}


def extract_nav_links(html: str, base_url: str) -> list:
    """Return absolute URLs found inside nav/header/footer/menu elements."""
    soup  = BeautifulSoup(html, "html.parser")
    links = []
    seen  = set()

    # Candidate containers for navigation links
    selectors = [
        "nav", "header", "footer",
        '[role="navigation"]',
        '[class*="menu"]', '[class*="nav"]',
        '[id*="menu"]',   '[id*="nav"]',
    ]

    containers = []
    for sel in selectors:
        containers.extend(soup.select(sel))

    # Fall back to all links if nothing found
    if not containers:
        containers = [soup]

    for container in containers:
        for tag in container.find_all("a", href=True):
            href = tag["href"].strip()
            if href.startswith(("#", "javascript", "mailto", "tel")):
                continue
            abs_url = urljoin(base_url, href).split("#")[0]
            if abs_url not in seen and same_domain(base_url, abs_url):
                seen.add(abs_url)
                links.append(abs_url)

    return links


# ── Core scraper ──────────────────────────────────────────────────────────────

async def fetch_page(page, url: str) -> str | None:
    """Navigate to *url* and return its HTML, or None on error."""
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)
        await page.wait_for_timeout(1500)   # let JS settle
        return await page.content()
    except Exception as exc:
        print(f"  [warn] Could not load {url}: {exc}")
        return None


async def scrape(start_url: str) -> dict:
    """Full scrape of *start_url* and its nav-linked sub-pages."""

    results = {
        "input_url": start_url,
        "pages_scraped": [],
        "phones":    [],
        "emails":    [],
        "social_urls": [],
    }

    visited  = set()
    all_phone_set  = set()
    all_email_set  = set()
    all_social_list = []
    social_seen    = set()

    def merge(data: dict, source_url: str):
        for p in data["phones"]:
            all_phone_set.add(p)
        for e in data["emails"]:
            all_email_set.add(e)
        for s in data["social_urls"]:
            norm = s.rstrip("/")
            if norm not in social_seen:
                social_seen.add(norm)
                all_social_list.append(norm)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        )
        page = await context.new_page()

        # ── Level 0: start URL ────────────────────────────────────────────
        print(f"\n[1/3] Loading start URL: {start_url}")
        html = await fetch_page(page, start_url)
        if html is None:
            print("ERROR: Could not load the start URL. Aborting.")
            await browser.close()
            return results

        visited.add(start_url)
        results["pages_scraped"].append(start_url)
        contact0 = extract_contact(html)
        merge(contact0, start_url)

        # ── Level 1 & 2: nav links ────────────────────────────────────────
        nav_links = extract_nav_links(html, start_url)
        print(f"[2/3] Found {len(nav_links)} nav links — visiting each …")

        queue = [(url, 1) for url in nav_links]

        while queue:
            url, depth = queue.pop(0)
            if url in visited or depth > NAV_DEPTH:
                continue
            visited.add(url)

            print(f"       → {url}")
            sub_html = await fetch_page(page, url)
            if sub_html is None:
                continue

            results["pages_scraped"].append(url)
            contact_n = extract_contact(sub_html)
            merge(contact_n, url)

            if depth < NAV_DEPTH:
                deeper = extract_nav_links(sub_html, url)
                for d_url in deeper:
                    if d_url not in visited:
                        queue.append((d_url, depth + 1))

        await browser.close()

    # ── Finalise ──────────────────────────────────────────────────────────
    results["phones"]      = sorted(all_phone_set)
    results["emails"]      = sorted(all_email_set)
    results["social_urls"] = all_social_list[:MAX_SOCIAL]

    print(f"\n[3/3] Done. Scraped {len(results['pages_scraped'])} page(s).")
    return results


# ── Output ────────────────────────────────────────────────────────────────────

def build_output(results: dict) -> dict:
    """Format results with input URL first, then labelled findings."""
    return {
        "input_url": results["input_url"],
        "pages_scraped": results["pages_scraped"],
        "contact_info": {
            "type:phone_numbers":   results["phones"],
            "type:email_addresses": results["emails"],
            "type:social_media_urls": results["social_urls"],
        },
    }


def save_json(data: dict, folder: str, filename: str) -> str:
    os.makedirs(folder, exist_ok=True)
    if not filename.lower().endswith(".json"):
        filename += ".json"
    path = os.path.join(folder, filename)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    return path


# ── Entry point ───────────────────────────────────────────────────────────────

def prompt(label: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value  = input(f"{label}{suffix}: ").strip()
    return value or default


def main():
    print("=" * 60)
    print("  Web Contact-Info Scraper")
    print("=" * 60)

    start_url = prompt("Enter URL to scrape").strip()
    if not start_url.startswith("http"):
        start_url = "https://" + start_url

    folder   = prompt("Save folder path", os.path.expanduser("~/Desktop"))
    filename = prompt("Output filename", "contact_info.json")

    results = asyncio.run(scrape(start_url))
    output  = build_output(results)

    saved = save_json(output, folder, filename)

    print("\n" + "=" * 60)
    print(f"  Results saved to: {saved}")
    print("=" * 60)
    print("\nSummary:")
    ci = output["contact_info"]
    print(f"  Phones found      : {len(ci['type:phone_numbers'])}")
    print(f"  Emails found      : {len(ci['type:email_addresses'])}")
    print(f"  Social URLs found : {len(ci['type:social_media_urls'])}")
    print(f"  Pages scraped     : {len(output['pages_scraped'])}")
    print()


if __name__ == "__main__":
    main()
