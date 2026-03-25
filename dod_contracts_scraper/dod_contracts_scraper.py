"""
DOD Contracts News Scraper
Fetches DOD contract awards since 1/1/2026 from:
  1. USASpending.gov API  - official federal spending data
  2. Defense.gov news RSS - official DoD press releases
"""

import json
import time
import xml.etree.ElementTree as ET
from datetime import date, datetime
from typing import Optional

import requests

# ── Constants ────────────────────────────────────────────────────────────────

START_DATE = "2026-01-01"
USASPENDING_URL = "https://api.usaspending.gov/api/v2/search/spending_by_award/"
DEFENSE_RSS_URL = "https://www.defense.gov/DesktopModules/ArticleCS/RSS.ashx?ContentType=1&Site=945&max=100"

# DOD agency codes on USASpending.gov (toptier)
DOD_AGENCY_CODE = "097"

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "dod-contracts-scraper/1.0"})


# ── USASpending.gov ───────────────────────────────────────────────────────────

def fetch_usaspending_contracts(
    start_date: str = START_DATE,
    end_date: Optional[str] = None,
    page_size: int = 100,
    max_pages: int = 10,
) -> list[dict]:
    """
    Query USASpending.gov for DOD contract awards since start_date.

    Returns a list of award dicts with fields:
        award_id, recipient_name, award_amount, description,
        awarding_agency, period_of_performance_start,
        period_of_performance_end, place_of_performance
    """
    if end_date is None:
        end_date = date.today().isoformat()

    payload = {
        "filters": {
            "agencies": [
                {
                    "type": "awarding",
                    "tier": "toptier",
                    "toptier_code": DOD_AGENCY_CODE,
                }
            ],
            "award_type_codes": ["A", "B", "C", "D"],  # contract types
            "time_period": [{"start_date": start_date, "end_date": end_date}],
        },
        "fields": [
            "Award ID",
            "Recipient Name",
            "Award Amount",
            "Description",
            "Awarding Agency",
            "Awarding Sub Agency",
            "Period of Performance Start Date",
            "Period of Performance Current End Date",
            "Place of Performance State Code",
            "Place of Performance Country Code",
            "Award Type",
            "Contract Award Type",
        ],
        "sort": "Award Amount",
        "order": "desc",
        "limit": page_size,
        "page": 1,
    }

    all_results: list[dict] = []

    for page in range(1, max_pages + 1):
        payload["page"] = page
        try:
            resp = SESSION.post(USASPENDING_URL, json=payload, timeout=30)
            resp.raise_for_status()
        except requests.RequestException as exc:
            print(f"[USASpending] Request error on page {page}: {exc}")
            break

        data = resp.json()
        results = data.get("results", [])
        if not results:
            break

        all_results.extend(results)
        print(f"[USASpending] Fetched page {page} — {len(results)} records "
              f"(total so far: {len(all_results)})")

        # Respect API rate limits
        time.sleep(0.5)

        if len(results) < page_size:
            break  # Last page

    return all_results


def normalize_usaspending(raw: dict) -> dict:
    """Flatten USASpending result into a consistent schema."""
    return {
        "source": "USASpending.gov",
        "award_id": raw.get("Award ID"),
        "recipient": raw.get("Recipient Name"),
        "amount_usd": raw.get("Award Amount"),
        "description": raw.get("Description"),
        "awarding_agency": raw.get("Awarding Agency"),
        "awarding_sub_agency": raw.get("Awarding Sub Agency"),
        "award_type": raw.get("Contract Award Type") or raw.get("Award Type"),
        "start_date": raw.get("Period of Performance Start Date"),
        "end_date": raw.get("Period of Performance Current End Date"),
        "place_of_performance": (
            f"{raw.get('Place of Performance State Code', '')} "
            f"{raw.get('Place of Performance Country Code', '')}".strip()
        ),
        "url": (
            f"https://www.usaspending.gov/award/{raw['Award ID']}"
            if raw.get("Award ID") else None
        ),
    }


# ── Defense.gov RSS ───────────────────────────────────────────────────────────

def fetch_defense_rss(
    start_date: str = START_DATE,
    url: str = DEFENSE_RSS_URL,
) -> list[dict]:
    """
    Fetch Defense.gov RSS feed and filter for contract-related items
    published on or after start_date.
    """
    since = datetime.fromisoformat(start_date).date()

    try:
        resp = SESSION.get(url, timeout=30)
        resp.raise_for_status()
    except requests.RequestException as exc:
        print(f"[Defense.gov RSS] Request error: {exc}")
        return []

    root = ET.fromstring(resp.content)
    channel = root.find("channel")
    if channel is None:
        print("[Defense.gov RSS] Unexpected XML structure — no <channel> found.")
        return []

    contract_keywords = {
        "contract", "award", "awarded", "procurement", "acquisition",
        "firm-fixed", "indefinite-delivery", "idiq", "sole source",
    }

    items: list[dict] = []
    for item in channel.findall("item"):
        pub_date_str = (item.findtext("pubDate") or "").strip()
        title = (item.findtext("title") or "").strip()
        description = (item.findtext("description") or "").strip()
        link = (item.findtext("link") or "").strip()

        # Parse publication date
        pub_date: Optional[date] = None
        for fmt in ("%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S %Z"):
            try:
                pub_date = datetime.strptime(pub_date_str, fmt).date()
                break
            except ValueError:
                continue

        if pub_date is None or pub_date < since:
            continue

        combined_text = (title + " " + description).lower()
        if not any(kw in combined_text for kw in contract_keywords):
            continue

        items.append({
            "source": "Defense.gov RSS",
            "title": title,
            "description": description,
            "pub_date": pub_date.isoformat(),
            "url": link,
        })

    print(f"[Defense.gov RSS] Found {len(items)} contract-related items since {start_date}.")
    return items


# ── Output helpers ────────────────────────────────────────────────────────────

def print_usaspending_summary(contracts: list[dict], top_n: int = 20) -> None:
    print(f"\n{'='*70}")
    print(f"  USASpending.gov — Top {top_n} DOD Contract Awards since {START_DATE}")
    print(f"{'='*70}")
    for i, c in enumerate(contracts[:top_n], 1):
        amount = c["amount_usd"]
        amount_str = f"${amount:,.0f}" if isinstance(amount, (int, float)) else str(amount)
        print(f"\n[{i}] {c['recipient']}")
        print(f"    Award ID  : {c['award_id']}")
        print(f"    Amount    : {amount_str}")
        print(f"    Agency    : {c['awarding_sub_agency'] or c['awarding_agency']}")
        print(f"    Type      : {c['award_type']}")
        print(f"    Period    : {c['start_date']} → {c['end_date']}")
        print(f"    Location  : {c['place_of_performance']}")
        if c["description"]:
            print(f"    Desc      : {c['description'][:120]}{'...' if len(c['description'] or '') > 120 else ''}")
        print(f"    URL       : {c['url']}")


def print_rss_summary(items: list[dict]) -> None:
    print(f"\n{'='*70}")
    print(f"  Defense.gov — Contract News since {START_DATE}")
    print(f"{'='*70}")
    for i, item in enumerate(items, 1):
        print(f"\n[{i}] {item['pub_date']}  {item['title']}")
        if item["description"]:
            snippet = item["description"][:200].replace("\n", " ")
            print(f"     {snippet}{'...' if len(item['description']) > 200 else ''}")
        print(f"     {item['url']}")


def save_json(data: object, path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, default=str)
    print(f"\nSaved → {path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    print(f"DOD Contracts Scraper  |  data since {START_DATE}  |  run date {date.today()}\n")

    # ── 1. USASpending.gov ──
    print("Querying USASpending.gov API …")
    raw_awards = fetch_usaspending_contracts(start_date=START_DATE, max_pages=5)
    contracts = [normalize_usaspending(r) for r in raw_awards]
    print_usaspending_summary(contracts)
    save_json(contracts, "dod_contracts_usaspending.json")

    # ── 2. Defense.gov RSS ──
    print("\nFetching Defense.gov RSS …")
    rss_items = fetch_defense_rss(start_date=START_DATE)
    print_rss_summary(rss_items)
    save_json(rss_items, "dod_contracts_defense_rss.json")

    print(f"\nDone.  {len(contracts)} contract awards + {len(rss_items)} news items collected.")


if __name__ == "__main__":
    main()
