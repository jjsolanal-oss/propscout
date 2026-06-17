#!/usr/bin/env python3
"""
PropScout V2 Phase 1 — Direct scraper + Claude batch classifier

Architecture:
  1. requests + BeautifulSoup  → scrape Crexi & LoopNet directly
  2. ONE Claude API call        → batch-classify + flag all listings
  3. openpyxl                  → export propscout_sa_phase1.xlsx
  4. JSON                      → save raw listings to propscout_sa_raw.json
"""

import anthropic
import json
import os
import re
import sys
import time
from dataclasses import dataclass, asdict
from datetime import datetime
from difflib import SequenceMatcher
from typing import Optional

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Constants ─────────────────────────────────────────────────────────────────

MODEL        = "claude-sonnet-4-6"
RAW_JSON     = "propscout_sa_raw.json"
EXCEL_PATH   = "propscout_sa_phase1.xlsx"
PAGE_DELAY   = 2.0   # seconds between page requests
MAX_PAGES    = 5     # pages to try per target URL
REQ_TIMEOUT  = 20    # seconds per HTTP request

SCRAPE_TARGETS = [
    {
        "platform":  "Crexi",
        "prop_type": "flex industrial",
        "url": "https://www.crexi.com/properties?types=Industrial&state=TX&city=San+Antonio",
    },
    {
        "platform":  "Crexi",
        "prop_type": "land",
        "url": "https://www.crexi.com/properties?types=Land&state=TX&city=San+Antonio",
    },
    {
        "platform":  "LoopNet",
        "prop_type": "mixed",   # inferred per-listing from name/details
        "url": "https://www.loopnet.com/search/commercial-real-estate/san-antonio-tx/for-sale/",
    },
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,image/apng,*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection":      "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest":  "document",
    "Sec-Fetch-Mode":  "navigate",
    "Sec-Fetch-Site":  "none",
    "Cache-Control":   "max-age=0",
}

SQF_PER_ACRE      = 43_560
VA_PSF_THRESHOLD  = 120.0
DEV_PSF_THRESHOLD = 5.0

CURRENT_YEAR = datetime.now().year


# ── Data model ────────────────────────────────────────────────────────────────

@dataclass
class Listing:
    listing_name:     str            = ""
    address:          str            = ""
    city:             str            = "San Antonio"
    state:            str            = "TX"
    zip_code:         str            = ""
    price:            Optional[float] = None
    total_sqft:       Optional[float] = None
    price_per_sqft:   Optional[float] = None
    year_built:       Optional[int]   = None
    lot_size_acres:   Optional[float] = None
    zoning:           Optional[str]   = None
    days_on_market:   Optional[int]   = None
    property_type:    str            = ""
    listing_source:   str            = ""
    listing_url:      str            = ""
    deal_type:        str            = ""
    deal_type_reason: str            = ""
    flag:             str            = ""


# ── Number parsing helpers ────────────────────────────────────────────────────

def _parse_price(text: str) -> Optional[float]:
    if not text:
        return None
    text = str(text).strip()
    m = re.search(r"([\d,]+\.?\d*)\s*([KkMmBb])?", text.replace("$", "").replace(",", ""))
    if not m:
        return None
    val = float(m.group(1))
    suffix = (m.group(2) or "").upper()
    if suffix == "K":
        val *= 1_000
    elif suffix == "M":
        val *= 1_000_000
    elif suffix == "B":
        val *= 1_000_000_000
    return val if val > 0 else None


def _parse_sqft(text: str) -> Optional[float]:
    if not text:
        return None
    text = re.sub(r"[,\s]*(sq\.?\s*ft\.?|sqft|sf)\b", "", str(text), flags=re.IGNORECASE)
    text = text.replace(",", "")
    m = re.search(r"[\d]+\.?\d*", text)
    if m:
        try:
            return float(m.group())
        except ValueError:
            pass
    return None


def _parse_int(text) -> Optional[int]:
    if text is None:
        return None
    m = re.search(r"\d+", str(text).replace(",", ""))
    return int(m.group()) if m else None


def _parse_acres(text) -> Optional[float]:
    if text is None:
        return None
    m = re.search(r"[\d]+\.?\d*", str(text).replace(",", ""))
    if m:
        val = float(m.group())
        return val if val > 0 else None
    return None


# ── Crexi scraping ────────────────────────────────────────────────────────────

def _deep_find_list(obj, depth=0) -> Optional[list]:
    """Recursively find first list of dicts that looks like property listings."""
    if depth > 8:
        return None
    if isinstance(obj, list) and len(obj) >= 1 and isinstance(obj[0], dict):
        keys = set(obj[0].keys())
        if keys & {"address", "price", "askingPrice", "listPrice", "streetAddress",
                   "sqft", "buildingSize", "slug", "propertyId"}:
            return obj
    if isinstance(obj, dict):
        for v in obj.values():
            r = _deep_find_list(v, depth + 1)
            if r:
                return r
    return None


def _listings_from_crexi_json(data: dict, prop_type: str) -> list[Listing]:
    """Extract Listing objects from Crexi's Next.js __NEXT_DATA__ blob."""
    page_props = data.get("props", {}).get("pageProps", {})

    raw_list = None
    for path in [
        lambda d: d.get("properties"),
        lambda d: d.get("listings"),
        lambda d: d.get("data", {}).get("properties"),
        lambda d: d.get("initialState", {}).get("properties", {}).get("list"),
        lambda d: d.get("searchResults", {}).get("properties"),
        lambda d: _deep_find_list(d),
    ]:
        raw_list = path(page_props)
        if isinstance(raw_list, list) and raw_list:
            break

    if not raw_list:
        raw_list = _deep_find_list(data)

    if not raw_list:
        return []

    out = []
    for item in raw_list:
        if not isinstance(item, dict):
            continue

        addr = str(item.get("address") or item.get("streetAddress") or
                   item.get("street") or "").strip()
        city = str(item.get("city") or item.get("cityName") or "San Antonio").strip()
        state = str(item.get("state") or item.get("stateCode") or "TX").strip()
        zip_code = str(item.get("zip") or item.get("zipCode") or "").strip()

        price_raw = (item.get("price") or item.get("askingPrice") or
                     item.get("listPrice") or item.get("salePrice"))
        price = _parse_price(str(price_raw)) if price_raw is not None else None

        sqft_raw = (item.get("sqft") or item.get("buildingSize") or
                    item.get("totalSize") or item.get("size"))
        sqft = _parse_sqft(str(sqft_raw)) if sqft_raw is not None else None

        year = _parse_int(item.get("yearBuilt") or item.get("year_built"))
        acres = _parse_acres(item.get("lotSize") or item.get("acreage") or
                             item.get("lot_size"))
        dom = _parse_int(item.get("daysOnMarket") or item.get("days_on_market"))

        slug = item.get("slug") or item.get("id") or item.get("propertyId") or ""
        url = str(item.get("url") or item.get("listingUrl") or
                  (f"https://www.crexi.com/properties/{slug}" if slug else "")).strip()

        name = str(item.get("name") or item.get("title") or item.get("propertyName") or
                   item.get("listingTitle") or addr or "Crexi Listing").strip()

        psf = round(price / sqft, 2) if (price and sqft and sqft > 0) else None

        if not addr and not name:
            continue

        out.append(Listing(
            listing_name   = name[:120],
            address        = addr[:120],
            city           = city,
            state          = state,
            zip_code       = zip_code[:10],
            price          = price,
            total_sqft     = sqft,
            price_per_sqft = psf,
            year_built     = year,
            lot_size_acres = acres,
            days_on_market = dom,
            property_type  = prop_type,
            listing_source = "Crexi",
            listing_url    = url[:300],
        ))

    return out


def _listings_from_crexi_html(soup: BeautifulSoup, prop_type: str) -> list[Listing]:
    """Fallback: parse Crexi listing cards from rendered HTML."""
    selectors = [
        ".property-card", "[data-testid='property-card']", ".listing-card",
        "article.property", ".search-result-item", ".property-row",
    ]
    cards = []
    for sel in selectors:
        cards = soup.select(sel)
        if cards:
            break

    out = []
    for card in cards:
        try:
            name_el = (card.select_one(".property-name, .listing-name, h2, h3"))
            name = name_el.get_text(strip=True) if name_el else ""

            addr_el = card.select_one(".property-address, .address, [data-testid='address']")
            addr_raw = addr_el.get_text(strip=True) if addr_el else ""

            price_el = card.select_one(".property-price, .price, [data-testid='price']")
            price_text = price_el.get_text(strip=True) if price_el else ""

            size_el = card.select_one(".property-size, .sqft, .size")
            size_text = size_el.get_text(strip=True) if size_el else ""

            link = card.select_one("a[href]")
            href = link["href"] if link else ""
            if href and not href.startswith("http"):
                href = f"https://www.crexi.com{href}"

            price = _parse_price(price_text)
            sqft  = _parse_sqft(size_text)
            psf   = round(price / sqft, 2) if (price and sqft and sqft > 0) else None

            # Parse city / zip from address text
            city = "San Antonio"
            zip_code = ""
            clean_addr = addr_raw
            parts = addr_raw.split(",")
            if len(parts) >= 2:
                clean_addr = parts[0].strip()
                rest = ",".join(parts[1:])
                cm = re.match(r"\s*(.+?)\s+(TX|Texas)", rest, re.IGNORECASE)
                if cm:
                    city = cm.group(1).strip()
            zm = re.search(r"\b(\d{5})\b", addr_raw)
            if zm:
                zip_code = zm.group(1)

            if not name and not clean_addr:
                continue

            out.append(Listing(
                listing_name   = name,
                address        = clean_addr,
                city           = city,
                state          = "TX",
                zip_code       = zip_code,
                price          = price,
                total_sqft     = sqft,
                price_per_sqft = psf,
                property_type  = prop_type,
                listing_source = "Crexi",
                listing_url    = href,
            ))
        except Exception:
            continue

    return out


def scrape_crexi(session: requests.Session, target: dict) -> list[Listing]:
    url       = target["url"]
    prop_type = target["prop_type"]
    listings  = []

    for page in range(1, MAX_PAGES + 1):
        page_url = url if page == 1 else f"{url}&page={page}"
        print(f"    page {page}: {page_url[:90]}")
        sys.stdout.flush()

        try:
            resp = session.get(page_url, timeout=REQ_TIMEOUT)
        except Exception as e:
            print(f"    ✗ request error: {e}")
            break

        print(f"    HTTP {resp.status_code}  ({len(resp.text):,} bytes)")

        if resp.status_code in (403, 429, 503):
            print(f"    ✗ blocked (HTTP {resp.status_code})")
            break
        if resp.status_code != 200:
            print(f"    ✗ unexpected status {resp.status_code}")
            break

        # Cloudflare challenge detection
        if ("cf-ray" in resp.headers or
                "checking your browser" in resp.text.lower() or
                "enable javascript" in resp.text.lower()):
            print("    ✗ Cloudflare/JS challenge — scraping blocked")
            break

        soup = BeautifulSoup(resp.text, "html.parser")

        # Strategy 1: Next.js __NEXT_DATA__
        next_tag = soup.find("script", id="__NEXT_DATA__")
        if next_tag and next_tag.string:
            try:
                next_data = json.loads(next_tag.string)
                found = _listings_from_crexi_json(next_data, prop_type)
                if found:
                    print(f"    ✓ Next.js JSON → {len(found)} listings")
                    listings.extend(found)
                    if len(found) < 5:
                        break   # last page
                    time.sleep(PAGE_DELAY)
                    continue
            except Exception as e:
                print(f"    ⚠ Next.js parse failed: {e}")

        # Strategy 2: any script tag containing JSON with address keys
        for script in soup.find_all("script"):
            src = script.string or ""
            if len(src) > 300 and "address" in src and "price" in src:
                for m in re.finditer(r"(\[\s*\{.*?\}\s*\])", src, re.DOTALL):
                    try:
                        candidate = json.loads(m.group(1))
                        if isinstance(candidate, list) and candidate:
                            found = _listings_from_crexi_json(
                                {"props": {"pageProps": {"properties": candidate}}},
                                prop_type
                            )
                            if found:
                                print(f"    ✓ script JSON → {len(found)} listings")
                                listings.extend(found)
                    except Exception:
                        pass

        # Strategy 3: HTML card parsing
        html_found = _listings_from_crexi_html(soup, prop_type)
        if html_found:
            print(f"    ✓ HTML cards → {len(html_found)} listings")
            listings.extend(html_found)

        if not listings and page == 1:
            print("    ⚠ no listings found on page 1 — stopping")
            break

        time.sleep(PAGE_DELAY)

    return listings


# ── LoopNet scraping ──────────────────────────────────────────────────────────

def _infer_prop_type(text: str) -> str:
    t = text.lower()
    if any(w in t for w in ("land", " lot", "acre", "tract", "vacant")):
        return "land"
    if any(w in t for w in ("flex", "industrial", "warehouse", "distribution", "manufacturing")):
        return "flex industrial"
    return "commercial"


def _listings_from_loopnet_html(soup: BeautifulSoup) -> list[Listing]:
    selectors = [
        ".placard", ".property-row", "[data-testid='property-card']",
        ".listing-row", ".srp-item", "article.property",
    ]
    cards = []
    for sel in selectors:
        cards = soup.select(sel)
        if cards:
            break

    out = []
    for card in cards:
        try:
            name_el = (card.select_one(".placard-content-title") or
                       card.select_one("h3, h4, [class*='title']"))
            name = name_el.get_text(strip=True) if name_el else ""

            addr_el = (card.select_one(".placard-content-address") or
                       card.select_one(".property-address, .address, [itemprop='streetAddress']"))
            addr_raw = addr_el.get_text(strip=True) if addr_el else ""

            price_el = (card.select_one("[class*='price'], .dollar-amount"))
            price_text = price_el.get_text(strip=True) if price_el else ""

            size_el = card.select_one("[class*='sqft'], [class*='size'], [class*='sf']")
            size_text = size_el.get_text(strip=True) if size_el else ""

            link = card.select_one("a[href]")
            href = link["href"] if link else ""
            if href and not href.startswith("http"):
                href = f"https://www.loopnet.com{href}"

            price = _parse_price(price_text)
            sqft  = _parse_sqft(size_text)
            psf   = round(price / sqft, 2) if (price and sqft and sqft > 0) else None

            city = "San Antonio"
            zip_code = ""
            clean_addr = addr_raw
            parts = addr_raw.split(",")
            if len(parts) >= 2:
                clean_addr = parts[0].strip()
                rest = ",".join(parts[1:])
                cm = re.match(r"\s*(.+?)\s+(TX|Texas)", rest, re.IGNORECASE)
                if cm:
                    city = cm.group(1).strip()
            zm = re.search(r"\b(\d{5})\b", addr_raw)
            if zm:
                zip_code = zm.group(1)

            if not name and not clean_addr:
                continue

            prop_type = _infer_prop_type(name + " " + clean_addr)

            out.append(Listing(
                listing_name   = name,
                address        = clean_addr,
                city           = city,
                state          = "TX",
                zip_code       = zip_code,
                price          = price,
                total_sqft     = sqft,
                price_per_sqft = psf,
                property_type  = prop_type,
                listing_source = "LoopNet",
                listing_url    = href,
            ))
        except Exception:
            continue

    return out


def _listings_from_json_ld(soup: BeautifulSoup, platform: str) -> list[Listing]:
    """Extract any JSON-LD structured data (ItemList, RealEstateListing, etc.)."""
    out = []
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            items = []
            if isinstance(data, dict):
                if data.get("@type") == "ItemList":
                    items = data.get("itemListElement", [])
                elif data.get("@type") in ("RealEstateListing", "Product"):
                    items = [data]
            elif isinstance(data, list):
                items = data

            for item in items:
                if not isinstance(item, dict):
                    continue
                thing = item.get("item", item)
                name = str(thing.get("name", "")).strip()
                addr_obj = thing.get("address", {})
                if isinstance(addr_obj, dict):
                    addr = str(addr_obj.get("streetAddress", "")).strip()
                    city = str(addr_obj.get("addressLocality", "San Antonio")).strip()
                    zip_code = str(addr_obj.get("postalCode", "")).strip()
                else:
                    addr = str(addr_obj).strip()
                    city = "San Antonio"
                    zip_code = ""
                price_raw = thing.get("price") or (thing.get("offers") or {}).get("price")
                price = _parse_price(str(price_raw)) if price_raw else None
                url = str(thing.get("url", "")).strip()

                if not name and not addr:
                    continue

                out.append(Listing(
                    listing_name   = name[:120],
                    address        = addr[:120],
                    city           = city,
                    state          = "TX",
                    zip_code       = zip_code[:10],
                    price          = price,
                    property_type  = _infer_prop_type(name + " " + addr),
                    listing_source = platform,
                    listing_url    = url[:300],
                ))
        except Exception:
            continue

    return out


def scrape_loopnet(session: requests.Session, target: dict) -> list[Listing]:
    url      = target["url"].rstrip("/") + "/"
    listings = []

    for page in range(1, MAX_PAGES + 1):
        page_url = url if page == 1 else f"{url}{page}/"
        print(f"    page {page}: {page_url}")
        sys.stdout.flush()

        try:
            resp = session.get(page_url, timeout=REQ_TIMEOUT)
        except Exception as e:
            print(f"    ✗ request error: {e}")
            break

        print(f"    HTTP {resp.status_code}  ({len(resp.text):,} bytes)")

        if resp.status_code in (403, 429, 503):
            print(f"    ✗ blocked (HTTP {resp.status_code})")
            break
        if resp.status_code != 200:
            break

        if ("cf-ray" in resp.headers or
                "checking your browser" in resp.text.lower()):
            print("    ✗ Cloudflare challenge — scraping blocked")
            break

        soup = BeautifulSoup(resp.text, "html.parser")

        jld = _listings_from_json_ld(soup, "LoopNet")
        if jld:
            print(f"    ✓ JSON-LD → {len(jld)} listings")
            listings.extend(jld)

        html_found = _listings_from_loopnet_html(soup)
        if html_found:
            print(f"    ✓ HTML cards → {len(html_found)} listings")
            listings.extend(html_found)

        if not jld and not html_found:
            print(f"    ⚠ no listings on page {page} — stopping")
            break

        time.sleep(PAGE_DELAY)

    return listings


# ── PSF calculation + deduplication ──────────────────────────────────────────

def calculate_psf(lst: Listing) -> Listing:
    if lst.price is None:
        return lst
    sqft = lst.total_sqft
    if sqft is None and lst.lot_size_acres:
        sqft = lst.lot_size_acres * SQF_PER_ACRE
    if sqft and sqft > 0:
        lst.price_per_sqft = round(lst.price / sqft, 2)
    return lst


def deduplicate(listings: list[Listing]) -> tuple[list[Listing], int]:
    def norm(s: str) -> str:
        s = s.lower().strip()
        s = re.sub(r"[^\w\s]", " ", s)
        return re.sub(r"\s+", " ", s)

    def similar(a: str, b: str) -> float:
        return SequenceMatcher(None, norm(a), norm(b)).ratio()

    def prices_close(p1, p2) -> bool:
        if not p1 or not p2:
            return True
        return abs(p1 - p2) / max(p1, p2) <= 0.05

    kept: list[Listing] = []
    removed = 0

    for lst in listings:
        dup = None
        for k in kept:
            if lst.address and k.address and similar(lst.address, k.address) >= 0.80:
                if prices_close(lst.price, k.price):
                    dup = k
                    break
        if dup:
            removed += 1
            for f in vars(lst):
                if f in ("listing_source", "deal_type", "deal_type_reason", "flag"):
                    continue
                if getattr(dup, f) in (None, "", 0):
                    setattr(dup, f, getattr(lst, f))
        else:
            kept.append(lst)

    return kept, removed


# ── Claude batch classification (ONE API call) ────────────────────────────────

RULES = f"""
CLASSIFICATION RULES (current year = {CURRENT_YEAR}, apply first match):
1. DEVELOPMENT  — property_type is "land"
2. DEVELOPMENT  — year_built is null AND total_sqft is null
3. UNKNOWN      — year_built is null AND total_sqft is not null  (building exists, age unknown)
4. UNKNOWN      — year_built present AND ({CURRENT_YEAR} - year_built) < 10  (too new)
5. VALUE-ADD    — year_built present AND age ≥ 10 AND
                  (lot_size_acres is null OR total_sqft/(lot_size_acres*43560) > 0.50)
6. DEVELOPMENT  — year_built present AND age ≥ 10 AND lot_size_acres present AND
                  total_sqft/(lot_size_acres*43560) ≤ 0.50  (≥50% undeveloped site)
7. UNKNOWN      — anything else with insufficient data

FLAGGING RULES (compute category averages across the full batch first):
  thresholds : VALUE-ADD = ${VA_PSF_THRESHOLD:.0f}/sqft,  non-VALUE-ADD = ${DEV_PSF_THRESHOLD:.0f}/sqft
  STRONG BUY : price_per_sqft ≤ threshold  AND  ≥ 10% below category average
  BELOW AVG  : price_per_sqft > threshold  AND  ≥ 10% below category average
  WATCH      : price_per_sqft ≤ threshold  AND  < 10% below category average
  (blank)    : none of the above, or price_per_sqft is null
"""


def classify_with_claude(client: anthropic.Anthropic,
                         listings: list[Listing]) -> list[Listing]:
    if not listings:
        return listings

    print(f"  Sending {len(listings)} listings to Claude for batch classification …")
    sys.stdout.flush()

    batch = [
        {
            "index":          i,
            "listing_name":   l.listing_name,
            "address":        l.address,
            "price":          l.price,
            "total_sqft":     l.total_sqft,
            "price_per_sqft": l.price_per_sqft,
            "year_built":     l.year_built,
            "lot_size_acres": l.lot_size_acres,
            "property_type":  l.property_type,
        }
        for i, l in enumerate(listings)
    ]

    prompt = f"""You are a commercial real estate analyst.

{RULES}

LISTINGS:
{json.dumps(batch, indent=2)}

Return ONLY a valid JSON array — no markdown, no explanation:
[
  {{"index": 0, "deal_type": "VALUE-ADD", "deal_type_reason": "30 yr old structure covers 68% of site", "flag": "WATCH"}},
  {{"index": 1, "deal_type": "DEVELOPMENT", "deal_type_reason": "vacant land parcel", "flag": "STRONG BUY"}},
  ...
]"""

    t0 = time.time()
    response = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    elapsed = time.time() - t0
    result_text = response.content[0].text if response.content else ""
    print(f"  Claude responded in {elapsed:.1f}s  ({len(result_text)} chars)")

    # Extract JSON array
    arr = None
    m = re.search(r"```(?:json)?\s*(\[[\s\S]*?\])\s*```", result_text)
    if m:
        try:
            arr = json.loads(m.group(1))
        except Exception:
            pass
    if arr is None:
        start = result_text.find("[")
        if start >= 0:
            depth = 0
            for i, ch in enumerate(result_text[start:], start):
                if ch == "[":
                    depth += 1
                elif ch == "]":
                    depth -= 1
                    if depth == 0:
                        try:
                            arr = json.loads(result_text[start : i + 1])
                        except Exception:
                            pass
                        break

    if not arr:
        print("  ⚠ Could not parse Claude's response — deal_type left blank")
        return listings

    for item in arr:
        idx = item.get("index")
        if idx is not None and 0 <= int(idx) < len(listings):
            listings[int(idx)].deal_type        = item.get("deal_type", "UNKNOWN")
            listings[int(idx)].deal_type_reason = item.get("deal_type_reason", "")
            listings[int(idx)].flag             = item.get("flag", "")

    return listings


# ── Excel export ──────────────────────────────────────────────────────────────

COLUMNS = [
    ("Listing Name",        "listing_name",       28),
    ("Address",             "address",            24),
    ("City",                "city",               14),
    ("State",               "state",               6),
    ("Zip Code",            "zip_code",            9),
    ("Price",               "price",              14),
    ("Total Sqft",          "total_sqft",         12),
    ("Price / Sqft",        "price_per_sqft",     12),
    ("Year Built",          "year_built",         10),
    ("Lot Size (Acres)",    "lot_size_acres",     14),
    ("Zoning",              "zoning",             10),
    ("Days on Market",      "days_on_market",     13),
    ("Property Type",       "property_type",      18),
    ("Source",              "listing_source",     12),
    ("Listing URL",         "listing_url",        42),
    ("Deal Type",           "deal_type",          13),
    ("Deal Type Reason",    "deal_type_reason",   46),
    ("FLAG",                "flag",               12),
]

C_HDR  = "1F3864"
C_VA   = "E2EFDA"
C_DEV  = "DDEBF7"
C_UNK  = "FFF2CC"
C_SB   = "C00000"
C_BA   = "ED7D31"
C_WA   = "FFE699"
T_SIDE = Side(style="thin", color="D9D9D9")
T_BORD = Border(left=T_SIDE, right=T_SIDE, top=T_SIDE, bottom=T_SIDE)


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def export_excel(listings: list[Listing], filepath: str, dupes: int):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Listings"

    hf = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    for ci, (hdr, _, w) in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=ci, value=hdr)
        c.font      = hf
        c.fill      = _fill(C_HDR)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = T_BORD
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    url_ci = next(i for i, (_, f, _) in enumerate(COLUMNS, 1) if f == "listing_url")

    for ri, lst in enumerate(listings, 2):
        row_bg = {"VALUE-ADD": C_VA, "DEVELOPMENT": C_DEV}.get(lst.deal_type, C_UNK)

        for ci, (_, fname, _) in enumerate(COLUMNS, 1):
            val = getattr(lst, fname, None)
            if fname == "price" and val is not None:
                val = float(val)
            elif fname in ("total_sqft", "year_built", "days_on_market") and val is not None:
                val = int(val)
            elif fname in ("price_per_sqft", "lot_size_acres") and val is not None:
                val = float(val)

            c = ws.cell(row=ri, column=ci, value=val)
            c.fill      = _fill(row_bg)
            c.border    = T_BORD
            c.alignment = Alignment(vertical="top", wrap_text=True)

            if fname == "price" and val is not None:
                c.number_format = '$#,##0'
            elif fname == "price_per_sqft" and val is not None:
                c.number_format = '$#,##0.00'
            elif fname == "lot_size_acres" and val is not None:
                c.number_format = '0.00'
            elif fname == "total_sqft" and val is not None:
                c.number_format = '#,##0'

            if fname == "flag":
                styles = {
                    "STRONG BUY": (C_SB, Font(name="Calibri", bold=True, color="FFFFFF", size=10)),
                    "BELOW AVG":  (C_BA, Font(name="Calibri", bold=True, size=10)),
                    "WATCH":      (C_WA, Font(name="Calibri", bold=True, size=10)),
                }
                if val in styles:
                    c.fill      = _fill(styles[val][0])
                    c.font      = styles[val][1]
                    c.alignment = Alignment(horizontal="center", vertical="center")

            if ci == url_ci and val and str(val).startswith("http"):
                c.hyperlink = str(val)
                c.font      = Font(name="Calibri", color="0563C1",
                                   underline="single", size=10)

        ws.row_dimensions[ri].height = 48

    # Summary sheet
    ss = wb.create_sheet("Summary")
    ss.column_dimensions["A"].width = 32
    ss.column_dimensions["B"].width = 16

    def sh(row, text):
        c = ss.cell(row=row, column=1, value=text)
        c.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        c.fill      = _fill(C_HDR)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ss.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ss.row_dimensions[row].height = 20

    def sr(row, label, value):
        a = ss.cell(row=row, column=1, value=label)
        b = ss.cell(row=row, column=2, value=value)
        a.font      = Font(name="Calibri", size=10)
        b.font      = Font(name="Calibri", size=10, bold=True)
        b.alignment = Alignment(horizontal="right")
        for x in (a, b):
            x.border = T_BORD

    total = len(listings)
    va  = sum(1 for l in listings if l.deal_type == "VALUE-ADD")
    dev = sum(1 for l in listings if l.deal_type == "DEVELOPMENT")
    unk = sum(1 for l in listings if l.deal_type == "UNKNOWN")
    sb  = sum(1 for l in listings if l.flag == "STRONG BUY")
    ba  = sum(1 for l in listings if l.flag == "BELOW AVG")
    wa  = sum(1 for l in listings if l.flag == "WATCH")

    r = 1
    sh(r, "PropScout V2 Phase 1 — SA Metro Summary");    r += 1
    sh(r, "Run info");                                     r += 1
    sr(r, "Report date",  datetime.now().strftime("%Y-%m-%d %H:%M")); r += 1
    sr(r, "Geography",    "San Antonio TX metro");         r += 1
    sh(r, "Listing counts");                               r += 1
    sr(r, "Total unique listings", total);                 r += 1
    sr(r, "Duplicates removed",    dupes);                 r += 1
    sh(r, "Deal type breakdown");                          r += 1
    sr(r, "VALUE-ADD",   va);   r += 1
    sr(r, "DEVELOPMENT", dev);  r += 1
    sr(r, "UNKNOWN",     unk);  r += 1
    sh(r, "Flagged listings");                             r += 1
    sr(r, "STRONG BUY",  sb);   r += 1
    sr(r, "BELOW AVG",   ba);   r += 1
    sr(r, "WATCH",       wa);   r += 1
    sh(r, "By platform");                                  r += 1
    for platform in ("Crexi", "LoopNet"):
        n = sum(1 for l in listings if l.listing_source == platform)
        sr(r, platform, n);                                r += 1

    wb.save(filepath)


# ── Terminal summary ──────────────────────────────────────────────────────────

def print_summary(listings: list[Listing], dupes: int,
                  excel_path: str, json_path: str):
    W  = 62
    hr = lambda c="─": print(c * W)

    total = len(listings)
    va    = sum(1 for l in listings if l.deal_type == "VALUE-ADD")
    dev   = sum(1 for l in listings if l.deal_type == "DEVELOPMENT")
    unk   = sum(1 for l in listings if l.deal_type == "UNKNOWN")
    sb    = sum(1 for l in listings if l.flag == "STRONG BUY")
    ba    = sum(1 for l in listings if l.flag == "BELOW AVG")
    wa    = sum(1 for l in listings if l.flag == "WATCH")

    print(); hr("═")
    print("  PROPSCOUT V2 PHASE 1 — FINAL SUMMARY".center(W))
    hr("═")
    print(f"  Total unique listings:     {total}")
    print(f"  Duplicates removed:        {dupes}")
    hr()
    print("  DEAL TYPE BREAKDOWN")
    print(f"    VALUE-ADD:               {va}")
    print(f"    DEVELOPMENT:             {dev}")
    print(f"    UNKNOWN:                 {unk}")
    hr()
    print("  FLAGGED LISTINGS")
    print(f"    STRONG BUY:              {sb}")
    print(f"    BELOW AVG:               {ba}")
    print(f"    WATCH:                   {wa}")
    hr()
    print("  BY PLATFORM")
    for p in ("Crexi", "LoopNet"):
        n = sum(1 for l in listings if l.listing_source == p)
        print(f"    {p:<14} {n} listings")
    hr()
    if sb or ba or wa:
        print("  FLAGGED LISTING DETAILS")
        for lst in listings:
            if lst.flag:
                addr = f"{lst.address}, {lst.city}"[:40]
                psf  = f"${lst.price_per_sqft:.2f}/sqft" if lst.price_per_sqft else "no PSF"
                print(f"    [{lst.flag:<10}] {addr:<42} {psf}")
        hr()
    hr("═")
    print(f"  Excel  → {excel_path}")
    print(f"  JSON   → {json_path}")
    hr("═"); print()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    W  = 62
    hr = lambda c="─": print(c * W)

    print(); hr("═")
    print("  PropScout V2 — Scraper + Claude Classifier".center(W))
    hr("═")
    print("  Scraping:  Crexi (industrial + land)  ·  LoopNet (commercial)")
    print("  API calls: 1  (Claude batch classification)")
    print(f"  Output:    {EXCEL_PATH}  ·  {RAW_JSON}")
    hr(); print()

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY not set.")
        sys.exit(1)

    session = requests.Session()
    session.headers.update(HEADERS)

    # ── Step 1: Scrape ────────────────────────────────────────────────────────
    all_raw: list[Listing] = []

    for target in SCRAPE_TARGETS:
        print(f"\n── {target['platform']} · {target['prop_type']} ──────────────")
        try:
            if target["platform"] == "Crexi":
                results = scrape_crexi(session, target)
            else:
                results = scrape_loopnet(session, target)
            print(f"  → {len(results)} listing(s) scraped")
            all_raw.extend(results)
        except Exception as e:
            print(f"  ✗ scrape error: {e}")

    print(f"\n  Total raw listings scraped: {len(all_raw)}")
    hr()

    if not all_raw:
        print("  No listings scraped — both platforms may be blocking requests.")
        print("  Try running with a residential IP or adding session cookies.")
        sys.exit(0)

    # ── Step 2: PSF + dedup ───────────────────────────────────────────────────
    all_raw = [calculate_psf(l) for l in all_raw]
    listings, dupes = deduplicate(all_raw)
    print(f"  Dedup: {dupes} removed  |  {len(listings)} unique")
    hr()

    # ── Step 3: Save raw JSON (before Claude) ─────────────────────────────────
    raw_dicts = []
    for l in listings:
        d = {f: getattr(l, f) for f in vars(l) if getattr(l, f) not in (None, "", 0)}
        raw_dicts.append(d)
    with open(RAW_JSON, "w", encoding="utf-8") as fh:
        json.dump(raw_dicts, fh, indent=2, default=str)
    print(f"  Raw JSON saved → {RAW_JSON}  ({len(raw_dicts)} records)")
    hr()

    # ── Step 4: One Claude call — classify + flag ─────────────────────────────
    client   = anthropic.Anthropic(api_key=api_key)
    listings = classify_with_claude(client, listings)
    va  = sum(1 for l in listings if l.deal_type == "VALUE-ADD")
    dev = sum(1 for l in listings if l.deal_type == "DEVELOPMENT")
    unk = sum(1 for l in listings if l.deal_type == "UNKNOWN")
    print(f"  Deal types: VALUE-ADD={va}  DEVELOPMENT={dev}  UNKNOWN={unk}")
    hr()

    # ── Step 5: Export ────────────────────────────────────────────────────────
    print(f"  Exporting → {EXCEL_PATH} …")
    export_excel(listings, EXCEL_PATH, dupes)

    print_summary(listings, dupes, EXCEL_PATH, RAW_JSON)


if __name__ == "__main__":
    main()
