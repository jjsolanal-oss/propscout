#!/usr/bin/env python3
"""
PropScout V2 Phase 1
Commercial RE scraper — San Antonio metro
Platforms: Crexi, LoopNet, Brevitas
Types:     flex industrial, small bay industrial, land
"""

import anthropic
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Constants ─────────────────────────────────────────────────────────────────

MODEL = "claude-sonnet-4-0"

PLATFORMS = ["Crexi", "LoopNet", "Brevitas"]

PROPERTY_TYPES = ["flex industrial", "small bay industrial", "land"]

METRO_CITIES = [
    "San Antonio", "Schertz", "New Braunfels", "Converse",
    "Universal City", "Live Oak", "Helotes", "Boerne", "Seguin", "San Marcos",
]

CURRENT_YEAR        = datetime.now().year
VALUE_ADD_MIN_AGE   = 10       # structure must be ≥ this many years old
SITE_UNCOV_MAX      = 0.50     # VALUE-ADD requires < 50% undeveloped site
SQF_PER_ACRE        = 43_560

# Flagging thresholds
VA_PSF_THRESHOLD    = 120.0    # $/sqft — value-add
DEV_PSF_THRESHOLD   = 5.0      # $/sqft — land/development
BELOW_AVG_PCT       = 0.10     # 10 % below category avg

# Dedup thresholds
ADDR_SIM_MIN        = 0.80
PRICE_TOL           = 0.05     # 5 % price tolerance

# Seconds to pause between API calls (avoid rate-limit bursts)
INTER_CALL_PAUSE    = 10

# ── Data model ────────────────────────────────────────────────────────────────

@dataclass
class Listing:
    listing_name:     str            = ""
    address:          str            = ""
    city:             str            = ""
    state:            str            = "TX"
    zip_code:         str            = ""
    price:            Optional[float] = None
    total_sqft:       Optional[float] = None
    price_per_sqft:   Optional[float] = None   # calculated
    year_built:       Optional[int]   = None
    lot_size_acres:   Optional[float] = None
    zoning:           Optional[str]   = None
    days_on_market:   Optional[int]   = None
    property_type:    str            = ""
    listing_source:   str            = ""
    listing_url:      str            = ""
    also_listed_on:   str            = ""
    deal_type:        str            = ""      # VALUE-ADD | DEVELOPMENT | UNKNOWN
    deal_type_reason: str            = ""
    flag:             str            = ""      # STRONG BUY | BELOW AVG | WATCH | ""


# ── Search prompt ─────────────────────────────────────────────────────────────

def build_search_prompt(platform: str, prop_type: str) -> str:
    cities = ", ".join(METRO_CITIES)
    platform_domain = f"{platform.lower()}.com"
    return f"""You are a commercial real estate data extraction agent.

YOUR TASK: Search {platform} for {prop_type} listings FOR SALE in the San Antonio TX metro area.

TARGET PLATFORM: {platform} ({platform_domain})
PROPERTY TYPE:   {prop_type}
GEOGRAPHY:       {cities} — all in Texas

SEARCH STRATEGY (run at least 5 searches):
1. Search {platform_domain} directly for "{prop_type} for sale San Antonio TX"
2. Google: site:{platform_domain} "San Antonio" "{prop_type}"
3. Google: {platform} "{prop_type}" "San Antonio" OR "Schertz" OR "New Braunfels" OR "Boerne" for sale price
4. Search for listings in each suburb: Schertz TX, New Braunfels TX, Helotes TX, Boerne TX, Seguin TX, San Marcos TX
5. Try broader commercial terms if specific terms fail (e.g. "industrial building for sale", "warehouse flex space")
6. Run additional searches to reach as many listings as possible — target 8–15 listings

EXTRACT FOR EACH LISTING:
- listing_name:    title or property name from the listing page
- address:         street address only (no city or state)
- city:            city name
- state:           TX
- zip_code:        5-digit ZIP
- price:           asking price as plain number (no $, no commas), null if not shown
- total_sqft:      building square footage as plain number (for land: total lot sqft if known), null if unavailable
- year_built:      4-digit integer, null if unavailable
- lot_size_acres:  lot size in decimal acres, null if unavailable
- zoning:          zoning code or designation, null if unavailable
- days_on_market:  integer number of days, null if unavailable
- listing_url:     direct URL to the specific listing on {platform_domain}

CRITICAL: Only include listings in these cities (TX): {cities}.
Skip any listing outside this geography.

After all searches, return ONLY a JSON object (no other text, no markdown):
{{
  "listings": [
    {{
      "listing_name": "string",
      "address": "string",
      "city": "string",
      "state": "TX",
      "zip_code": "string or null",
      "price": 1234567,
      "total_sqft": 12000,
      "year_built": 1998,
      "lot_size_acres": 1.5,
      "zoning": "I-1",
      "days_on_market": 45,
      "listing_url": "https://..."
    }}
  ],
  "search_notes": "brief summary of what you found and any limitations"
}}"""


# ── JSON extraction ───────────────────────────────────────────────────────────

def extract_json(text: str) -> Optional[dict]:
    """Pull the first valid JSON object out of Claude's response."""
    # Try fenced code block first
    m = re.search(r"```(?:json)?\s*(\{[\s\S]*?\})\s*```", text)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass

    # Walk the string to find the outermost { ... }
    start = text.find("{")
    if start == -1:
        return None
    depth = 0
    for i, ch in enumerate(text[start:], start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(text[start : i + 1])
                except json.JSONDecodeError:
                    break
    return None


# ── Type coercions ────────────────────────────────────────────────────────────

def _float(val) -> Optional[float]:
    if val is None:
        return None
    try:
        v = float(str(val).replace(",", "").replace("$", "").strip())
        return v if v > 0 else None
    except (ValueError, TypeError):
        return None


def _int(val) -> Optional[int]:
    if val is None:
        return None
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None


# ── Platform search (one API call) ───────────────────────────────────────────

def search_platform(client: anthropic.Anthropic,
                    platform: str, prop_type: str) -> list[Listing]:
    """Call Claude with web_search to extract listings from one platform/type."""
    prompt   = build_search_prompt(platform, prop_type)
    messages = [{"role": "user", "content": prompt}]

    full_text    = ""
    search_count = 0
    tool_buf     = ""
    in_tool      = False

    metro_lower = {c.lower() for c in METRO_CITIES}

    for _continuation in range(4):
        try:
            with client.messages.stream(
                model=MODEL,
                max_tokens=6000,
                tools=[{
                    "type": "web_search_20260209",
                    "name": "web_search",
                    "allowed_callers": ["direct"],
                }],
                messages=messages,
            ) as stream:
                for event in stream:
                    etype = getattr(event, "type", None)

                    if etype == "content_block_start":
                        cb = event.content_block
                        if cb.type == "server_tool_use":
                            in_tool  = True
                            tool_buf = ""
                        elif cb.type == "text":
                            in_tool = False

                    elif etype == "content_block_delta":
                        delta = event.delta
                        dtype = getattr(delta, "type", None)
                        if dtype == "input_json_delta" and in_tool:
                            tool_buf += getattr(delta, "partial_json", "")
                        elif dtype == "text_delta":
                            full_text += delta.text

                    elif etype == "content_block_stop" and in_tool:
                        search_count += 1
                        try:
                            q = json.loads(tool_buf).get("query", "")
                            print(f"      [{search_count}] {q[:70]}")
                        except Exception:
                            pass
                        in_tool = False

                final = stream.get_final_message()

        except anthropic.RateLimitError:
            wait = 60
            print(f"      ⚠ Rate limit — waiting {wait}s …")
            time.sleep(wait)
            continue

        if final.stop_reason != "pause_turn":
            break
        messages.append({"role": "assistant", "content": final.content})

    # ── Parse JSON ────────────────────────────────────────────────────────
    data = extract_json(full_text)
    if not data or "listings" not in data:
        print(f"      ⚠ No parseable JSON returned")
        return []

    notes = data.get("search_notes", "")
    if notes:
        print(f"      Notes: {notes[:120]}")

    listings: list[Listing] = []
    for raw in data.get("listings", []):
        if not isinstance(raw, dict):
            continue

        city = str(raw.get("city") or "").strip()

        # Geography filter
        if city and city.lower() not in metro_lower:
            if "san antonio" not in city.lower():
                continue

        # Need at least an address or a name to be useful
        if not raw.get("address") and not raw.get("listing_name"):
            continue

        lst = Listing(
            listing_name   = str(raw.get("listing_name") or "").strip(),
            address        = str(raw.get("address")      or "").strip(),
            city           = city,
            state          = "TX",
            zip_code       = str(raw.get("zip_code")     or "").strip(),
            price          = _float(raw.get("price")),
            total_sqft     = _float(raw.get("total_sqft")),
            year_built     = _int(raw.get("year_built")),
            lot_size_acres = _float(raw.get("lot_size_acres")),
            zoning         = raw.get("zoning") or None,
            days_on_market = _int(raw.get("days_on_market")),
            property_type  = prop_type,
            listing_source = platform,
            listing_url    = str(raw.get("listing_url") or "").strip(),
        )
        listings.append(lst)

    return listings


# ── PSF calculation ───────────────────────────────────────────────────────────

def calculate_psf(lst: Listing) -> Listing:
    """Compute price_per_sqft; for land use lot acres → sqft if needed."""
    if lst.price is None:
        return lst

    sqft = lst.total_sqft
    if sqft is None and lst.lot_size_acres:
        sqft = lst.lot_size_acres * SQF_PER_ACRE

    if sqft and sqft > 0:
        lst.price_per_sqft = round(lst.price / sqft, 2)

    return lst


# ── Deduplication ─────────────────────────────────────────────────────────────

def _normalize_addr(addr: str) -> str:
    addr = addr.lower().strip()
    abbrevs = {
        r"\bst\b": "street",   r"\bave?\b": "avenue",  r"\bblvd\b": "boulevard",
        r"\bdr\b": "drive",    r"\brd\b":   "road",     r"\bln\b":   "lane",
        r"\bct\b": "court",    r"\bpl\b":   "place",    r"\bpkwy\b": "parkway",
        r"\bhwy\b": "highway", r"\bfwy\b":  "freeway",  r"\bste\b":  "suite",
        r"\bn\b":  "north",    r"\bs\b":    "south",    r"\be\b":    "east",
        r"\bw\b":  "west",
    }
    for pat, rep in abbrevs.items():
        addr = re.sub(pat, rep, addr)
    addr = re.sub(r"[^\w\s]", " ", addr)
    return re.sub(r"\s+", " ", addr).strip()


def _addr_sim(a: str, b: str) -> float:
    return SequenceMatcher(None, _normalize_addr(a), _normalize_addr(b)).ratio()


def _completeness(lst: Listing) -> int:
    fields = [
        lst.listing_name, lst.address, lst.city, lst.zip_code,
        lst.price, lst.total_sqft, lst.year_built, lst.lot_size_acres,
        lst.zoning, lst.days_on_market, lst.listing_url,
    ]
    return sum(1 for f in fields if f is not None and str(f).strip() != "")


def _prices_close(p1: Optional[float], p2: Optional[float]) -> bool:
    if p1 is None or p2 is None or p1 == 0 or p2 == 0:
        return True   # unknown price → can't rule out duplicate
    return abs(p1 - p2) / max(p1, p2) <= PRICE_TOL


def deduplicate(listings: list[Listing]) -> tuple[list[Listing], int]:
    """Remove cross-platform duplicates. Returns (unique_list, removed_count)."""
    kept: list[Listing] = []
    removed = 0

    for lst in listings:
        matched = None
        for k in kept:
            if not lst.address or not k.address:
                continue
            if (_addr_sim(lst.address, k.address) >= ADDR_SIM_MIN
                    and _prices_close(lst.price, k.price)):
                matched = k
                break

        if matched is None:
            kept.append(lst)
        else:
            removed += 1
            # Record the additional source
            src = lst.listing_source
            if src and src not in (matched.listing_source or ""):
                if matched.also_listed_on:
                    if src not in matched.also_listed_on:
                        matched.also_listed_on += f", {src}"
                else:
                    matched.also_listed_on = src

            # Merge any fields that the keeper is missing
            for fname, fval in vars(lst).items():
                if fname in ("listing_source", "also_listed_on",
                             "deal_type", "deal_type_reason", "flag"):
                    continue
                if getattr(matched, fname) in (None, "", 0):
                    setattr(matched, fname, fval)

    return kept, removed


# ── Deal classification ───────────────────────────────────────────────────────

def classify_deal(lst: Listing) -> Listing:
    ptype = lst.property_type.lower()

    # ── Land is always DEVELOPMENT ────────────────────────────────────────
    if ptype == "land":
        lst.deal_type        = "DEVELOPMENT"
        lst.deal_type_reason = "Vacant land parcel — classified as development opportunity"
        return lst

    # ── Derive building age ───────────────────────────────────────────────
    age = (CURRENT_YEAR - lst.year_built) if lst.year_built else None

    # ── Derive site coverage ratio ────────────────────────────────────────
    site_coverage = None
    if lst.total_sqft and lst.lot_size_acres and lst.lot_size_acres > 0:
        site_coverage = lst.total_sqft / (lst.lot_size_acres * SQF_PER_ACRE)

    # ── No structure evidence → DEVELOPMENT ──────────────────────────────
    if lst.year_built is None and lst.total_sqft is None:
        lst.deal_type        = "DEVELOPMENT"
        lst.deal_type_reason = ("No year-built or building sqft recorded — "
                                "likely vacant or land-only")
        return lst

    # ── Structure too new → UNKNOWN ───────────────────────────────────────
    if age is not None and age < VALUE_ADD_MIN_AGE:
        lst.deal_type        = "UNKNOWN"
        lst.deal_type_reason = (f"Structure is only {age} yr(s) old; VALUE-ADD requires "
                                f"≥{VALUE_ADD_MIN_AGE} yrs — may be new construction")
        return lst

    # ── Structure old enough; check site coverage ─────────────────────────
    if age is not None and age >= VALUE_ADD_MIN_AGE:
        if site_coverage is None:
            lst.deal_type        = "UNKNOWN"
            lst.deal_type_reason = (f"Structure is {age} yrs old (qualifies), but lot-size "
                                    "data unavailable to verify < 50% undeveloped rule")
        elif site_coverage > (1 - SITE_UNCOV_MAX):
            # Building covers > 50% of site → VALUE-ADD
            lst.deal_type        = "VALUE-ADD"
            lst.deal_type_reason = (f"Structure {age} yrs old; building covers "
                                    f"~{site_coverage*100:.0f}% of site (< 50% undeveloped)")
        else:
            # ≥ 50% of site is undeveloped → DEVELOPMENT
            undev_pct = (1 - site_coverage) * 100
            lst.deal_type        = "DEVELOPMENT"
            lst.deal_type_reason = (f"Structure {age} yrs old but ~{undev_pct:.0f}% of site "
                                    "is undeveloped (≥ 50% threshold)")
        return lst

    # ── Year built missing but sqft present ──────────────────────────────
    if lst.total_sqft:
        lst.deal_type        = "UNKNOWN"
        lst.deal_type_reason = ("Building sqft present but year built unknown — "
                                "cannot confirm ≥10-yr VALUE-ADD age requirement")
    else:
        lst.deal_type        = "DEVELOPMENT"
        lst.deal_type_reason = "No year built and no building sqft — likely vacant"

    return lst


# ── Flagging ──────────────────────────────────────────────────────────────────

def apply_flags(listings: list[Listing]) -> list[Listing]:
    """Compute category PSF averages and stamp FLAG on every listing."""
    va_psf  = [l.price_per_sqft for l in listings
               if l.deal_type == "VALUE-ADD"   and l.price_per_sqft]
    dev_psf = [l.price_per_sqft for l in listings
               if l.deal_type != "VALUE-ADD"   and l.price_per_sqft]

    va_avg  = sum(va_psf)  / len(va_psf)  if va_psf  else None
    dev_avg = sum(dev_psf) / len(dev_psf) if dev_psf else None

    print(f"  VALUE-ADD avg $/sqft:        "
          f"{'${:.2f}'.format(va_avg)  if va_avg  else 'n/a'} "
          f"(n={len(va_psf)})")
    print(f"  DEVELOPMENT avg $/sqft:      "
          f"{'${:.2f}'.format(dev_avg) if dev_avg else 'n/a'} "
          f"(n={len(dev_psf)})")

    for lst in listings:
        psf = lst.price_per_sqft
        if psf is None:
            lst.flag = ""
            continue

        is_va       = (lst.deal_type == "VALUE-ADD")
        threshold   = VA_PSF_THRESHOLD  if is_va else DEV_PSF_THRESHOLD
        avg         = va_avg            if is_va else dev_avg

        at_threshold = psf <= threshold
        below_avg    = (avg is not None and psf <= avg * (1 - BELOW_AVG_PCT))

        if at_threshold and below_avg:
            lst.flag = "STRONG BUY"
        elif below_avg:
            lst.flag = "BELOW AVG"
        elif at_threshold:
            lst.flag = "WATCH"
        else:
            lst.flag = ""

    return listings


# ── Excel export ──────────────────────────────────────────────────────────────

COLUMNS = [
    # (header,              field_name,         col_width)
    ("Listing Name",        "listing_name",      28),
    ("Address",             "address",           24),
    ("City",                "city",              14),
    ("State",               "state",              6),
    ("Zip Code",            "zip_code",           9),
    ("Price",               "price",             14),
    ("Total Sqft",          "total_sqft",        12),
    ("Price / Sqft",        "price_per_sqft",    12),
    ("Year Built",          "year_built",        10),
    ("Lot Size (Acres)",    "lot_size_acres",    14),
    ("Zoning",              "zoning",            10),
    ("Days on Market",      "days_on_market",    13),
    ("Property Type",       "property_type",     18),
    ("Listing Source",      "listing_source",    14),
    ("Also Listed On",      "also_listed_on",    16),
    ("Listing URL",         "listing_url",       42),
    ("Deal Type",           "deal_type",         13),
    ("Deal Type Reason",    "deal_type_reason",  46),
    ("FLAG",                "flag",              12),
]

# Palette
C_HDR_BG  = "1F3864"   # dark navy
C_VA_BG   = "E2EFDA"   # soft green  → VALUE-ADD rows
C_DEV_BG  = "DDEBF7"   # soft blue   → DEVELOPMENT rows
C_UNK_BG  = "FFF2CC"   # soft yellow → UNKNOWN rows
C_SB_BG   = "C00000"   # red         → STRONG BUY cell
C_BA_BG   = "ED7D31"   # orange      → BELOW AVG cell
C_WA_BG   = "FFE699"   # pale amber  → WATCH cell
C_ALT_BG  = "F2F2F2"   # alternating even-row tint
THIN_SIDE = Side(style="thin", color="D9D9D9")
THIN_BORD = Border(left=THIN_SIDE, right=THIN_SIDE,
                   top=THIN_SIDE,  bottom=THIN_SIDE)


def _cell_fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def export_excel(listings: list[Listing], filepath: str,
                 dupes_removed: int):
    wb = openpyxl.Workbook()

    # ── Main data sheet ───────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Listings"

    # Header
    hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    for ci, (header, _, width) in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=ci, value=header)
        c.font      = hdr_font
        c.fill      = _cell_fill(C_HDR_BG)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border    = THIN_BORD
        ws.column_dimensions[get_column_letter(ci)].width = width

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    # Data rows
    url_col_idx = next(i for i, (_, f, _) in enumerate(COLUMNS, 1)
                       if f == "listing_url")

    for ri, lst in enumerate(listings, 2):
        # Row background by deal type
        if lst.deal_type == "VALUE-ADD":
            row_bg = C_VA_BG
        elif lst.deal_type == "DEVELOPMENT":
            row_bg = C_DEV_BG
        else:
            row_bg = C_UNK_BG

        for ci, (_, fname, _) in enumerate(COLUMNS, 1):
            val = getattr(lst, fname, None)

            # Coerce for Excel types
            if fname == "price" and val is not None:
                val = float(val)
            elif fname in ("total_sqft", "year_built", "days_on_market") and val is not None:
                val = int(val)
            elif fname in ("price_per_sqft", "lot_size_acres") and val is not None:
                val = float(val)

            c = ws.cell(row=ri, column=ci, value=val)
            c.fill      = _cell_fill(row_bg)
            c.border    = THIN_BORD
            c.alignment = Alignment(vertical="top", wrap_text=True)

            # Number formats
            if fname == "price" and val is not None:
                c.number_format = '$#,##0'
            elif fname == "price_per_sqft" and val is not None:
                c.number_format = '$#,##0.00'
            elif fname == "lot_size_acres" and val is not None:
                c.number_format = '0.00'
            elif fname == "total_sqft" and val is not None:
                c.number_format = '#,##0'

            # FLAG cell coloring
            if fname == "flag":
                if val == "STRONG BUY":
                    c.fill = _cell_fill(C_SB_BG)
                    c.font = Font(name="Calibri", bold=True,
                                  color="FFFFFF", size=10)
                    c.alignment = Alignment(horizontal="center",
                                            vertical="center")
                elif val == "BELOW AVG":
                    c.fill = _cell_fill(C_BA_BG)
                    c.font = Font(name="Calibri", bold=True, size=10)
                    c.alignment = Alignment(horizontal="center",
                                            vertical="center")
                elif val == "WATCH":
                    c.fill = _cell_fill(C_WA_BG)
                    c.font = Font(name="Calibri", bold=True, size=10)
                    c.alignment = Alignment(horizontal="center",
                                            vertical="center")

            # Hyperlink for URL column
            if ci == url_col_idx and val and str(val).startswith("http"):
                c.hyperlink = str(val)
                c.font      = Font(name="Calibri", color="0563C1",
                                   underline="single", size=10)

        ws.row_dimensions[ri].height = 48

    # ── Summary sheet ─────────────────────────────────────────────────────
    ss = wb.create_sheet("Summary")
    ss.column_dimensions["A"].width = 30
    ss.column_dimensions["B"].width = 16

    def ss_hdr(row, text):
        c = ss.cell(row=row, column=1, value=text)
        c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        c.fill = _cell_fill(C_HDR_BG)
        c.alignment = Alignment(horizontal="left", vertical="center",
                                indent=1)
        ss.merge_cells(start_row=row, start_column=1,
                       end_row=row,   end_column=2)
        ss.row_dimensions[row].height = 20

    def ss_row(row, label, value):
        a = ss.cell(row=row, column=1, value=label)
        b = ss.cell(row=row, column=2, value=value)
        a.font = Font(name="Calibri", size=10)
        b.font = Font(name="Calibri", size=10, bold=True)
        b.alignment = Alignment(horizontal="right")
        for c in (a, b):
            c.border = THIN_BORD

    total     = len(listings)
    va_n      = sum(1 for l in listings if l.deal_type == "VALUE-ADD")
    dev_n     = sum(1 for l in listings if l.deal_type == "DEVELOPMENT")
    unk_n     = sum(1 for l in listings if l.deal_type == "UNKNOWN")
    sb_n      = sum(1 for l in listings if l.flag == "STRONG BUY")
    ba_n      = sum(1 for l in listings if l.flag == "BELOW AVG")
    wa_n      = sum(1 for l in listings if l.flag == "WATCH")

    r = 1
    ss_hdr(r, "PropScout V2 Phase 1 — SA Metro Summary"); r += 1
    ss_hdr(r, "Run info"); r += 1
    ss_row(r, "Report date",
           datetime.now().strftime("%Y-%m-%d %H:%M")); r += 1
    ss_row(r, "Geography",
           "San Antonio TX metro + suburbs"); r += 1
    ss_hdr(r, "Listing counts"); r += 1
    ss_row(r, "Total unique listings", total); r += 1
    ss_row(r, "Duplicates removed", dupes_removed); r += 1
    ss_hdr(r, "Deal type breakdown"); r += 1
    ss_row(r, "VALUE-ADD",   va_n);  r += 1
    ss_row(r, "DEVELOPMENT", dev_n); r += 1
    ss_row(r, "UNKNOWN",     unk_n); r += 1
    ss_hdr(r, "Flagged listings"); r += 1
    ss_row(r, "STRONG BUY",  sb_n); r += 1
    ss_row(r, "BELOW AVG",   ba_n); r += 1
    ss_row(r, "WATCH",       wa_n); r += 1
    ss_hdr(r, "By platform"); r += 1
    for platform in PLATFORMS:
        primary = sum(1 for l in listings if l.listing_source == platform)
        cross   = sum(1 for l in listings
                      if platform in (l.also_listed_on or ""))
        ss_row(r, f"{platform}  (primary | cross-listed)",
               f"{primary} | {cross}"); r += 1
    ss_hdr(r, "By property type"); r += 1
    for pt in PROPERTY_TYPES:
        ss_row(r, pt.title(),
               sum(1 for l in listings if l.property_type == pt)); r += 1

    wb.save(filepath)


# ── Terminal summary ──────────────────────────────────────────────────────────

def print_summary(listings: list[Listing], dupes_removed: int, filepath: str):
    W = 62

    def hr(c="─"):
        print(c * W)

    total = len(listings)
    va_n  = sum(1 for l in listings if l.deal_type == "VALUE-ADD")
    dev_n = sum(1 for l in listings if l.deal_type == "DEVELOPMENT")
    unk_n = sum(1 for l in listings if l.deal_type == "UNKNOWN")
    sb_n  = sum(1 for l in listings if l.flag == "STRONG BUY")
    ba_n  = sum(1 for l in listings if l.flag == "BELOW AVG")
    wa_n  = sum(1 for l in listings if l.flag == "WATCH")

    print()
    hr("═")
    print("  PROPSCOUT V2 PHASE 1 — FINAL SUMMARY".center(W))
    hr("═")
    print(f"  Total unique listings:       {total}")
    print(f"  Duplicates removed:          {dupes_removed}")
    hr()
    print("  DEAL TYPE BREAKDOWN")
    print(f"    VALUE-ADD:                 {va_n}")
    print(f"    DEVELOPMENT:               {dev_n}")
    print(f"    UNKNOWN:                   {unk_n}")
    hr()
    print("  FLAGGED LISTINGS")
    print(f"    STRONG BUY  (≤threshold & ≥10% below avg):  {sb_n}")
    print(f"    BELOW AVG   (≥10% below category avg):      {ba_n}")
    print(f"    WATCH       (≤threshold only):              {wa_n}")
    hr()
    print("  BY PLATFORM")
    for p in PLATFORMS:
        pri   = sum(1 for l in listings if l.listing_source == p)
        cross = sum(1 for l in listings if p in (l.also_listed_on or ""))
        print(f"    {p:<12}  primary={pri:<4} cross-listed={cross}")
    hr()
    print("  BY PROPERTY TYPE")
    for pt in PROPERTY_TYPES:
        n = sum(1 for l in listings if l.property_type == pt)
        print(f"    {pt:<26} {n}")
    hr()
    if sb_n or ba_n or wa_n:
        print("  FLAGGED LISTING DETAILS")
        for lst in listings:
            if lst.flag:
                addr = f"{lst.address}, {lst.city}"[:40]
                psf  = (f"${lst.price_per_sqft:.2f}/sqft"
                        if lst.price_per_sqft else "no PSF")
                print(f"    [{lst.flag:<10}] {addr:<42} {psf}")
        hr()
    hr("═")
    print(f"  Excel saved → {filepath}")
    hr("═")
    print()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    W = 62

    def hr(c="─"):
        print(c * W)

    print()
    hr("═")
    print("  PropScout V2 Phase 1 — SA Metro Commercial Scraper".center(W))
    hr("═")
    print(f"  Platforms:    {', '.join(PLATFORMS)}")
    print(f"  Types:        {', '.join(PROPERTY_TYPES)}")
    print(f"  Cities:       {len(METRO_CITIES)} metro cities")
    total_calls = len(PLATFORMS) * len(PROPERTY_TYPES)
    print(f"  API calls:    {total_calls} ({len(PLATFORMS)} platforms × "
          f"{len(PROPERTY_TYPES)} property types)")
    print(f"  Est. runtime: 8–15 minutes")
    hr()
    print()

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY not set.")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    # ── Phase 1: Search & extract ─────────────────────────────────────────
    all_raw: list[Listing] = []
    run = 0

    for platform in PLATFORMS:
        for prop_type in PROPERTY_TYPES:
            run += 1
            print(f"  [{run}/{total_calls}] {platform} · {prop_type}")
            try:
                results = search_platform(client, platform, prop_type)
            except Exception as exc:
                print(f"      ✗ Error: {exc}")
                results = []
            print(f"      → {len(results)} listing(s) extracted")
            all_raw.extend(results)

            if run < total_calls:
                time.sleep(INTER_CALL_PAUSE)

    print()
    print(f"  Raw listings before dedup: {len(all_raw)}")
    hr()

    if not all_raw:
        print("  No listings found. Check API key and network, then retry.")
        sys.exit(0)

    # ── Phase 2: PSF calculation ──────────────────────────────────────────
    all_raw = [calculate_psf(l) for l in all_raw]

    # ── Phase 3: Deduplicate ──────────────────────────────────────────────
    print("  Running deduplication …")
    listings, dupes_removed = deduplicate(all_raw)
    print(f"  Duplicates removed: {dupes_removed}")
    print(f"  Unique listings:    {len(listings)}")
    hr()

    # ── Phase 4: Classify ─────────────────────────────────────────────────
    print("  Classifying deals …")
    listings = [classify_deal(l) for l in listings]
    va_n  = sum(1 for l in listings if l.deal_type == "VALUE-ADD")
    dev_n = sum(1 for l in listings if l.deal_type == "DEVELOPMENT")
    unk_n = sum(1 for l in listings if l.deal_type == "UNKNOWN")
    print(f"  VALUE-ADD={va_n}  DEVELOPMENT={dev_n}  UNKNOWN={unk_n}")
    hr()

    # ── Phase 5: Flagging ─────────────────────────────────────────────────
    print("  Applying price/sqft flags …")
    listings = apply_flags(listings)
    hr()

    # ── Phase 6: Export ───────────────────────────────────────────────────
    filepath = "propscout_sa_phase1.xlsx"
    print(f"  Exporting to {filepath} …")
    export_excel(listings, filepath, dupes_removed)

    # ── Summary ───────────────────────────────────────────────────────────
    print_summary(listings, dupes_removed, filepath)


if __name__ == "__main__":
    main()
