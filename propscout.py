#!/usr/bin/env python3
"""PropScout — Real Estate Market Research CLI Agent"""

import anthropic
import json
import os
import sys
from datetime import datetime

# ── Property taxonomy ────────────────────────────────────────────────────────

PROPERTY_CATEGORIES = {
    "1": "residential",
    "2": "commercial",
    "3": "land",
}

PROPERTY_SUBTYPES = {
    "residential": {
        "1": "house",
        "2": "condo",
        "3": "townhouse",
        "4": "multi-family",
    },
    "commercial": {
        "1": "retail",
        "2": "warehouse",
        "3": "office",
        "4": "flex building",
        "5": "industrial",
    },
    "land": {
        "1": "residential lot",
        "2": "agricultural",
        "3": "ranch",
        "4": "timberland",
    },
}

SOURCE_MAP = {
    "residential": ["Zillow", "Redfin", "Realtor.com"],
    "commercial":  ["LoopNet", "Crexi", "Realtor.com"],
    "land":        ["Land.com", "Zillow", "Realtor.com"],
}

# ── UI helpers ────────────────────────────────────────────────────────────────

W = 64  # terminal width

def hr(char="─"):
    print(char * W)

def banner():
    print()
    hr("═")
    print("  PropScout  ·  Real Estate Market Research Agent".center(W))
    hr("═")
    print()

def section(title):
    print()
    hr()
    print(f"  {title}")
    hr()

def menu_choice(prompt: str, options: dict) -> str:
    print(f"\n{prompt}")
    for k, v in options.items():
        print(f"   {k}.  {v.title()}")
    while True:
        choice = input("\n   Enter choice: ").strip()
        if choice in options:
            return options[choice]
        print(f"   Invalid — enter one of: {', '.join(options.keys())}")

def parse_dollars(raw: str) -> int:
    cleaned = raw.replace(",", "").replace("$", "").replace(" ", "")
    return int(float(cleaned))

def get_budget() -> tuple[int, int]:
    print("\nEnter your budget range:")
    while True:
        try:
            lo = parse_dollars(input("   Minimum ($): "))
            break
        except ValueError:
            print("   Please enter a valid number.")
    while True:
        try:
            hi = parse_dollars(input("   Maximum ($): "))
            if hi <= lo:
                print("   Maximum must be greater than minimum.")
                continue
            break
        except ValueError:
            print("   Please enter a valid number.")
    return lo, hi

# ── Prompt construction ───────────────────────────────────────────────────────

def build_prompt(location: str, category: str, subtype: str,
                 lo: int, hi: int) -> str:
    sources = ", ".join(SOURCE_MAP.get(category, ["Zillow", "Realtor.com"]))
    today   = datetime.now().strftime("%B %d, %Y")

    unit_label = "per acre" if category == "land" else "per sq ft"

    return f"""You are PropScout, an expert real estate market research agent. \
A buyer has hired you to research current listings and market conditions.

BUYER CRITERIA
--------------
Location:      {location}
Property Type: {subtype.title()} ({category.title()})
Budget Range:  ${lo:,} – ${hi:,}
Primary Data Sources: {sources}

RESEARCH INSTRUCTIONS
---------------------
1. Search {sources} for ACTIVE {subtype} listings in {location} priced \
between ${lo:,} and ${hi:,}. Use multiple searches to gather comprehensive data.
2. Look for recent sold data and market statistics for this property type in \
this specific area.
3. Identify 3–5 specific real current listings (real addresses, real prices, \
real details from your search results).
4. Collect: average asking price, price {unit_label}, average days on market, \
total active inventory near the budget range, and any price-trend data.

OUTPUT FORMAT
-------------
Produce the full report below. Fill in every placeholder with real data \
from your searches. Be specific — use actual numbers, actual addresses, \
actual listing details you found.

{'═' * 62}
      PROPSCOUT MARKET INTELLIGENCE REPORT
{'═' * 62}

  Location:       {location}
  Property Type:  {subtype.title()} ({category.title()})
  Budget Range:   ${lo:,} – ${hi:,}
  Report Date:    {today}

{'─' * 62}
  MARKET KPIs
{'─' * 62}

  Average Listing Price:        $[fill in]
  Median Price {unit_label:15s}   $[fill in]
  Average Days on Market:       [fill in] days
  Active Listings in Budget:    [fill in] properties
  30-Day Price Trend:           [↑ Up X% / ↓ Down X% / → Stable]
  Market Condition:             [Buyer's Market / Seller's Market / Balanced]

{'─' * 62}
  PRICE DISTRIBUTION
{'─' * 62}

  Below budget  (< ${lo:,}):           [X]% of active listings
  Within budget (${lo:,} – ${hi:,}):   [X]% of active listings   ← YOUR TARGET
  Above budget  (> ${hi:,}):           [X]% of active listings

  [2–3 sentences: what does this distribution mean for negotiating power?]

{'─' * 62}
  TOP RECOMMENDED LISTINGS
{'─' * 62}

  For each listing you found, use this format:

  #1
  Address:  [full street address, city, state]
  Price:    $[list price]
  Details:  [beds / baths / sq ft — or acreage / lot details for land]
  Source:   [website where found, e.g. Zillow or LoopNet]
  Listed:   [X days on market]
  Why buy:  [1–2 sentences on what makes this listing stand out]

  #2
  [same format]

  [Continue for 3–5 listings total]

{'─' * 62}
  MARKET NARRATIVE
{'─' * 62}

  PARAGRAPH 1 — CURRENT CONDITIONS:
  [Describe what is happening in this market right now. Is inventory rising
  or falling? How competitive is it for buyers in this price range? Any
  notable shifts in the past 30–60 days?]

  PARAGRAPH 2 — WHAT YOUR BUDGET BUYS:
  [Explain specifically what ${lo:,}–${hi:,} gets a buyer in {location}
  today. Are there specific neighborhoods or zip codes where the budget
  stretches further? Any emerging areas worth considering?]

  PARAGRAPH 3 — BUYER GUIDANCE:
  [Concrete action items: should the buyer act fast or wait? Offer strategy
  tips given current competition. Key risks to watch for. Any seasonal
  factors affecting timing?]

{'─' * 62}
  SOURCES CONSULTED
{'─' * 62}

  [List each URL, search query, or source you actually retrieved data from]

{'═' * 62}
  END OF REPORT  ·  Generated by PropScout
{'═' * 62}"""

# ── Research engine ───────────────────────────────────────────────────────────

def run_research(location: str, category: str, subtype: str,
                 lo: int, hi: int) -> str:
    client  = anthropic.Anthropic()
    prompt  = build_prompt(location, category, subtype, lo, hi)
    messages = [{"role": "user", "content": prompt}]

    print()
    hr()
    print("  Connecting to research agent …")
    hr()

    full_report     = ""
    search_count    = 0
    report_printing = False
    tool_input_buf  = ""
    in_tool_block   = False

    max_continuations = 5
    for _ in range(max_continuations):
        with client.messages.stream(
            model="claude-sonnet-4-0",
            max_tokens=8000,
            tools=[{"type": "web_search_20260209", "name": "web_search"}],
            messages=messages,
        ) as stream:
            for event in stream:
                etype = getattr(event, "type", None)

                if etype == "content_block_start":
                    cb = event.content_block
                    if cb.type == "server_tool_use":
                        in_tool_block  = True
                        tool_input_buf = ""
                    elif cb.type == "text":
                        in_tool_block = False
                        if not report_printing:
                            report_printing = True
                            if search_count:
                                print(f"\n  ✓ {search_count} web search(es) complete — "
                                      "generating report …\n")
                                hr()
                                print()
                    else:
                        in_tool_block = False

                elif etype == "content_block_delta":
                    delta = event.delta
                    dtype = getattr(delta, "type", None)

                    if dtype == "input_json_delta" and in_tool_block:
                        tool_input_buf += getattr(delta, "partial_json", "")

                    elif dtype == "text_delta":
                        text = delta.text
                        sys.stdout.write(text)
                        sys.stdout.flush()
                        full_report += text

                elif etype == "content_block_stop" and in_tool_block:
                    search_count += 1
                    # Parse search query for display
                    try:
                        query = json.loads(tool_input_buf).get("query", "…")
                    except (json.JSONDecodeError, AttributeError):
                        query = "market data"
                    sys.stdout.write(f"\r  [{search_count}] Searching: {query[:50]:<50}")
                    sys.stdout.flush()
                    in_tool_block  = False
                    tool_input_buf = ""

            final = stream.get_final_message()

        if final.stop_reason != "pause_turn":
            break

        # Server-side tool loop hit its iteration cap — continue
        messages.append({"role": "assistant", "content": final.content})

    return full_report

# ── Report persistence ────────────────────────────────────────────────────────

def save_report(report_text: str, location: str, category: str,
                subtype: str, lo: int, hi: int) -> str:
    timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
    loc_slug     = location.lower().replace(",", "").replace(" ", "_")
    sub_slug     = subtype.replace(" ", "_")
    filename     = f"propscout_{loc_slug}_{sub_slug}_{timestamp}.txt"

    header = (
        f"PropScout Market Research Report\n"
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"Location:  {location}\n"
        f"Type:      {subtype.title()} ({category.title()})\n"
        f"Budget:    ${lo:,} – ${hi:,}\n"
        f"{'=' * 62}\n\n"
    )

    with open(filename, "w", encoding="utf-8") as fh:
        fh.write(header + report_text)

    return filename

# ── Main flow ─────────────────────────────────────────────────────────────────

def main():
    banner()

    # API key check
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ERROR: ANTHROPIC_API_KEY is not set.")
        print("  Export your key and re-run:  export ANTHROPIC_API_KEY=sk-ant-...")
        sys.exit(1)

    print("Welcome! Let's research your target real estate market.\n")

    # ── Collect inputs ──────────────────────────────────────────────────────
    location = input("Target location (city, state — e.g. Austin, TX): ").strip()
    if not location:
        print("Location is required.")
        sys.exit(1)

    category = menu_choice("Property type category:", PROPERTY_CATEGORIES)
    subtype  = menu_choice(f"Select {category} subtype:", PROPERTY_SUBTYPES[category])
    lo, hi   = get_budget()

    # ── Confirm ─────────────────────────────────────────────────────────────
    section("SEARCH PARAMETERS")
    print(f"  Location:      {location}")
    print(f"  Property Type: {subtype.title()} ({category.title()})")
    print(f"  Budget:        ${lo:,} – ${hi:,}")
    hr()

    go = input("\n  Start research? (y/n): ").strip().lower()
    if go not in ("y", "yes"):
        print("\n  Research cancelled.")
        sys.exit(0)

    # ── Research ─────────────────────────────────────────────────────────────
    try:
        report = run_research(location, category, subtype, lo, hi)
    except anthropic.APIConnectionError:
        print("\n\nERROR: Could not connect to the API. Check your internet connection.")
        sys.exit(1)
    except anthropic.AuthenticationError:
        print("\n\nERROR: Invalid API key. Verify ANTHROPIC_API_KEY.")
        sys.exit(1)
    except anthropic.RateLimitError:
        print("\n\nERROR: Rate limit exceeded. Wait a moment and try again.")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n  Research interrupted by user.")
        sys.exit(0)

    if not report.strip():
        print("\n\nERROR: No report content was returned. Please try again.")
        sys.exit(1)

    # ── Save ──────────────────────────────────────────────────────────────────
    filename = save_report(report, location, category, subtype, lo, hi)

    print()
    hr("═")
    print(f"  Report saved → {filename}")
    hr("═")
    print()


if __name__ == "__main__":
    main()
