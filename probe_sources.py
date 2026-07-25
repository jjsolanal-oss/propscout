#!/usr/bin/env python3
"""Probe each PropScout source with a plain requests GET and report what
comes back: status, body size, and any recognisable bot-block signature.

Distinguishes three failure modes that look alike in the scraper diagnostics:
  reachable   — real content arrived; any zero-listing result is a parser bug
  edge block  — 403/503 with a tiny body; blocked before content is served
  challenge   — 200 with a JS interstitial instead of listings

Run:  python probe_sources.py
"""

import sys
import time

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

HEADERS = {
    "User-Agent":                UA,
    "Accept":                    "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language":           "en-US,en;q=0.9",
    "Connection":                "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

TARGETS = {
    "Brevitas":           "https://brevitas.com/buy?q=San+Antonio+TX+industrial",
    "LandBrokerMLS":      "https://www.landbrokermls.com/listings/?state_name=Texas&q=San+Antonio",
    "CommercialExchange": "https://www.commercialexchange.com/commercial-real-estate/san-antonio-tx/industrial/for-sale/",
    "Catylist":           "https://www.catylist.com/san-antonio-tx-commercial-real-estate/industrial/for-sale/",
    "Land.com":           "https://www.land.com/results/Texas/San-Antonio/",
}

# Substrings that identify the vendor doing the blocking.
SIGNATURES = [
    ("Cloudflare",   ("cf-ray", "cloudflare", "attention required", "cf-browser-verification")),
    ("Akamai",       ("akamai", "reference #", "access denied")),
    ("PerimeterX",   ("perimeterx", "px-captcha", "_px")),
    ("DataDome",     ("datadome", "dd_cookie")),
    ("Imperva",      ("incapsula", "imperva", "_incap_")),
    ("hCaptcha",     ("hcaptcha",)),
    ("reCAPTCHA",    ("recaptcha",)),
]

# Markers suggesting real listing content rather than a shell or interstitial.
CONTENT_MARKERS = ("__next_data__", "application/ld+json", "listing", "property",
                   "for sale", "$")


def detect(resp: requests.Response) -> str:
    blob = (resp.text[:20000] + " " + " ".join(resp.headers.keys())).lower()
    hits = [name for name, keys in SIGNATURES if any(k in blob for k in keys)]
    return ", ".join(hits) if hits else "none identified"


def classify(resp: requests.Response) -> str:
    size = len(resp.text)
    body = resp.text.lower()
    if resp.status_code in (401, 403, 405, 406, 503):
        return "EDGE BLOCK" if size < 5000 else "BLOCK PAGE"
    if resp.status_code == 429:
        return "RATE LIMITED"
    if resp.status_code != 200:
        return f"HTTP {resp.status_code}"
    if size < 5000:
        return "SUSPICIOUS (200 but tiny body)"
    if any(s in body for s in ("checking your browser", "enable javascript",
                               "verify you are human", "just a moment")):
        return "JS CHALLENGE"
    if any(m in body for m in CONTENT_MARKERS):
        return "REACHABLE"
    return "200 but no listing markers"


def main() -> int:
    print()
    print("=" * 78)
    print("  SOURCE REACHABILITY PROBE  (plain requests, browser headers)")
    print("=" * 78)

    reachable = []
    for name, url in TARGETS.items():
        print(f"\n  {name}")
        print(f"    url    : {url}")
        try:
            resp = requests.get(url, headers=HEADERS, timeout=20)
        except requests.RequestException as exc:
            print(f"    result : REQUEST FAILED — {type(exc).__name__}: {exc}")
            time.sleep(2)
            continue

        verdict = classify(resp)
        if verdict == "REACHABLE":
            reachable.append(name)

        print(f"    status : HTTP {resp.status_code}")
        print(f"    bytes  : {len(resp.text):,}")
        print(f"    server : {resp.headers.get('Server', '—')}")
        print(f"    blocker: {detect(resp)}")
        print(f"    result : {verdict}")

        fname = f"sample_{name.replace('.', '_').lower()}.html"
        try:
            with open(fname, "w", encoding="utf-8") as fh:
                fh.write(resp.text)
            print(f"    saved  : {fname}")
        except OSError as exc:
            print(f"    saved  : failed ({exc})")

        time.sleep(2)

    print()
    print("=" * 78)
    if reachable:
        print(f"  Reachable: {', '.join(reachable)}")
        print("  For these, HTTP works and zero listings means the parser needs fixing.")
    else:
        print("  No source served real content to a plain requests GET.")
        print("  Header tuning will not help — these blocks key on the TLS")
        print("  fingerprint, which requests cannot change.")
    print("=" * 78)
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
