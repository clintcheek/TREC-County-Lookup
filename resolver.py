from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import os
import re
import signal
import threading
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, fields
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz.fuzz import token_set_ratio

RESOLVER_VERSION = "v6"
USER_AGENT = "TexasBrokerCountyResolver/6.0 (private license-location research)"
ADDRESS_RE = re.compile(
    r"\b(\d{1,6}\s+[A-Za-z0-9.'#&\- ]{2,80}?\s(?:Street|St|Road|Rd|Avenue|Ave|Boulevard|Blvd|Drive|Dr|Lane|Ln|Court|Ct|Parkway|Pkwy|Highway|Hwy|Loop|Way|Trail|Trl|Circle|Cir|Plaza|Place|Pl|Terrace|Ter)\.?"
    r"(?:\s*(?:Suite|Ste|Unit|#)\s*[A-Za-z0-9\-]+)?\s*,?\s*[A-Za-z.'\- ]{2,45},?\s*(?:TX|Texas|[A-Z]{2})\s+\d{5}(?:-\d{4})?)\b",
    re.I,
)
ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")
PROPERTY_TERMS = re.compile(
    r"\b(for sale|for lease|sold|pending|bed(?:room)?s?|bath(?:room)?s?|sq\.?\s*ft|acre(?:s)?|mls|listing price|property details|home value|open house)\b",
    re.I,
)
PROPERTY_PATHS = ("/property/", "/homedetail/", "/listing/", "/listings/", "/realestateandhomes-detail/", "/homes/", "/home-details/", "/real-estate/")
LISTING_DOMAINS = {"zillow.com", "redfin.com", "realtor.com", "homes.com", "har.com", "trulia.com", "loopnet.com", "crexi.com", "land.com", "landwatch.com"}
OFFICE_PATH_TERMS = ("contact", "about", "office", "location", "locations", "company", "team", "brokerage")
CONTACT_TERMS = re.compile(r"\b(contact|office|location|headquarters|address|about us|our office)\b", re.I)
CORP_SUFFIXES = {"llc", "lp", "ltd", "inc", "corp", "corporation", "company", "co", "pllc", "llp", "realty", "real", "estate", "group"}

_thread = threading.local()
_cache_lock = threading.Lock()
_audit_lock = threading.Lock()


def now_utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def session() -> requests.Session:
    if not hasattr(_thread, "session"):
        s = requests.Session()
        s.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"})
        _thread.session = s
    return _thread.session


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def canonical_license(value: Any) -> str:
    return clean(value).upper()


def normalized_name(value: str) -> str:
    words = re.findall(r"[a-z0-9]+", clean(value).lower())
    kept = [w for w in words if w not in CORP_SUFFIXES]
    return " ".join(kept or words)


def normalize_address(value: str) -> str:
    x = clean(value).lower()
    replacements = {
        " street": " st", " road": " rd", " avenue": " ave", " boulevard": " blvd",
        " drive": " dr", " lane": " ln", " court": " ct", " parkway": " pkwy",
        " highway": " hwy", " suite ": " ste ", " texas ": " tx ",
    }
    for a, b in replacements.items():
        x = x.replace(a, b)
    return re.sub(r"[^a-z0-9]", "", x)


@dataclass
class Candidate:
    license_number: str
    brokerage_name: str
    related_broker_name: str
    address: str
    url: str
    source_domain: str
    source_tier: int
    source_type: str
    page_title: str
    evidence_text: str
    identity_score: float
    address_score: float
    negative_score: float
    total_score: float
    reject_reason: str = ""
    matched_address: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""
    county: str = ""


@dataclass
class Result:
    license_number: str
    brokerage_name: str
    related_broker_name: str = ""
    office_address: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""
    county: str = ""
    status: str = "Unresolved"
    confidence: float = 0.0
    identity_score: float = 0.0
    consensus_sources: int = 0
    evidence_url: str = ""
    secondary_evidence_url: str = ""
    evidence_type: str = ""
    notes: str = ""
    updated_at: str = ""
    resolver_version: str = RESOLVER_VERSION
    last_checked: str = ""
    evidence_count: int = 0
    needs_recheck: str = "Yes"
    attempt_count: int = 1



def finalize_result(result: Result, previous: Result | None = None) -> Result:
    result.resolver_version = RESOLVER_VERSION
    result.last_checked = result.updated_at or now_utc()
    result.updated_at = result.last_checked
    result.evidence_count = max(result.evidence_count, result.consensus_sources)
    result.needs_recheck = "No" if result.status == "Verified" and result.confidence >= 0.90 else "Yes"
    result.attempt_count = (previous.attempt_count if previous else 0) + 1
    return result


def status_rank(status: str) -> int:
    return {"Verified": 4, "Needs Review": 3, "Unresolved": 2, "Error": 1}.get(status, 0)


def choose_result(previous: Result | None, current: Result, force_replace: bool = False) -> Result:
    current = finalize_result(current, previous)
    if previous is None or force_replace:
        return current
    # Never downgrade a previously stronger answer. A new result replaces the old one
    # only when status improves, confidence improves materially, or evidence increases.
    better = (
        status_rank(current.status) > status_rank(previous.status)
        or (status_rank(current.status) == status_rank(previous.status) and current.confidence > previous.confidence + 0.01)
        or (current.county == previous.county and current.evidence_count > previous.evidence_count and current.confidence >= previous.confidence - 0.02)
    )
    if better:
        return current
    previous.last_checked = current.last_checked
    previous.attempt_count = current.attempt_count
    previous.resolver_version = RESOLVER_VERSION
    previous.needs_recheck = "No" if previous.status == "Verified" and previous.confidence >= 0.90 else "Yes"
    previous.notes = clean(previous.notes + " | Rechecked by " + RESOLVER_VERSION + "; existing stronger result retained.")
    return previous


def append_history(path: Path, previous: Result | None, attempted: Result, selected: Result, mode: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "checked_at", "mode", "resolver_version", "license_number", "brokerage_name",
        "previous_status", "previous_county", "previous_confidence",
        "attempted_status", "attempted_county", "attempted_confidence",
        "selected_status", "selected_county", "selected_confidence", "selection_note",
        "evidence_url", "secondary_evidence_url", "notes"
    ]
    row = {
        "checked_at": attempted.last_checked or attempted.updated_at or now_utc(),
        "mode": mode, "resolver_version": RESOLVER_VERSION,
        "license_number": attempted.license_number, "brokerage_name": attempted.brokerage_name,
        "previous_status": previous.status if previous else "",
        "previous_county": previous.county if previous else "",
        "previous_confidence": previous.confidence if previous else "",
        "attempted_status": attempted.status, "attempted_county": attempted.county,
        "attempted_confidence": attempted.confidence,
        "selected_status": selected.status, "selected_county": selected.county,
        "selected_confidence": selected.confidence,
        "selection_note": "new result selected" if selected is attempted else "existing stronger result retained",
        "evidence_url": attempted.evidence_url, "secondary_evidence_url": attempted.secondary_evidence_url,
        "notes": attempted.notes,
    }
    with _audit_lock:
        exists = path.exists()
        with path.open("a", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            if not exists:
                w.writeheader()
            w.writerow(row)


def registrable_host(url: str) -> str:
    host = urlparse(url).netloc.lower().split(":")[0].removeprefix("www.")
    return host


def is_listing_domain(url: str) -> bool:
    host = registrable_host(url)
    return any(host == d or host.endswith("." + d) for d in LISTING_DOMAINS)


def is_office_page(url: str, title: str = "", snippet: str = "") -> bool:
    path = urlparse(url).path.lower()
    sample = f"{title} {snippet}".lower()
    return any(term in path for term in OFFICE_PATH_TERMS) or bool(CONTACT_TERMS.search(sample))


def source_profile(url: str) -> tuple[int, str]:
    host = registrable_host(url)
    path = urlparse(url).path.lower()
    if any(p in path for p in PROPERTY_PATHS):
        return 0, "Property listing"
    if host.endswith("trec.texas.gov") or host.endswith("texas.gov"):
        return 5, "Government"
    if is_listing_domain(url):
        return 1, "Listing-site corroboration only"
    if any(x in host for x in ("bbb.org", "chamberofcommerce.com")):
        return 3, "Business directory"
    if any(x in host for x in ("bizapedia.com", "mapquest.com", "yelp.com", "facebook.com", "instagram.com", "linkedin.com")):
        return 1, "Low-authority directory"
    return 4, "Possible official website"


def extract_addresses(text: str) -> list[str]:
    text = html.unescape(re.sub(r"\s+", " ", text or " "))
    out: list[str] = []
    for m in ADDRESS_RE.findall(text):
        a = clean(m).strip(" ,.;")
        if a not in out:
            out.append(a)
    return out


def identity_score(name: str, broker_name: str, license_number: str, title: str, text: str, url: str) -> float:
    hay = clean(f"{title} {text} {url}").lower()
    name_score = token_set_ratio(normalized_name(name), hay) / 100
    broker_score = token_set_ratio(normalized_name(broker_name), hay) / 100 if broker_name else 0
    license_bonus = 0.35 if license_number.lower() in hay else 0
    exact_bonus = 0.20 if normalized_name(name) and normalized_name(name) in re.sub(r"[^a-z0-9 ]", " ", hay) else 0
    broker_bonus = 0.12 if broker_name and broker_score >= 0.90 else 0
    return min(1.0, 0.52 * name_score + 0.16 * broker_score + license_bonus + exact_bonus + broker_bonus)


def property_penalty(url: str, title: str, text: str, address: str) -> tuple[float, str]:
    path = urlparse(url).path.lower()
    sample = f"{title} {text}"
    if any(p in path for p in PROPERTY_PATHS):
        return 1.0, "property-listing URL"
    hits = len(PROPERTY_TERMS.findall(sample))
    if hits >= 3:
        return 0.85, "property-listing language"
    if hits >= 1 and address.lower() in sample.lower():
        return 0.45, "possible advertised property"
    return 0.0, ""


def address_context_score(title: str, text: str, source_tier: int) -> float:
    sample = f"{title} {text}"
    score = 0.30 + source_tier * 0.08
    if CONTACT_TERMS.search(sample):
        score += 0.18
    if ZIP_RE.search(sample):
        score += 0.08
    return min(1.0, score)


def read_search_cache(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_search_cache(path: Path, cache: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def serper_search(query: str, api_key: str, delay: float, cache: dict[str, Any], cache_path: Path, max_results: int) -> list[dict[str, str]]:
    key = hashlib.sha256(query.encode()).hexdigest()
    with _cache_lock:
        if key in cache:
            return cache[key]
    time.sleep(delay)
    r = session().post(
        "https://google.serper.dev/search",
        headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
        json={"q": query, "gl": "us", "hl": "en", "num": max_results},
        timeout=40,
    )
    r.raise_for_status()
    payload = r.json()
    rows = [{
        "title": clean(i.get("title")), "link": clean(i.get("link")), "snippet": clean(i.get("snippet"))
    } for i in payload.get("organic", [])]
    with _cache_lock:
        cache[key] = rows
        if len(cache) % 25 == 0:
            save_search_cache(cache_path, cache)
    return rows


def fetch_page(url: str, delay: float) -> tuple[str, str, list[str], list[str]]:
    try:
        time.sleep(delay)
        r = session().get(url, timeout=30, allow_redirects=True)
        if r.status_code >= 400 or "text/html" not in r.headers.get("content-type", ""):
            return "", "", [], []
        soup = BeautifulSoup(r.text[:2_000_000], "lxml")
        title = clean(soup.title.get_text(" ", strip=True) if soup.title else "")
        text = clean(soup.get_text(" ", strip=True))[:120000]
        addresses: list[str] = []
        office_links: list[str] = []
        base_host = registrable_host(r.url)
        for a in soup.find_all("a", href=True):
            label = clean(a.get_text(" ", strip=True)).lower()
            href = clean(a.get("href"))
            absolute = urljoin(r.url, href)
            if registrable_host(absolute) != base_host:
                continue
            path = urlparse(absolute).path.lower()
            if any(term in label or term in path for term in OFFICE_PATH_TERMS):
                if absolute not in office_links:
                    office_links.append(absolute)
        for node in soup.find_all("script", attrs={"type": "application/ld+json"}):
            try:
                obj = json.loads(node.get_text(" ", strip=True))
                stack = obj if isinstance(obj, list) else [obj]
                while stack:
                    item = stack.pop()
                    if isinstance(item, list):
                        stack.extend(item)
                    elif isinstance(item, dict):
                        stack.extend(item.values())
                        addr = item.get("address")
                        if isinstance(addr, dict):
                            line = ", ".join(clean(addr.get(k)) for k in ("streetAddress", "addressLocality", "addressRegion", "postalCode") if clean(addr.get(k)))
                            addresses.extend(extract_addresses(line))
            except Exception:
                pass
        addresses.extend(extract_addresses(text))
        return title, text, list(dict.fromkeys(addresses))[:15], office_links[:8]
    except Exception:
        return "", "", [], []


def census_geocode(address: str, delay: float) -> dict[str, str] | None:
    try:
        time.sleep(delay)
        r = session().get(
            "https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress",
            params={"address": address, "benchmark": "Public_AR_Current", "vintage": "Current_Current", "format": "json"},
            timeout=40,
        )
        r.raise_for_status()
        matches = r.json().get("result", {}).get("addressMatches", [])
        if not matches:
            return None
        m = matches[0]
        comps = m.get("addressComponents", {})
        counties = m.get("geographies", {}).get("Counties", [])
        county = clean(counties[0].get("NAME")) if counties else ""
        county = re.sub(r"\s+County$", "", county, flags=re.I)
        return {
            "matched_address": clean(m.get("matchedAddress")) or address,
            "city": clean(comps.get("city")), "state": clean(comps.get("state")),
            "zip_code": clean(comps.get("zip")), "county": county,
        }
    except Exception:
        return None


def build_candidate(lic: str, name: str, broker: str, address: str, url: str, title: str, text: str, source_type: str) -> Candidate:
    tier, profile = source_profile(url)
    ident = identity_score(name, broker, lic, title, text, url)
    neg, reason = property_penalty(url, title, text, address)
    addr_score = address_context_score(title, text, tier)
    total = max(0.0, min(1.0, 0.58 * ident + 0.27 * addr_score + 0.15 * (tier / 5) - 0.75 * neg))
    return Candidate(lic, name, broker, address, url, urlparse(url).netloc.lower(), tier, f"{profile}; {source_type}", title, clean(text)[:600], round(ident, 3), round(addr_score, 3), round(neg, 3), round(total, 3), reason)


def audit_candidates(path: Path, candidates: list[Candidate]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    names = [f.name for f in fields(Candidate)]
    with _audit_lock:
        exists = path.exists()
        with path.open("a", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=names)
            if not exists:
                w.writeheader()
            for c in candidates:
                w.writerow(asdict(c))


def resolve_one(lic: str, name: str, broker: str, api_key: str, cfg: dict[str, Any], cache: dict[str, Any], cache_path: Path, audit_path: Path) -> Result:
    delay = float(cfg["request_delay_seconds"])
    # Office-first search strategy. Listing domains are excluded from primary searches
    # and may only corroborate a location already found on an office/business source.
    negative_sites = "-site:zillow.com -site:redfin.com -site:realtor.com -site:homes.com -site:har.com -site:trulia.com -site:loopnet.com -site:crexi.com"
    queries = [
        f'"{name}" contact office address Texas {negative_sites}',
        f'"{name}" (contact OR office OR headquarters OR location) Texas {negative_sites}',
        f'"{name}" "{lic}" Texas {negative_sites}',
        f'"{name}" "{broker}" contact Texas {negative_sites}' if broker else f'"{name}" broker contact Texas {negative_sites}',
        f'"{name}" (BBB OR chamber OR business) Texas {negative_sites}',
    ]
    raw: list[Candidate] = []
    corroboration: list[Candidate] = []
    seen_pages: set[str] = set()
    page_budget = int(cfg["max_pages_per_record"])
    contact_budget = int(cfg.get("max_contact_pages_per_record", 4))

    for q in queries:
        try:
            rows = serper_search(q, api_key, delay, cache, cache_path, int(cfg["max_search_results"]))
        except Exception:
            continue
        for row in rows:
            url, title, snippet = row["link"], row["title"], row["snippet"]
            if not url:
                continue
            listing = is_listing_domain(url)
            target = corroboration if listing else raw
            for address in extract_addresses(f"{title} {snippet}"):
                target.append(build_candidate(lic, name, broker, address, url, title, snippet, "search snippet"))

            tier, _ = source_profile(url)
            preliminary_ident = identity_score(name, broker, lic, title, snippet, url)
            should_fetch = (not listing and (is_office_page(url, title, snippet) or tier >= 3) and preliminary_ident >= 0.32)
            if url not in seen_pages and page_budget > 0 and should_fetch:
                seen_pages.add(url)
                page_budget -= 1
                ptitle, ptext, addresses, office_links = fetch_page(url, delay)
                for address in addresses:
                    raw.append(build_candidate(lic, name, broker, address, url, ptitle or title, ptext or snippet, "office/business page"))

                # Follow same-domain Contact/About/Office pages discovered on an official site.
                for office_url in office_links:
                    if contact_budget <= 0 or office_url in seen_pages:
                        break
                    seen_pages.add(office_url)
                    contact_budget -= 1
                    ctitle, ctext, caddresses, _ = fetch_page(office_url, delay)
                    for address in caddresses:
                        raw.append(build_candidate(lic, name, broker, address, office_url, ctitle, ctext, "linked contact/office page"))

    dedup: dict[tuple[str, str], Candidate] = {}
    for c in raw + corroboration:
        key = (normalize_address(c.address), c.source_domain)
        if key not in dedup or c.total_score > dedup[key].total_score:
            dedup[key] = c
    candidates = sorted(dedup.values(), key=lambda c: c.total_score, reverse=True)

    minimum_identity = float(cfg["minimum_identity_score"])
    primary = [
        c for c in candidates
        if not is_listing_domain(c.url)
        and c.negative_score < 0.85
        and (
            (c.source_tier >= 2 and c.identity_score >= minimum_identity)
            or (c.source_tier >= 4 and c.identity_score >= max(0.30, minimum_identity - 0.16))
        )
    ]

    for c in primary[:20]:
        geo = census_geocode(c.address, delay)
        if geo and geo.get("county") and geo.get("state", "TX").upper() == "TX":
            c.matched_address, c.city, c.state, c.zip_code, c.county = geo["matched_address"], geo["city"], geo["state"], geo["zip_code"], geo["county"]

    # Listing sites can only add corroboration after a primary business address exists.
    geocoded_primary = [c for c in primary if c.county]
    for c in corroboration[:10]:
        if not geocoded_primary or c.negative_score >= 0.85:
            continue
        geo = census_geocode(c.address, delay)
        if geo and geo.get("county"):
            c.matched_address, c.city, c.state, c.zip_code, c.county = geo["matched_address"], geo["city"], geo["state"], geo["zip_code"], geo["county"]

    groups: dict[str, list[Candidate]] = defaultdict(list)
    for c in geocoded_primary:
        groups[normalize_address(c.matched_address or c.address)].append(c)
    ranked: list[tuple[float, list[Candidate]]] = []
    for key, group in groups.items():
        domains = {c.source_domain for c in group}
        best = max(c.total_score for c in group)
        authoritative = max(c.source_tier for c in group)
        consensus_bonus = min(0.16, 0.08 * (len(domains) - 1))
        # Corroboration only counts when it resolves to the same county as the primary group.
        county = group[0].county
        listing_support = len({c.source_domain for c in corroboration if c.county == county})
        corroboration_bonus = min(0.05, 0.02 * listing_support)
        final = min(0.99, best + consensus_bonus + corroboration_bonus + (0.05 if authoritative >= 4 else 0))
        ranked.append((final, sorted(group, key=lambda c: c.total_score, reverse=True)))
    ranked.sort(key=lambda x: x[0], reverse=True)

    audit_candidates(audit_path, candidates[:35])
    if not ranked:
        return Result(lic, name, broker, status="Unresolved", notes="No office/contact-page Texas business address could be verified and geocoded. Listing-site addresses were retained only as secondary evidence.", updated_at=now_utc())

    score, group = ranked[0]
    best = group[0]
    source_count = len({c.source_domain for c in group})
    has_authoritative = any(c.source_tier >= 4 for c in group)
    auto_accept = float(cfg["auto_accept_score"])
    review_score = float(cfg["review_score"])
    if score >= auto_accept and (has_authoritative or source_count >= 2):
        status = "Verified"
    elif score >= review_score:
        status = "Needs Review"
    else:
        status = "Unresolved"
    county = best.county if status != "Unresolved" else ""
    secondary = next((c.url for c in group[1:] if c.source_domain != best.source_domain), "")
    notes = f"Office/contact pages were prioritized. {source_count} independent primary source domain(s) support the selected business address; listing sites were secondary corroboration only."
    return Result(
        lic, name, broker, best.matched_address, best.city, best.state, best.zip_code, county,
        status, round(score, 3), best.identity_score, source_count, best.url, secondary,
        best.source_type, notes, now_utc(),
    )


def find_columns(ws) -> tuple[int, int, int, int | None]:
    headers = {clean(c.value).lower(): c.column for c in ws[1]}
    lic = headers.get("license number") or headers.get("license")
    name = headers.get("full name") or headers.get("brokerage") or headers.get("brokerage name")
    broker = headers.get("related license full name") or headers.get("related broker name")
    if not lic or not name:
        raise RuntimeError(f"Required columns not found: {list(headers)}")
    return 1, lic, name, broker


def load_checkpoint(path: Path) -> dict[str, Result]:
    if not path.exists():
        return {}
    out: dict[str, Result] = {}
    names = {f.name for f in fields(Result)}
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            data = {k: row.get(k, "") for k in names}
            for k in ("confidence", "identity_score"):
                data[k] = float(data.get(k) or 0)
            data["consensus_sources"] = int(float(data.get("consensus_sources") or 0))
            data["evidence_count"] = int(float(data.get("evidence_count") or data.get("consensus_sources") or 0))
            data["attempt_count"] = int(float(data.get("attempt_count") or 0))
            r = Result(**data)
            out[canonical_license(r.license_number)] = r
    return out


def save_checkpoint(path: Path, results: dict[str, Result]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    names = [f.name for f in fields(Result)]
    with tmp.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=names); w.writeheader()
        for key in sorted(results): w.writerow(asdict(results[key]))
    tmp.replace(path)


def write_output(input_file: Path, output_file: Path, checkpoint: dict[str, Result]) -> None:
    wb = load_workbook(input_file)
    ws = wb.active
    _, lic_col, _, _ = find_columns(ws)
    headers = ["Office Address", "City", "State", "ZIP", "County", "Resolution Status", "Confidence", "Identity Score", "Consensus Sources", "Evidence Type", "Evidence URL", "Secondary Evidence URL", "Resolution Notes", "Last Updated UTC", "Resolver Version", "Last Checked UTC", "Evidence Count", "Needs Recheck", "Attempt Count"]
    existing = {clean(c.value): c.column for c in ws[1]}
    col = ws.max_column + 1
    for h in headers:
        if h not in existing:
            ws.cell(1, col, h); existing[h] = col; col += 1
    mapping = {
        "Office Address":"office_address", "City":"city", "State":"state", "ZIP":"zip_code", "County":"county",
        "Resolution Status":"status", "Confidence":"confidence", "Identity Score":"identity_score",
        "Consensus Sources":"consensus_sources", "Evidence Type":"evidence_type", "Evidence URL":"evidence_url",
        "Secondary Evidence URL":"secondary_evidence_url", "Resolution Notes":"notes", "Last Updated UTC":"updated_at",
        "Resolver Version":"resolver_version", "Last Checked UTC":"last_checked", "Evidence Count":"evidence_count",
        "Needs Recheck":"needs_recheck", "Attempt Count":"attempt_count",
    }
    for row in range(2, ws.max_row + 1):
        r = checkpoint.get(canonical_license(ws.cell(row, lic_col).value))
        if r:
            for h, attr in mapping.items(): ws.cell(row, existing[h], getattr(r, attr))
    fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = fill; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    widths = {"Office Address":38,"City":18,"State":9,"ZIP":12,"County":18,"Resolution Status":18,"Confidence":12,"Identity Score":14,"Consensus Sources":18,"Evidence Type":26,"Evidence URL":52,"Secondary Evidence URL":52,"Resolution Notes":55,"Last Updated UTC":22,"Resolver Version":16,"Last Checked UTC":22,"Evidence Count":15,"Needs Recheck":15,"Attempt Count":14}
    for h, width in widths.items(): ws.column_dimensions[get_column_letter(existing[h])].width = width
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    output_file.parent.mkdir(parents=True, exist_ok=True)
    temp_output = output_file.with_name(output_file.stem + ".checkpoint" + output_file.suffix)
    wb.save(temp_output)
    temp_output.replace(output_file)


def version_number(value: str) -> int:
    m = re.search(r"(\d+)", clean(value))
    return int(m.group(1)) if m else 0


def age_days(timestamp: str) -> int:
    try:
        dt = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
        return max(0, (datetime.now(timezone.utc) - dt).days)
    except Exception:
        return 999999


def should_process(old: Result | None, mode: str, cfg: dict[str, Any]) -> bool:
    if old is None:
        return True
    threshold = float(cfg.get("low_confidence_threshold", 0.90))
    stale_days = int(cfg.get("recheck_after_days", 180))
    if mode == "recheck_all":
        return True
    if mode in {"recheck_review", "recheck_unresolved"}:
        return old.status in ({"Needs Review", "Unresolved"} if mode == "recheck_review" else {"Unresolved", "Error"})
    if mode == "upgrade_confidence":
        return old.confidence < threshold or old.status != "Verified"
    if mode == "upgrade_version":
        return version_number(old.resolver_version) < version_number(RESOLVER_VERSION)
    if mode == "recheck_stale":
        return age_days(old.last_checked or old.updated_at) >= stale_days
    if mode == "flagged":
        return clean(old.needs_recheck).lower() in {"yes", "true", "1"}
    return False


def save_progress_metadata(path: Path, checkpoint: dict[str, Result], completed_this_run: int, planned_this_run: int, mode: str) -> None:
    counts: dict[str, int] = defaultdict(int)
    for result in checkpoint.values():
        counts[result.status] += 1
    payload = {
        "last_saved_utc": now_utc(),
        "completed_this_run": completed_this_run,
        "planned_this_run": planned_this_run,
        "total_checkpointed_records": len(checkpoint),
        "mode": mode,
        "resolver_version": RESOLVER_VERSION,
        "status_counts": dict(counts),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    temp.replace(path)


def persist_all(
    checkpoint_path: Path,
    metadata_path: Path,
    cache_path: Path,
    input_file: Path,
    output_file: Path,
    checkpoint: dict[str, Result],
    cache: dict[str, Any],
    completed_this_run: int,
    planned_this_run: int,
    mode: str,
) -> None:
    """Atomically persist CSV state, cache, metadata, and the current workbook."""
    save_checkpoint(checkpoint_path, checkpoint)
    save_search_cache(cache_path, cache)
    save_progress_metadata(metadata_path, checkpoint, completed_this_run, planned_this_run, mode)
    write_output(input_file, output_file, checkpoint)
    print(f"Saved recovery checkpoint: {completed_this_run}/{planned_this_run} completed this run; {len(checkpoint)} total records saved.", flush=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--max-rows", type=int)
    ap.add_argument("--mode", choices=["new", "recheck_review", "recheck_unresolved", "upgrade_confidence", "upgrade_version", "recheck_stale", "flagged", "recheck_all"], default="new")
    ap.add_argument("--checkpoint-every", type=int, default=None)
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    api_key = os.environ.get("SERPER_API_KEY", "").strip()
    if not api_key:
        raise SystemExit("SERPER_API_KEY is missing.")

    input_file = Path(cfg["input_file"])
    output_file = Path(cfg["output_file"])
    checkpoint_path = Path(cfg["checkpoint_file"])
    audit_path = Path(cfg["candidate_audit_file"])
    cache_path = Path(cfg["search_cache_file"])
    metadata_path = Path(cfg.get("checkpoint_metadata_file", "state/checkpoint.json"))
    history_path = Path(cfg.get("history_file", "state/resolution_history.csv"))
    checkpoint_every = max(1, args.checkpoint_every or int(cfg.get("checkpoint_every", 25)))

    checkpoint = load_checkpoint(checkpoint_path)
    cache = read_search_cache(cache_path)

    wb = load_workbook(input_file, read_only=True, data_only=True)
    ws = wb.active
    _, lic_col, name_col, broker_col = find_columns(ws)
    pending: list[tuple[str, str, str]] = []
    limit = args.max_rows or int(cfg["max_rows_per_run"])
    for row in ws.iter_rows(min_row=2, values_only=True):
        lic = canonical_license(row[lic_col - 1])
        name = clean(row[name_col - 1])
        broker = clean(row[broker_col - 1]) if broker_col else ""
        if lic and name and should_process(checkpoint.get(lic), args.mode, cfg):
            pending.append((lic, name, broker))
            if len(pending) >= limit:
                break
    wb.close()

    planned = len(pending)
    completed = 0
    interrupted = False

    def handle_stop(signum, frame):
        nonlocal interrupted
        interrupted = True
        print(f"Received stop signal {signum}. Finishing current completed records and saving progress.", flush=True)

    signal.signal(signal.SIGTERM, handle_stop)
    signal.signal(signal.SIGINT, handle_stop)

    print(
        f"Checkpoint contains {len(checkpoint)} records; mode={args.mode}; processing={planned}; save interval={checkpoint_every}.",
        flush=True,
    )

    # Create a valid recovery workbook and metadata even when there is nothing new to process.
    if not pending:
        persist_all(
            checkpoint_path, metadata_path, cache_path, input_file, output_file,
            checkpoint, cache, completed, planned, args.mode,
        )
        print("No eligible records remain for this mode.", flush=True)
        return 0

    with ThreadPoolExecutor(max_workers=int(cfg["workers"])) as pool:
        futures = {
            pool.submit(resolve_one, lic, name, broker, api_key, cfg, cache, cache_path, audit_path): (lic, name, broker)
            for lic, name, broker in pending
        }
        try:
            for future in as_completed(futures):
                lic, name, broker = futures[future]
                try:
                    result = future.result()
                except Exception as exc:
                    result = Result(
                        lic,
                        name,
                        broker,
                        status="Error",
                        notes=f"{type(exc).__name__}: {exc}",
                        updated_at=now_utc(),
                    )
                previous = checkpoint.get(result.license_number)
                attempted = finalize_result(result, previous)
                selected = choose_result(previous, attempted, force_replace=False)
                checkpoint[result.license_number] = selected
                append_history(history_path, previous, attempted, selected, args.mode)
                completed += 1

                if completed % checkpoint_every == 0 or completed == planned:
                    persist_all(
                        checkpoint_path, metadata_path, cache_path, input_file, output_file,
                        checkpoint, cache, completed, planned, args.mode,
                    )

                if interrupted:
                    for pending_future in futures:
                        pending_future.cancel()
                    break
        finally:
            # This final save covers ordinary exceptions and manual cancellation signals.
            persist_all(
                checkpoint_path, metadata_path, cache_path, input_file, output_file,
                checkpoint, cache, completed, planned, args.mode,
            )

    counts: dict[str, int] = defaultdict(int)
    for result in checkpoint.values():
        counts[result.status] += 1
    print("Done:", dict(counts), flush=True)
    return 130 if interrupted else 0

if __name__ == "__main__":
    raise SystemExit(main())
