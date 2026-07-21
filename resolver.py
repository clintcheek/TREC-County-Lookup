from __future__ import annotations

import argparse
import csv
import html
import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

USER_AGENT = "TexasBrokerCountyResolver/1.0 (+private research workflow)"
ADDRESS_RE = re.compile(
    r"\b(\d{1,6}\s+[A-Za-z0-9.'#\- ]{2,70}?\s(?:Street|St|Road|Rd|Avenue|Ave|Boulevard|Blvd|Drive|Dr|Lane|Ln|Court|Ct|Parkway|Pkwy|Highway|Hwy|Loop|Way|Trail|Trl|Circle|Cir|Plaza|Pl|Terrace|Ter)\.?"
    r"(?:\s*(?:Suite|Ste|Unit|#)\s*[A-Za-z0-9\-]+)?\s*,?\s*[A-Za-z.'\- ]{2,40},?\s*(?:TX|Texas|[A-Z]{2})\s+\d{5}(?:-\d{4})?)\b",
    re.IGNORECASE,
)
ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")

_thread = threading.local()


def session() -> requests.Session:
    if not hasattr(_thread, "session"):
        s = requests.Session()
        s.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"})
        _thread.session = s
    return _thread.session


@dataclass
class Result:
    license_number: str
    brokerage_name: str
    office_address: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""
    county: str = ""
    status: str = "Unresolved"
    confidence: float = 0.0
    evidence_url: str = ""
    evidence_type: str = ""
    notes: str = ""
    updated_at: str = ""


def normalize(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def canonical_license(value: Any) -> str:
    return normalize(value).upper()


def search_serper(query: str, api_key: str, delay: float) -> list[dict[str, str]]:
    time.sleep(delay)
    r = session().post(
        "https://google.serper.dev/search",
        headers={"X-API-KEY": api_key, "Content-Type": "application/json"},
        json={"q": query, "gl": "us", "hl": "en", "num": 10},
        timeout=35,
    )
    r.raise_for_status()
    payload = r.json()
    rows: list[dict[str, str]] = []
    for item in payload.get("organic", []):
        rows.append({
            "title": normalize(item.get("title")),
            "link": normalize(item.get("link")),
            "snippet": normalize(item.get("snippet")),
        })
    return rows


def likely_official(url: str, name: str) -> bool:
    host = urlparse(url).netloc.lower().removeprefix("www.")
    blocked = ("facebook.com", "instagram.com", "linkedin.com", "yelp.com", "mapquest.com")
    if any(x in host for x in blocked):
        return False
    tokens = [t.lower() for t in re.findall(r"[A-Za-z0-9]+", name) if len(t) > 3]
    return bool(tokens) and any(t in host.replace("-", "") for t in tokens[:4])


def extract_addresses(text: str) -> list[str]:
    cleaned = html.unescape(re.sub(r"\s+", " ", text or " "))
    out: list[str] = []
    for match in ADDRESS_RE.findall(cleaned):
        address = normalize(match).strip(" ,.;")
        if address not in out:
            out.append(address)
    return out


def fetch_page_addresses(url: str, delay: float) -> list[str]:
    try:
        time.sleep(delay)
        r = session().get(url, timeout=25, allow_redirects=True)
        if r.status_code >= 400 or "text/html" not in r.headers.get("content-type", ""):
            return []
        soup = BeautifulSoup(r.text[:2_000_000], "lxml")
        candidates: list[str] = []
        for node in soup.find_all("script", attrs={"type": "application/ld+json"}):
            try:
                obj = json.loads(node.get_text(" ", strip=True))
                stack = obj if isinstance(obj, list) else [obj]
                for item in stack:
                    if isinstance(item, dict):
                        addr = item.get("address")
                        if isinstance(addr, dict):
                            line = ", ".join(normalize(addr.get(k)) for k in (
                                "streetAddress", "addressLocality", "addressRegion", "postalCode"
                            ) if normalize(addr.get(k)))
                            candidates.extend(extract_addresses(line))
            except Exception:
                pass
        candidates.extend(extract_addresses(soup.get_text(" ", strip=True)))
        return list(dict.fromkeys(candidates))[:10]
    except Exception:
        return []


def census_geocode(address: str, delay: float) -> dict[str, str] | None:
    try:
        time.sleep(delay)
        r = session().get(
            "https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress",
            params={
                "address": address,
                "benchmark": "Public_AR_Current",
                "vintage": "Current_Current",
                "format": "json",
            },
            timeout=35,
        )
        r.raise_for_status()
        matches = r.json().get("result", {}).get("addressMatches", [])
        if not matches:
            return None
        m = matches[0]
        comps = m.get("addressComponents", {})
        geos = m.get("geographies", {})
        counties = geos.get("Counties") or geos.get("County Subdivisions") or []
        county = normalize(counties[0].get("NAME")) if counties else ""
        county = re.sub(r"\s+County$", "", county, flags=re.I)
        return {
            "matched_address": normalize(m.get("matchedAddress")) or address,
            "city": normalize(comps.get("city")),
            "state": normalize(comps.get("state")),
            "zip_code": normalize(comps.get("zip")),
            "county": county,
        }
    except Exception:
        return None


def score_candidate(name: str, url: str, snippet: str, address: str, source: str) -> float:
    score = 0.35
    if likely_official(url, name):
        score += 0.30
    if source == "page":
        score += 0.12
    if "har.com" in url or "realtor.com" in url or "loopnet.com" in url:
        score += 0.12
    name_tokens = [t.lower() for t in re.findall(r"[A-Za-z0-9]+", name) if len(t) > 3]
    hay = (url + " " + snippet).lower()
    overlap = sum(1 for t in name_tokens[:6] if t in hay)
    score += min(0.12, overlap * 0.03)
    if not ZIP_RE.search(address):
        score -= 0.20
    return max(0.0, min(0.99, score))


def resolve_one(license_number: str, brokerage_name: str, api_key: str, delay: float, threshold: float) -> Result:
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    queries = [
        f'"{brokerage_name}" "{license_number}" address',
        f'"{brokerage_name}" Texas office address',
        f'"{brokerage_name}" broker address',
    ]
    seen_urls: set[str] = set()
    candidates: list[tuple[float, str, str, str, str]] = []
    errors: list[str] = []

    for query in queries:
        try:
            results = search_serper(query, api_key, delay)
        except Exception as exc:
            errors.append(f"search:{type(exc).__name__}")
            continue
        for item in results:
            url, snippet = item["link"], item["snippet"]
            if not url or url in seen_urls:
                continue
            seen_urls.add(url)
            for address in extract_addresses(item["title"] + " " + snippet):
                candidates.append((score_candidate(brokerage_name, url, snippet, address, "snippet"), address, url, "Search snippet", snippet))
            if likely_official(url, brokerage_name) or any(d in url for d in ("har.com", "realtor.com", "loopnet.com", "bbb.org")):
                for address in fetch_page_addresses(url, delay):
                    candidates.append((score_candidate(brokerage_name, url, snippet, address, "page"), address, url, "Web page", snippet))

    candidates.sort(key=lambda x: x[0], reverse=True)
    for score, address, url, evidence_type, snippet in candidates[:12]:
        geo = census_geocode(address, delay)
        if not geo or not geo.get("county"):
            continue
        status = "Resolved" if score >= threshold else "Needs Review"
        return Result(
            license_number=license_number,
            brokerage_name=brokerage_name,
            office_address=geo["matched_address"],
            city=geo["city"],
            state=geo["state"],
            zip_code=geo["zip_code"],
            county=geo["county"],
            status=status,
            confidence=round(score, 2),
            evidence_url=url,
            evidence_type=evidence_type,
            notes="Address extracted from public source and county returned by U.S. Census geocoder.",
            updated_at=now,
        )

    return Result(
        license_number=license_number,
        brokerage_name=brokerage_name,
        status="Unresolved",
        confidence=0.0,
        notes="No geocodable public office address found. " + ";".join(errors[:3]),
        updated_at=now,
    )


def find_columns(ws) -> tuple[int, int, int]:
    header_row = 1
    headers = {normalize(c.value).lower(): c.column for c in ws[header_row]}
    license_col = headers.get("license number") or headers.get("license")
    name_col = headers.get("full name") or headers.get("brokerage") or headers.get("brokerage name")
    if not license_col or not name_col:
        raise RuntimeError(f"Required columns not found. Headers: {list(headers)}")
    return header_row, license_col, name_col


def load_checkpoint(path: Path) -> dict[str, Result]:
    if not path.exists():
        return {}
    out: dict[str, Result] = {}
    with path.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            row["confidence"] = float(row.get("confidence") or 0)
            r = Result(**{k: row.get(k, "") for k in Result.__dataclass_fields__})
            out[canonical_license(r.license_number)] = r
    return out


def save_checkpoint(path: Path, results: dict[str, Result]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp")
    fields = list(Result.__dataclass_fields__)
    with temp.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for key in sorted(results):
            writer.writerow(asdict(results[key]))
    temp.replace(path)


def write_output(input_file: Path, output_file: Path, checkpoint: dict[str, Result]) -> None:
    wb = load_workbook(input_file)
    ws = wb.active
    _, license_col, _ = find_columns(ws)
    new_headers = [
        "Office Address", "City", "State", "ZIP", "County", "Resolution Status",
        "Confidence", "Evidence Type", "Evidence URL", "Resolution Notes", "Last Updated UTC"
    ]
    existing = {normalize(c.value): c.column for c in ws[1]}
    start_col = ws.max_column + 1
    for idx, header in enumerate(new_headers):
        if header not in existing:
            ws.cell(1, start_col, header)
            existing[header] = start_col
            start_col += 1
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    field_map = {
        "Office Address": "office_address", "City": "city", "State": "state", "ZIP": "zip_code",
        "County": "county", "Resolution Status": "status", "Confidence": "confidence",
        "Evidence Type": "evidence_type", "Evidence URL": "evidence_url",
        "Resolution Notes": "notes", "Last Updated UTC": "updated_at",
    }
    for row in range(2, ws.max_row + 1):
        key = canonical_license(ws.cell(row, license_col).value)
        result = checkpoint.get(key)
        if not result:
            continue
        for header, attr in field_map.items():
            ws.cell(row, existing[header], getattr(result, attr))
    widths = {"Office Address": 38, "City": 18, "State": 9, "ZIP": 12, "County": 18,
              "Resolution Status": 18, "Confidence": 12, "Evidence Type": 18,
              "Evidence URL": 55, "Resolution Notes": 48, "Last Updated UTC": 22}
    for header, width in widths.items():
        ws.column_dimensions[get_column_letter(existing[header])].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--max-rows", type=int, default=None)
    args = ap.parse_args()
    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    api_key = os.environ.get("SERPER_API_KEY", "").strip()
    if not api_key:
        raise SystemExit("SERPER_API_KEY is missing. Add it as a GitHub Actions repository secret.")
    input_file = Path(cfg["input_file"])
    output_file = Path(cfg["output_file"])
    checkpoint_file = Path(cfg["checkpoint_file"])
    max_rows = args.max_rows or int(cfg.get("max_rows_per_run", 500))
    workers = int(cfg.get("workers", 5))
    delay = float(cfg.get("request_delay_seconds", 0.45))
    threshold = float(cfg.get("minimum_confidence_to_resolve", 0.72))

    wb = load_workbook(input_file, read_only=True, data_only=True)
    ws = wb.active
    _, license_col, name_col = find_columns(ws)
    checkpoint = load_checkpoint(checkpoint_file)
    pending: list[tuple[str, str]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        lic = canonical_license(row[license_col - 1])
        name = normalize(row[name_col - 1])
        if not lic or not name:
            continue
        old = checkpoint.get(lic)
        if old and old.status in {"Resolved", "Needs Review", "Unresolved"}:
            continue
        pending.append((lic, name))
        if len(pending) >= max_rows:
            break
    wb.close()

    print(f"Loaded {len(checkpoint)} checkpoint rows; processing {len(pending)} new rows.")
    if pending:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = {pool.submit(resolve_one, lic, name, api_key, delay, threshold): lic for lic, name in pending}
            completed = 0
            for fut in as_completed(futures):
                result = fut.result()
                checkpoint[canonical_license(result.license_number)] = result
                completed += 1
                if completed % 10 == 0 or completed == len(pending):
                    save_checkpoint(checkpoint_file, checkpoint)
                    print(f"Checkpoint: {completed}/{len(pending)} in this run")
    save_checkpoint(checkpoint_file, checkpoint)
    write_output(input_file, output_file, checkpoint)
    resolved = sum(1 for r in checkpoint.values() if r.status == "Resolved")
    review = sum(1 for r in checkpoint.values() if r.status == "Needs Review")
    unresolved = sum(1 for r in checkpoint.values() if r.status == "Unresolved")
    print(f"Done. Resolved={resolved}, Needs Review={review}, Unresolved={unresolved}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
