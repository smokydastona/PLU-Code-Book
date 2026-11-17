"""Inspect external image-source pages without downloading or approving images."""
import argparse
import csv
import json
import re
import sys
from html.parser import HTMLParser
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import HTTPRedirectHandler, Request, build_opener

from network_fetch import is_dns_error, open_with_dns_retry
from working_list import read_working_list

ROOT = Path(__file__).resolve().parents[1]
MAX_PAGE_BYTES = 1_000_000
ALLOWED_HOSTS = {
    "www.istockphoto.com", "pixabay.com", "depositphotos.com", "www.google.com",
    "www.shutterstock.com", "atgroupeurope.eu", "www.123rf.com", "www.vecteezy.com",
    "www.pexels.com", "static1.squarespace.com", "www.pickpik.com",
    "photodune.net", "www.pixtastock.com",
}


def validate_url(value):
    parsed = urlparse(value)
    if (parsed.scheme != "https" or parsed.hostname not in ALLOWED_HOSTS
            or parsed.username or parsed.password or parsed.port not in (None, 443)):
        raise ValueError(f"URL outside approved HTTPS hosts: {value}")


def site_host(value):
    host = urlparse(value).hostname or ""
    return host.removeprefix("www.")


class SafeRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        validate_url(newurl)
        if site_host(req.full_url) != site_host(newurl):
            raise ValueError("cross-site redirect blocked")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


class PageMetadata(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.title = ""
        self.in_title = False
        self.metadata = {}

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag.lower() == "title":
            self.in_title = True
        if tag.lower() == "meta":
            key = (attrs.get("property") or attrs.get("name") or attrs.get("itemprop") or "").lower()
            content = attrs.get("content", "")
            if key and content and key in {
                "description", "license", "copyright", "rights",
                "og:title", "og:description", "og:image", "og:image:width",
                "og:image:height", "og:url", "twitter:title", "twitter:description",
            }:
                self.metadata.setdefault(key, content.strip())
        if tag.lower() == "link" and "canonical" in attrs.get("rel", "").lower():
            self.metadata.setdefault("canonical", attrs.get("href", ""))
        if tag.lower() == "link" and "license" in attrs.get("rel", "").lower():
            self.metadata.setdefault("license_url", attrs.get("href", ""))

    def handle_endtag(self, tag):
        if tag.lower() == "title":
            self.in_title = False

    def handle_data(self, data):
        if self.in_title and not self.title:
            self.title = " ".join(data.split())[:300]


def read_leads(path):
    rows = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(rows, list) or not rows:
        raise ValueError("manifest must be a nonempty JSON list")
    seen = set()
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError("each lead must be a JSON object")
        for key in ("plu", "product", "source_url"):
            if not isinstance(row.get(key), str) or not row[key].strip():
                raise ValueError(f"each lead requires a nonempty {key}")
        if not re.fullmatch(r"\d+", row["plu"]) or row["plu"] in seen:
            raise ValueError(f"invalid or duplicate PLU: {row['plu']}")
        seen.add(row["plu"])
        validate_url(row["source_url"])
        if type(row.get("search_only", False)) is not bool:
            raise ValueError(f"search_only must be boolean for PLU {row['plu']}")
    return rows


def fetch_page(url):
    opener = build_opener(SafeRedirect())
    request = Request(url, headers={
        "User-Agent": "ProducePLUReview/1.0 (https://github.com/smokydastona/produce-plu-code-book)",
        "Accept": "text/html,application/xhtml+xml,application/pdf;q=0.9",
    })
    with open_with_dns_retry(opener.open, request, timeout=30) as response:
        content_type = response.headers.get("Content-Type", "").split(";", 1)[0].lower()
        result = {
            "http_status": response.status,
            "final_url": response.url,
            "content_type": content_type,
            "content_length": response.headers.get("Content-Length", ""),
        }
        if content_type == "application/pdf":
            result.update(status="PDF_PAGE_ONLY",
                          details="PDF identified by response headers; body not downloaded or parsed.")
            return result
        if content_type not in {"text/html", "application/xhtml+xml"}:
            result.update(status="NON_HTML_RESPONSE",
                          details="Response was not HTML; body not downloaded or parsed.")
            return result
        body = response.read(MAX_PAGE_BYTES + 1)
    truncated = len(body) > MAX_PAGE_BYTES
    parser = PageMetadata()
    parser.feed(body[:MAX_PAGE_BYTES].decode("utf-8", errors="replace"))
    metadata = parser.metadata
    result.update(
        status="PAGE_METADATA_FETCHED",
        title=parser.title or metadata.get("og:title", ""),
        description=metadata.get("description", metadata.get("og:description", ""))[:1000],
        canonical_url=metadata.get("canonical", metadata.get("og:url", "")),
        preview_url=metadata.get("og:image", ""),
        preview_width=metadata.get("og:image:width", ""),
        preview_height=metadata.get("og:image:height", ""),
        license_claim=metadata.get("license", metadata.get("copyright", metadata.get("rights", "")))[:500],
        license_url=metadata.get("license_url", ""),
        details=("HTML metadata only; preview image not downloaded. "
                 + ("HTML response truncated at 1 MB." if truncated else "")),
    )
    return result


def examine(manifest, output, root=ROOT):
    if output.exists() and any(output.iterdir()):
        raise ValueError("output must be empty")
    output.mkdir(parents=True, exist_ok=True)
    leads = read_leads(manifest)
    items, picklist_digest = read_working_list(root)
    active = {(item["plu"], item["name"]): item
              for item in items if item["included"]}
    results = []
    for lead in leads:
        item = active.get((lead["plu"], lead["product"]))
        row = dict(
            plu=lead["plu"], product=lead["product"], source_url=lead["source_url"],
            working_list_id=item["source_id"] if item else "",
            working_source_ids=";".join(item.get("source_ids", [])) if item else "",
            working_record_sha256=item.get("record_sha256", "") if item else "",
            picklist_sha256=picklist_digest,
            status="IDENTITY_NOT_FOUND" if not item else "",
            title="", description="", canonical_url="", preview_url="",
            preview_width="", preview_height="", content_type="", content_length="",
            license_claim="", license_url="", http_status="", final_url="",
            license_verification_status="NOT_VERIFIED",
            rights_verification_status="NOT_VERIFIED",
            photographic_origin_verification_status="NOT_VERIFIED",
            visual_approval_status="NOT_APPROVED",
            resolution_verification_status="NOT_VERIFIED",
            production_status="NOT_APPROVED",
            details="No exact active WorkingList identity." if not item else "",
        )
        if item:
            try:
                if lead.get("search_only"):
                    row.update(status="SEARCH_URL_ONLY",
                               details="Search page is not a specific image source; no search results fetched.")
                else:
                    row.update(fetch_page(lead["source_url"]))
                    row["license_verification_status"] = "NOT_VERIFIED"
                    row["rights_verification_status"] = "NOT_VERIFIED"
                    row["photographic_origin_verification_status"] = "NOT_VERIFIED"
                    row["visual_approval_status"] = "NOT_APPROVED"
                    row["resolution_verification_status"] = "NOT_VERIFIED"
                    row["production_status"] = "NOT_APPROVED"
            except (HTTPError, URLError, OSError, ValueError) as exc:
                row.update(status="DNS_ERROR" if is_dns_error(exc) else "FETCH_ERROR",
                            details=str(exc)[:500])
        results.append(row)
        print(f"{row['plu']} {row['product']}: {row['status']}", flush=True)

    fields = list(results[0])
    with (output / "source-findings.csv").open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(results)
    (output / "source-findings.json").write_text(
        json.dumps(results, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (output / "REVIEW.txt").write_text(
        "These are external source-page leads, not approved images. The inspector fetched "
        "HTML metadata only; it did not download previews, images, or PDF bodies. Page text "
        "and declared licenses are unverified claims. Rights verification, photographic "
        "origin, visual review, image resolution, reuse decisions, and production assignment "
        "remain outstanding. Search URLs were not queried.\n", encoding="utf-8")
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    try:
        examine(args.manifest, args.output, args.root)
    except (OSError, ValueError, KeyError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
