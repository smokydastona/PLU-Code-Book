"""Examine owner-supplied Commons leads without granting image approvals."""
import argparse
import csv
import json
import re
import sys
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlencode, urlsplit

from retrieve_commons import API, contact_sheets, download, fetch, plain, search

CATEGORIES = (
    "Category:Fruits on white background",
    "Category:Vegetables on white background",
)


def category_pages(category, user_agent):
    params = {
        "action": "query", "format": "json", "generator": "categorymembers",
        "gcmtitle": category, "gcmnamespace": 6, "gcmtype": "file",
        "gcmlimit": 100, "prop": "imageinfo", "iiprop": "url|size|mime|extmetadata",
    }
    response = json.loads(fetch(API + "?" + urlencode(params), user_agent))
    if "error" in response:
        raise ValueError(f"Commons API error: {response['error']}")
    return list(response.get("query", {}).get("pages", {}).values()), response.get("continue")


def page_details(page):
    info = page.get("imageinfo", [{}])[0]
    metadata = info.get("extmetadata", {})
    get = lambda key: plain(metadata.get(key, {}).get("value", ""))
    return dict(
        title=page.get("title", ""), page_url=info.get("descriptionurl", ""),
        original_url=info.get("url", ""), width=info.get("width", 0),
        height=info.get("height", 0), mime=info.get("mime", ""),
        creator=get("Artist"), license=get("LicenseShortName"),
        license_url=get("LicenseUrl"), restrictions=get("Restrictions"),
        description=get("ImageDescription"),
    )


def examine(leads, output, user_agent):
    if output.exists() and any(output.iterdir()):
        raise ValueError("output must be empty")
    output.mkdir(parents=True, exist_ok=True)
    with leads.open(encoding="utf-8-sig", newline="") as file:
        rows = list(csv.DictReader(file))
    if not rows or len({row["plu"] for row in rows}) != len(rows):
        raise ValueError("leads must have distinct PLUs")
    categories, category_errors = {}, {}
    for category in CATEGORIES:
        try:
            pages, continuation = category_pages(category, user_agent)
            categories[category] = dict(
                pages=[page_details(page) for page in pages],
                continuation=continuation,
                scope="First 100 direct file members only; subcategories not traversed",
            )
        except (OSError, ValueError, KeyError) as exc:
            category_errors[category] = str(exc)
    (output / "categories.json").write_text(json.dumps(
        dict(categories=categories, errors=category_errors),
        ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    results, records, queries = [], [], {}
    for number, row in enumerate(rows, 1):
        parsed = urlsplit(row["source_url"])
        if parsed.scheme != "https" or parsed.hostname != "commons.wikimedia.org":
            raise ValueError(f"invalid Commons URL for {row['plu']}")
        result = dict(plu=row["plu"], product=row["product"], source_url=row["source_url"],
                      status="UNRESOLVED", details="", candidates=[])
        try:
            if row["source_type"] == "FILE_PAGE" and parsed.path.startswith("/wiki/File:"):
                title = unquote(parsed.path.removeprefix("/wiki/")).replace("_", " ")
                pages = search(row["product"], user_agent, title=title)
            elif row["source_type"] == "SEARCH_QUERY" and parsed.path == "/w/index.php":
                query = parse_qs(parsed.query).get("search", [""])[0]
                if not query:
                    raise ValueError("missing Commons search query")
                if query not in queries:
                    queries[query] = search(query, user_agent)
                pages = queries[query]
            else:
                raise ValueError("unsupported Commons lead")
            tokens = [t for t in re.findall(r"[a-z]+", row["product"].lower())
                      if len(t) >= 5 and t not in {"organic", "small", "seedless"}]
            for category in categories.values():
                for page in category["pages"]:
                    text = (page["title"] + " " + page["description"]).lower()
                    if tokens and all(token.rstrip("s") in text for token in tokens):
                        result.setdefault("category_matches", []).append(page)
            result["pages_returned"] = len(pages)
            result["metadata"] = [page_details(page) for page in pages[:10]]
            folder = output / "originals" / row["plu"]
            rejections = []
            if download(row["plu"], row["product"], folder, user_agent,
                        pages=pages[:10], used_images={}, rejections=rejections):
                sidecar = next(folder.glob("*.json"))
                metadata = json.loads(sidecar.read_text(encoding="utf-8"))
                metadata.update(idx=str(number), image=str(sidecar.with_suffix("").relative_to(output)))
                records.append(metadata)
                result.update(status="CANDIDATE_NOT_APPROVED",
                              details="Candidate requires visual, rights, variety and reuse review")
            else:
                result["details"] = "No eligible downloadable candidate in first ten results"
            if rejections:
                result["details"] += "; " + "; ".join(rejections)
        except (OSError, ValueError, KeyError) as exc:
            result.update(status="ERROR", details=str(exc))
        results.append(result)
    (output / "lead-results.json").write_text(json.dumps(
        results, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    with (output / "lead-results.csv").open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=["plu", "product", "source_url", "status", "details"])
        writer.writeheader()
        writer.writerows({key: row[key] for key in writer.fieldnames} for row in results)
    contact_sheets(records, output)
    (output / "REVIEW.txt").write_text(
        "CANDIDATES ARE UNREVIEWED AND NOT APPROVED. Search results and category membership "
        "do not verify exact variety, rights, photographic origin or image reuse. "
        "No WorkingList assignment or production gate was changed. "
        "Category results cover the first 100 direct file members only.\n", encoding="utf-8")
    print(json.dumps(dict(leads=len(results), candidates=len(records),
                          errors=sum(row["status"] == "ERROR" for row in results),
                          category_errors=category_errors)))
    return bool(category_errors or any(row["status"] == "ERROR" for row in results))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--leads", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--user-agent", required=True)
    args = parser.parse_args()
    try:
        return int(examine(args.leads, args.output, args.user_agent))
    except (OSError, ValueError, KeyError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
