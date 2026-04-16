#!/usr/bin/env python3
import argparse
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


def normalize_title(text: str) -> str:
    text = (text or "").strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_title(a), normalize_title(b)).ratio()


def make_session(email: str) -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=1.0,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(
        {
            "User-Agent": f"title-doi-matcher/1.0 (mailto:{email})",
            "Accept": "application/json",
        }
    )
    return session


def crossref_lookup(session: requests.Session, title: str) -> Tuple[Optional[str], Optional[str], float, str]:
    url = "https://api.crossref.org/works"
    params = {
        "query.title": title,
        "rows": 5,
        "select": "DOI,title,type,issued,published,container-title,author",
    }
    response = session.get(url, params=params, timeout=30)
    response.raise_for_status()
    items = response.json().get("message", {}).get("items", [])

    best_doi = None
    best_title = None
    best_score = -1.0
    best_type = ""
    for item in items:
        cand_title = ""
        titles = item.get("title") or []
        if titles:
            cand_title = titles[0]
        score = similarity(title, cand_title)
        if score > best_score:
            best_score = score
            best_doi = item.get("DOI")
            best_title = cand_title
            best_type = item.get("type") or ""
    return best_doi, best_title, best_score, best_type


def openalex_lookup(session: requests.Session, title: str) -> Tuple[Optional[str], Optional[str], float, str]:
    url = "https://api.openalex.org/works"
    params = {
        "search": title,
        "per-page": 5,
        "mailto": session.headers.get("User-Agent", ""),
    }
    response = session.get(url, params=params, timeout=30)
    response.raise_for_status()
    items = response.json().get("results", [])

    best_doi = None
    best_title = None
    best_score = -1.0
    best_type = ""
    for item in items:
        cand_title = item.get("title") or ""
        score = similarity(title, cand_title)
        if score > best_score:
            best_score = score
            doi = item.get("doi")
            if doi and doi.startswith("https://doi.org/"):
                doi = doi.replace("https://doi.org/", "")
            best_doi = doi
            best_title = cand_title
            best_type = item.get("type") or ""
    return best_doi, best_title, best_score, best_type


def lookup_one(
    session: requests.Session,
    title: str,
    output_title_column: str,
    min_score: float,
    delay: float,
) -> Dict[str, Any]:
    title = (title or "").strip()
    out: Dict[str, Any] = {
        output_title_column: title,
        "doi": "",
    }
    if not title:
        return out

    try:
        doi, _matched_title, score, _record_type = crossref_lookup(session, title)
        if (not doi) or (score < min_score):
            doi2, _matched_title2, score2, _record_type2 = openalex_lookup(session, title)
            if doi2 and score2 >= max(score, min_score):
                doi, score = doi2, score2
        if doi and score >= min_score:
            out["doi"] = doi
    except Exception:
        pass

    if delay > 0:
        time.sleep(delay)
    return out


def autosize_worksheet(ws) -> None:
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 80)


def write_output(df: pd.DataFrame, output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.book["Sheet1"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        autosize_worksheet(ws)


def main() -> None:
    parser = argparse.ArgumentParser(description="Match paper titles to DOI and write a new Excel file.")
    parser.add_argument("--input", required=True, help="Input .xlsx file")
    parser.add_argument("--output", required=True, help="Output .xlsx file")
    parser.add_argument(
        "--title-column",
        default=None,
        help="Column name for the title field. Defaults to the first column.",
    )
    parser.add_argument(
        "--output-title-column",
        default="title",
        help="Column name to use for the title field in the output workbook.",
    )
    parser.add_argument("--email", required=True, help="Contact email for API polite usage")
    parser.add_argument("--workers", type=int, default=4, help="Number of concurrent lookup workers")
    parser.add_argument("--min-score", type=float, default=0.90, help="Minimum title similarity threshold")
    parser.add_argument("--delay", type=float, default=0.2, help="Sleep after each lookup to avoid hammering APIs")
    parser.add_argument("--limit", type=int, default=0, help="Optional row limit for testing")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    df = pd.read_excel(input_path)
    if df.empty:
        raise ValueError("Input workbook has no rows")

    title_col = args.title_column or df.columns[0]
    if title_col not in df.columns:
        raise ValueError(f"Title column not found: {title_col}")

    titles: List[str] = df[title_col].fillna("").astype(str).tolist()
    if args.limit and args.limit > 0:
        titles = titles[: args.limit]

    session = make_session(args.email)
    results: List[Optional[Dict[str, Any]]] = [None] * len(titles)

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_to_idx = {
            executor.submit(
                lookup_one,
                session,
                title,
                args.output_title_column,
                args.min_score,
                args.delay,
            ): idx
            for idx, title in enumerate(titles)
        }
        completed = 0
        total = len(future_to_idx)
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            results[idx] = future.result()
            completed += 1
            if completed % 50 == 0 or completed == total:
                print(f"Processed {completed}/{total}")

    out_df = pd.DataFrame(results)
    out_df = out_df[[args.output_title_column, "doi"]]
    write_output(out_df, output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
