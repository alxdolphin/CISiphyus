from __future__ import annotations

import re
from typing import Any
from urllib.parse import parse_qsl, urlsplit

import config


def is_excel_export_url(candidate: str) -> bool:
    return "/caseworthy/excelexport.aspx" in candidate.lower()


def discover_export_url_candidates(page: Any) -> list[str]:
    candidates: list[str] = []
    seen: set[str] = set()

    def add_candidate(candidate: str) -> None:
        normalized = candidate.strip()
        if not normalized:
            return
        if not normalized.startswith(("http://", "https://")):
            return
        if not is_excel_export_url(normalized):
            return
        if normalized in seen:
            return
        seen.add(normalized)
        candidates.append(normalized)

    hrefs = page.eval_on_selector_all("a[href]", "els => els.map(el => el.href)")
    if isinstance(hrefs, list):
        for href in hrefs:
            if isinstance(href, str):
                add_candidate(href)

    storage_values = page.evaluate(
        """() => {
            const values = [];
            for (let i = 0; i < window.localStorage.length; i += 1) {
                const key = window.localStorage.key(i);
                if (!key) {
                    continue;
                }
                values.push(key);
                const value = window.localStorage.getItem(key);
                if (value) {
                    values.push(value);
                }
            }
            for (let i = 0; i < window.sessionStorage.length; i += 1) {
                const key = window.sessionStorage.key(i);
                if (!key) {
                    continue;
                }
                values.push(key);
                const value = window.sessionStorage.getItem(key);
                if (value) {
                    values.push(value);
                }
            }
            return values;
        }"""
    )

    if isinstance(storage_values, list):
        url_pattern = re.compile(r"https?://[^\s\"'<>]+", flags=re.IGNORECASE)
        for value in storage_values:
            if not isinstance(value, str):
                continue
            if is_excel_export_url(value):
                add_candidate(value)
            for match in url_pattern.findall(value):
                add_candidate(match)

    return candidates


def choose_refreshed_export_url(
    original_url: str,
    candidates: list[str],
) -> tuple[str, dict[str, Any]]:
    original_split = urlsplit(original_url)
    original_params = dict(parse_qsl(original_split.query, keep_blank_values=True))
    original_form_id = original_params.get("FormID")
    original_host = original_split.netloc.lower()

    scored_candidates: list[tuple[int, str]] = []
    for candidate in candidates:
        candidate_split = urlsplit(candidate)
        candidate_params = dict(parse_qsl(candidate_split.query, keep_blank_values=True))
        score = 0
        if candidate_split.netloc.lower() == original_host:
            score += 2
        if original_form_id and candidate_params.get("FormID") == original_form_id:
            score += 2
        if candidate != original_url:
            score += 1
        scored_candidates.append((score, candidate))

    scored_candidates.sort(key=lambda item: item[0], reverse=True)
    selected_url = original_url
    if scored_candidates and scored_candidates[0][0] > 0:
        selected_url = scored_candidates[0][1]

    metadata = {
        "export_url_candidates_found_count": len(candidates),
        "export_url_candidates_preview": [config.redact_url(item) for item in candidates[:5]],
        "redacted_export_url_selected": config.redact_url(selected_url),
        "export_url_refreshed": selected_url != original_url,
    }
    return selected_url, metadata
