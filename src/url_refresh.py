#CISiphyus URL Refresh Module

# This module refreshs the export URL for a given report by refreshing the LS parameter

from __future__ import annotations

import re
from typing import Any
from urllib.parse import parse_qsl, quote, unquote, urlsplit, urlunsplit

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


def decode_ls_pairs(ls_value: str) -> tuple[str, list[tuple[str, str]]]:
    raw = ls_value.strip()
    if not raw:
        raise ValueError("ls value empty")

    decoded = raw
    for _ in range(3):
        nxt = unquote(decoded)
        if nxt == decoded:
            break
        decoded = nxt

    if "procedure=" not in decoded.lower():
        raise ValueError("missing procedure prefix in ls")

    if "&" in decoded:
        procedure_part, rest = decoded.split("&", 1)
    else:
        procedure_part, rest = decoded, ""

    if not procedure_part.lower().startswith("procedure="):
        raise ValueError("invalid procedure prefix in ls")
    procedure = procedure_part.split("=", 1)[1]

    pairs: list[tuple[str, str]] = []
    for token in rest.split("&"):
        if not token:
            continue
        if "=" not in token:
            raise ValueError(f"invalid ls token: {token}")
        key, value = token.split("=", 1)
        pairs.append((key, value))
    return procedure, pairs


def _encode_ls_value(value: str) -> str:
    encoded = ""
    for ch in value:
        if ch == "~":
            encoded += "%257e"
        else:
            encoded += ch
    return encoded


def encode_ls_blob(procedure: str, pairs: list[tuple[str, str]]) -> str:
    segments = [f"Procedure%3D{procedure}"]
    for key, value in pairs:
        if not key.startswith("@"):
            raise ValueError(f"expected @-prefixed ls key, got {key!r}")
        segments.append(f"%2540{key[1:]}%3D{_encode_ls_value(value)}")
    return "%26".join(segments)


def replace_ls_param(url: str, param: str, value: str) -> str:
    split = urlsplit(url)
    query_pairs = list(parse_qsl(split.query, keep_blank_values=True))
    ls_index = next((idx for idx, (key, _) in enumerate(query_pairs) if key == "LS"), None)
    if ls_index is None:
        raise ValueError("url missing LS query parameter")

    _, ls_value = query_pairs[ls_index]
    procedure, pairs = decode_ls_pairs(ls_value)
    if not param.startswith("@"):
        raise ValueError(f"year scope param must start with @, got {param!r}")

    updated = False
    new_pairs: list[tuple[str, str]] = []
    for key, existing in pairs:
        if key == param:
            new_pairs.append((key, value))
            updated = True
        else:
            new_pairs.append((key, existing))
    if not updated:
        raise ValueError(f"ls param not found: {param}")

    query_pairs[ls_index] = ("LS", encode_ls_blob(procedure, new_pairs))
    encoded_parts: list[str] = []
    for key, value in query_pairs:
        if key == "LS":
            encoded_parts.append(f"LS={value}")
        else:
            encoded_parts.append(f"{quote(key, safe='')}={quote(value, safe='')}")
    return urlunsplit(
        (
            split.scheme,
            split.netloc,
            split.path,
            "&".join(encoded_parts),
            "",
        )
    )


def resolve_year_scoped_url(base_url: str, *, param: str, program_id: str | int) -> str:
    return replace_ls_param(base_url, param, str(program_id))
