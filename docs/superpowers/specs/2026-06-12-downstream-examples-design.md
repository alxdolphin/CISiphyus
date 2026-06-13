# downstream applications design

**Status:** Approved for implementation | **Date:** 2026-06-12

## goal

Add prototypical, fully functional downstream applications to the CISiphyus repo under
`examples/` — practical reference implementations that consume CISiphyus exports and
produce the same artifacts as their production counterparts.

## decisions

- scope matches production minus Cloud Run, Drive, Apps Script, and docx pipeline
- CIS monorepo stays canonical; examples are trimmed copies refreshed via `examples/sync_from_cis.sh`
- QPR example: age-gated prefetch of `student_metrics_summary` + production `qpr_tools.py` provision-batch logic
- top-level folder is `examples/`

## layout

```
examples/
  README.md                     # overview + quickstart for all applications
  sync_from_cis.sh              # refreshes vendored files from a CIS monorepo checkout
  accreditation/
    accreditation_cisiphyus_fetch.py
    accreditation_workbook.py
    accreditation_monitor.py
    fixtures/
      accreditation_monitor_minimal.xlsx
  qpr/
    qpr_cisiphyus_prefetch.py   # age-gated fetch library
    qpr_provision.py            # sole QPR CLI (fetch + provision)
    qpr_tools_loader.py
    fixtures/
      student_metrics_summary_minimal.xlsx
      reporting_deadlines_minimal.xlsx
      qpr_import_template.xlsx
tests/
  test_examples_accreditation.py
  test_examples_qpr_provision.py
```

## export map

| Application | CISiphyus report | Output |
|---|---|---|
| Accreditation monitoring | `accreditation` | Site flag CSVs / summaries |
| QPR provisioning | `student_metrics_summary` | Per-site `Q{n}_{Site}_QPR.xlsx` |

## contract adjustments vs production copies

- `_default_cisiphyus_root()` resolves to this repo root (`Path(__file__).resolve().parents[2]`),
  overridable via `CISIPHYUS_ROOT` / `ACCREDITATION_CISPHYUS_ROOT` / `QPR_CISPHYUS_ROOT`.
- subprocess command matches the current slim CLI surface only:
  `python run.py <report_id>` plus optional `--headed`.
- artifact contract unchanged: success is exit 0 + `artifacts/latest/<report_id>/raw.xlsx`.

## application 1: accreditation site monitoring

Live export resolution is implicit: omit `--workbook` / `--metric-workbook` and each application
retrieves via cisiphyus when the local workbook is missing or stale (24h gate). Pass those flags
only to pin a local file for CI or reproducible reruns.

## application 2: qpr workbook provisioning

`qpr_provision.py` is the sole QPR entrypoint:

1. implicitly prefetch `student_metrics_summary` via `qpr_cisiphyus_prefetch.py` when
   `--metric-workbook` is omitted and the local workbook is missing or stale
2. load production `tools/qpr/qpr_tools.py` from `CIS_MONOREPO_ROOT`
3. call `provision_all_site_templates` / `provision_site_template`

```bash
python examples/qpr/qpr_provision.py /tmp/qpr-out \
  --grading-period 2.0 --no-site-staff-filter
```

Tests pin `--metric-workbook` and related fixtures; runtime defaults always use cisiphyus retrieval.

Outputs: `{output}/{school-year}/Q{n}/Q{n}_{Site}_QPR.xlsx` per site.

## sync workflow

`examples/sync_from_cis.sh` copies the manifest from a CIS checkout; contract
adjustments must be re-applied after sync (see script header).

## out of scope

- replacing production `tools/qpr` / `tools/accreditation` in the CIS monorepo
- student-metrics audit (`audit_metrics_all`) as a standalone QPR example
- docx generation, uploads, notifications, Cloud Run
- credentials or export URLs in git
