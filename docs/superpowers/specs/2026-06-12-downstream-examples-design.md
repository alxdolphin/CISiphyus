# downstream applications design

**Status:** Approved for implementation | **Date:** 2026-06-12

## goal

Add prototypical, fully functional downstream applications to the CISiphyus repo under
`examples/` — practical reference implementations that affiliates can clone and run without
access to any private monorepo.

## decisions

- scope matches production minus Cloud Run, Drive, Apps Script, and docx pipeline
- CISiphyus is self-contained; examples bundle their Python modules and fixture workbooks
- QPR example: age-gated prefetch of `student_metrics_summary` + bundled `qpr_tools.py` provision logic
- top-level folder is `examples/`

## layout

```
examples/
  README.md
  accreditation/
    accreditation_cisiphyus_fetch.py
    accreditation_workbook.py
    accreditation_monitor.py
    fixtures/
      accreditation_monitor_minimal.xlsx
  qpr/
    qpr_cisiphyus_prefetch.py
    qpr_provision.py
    qpr_tools.py                  # bundled provision logic
    qpr_tools_loader.py
    fixtures/
      student_metrics_summary_minimal.xlsx
      reporting_deadlines_minimal.xlsx
      qpr_import_template.xlsx
      site_staff_list_minimal.xlsx
tests/
  test_examples_accreditation.py
  test_examples_qpr_provision.py
```

## export map

| Application | CISiphyus report | Output |
|---|---|---|
| Accreditation monitoring | `accreditation` | Site flag CSVs / summaries |
| QPR provisioning | `student_metrics_summary` | Per-site `Q{n}_{Site}_QPR.xlsx` |

## contract adjustments in vendored modules

- `_default_cisiphyus_root()` resolves to this repo root (`Path(__file__).resolve().parents[2]`),
  overridable via `CISIPHYUS_ROOT` / `ACCREDITATION_CISPHYUS_ROOT` / `QPR_CISPHYUS_ROOT`.
- subprocess command matches the current slim CLI surface only:
  `python run.py <report_id>` plus optional `--headed`.
- artifact contract unchanged: success is exit 0 + `artifacts/latest/<report_id>/raw.xlsx`.
- bundled `qpr_tools.py` path constants resolve under `examples/qpr/fixtures/` and `local_inputs/`.

## application 1: accreditation site monitoring

Live export resolution is implicit: omit `--workbook` and the monitor retrieves via cisiphyus
when the local workbook is missing or stale (24h gate). Pass `--workbook` only to pin a local
file for CI or reproducible reruns.

## application 2: qpr workbook provisioning

`qpr_provision.py` is the sole QPR entrypoint:

1. implicitly prefetch `student_metrics_summary` via `qpr_cisiphyus_prefetch.py` when
   `--metric-workbook` is omitted and the local workbook is missing or stale
2. load bundled `examples/qpr/qpr_tools.py`
3. call `provision_all_site_templates` / `provision_site_template`

```bash
python examples/qpr/qpr_provision.py /tmp/qpr-out --grading-period 2.0
```

Tests pin `--metric-workbook` and related fixtures; runtime defaults always use cisiphyus retrieval.

Outputs: `{output}/{school-year}/Q{n}/Q{n}_{Site}_QPR.xlsx` per site.

## out of scope

- publishing or mirroring a private CIS monorepo
- student-metrics audit (`audit_metrics_all`) as a standalone QPR example
- docx generation, uploads, notifications, Cloud Run
- credentials or export URLs in git
