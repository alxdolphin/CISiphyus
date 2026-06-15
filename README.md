# CISIPHYUS - CISDM Report Automation Framework

A [Playwright](https://github.com/microsoft/playwright)-based automaton for secure, programmatic export of reports from CISDM.

## SUMMARY

CISiphyus automates the retrieval of CISDM reports, enabling scripted, repeatable, and auditable exports. It securely manages authentication via dedicated Chrome profiles, supports both direct-export and UI-based reports, and ensures report integrity through configurable validations.

### Dependencies

* Python 3.10+

  * [playwright](https://github.com/microsoft/playwright)
      `chromium`
  * [PyYAML](https://github.com/yaml/pyyaml)
  * [openpyxl](https://github.com/openpyxl/openpyxl)
* Chromium / Chrome browser

## SETUP

```bash
git clone https://github.com/alxdolphin/CISiphyus.git
cd CISiphyus
python -m venv .venv && source .venv/bin/activate
pip install -e . && playwright install chromium

# OR

nix develop
```

**CISDM Export / Report URLs**

```bash
cp config/export_urls.env.example local/export_urls.env
```

For each export / report you want to retrieve, you need to get the URL from CISDM and add it to the `local/export_urls.env` file.

| Type   | URL                          |
| ------ | ---------------------------- |
| Export | Page Options → Excel Export |
| Report | Report Viewer → 💾 → Excel |

**Authentication**

```bash
cisiphyus --bootstrap
```

The CISDM session is persisted in `config/chrome-user-data/` (Playwright profile `Profile 1`) and reused on every later run. Log in via the browser, then press Enter in the terminal.

## USAGE

```bash
# Pull raw exports by report ID (according to `config/reports.yaml`)
cisiphyus pull accreditation
cisiphyus pull student_metrics_summary

# Pull a historical school year (year-scoped direct exports only)
cisiphyus pull student_metrics_summary --school-year SY24-25
```

### School-year scoped exports

Reports with `year_scope` in `config/reports.yaml` (for example
`student_metrics_summary`) inject the enrollment program ID for the selected school
year into the CISDM export URL.

Configure school years in `config/reports.yaml`:

```yaml
school_year_programs:
  SY24-25: 1330
  SY25-26: 1331
```

| Setting | Purpose |
| --- | --- |
| `--school-year SYxx-yy` | Override the school year for this pull |
| `CISDM_DEFAULT_SCHOOL_YEAR` | Default year when `--school-year` is omitted (must appear in `school_year_programs`) |
| (implicit default) | Latest school year in `school_year_programs` when env is unset |

**Output layout**

| Pull | Output path |
| --- | --- |
| Default / current school year | `artifacts/latest/<report_id>/raw.xlsx` |
| Non-default school year | `artifacts/archives/<SY>/pulls/<report_id>/raw.xlsx` |

Each run writes `result.json` (and `diag.json` on failure) alongside `raw.xlsx`.
Failures for archived pulls are written under the archive path, not `artifacts/latest/`.

## EXAMPLES

CISiphyus allows for a variety of downstream applications to improve operations. The [`examples/`](examples/README.md) directory contains prototypical, fully functional applications of the framework.

**Accreditation site monitoring**
> Retrieve `accreditation` report and flag site-level compliance issues (reporting gaps, Tier I counts, case-management completeness).

```bash
cisiphyus audit accreditation
```

**Audit Student Metrics**
> Retrieve `student_metrics_summary` and flag row-level data-quality issues.

```bash
cisiphyus audit metrics
```

**Provision Quarterly Progress Reports**
> Retrieve `student_metrics_summary` and generate per-site QPR import workbooks using bundled tools and fixtures.

```bash
cisiphyus qpr --grading-period 2.0
```

## CHANGELOG

* 0.4.0
  * Group downstream audits under `cisiphyus audit accreditation` and `cisiphyus audit metrics`
  * Add `cisiphyus pull <report_id>` for raw exports
  * Default example outputs to `artifacts/<app>/` (override with `*_OUTPUT_DIR` env vars)
  * Consolidate accreditation example into single `accreditation.py` (fetch, parse, rules, CLI)
  * Consolidate QPR example into `qpr.py` (fetch, loader, provision CLI); keep `qpr_tools.py` as bundled library

* 0.3.0
  * Add `examples/` directory with prototypical downstream applications

* 0.2.0
  * Verify `Auth_CaseWorthy` / `Context_CaseWorthy` and export them to `config/CISDM_cookies.json` during bootstrap
  * Failures now write `result.json` / `diag.json` with redacted cookie inventories for diagnosis
  * Removed legacy `session.py` cookie-overlay module (superseded by bootstrap-generated export)

* 0.1.0
  * Initial Release
