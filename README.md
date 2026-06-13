# CISIPHYUS - CISDM Report Automation Framework

A [Playwright](https://github.com/microsoft/playwright)-based automaton for secure, programmatic export of reports from CISDM.

## SUMMARY

CISiphyus automates the retrieval of CISDM reports, enabling scripted, repeatable, and auditable exports. It securely manages authentication via dedicated Chrome profiles, supports both direct-export and UI-based reports, and ensures report integrity through configurable validations.

### Dependencies

* Python 3.8+

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
pip install -r requirements.txt && playwright install chromium

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
python run.py --bootstrap
```

The CISDM session is persisted in `config/chrome-user-data/` (Playwright profile `Profile 1`) and reused on every later run. Log in via the browser, then press Enter in the terminal.

## USAGE

```bash
python run.py student_metrics_summary # Retrieve Metrics Tracking (All) export
# equivalent after pip install -e .:
cisiphyus student_metrics_summary
```

## EXAMPLES

CISiphyus allows for a variety of downstream applications to improve operations. The [`examples/`](examples/README.md) directory contains prototypical, fully functional applications of the framework.

**Accreditation site monitoring**
> Retrieve `accreditation` report and flag site-level compliance issues (reporting gaps, Tier I counts, case-management completeness).

```bash
python examples/accreditation/accreditation_monitor.py \
  --output-dir examples/accreditation/local_inputs/monitor-out

# equivalent after pip install -e .
cisiphyus accreditation \
  --output-dir examples/accreditation/local_inputs/monitor-out
```

**QPR workbook provisioning**
> Retrieve `student_metrics_summary` and generate per-site QPR import workbooks using bundled tools and fixtures.

```bash
python examples/qpr/qpr_provision.py examples/qpr/output \
  --grading-period 2.0
# equivalent after pip install -e .:
cisiphyus qpr examples/qpr/output \
  --grading-period 2.0
```

## CHANGELOG

* 0.3.0
  * Add `examples/` directory with prototypical downstream applications

* 0.2.0

  * Verify `Auth_CaseWorthy` / `Context_CaseWorthy` and export them to `config/CISDM_cookies.json` during bootstrap
  * Failures now write `result.json` / `diag.json` with redacted cookie inventories for diagnosis
  * Removed legacy `session.py` cookie-overlay module (superseded by bootstrap-generated export)
* 0.1.0

  * Initial Release
