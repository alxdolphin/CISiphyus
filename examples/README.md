# downstream applications

Prototypical, fully functional applications that consume CISiphyus exports. Clone this
repo, bootstrap auth, and run — no private monorepo checkout is required.

Each application retrieves its CIS export via cisiphyus when you omit `--workbook` /
`--metric-workbook` (age-gated refresh when missing or stale). Pass those flags only to
pin a local file for CI or reproducible reruns.

| Application | CISiphyus report | Output |
|---|---|---|
| Accreditation monitoring | `accreditation` | Site flag CSVs / summaries |
| Student metrics audit | `student_metrics_summary` | Row-level audit CSVs / JSON |
| QPR provisioning | `student_metrics_summary` | Per-site `Q{n}_{Site}_QPR.xlsx` |

## accreditation site monitoring (`accreditation/`)

Retrieves the `accreditation` report (`ui_export` strategy) and flags site-level
compliance issues: site-coordination reporting gaps, Tier I support counts, and
case-management completeness.

```bash
cisiphyus audit accreditation
```

Force a fresh export even when the cached workbook is still within the 24h gate:

```bash
cisiphyus audit accreditation --force-fetch
```

Outputs: `all_flags.csv`, `sc_flags.csv`, `cm_flags.csv`, `tier1_flags.csv`,
`monitoring_summary.json`, `monitoring_summary.md`.

## student metrics audit (`audit/`)

Retrieves `student_metrics_summary` and flags row-level data-quality issues (invalid
metric/scale combinations, duplicate keys, and related checks).

```bash
cisiphyus audit metrics
```

Pin a local workbook or pull a specific school year:

```bash
cisiphyus audit metrics --workbook path/to/StudentMetricsSummary.xlsx
cisiphyus audit metrics --school-year SY24-25
```

## qpr workbook provisioning (`qpr/`)

Retrieves `student_metrics_summary` (`direct_url` strategy), then generates per-site QPR
import workbooks (`.xlsx`) using bundled [`qpr_tools.py`](qpr/qpr_tools.py) provision
logic and default fixtures under `examples/qpr/fixtures/`.

```bash
cisiphyus qpr --grading-period 2.0
```

Single site:

```bash
cisiphyus qpr --grading-period 2.0 --school-name "Lincoln HS"
```

Override bundled template, deadlines, or site-staff inputs when your affiliate uses
local workbooks:

```bash
cisiphyus qpr --grading-period 2.0 \
  --template /path/to/QPR_Import_Template.xlsx \
  --deadlines /path/to/ReportingDeadlines.xlsx \
  --site-staff-list /path/to/Site_Staff_List.xlsx
```

Historical school year:

```bash
cisiphyus qpr --grading-period 2.0 --school-year SY24-25
```

Outputs: `artifacts/qpr/{school-year}/Q{n}/Q{n}_{Site}_QPR.xlsx` per site.

## environment variables

| Variable | Purpose |
|---|---|
| `CISIPHYUS_ROOT` | Override the cisiphyus repo root (default: this repo) |
| `CISPHYUS_HEADED` | `1` to show the browser during cisiphyus retrieval |
| `CISDM_DEFAULT_SCHOOL_YEAR` | Default school year for pulls and example apps |
| `ACCREDITATION_WORKBOOK` / `ACCREDITATION_LOCAL_INPUTS_DIR` | Accreditation workbook destination |
| `ACCREDITATION_MAX_AGE_HOURS` | Freshness gate for the accreditation workbook (default 24) |
| `AUDIT_STUDENT_METRICS_FILENAME` / `AUDIT_LOCAL_INPUTS_DIR` | Student metrics workbook destination for audit |
| `QPR_STUDENT_METRICS_WORKBOOK` / `QPR_LOCAL_INPUTS_DIR` | Student metrics workbook destination for QPR |
| `QPR_OUTPUT_DIR` | QPR provision output root (default `artifacts/qpr/`) |
| `QPR_FETCH_STUDENT_METRICS` | `1` to force a refresh even when fresh |

Example apps use `cisiphyus pull <report_id>` under the hood. When `--school-year` is
omitted, they follow `CISDM_DEFAULT_SCHOOL_YEAR` or the latest year in
`config/reports.yaml`.

## tests

```bash
pip install -e .[dev]
pytest tests/
```

Tests pass `--workbook` / `--metric-workbook` fixtures explicitly; no live CISDM access is required.
