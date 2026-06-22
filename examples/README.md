# downstream applications

Prototypical, fully functional applications that consume CISiphyus exports. Clone this
repo, bootstrap auth, and run — no private monorepo checkout is required.

Each application retrieves its CIS export via cisiphyus when you omit `--workbook` /
`--metric-workbook` (age-gated refresh when missing or stale). Pass those flags only to
pin a local file for CI or reproducible reruns.

| Application | CISiphyus report | Output |
|---|---|---|
| Accreditation monitoring | `accreditation` | Site flag CSVs / summaries |
| QPR provisioning | `student_metrics_summary` | Per-site `Q{n}_{Site}_QPR.xlsx` |
| Longitudinal trend tracker | `accreditation` + `student_metrics_summary` | QoQ movement + regression CSVs / summaries |

## accreditation site monitoring (`accreditation/`)

Retrieves the `accreditation` report (`ui_export` strategy) and flags site-level
compliance issues: site-coordination reporting gaps, Tier I support counts, and
case-management completeness.

```bash
python examples/accreditation/accreditation_monitor.py --output-dir /tmp/accreditation-out
```

Force a fresh export even when the cached workbook is still within the 24h gate:

```bash
python examples/accreditation/accreditation_monitor.py \
  --output-dir /tmp/accreditation-out \
  --force-fetch
```

Outputs: `all_flags.csv`, `sc_flags.csv`, `cm_flags.csv`, `tier1_flags.csv`,
`monitoring_summary.json`, `monitoring_summary.md`.

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

Outputs: `artifacts/qpr/{school-year}/Q{n}/Q{n}_{Site}_QPR.xlsx` per site.

## longitudinal trend tracker (`trends/`)

Discovers available reporting periods for each school year, captures snapshots, and
compares consecutive quarters.

```bash
cisiphyus trend SY25-26
cisiphyus trend SY25-26 SY24-25
cisiphyus trend --all
```

### period discovery

Periods come from `artifacts/qpr/{school-year}/Q{n}/` folders (same layout as QPR output),
cached archives under `artifacts/archives/{school-year}/Q{n}/`, and existing snapshots.
`--all` also includes every school year listed in `school_year_programs` inside
`config/reports.yaml`. Years without QPR quarter folders are pulled from CISDM,
archived under `artifacts/archives/{school-year}/EOY/`, and captured as
`artifacts/snapshots/{school-year}/EOY/` snapshots (metrics-only for historical years).

`--all` also runs cross-year compares (`artifacts/trends/cross_year/`) between
consecutive EOY snapshots (prior school year EOY vs next school year EOY). Years
without an EOY snapshot are skipped until one is captured.

The cross-year HTML report has two sections:

- **Student Metrics Summary data quality** — changed issue counts and issue counts increased (issue-flag / flagged-row deltas).
- **Goal progress** — Goal–Metric rows with Baseline and Target where the latest filled Grading Period value is compared to Target; includes % meets Target and School Year–over–School Year Goal–Metric row improved/worsened counts.

Snapshots store `metrics/progress_rollup.json` for goal progress. Re-run `backfill_progress_rollups` or `cisiphyus trend --all` to populate historical snapshots.

Compare output warns when period snapshots share identical workbook hashes (CISDM exports
are school-year-scoped, not quarter-scoped — run trend at each period close for meaningful QoQ).

Per-year status is printed: `snapshotted`, `compared`, `skipped`, or `failed`.

Run at the end of each grading period so each quarter gets its own CISDM pull. Older
quarters without a cached pull are skipped until you have snapshot history for them.

```bash
cisiphyus trend SY25-26          # pull + capture latest missing period, compare pairs
cisiphyus trend SY25-26 --force-fetch
```

Snapshots: `artifacts/snapshots/{school-year}/{period}/`  
Compare outputs: `artifacts/trends/{school-year}/{Q1_vs_Q2}/`

## environment variables

| Variable | Purpose |
|---|---|
| `CISIPHYUS_ROOT` | Override the cisiphyus repo root (default: this repo) |
| `CISPHYUS_HEADED` | `1` to show the browser during cisiphyus retrieval |
| `ACCREDITATION_WORKBOOK` / `ACCREDITATION_LOCAL_INPUTS_DIR` | Accreditation workbook destination |
| `ACCREDITATION_MAX_AGE_HOURS` | Freshness gate for the accreditation workbook (default 24) |
| `QPR_STUDENT_METRICS_WORKBOOK` / `QPR_LOCAL_INPUTS_DIR` | Student metrics workbook destination |
| `QPR_OUTPUT_DIR` | QPR provision output root (default `artifacts/qpr/`) |
| `QPR_FETCH_STUDENT_METRICS` | `1` to force a refresh even when fresh |
| `TREND_SNAPSHOTS_DIR` | Snapshot archive root (default `artifacts/snapshots/`) |
| `TREND_INPUTS_DIR` | Cached CISDM pull copies (default `artifacts/archives/`) |
| `TREND_OUTPUT_DIR` | Trend compare output root (default `artifacts/trends/`) |
| `TREND_SCHOOL_YEAR` | Default school year when none passed to `cisiphyus trend` |

## tests

```bash
pip install -e .[dev]
pytest tests/
```

Tests pass `--workbook` / `--metric-workbook` fixtures explicitly; no live CISDM access is required.
