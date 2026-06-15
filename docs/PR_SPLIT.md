# PR split: school-year scoped exports

`main` already includes downstream example apps (PR 2) and the Nix dev environment
(PR 3). The remaining work is school-year scoped CISDM export plumbing.

## PR 1 — School-year scoped exports (`feat/school-year-scoping`) — **open**

Core CISDM export plumbing plus review fixes:

- `src/config.py`, `src/report.py`, `src/url_refresh.py`, `src/artifacts.py`, `src/cli.py`
- `config/reports.example.yaml`, `config/export_urls.env.example`
- `tests/test_school_year_scoping.py`, `.gitignore` test allowlist
- `README.md` school-year documentation
- Example integration: `examples/_cisiphyus_fetch.py`, `--school-year` on audit/qpr,
  standardized `pull` subcommand

## PR 2 — Downstream example apps — **merged to `main`**

## PR 3 — Nix dev environment — **merged to `main`**

## Integration branch

`feat/school-year-scoped-exports` retains the full combined branch history for reference.
