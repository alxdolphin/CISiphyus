# PR split: `feat/school-year-scoped-exports`

The original branch combined ~6,600 LOC across unrelated features. Merge via three
reviewable pull requests:

## PR 1 — School-year scoped exports (`feat/school-year-scoping`)

Core CISDM export plumbing:

- `src/config.py`, `src/report.py`, `src/url_refresh.py`, `src/artifacts.py`
- `src/cli.py` (pull subcommand, `--school-year`, failure-path fix; no example dispatch)
- `config/reports.example.yaml`, `config/export_urls.env.example`
- `tests/test_school_year_scoping.py`, `.gitignore` test allowlist
- `README.md` school-year documentation

## PR 2 — Downstream example apps (`feat/downstream-examples`)

Stacks on PR 1:

- `examples/` (accreditation, audit, qpr, `_cisiphyus_fetch.py`)
- `src/example_cli.py`, full `src/cli.py` with audit/qpr dispatch
- `run.py`, `pyproject.toml`
- `tests/test_examples_*.py`, `examples/README.md`

## PR 3 — Nix dev environment (`chore/nix-dev`)

Independent chore off `main`:

- `flake.nix`, `flake.lock`, `shell.nix`

## Integration branch

`feat/school-year-scoped-exports` retains the full combined branch for reference until
the split PRs land.
