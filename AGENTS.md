# AGENTS.md

## Cursor Cloud specific instructions

CISiphyus is a Python 3 + Playwright CLI that exports CISDM reports. The startup
update script provisions a `.venv` (editable install of `.[dev]`) and the
Playwright Chromium browser, so dependencies are already in place when a session
starts. Activate with `source .venv/bin/activate` or call tools via `.venv/bin/…`.

### Running things
- Tests (self-contained, no live CISDM access): `.venv/bin/python -m pytest tests/`
- Downstream example apps run end-to-end against bundled synthetic fixtures —
  these are the recommended smoke tests since they exercise the framework without
  credentials. See `examples/README.md` and the root `README.md` for exact
  commands. Pass `--metric-workbook` / `--workbook` (and `--no-site-staff-filter`
  for QPR) to pin fixtures and skip the live fetch.
- CLI entry points: `.venv/bin/python run.py <report_id>` or the installed
  `cisiphyus` console script.

### Non-obvious caveats
- The core report fetch (`run.py <report_id>` without an example subcommand)
  requires live CISDM authentication and per-report export URLs. It fails fast
  with `Missing export URL env var: …` unless `config/reports.yaml` and
  `config/export_urls.env` exist and a Chrome profile has been bootstrapped via
  `run.py --bootstrap`. None of this is available in the cloud VM, so validate
  changes through the test suite and the example apps instead.
- `config/reports.yaml` is git-ignored machine-local config. Copy it from the
  template when the core fetch path is needed: `cp config/reports.example.yaml config/reports.yaml`.
- Example apps auto-fetch via cisiphyus when workbook flags are omitted; always
  pass the bundled fixture flags in CI/headless runs to avoid the live fetch.
