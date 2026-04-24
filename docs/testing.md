# Testing

This document is the source of truth for local verification commands and
completion expectations in `pubify-ppt`.

## Environment

Use the repo-local `./.conda` environment for Python commands, test runs, and
docs builds unless a task explicitly requires something else.

## Canonical Verification Commands

Run the package test suite with:

```bash
./.conda/bin/pytest tests -q
```

Build the docs with:

```bash
./.conda/bin/mkdocs build --strict
```

## Important Test Surfaces

- `tests/test_cli.py`: CLI parser and entrypoint behavior
- `tests/test_config.py`: workspace config loading
- `tests/test_discovery.py`: presentation path resolution and
  `pubify-data` adapter loading
- `tests/test_check.py`: static presentation validation and deck token checks
- `tests/test_figures.py`: figure rendering and PowerPoint anchor replacement
- `tests/test_runtime.py`: package/runtime smoke coverage
- `tests/test_anchors.py`: PowerPoint anchor module smoke coverage

## Completion Expectations

Run the full test suite before concluding substantial code changes.

Run the strict docs build for changes that affect:

- `README.md`
- `docs/*`
- `AGENTS.md`
- `mkdocs.yml`
- package metadata or docstrings that feed the public docs surface

Targeted tests are acceptable during iteration, but final verification should
match the changed surface area.
