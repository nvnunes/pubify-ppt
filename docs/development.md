# Development Setup

This document covers local setup and daily workflow. For repo boundaries and
package ownership, use `architecture.md`. For canonical verification commands
and completion expectations, use `testing.md`.

## Shared Guidance

This repo adopts the shared guidance in:

- `astro-agents/guidance/public-python-projects.md`
- `astro-agents/guidance/python-development.md`

Repo-local environment setup, toolchain choices, and daily commands in this
document remain the source of truth for this repo.

## Environment

- target Python 3.10+
- use the repo-local `./.conda` environment by default

For a fresh clone, create the local environment with:

```bash
conda create -p ./.conda python=3.12 pip -y
```

Then install the package and dev dependencies with:

```bash
./.conda/bin/pip install -e ".[dev]"
```

## Daily Commands

Prefer commands from the local environment instead of bare `python`, `pip`, or
`mkdocs` invocations:

```bash
./.conda/bin/pip install -e ".[dev]"
./.conda/bin/pytest tests -q
./.conda/bin/mkdocs build --strict
```

Use the installed CLI from the same environment when checking the local command
surface:

```bash
./.conda/bin/ppt --help
```
