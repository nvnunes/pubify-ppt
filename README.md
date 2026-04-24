# pubify-ppt

`pubify-ppt` is the PowerPoint-oriented downstream package for the
presentation-agnostic `pubify-data` runtime.

It is intended for host workspaces that keep presentations, presentation-local
data, and editable `.pptx` decks under version control. `pubify-ppt` owns the
PowerPoint-specific workflow: locating deck anchors, rendering figures into
slides, replacing inline stats, writing decks, and managing generated
PowerPoint artifacts.

This package does not own your presentations. A host workspace does.

## Project Docs

- [Architecture](docs/architecture.md)
- [Development setup](docs/development.md)
- [Testing and validation](docs/testing.md)
- [Implementation plan](docs/plan.md)

## Current Status

This repository is in the foundation phase. The package skeleton, CLI
entrypoint, docs, and baseline tests are being established before the
PowerPoint runtime is implemented.

## Planned Workspace Model

`pubify-ppt` discovers a host workspace from `pubify.yaml`:

```yaml
pubify-ppt:
  presentations_root: slides
```

Each presentation will live under `slides/<presentation-id>/` and use:

```text
slides/<presentation-id>/
  figures.py
  ppt.yaml
  deck.pptx
  data/
    ppt-artifacts/
```

See `docs/plan.md` for the phased implementation contract.
