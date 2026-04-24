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

The package skeleton and initialization workflow are in place. `ppt init`
creates or updates workspace config, and `ppt init <presentation-id>` creates a
starter presentation scaffold with `ppt.yaml`, `figures.py`, `deck.pptx`, and
`data/ppt-artifacts/`.

Inventory, check, update, figure rendering, stat replacement, and deck backup
commands are still planned work.

## Workspace Model

`pubify-ppt` discovers a host workspace from `pubify.yaml`:

```yaml
pubify-ppt:
  presentations_root: slides
```

Each presentation lives under `slides/<presentation-id>/` and uses:

```text
slides/<presentation-id>/
  figures.py
  ppt.yaml
  deck.pptx
  data/
    ppt-artifacts/
```

See `docs/plan.md` for the phased implementation contract.
