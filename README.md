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
- [PowerPoint authoring](docs/powerpoint.md)
- [Migration guide](docs/migration.md)
- [Development setup](docs/development.md)
- [Testing and validation](docs/testing.md)
- [Implementation plan](docs/plan.md)

## Current Status

The package skeleton, initialization workflow, inventory commands, static
presentation checks, figure updates, stat updates, full updates, generated-copy
output, and in-place deck backups are in place. `ppt init` creates or updates
workspace config, and `ppt init <presentation-id>` creates a starter
presentation scaffold with `ppt.yaml`, `figures.py`, `deck.pptx`, and
`data/ppt-artifacts/`.

Tables and backup inspection/restore commands are still planned work.

## Quick Start

Create or update workspace config:

```bash
ppt init
```

Create a starter presentation:

```bash
ppt init demo
```

Inspect the discovered presentation surface:

```bash
ppt list
ppt demo data list
ppt demo figure list
ppt demo stat list
```

Validate anchors, tokens, data paths, and dependencies:

```bash
ppt demo check
```

Update the source deck in place:

```bash
ppt demo update
```

Write a generated copy instead of mutating `deck.pptx`:

```bash
ppt demo update --output exports/demo.pptx
```

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
      figures/
      backups/
```

See `docs/plan.md` for the phased implementation contract.

## Implemented Commands

```bash
ppt list
ppt init
ppt init <presentation-id>
ppt <presentation-id> check
ppt <presentation-id> data list
ppt <presentation-id> figure list
ppt <presentation-id> figure update
ppt <presentation-id> figure <figure-id> update
ppt <presentation-id> stat list
ppt <presentation-id> stat update
ppt <presentation-id> stat <stat-id> update
ppt <presentation-id> update
```

Write commands accept `--output <path>` to write a generated copy instead of
mutating `deck.pptx`.

## Authoring Summary

- Figure anchors live in PowerPoint alt text, for example
  `{{fig:example}}`.
- Multi-panel figures use one explicit anchor per panel, for example
  `{{fig:comparison:1}}`.
- Reused paper outputs are declared in `ppt.yaml` under `sources:` and exposed
  through presentation-local wrapper functions in `figures.py`; PowerPoint
  anchors should reference only local IDs.
- Stat tokens live in text boxes, for example `{{stat:example.count}}`.
- Valid stat tokens remain supported if PowerPoint splits them across internal
  text runs.
- After the first update, stat text boxes retain
  `{{stat:example.count=<previous_value>}}` in Alt Text so later updates can
  safely replace the previous visible value.
- Simple rectangle placeholders and previously generated pictures are the
  supported figure anchor shapes.
- Grouped, rotated, cropped, animated, table-contained, chart-contained,
  SmartArt-contained, and embedded-object-contained anchors are unsupported.
