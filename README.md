# pubify-ppt

`pubify-ppt` is the PowerPoint-oriented downstream package for the
presentation-agnostic `pubify-data` runtime.

It is intended for host workspaces that keep presentations, presentation-local
data, and editable `.pptx` decks under version control. `pubify-ppt` owns the
PowerPoint-specific workflow: locating deck anchors, rendering figures into
slides, replacing inline stats, updating native tables, writing decks, and
managing generated PowerPoint artifacts.

This package does not own your presentations. A host workspace does.

## Project Docs

- [Documentation home](https://nvnunes.github.io/pubify-ppt/)
- [Architecture](https://nvnunes.github.io/pubify-ppt/architecture/)
- [Usage](https://nvnunes.github.io/pubify-ppt/usage/)
- [Development setup](https://nvnunes.github.io/pubify-ppt/development/)
- [Testing and validation](https://nvnunes.github.io/pubify-ppt/testing/)
- [API reference](https://nvnunes.github.io/pubify-ppt/api/)

## Current Status

The package skeleton, initialization workflow, inventory commands, static
presentation checks, figure updates, stat updates, table updates, full updates,
generated-copy output, and in-place deck backups are in place. `ppt init`
creates or updates workspace config, and `ppt init <presentation-id>` creates a
starter presentation scaffold with `ppt.yaml`, `figures.py`, `deck.pptx`, and
`data/ppt-artifacts/`.

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
ppt demo table list
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
slides/
slides/<presentation-id>/
  figures.py
  ppt.yaml
  deck.pptx
  data/
    ppt-artifacts/
      figures/
      backups/
```

See `docs/architecture.md` for the stable package and workspace contracts.

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
ppt <presentation-id> table list
ppt <presentation-id> table update
ppt <presentation-id> table <table-id> update
ppt <presentation-id> update
```

Write commands accept `--output <path>` to write a generated copy instead of
mutating `deck.pptx`.

## Authoring Summary

- Figure anchors live in PowerPoint alt text, for example
  `{{fig:example}}`.
- To bootstrap a figure, draw a supported shape and set its visible text
  exactly to `{{fig:example}}`; update will replace it with a picture and carry
  the token forward in Alt Text.
- Figure metadata supports export padding, for example
  `FigureResult(fig, metadata={"export_pad_inches": 0.02})`, with optional
  side-specific `export_pad_left_inches`, `export_pad_right_inches`,
  `export_pad_top_inches`, and `export_pad_bottom_inches`. Side-specific
  padding expands Matplotlib's tight export bounding box before rendering, which
  helps keep labels and spines away from PowerPoint image boundaries.
- Figure PNGs are rendered on an anchor-sized canvas at the configured DPI, so
  PowerPoint does not rescale generated text after export.
- Figure text uses the PowerPoint theme body font by default. Set
  `defaults.figure_font_family` in `ppt.yaml` to override the discovered theme
  font for generated figures; unavailable fonts are ignored to avoid Matplotlib
  `findfont` noise. Set `defaults.figure_*_fontsize_pt` values to control
  presentation-level generated figure text sizes.
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
- Table anchors use `{{table:example}}` in a placeholder's visible text or Alt
  Text. Update replaces the placeholder with a native PowerPoint table and
  persists the token in Alt Text.
- Table headings can come from `TableResult(..., metadata={"columns": (...)})`
  when a table is first created. Later updates preserve the PowerPoint heading
  row and require the table to keep the same total row count and column count.
- Simple rectangle placeholders and previously generated pictures are the
  supported figure anchor shapes. Simple placeholders and generated native
  tables are the supported table anchor shapes.
- Grouped, rotated, cropped, animated, table-contained, chart-contained,
  SmartArt-contained, and embedded-object-contained anchors are unsupported.
