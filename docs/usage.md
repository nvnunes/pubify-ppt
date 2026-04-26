# Usage

This document describes the user-facing workflow for authoring, checking, and
updating `pubify-ppt` presentations.

## Workspace And Presentation Setup

Create or update workspace config from the host workspace root:

```bash
ppt init
```

This also creates a shared presentations-root `AGENTS.md`, usually
`slides/AGENTS.md`, with guidance for agents working in downstream presentation
workspaces.

Create a starter presentation:

```bash
ppt init demo
```

This creates:

```text
slides/demo/
  deck.pptx
  figures.py
  ppt.yaml
  data/
    ppt-artifacts/
```

The starter presentation includes a runnable example loader, figure, stat, and
table plus matching deck anchors. Running `ppt demo update` replaces the
starter figure, stat, and table tokens in `deck.pptx`.

If you already have a deck, replace `slides/demo/deck.pptx` with your existing
PowerPoint file after initialization.

Put presentation-local source data under:

```text
slides/demo/data/
```

If the data must physically live elsewhere, make `slides/demo/data` a symlink
to that location before running presentation init or updates. `pubify-ppt`
still treats the presentation-local `data/` path as canonical.

## Deck Source Model

Each presentation owns an editable source deck:

```text
slides/<presentation-id>/deck.pptx
```

The `.pptx` file is both source and layout editor. Resize placeholders or
previously generated pictures in PowerPoint, then rerun `ppt <presentation-id>
update` to regenerate content at the current deck geometry.

## Authoring Figures

Figure anchors use PowerPoint alt text as their persistent managed state:

```text
{{fig:<figure_id>}}
{{fig:<figure_id>:<panel_number>}}
```

Use `{{fig:<figure_id>}}` for single-panel figures. Multi-panel figures require
one explicit anchor per panel, using one-based panel numbers such as
`{{fig:comparison:1}}` and `{{fig:comparison:2}}`.

For first-time bootstrapping, you can also draw a supported shape and set its
visible text to exactly one figure token, with no surrounding text:

```text
{{fig:<figure_id>}}
```

On update, `pubify-ppt` treats that shape as a figure anchor, replaces it with
the rendered picture, and writes the same token into the picture's alt text. If
a shape has both figure-token alt text and visible figure-token text, they must
match.

Figure anchors must reference presentation-local figure IDs defined in the
presentation's `figures.py`. Source publication IDs are not valid PowerPoint
anchors; expose reused source outputs through local wrapper functions instead.

To convert a manual figure incrementally:

- add a loader and `@figure` function to `figures.py`
- draw a simple rectangle in PowerPoint where the figure should appear
- set the rectangle's alt text description to `{{fig:<figure_id>}}`, or set its
  visible text exactly to that token
- run `ppt demo figure <figure_id> update`

Once the figure appears correctly, resize the generated picture in PowerPoint
as needed and rerun the same command. The deck geometry drives future render
size.

## Reusing Paper Outputs

A presentation can reuse figures, stats, and tables from another pubify
publication by declaring a source in `ppt.yaml`:

```yaml
sources:
  ao4elt8: papers/ao4elt8
```

The source root uses the conventional pubify layout:

```text
papers/ao4elt8/
  figures.py
  data/
```

Source outputs are reused from presentation-local wrapper functions:

```python
from pubify_data import figure, stat, table
from pubify_ppt import FigureResult, StatResult, TableResult


@figure
def plot_training_fov_ee(ctx):
    panel = ctx.source("ao4elt8").figure("training_compare_fov_ee").panel(1)
    return FigureResult(panel)


@stat
def compute_training_count(ctx):
    source_stat = ctx.source("ao4elt8").stat("training_count")
    return StatResult(source_stat.values[0].value)


@table
def tabulate_training_summary(ctx):
    source_table = ctx.source("ao4elt8").table("training_summary")
    return TableResult(source_table.bodies[0], metadata=source_table.metadata)
```

The PowerPoint deck then uses only the local IDs:

```text
{{fig:training_fov_ee}}
{{stat:training_count}}
{{table:training_summary}}
```

This keeps the editable deck independent of source publication internals such
as paper IDs, source figure names, and panel numbering. Rename or adapt source
outputs in `figures.py`, not in PowerPoint alt text.

## Supported Figure Anchor Shapes

Supported in the current implementation:

- plain rectangle or freeform placeholder shapes with a valid `{{fig:...}}`
  token in alt text
- plain rectangle or freeform placeholder shapes whose visible text is exactly
  one valid `{{fig:...}}` token
- existing picture shapes created by a previous `ppt ... figure update` or
  `ppt ... update`

Unsupported in the current implementation:

- grouped shapes
- rotated shapes
- cropped picture shapes
- shapes inside tables, charts, SmartArt, or embedded objects
- animated shapes

If an unsupported anchor is detected, `ppt <presentation-id> check` or an
update command fails with the slide number, token, and unsupported feature. The
reliable fix is to replace the anchor with a simple rectangle, then rerun the
command.

## Authoring Stats

New stats are authored as inline text tokens in PowerPoint text boxes:

```text
{{stat:<stat_id>}}
{{stat:<stat_id>.<key>}}
```

Scalar stats use `{{stat:<stat_id>}}`. Dictionary stats use
`{{stat:<stat_id>.<key>}}`.

PowerPoint may split a visible token across multiple internal text runs,
especially around punctuation or spell-check boundaries. `pubify-ppt` treats
the paragraph text as the source token stream, so valid split-run tokens are
supported.

On first update, `pubify-ppt` replaces the visible token with the computed
value and writes the prior rendered value into the text box Alt Text:

```text
{{stat:<stat_id>=<previous_value>}}
{{stat:<stat_id>.<key>=<previous_value>}}
```

On later updates, `pubify-ppt` reads that Alt Text anchor, finds the previous
value in the same text box, replaces it with the newly computed value, and
updates the Alt Text. The previous value must appear exactly once. If it is not
found, restore the previous value or reinsert the original `{{stat:...}}` token
so the update can proceed safely.

Each stat-managed text box supports one stat token. Split multiple stats into
separate text boxes.

## Authoring Tables

New tables are authored with one placeholder token:

```text
{{table:<table_id>}}
```

The token can be either the placeholder's exact visible text or its Alt Text
description. On first update, `pubify-ppt` replaces the placeholder with a
native PowerPoint table at the same geometry and writes the same token into
the table's Alt Text.

Table functions can provide default column headings in `metadata["columns"]`:

```python
from pubify_data import table
from pubify_ppt import TableResult


@table
def tabulate_summary(ctx):
    rows = [
        ("A", 10),
        ("B", 20),
    ]
    return TableResult(rows, metadata={"columns": ("label", "value")})
```

When present, `metadata["columns"]` must be an ordered sequence of strings and
its length must match the table data width. If it is omitted, first-time table
creation uses blank heading cells. PowerPoint table cells are text-only in
this implementation; `None` becomes an empty string and other values are
converted with `str(...)`.

On later updates, `pubify-ppt` finds the native table by Alt Text and mutates
the existing body cells in place. The heading row is left as-is, so edits made
in PowerPoint are preserved and `metadata["columns"]` is ignored after the
native table exists. The table must keep the same total row count, including
the heading row, and the same column count as the computed table. If the
dimensions differ, adjust the PowerPoint table size or replace the table with
the original `{{table:...}}` token and rerun the update.

One `{{table:...}}` anchor maps to one native PowerPoint table. Multi-body
table results are not supported in v1.

## Checking A Deck

Run:

```bash
ppt demo check
```

This catches missing data files, unknown anchors, duplicate anchors, malformed
tokens, unsupported figure/table anchor shapes, and ambiguous stat-managed text
boxes before an update writes the deck.

## Updating A Deck

Use default in-place updates during normal iteration:

```bash
ppt demo update
```

Targeted updates are available when iterating on one surface:

```bash
ppt demo figure <figure-id> update
ppt demo stat <stat-id> update
ppt demo table <table-id> update
```

Update output reports the slide and replacement that changed:

```text
Slide 1: {{fig:example}} = example.png
Slide 1: {{stat:example.count}} = 3
Slide 1: {{table:example}} = 4 rows x 2 columns
```

When the same token appears multiple times on a slide, repeated replacements
are indexed:

```text
Slide 1: {{stat:example.count}} [1/2] = 3
Slide 1: {{stat:example.count}} [2/2] = 3
```

## Export Copies

Use `--output <path>` for export copies:

```bash
ppt demo update --output exports/demo-review.pptx
```

Generated copies do not create backups and do not update the source deck's
embedded pictures, stat text, or native tables. Do not treat generated copies
as the canonical editable source. Continue editing `slides/demo/deck.pptx`.

## Backups

Write commands mutate `deck.pptx` in place by default and create a timestamped
backup under:

```text
data/ppt-artifacts/backups/
```

Backup retention is configured per presentation in `ppt.yaml` with
`backup_retention`.
