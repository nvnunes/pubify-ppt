# PowerPoint Authoring

This document describes the supported PowerPoint authoring model for
`pubify-ppt` decks.

## Deck Source

Each presentation owns an editable source deck:

```text
slides/<presentation-id>/deck.pptx
```

The `.pptx` file is both source and layout editor. Resize placeholders or
previously generated pictures in PowerPoint, then rerun `ppt <presentation-id>
update` to regenerate content at the current deck geometry.

## Figure Anchors

Figure anchors use PowerPoint alt text. In PowerPoint, set the shape's
description alt text to one of these tokens:

```text
{{fig:<figure_id>}}
{{fig:<figure_id>:<panel_number>}}
```

Use `{{fig:<figure_id>}}` for single-panel figures. Multi-panel figures require
one explicit anchor per panel, using one-based panel numbers such as
`{{fig:comparison:1}}` and `{{fig:comparison:2}}`.

Figure updates replace the anchor shape with an embedded PNG and preserve the
same figure token in the inserted picture's alt text. Future updates use that
picture as the next anchor.

Figure anchors must reference presentation-local figure IDs defined in the
presentation's `figures.py`. Source publication IDs are not valid PowerPoint
anchors; expose reused source outputs through local wrapper functions instead.

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
from pubify_data import figure, stat
from pubify_ppt import FigureResult, StatResult


@figure
def plot_training_fov_ee(ctx):
    panel = ctx.source("ao4elt8").figure("training_compare_fov_ee").panel(1)
    return FigureResult(panel)


@stat
def compute_training_count(ctx):
    source_stat = ctx.source("ao4elt8").stat("training_count")
    return StatResult(source_stat.values[0].value)
```

The PowerPoint deck then uses only the local IDs:

```text
{{fig:training_fov_ee}}
{{stat:training_count}}
```

This keeps the editable deck independent of source publication internals such
as paper IDs, source figure names, and panel numbering. Rename or adapt source
outputs in `figures.py`, not in PowerPoint alt text.

## Supported Figure Anchor Shapes

Supported in the current implementation:

- plain rectangle or freeform placeholder shapes with a valid `{{fig:...}}`
  token in alt text
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

## Stat Tokens

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

## Output And Backups

Write commands mutate `deck.pptx` in place by default and create a backup under:

```text
data/ppt-artifacts/backups/
```

Use `--output <path>` when you want a generated copy instead of mutating the
source deck:

```bash
ppt demo update --output exports/demo.pptx
```

Generated copies do not create backups and do not update the source deck's
embedded pictures or stat text.
