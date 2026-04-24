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

Stats are inline text tokens in PowerPoint text boxes:

```text
{{stat:<stat_id>}}
{{stat:<stat_id>.<key>}}
```

Scalar stats use `{{stat:<stat_id>}}`. Dictionary stats use
`{{stat:<stat_id>.<key>}}`.

When a stat token is fully contained in one PowerPoint text run, replacement
preserves that run's formatting. If PowerPoint splits a token across multiple
runs, `check` and update commands fail. Retype the full token in one operation
inside PowerPoint to put it back into a single run.

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
