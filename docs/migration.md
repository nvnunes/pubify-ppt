# Migration Guide

This guide describes a low-risk path for moving a manually maintained
PowerPoint deck into `pubify-ppt`.

## 1. Initialize The Workspace

From the host workspace root:

```bash
ppt init
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

If you already have a deck, replace `slides/demo/deck.pptx` with your existing
PowerPoint file after initialization.

## 2. Move Stable Inputs Into `data/`

Put presentation-local source data under:

```text
slides/demo/data/
```

If the data must physically live elsewhere, make `slides/demo/data` a symlink
to that location before running presentation init or updates. `pubify-ppt`
still treats the presentation-local `data/` path as canonical.

## 3. Convert Manual Figures Incrementally

Start with one manually maintained figure:

- add a loader and `@figure` function to `figures.py`
- draw a simple rectangle in PowerPoint where the figure should appear
- set the rectangle's alt text description to `{{fig:<figure_id>}}`
- run `ppt demo figure <figure_id> update`

Once the figure appears correctly, resize the generated picture in PowerPoint
as needed and rerun the same command. The deck geometry drives future render
size.

## 4. Convert Repeated Numbers To Stats

For scalar values, place:

```text
{{stat:<stat_id>}}
```

For dictionary values, place:

```text
{{stat:<stat_id>.<key>}}
```

Then run:

```bash
ppt demo stat update
```

If a stat token fails because it is split across PowerPoint runs, retype the
entire token in one edit.

## 5. Use `check` Before Larger Updates

Run:

```bash
ppt demo check
```

This catches missing data files, unknown anchors, duplicate anchors, malformed
tokens, unsupported figure anchor shapes, and split-run stat tokens before an
update writes the deck.

## 6. Keep The Source Deck Editable

Use default in-place updates during normal iteration. `pubify-ppt` creates
timestamped backups under `data/ppt-artifacts/backups/` before replacing
`deck.pptx`.

Use `--output <path>` for export copies:

```bash
ppt demo update --output exports/demo-review.pptx
```

Do not treat generated copies as the canonical editable source. Continue
editing `slides/demo/deck.pptx`.
