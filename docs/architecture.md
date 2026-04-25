# Architecture

This document is the source of truth for `pubify-ppt` package boundaries,
workspace contracts, and generated artifact ownership.

## Package Role

`pubify-ppt` is a downstream package built on `pubify-data`.

Package-owned behavior:

- `ppt` CLI command surface
- `pubify-ppt` workspace config loading
- presentation-local `ppt.yaml` config loading
- PowerPoint deck loading and writing
- PowerPoint anchor and token discovery
- figure rendering into deck image placeholders
- inline stat replacement
- native table replacement
- deck backups and generated PowerPoint artifact lifecycle

Upstream-owned behavior:

- data, external data, figure, stat, and table decorators
- publication entrypoint import and decorated-object discovery
- loader dependency validation and execution
- neutral figure, stat, and table result models
- artifact namespace helpers under downstream-supplied data roots

## Workspace Model

A host workspace is rooted by `pubify.yaml`. The `pubify-ppt` section owns the
presentation root:

```yaml
pubify-ppt:
  presentations_root: slides
```

Presentation-local data is canonical under:

```text
slides/<presentation-id>/data/
```

If a host wants data stored elsewhere, it should make that `data/` path a
filesystem redirect such as a symlink. `pubify-ppt` still treats the
presentation-local path as the data root.

## Generated Artifacts

Generated PowerPoint artifacts live under:

```text
slides/<presentation-id>/data/ppt-artifacts/
```

The editable deck remains presentation source. Rendered figures and backups are
derived artifacts owned by `pubify-ppt`.

Generated figure PNGs are intermediate artifacts. Decks embed image data, so an
opened `.pptx` does not depend on the PNG files being present.

## Initialization

`ppt init` creates `pubify.yaml` when missing, appends a `pubify-ppt` section
when the file already exists without one, and creates the configured
`presentations_root`.

`ppt init <presentation-id>` creates the presentation folder, preserves an
existing `data/` directory or symlink, creates `data/ppt-artifacts/figures/`
and `data/ppt-artifacts/backups/`, and writes starter `ppt.yaml`, `figures.py`,
`example.csv`, and `deck.pptx` only when those files are missing. The starter
entrypoint defines one example loader, figure, stat, and table; the starter
deck contains matching figure, stat, and table anchors.

## Discovery And Checks

`pubify-ppt` loads presentation `figures.py` through `pubify-data` using a
`PublicationAdapter` whose data root is `slides/<presentation-id>/data/`.
Workspace-relative `external_data_roots` in `ppt.yaml` resolve relative to the
workspace root.

Read-only inventory commands list decorated loaders, figures, stats, and tables
from the loaded presentation entrypoint.

`ppt <presentation-id> check` validates config, required presentation paths,
loader data paths, `pubify-data` dependencies, figure/table anchors in
PowerPoint alt text or exact placeholder text, stat tokens in PowerPoint text,
duplicate or unknown figure/table anchors, unknown stat ids, stat Alt Text
anchors, and ambiguous stat-managed text boxes.

See `usage.md` for user-facing anchor authoring rules and supported shape
behavior.

Source publications declared in `ppt.yaml` are code dependencies, not deck
anchor namespaces. Presentation `figures.py` owns any remapping from
`ctx.source("<source-id>")` to local figure/stat/table IDs, and the editable
PowerPoint deck references only those local IDs.

## Figure Updates

`ppt <presentation-id> figure update` renders all declared figures and replaces
matching figure anchors in the editable source deck. `ppt <presentation-id>
figure <figure-id> update` updates only one selected figure and leaves
unrelated figure anchors and PNG artifacts unchanged.

V1 figure updates render PNGs under
`data/ppt-artifacts/figures/`. Single-panel figures use
`<figure-id>.png`; multi-panel figures use `<figure-id>_<panel-number>.png`.
The current PowerPoint anchor dimensions drive the Matplotlib export size, and
inserted pictures are centered inside the anchor box with contain-fit
geometry. Inserted pictures keep the original `{{fig:...}}` token as alt text
so future updates can find them.

For bootstrap authoring, a supported shape whose visible text is exactly one
`{{fig:...}}` token is treated as an anchor. The update replaces the shape with
a picture and persists the token in alt text; after that, alt text is the
managed anchor source.

Figure updates support simple placeholder shapes and previously generated
pictures. Unsupported anchor features, such as grouping, rotation, or cropping,
are validation errors rather than silently changed layout.

## Stat Updates

`ppt <presentation-id> stat update` computes all declared stats and updates
matching stat-managed text boxes in the editable source deck. `ppt
<presentation-id> stat <stat-id> update` computes and replaces only one
selected stat, leaving unrelated stat tokens or anchored values unchanged.

New stats are authored with visible `{{stat:...}}` tokens. On first update,
`pubify-ppt` replaces the visible token with the computed plain-text value and
writes `{{stat:<id>=<previous_value>}}` or
`{{stat:<id>.<key>=<previous_value>}}` into the text box Alt Text. Later
updates read that Alt Text anchor, find the previous value exactly once in the
same text box, replace it with the new value, and update the Alt Text marker.

Valid visible stat tokens may span multiple internal PowerPoint runs. Each
stat-managed text box supports one stat token; multiple stats should use
separate text boxes.

## Table Updates

`ppt <presentation-id> table update` computes all table results with matching
anchors and updates native PowerPoint tables in the editable source deck. `ppt
<presentation-id> table <table-id> update` computes and replaces only one
selected table.

New tables are authored with a simple placeholder whose visible text or Alt
Text description is exactly `{{table:<table_id>}}`. On first update,
`pubify-ppt` replaces that placeholder with a native PowerPoint table at the
same geometry and persists the token in the table shape Alt Text. Future
updates find the table through that Alt Text.

`pubify-data` owns neutral table result normalization. `pubify-ppt` requires
one body per PowerPoint table and can read default creation-time column
headings from `metadata["columns"]`. When present, headings must be strings
and must match the computed data width. Table cells are plain text: `None`
becomes an empty string and all other values use `str(...)`.

Existing native tables are updated in place only when the PowerPoint table has
the same total row count, including the heading row, and the same column count
as the computed table. Refreshes rewrite only body cells and preserve the
PowerPoint heading row exactly as the user edited it. Dimension mismatches are
validation errors for that update; the user can either adjust the native table
size or replace it with the original `{{table:...}}` placeholder.

## Deck Writes And Backups

Write commands mutate `deck.pptx` in place by default. Before replacing the
source deck, `pubify-ppt` copies the current source deck to:

```text
data/ppt-artifacts/backups/deck-YYYYMMDD-HHMMSS.pptx
```

Deck writes save to a temporary `.pptx` beside the source deck and then replace
the source path. If the save fails after backup creation, the original source
deck and created backup are retained. Backup pruning runs only after a
successful in-place write and keeps the newest `backup_retention` backups from
`ppt.yaml`.

`--output <path>` writes a generated deck copy and does not create a backup or
mutate `deck.pptx`. Generated figure PNGs are still refreshed under
`data/ppt-artifacts/figures/` when figure rendering is part of the command.

`ppt <presentation-id> update` refreshes all figures, stats, and tables in one
open deck and performs one final write through the same backup/output policy.
