# pubify-ppt Foundation Plan

## Summary

`pubify-ppt` will be a sibling downstream package built on the `pubify-data`
runtime. It will use the same `figures.py` authoring model, loader execution,
publication-local data-root behavior, and artifact-namespace foundation as
other `pubify-data` downstreams, while owning only PowerPoint-specific
behavior: `.pptx` deck loading, anchor discovery, artifact insertion, native
table updates, and deck writing.

V1 supports figures, stats, and native PowerPoint tables.

## Current Decisions

- Package/repo: new sibling package `pubify-ppt`.
- CLI: `ppt`.
- Workspace root: `slides`.
- Source deck filename: `deck.pptx`.
- Runtime foundation: `pubify-data`.
- PowerPoint backend: `python-pptx`.
- Figure anchors: PowerPoint alt text.
- Stat anchors: visible text tokens for first update, then text box Alt Text
  markers that store the previous rendered value.
- Table anchors: visible text or Alt Text tokens for first update, then native
  PowerPoint table Alt Text.
- Table headings: optional `metadata["columns"]` defaults are used only when a
  table is first created; later updates preserve the PowerPoint heading row.
- Table compatibility: updates mutate existing table cells only when total row
  count and column count match the computed table.
- Default output policy: update `deck.pptx` in place.
- Optional output policy: support explicit `--output <path>` for generated
  copies.
- Backup retention default: keep the latest 5 in-place deck backups.
- Canonical generated artifacts: presentation-local `data/ppt-artifacts/`.
- External storage policy: symlink the presentation-local `data/` directory
  when data should physically live elsewhere.
- Tables: native PowerPoint tables with optional creation-time headings from
  `metadata["columns"]`.

## Proposed Folder Structure

Workspace config in `pubify.yaml`:

```yaml
pubify-ppt:
  presentations_root: slides
```

Presentation folder:

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

Presentation config in `ppt.yaml`:

```yaml
deck: deck.pptx
backup_retention: 5
defaults:
  image_format: png
  dpi: 200
  fit: contain
external_data_roots: {}
```

V1 `ppt.yaml` fields:

- `deck`: presentation-local path to the editable source deck. Defaults to
  `deck.pptx`.
- `backup_retention`: number of in-place deck backups to keep. Defaults to 5.
- `defaults.image_format`: generated figure image format. V1 supports `png`.
- `defaults.dpi`: DPI-equivalent render scale at the current anchor size.
  Defaults to 200.
- `defaults.fit`: image placement mode inside the anchor box. V1 supports
  `contain`.
- `external_data_roots`: optional mapping for `@external_data(...)` loaders,
  resolved the same way as other `pubify-data` downstreams.

Reserved future fields:

- `outputs`: optional named generated-copy destinations if the single
  `--output <path>` flag becomes too limiting.

## Package Structure

Repository layout:

```text
pubify-ppt/
  AGENTS.md
  README.md
  pyproject.toml
  docs/
    architecture.md
    development.md
    index.md
    plan.md
    testing.md
  src/
    pubify_ppt/
      __init__.py
      cli.py
      config.py
      discovery.py
      runtime.py
      anchors.py
      figures.py
      stats.py
      tables.py
      backups.py
      assets/
        init/
  tests/
    conftest.py
    test_cli.py
    test_config.py
    test_runtime.py
    test_anchors.py
    test_tables.py
```

Initial dependencies:

- `pubify-data`
- `python-pptx`
- `matplotlib`
- `numpy`
- `Pillow` for image-size inspection during contain-fit placement
- `pytest` and `mkdocs` in the dev extra

Development environment:

- Create a repo-local `./.conda` environment, matching the `pubify-data` and
  `pubify-pubs` workflow.
- Use `./.conda/bin/python -m pytest` for tests.
- Add MkDocs early so docs and the public API surface stay aligned as the
  package grows.

`pubify-ppt` should always treat `slides/<presentation-id>/data` as the
canonical presentation data root. If a host wants data outside the presentation
folder, it should create `slides/<presentation-id>/data` as a symlink before
running presentation init.

Generated PowerPoint artifacts live under the fixed `ppt-artifacts` namespace
inside that data root. The source deck remains presentation source; rendered
figures and backups are derived artifacts.

Folder-structure decisions:

- Use `slides` as the default workspace root because it is short, concrete, and
  parallel to `papers` without overloading the PowerPoint file format name.
- Use `deck.pptx` as the default source deck name because it describes the
  editable presentation source rather than implying a disposable template.
- Use `ppt` as the CLI name to mirror the `pubs` command from `pubify-pubs`.
- Keep generated artifacts under `data/ppt-artifacts` for v1. Revisit
  configurability only after the in-place update, backup, and figure lifecycle
  is stable.

## Authoring Model

Presentation `figures.py` imports reusable decorators from `pubify_data`:

```python
from pubify_data import data, external_data, figure, stat
```

Data loaders use `@data(...)` and `@external_data(...)` with the same semantics
as other `pubify-data` downstreams. `@data(...)` paths resolve under the
presentation-local `data/` root.

Figures return values compatible with `pubify-data` neutral figure results. V1
should support Matplotlib figures rendered to images and embedded into the deck.
Multi-panel figure results are supported as one coordinated `@figure` method
that emits multiple related visual outputs. PowerPoint owns the visual layout:
each panel is inserted into its own explicit anchor box rather than being
automatically composed by `pubify-ppt`.

Stats return scalar values or dictionaries. New stats are authored with visible
tokens in PowerPoint text boxes:

```text
{{stat:<stat_id>}}
{{stat:<stat_id>.<key>}}
```

On first update, the visible token is replaced and the text box Alt Text is set
to `{{stat:<stat_id>=<previous_value>}}` or
`{{stat:<stat_id>.<key>=<previous_value>}}`. Later updates use the Alt Text
marker to find the previous visible value and replace it safely.

## PowerPoint Anchoring

Figure anchors use PowerPoint alt text tokens:

```text
{{fig:<figure_id>}}
{{fig:<figure_id>:<panel_number>}}
```

V1 anchor rules:

- Read and write the shape description field exposed by `python-pptx` as the
  anchor alt text.
- The `{{...}}` wrapper is the managed-token marker.
- A scalar/single-panel figure uses `{{fig:<figure_id>}}`.
- A multi-panel figure uses one anchor per panel:
  `{{fig:<figure_id>:<panel_number>}}`, with one-based panel numbers.
- Panel anchors mean "the Nth visual output from one coordinated figure
  method"; they do not imply equal sizing or automatic grid layout.
- Missing anchors are errors during `figure update` and `update`.
- Duplicate anchors for the same figure or panel are errors.
- Extra anchors for unknown figures are errors during `check` and `update`.
- Anchors for figures not selected in a targeted update are left unchanged.
- Replacement preserves the anchor shape's slide, left, top, width, and height.
- Replacement images must keep the same `{{fig:...}}` token in their alt text
  so future updates can find them.
- On later updates, existing generated images are treated as the anchor. Their
  current width and height define the next export size.
- V1 does not promise preservation of z-order, grouping, rotation, crop, or
  animation. If those are present on an anchor shape, `update` should fail with
  a clear unsupported-anchor error rather than silently producing a different
  layout.

Supported v1 figure anchor shapes:

- plain rectangle/freeform placeholder shapes with the figure token in alt text
- existing picture shapes inserted by a previous `ppt update`

Unsupported v1 figure anchor shapes:

- grouped shapes
- rotated shapes
- cropped picture shapes
- animated shapes
- shapes inside tables, charts, SmartArt, or embedded objects

Unsupported anchors should produce an error naming the slide number, figure
token, and unsupported feature. Users can resolve the error by replacing the
anchor with a simple rectangle or an existing generated picture.

The update flow should locate matching shapes, render each figure panel under
`data/ppt-artifacts/figures/`, delete or replace the anchor shape, and insert
the rendered figure image into the same slide geometry.

Stats are updated by scanning text shapes for new visible `{{stat:...}}`
tokens and by scanning text box Alt Text for previously rendered stat markers.

V1 stat replacement rules:

- Supported token forms are `{{stat:<stat_id>}}` and
  `{{stat:<stat_id>.<key>}}`.
- Persisted rendered marker forms are `{{stat:<stat_id>=<previous_value>}}`
  and `{{stat:<stat_id>.<key>=<previous_value>}}` in text box Alt Text.
- Scalar stats use `{{stat:<stat_id>}}`.
- Dictionary stats use `{{stat:<stat_id>.<key>}}`, with keys normalized the
  same way as the TeX macro key suffix in `pubify-pubs`.
- Missing stat ids and missing dictionary keys are errors.
- Repeated tokens are allowed when each occurrence lives in its own text box.
- Replacement values are plain text.
- Replacement should preserve the formatting of the run containing the token
  when possible.
- Valid visible tokens may be split across multiple internal PowerPoint runs.
- Each stat-managed text box supports one stat token. Multiple stats in one
  text box are errors.
- On later updates, the previous rendered value must appear exactly once in the
  text box. If it is missing or ambiguous, the user should restore the previous
  value or reinsert the original `{{stat:...}}` token.
- Text outside the token is preserved on first update.

By default, updates write back to `deck.pptx` in place after creating a backup.
`--output <path>` writes a generated copy instead and does not mutate
`deck.pptx`.

## Figure Rendering Policy

V1 renders figures as PNG files under `data/ppt-artifacts/figures/`.

- Image format comes from `ppt.yaml` `defaults.image_format`; v1 supports PNG.
- Render scale comes from `ppt.yaml` `defaults.dpi`; the default is 200
  DPI-equivalent at the current anchor's physical size.
- The PowerPoint anchor box is authoritative for placement.
- The current anchor width and height in the `.pptx` drive the next figure
  export size. Users resize the image or placeholder in PowerPoint, then the
  next update re-renders at that size.
- Inserted images preserve their aspect ratio and are contained inside the
  anchor box when `defaults.fit` is `contain`.
- If the rendered image aspect ratio differs from the anchor box, center it
  within the anchor bounds.
- Do not stretch or crop images in v1.
- Multi-panel figure results render one PNG per panel and require one explicit
  `{{fig:<figure_id>:<panel_number>}}` anchor per panel.
- A multi-panel result without the required panel anchors is an error.
- A single-panel result may use either `{{fig:<figure_id>}}` or
  `{{fig:<figure_id>:1}}`, but `{{fig:<figure_id>}}` is preferred.
- Targeted figure updates only regenerate selected figure images and leave
  unrelated figure anchors and generated image files unchanged.

## Artifact Lifecycle And Backups

Generated artifacts are authoritative snapshots under `data/ppt-artifacts/`.

- Default deck update target: source `deck.pptx` in place.
- Optional output deck path: caller-supplied `--output <path>`.
- Generated figure image names:
  - single-panel figure: `<figure_id>.png`
  - multi-panel figure: `<figure_id>_<panel_number>.png`
- Full `ppt <presentation-id> update` refreshes all figures and stats and
  writes the deck.
- Full `figure update` clears stale generated PNGs for known figures before
  rewriting selected outputs.
- Targeted `figure <figure-id> update` rewrites only that figure's PNGs and
  leaves unrelated generated PNGs untouched.
- `stat update` rewrites only the deck; it does not remove or rewrite figure
  PNGs.
- Generated figure PNGs are intermediate artifacts used while updating the
  deck. The deck embeds image data and should not depend on external PNG files
  at presentation-open time.
- In-place writes create a backup before replacing `deck.pptx`.
- Backup path format:
  `data/ppt-artifacts/backups/deck-YYYYMMDD-HHMMSS.pptx`.
- Backup retention defaults to 5 and is configured per presentation in
  `ppt.yaml` with `backup_retention`.
- Delete older backups only after a successful in-place write.
- If an update fails after backup creation, keep the backup.
- In-place writes should be atomic enough for normal local use: save to a temp
  file beside `deck.pptx`, then replace `deck.pptx`.
- `--output <path>` writes a generated copy and does not create a backup unless
  the output path is the source deck path.
## CLI Behavior

Initial command surface:

```text
ppt list
ppt init
ppt init <presentation-id>
ppt <presentation-id> check
ppt <presentation-id> data list
ppt <presentation-id> figure list
ppt <presentation-id> figure update [--output <path>]
ppt <presentation-id> figure <figure-id> update [--output <path>]
ppt <presentation-id> stat list
ppt <presentation-id> stat update [--output <path>]
ppt <presentation-id> stat <stat-id> update [--output <path>]
ppt <presentation-id> update [--output <path>]
```

Use `pubify_data.PublicationAdapter`, `WorkspaceAdapter`, runtime helpers, and
command registry patterns. Use `pubify-data` artifact namespace helpers to
resolve `ppt-artifacts`; keep all `python-pptx` logic inside `pubify-ppt`.

CLI write semantics:

- `ppt list` and all `list` subcommands are read-only.
- `ppt <presentation-id> check` validates config, data roots, anchors, tokens,
  and unsupported shapes without writing the deck.
- `ppt <presentation-id> update` refreshes all figures and stats and writes the
  deck in place by default.
- `ppt <presentation-id> figure update` refreshes all figure anchors and writes
  the deck in place by default.
- `ppt <presentation-id> figure <figure-id> update` refreshes only that figure's
  anchors and writes the deck in place by default.
- `ppt <presentation-id> stat update` refreshes all stat-managed text boxes and
  writes the deck in place by default.
- `ppt <presentation-id> stat <stat-id> update` refreshes only text boxes for
  that stat and writes the deck in place by default.
- `--output <path>` is accepted on write commands and writes a generated copy
  instead of mutating the source deck.
- `--output <path>` never updates the source deck's anchor/image state, so it is
  intended for export/checking rather than normal iteration.

## Init Contract

`ppt init` for a workspace should:

- create `pubify.yaml` if missing
- add a `pubify-ppt` section with `presentations_root: slides`
- create the configured `presentations_root`
- not create presentation-local data roots
- reject `--force`

Target workspace config:

```yaml
pubify-ppt:
  presentations_root: slides
```

`ppt init <presentation-id>` should:

- create `slides/<presentation-id>/`
- create `slides/<presentation-id>/data/` if missing
- preserve `slides/<presentation-id>/data/` when it already exists as a
  directory or symlink
- create `data/ppt-artifacts/`
- create `data/ppt-artifacts/figures/`
- create `data/ppt-artifacts/backups/`
- create `figures.py` when missing
- create `ppt.yaml` when missing
- create a minimal starter `deck.pptx` when missing
- never overwrite existing `deck.pptx`, `figures.py`, or `ppt.yaml`

The starter deck should be created with `python-pptx`, not a bundled binary
deck asset. It should contain one starter slide with:

- one figure placeholder shape whose alt text is `{{fig:example}}`
- one text box containing `{{stat:example.count}}`
- one table placeholder shape containing `{{table:example}}`

The starter `figures.py` should define matching example data, one example
figure, one dictionary stat with a `count` key, and one example table. This
makes `ppt init demo && ppt demo update` visibly useful and gives tests an
end-to-end fixture. The source deck remains user-editable presentation source.

## Test Plan

- Workspace and presentation config parsing.
- `pubify-data` adapter construction for `slides/<id>/figures.py`.
- Presentation-local data root resolution, including when `data/` is a symlink.
- `ppt-artifacts` namespace creation under the presentation-local data root.
- Starter `ppt init <id>` creates `deck.pptx` with figure, stat, and table
  anchors using `python-pptx`.
- Starter `ppt init <id>` creates matching example loaders, figure, stat, and
  table.
- Figure anchor detection from PowerPoint alt text.
- Generated deck update replaces figure anchors with images at matching
  geometry.
- Rendered figure images are written under `data/ppt-artifacts/figures/`.
- Stat replacement for scalar and dictionary stats, including first-update
  visible tokens and later Alt Text anchored values.
- In-place updates create timestamped backups and respect `backup_retention`.
- `--output <path>` writes a generated copy without mutating the source deck.
- Integration test with a tiny `.pptx` deck, one loader, one figure, and two
  stats.
- Backup retention test with `backup_retention: 2`.
- Atomic-write test verifies failed updates keep the original deck and any
  created backup.

## Assumptions

- `pubify-ppt` remains independent from `pubify-pubs`.
- `pubify-data` remains presentation-format agnostic.
- `pubify-data` already provides artifact namespace helpers under a supplied
  publication data root.
- PowerPoint iteration requires mutating the source deck by default because the
  `.pptx` file is both source and layout editor.

## Phase Plan

### Phase 0: Package Skeleton

- Create package metadata, source layout, tests, docs, and local development
  environment.
- Add `ppt` console script.
- Add config loading for `pubify-ppt.presentations_root`.

### Phase 1: Workspace And Presentation Init

- Implement `ppt init`.
- Implement `ppt init <presentation-id>`.
- Create `ppt.yaml`, `figures.py`, starter `deck.pptx`, presentation-local
  `data/`, and `data/ppt-artifacts/{figures,backups}/`.
- Starter deck contains `{{fig:example}}` and `{{stat:example.count}}`.
- Starter `figures.py` contains matching example data, figure, and stat.

### Phase 2: Data And Inventory Commands

- Load presentations through `pubify-data.PublicationAdapter`.
- Implement `ppt list`.
- Implement `data list`, `figure list`, and `stat list`.
- Validate presentation-local data roots, including symlinked `data/`.
- Implement `ppt <presentation-id> check` for config, anchor, token, and
  unsupported-shape validation.

### Phase 3: Figures

- Discover figure anchors from shape alt text.
- Render figures to PNG under `data/ppt-artifacts/figures/`.
- Replace figure anchors/images in place while preserving current geometry.
- Preserve figure token alt text on inserted images.
- Support targeted figure updates.
- Support multi-panel figure results through explicit panel anchors.

### Phase 4: Stats

- Discover visible stat tokens in text shapes.
- Persist rendered stat anchors in text box Alt Text.
- Replace scalar and dictionary stat tokens and anchored rendered values.
- Preserve formatting for single-run tokens.
- Accept valid split-run tokens.
- Fail clearly when a stat-managed text box contains multiple stat tokens or
  when the previous rendered value cannot be found exactly once.
- Support targeted stat updates.

### Phase 5: In-Place Deck Writes And Backups

- Implement default in-place deck update.
- Create timestamped backups before writes.
- Enforce `backup_retention`, default 5.
- Save through temp-file replacement.
- Implement `--output <path>` for generated copies.

### Phase 6: Polish And Docs

- Add README, architecture docs, and migration guidance for teams moving from
  manually maintained decks.
- Document supported/unsupported PowerPoint shape behavior.

### Phase 7: Tables

- Add native PowerPoint table support using `python-pptx`.
- Table anchors use visible placeholder text for first update or persisted Alt
  Text after creation:

```text
{{table:<table_id>}}
```

- Replace a simple placeholder shape with a PowerPoint table sized to the
  anchor box.
- Use `pubify-data` table results as the neutral input.
- Read default column headings from `TableResult.metadata["columns"]` only
  when creating a native table from a token.
- Update existing native table cells only when the PowerPoint table has the
  same total row count and column count as the computed table.
- Preserve existing native table heading cells on refresh.
- On dimension mismatch, tell the user to adjust the native table size or
  replace it with the original `{{table:...}}` token and rerun.
- V1 table phase supports text-only cells, because PowerPoint table cells do
  not hold images or nested shapes.
