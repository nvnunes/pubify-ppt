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

## Initialization

`ppt init` creates `pubify.yaml` when missing, appends a `pubify-ppt` section
when the file already exists without one, and creates the configured
`presentations_root`.

`ppt init <presentation-id>` creates the presentation folder, preserves an
existing `data/` directory or symlink, creates `data/ppt-artifacts/figures/`
and `data/ppt-artifacts/backups/`, and writes starter `ppt.yaml`, `figures.py`,
`example.csv`, and `deck.pptx` only when those files are missing.

## Discovery And Checks

`pubify-ppt` loads presentation `figures.py` through `pubify-data` using a
`PublicationAdapter` whose data root is `slides/<presentation-id>/data/`.
Workspace-relative `external_data_roots` in `ppt.yaml` resolve relative to the
workspace root.

Read-only inventory commands list decorated loaders, figures, and stats from
the loaded presentation entrypoint.

`ppt <presentation-id> check` validates config, required presentation paths,
loader data paths, `pubify-data` dependencies, figure anchors in PowerPoint alt
text, stat tokens in PowerPoint text, duplicate or unknown figure anchors,
unknown stat ids, and split-run stat tokens.
