# pubify-ppt

`pubify-ppt` is the PowerPoint-oriented downstream package for `pubify-data`.

The package skeleton, initialization workflow, inventory commands, static
presentation checks, figure updates, stat updates, full updates,
generated-copy output, and in-place deck backups are in place. Use the
implementation plan for the current sequencing and `architecture.md` for the
stable ownership model as it is implemented.

## Starting Points

- `architecture.md`: package boundaries, workspace model, and generated
  artifact ownership
- `powerpoint.md`: deck authoring rules, anchor support, stat tokens, and
  output behavior
- `migration.md`: practical path for converting a manual deck incrementally
- `development.md`: local environment and daily commands
- `testing.md`: canonical verification commands
- `plan.md`: phased implementation plan
