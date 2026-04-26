# AGENTS.md

## First Reads
- If this workspace includes a local `pubify-ppt` checkout, read:
  - `pubify-ppt/AGENTS.md`
  - `pubify-ppt/README.md`
- Otherwise, use the public `pubify-ppt` docs:
  - https://nvnunes.github.io/pubify-ppt/
  - https://nvnunes.github.io/pubify-ppt/architecture/
  - https://nvnunes.github.io/pubify-ppt/usage/

## Working Rules
- Keep presentation-specific analysis code, source data, and editable PowerPoint
  decks in the host presentation workspace.
- Use `ppt init <presentation-id>` to initialize or repair presentation scaffolding.
- Generated PowerPoint artifacts live under each presentation's `data/ppt-artifacts/` tree.
- Treat `deck.pptx` as the editable source deck; use `--output <path>` only for
  generated review copies.
