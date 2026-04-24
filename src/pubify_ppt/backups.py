from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import os
import shutil
import tempfile

from pptx.presentation import Presentation as PresentationObject

from pubify_ppt.discovery import PresentationDefinition
from pubify_ppt.runtime import ensure_generated_artifact_paths


@dataclass(frozen=True)
class DeckWriteResult:
    """Result of writing a generated deck."""

    deck_path: Path
    backup_path: Path | None = None


def write_deck(
    presentation: PresentationDefinition,
    deck: PresentationObject,
    *,
    output: Path | None = None,
) -> DeckWriteResult:
    """Write a deck using the package backup and output policy."""

    source_path = presentation.paths.deck_path
    output_path = _resolve_output_path(presentation, output)
    if output_path == source_path:
        return _write_source_deck(presentation, deck)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    deck.save(output_path)
    return DeckWriteResult(deck_path=output_path)


def _write_source_deck(
    presentation: PresentationDefinition,
    deck: PresentationObject,
) -> DeckWriteResult:
    ensure_generated_artifact_paths(presentation)
    source_path = presentation.paths.deck_path
    backup_path = create_deck_backup(presentation)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{source_path.stem}-",
        suffix=source_path.suffix,
        dir=source_path.parent,
    )
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        deck.save(temp_path)
        os.replace(temp_path, source_path)
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise
    prune_deck_backups(presentation)
    return DeckWriteResult(deck_path=source_path, backup_path=backup_path)


def create_deck_backup(presentation: PresentationDefinition) -> Path:
    """Copy the current source deck into the backup artifact directory."""

    backup_path = _next_backup_path(presentation)
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(presentation.paths.deck_path, backup_path)
    return backup_path


def prune_deck_backups(presentation: PresentationDefinition) -> None:
    """Retain only the configured number of newest deck backups."""

    retention = presentation.config.backup_retention
    if retention < 0:
        return
    backups = sorted(
        presentation.paths.backups_root.glob("deck-*.pptx"),
        key=lambda path: (path.stat().st_mtime_ns, path.name),
    )
    stale = backups[:-retention] if retention else backups
    for path in stale:
        path.unlink()


def _next_backup_path(presentation: PresentationDefinition) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = presentation.paths.backups_root / f"deck-{timestamp}.pptx"
    suffix = 1
    while candidate.exists():
        candidate = presentation.paths.backups_root / f"deck-{timestamp}-{suffix}.pptx"
        suffix += 1
    return candidate


def _resolve_output_path(presentation: PresentationDefinition, output: Path | None) -> Path:
    if output is None:
        return presentation.paths.deck_path
    output_path = output.expanduser()
    if not output_path.is_absolute():
        output_path = (Path.cwd() / output_path).resolve()
    return output_path
