from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation

from pubify_ppt.backups import write_deck
from pubify_ppt.discovery import PresentationDefinition
from pubify_ppt.figures import FigureOutput, update_figures_in_deck
from pubify_ppt.runtime import check_presentation, ensure_generated_artifact_paths
from pubify_ppt.stats import StatReplacement, update_stats_in_deck
from pubify_ppt.tables import TableReplacement, update_tables_in_deck


@dataclass(frozen=True)
class PresentationUpdateResult:
    """Artifacts and replacements produced by a full presentation update."""

    figure_outputs: tuple[FigureOutput, ...]
    stat_replacements: tuple[StatReplacement, ...]
    table_replacements: tuple[TableReplacement, ...]


def update_presentation(
    presentation: PresentationDefinition,
    *,
    output: Path | None = None,
) -> PresentationUpdateResult:
    """Refresh all figures, stats, and tables, then write the deck once."""

    check_presentation(presentation)
    ensure_generated_artifact_paths(presentation)
    deck = Presentation(presentation.paths.deck_path)
    figure_result = update_figures_in_deck(presentation, deck=deck)
    stat_result = update_stats_in_deck(presentation, deck=deck)
    table_result = update_tables_in_deck(presentation, deck=deck)
    write_deck(presentation, deck, output=output)
    return PresentationUpdateResult(
        figure_outputs=figure_result.outputs,
        stat_replacements=stat_result.replacements,
        table_replacements=table_result.replacements,
    )
