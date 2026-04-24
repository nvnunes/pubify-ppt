from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from pptx import Presentation
from pptx.presentation import Presentation as PresentationObject
import pubify_data

from pubify_ppt.anchors import STAT_TOKEN_RE
from pubify_ppt.backups import write_deck
from pubify_ppt.discovery import PresentationDefinition
from pubify_ppt.runtime import check_presentation


@dataclass(frozen=True)
class StatReplacement:
    """One stat token replacement applied to a deck."""

    stat_id: str
    key: str | None
    token: str
    value: str


@dataclass(frozen=True)
class StatUpdateResult:
    """Stat replacements applied to an open deck."""

    deck: PresentationObject
    replacements: tuple[StatReplacement, ...]


def update_stats(
    presentation: PresentationDefinition,
    *,
    stat_id: str | None = None,
) -> tuple[StatReplacement, ...]:
    """Compute selected stats and replace matching inline PowerPoint tokens."""

    result = update_stats_in_deck(presentation, stat_id=stat_id)
    write_deck(presentation, result.deck)
    return result.replacements


def update_stats_to_output(
    presentation: PresentationDefinition,
    *,
    stat_id: str | None = None,
    output: Path | None = None,
) -> tuple[StatReplacement, ...]:
    """Compute selected stats and write the updated deck through the output policy."""

    result = update_stats_in_deck(presentation, stat_id=stat_id)
    write_deck(presentation, result.deck, output=output)
    return result.replacements


def update_stats_in_deck(
    presentation: PresentationDefinition,
    *,
    stat_id: str | None = None,
    deck: PresentationObject | None = None,
) -> StatUpdateResult:
    """Compute selected stats and replace matching tokens in an open deck."""

    check_presentation(presentation)
    selected_ids = _selected_stat_ids(presentation, stat_id)
    active_deck = deck if deck is not None else Presentation(presentation.paths.deck_path)
    if not selected_ids:
        return StatUpdateResult(active_deck, ())

    values = _run_selected_stats(presentation, selected_ids)
    replacements: list[StatReplacement] = []

    for slide in active_deck.slides:
        for shape in slide.shapes:
            text_frame = getattr(shape, "text_frame", None)
            if text_frame is None:
                continue
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.text, run_replacements = _replace_run_tokens(run.text, selected_ids=selected_ids, values=values)
                    replacements.extend(run_replacements)

    return StatUpdateResult(active_deck, tuple(replacements))


def _selected_stat_ids(presentation: PresentationDefinition, stat_id: str | None) -> tuple[str, ...]:
    if stat_id is None:
        return tuple(sorted(presentation.stats))
    if stat_id not in presentation.stats:
        raise KeyError(f"Unknown stat '{stat_id}'")
    return (stat_id,)


def _run_selected_stats(
    presentation: PresentationDefinition,
    selected_ids: tuple[str, ...],
) -> dict[tuple[str, str | None], str]:
    ctx = pubify_data.build_run_context(presentation.upstream)
    values: dict[tuple[str, str | None], str] = {}
    for current_id in selected_ids:
        (computed,) = pubify_data.run_stats(presentation.upstream, current_id, ctx=ctx)
        for value in computed.values:
            values[(computed.stat_id, value.key)] = value.value
    return values


def _replace_run_tokens(
    text: str,
    *,
    selected_ids: tuple[str, ...],
    values: dict[tuple[str, str | None], str],
) -> tuple[str, list[StatReplacement]]:
    replacements: list[StatReplacement] = []

    def replace(match: object) -> str:
        stat_id = match.group(1)
        key = match.group(2)
        if stat_id not in selected_ids:
            return match.group(0)
        value_key = (stat_id, key)
        if value_key not in values:
            if key is None:
                raise ValueError(f"Missing scalar value for stat '{stat_id}'")
            raise ValueError(f"Missing key '{key}' for stat '{stat_id}'")
        value = values[value_key]
        replacements.append(StatReplacement(stat_id, key, match.group(0), value))
        return value

    return STAT_TOKEN_RE.sub(replace, text), replacements
