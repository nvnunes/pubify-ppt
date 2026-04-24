"""PowerPoint anchor and inline-token discovery for ``pubify-ppt``."""

from __future__ import annotations

from dataclasses import dataclass
import re

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


FIGURE_TOKEN_RE = re.compile(r"^\{\{fig:([A-Za-z0-9_.-]+)(?::([1-9][0-9]*))?\}\}$")
STAT_TOKEN_RE = re.compile(r"\{\{stat:([A-Za-z0-9_-]+)(?:\.([A-Za-z0-9_.-]+))?\}\}")
MANAGED_TOKEN_RE = re.compile(r"\{\{([^{}]+)\}\}")


@dataclass(frozen=True)
class FigureAnchor:
    """One PowerPoint shape carrying a managed figure token."""

    slide_number: int
    figure_id: str
    panel_number: int | None
    token: str
    shape: object


@dataclass(frozen=True)
class StatToken:
    """One inline stat token discovered in PowerPoint text."""

    slide_number: int
    stat_id: str
    key: str | None
    token: str


def discover_figure_anchors(deck_path: object) -> tuple[FigureAnchor, ...]:
    """Discover figure anchors from PowerPoint shape alt text."""

    deck = Presentation(deck_path)
    anchors: list[FigureAnchor] = []
    for slide_number, slide in enumerate(deck.slides, start=1):
        for shape in slide.shapes:
            alt_text = shape_alt_text(shape)
            if not alt_text:
                continue
            match = FIGURE_TOKEN_RE.fullmatch(alt_text.strip())
            if match is None:
                continue
            panel = int(match.group(2)) if match.group(2) is not None else None
            anchors.append(FigureAnchor(slide_number, match.group(1), panel, match.group(0), shape))
    return tuple(anchors)


def discover_stat_tokens(deck_path: object) -> tuple[StatToken, ...]:
    """Discover supported stat tokens from PowerPoint text shapes."""

    deck = Presentation(deck_path)
    tokens: list[StatToken] = []
    for slide_number, slide in enumerate(deck.slides, start=1):
        for shape in slide.shapes:
            text_frame = getattr(shape, "text_frame", None)
            if text_frame is None:
                continue
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    for match in STAT_TOKEN_RE.finditer(run.text):
                        tokens.append(StatToken(slide_number, match.group(1), match.group(2), match.group(0)))
    return tuple(tokens)


def validate_deck_anchors(deck_path: object, *, figure_ids: set[str], stat_ids: set[str]) -> list[str]:
    """Return static deck-anchor validation errors without mutating the deck."""

    deck = Presentation(deck_path)
    errors: list[str] = []
    figure_keys: dict[tuple[str, int | None], FigureAnchor] = {}
    for slide_number, slide in enumerate(deck.slides, start=1):
        for shape in slide.shapes:
            alt_text = shape_alt_text(shape)
            if alt_text:
                _validate_alt_text(
                    errors,
                    figure_keys,
                    shape,
                    slide_number=slide_number,
                    alt_text=alt_text.strip(),
                    figure_ids=figure_ids,
                )
            _validate_text_shape(errors, shape, slide_number=slide_number, stat_ids=stat_ids)
    return errors


def shape_alt_text(shape: object) -> str | None:
    """Return a shape's description alt text as exposed in the Open XML tree."""

    matches = shape._element.xpath(".//p:cNvPr")
    if not matches:
        return None
    return matches[0].get("descr")


def _validate_alt_text(
    errors: list[str],
    figure_keys: dict[tuple[str, int | None], FigureAnchor],
    shape: object,
    *,
    slide_number: int,
    alt_text: str,
    figure_ids: set[str],
) -> None:
    managed_match = MANAGED_TOKEN_RE.fullmatch(alt_text)
    if managed_match is None:
        return
    figure_match = FIGURE_TOKEN_RE.fullmatch(alt_text)
    if figure_match is None:
        if managed_match.group(1).startswith("fig:"):
            errors.append(f"Slide {slide_number}: malformed figure anchor token {alt_text!r}")
        return

    figure_id = figure_match.group(1)
    panel_number = int(figure_match.group(2)) if figure_match.group(2) is not None else None
    token = figure_match.group(0)
    if figure_id not in figure_ids:
        errors.append(f"Slide {slide_number}: unknown figure anchor {token}")

    key = (figure_id, panel_number)
    if key in figure_keys:
        errors.append(f"Slide {slide_number}: duplicate figure anchor {token}")
    else:
        figure_keys[key] = FigureAnchor(slide_number, figure_id, panel_number, token, shape)

    unsupported = _unsupported_figure_anchor_features(shape)
    if unsupported:
        errors.append(f"Slide {slide_number}: unsupported figure anchor {token}: {unsupported}")


def _validate_text_shape(
    errors: list[str],
    shape: object,
    *,
    slide_number: int,
    stat_ids: set[str],
) -> None:
    text_frame = getattr(shape, "text_frame", None)
    if text_frame is None:
        return
    for paragraph in text_frame.paragraphs:
        paragraph_text = "".join(run.text for run in paragraph.runs)
        for managed_match in MANAGED_TOKEN_RE.finditer(paragraph_text):
            token = managed_match.group(0)
            if not managed_match.group(1).startswith("stat:"):
                continue
            stat_match = STAT_TOKEN_RE.fullmatch(token)
            if stat_match is None:
                errors.append(f"Slide {slide_number}: malformed stat token {token!r}")
                continue
            if stat_match.group(1) not in stat_ids:
                errors.append(f"Slide {slide_number}: unknown stat token {token}")
            if not any(token in run.text for run in paragraph.runs):
                errors.append(
                    f"Slide {slide_number}: stat token {token} is split across PowerPoint runs; "
                    "retype the token in one text run"
                )


def _unsupported_figure_anchor_features(shape: object) -> str | None:
    features: list[str] = []
    if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
        features.append("grouped shape")
    rotation = getattr(shape, "rotation", 0)
    if rotation not in (0, 0.0):
        features.append("rotation")
    if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
        crop_values = (
            getattr(shape, "crop_left", 0),
            getattr(shape, "crop_right", 0),
            getattr(shape, "crop_top", 0),
            getattr(shape, "crop_bottom", 0),
        )
        if any(value not in (0, 0.0) for value in crop_values):
            features.append("crop")
    return ", ".join(features) if features else None
