from __future__ import annotations

from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Emu
import pubify_data

from pubify_ppt.anchors import (
    FigureAnchor,
    discover_figure_anchors_in_deck,
    set_shape_alt_text,
)
from pubify_ppt.discovery import PresentationDefinition
from pubify_ppt.runtime import check_presentation, ensure_generated_artifact_paths


EMU_PER_INCH = 914400


def update_figures(
    presentation: PresentationDefinition,
    *,
    figure_id: str | None = None,
) -> tuple[Path, ...]:
    """Render selected figures and replace their PowerPoint anchors in place.

    Phase 3 writes the source deck directly. Backup creation and ``--output``
    generated copies are owned by the later deck-write phase.
    """

    check_presentation(presentation)
    ensure_generated_artifact_paths(presentation)
    selected_ids = _selected_figure_ids(presentation, figure_id)
    if not selected_ids:
        return ()

    deck = Presentation(presentation.paths.deck_path)
    anchors = discover_figure_anchors_in_deck(deck)
    rendered = _run_selected_figures(presentation, selected_ids)
    outputs: list[Path] = []
    _clear_selected_figure_pngs(presentation, selected_ids)

    for current_id in selected_ids:
        figure_result = rendered[current_id]
        panel_bindings = _bind_anchors(current_id, figure_result, anchors)
        for panel_index, anchor in panel_bindings:
            output_path = _figure_output_path(presentation, current_id, len(figure_result.panels), panel_index)
            _render_panel_png(
                figure_result.panels[panel_index - 1].payload,
                output_path,
                width_emu=anchor.shape.width,
                height_emu=anchor.shape.height,
                dpi=presentation.config.defaults.dpi,
            )
            _replace_anchor_with_picture(anchor, output_path)
            outputs.append(output_path)

    deck.save(presentation.paths.deck_path)
    return tuple(outputs)


def _selected_figure_ids(presentation: PresentationDefinition, figure_id: str | None) -> tuple[str, ...]:
    if figure_id is None:
        return tuple(sorted(presentation.figures))
    if figure_id not in presentation.figures:
        raise KeyError(f"Unknown figure '{figure_id}'")
    return (figure_id,)


def _run_selected_figures(
    presentation: PresentationDefinition,
    selected_ids: tuple[str, ...],
) -> dict[str, pubify_data.FigureResult]:
    ctx = pubify_data.build_run_context(presentation.upstream)
    rendered: dict[str, pubify_data.FigureResult] = {}
    for current_id in selected_ids:
        ((returned_id, result),) = pubify_data.run_figures(presentation.upstream, current_id, ctx=ctx)
        rendered[returned_id] = result
    return rendered


def _bind_anchors(
    figure_id: str,
    figure_result: pubify_data.FigureResult,
    anchors: tuple[FigureAnchor, ...],
) -> tuple[tuple[int, FigureAnchor], ...]:
    figure_anchors = [anchor for anchor in anchors if anchor.figure_id == figure_id]
    panel_count = len(figure_result.panels)
    if panel_count == 1:
        scalar_anchors = [anchor for anchor in figure_anchors if anchor.panel_number is None]
        panel_anchors = [anchor for anchor in figure_anchors if anchor.panel_number == 1]
        extra_anchors = [anchor for anchor in figure_anchors if anchor.panel_number not in (None, 1)]
        if extra_anchors:
            tokens = ", ".join(anchor.token for anchor in extra_anchors)
            raise ValueError(f"Figure '{figure_id}' has extra panel anchors for a single-panel result: {tokens}")
        if len(scalar_anchors) + len(panel_anchors) != 1:
            raise ValueError(f"Figure '{figure_id}' requires exactly one anchor")
        return ((1, (scalar_anchors or panel_anchors)[0]),)

    bindings: list[tuple[int, FigureAnchor]] = []
    by_panel = {anchor.panel_number: anchor for anchor in figure_anchors if anchor.panel_number is not None}
    if any(anchor.panel_number is None for anchor in figure_anchors):
        raise ValueError(f"Figure '{figure_id}' returned multiple panels but has a scalar anchor")
    for panel_index in range(1, panel_count + 1):
        anchor = by_panel.get(panel_index)
        if anchor is None:
            raise ValueError(f"Figure '{figure_id}' is missing panel anchor {{fig:{figure_id}:{panel_index}}}")
        bindings.append((panel_index, anchor))
    extra_panels = sorted(panel for panel in by_panel if panel is not None and panel > panel_count)
    if extra_panels:
        raise ValueError(f"Figure '{figure_id}' has extra panel anchors: {extra_panels}")
    return tuple(bindings)


def _clear_selected_figure_pngs(
    presentation: PresentationDefinition,
    selected_ids: tuple[str, ...],
) -> None:
    for current_id in selected_ids:
        for path in presentation.paths.figures_root.glob(f"{current_id}*.png"):
            if path.name == f"{current_id}.png" or path.name.startswith(f"{current_id}_"):
                path.unlink()


def _figure_output_path(
    presentation: PresentationDefinition,
    figure_id: str,
    panel_count: int,
    panel_index: int,
) -> Path:
    if panel_count == 1:
        return presentation.paths.figures_root / f"{figure_id}.png"
    return presentation.paths.figures_root / f"{figure_id}_{panel_index}.png"


def _render_panel_png(
    payload: object,
    output_path: Path,
    *,
    width_emu: int,
    height_emu: int,
    dpi: int,
) -> None:
    figure = _matplotlib_figure(payload)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.set_size_inches(width_emu / EMU_PER_INCH, height_emu / EMU_PER_INCH, forward=True)
    figure.savefig(output_path, format="png", dpi=dpi)


def _matplotlib_figure(payload: object) -> object:
    if hasattr(payload, "savefig") and hasattr(payload, "set_size_inches"):
        return payload
    figure = getattr(payload, "figure", None)
    if figure is not None and hasattr(figure, "savefig") and hasattr(figure, "set_size_inches"):
        return figure
    raise ValueError("Figure panels must be Matplotlib Figure or Axes objects")


def _replace_anchor_with_picture(anchor: FigureAnchor, image_path: Path) -> None:
    shape = anchor.shape
    slide = shape.part.slide
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    picture_left, picture_top, picture_width, picture_height = _contained_geometry(
        image_path,
        left=left,
        top=top,
        width=width,
        height=height,
    )
    parent = shape._element.getparent()
    parent.remove(shape._element)
    picture = slide.shapes.add_picture(
        str(image_path),
        picture_left,
        picture_top,
        width=picture_width,
        height=picture_height,
    )
    set_shape_alt_text(picture, anchor.token)


def _contained_geometry(
    image_path: Path,
    *,
    left: int,
    top: int,
    width: int,
    height: int,
) -> tuple[Emu, Emu, Emu, Emu]:
    with Image.open(image_path) as image:
        image_width, image_height = image.size
    image_aspect = image_width / image_height
    box_aspect = width / height
    if image_aspect >= box_aspect:
        picture_width = width
        picture_height = round(width / image_aspect)
    else:
        picture_height = height
        picture_width = round(height * image_aspect)
    picture_left = left + round((width - picture_width) / 2)
    picture_top = top + round((height - picture_height) / 2)
    return Emu(picture_left), Emu(picture_top), Emu(picture_width), Emu(picture_height)
