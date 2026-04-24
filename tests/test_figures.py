from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
import pytest

from pubify_ppt.discovery import load_presentation_definition
from pubify_ppt.figures import update_figures
from pubify_ppt.init import init_presentation_by_id, init_workspace


def test_update_figures_replaces_starter_anchor_with_picture(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation)

    assert outputs == (tmp_path / "slides" / "demo" / "data" / "ppt-artifacts" / "figures" / "example.png",)
    assert outputs[0].is_file()
    deck = Presentation(tmp_path / "slides" / "demo" / "deck.pptx")
    pictures = [shape for shape in deck.slides[0].shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(pictures) == 1
    assert _shape_alt_text(pictures[0]) == "{{fig:example}}"


def test_update_figures_can_update_existing_picture_anchor(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation = load_presentation_definition(tmp_path, "demo")

    update_figures(presentation)
    presentation = load_presentation_definition(tmp_path, "demo")
    outputs = update_figures(presentation)

    assert outputs[0].name == "example.png"
    deck = Presentation(tmp_path / "slides" / "demo" / "deck.pptx")
    pictures = [shape for shape in deck.slides[0].shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert len(pictures) == 1
    assert _shape_alt_text(pictures[0]) == "{{fig:example}}"


def test_update_figures_supports_targeted_multi_panel_results(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "import matplotlib.pyplot as plt",
                "from pubify_data import figure",
                "@figure",
                "def plot_pair(ctx):",
                "    fig1, ax1 = plt.subplots()",
                "    ax1.plot([1, 2], [1, 2])",
                "    fig2, ax2 = plt.subplots()",
                "    ax2.plot([1, 2], [2, 1])",
                "    return [fig1, fig2]",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    first = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    second = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(first, "{{fig:pair:1}}")
    _set_shape_alt_text(second, "{{fig:pair:2}}")
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation, figure_id="pair")

    assert [path.name for path in outputs] == ["pair_1.png", "pair_2.png"]
    deck = Presentation(presentation_root / "deck.pptx")
    pictures = [shape for shape in deck.slides[0].shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert sorted(_shape_alt_text(shape) for shape in pictures) == ["{{fig:pair:1}}", "{{fig:pair:2}}"]


def test_update_figures_requires_selected_anchor(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    deck = Presentation()
    deck.slides.add_slide(deck.slide_layouts[6])
    deck.save(tmp_path / "slides" / "demo" / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    with pytest.raises(ValueError, match="requires exactly one anchor"):
        update_figures(presentation, figure_id="example")


def _shape_alt_text(shape: object) -> str | None:
    c_nv_pr = shape._element.xpath(".//p:cNvPr")[0]
    return c_nv_pr.get("descr")


def _set_shape_alt_text(shape: object, value: str) -> None:
    c_nv_pr = shape._element.xpath(".//p:cNvPr")[0]
    c_nv_pr.set("descr", value)
