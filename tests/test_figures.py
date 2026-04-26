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


def test_update_figures_bootstraps_anchor_from_exact_shape_text(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    shape.text = "{{fig:example}}"
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation)

    assert [path.name for path in outputs] == ["example.png"]
    deck = Presentation(presentation_root / "deck.pptx")
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


def test_update_figures_prepares_axes_payload_without_mutating_source(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "import matplotlib.pyplot as plt",
                "from pubify_data import figure",
                "from pubify_ppt import FigureResult",
                "@figure",
                "def plot_axes(ctx):",
                "    fig, axs = plt.subplots(1, 2)",
                "    axs[0].set_xlabel('left')",
                "    axs[1].set_xlabel('right')",
                "    axs[0].plot([1, 2], [1, 2])",
                "    axs[1].plot([1, 2], [2, 1])",
                "    return FigureResult(axs[1], metadata={'hide_labels': True})",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(shape, "{{fig:axes}}")
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation, figure_id="axes")

    assert [path.name for path in outputs] == ["axes.png"]
    assert outputs[0].is_file()


def test_targeted_figure_update_preserves_prefix_matching_png(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    figures_root = presentation_root / "data" / "ppt-artifacts" / "figures"
    figures_root.mkdir(parents=True, exist_ok=True)
    (figures_root / "foo_1.png").write_text("stale panel", encoding="utf-8")
    (figures_root / "foo_bar.png").write_text("separate figure", encoding="utf-8")
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "import matplotlib.pyplot as plt",
                "from pubify_data import figure",
                "@figure",
                "def plot_foo(ctx):",
                "    fig, ax = plt.subplots()",
                "    ax.plot([1, 2], [1, 2])",
                "    return fig",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(shape, "{{fig:foo}}")
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation, figure_id="foo")

    assert [path.name for path in outputs] == ["foo.png"]
    assert not (figures_root / "foo_1.png").exists()
    assert (figures_root / "foo_bar.png").read_text(encoding="utf-8") == "separate figure"


def test_update_figures_rejects_non_integer_metadata_dpi(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "import matplotlib.pyplot as plt",
                "from pubify_data import figure",
                "from pubify_ppt import FigureResult",
                "@figure",
                "def plot_bad_dpi(ctx):",
                "    fig, ax = plt.subplots()",
                "    ax.plot([1, 2], [1, 2])",
                "    return FigureResult(fig, metadata={'dpi': '300'})",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(shape, "{{fig:bad_dpi}}")
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    with pytest.raises(ValueError, match="dpi must be a positive integer"):
        update_figures(presentation, figure_id="bad_dpi")


def test_update_figures_renders_local_wrapper_around_source_publication_figure(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    source_root = tmp_path / "papers" / "ao4elt8"
    source_data = source_root / "data"
    source_data.mkdir(parents=True)
    (source_data / "points.csv").write_text("x,y\n1,1\n2,4\n", encoding="utf-8")
    (source_root / "figures.py").write_text(
        "\n".join(
            [
                "import csv",
                "import matplotlib.pyplot as plt",
                "from pubify_data import data, figure",
                "@data('points.csv')",
                "def load_points(ctx, path):",
                "    with path.open(newline='', encoding='utf-8') as handle:",
                "        return [(float(row['x']), float(row['y'])) for row in csv.DictReader(handle)]",
                "@figure",
                "def plot_source(ctx, points):",
                "    fig, ax = plt.subplots()",
                "    ax.plot([point[0] for point in points], [point[1] for point in points])",
                "    return fig",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    presentation_root = tmp_path / "slides" / "demo"
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "from pubify_data import figure",
                "from pubify_ppt import FigureResult",
                "@figure",
                "def plot_source_panel(ctx):",
                "    return FigureResult(ctx.source('ao4elt8').figure('source').panel(1))",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    (presentation_root / "ppt.yaml").write_text(
        "\n".join(
            [
                "deck: deck.pptx",
                "backup_retention: 5",
                "defaults:",
                "  image_format: png",
                "  dpi: 200",
                "  fit: contain",
                "external_data_roots:",
                "sources:",
                "  ao4elt8: papers/ao4elt8",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(shape, "{{fig:source_panel}}")
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation, figure_id="source_panel")

    assert [path.name for path in outputs] == ["source_panel.png"]
    deck = Presentation(presentation_root / "deck.pptx")
    pictures = [shape for shape in deck.slides[0].shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    assert [_shape_alt_text(shape) for shape in pictures] == ["{{fig:source_panel}}"]


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
