from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from PIL import Image
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


def test_update_figures_patches_only_figure_slide_rels_and_media_parts(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    _write_single_line_figure_module(presentation_root, [1, 2, 3])
    deck = Presentation()
    first = deck.slides.add_slide(deck.slide_layouts[6])
    shape = first.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(shape, "{{fig:line}}")
    second = deck.slides.add_slide(deck.slide_layouts[6])
    second.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(2), Inches(0.5)).text = "untouched"
    deck.save(presentation_root / "deck.pptx")
    before = _package_payloads(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    update_figures(presentation, figure_id="line")

    changed = _changed_package_entries(before, _package_payloads(presentation_root / "deck.pptx"))
    assert changed <= {
        "[Content_Types].xml",
        "ppt/slides/slide1.xml",
        "ppt/slides/_rels/slide1.xml.rels",
        "ppt/media/image1.png",
    }
    assert "ppt/slides/slide2.xml" not in changed


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


def test_update_figures_deletes_unreferenced_replaced_media(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    _write_single_line_figure_module(presentation_root, [1, 2, 3])
    _write_deck_with_figure_anchor(presentation_root, "{{fig:line}}")
    presentation = load_presentation_definition(tmp_path, "demo")
    update_figures(presentation, figure_id="line")
    first_media = _media_payloads(presentation_root / "deck.pptx")
    before_second_update = _package_payloads(presentation_root / "deck.pptx")

    _write_single_line_figure_module(presentation_root, [3, 1, 2])
    presentation = load_presentation_definition(tmp_path, "demo")
    update_figures(presentation, figure_id="line")

    changed = _changed_package_entries(before_second_update, _package_payloads(presentation_root / "deck.pptx"))
    assert changed == {"ppt/media/image1.png"}
    second_media = _media_payloads(presentation_root / "deck.pptx")
    assert len(second_media) == 1
    assert second_media != first_media
    assert _unused_image_relationship_ids(presentation_root / "deck.pptx", slide_number=1) == []
    assert _duplicate_c_nv_pr_ids(presentation_root / "deck.pptx", slide_number=1) == []


def test_update_figures_detaches_copied_managed_picture_anchors(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    _write_two_line_figure_module(presentation_root)
    seed = presentation_root / "seed.png"
    Image.new("RGB", (20, 20), "blue").save(seed)
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    first = slide.shapes.add_picture(str(seed), Inches(0.5), Inches(0.5), width=Inches(2))
    second = slide.shapes.add_picture(str(seed), Inches(3), Inches(0.5), width=Inches(2))
    _set_shape_alt_text(first, "{{fig:first}}")
    _set_shape_alt_text(second, "{{fig:second}}")
    deck.save(presentation_root / "deck.pptx")
    assert _picture_blip_ids(presentation_root / "deck.pptx", slide_number=1) == ["rId2", "rId2"]
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation)

    assert [path.name for path in outputs] == ["first.png", "second.png"]
    blip_ids = _picture_blip_ids(presentation_root / "deck.pptx", slide_number=1)
    assert len(set(blip_ids)) == 2
    targets = _image_relationship_targets(presentation_root / "deck.pptx", slide_number=1)
    assert len({targets[r_id] for r_id in blip_ids}) == 2
    media_payloads = _package_payloads(presentation_root / "deck.pptx")
    assert media_payloads[f"ppt/media/{targets[blip_ids[0]].split('/')[-1]}"] == outputs[0].read_bytes()
    assert media_payloads[f"ppt/media/{targets[blip_ids[1]].split('/')[-1]}"] == outputs[1].read_bytes()


def test_update_figures_removes_stale_image_relationships(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    _write_single_line_figure_module(presentation_root, [1, 2, 3])
    _write_deck_with_figure_anchor(presentation_root, "{{fig:line}}")
    presentation = load_presentation_definition(tmp_path, "demo")
    update_figures(presentation, figure_id="line")

    _write_single_line_figure_module(presentation_root, [3, 1, 2])
    presentation = load_presentation_definition(tmp_path, "demo")
    update_figures(presentation, figure_id="line")

    assert _unused_image_relationship_ids(presentation_root / "deck.pptx", slide_number=1) == []


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


def test_update_figures_saves_reused_axes_panel_with_tight_bbox(tmp_path: Path) -> None:
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
                "def plot_reused_axes(ctx):",
                "    fig = plt.figure(figsize=(5, 5))",
                "    ax = fig.add_axes([0.1275, 0.11, 0.77, 0.77])",
                "    ax.plot([0, 1, 2], [0, 1, 0])",
                "    ax.set_xlabel('x label that must remain visible')",
                "    ax.set_ylabel('y label that must remain visible')",
                "    return FigureResult(ax)",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(shape, "{{fig:reused_axes}}")
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    outputs = update_figures(presentation, figure_id="reused_axes")

    assert [path.name for path in outputs] == ["reused_axes.png"]
    with Image.open(outputs[0]) as image:
        assert image.size != (600, 400)


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


def _write_deck_with_figure_anchor(presentation_root: Path, token: str) -> None:
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
    _set_shape_alt_text(shape, token)
    deck.save(presentation_root / "deck.pptx")


def _write_single_line_figure_module(presentation_root: Path, y_values: list[int]) -> None:
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "import matplotlib.pyplot as plt",
                "from pubify_data import figure",
                "@figure",
                "def plot_line(ctx):",
                "    fig, ax = plt.subplots()",
                f"    ax.plot([1, 2, 3], {y_values!r})",
                "    return fig",
            ]
        )
        + "\n",
        encoding="utf-8",
    )


def _write_two_line_figure_module(presentation_root: Path) -> None:
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "import matplotlib.pyplot as plt",
                "from pubify_data import figure",
                "@figure",
                "def plot_first(ctx):",
                "    fig, ax = plt.subplots()",
                "    ax.plot([1, 2, 3], [1, 2, 3])",
                "    return fig",
                "@figure",
                "def plot_second(ctx):",
                "    fig, ax = plt.subplots()",
                "    ax.plot([1, 2, 3], [3, 1, 2])",
                "    return fig",
            ]
        )
        + "\n",
        encoding="utf-8",
    )


def _unused_image_relationship_ids(deck_path: Path, *, slide_number: int) -> list[str]:
    relationship_namespace = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    relationship_attribute_namespace = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    image_relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    with ZipFile(deck_path) as package:
        slide_xml = package.read(f"ppt/slides/slide{slide_number}.xml")
        rels_xml = package.read(f"ppt/slides/_rels/slide{slide_number}.xml.rels")
    slide = ElementTree.fromstring(slide_xml)
    rels = ElementTree.fromstring(rels_xml)
    used_ids = {
        value
        for element in slide.iter()
        if (value := element.attrib.get(f"{relationship_attribute_namespace}embed")) is not None
    }
    image_ids = {
        relationship.attrib["Id"]
        for relationship in rels.findall(f"{relationship_namespace}Relationship")
        if relationship.attrib.get("Type") == image_relationship_type
    }
    return sorted(image_ids - used_ids)


def _picture_blip_ids(deck_path: Path, *, slide_number: int) -> list[str]:
    relationship_attribute_namespace = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    with ZipFile(deck_path) as package:
        slide_xml = package.read(f"ppt/slides/slide{slide_number}.xml")
    slide = ElementTree.fromstring(slide_xml)
    return [
        value
        for element in slide.iter()
        if element.tag.endswith("}blip")
        if (value := element.attrib.get(f"{relationship_attribute_namespace}embed")) is not None
    ]


def _image_relationship_targets(deck_path: Path, *, slide_number: int) -> dict[str, str]:
    relationship_namespace = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    image_relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    with ZipFile(deck_path) as package:
        rels_xml = package.read(f"ppt/slides/_rels/slide{slide_number}.xml.rels")
    rels = ElementTree.fromstring(rels_xml)
    return {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in rels.findall(f"{relationship_namespace}Relationship")
        if relationship.attrib.get("Type") == image_relationship_type
    }


def _package_payloads(deck_path: Path) -> dict[str, bytes]:
    with ZipFile(deck_path) as package:
        return {name: package.read(name) for name in package.namelist()}


def _changed_package_entries(before: dict[str, bytes], after: dict[str, bytes]) -> set[str]:
    names = set(before) | set(after)
    return {name for name in names if before.get(name) != after.get(name)}


def _media_payloads(deck_path: Path) -> dict[str, bytes]:
    with ZipFile(deck_path) as package:
        return {
            name: package.read(name)
            for name in package.namelist()
            if name.startswith("ppt/media/")
        }


def _duplicate_c_nv_pr_ids(deck_path: Path, *, slide_number: int) -> list[str]:
    presentation_namespace = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    with ZipFile(deck_path) as package:
        slide_xml = package.read(f"ppt/slides/slide{slide_number}.xml")
    slide = ElementTree.fromstring(slide_xml)
    ids = [
        element.attrib["id"]
        for element in slide.iter(f"{presentation_namespace}cNvPr")
        if "id" in element.attrib
    ]
    return sorted({value for value in ids if ids.count(value) > 1})
