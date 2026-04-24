from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches
import pytest

from pubify_ppt.discovery import load_presentation_definition
from pubify_ppt.init import init_presentation_by_id, init_workspace
from pubify_ppt.stats import update_stats


def test_update_stats_replaces_starter_dictionary_token(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation = load_presentation_definition(tmp_path, "demo")

    replacements = update_stats(presentation)

    assert [(item.token, item.value) for item in replacements] == [("{{stat:example.count}}", "3")]
    assert "Example count: 3" in _deck_text(tmp_path / "slides" / "demo" / "deck.pptx")


def test_update_stats_replaces_repeated_scalar_tokens_and_preserves_run_format(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "from pubify_data import stat",
                "@stat",
                "def compute_total(ctx):",
                "    return 42",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    text_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(5.0), Inches(0.5))
    run = text_box.text_frame.paragraphs[0].add_run()
    run.text = "{{stat:total}} and {{stat:total}}"
    run.font.bold = True
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    replacements = update_stats(presentation)

    assert [item.value for item in replacements] == ["42", "42"]
    deck = Presentation(presentation_root / "deck.pptx")
    run = deck.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    assert run.text == "42 and 42"
    assert run.font.bold is True


def test_update_stats_targeted_update_leaves_other_tokens_unchanged(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    (presentation_root / "figures.py").write_text(
        "\n".join(
            [
                "from pubify_data import stat",
                "@stat",
                "def compute_first(ctx):",
                "    return 'one'",
                "@stat",
                "def compute_second(ctx):",
                "    return 'two'",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(6.0), Inches(0.5)).text = (
        "{{stat:first}} {{stat:second}}"
    )
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    update_stats(presentation, stat_id="first")

    assert "one {{stat:second}}" in _deck_text(presentation_root / "deck.pptx")


def test_update_stats_reports_missing_dictionary_key(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")
    presentation_root = tmp_path / "slides" / "demo"
    deck = Presentation(presentation_root / "deck.pptx")
    slide = deck.slides[0]
    slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(4.0), Inches(0.5)).text = "{{stat:example.missing}}"
    deck.save(presentation_root / "deck.pptx")
    presentation = load_presentation_definition(tmp_path, "demo")

    with pytest.raises(ValueError, match="Missing key 'missing' for stat 'example'"):
        update_stats(presentation)


def _deck_text(deck_path: Path) -> str:
    deck = Presentation(deck_path)
    return "\n".join(shape.text for slide in deck.slides for shape in slide.shapes if hasattr(shape, "text"))
