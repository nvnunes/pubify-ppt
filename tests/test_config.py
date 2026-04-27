from __future__ import annotations

from pathlib import Path

import pytest

from pubify_ppt.config import (
    DEFAULT_PRESENTATIONS_ROOT,
    load_presentation_config,
    load_workspace_config,
    render_default_presentation_config,
    render_default_workspace_config,
    write_default_presentation_config,
    write_default_workspace_config,
)


def test_render_default_workspace_config_uses_slides_root() -> None:
    assert render_default_workspace_config() == (
        "pubify-ppt:\n"
        f"  presentations_root: {DEFAULT_PRESENTATIONS_ROOT}\n"
    )


def test_write_default_workspace_config(tmp_path: Path) -> None:
    config_path = tmp_path / "pubify.yaml"

    write_default_workspace_config(config_path)

    assert config_path.read_text(encoding="utf-8") == render_default_workspace_config()


def test_render_default_presentation_config() -> None:
    assert render_default_presentation_config() == (
        "deck: deck.pptx\n"
        "backup_retention: 5\n"
        "defaults:\n"
        "  dpi: 200\n"
        "external_data_roots:\n"
        "sources:\n"
    )


def test_load_presentation_config_parses_defaults(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    write_default_presentation_config(config_path)

    config = load_presentation_config(config_path)

    assert config.deck == "deck.pptx"
    assert config.backup_retention == 5
    assert config.defaults.dpi == 200
    assert config.defaults.figure_font_family is None
    assert config.defaults.figure_style == {}
    assert config.external_data_roots == {}
    assert config.sources == {}


def test_load_presentation_config_accepts_legacy_image_format_and_fit_defaults(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "defaults:\n"
        "  image_format: png\n"
        "  fit: contain\n",
        encoding="utf-8",
    )

    config = load_presentation_config(config_path)

    assert config.defaults.dpi == 200


def test_load_presentation_config_parses_figure_font_family_default(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "defaults:\n"
        "  dpi: 200\n"
        "  figure_font_family: Aptos\n",
        encoding="utf-8",
    )

    config = load_presentation_config(config_path)

    assert config.defaults.figure_font_family == "Aptos"


def test_load_presentation_config_parses_figure_font_size_defaults(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "defaults:\n"
        "  figure_base_fontsize_pt: 11\n"
        "  figure_axes_labelsize_pt: 11\n"
        "  figure_tick_labelsize_pt: 10\n"
        "  figure_legend_fontsize_pt: 10\n"
        "  figure_title_fontsize_pt: 12\n",
        encoding="utf-8",
    )

    config = load_presentation_config(config_path)

    assert config.defaults.figure_style == {
        "base_fontsize_pt": 11.0,
        "axes_labelsize_pt": 11.0,
        "tick_labelsize_pt": 10.0,
        "legend_fontsize_pt": 10.0,
        "title_fontsize_pt": 12.0,
    }


@pytest.mark.parametrize("value", ["0", "-1", "true", "large"])
def test_load_presentation_config_rejects_invalid_figure_font_size_defaults(tmp_path: Path, value: str) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "defaults:\n"
        f"  figure_base_fontsize_pt: {value}\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="defaults.figure_base_fontsize_pt must be a positive number"):
        load_presentation_config(config_path)


@pytest.mark.parametrize("value", ["", "  "])
def test_load_presentation_config_rejects_empty_figure_font_family(tmp_path: Path, value: str) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "defaults:\n"
        f"  figure_font_family: {value}\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="defaults.figure_font_family must be a non-empty string"):
        load_presentation_config(config_path)


def test_load_presentation_config_rejects_unknown_top_level_keys(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "backup_retensions: 5\n"
        "defaults:\n"
        "  dpi: 200\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match=r"unknown ppt\.yaml key\(s\): backup_retensions"):
        load_presentation_config(config_path)


def test_load_presentation_config_rejects_unknown_defaults_keys(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "defaults:\n"
        "  dp: 200\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match=r"unknown ppt\.yaml defaults key\(s\): dp"):
        load_presentation_config(config_path)


def test_load_presentation_config_rejects_bool_numeric_fields(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    config_path.write_text(
        "deck: deck.pptx\n"
        "backup_retention: true\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="backup_retention must be a non-negative integer"):
        load_presentation_config(config_path)


def test_load_workspace_config_resolves_presentations_root(tmp_path: Path) -> None:
    (tmp_path / "pubify.yaml").write_text(
        "pubify-ppt:\n  presentations_root: slides\n",
        encoding="utf-8",
    )

    config = load_workspace_config(tmp_path)

    assert config.workspace_root == tmp_path.resolve()
    assert config.presentations_root == tmp_path.resolve() / "slides"


def test_load_workspace_config_requires_pubify_ppt_section(tmp_path: Path) -> None:
    (tmp_path / "pubify.yaml").write_text(
        "pubify-pubs:\n  publications_root: papers\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="missing required pubify-ppt section"):
        load_workspace_config(tmp_path)
