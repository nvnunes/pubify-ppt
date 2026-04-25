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
        "  image_format: png\n"
        "  dpi: 200\n"
        "  fit: contain\n"
        "external_data_roots:\n"
        "sources:\n"
    )


def test_load_presentation_config_parses_defaults(tmp_path: Path) -> None:
    config_path = tmp_path / "ppt.yaml"
    write_default_presentation_config(config_path)

    config = load_presentation_config(config_path)

    assert config.deck == "deck.pptx"
    assert config.backup_retention == 5
    assert config.defaults.image_format == "png"
    assert config.defaults.dpi == 200
    assert config.defaults.fit == "contain"
    assert config.external_data_roots == {}
    assert config.sources == {}


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
