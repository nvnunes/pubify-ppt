from __future__ import annotations

from pathlib import Path

import pytest

from pubify_ppt.config import (
    DEFAULT_PRESENTATIONS_ROOT,
    load_workspace_config,
    render_default_workspace_config,
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
