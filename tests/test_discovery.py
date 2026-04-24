from __future__ import annotations

from pathlib import Path

from pubify_ppt.discovery import (
    build_presentation_paths,
    list_presentation_ids,
    load_presentation_definition,
)
from pubify_ppt.init import init_presentation_by_id, init_workspace


def test_list_presentation_ids_uses_configured_presentations_root(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "b-talk")
    init_presentation_by_id(tmp_path, "a-talk")

    assert list_presentation_ids(tmp_path) == ["a-talk", "b-talk"]


def test_load_presentation_definition_uses_pubify_data_adapter(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")

    presentation = load_presentation_definition(tmp_path, "demo")

    assert presentation.presentation_id == "demo"
    assert sorted(presentation.loaders) == ["example"]
    assert sorted(presentation.figures) == ["example"]
    assert sorted(presentation.stats) == ["example"]
    assert presentation.upstream.adapter.data_root == tmp_path / "slides" / "demo" / "data"


def test_build_presentation_paths_resolves_artifact_namespace(tmp_path: Path) -> None:
    init_workspace(tmp_path)
    init_presentation_by_id(tmp_path, "demo")

    paths = build_presentation_paths(tmp_path, "demo")

    assert paths.ppt_artifacts_root == tmp_path / "slides" / "demo" / "data" / "ppt-artifacts"
    assert paths.figures_root == paths.ppt_artifacts_root / "figures"
    assert paths.backups_root == paths.ppt_artifacts_root / "backups"
