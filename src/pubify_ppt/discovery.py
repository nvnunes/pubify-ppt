from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from types import ModuleType

import pubify_data

from pubify_ppt.config import (
    PRESENTATION_CONFIG_FILENAME,
    PresentationConfig,
    load_presentation_config,
    load_workspace_config,
)
from pubify_ppt.init import (
    BACKUPS_ARTIFACT_DIR,
    FIGURES_ARTIFACT_DIR,
    PRESENTATION_ENTRYPOINT,
    PPT_ARTIFACTS_NAMESPACE,
)


LoaderSpec = pubify_data.LoaderSpec
FigureSpec = pubify_data.FigureSpec
StatSpec = pubify_data.StatSpec
TableSpec = pubify_data.TableSpec


@dataclass(frozen=True)
class PresentationPaths:
    """Resolved workspace and presentation paths used by the runtime.

    Attributes:
        workspace_root: Host workspace root containing ``pubify.yaml``.
        presentation_root: Directory for one presentation id.
        data_root: Presentation-local canonical data root.
        ppt_artifacts_root: Generated PowerPoint artifact namespace.
        figures_root: Generated figure PNG artifact directory.
        backups_root: Generated deck backup artifact directory.
        entrypoint: Presentation-local ``figures.py`` entrypoint.
        config_path: Presentation-local ``ppt.yaml`` config.
        deck_path: Editable source deck path resolved from ``ppt.yaml``.
    """

    workspace_root: Path
    presentation_root: Path
    data_root: Path
    ppt_artifacts_root: Path
    figures_root: Path
    backups_root: Path
    entrypoint: Path
    config_path: Path
    deck_path: Path


@dataclass(frozen=True)
class PresentationDefinition:
    """Loaded presentation module plus resolved config, paths, and decorators."""

    presentation_id: str
    paths: PresentationPaths
    config: PresentationConfig
    upstream: pubify_data.PublicationDefinition

    @property
    def module(self) -> ModuleType:
        return self.upstream.module

    @property
    def loaders(self) -> dict[str, LoaderSpec]:
        return self.upstream.loaders

    @property
    def figures(self) -> dict[str, FigureSpec]:
        return self.upstream.figures

    @property
    def stats(self) -> dict[str, StatSpec]:
        return self.upstream.stats

    @property
    def tables(self) -> dict[str, TableSpec]:
        return self.upstream.tables


def list_presentation_ids(workspace_root: Path) -> list[str]:
    """List presentation ids under the configured workspace presentations root."""

    presentations_root = load_workspace_config(workspace_root).presentations_root
    if not presentations_root.exists():
        return []
    return sorted(path.name for path in presentations_root.iterdir() if path.is_dir())


def build_presentation_paths(
    workspace_root: Path,
    presentation_id: str,
    config: PresentationConfig | None = None,
) -> PresentationPaths:
    """Resolve framework-owned paths for one presentation under a workspace."""

    workspace = load_workspace_config(workspace_root)
    presentation_root = workspace.presentations_root / presentation_id
    config_path = presentation_root / PRESENTATION_CONFIG_FILENAME
    if config is None and config_path.exists():
        config = load_presentation_config(config_path)
    deck_relative = config.deck_path if config is not None else Path("deck.pptx")
    data_root = presentation_root / "data"
    ppt_artifacts_root = pubify_data.artifact_namespace_root(data_root, PPT_ARTIFACTS_NAMESPACE)
    return PresentationPaths(
        workspace_root=workspace_root.resolve(),
        presentation_root=presentation_root,
        data_root=data_root,
        ppt_artifacts_root=ppt_artifacts_root,
        figures_root=ppt_artifacts_root / FIGURES_ARTIFACT_DIR,
        backups_root=ppt_artifacts_root / BACKUPS_ARTIFACT_DIR,
        entrypoint=presentation_root / PRESENTATION_ENTRYPOINT,
        config_path=config_path,
        deck_path=presentation_root / deck_relative,
    )


def load_presentation_definition(workspace_root: Path, presentation_id: str) -> PresentationDefinition:
    """Load one presentation's config, entrypoint module, loaders, figures, and stats."""

    initial_paths = build_presentation_paths(workspace_root, presentation_id)
    if not initial_paths.presentation_root.exists():
        raise FileNotFoundError(f"Unknown presentation '{presentation_id}'")
    if not initial_paths.config_path.exists():
        raise FileNotFoundError(f"Missing presentation config: {initial_paths.config_path}")
    config = load_presentation_config(initial_paths.config_path)
    paths = build_presentation_paths(workspace_root, presentation_id, config=config)
    if not paths.entrypoint.exists():
        raise FileNotFoundError(f"Missing figures entrypoint: {paths.entrypoint}")

    adapter = pubify_data.PublicationAdapter(
        publication_id=presentation_id,
        publication_root=paths.presentation_root,
        entrypoint=paths.entrypoint,
        data_root=paths.data_root,
        external_data_roots=_resolve_external_data_roots(paths.workspace_root, config.external_data_roots),
        source_roots=_resolve_source_roots(paths.workspace_root, config.sources),
        workspace=pubify_data.WorkspaceAdapter(paths.workspace_root),
    )
    upstream = pubify_data.load_publication_from_entrypoint(presentation_id, adapter=adapter)
    return PresentationDefinition(
        presentation_id=presentation_id,
        paths=paths,
        config=config,
        upstream=upstream,
    )


def _resolve_external_data_roots(workspace_root: Path, roots: dict[str, str]) -> dict[str, Path]:
    resolved: dict[str, Path] = {}
    for name, root in roots.items():
        root_path = Path(root).expanduser()
        if not root_path.is_absolute():
            root_path = (workspace_root / root_path).resolve()
        resolved[name] = root_path
    return resolved


def _resolve_source_roots(workspace_root: Path, roots: dict[str, str]) -> dict[str, Path]:
    resolved: dict[str, Path] = {}
    for name, root in roots.items():
        root_path = Path(root).expanduser()
        if not root_path.is_absolute():
            root_path = (workspace_root / root_path).resolve()
        resolved[name] = root_path
    return resolved
