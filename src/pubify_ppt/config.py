from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pubify_data
from pubify_data.config import load_pubify_config, parse_simple_yaml, resolve_workspace_relative_path


WORKSPACE_CONFIG_FILENAME = "pubify.yaml"
WORKSPACE_CONFIG_SECTION = "pubify-ppt"
DEFAULT_PRESENTATIONS_ROOT = "slides"
PRESENTATION_CONFIG_FILENAME = "ppt.yaml"
DEFAULT_DECK_FILENAME = "deck.pptx"
DEFAULT_BACKUP_RETENTION = 5
DEFAULT_IMAGE_FORMAT = "png"
DEFAULT_DPI = 200
DEFAULT_FIT = "contain"
PRESENTATION_CONFIG_KEYS = frozenset({"deck", "backup_retention", "defaults", "external_data_roots", "sources"})
PRESENTATION_DEFAULTS_KEYS = frozenset({"image_format", "dpi", "fit"})


@dataclass(frozen=True)
class WorkspaceConfig:
    """Workspace-level presentation roots loaded from ``pubify.yaml``.

    Attributes:
        workspace_root: Resolved host workspace root containing ``pubify.yaml``.
        presentations_root: Resolved directory that contains presentation
            folders managed by ``pubify-ppt``.
    """

    workspace_root: Path
    presentations_root: Path


@dataclass(frozen=True)
class PresentationDefaults:
    """Presentation-local defaults loaded from ``ppt.yaml``.

    Attributes:
        image_format: Generated figure image format. V1 supports ``png``.
        dpi: DPI-equivalent figure render scale at the current anchor size.
        fit: Image placement mode inside the anchor box. V1 supports
            ``contain``.
    """

    image_format: str = DEFAULT_IMAGE_FORMAT
    dpi: int = DEFAULT_DPI
    fit: str = DEFAULT_FIT


@dataclass(frozen=True)
class PresentationConfig:
    """Presentation-local workflow settings loaded from ``ppt.yaml``.

    Attributes:
        deck: Presentation-local editable source deck path.
        backup_retention: Number of in-place backups to retain.
        defaults: Figure rendering defaults for deck updates.
        external_data_roots: Named roots used by ``@external_data(...)``.
        sources: Named source publication roots.
    """

    deck: str = DEFAULT_DECK_FILENAME
    backup_retention: int = DEFAULT_BACKUP_RETENTION
    defaults: PresentationDefaults = field(default_factory=PresentationDefaults)
    external_data_roots: dict[str, str] = field(default_factory=dict)
    sources: dict[str, str] = field(default_factory=dict)

    @property
    def deck_path(self) -> Path:
        return Path(self.deck)


def find_workspace_root(start: Path | None = None) -> Path:
    """Walk upward from ``start`` until a ``pubify.yaml`` workspace root is found."""

    return pubify_data.find_workspace_root(start)


def load_workspace_config(workspace_root: Path) -> WorkspaceConfig:
    """Load workspace-level presentation settings from ``pubify.yaml``."""

    root = workspace_root.resolve()
    config_path = root / WORKSPACE_CONFIG_FILENAME
    section = pubify_data.load_config_section(root, WORKSPACE_CONFIG_SECTION)
    if not section:
        raise ValueError(f"{config_path}: missing required {WORKSPACE_CONFIG_SECTION} section")
    presentations_root = resolve_workspace_relative_path(section, config_path, "presentations_root")
    if presentations_root is None:
        raise ValueError(f"{config_path}: presentations_root must be a non-empty string")
    return WorkspaceConfig(workspace_root=root, presentations_root=presentations_root)


def workspace_config_has_section(workspace_root: Path) -> bool:
    """Return whether ``pubify.yaml`` already contains a ``pubify-ppt`` section."""

    return WORKSPACE_CONFIG_SECTION in load_pubify_config(workspace_root).raw


def load_presentation_config(path: Path) -> PresentationConfig:
    """Load and validate one presentation-local ``ppt.yaml`` file."""

    raw = parse_simple_yaml(path.read_text(encoding="utf-8"))
    _reject_unknown_mapping_keys(raw, PRESENTATION_CONFIG_KEYS, path, "ppt.yaml")
    deck = raw.get("deck", DEFAULT_DECK_FILENAME)
    if not isinstance(deck, str) or not deck:
        raise ValueError(f"{path}: deck must be a non-empty string")
    deck_path = Path(deck)
    if deck_path.is_absolute() or ".." in deck_path.parts:
        raise ValueError(f"{path}: deck must be a presentation-local relative path")

    backup_retention = raw.get("backup_retention", DEFAULT_BACKUP_RETENTION)
    if isinstance(backup_retention, bool) or not isinstance(backup_retention, int) or backup_retention < 0:
        raise ValueError(f"{path}: backup_retention must be a non-negative integer")

    defaults_raw = raw.get("defaults", {})
    if not isinstance(defaults_raw, dict):
        raise ValueError(f"{path}: defaults must be a mapping when set")
    _reject_unknown_mapping_keys(defaults_raw, PRESENTATION_DEFAULTS_KEYS, path, "ppt.yaml defaults")
    image_format = defaults_raw.get("image_format", DEFAULT_IMAGE_FORMAT)
    if image_format != DEFAULT_IMAGE_FORMAT:
        raise ValueError(f"{path}: defaults.image_format must be png")
    dpi = defaults_raw.get("dpi", DEFAULT_DPI)
    if isinstance(dpi, bool) or not isinstance(dpi, int) or dpi <= 0:
        raise ValueError(f"{path}: defaults.dpi must be a positive integer")
    fit = defaults_raw.get("fit", DEFAULT_FIT)
    if fit != DEFAULT_FIT:
        raise ValueError(f"{path}: defaults.fit must be contain")

    external_data_roots = raw.get("external_data_roots", {})
    if not isinstance(external_data_roots, dict):
        raise ValueError(f"{path}: external_data_roots must be a mapping when set")
    normalized_external_roots: dict[str, str] = {}
    for root_name, root_path in external_data_roots.items():
        if not isinstance(root_name, str) or not root_name:
            raise ValueError(f"{path}: external_data_roots keys must be non-empty strings")
        if not isinstance(root_path, str) or not root_path:
            raise ValueError(f"{path}: external_data_roots.{root_name} must be a non-empty string")
        normalized_external_roots[root_name] = root_path

    sources = raw.get("sources", {})
    if not isinstance(sources, dict):
        raise ValueError(f"{path}: sources must be a mapping when set")
    normalized_sources: dict[str, str] = {}
    for source_name, source_path in sources.items():
        if not isinstance(source_name, str) or not source_name:
            raise ValueError(f"{path}: sources keys must be non-empty strings")
        if not isinstance(source_path, str) or not source_path:
            raise ValueError(f"{path}: sources.{source_name} must be a non-empty string")
        normalized_sources[source_name] = source_path

    return PresentationConfig(
        deck=deck,
        backup_retention=backup_retention,
        defaults=PresentationDefaults(image_format=image_format, dpi=dpi, fit=fit),
        external_data_roots=normalized_external_roots,
        sources=normalized_sources,
    )


def _reject_unknown_mapping_keys(
    raw: dict[str, object],
    allowed_keys: frozenset[str],
    path: Path,
    label: str,
) -> None:
    unknown = sorted(set(raw) - allowed_keys)
    if unknown:
        joined = ", ".join(unknown)
        raise ValueError(f"{path}: unknown {label} key(s): {joined}")


def render_default_workspace_config() -> str:
    """Render the default workspace config used by ``ppt init``."""

    return "\n".join(
        [
            f"{WORKSPACE_CONFIG_SECTION}:",
            f"  presentations_root: {DEFAULT_PRESENTATIONS_ROOT}",
            "",
        ]
    )


def write_default_workspace_config(path: Path) -> None:
    """Write the default ``pubify.yaml`` scaffold for ``ppt init``."""

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_default_workspace_config(), encoding="utf-8")


def render_default_presentation_config() -> str:
    """Render the default presentation-local ``ppt.yaml`` scaffold."""

    return "\n".join(
        [
            f"deck: {DEFAULT_DECK_FILENAME}",
            f"backup_retention: {DEFAULT_BACKUP_RETENTION}",
            "defaults:",
            f"  image_format: {DEFAULT_IMAGE_FORMAT}",
            f"  dpi: {DEFAULT_DPI}",
            f"  fit: {DEFAULT_FIT}",
            "external_data_roots:",
            "sources:",
            "",
        ]
    )


def write_default_presentation_config(path: Path) -> None:
    """Write the default ``ppt.yaml`` scaffold for ``ppt init <presentation-id>``."""

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_default_presentation_config(), encoding="utf-8")
