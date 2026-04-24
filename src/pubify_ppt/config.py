from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pubify_data
from pubify_data.config import resolve_workspace_relative_path


WORKSPACE_CONFIG_FILENAME = "pubify.yaml"
WORKSPACE_CONFIG_SECTION = "pubify-ppt"
DEFAULT_PRESENTATIONS_ROOT = "slides"


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
