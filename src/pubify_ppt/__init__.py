"""PowerPoint presentation workflow package built on ``pubify-data``."""

from .config import (
    PresentationConfig,
    PresentationDefaults,
    WorkspaceConfig,
    find_workspace_root,
    load_presentation_config,
    load_workspace_config,
)
from .discovery import PresentationDefinition, PresentationPaths, load_presentation_definition
from .figures import update_figures
from .runtime import check_presentation
from .stats import update_stats

__all__ = [
    "PresentationConfig",
    "PresentationDefinition",
    "PresentationDefaults",
    "PresentationPaths",
    "WorkspaceConfig",
    "check_presentation",
    "find_workspace_root",
    "load_presentation_config",
    "load_presentation_definition",
    "load_workspace_config",
    "update_figures",
    "update_stats",
]
