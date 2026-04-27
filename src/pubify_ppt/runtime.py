from __future__ import annotations

from pathlib import Path

import pubify_data

from pubify_ppt.anchors import validate_deck_anchors
from pubify_ppt.discovery import PresentationDefinition


def ensure_generated_artifact_paths(presentation: PresentationDefinition) -> None:
    """Create canonical generated PowerPoint artifact directories."""

    presentation.paths.data_root.mkdir(parents=True, exist_ok=True)
    presentation.paths.ppt_artifacts_root.mkdir(parents=True, exist_ok=True)
    presentation.paths.figures_root.mkdir(parents=True, exist_ok=True)
    presentation.paths.backups_root.mkdir(parents=True, exist_ok=True)


def check_presentation(
    presentation: PresentationDefinition,
    *,
    allow_shared_figure_relationships: bool = False,
) -> None:
    """Raise ``ValueError`` if the presentation fails static validation."""

    errors = validate_presentation_definition(
        presentation,
        allow_shared_figure_relationships=allow_shared_figure_relationships,
    )
    if errors:
        joined = "\n".join(f"- {message}" for message in errors)
        raise ValueError(f"Presentation '{presentation.presentation_id}' failed validation:\n{joined}")


def validate_presentation_definition(
    presentation: PresentationDefinition,
    *,
    allow_shared_figure_relationships: bool = False,
) -> list[str]:
    """Return static validation errors without running loaders, figures, or stats."""

    errors: list[str] = []
    paths = presentation.paths

    if not paths.data_root.exists():
        errors.append(f"Missing data directory: {paths.data_root}")
    elif not paths.data_root.is_dir():
        errors.append(f"Data path must be a directory or symlink to a directory: {paths.data_root}")

    if not paths.ppt_artifacts_root.exists():
        errors.append(f"Missing PowerPoint artifacts directory: {paths.ppt_artifacts_root}")
    if not paths.figures_root.exists():
        errors.append(f"Missing generated figures directory: {paths.figures_root}")
    if not paths.backups_root.exists():
        errors.append(f"Missing backup directory: {paths.backups_root}")

    if not paths.deck_path.exists():
        errors.append(f"Missing deck: {paths.deck_path}")
    elif not paths.deck_path.is_file():
        errors.append(f"Deck path must be a file: {paths.deck_path}")

    errors.extend(_validate_loader_paths(presentation))
    errors.extend(pubify_data.validate_dependencies(presentation.upstream))

    if paths.deck_path.exists() and paths.deck_path.is_file():
        errors.extend(
            validate_deck_anchors(
                paths.deck_path,
                figure_ids=set(presentation.figures),
                stat_ids=set(presentation.stats),
                table_ids=set(presentation.tables),
                report_shared_figure_relationships=not allow_shared_figure_relationships,
            )
        )

    return errors


def _validate_loader_paths(presentation: PresentationDefinition) -> list[str]:
    errors: list[str] = []
    for loader in presentation.loaders.values():
        root = _loader_root_for_validation(presentation, loader, errors)
        if root is None:
            continue
        for relative_path in loader.relative_paths.values():
            data_path = root / relative_path
            if not data_path.exists():
                errors.append(f"Missing {loader.kind} path for loader '{loader.loader_id}': {data_path}")
    return errors


def _loader_root_for_validation(
    presentation: PresentationDefinition,
    loader: object,
    errors: list[str],
) -> Path | None:
    if loader.kind == "data":
        return presentation.paths.data_root
    if loader.kind == "external_data":
        if loader.root_name is None or loader.root_name not in presentation.config.external_data_roots:
            errors.append(f"Missing external data root config for loader '{loader.loader_id}': {loader.root_name}")
            return None
        root = Path(presentation.upstream.external_data_roots[loader.root_name]).expanduser()
        if not root.exists():
            errors.append(f"Missing external data root path for loader '{loader.loader_id}': {root}")
            return None
        if not root.is_dir():
            errors.append(f"External data root must be a directory for loader '{loader.loader_id}': {root}")
            return None
        return root
    errors.append(f"Unsupported loader kind for loader '{loader.loader_id}': {loader.kind}")
    return None
