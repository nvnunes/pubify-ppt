from __future__ import annotations

from collections.abc import Sequence

import pubify_data


class FigureResult(pubify_data.BaseFigureResult):
    """PowerPoint figure result with slide-facing metadata reserved for pubify-ppt."""

    def __init__(
        self,
        panels_or_panel: object | Sequence[object],
        *,
        layout: str | None = None,
        metadata: dict[str, object] | None = None,
    ) -> None:
        super().__init__(panels_or_panel, layout=layout, metadata=metadata)


class StatResult(pubify_data.BaseStatResult):
    """PowerPoint stat result with slide-facing metadata reserved for pubify-ppt."""


class TableResult(pubify_data.BaseTableResult):
    """PowerPoint table result with slide-facing metadata reserved for pubify-ppt."""
