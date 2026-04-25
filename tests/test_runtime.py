from __future__ import annotations

import pubify_ppt


def test_package_exports_workspace_config_helpers() -> None:
    assert "WorkspaceConfig" in pubify_ppt.__all__
    assert "load_workspace_config" in pubify_ppt.__all__
    assert "update_tables" in pubify_ppt.__all__
