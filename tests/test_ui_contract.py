from __future__ import annotations

import ast
from pathlib import Path

UI_SOURCE = Path(__file__).parents[1] / "src" / "excel_git_viewer" / "ui.py"


def load_main_window_method(name: str) -> ast.FunctionDef:
    module = ast.parse(UI_SOURCE.read_text(encoding="utf-8"))
    main_window = next(
        node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "MainWindow"
    )
    return next(
        node for node in main_window.body if isinstance(node, ast.FunctionDef) and node.name == name
    )


def self_calls(method: ast.FunctionDef, name: str) -> list[ast.Call]:
    return [
        node
        for node in ast.walk(method)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and isinstance(node.func.value, ast.Name)
        and node.func.value.id == "self"
        and node.func.attr == name
    ]


def connected_handler(method: ast.FunctionDef, owner: str, signal: str) -> str | None:
    for node in ast.walk(method):
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
            continue
        signal_node = node.func.value
        if node.func.attr != "connect" or not isinstance(signal_node, ast.Attribute):
            continue
        owner_node = signal_node.value
        if (
            signal_node.attr == signal
            and isinstance(owner_node, ast.Attribute)
            and isinstance(owner_node.value, ast.Name)
            and owner_node.value.id == "self"
            and owner_node.attr == owner
            and len(node.args) == 1
            and isinstance(node.args[0], ast.Attribute)
        ):
            return node.args[0].attr
    return None


def test_history_ui_restores_cache_without_an_excel_commit_filter() -> None:
    init_method = load_main_window_method("__init__")
    open_method = load_main_window_method("_open_repository")
    module = ast.parse(UI_SOURCE.read_text(encoding="utf-8"))

    assert len(self_calls(init_method, "_restore_last_repository")) == 1
    [reload_call] = self_calls(open_method, "_reload_commits")
    assert reload_call.keywords == []
    assert not any(
        isinstance(node, ast.Attribute) and node.attr == "only_excel_checkbox"
        for node in ast.walk(module)
    )


def test_refresh_button_forces_git_history_reload() -> None:
    build_method = load_main_window_method("_build_ui")
    force_method = load_main_window_method("_force_reload_commits")

    assert connected_handler(build_method, "refresh_button", "clicked") == "_force_reload_commits"
    [reload_call] = self_calls(force_method, "_reload_commits")
    assert len(reload_call.keywords) == 1
    keyword = reload_call.keywords[0]
    assert keyword.arg == "force_refresh"
    assert isinstance(keyword.value, ast.Constant) and keyword.value.value is True


def test_background_workers_are_retained_until_their_ui_callback() -> None:
    init_method = load_main_window_method("__init__")
    start_method = load_main_window_method("_start_task")
    terminal_methods = [
        load_main_window_method("_task_completed"),
        load_main_window_method("_task_failed"),
        load_main_window_method("_task_cancelled"),
    ]

    assert any(
        isinstance(node, ast.Attribute) and node.attr == "_workers"
        for node in ast.walk(init_method)
    )
    assert len(self_calls(start_method, "_retain_worker")) == 1
    assert all(len(self_calls(method, "_release_worker")) == 1 for method in terminal_methods)


def test_workspace_uses_tables_and_nested_resizable_splitters() -> None:
    build_method = load_main_window_method("_build_ui")
    attributes = {
        node.attr
        for node in ast.walk(build_method)
        if isinstance(node, ast.Attribute)
        and isinstance(node.value, ast.Name)
        and node.value.id == "self"
    }

    assert {"commit_table", "file_table"} <= attributes
    assert {
        "main_splitter",
        "navigation_splitter",
        "details_splitter",
        "context_splitter",
    } <= attributes
    assert "commit_list" not in attributes
    assert "file_list" not in attributes
