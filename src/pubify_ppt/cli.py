from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from pathlib import Path
import sys

from pubify_ppt.config import WORKSPACE_CONFIG_SECTION
from pubify_ppt.config import find_workspace_root
from pubify_ppt.backups import write_deck
from pubify_ppt.discovery import list_presentation_ids, load_presentation_definition
from pubify_ppt.figures import FigureOutput
from pubify_ppt.figures import update_figures_in_deck
from pubify_ppt.init import init_presentation_by_id, init_workspace
from pubify_ppt.runtime import check_presentation
from pubify_ppt.stats import StatReplacement, update_stats_to_output
from pubify_ppt.tables import TableReplacement, update_tables_to_output
from pubify_ppt.update import update_presentation


def build_parser() -> argparse.ArgumentParser:
    """Build the ``ppt`` CLI parser."""

    parser = argparse.ArgumentParser(
        prog="ppt",
        usage="ppt <command>",
        description="\n".join(
            [
                "Commands:",
                "  ppt list",
                "  ppt init",
                "  ppt init <presentation-id>",
                "  ppt <presentation-id> check",
                "  ppt <presentation-id> data list",
                "  ppt <presentation-id> figure list",
                "  ppt <presentation-id> figure update [--output <path>]",
                "  ppt <presentation-id> figure <figure-id> update [--output <path>]",
                "  ppt <presentation-id> stat list",
                "  ppt <presentation-id> stat update [--output <path>]",
                "  ppt <presentation-id> stat <stat-id> update [--output <path>]",
                "  ppt <presentation-id> table list",
                "  ppt <presentation-id> table update [--output <path>]",
                "  ppt <presentation-id> table <table-id> update [--output <path>]",
                "  ppt <presentation-id> update [--output <path>]",
                "",
                f"Workspace config section: {WORKSPACE_CONFIG_SECTION}",
            ]
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("subject", nargs="?", help=argparse.SUPPRESS)
    parser.add_argument("arg2", nargs="?", help=argparse.SUPPRESS)
    parser.add_argument("arg3", nargs="?", help=argparse.SUPPRESS)
    parser.add_argument("arg4", nargs="?", help=argparse.SUPPRESS)
    parser.add_argument("arg5", nargs="?", help=argparse.SUPPRESS)
    parser.add_argument("--output", help=argparse.SUPPRESS)
    parser.add_argument("--force", action="store_true", help=argparse.SUPPRESS)
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the ``ppt`` CLI and return its process exit code."""

    parser = build_parser()
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    try:
        return _run_main(parser, args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


def _run_main(parser: argparse.ArgumentParser, args: argparse.Namespace) -> int:
    if args.subject == "list":
        _reject_output(parser, "list", args.output)
        _reject_force(parser, "list", args.force)
        if any(value is not None for value in (args.arg2, args.arg3, args.arg4, args.arg5)):
            parser.error("list does not accept additional arguments")
        workspace_root = find_workspace_root()
        for presentation_id in list_presentation_ids(workspace_root):
            print(presentation_id)
        return 0

    if args.subject == "init":
        _reject_output(parser, "init", args.output)
        _reject_force(parser, "init", args.force)
        if args.arg3 is not None or args.arg4 is not None or args.arg5 is not None:
            parser.error("init accepts at most optional <presentation-id>")
        if args.arg2 is None:
            workspace_root = init_workspace(Path.cwd())
            print(workspace_root)
            return 0
        workspace_root = find_workspace_root()
        presentation_root = init_presentation_by_id(workspace_root, args.arg2)
        print(presentation_root)
        return 0

    if args.subject is not None and args.arg2 is not None:
        return _run_presentation_command(parser, args)

    parser.error("missing command; run 'ppt --help'")
    return 2


def _reject_output(parser: argparse.ArgumentParser, command: str, output: str | None) -> None:
    if output is not None:
        parser.error(f"{command} does not accept --output")


def _reject_force(parser: argparse.ArgumentParser, command: str, force: bool) -> None:
    if force:
        parser.error(f"{command} does not accept --force")


def _run_presentation_command(parser: argparse.ArgumentParser, args: argparse.Namespace) -> int:
    _reject_force(parser, args.arg2, args.force)
    workspace_root = find_workspace_root()
    presentation = load_presentation_definition(workspace_root, args.subject)

    if args.arg2 == "check":
        _reject_output(parser, "check", args.output)
        if any(value is not None for value in (args.arg3, args.arg4, args.arg5)):
            parser.error("check does not accept additional arguments")
        check_presentation(presentation)
        print(f"{presentation.presentation_id}: ok")
        return 0

    if args.arg2 == "update":
        if any(value is not None for value in (args.arg3, args.arg4, args.arg5)):
            parser.error("update does not accept additional arguments")
        result = update_presentation(presentation, output=_output_path(args.output))
        for line in _replacement_lines(result.figure_outputs, result.stat_replacements, result.table_replacements):
            print(line)
        return 0

    if args.arg2 in {"data", "figure", "stat", "table"}:
        if args.arg3 == "list" and args.arg4 is None and args.arg5 is None:
            _reject_output(parser, f"{args.arg2} list", args.output)
            values = {
                "data": presentation.loaders,
                "figure": presentation.figures,
                "stat": presentation.stats,
                "table": presentation.tables,
            }[args.arg2]
            for item_id in sorted(values):
                print(item_id)
            return 0
        if args.arg2 == "figure" and args.arg3 == "update" and args.arg4 is None and args.arg5 is None:
            result = update_figures_in_deck(presentation)
            write_deck(presentation, result.deck, output=_output_path(args.output))
            for line in _replacement_lines(result.outputs, (), ()):
                print(line)
            return 0
        if args.arg2 == "figure" and args.arg4 == "update" and args.arg3 is not None and args.arg5 is None:
            result = update_figures_in_deck(presentation, figure_id=args.arg3)
            write_deck(presentation, result.deck, output=_output_path(args.output))
            for line in _replacement_lines(result.outputs, (), ()):
                print(line)
            return 0
        if args.arg2 == "stat" and args.arg3 == "update" and args.arg4 is None and args.arg5 is None:
            replacements = update_stats_to_output(presentation, output=_output_path(args.output))
            for line in _replacement_lines((), replacements, ()):
                print(line)
            return 0
        if args.arg2 == "stat" and args.arg4 == "update" and args.arg3 is not None and args.arg5 is None:
            replacements = update_stats_to_output(presentation, stat_id=args.arg3, output=_output_path(args.output))
            for line in _replacement_lines((), replacements, ()):
                print(line)
            return 0
        if args.arg2 == "table" and args.arg3 == "update" and args.arg4 is None and args.arg5 is None:
            replacements = update_tables_to_output(presentation, output=_output_path(args.output))
            for line in _replacement_lines((), (), replacements):
                print(line)
            return 0
        if args.arg2 == "table" and args.arg4 == "update" and args.arg3 is not None and args.arg5 is None:
            replacements = update_tables_to_output(presentation, table_id=args.arg3, output=_output_path(args.output))
            for line in _replacement_lines((), (), replacements):
                print(line)
            return 0
        parser.error(f"unsupported {args.arg2} command")

    parser.error(f"unsupported command '{args.arg2}'")
    return 2


def _output_path(value: str | None) -> Path | None:
    return Path(value) if value is not None else None


def _slide_line(slide_number: int, value: str) -> str:
    return f"Slide {slide_number}: {value}"


def _replacement_lines(
    figure_outputs: tuple[FigureOutput, ...],
    stat_replacements: tuple[StatReplacement, ...],
    table_replacements: tuple[TableReplacement, ...],
) -> list[str]:
    records = [
        (output.slide_number, output.shape_index, output.token, output.path.name)
        for output in figure_outputs
    ] + [
        (replacement.slide_number, replacement.shape_index, replacement.token, replacement.value)
        for replacement in stat_replacements
    ] + [
        (replacement.slide_number, replacement.shape_index, replacement.token, replacement.summary)
        for replacement in table_replacements
    ]
    records.sort(key=lambda item: (item[0], item[1]))
    totals = Counter((slide_number, token) for slide_number, _shape_index, token, _value in records)
    seen: defaultdict[tuple[int, str], int] = defaultdict(int)
    lines: list[str] = []
    for slide_number, _shape_index, token, value in records:
        key = (slide_number, token)
        seen[key] += 1
        suffix = f" [{seen[key]}/{totals[key]}]" if totals[key] > 1 else ""
        lines.append(_slide_line(slide_number, f"{token}{suffix} = {value}"))
    return lines
