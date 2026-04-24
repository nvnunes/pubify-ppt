from __future__ import annotations

import argparse
from pathlib import Path
import sys

from pubify_ppt.config import WORKSPACE_CONFIG_SECTION
from pubify_ppt.config import find_workspace_root
from pubify_ppt.discovery import list_presentation_ids, load_presentation_definition
from pubify_ppt.init import init_presentation_by_id, init_workspace
from pubify_ppt.runtime import check_presentation


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
                "  ppt <presentation-id> stat list",
                "  ppt <presentation-id> stat update [--output <path>]",
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
    except Exception as exc:
        parser.error(str(exc))

    parser.error("command is not implemented yet")
    return 2


def _reject_output(parser: argparse.ArgumentParser, command: str, output: str | None) -> None:
    if output is not None:
        parser.error(f"{command} does not accept --output")


def _reject_force(parser: argparse.ArgumentParser, command: str, force: bool) -> None:
    if force:
        parser.error(f"{command} does not accept --force")


def _run_presentation_command(parser: argparse.ArgumentParser, args: argparse.Namespace) -> int:
    _reject_output(parser, args.arg2, args.output)
    _reject_force(parser, args.arg2, args.force)
    workspace_root = find_workspace_root()
    presentation = load_presentation_definition(workspace_root, args.subject)

    if args.arg2 == "check":
        if any(value is not None for value in (args.arg3, args.arg4, args.arg5)):
            parser.error("check does not accept additional arguments")
        check_presentation(presentation)
        print(f"{presentation.presentation_id}: ok")
        return 0

    if args.arg2 in {"data", "figure", "stat"}:
        if args.arg3 != "list" or args.arg4 is not None or args.arg5 is not None:
            parser.error(f"{args.arg2} supports only 'list' in this phase")
        values = {
            "data": presentation.loaders,
            "figure": presentation.figures,
            "stat": presentation.stats,
        }[args.arg2]
        for item_id in sorted(values):
            print(item_id)
        return 0

    parser.error(f"unsupported command '{args.arg2}'")
    return 2
