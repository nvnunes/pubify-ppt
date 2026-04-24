from __future__ import annotations

import argparse
from pathlib import Path
import sys

from pubify_ppt.config import WORKSPACE_CONFIG_SECTION
from pubify_ppt.config import find_workspace_root
from pubify_ppt.init import init_presentation_by_id, init_workspace


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
        if args.subject == "init":
            _reject_output(parser, "init", args.output)
            if args.force:
                parser.error("init does not accept --force")
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
    except Exception as exc:
        parser.error(str(exc))

    parser.error("command is not implemented yet")
    return 2


def _reject_output(parser: argparse.ArgumentParser, command: str, output: str | None) -> None:
    if output is not None:
        parser.error(f"{command} does not accept --output")
