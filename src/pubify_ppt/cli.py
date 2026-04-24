from __future__ import annotations

import argparse
import sys

from pubify_ppt.config import WORKSPACE_CONFIG_SECTION


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
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the ``ppt`` CLI and return its process exit code."""

    parser = build_parser()
    parser.parse_args(sys.argv[1:] if argv is None else argv)
    parser.error("Phase 0 provides the CLI entrypoint; commands arrive in later phases")
    return 2
