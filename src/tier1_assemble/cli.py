from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional

from .assembler import COMPOSITE_POSITIONS, assemble_presentation, combine_map_legend, write_assembly_report
from .models import AssemblyOptions, AssemblyReport


def main(argv: Optional[list] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "assemble":
            report = assemble_presentation(
                Path(args.template),
                Path(args.manifest) if args.manifest else None,
                Path(args.output_dir),
                scratch_dir=Path(args.scratch_dir) if args.scratch_dir else None,
                asset_dir=Path(args.asset_dir) if args.asset_dir else None,
                run_id=args.run_id,
                output_path=Path(args.output) if args.output else None,
                options=AssemblyOptions(overwrite=args.overwrite),
            )
            if args.report:
                write_assembly_report(report, Path(args.report))
            _print_assembly_report(report, args.json)
            return 0 if report.valid else 1

        if args.command == "combine-map-legend":
            output_path = combine_map_legend(
                Path(args.map),
                Path(args.legend),
                args.position,
                output_path=Path(args.output) if args.output else None,
            )
            print(f"output_path: {output_path}")
            return 0

        parser.print_help()
        return 2
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tier1-assemble",
        description="Assemble BEAD Tier 1 PowerPoint drafts from tokenized templates and local assets.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    assemble = subparsers.add_parser("assemble", help="Assemble a Tier 1 PPTX from a tokenized template and assets.")
    assemble.add_argument("--template", required=True, help="Tokenized master PPTX template.")
    assemble.add_argument("--manifest", help="Assembly asset manifest JSON.")
    assemble.add_argument("--output-dir", required=True, help="Directory for the completed PPTX.")
    assemble.add_argument("--scratch-dir", help="Run scratch directory; manifest asset paths resolve relative to this.")
    assemble.add_argument(
        "--asset-dir",
        help=(
            "Directory of PNG/JPEG assets named after tokens. Defaults to "
            "placeholder_assets under --scratch-dir, or next to the template when --scratch-dir is omitted. "
            "Combined map+legend PNGs are also written here."
        ),
    )
    assemble.add_argument("--run-id", help="Run id override. Defaults to manifest run_id.")
    assemble.add_argument("--output", help="Optional explicit PPTX output path.")
    assemble.add_argument("--report", help="Optional JSON audit report path.")
    assemble.add_argument("--overwrite", action="store_true", help="Allow replacing an existing output file.")
    assemble.add_argument("--json", action="store_true", help="Print report as JSON.")

    composite = subparsers.add_parser("combine-map-legend", help="Create a map PNG with a cropped legend overlaid.")
    composite.add_argument("--map", required=True, help="Base map PNG/JPEG path.")
    composite.add_argument("--legend", required=True, help="Transparent legend PNG/JPEG path.")
    composite.add_argument("--position", required=True, choices=sorted(COMPOSITE_POSITIONS), help="Legend corner position.")
    composite.add_argument("--output", help="Optional output PNG path. Defaults to <map_stem>_with_legend.png.")
    return parser


def _print_assembly_report(report: AssemblyReport, as_json: bool) -> None:
    if as_json:
        print(json.dumps(report.as_dict(), indent=2, sort_keys=True))
        return
    print(f"valid: {'yes' if report.valid else 'no'}")
    print(f"template_path: {report.template_path}")
    print(f"manifest_path: {report.manifest_path}")
    print(f"output_path: {report.output_path}")
    print(f"run_id: {report.run_id}")
    print(f"slides_scanned: {report.slides_scanned}")
    print(f"tokens_found: {report.tokens_found}")
    print(f"assets_inserted: {report.assets_inserted}")
    print(f"derived_assets_created: {report.derived_assets_created}")
    if report.unresolved_tokens:
        print("unresolved_tokens:")
        for token in report.unresolved_tokens:
            print(f"  - {token}")
    if report.unused_assets:
        print("unused_assets:")
        for token in report.unused_assets:
            print(f"  - {token}")
    if report.warnings:
        print("warnings:")
        for warning in report.warnings:
            print(f"  - {warning}")
    if report.replacements:
        print("replacements:")
        for replacement in report.replacements:
            status = "inserted" if replacement.inserted else "not inserted"
            print(f"  - slide {replacement.slide_number}: {replacement.token} ({status})")
