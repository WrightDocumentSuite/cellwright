#!/usr/bin/env python3
"""Reject missing/failed corpus runs and any test skip not explicitly allow-listed."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--results", type=Path, required=True)
    args = parser.parse_args()

    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    allowed_skips = set(manifest["allowedSkippedTests"])
    result_files = sorted(args.results.rglob("*.trx"))
    if not result_files:
        raise RuntimeError(f"No TRX files found below {args.results}")

    minimum_by_file = manifest["minimumResultsByFile"]
    found_by_name: dict[str, list[Path]] = {}
    for result_file in result_files:
        found_by_name.setdefault(result_file.name, []).append(result_file)

    missing_files = sorted(set(minimum_by_file) - set(found_by_name))
    unexpected_files = sorted(set(found_by_name) - set(minimum_by_file))
    duplicate_files = sorted(name for name, paths in found_by_name.items() if len(paths) > 1)
    if missing_files:
        raise RuntimeError(f"Required TRX files are missing: {', '.join(missing_files)}")
    if unexpected_files:
        raise RuntimeError(f"Unexpected TRX files make the gate ambiguous: {', '.join(unexpected_files)}")
    if duplicate_files:
        raise RuntimeError(f"Duplicate TRX file names make the gate ambiguous: {', '.join(duplicate_files)}")

    total = 0
    failures: list[str] = []
    unexpected_skips: list[str] = []
    allowed_seen: list[str] = []

    for result_file in result_files:
        tree = ET.parse(result_file)
        file_total = 0
        for element in tree.iter():
            if local_name(element.tag) != "UnitTestResult":
                continue
            total += 1
            file_total += 1
            name = element.attrib.get("testName", "<unnamed>")
            outcome = element.attrib.get("outcome", "").lower()
            if outcome in {"passed", "completed"}:
                continue
            if outcome in {"notexecuted", "skipped"}:
                if any(name == allowed or name.startswith(allowed + "(") for allowed in allowed_skips):
                    allowed_seen.append(name)
                else:
                    unexpected_skips.append(name)
            else:
                failures.append(f"{name} ({outcome or 'unknown outcome'})")

        if result_file.name in minimum_by_file and file_total < int(minimum_by_file[result_file.name]):
            failures.append(
                f"{result_file.name} contains {file_total} results; "
                f"expected at least {minimum_by_file[result_file.name]}"
            )

    minimum = int(manifest["minimumTestResults"])
    print(
        f"Read {total} results from {len(result_files)} TRX files; "
        f"{len(allowed_seen)} allowed skips, {len(unexpected_skips)} unexpected skips, {len(failures)} failures"
    )
    if total < minimum:
        failures.append(f"Only {total} test results were recorded; expected at least {minimum}")
    if unexpected_skips:
        failures.extend(f"Unexpected skipped test: {name}" for name in unexpected_skips)
    if failures:
        raise RuntimeError("\n".join(failures))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, ET.ParseError) as error:
        print(f"error: {error}", file=sys.stderr)
        raise SystemExit(1)
