#!/usr/bin/env python3
"""Fail when NuGet reports a direct or transitive package advisory."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import subprocess
import sys


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("projects", type=Path, nargs="+")
    args = parser.parse_args()

    findings: list[str] = []
    inspected_projects = 0
    for project_or_solution in args.projects:
        command = [
            "dotnet",
            "package",
            "list",
            "--project",
            str(project_or_solution),
            "--vulnerable",
            "--include-transitive",
            "--format",
            "json",
            "--output-version",
            "1",
            "--no-restore",
        ]
        result = subprocess.run(command, check=False, capture_output=True, text=True)
        if result.returncode != 0:
            sys.stdout.write(result.stdout)
            sys.stderr.write(result.stderr)
            raise RuntimeError(
                f"dotnet package list failed for {project_or_solution} with {result.returncode}"
            )

        start = result.stdout.find("{")
        if start < 0:
            raise RuntimeError(f"dotnet package list returned no JSON for {project_or_solution}")
        report = json.loads(result.stdout[start:])
        inspected_projects += len(report.get("projects", []))

        for project in report.get("projects", []):
            for framework in project.get("frameworks", []):
                packages = framework.get("topLevelPackages", []) + framework.get("transitivePackages", [])
                for package in packages:
                    for advisory in package.get("vulnerabilities", []):
                        advisory_url = (
                            advisory.get("advisoryurl")
                            or advisory.get("advisoryUrl")
                            or "<unknown advisory>"
                        )
                        findings.append(
                            f"{project.get('path', '<unknown project>')}: "
                            f"{package.get('id', '<unknown package>')} "
                            f"{package.get('resolvedVersion', '<unknown version>')} "
                            f"[{advisory.get('severity', 'unknown')}] {advisory_url}"
                        )

    if findings:
        raise RuntimeError("NuGet vulnerabilities found:\n" + "\n".join(findings))
    print(f"NuGet audit is clean across {inspected_projects} projects")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, json.JSONDecodeError) as error:
        print(f"error: {error}", file=sys.stderr)
        raise SystemExit(1)
