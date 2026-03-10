#!/usr/bin/env python3

import argparse
import subprocess
import sys
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser(description="Backward-compatible wrapper for build_platform_bundle.py.")
    parser.add_argument(
        "--binary",
        default="target/release/zeroslide",
        help="Path to the compiled zeroslide binary.",
    )
    parser.add_argument(
        "--target-name",
        required=True,
        help="Bundle target label, for example zeroslide-macos-arm64.",
    )
    parser.add_argument(
        "--output-root",
        default="distribution/bundles",
        help="Output folder for the bundle.",
    )
    args = parser.parse_args()

    platform = args.target_name.removeprefix("zeroslide-")
    repo_root = Path(__file__).resolve().parent.parent
    script = repo_root / "scripts" / "build_platform_bundle.py"
    command = [
        sys.executable,
        str(script),
        "--platform",
        platform,
        "--binary-path",
        args.binary,
        "--output-root",
        args.output_root,
        "--project-name",
        "ZeroSlide",
        "--binary-name",
        "zeroslide",
    ]
    subprocess.run(command, check=True, cwd=repo_root)


if __name__ == "__main__":
    main()
