#!/usr/bin/env python3

import argparse
import hashlib
import json
import shutil
from pathlib import Path


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a local ZeroSlide skill bundle.")
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

    repo_root = Path(__file__).resolve().parent.parent
    binary = (repo_root / args.binary).resolve()
    if not binary.exists():
        raise SystemExit(f"Binary not found: {binary}")

    bundle_root = repo_root / args.output_root / args.target_name
    if bundle_root.exists():
        shutil.rmtree(bundle_root)
    (bundle_root / "bin").mkdir(parents=True)

    binary_name = args.target_name
    bundled_binary = bundle_root / "bin" / binary_name
    shutil.copy2(binary, bundled_binary)

    for rel_path in ["SKILL.md", "mcp.json", "README.md"]:
        shutil.copy2(repo_root / rel_path, bundle_root / Path(rel_path).name)

    manifest = {
        "name": "ZeroSlide",
        "target": args.target_name,
        "binary": f"bin/{binary_name}",
        "sha256": sha256(bundled_binary),
    }
    (bundle_root / "bundle_manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
