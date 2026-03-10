#!/usr/bin/env python3

import argparse
import hashlib
import json
from pathlib import Path


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_manifest(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build aggregate manifests and checksums for platform bundles.")
    parser.add_argument("--input-root", required=True, help="Directory containing per-platform bundle files.")
    parser.add_argument("--output-root", required=True, help="Directory to write aggregate outputs into.")
    parser.add_argument(
        "--version",
        help="Only include manifests for the specified normalized version and reject mixed-version inputs.",
    )
    parser.add_argument(
        "--manifest-name",
        default="ZeroSlide-bundles.manifest.json",
        help="Filename for the aggregate manifest JSON.",
    )
    parser.add_argument(
        "--checksums-name",
        default="ZeroSlide-bundles.SHA256SUMS",
        help="Filename for the aggregate checksum file.",
    )
    args = parser.parse_args()

    input_root = Path(args.input_root).resolve()
    output_root = Path(args.output_root).resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    manifests = sorted(input_root.glob("*.manifest.json"))
    if not manifests:
        raise SystemExit(f"No per-platform manifests found in {input_root}")

    bundles = [read_manifest(path) for path in manifests]
    if args.version:
        bundles = [bundle for bundle in bundles if bundle["version"] == args.version]
    if not bundles:
        raise SystemExit(
            f"No per-platform manifests matched version {args.version!r} in {input_root}"
            if args.version
            else f"No per-platform manifests found in {input_root}"
        )

    versions = {bundle["version"] for bundle in bundles}
    if len(versions) != 1:
        raise SystemExit(f"Aggregate manifest input must contain one version, found {sorted(versions)}")
    bundles.sort(key=lambda item: item["platform"])

    aggregate_manifest = {
        "project_name": bundles[0]["project_name"],
        "version": bundles[0]["version"],
        "bundle_count": len(bundles),
        "bundles": bundles,
    }
    manifest_path = output_root / args.manifest_name
    write_text(manifest_path, json.dumps(aggregate_manifest, indent=2) + "\n")

    checksum_lines = []
    for bundle in bundles:
        zip_path = input_root / bundle["zip_file"]
        if not zip_path.exists():
            raise SystemExit(f"Bundle zip referenced by manifest is missing: {zip_path}")
        checksum_lines.append(f"{sha256(zip_path)}  {zip_path.name}")
    checksum_path = output_root / args.checksums_name
    write_text(checksum_path, "\n".join(checksum_lines) + "\n")


if __name__ == "__main__":
    main()
