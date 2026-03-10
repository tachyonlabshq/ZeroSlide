#!/usr/bin/env python3

import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import tempfile
import zipfile
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # pragma: no cover
    tomllib = None


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def cargo_package_metadata(repo_root: Path) -> dict[str, str]:
    cargo_toml = repo_root / "Cargo.toml"
    if not cargo_toml.exists() or tomllib is None:
        return {}
    data = tomllib.loads(read_text(cargo_toml))
    package = data.get("package", {})
    return {
        "name": package.get("name", ""),
        "version": package.get("version", ""),
    }


def bundled_binary_name(binary_name: str, platform: str) -> str:
    if platform.startswith("windows-") and not binary_name.endswith(".exe"):
        return f"{binary_name}.exe"
    return binary_name


def release_version(args_version: str | None, env_version: str | None, cargo_version: str) -> str:
    if args_version:
        return args_version
    if env_version:
        return env_version
    return cargo_version or "dev"


def normalize_version(raw_version: str) -> str:
    return raw_version[1:] if raw_version.startswith("v") else raw_version


def template_vars(
    *,
    project_name: str,
    binary_name: str,
    platform: str,
    version: str,
    binary_relpath: str,
) -> dict[str, str]:
    return {
        "PROJECT_NAME": project_name,
        "BINARY_NAME": binary_name,
        "PLATFORM": platform,
        "VERSION": version,
        "BINARY_RELPATH": binary_relpath,
    }


def render_template(path: Path, substitutions: dict[str, str]) -> str:
    content = read_text(path)
    for key, value in substitutions.items():
        content = content.replace(f"{{{{{key}}}}}", value)
    return content


def safe_unlink(path: Path) -> None:
    if path.exists():
        path.unlink()


def add_file_to_zip(archive: zipfile.ZipFile, source: Path, arcname: str) -> None:
    file_stat = source.stat()
    zipinfo = zipfile.ZipInfo.from_file(source, arcname)
    permissions = stat.S_IMODE(file_stat.st_mode) or 0o644
    zipinfo.external_attr = permissions << 16
    with source.open("rb") as handle:
        archive.writestr(zipinfo, handle.read(), compress_type=zipfile.ZIP_DEFLATED)


def validate_top_level_folder(zip_path: Path, expected_folder: str) -> None:
    with zipfile.ZipFile(zip_path) as archive:
        names = archive.namelist()
    top_levels = {Path(name).parts[0] for name in names if name}
    if top_levels != {expected_folder}:
        raise SystemExit(
            f"Bundle must contain exactly one top-level folder named {expected_folder}, got {sorted(top_levels)}"
        )


def build_bundle(args: argparse.Namespace) -> None:
    repo_root = Path(__file__).resolve().parent.parent
    cargo_metadata = cargo_package_metadata(repo_root)

    project_name = args.project_name or os.environ.get("PROJECT_NAME") or "ZeroSlide"
    binary_name = args.binary_name or os.environ.get("BINARY_NAME") or cargo_metadata.get("name") or "zeroslide"
    version = release_version(
        args.version,
        os.environ.get("RELEASE_VERSION"),
        cargo_metadata.get("version", ""),
    )
    normalized_version = normalize_version(version)
    output_root = (repo_root / args.output_root).resolve()
    binary_path = (repo_root / args.binary_path).resolve()
    template_root = (repo_root / args.template_root).resolve()

    if not binary_path.exists():
        raise SystemExit(f"Binary not found: {binary_path}")

    if not template_root.exists():
        raise SystemExit(f"Template root not found: {template_root}")

    readme_template = template_root / "platform-package-README.template.md"
    mcp_template = template_root / "platform-package-mcp.template.json"
    for required in [readme_template, mcp_template]:
        if not required.exists():
            raise SystemExit(f"Template not found: {required}")

    output_root.mkdir(parents=True, exist_ok=True)

    packaged_binary_name = bundled_binary_name(binary_name, args.platform)
    binary_relpath = f"./bin/{packaged_binary_name}"
    zip_base_name = f"{project_name}-{args.platform}-{normalized_version}"
    zip_path = output_root / f"{zip_base_name}.zip"
    manifest_path = output_root / f"{zip_base_name}.manifest.json"
    checksum_path = output_root / f"{zip_base_name}.sha256"

    safe_unlink(zip_path)
    safe_unlink(manifest_path)
    safe_unlink(checksum_path)

    with tempfile.TemporaryDirectory(prefix="zeroslide-bundle-") as temp_dir:
        staging_root = Path(temp_dir) / project_name
        (staging_root / "bin").mkdir(parents=True)

        bundled_binary_path = staging_root / "bin" / packaged_binary_name
        shutil.copy2(binary_path, bundled_binary_path)

        substitutions = template_vars(
            project_name=project_name,
            binary_name=packaged_binary_name,
            platform=args.platform,
            version=normalized_version,
            binary_relpath=binary_relpath,
        )

        write_text(staging_root / "README.md", render_template(readme_template, substitutions))
        write_text(staging_root / "mcp.json", render_template(mcp_template, substitutions))
        shutil.copy2(repo_root / "SKILL.md", staging_root / "SKILL.md")

        staged_files = sorted(path for path in staging_root.rglob("*") if path.is_file())
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for staged_file in staged_files:
                arcname = staged_file.relative_to(staging_root.parent).as_posix()
                add_file_to_zip(archive, staged_file, arcname)

    validate_top_level_folder(zip_path, project_name)

    manifest = {
        "project_name": project_name,
        "platform": args.platform,
        "version": normalized_version,
        "zip_file": zip_path.name,
        "bundle_folder": project_name,
        "binary": f"bin/{packaged_binary_name}",
        "binary_command": [binary_relpath, "mcp-stdio"],
        "files": [
            "README.md",
            "SKILL.md",
            "mcp.json",
            f"bin/{packaged_binary_name}",
        ],
        "sha256": sha256(zip_path),
    }
    write_text(manifest_path, json.dumps(manifest, indent=2) + "\n")
    write_text(checksum_path, f"{manifest['sha256']}  {zip_path.name}\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a platform-specific Zero family skill bundle zip.")
    parser.add_argument("--platform", required=True, help="Platform label such as macos-arm64 or windows-x64.")
    parser.add_argument("--binary-path", required=True, help="Path to the compiled binary to package.")
    parser.add_argument(
        "--output-root",
        required=True,
        help="Directory that will receive the generated zip, manifest, and checksum.",
    )
    parser.add_argument("--project-name", help="Skill folder and zip prefix name. Defaults to ZeroSlide.")
    parser.add_argument("--binary-name", help="Bundled binary name without platform suffix.")
    parser.add_argument("--version", help="Version string. Defaults to RELEASE_VERSION or Cargo.toml version.")
    parser.add_argument(
        "--template-root",
        default="distribution/templates",
        help="Template directory containing platform-package templates.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)+", args.platform):
        raise SystemExit(f"Invalid platform label: {args.platform}")
    build_bundle(args)


if __name__ == "__main__":
    main()
