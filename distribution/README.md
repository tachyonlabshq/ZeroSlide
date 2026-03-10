# Distribution Assets

Build a local platform bundle zip:

```bash
python3 scripts/build_platform_bundle.py \
  --platform macos-arm64 \
  --binary-path target/release/zeroslide \
  --output-root distribution/artifacts
```

Each bundle zip contains exactly one top-level `ZeroSlide/` folder with:

- `README.md`
- `SKILL.md`
- `mcp.json`
- `bin/zeroslide` or `bin/zeroslide.exe`

The script also writes:

- a per-platform manifest JSON
- a per-platform SHA256 checksum file

Generate an aggregate manifest and checksum index from a directory of per-platform outputs:

```bash
python3 scripts/build_bundle_index.py \
  --input-root distribution/artifacts \
  --output-root distribution/aggregate-artifacts \
  --version 0.1.0
```

For cross-platform binaries, use the GitHub Actions platform-bundle workflow in [.github/workflows/platform-bundles.yml](/Users/michaelwong/Developer/ZeroSlide/.github/workflows/platform-bundles.yml). It builds:

- macOS arm64 on `macos-15`
- macOS x64 on `macos-15-intel`
- Windows x64 on `ubuntu-latest` via `cargo-xwin`
- Windows arm64 on `ubuntu-latest` via `cargo-xwin`

Each matrix job packages the bundle, uploads the zip plus manifest/checksum artifacts, and the aggregate job emits a run-level manifest and `SHA256SUMS` file. Successful `main` and manual runs publish prerelease snapshot releases, while `v*` tag runs publish stable GitHub Releases.
