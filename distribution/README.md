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

For cross-platform binaries, use the GitHub Actions platform-bundle workflow. The public repo keeps source, packaging templates, workflow automation, and release artifacts in one place.
