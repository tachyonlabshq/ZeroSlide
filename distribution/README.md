# Distribution Assets

Build a local skill bundle:

```bash
python3 scripts/build_release_bundle.py --target-name zeroslide-macos-arm64
```

The bundle contains:

- a copied binary under `distribution/bundles/<target>/bin/`
- `SKILL.md`
- `mcp.json`
- a release manifest with checksums

For cross-platform binaries, use the GitHub Actions release workflow. The public repo keeps source, packaging metadata, and release assets in one place.
