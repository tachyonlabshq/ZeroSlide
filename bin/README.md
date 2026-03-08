# Binary Artifacts

The checked-in local artifacts currently include:

- `zeroslide`
- `zeroslide-macos-arm64`

`ZeroSlide` is configured so the GitHub release workflow builds the remaining target binaries on GitHub runners:

- `zeroslide-macos-x64`
- `zeroslide-linux-x64`
- `zeroslide-windows-x64.exe`

The roadmap keeps the broader multi-platform matrix open so Linux arm64 and Windows arm64 artifacts can be added once the release pipeline is extended or cross-build tooling is standardized.
