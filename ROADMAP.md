# ZeroSlide Roadmap

## Phase 1 - Repository Foundation
- [ ] Initialize the Rust workspace for `zeroslide_pptx_agent` with a hardened release profile (`lto`, stripped symbols, abort-on-panic) and a clean `src/` module split for CLI, MCP, PPTX operations, schema, errors, and packaging helpers.
- [ ] Mirror the successful `ZeroCell` repository conventions where useful:
  - source and tests in a single repo
  - `docs/`, `distribution/`, `security/`, `scripts/`, and `examples/` folders
  - public-facing skill metadata and MCP config templates
- [ ] Vendor or pin `ppt-rs` as the base PPTX engine through a git dependency first, then evaluate whether any required fixes should be upstreamed or patched locally.
- [ ] Define a stable product identity for the Zero family:
  - binary name: `zeroslide`
  - MCP server name: `zeroslide`
  - skill name: `ZeroSlide`
  - contract versioning documented from the start

## Phase 2 - Presentation Read And Inspection Core
- [ ] Implement presentation inspection commands on top of `ppt-rs` reader APIs:
  - `inspect-presentation`
  - `inspect-slide`
  - `extract-text`
  - `extract-outline`
- [ ] Normalize slide metadata into agent-friendly JSON:
  - deck title and core properties
  - slide count and per-slide index/id
  - titles, bullets, body text, tables, notes presence, shape counts
  - warnings for empty slides, duplicate titles, or unreadable content
- [ ] Add resilient fallbacks for imperfect files:
  - corrupted or partial slides
  - missing title placeholders
  - nonstandard OOXML relationship layouts
  - graceful handling of notes/comments parts that are absent
- [ ] Build compact outputs optimized for LLM use so agents can inspect a deck without consuming the entire document payload at once.

## Phase 3 - Authoring And Update Operations
- [ ] Implement creation commands for new decks and appended slides using `ppt-rs` generation primitives:
  - `new-presentation`
  - `add-slide`
  - `add-title-slide`
  - `add-bullets-slide`
  - `add-speaker-notes`
- [ ] Implement safe edit operations for existing decks:
  - `replace-slide-text`
  - `append-bullets`
  - `remove-slide`
  - `reorder-slides`
- [ ] Prefer non-destructive writes:
  - read from source deck
  - apply requested changes
  - write to a new output path by default
  - allow in-place writes only behind an explicit flag
- [ ] Track deck-level and slide-level change summaries so agent runtimes can explain exactly what changed.

## Phase 4 - Agent Follow-Up Comments (`@Agent`)
- [ ] Add first-class support for PowerPoint review comments as an agent inbox layer.
- [ ] Implement `scan-agent-comments` to discover PowerPoint comments that contain an `@Agent` mention and emit a structured queue:
  - slide number
  - author
  - timestamp if available
  - comment text
  - extracted instruction body after the mention
  - anchor position if present
- [ ] Implement `add-agent-comment` so an agent can leave a response or follow-up request directly in the `.pptx` comment system.
- [ ] Implement `resolve-agent-comment` / `mark-agent-comment-processed` conventions:
  - preserve the original user comment
  - append an agent response comment or status marker
  - avoid destructive deletion unless explicitly requested
- [ ] Support a pragmatic interoperability mode:
  - use native PPTX comments when present
  - fall back to speaker notes or a custom metadata part only if the source file lacks comment structures and the caller explicitly allows fallback behavior
- [ ] Define the `@Agent` parsing rules:
  - support `@Agent`, `@agent`, and configurable aliases
  - preserve the original freeform comment text
  - ignore false positives inside unrelated text when mention parsing is disabled

## Phase 5 - MCP And Skill Surface
- [ ] Implement an MCP stdio server with a stable tool set designed for OpenCode but generic enough for other agent runtimes.
- [ ] Initial MCP tools:
  - `inspect_presentation`
  - `inspect_slide`
  - `extract_outline`
  - `create_presentation`
  - `add_slide`
  - `replace_slide_text`
  - `add_speaker_notes`
  - `scan_agent_comments`
  - `add_agent_comment`
  - `resolve_agent_comment`
  - `schema_info`
  - `skill_api_contract`
- [ ] Keep tool schemas explicit and narrow so agent clients can reliably discover the right operation.
- [ ] Publish a `SKILL.md` plus `mcp.json` template for local OpenCode installation and compatible MCP clients.
- [ ] Document command aliases and transport behavior so the same binary works as:
  - standalone CLI
  - MCP stdio server
  - installable skill bundle

## Phase 6 - Packaging And Distribution
- [ ] Keep one repository for source, release automation, and prepackaged binaries.
- [ ] Create a `bin/` layout for release artifacts:
  - macOS arm64
  - macOS x64
  - Linux x64
  - Linux arm64
  - Windows x64
  - Windows arm64
- [ ] Add scripts to assemble a single-binary skill bundle similar to `ZeroCell`, including:
  - binary
  - `SKILL.md`
  - `mcp.json` or MCP template config
  - install helper
  - manifest with checksums
- [ ] Add GitHub Actions workflows for matrix builds and artifact publishing so the checked-in public repo and tagged releases stay synchronized.
- [ ] Decide whether committed binaries live directly under `bin/` on `main` or in a release branch/artifact sync workflow; document the rule clearly before the first public release.

## Phase 7 - Testing And Validation
- [ ] Build a layered test suite:
  - unit tests for schema, mention parsing, and error handling
  - golden tests for PPTX read/write behavior
  - integration tests for CLI and MCP request/response flows
  - packaging tests for skill bundle manifests
- [ ] Create fixture decks that cover edge cases:
  - empty deck
  - missing titles
  - notes only
  - comments only
  - comments with `@Agent`
  - multi-comment thread on the same slide
  - special characters, Unicode, and XML-escaping cases
  - damaged or partially valid PPTX zip structures
- [ ] Validate generated files in PowerPoint-compatible environments where feasible:
  - PowerPoint
  - LibreOffice Impress
  - Google Slides import
- [ ] Add regression tests for comment preservation so agent comment operations never silently strip unrelated review threads.

## Phase 8 - Security And Supply Chain
- [ ] Add dependency and license gates:
  - `cargo audit`
  - `cargo deny`
  - lockfile review
- [ ] Add secure file-handling rules:
  - reject zip-slip style path traversal when unpacking PPTX internals
  - bound XML and ZIP processing to avoid trivial resource-exhaustion cases
  - sanitize generated XML content and comment text
- [ ] Add release integrity artifacts:
  - checksums
  - SBOM generation
  - optional signing/attestation workflow
- [ ] Review the `@Agent` comment flow for prompt-injection risk and document agent-side handling guidance:
  - comments are untrusted user instructions
  - tool output should preserve provenance and author identity
  - auto-apply behaviors must remain opt-in

## Phase 9 - Documentation And Ecosystem Fit
- [ ] Write a concise root `README.md` focused on agent workflows instead of generic PPTX marketing.
- [ ] Publish `docs/SKILL_API_CONTRACT.md` with stable command/tool schemas and compatibility guarantees.
- [ ] Publish `docs/COMMENT_WORKFLOW.md` for how humans and agents collaborate through PowerPoint comments.
- [ ] Provide end-to-end examples for:
  - creating a deck from JSON
  - scanning `@Agent` comments
  - replying in-thread with agent output
  - wiring `ZeroSlide` into OpenCode and other MCP-capable agents

## Phase 10 - Release And Publication
- [ ] Create the Git repository and push to `tachyonlabshq/ZeroSlide`.
- [ ] Tag the first public milestone once the following are green:
  - build
  - tests
  - security gates
  - packaging validation
  - MCP handshake validation
- [ ] Publish a release with binaries, checksums, and install instructions.
- [ ] Capture post-release follow-up items for upstream `ppt-rs` contributions or internal patches that should be reduced over time.
