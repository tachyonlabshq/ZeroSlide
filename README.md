# ZeroSlide

Rust-first PPTX tooling for AI agents.

`ZeroSlide` is the PowerPoint member of the Zero family: a single binary that works as:

- a CLI for inspecting and editing `.pptx` decks
- an MCP stdio server for OpenCode and other agent runtimes
- a local skill bundle with installable metadata and config templates

It uses [`ppt-rs`](https://github.com/yingkitw/ppt-rs) as the base presentation engine, then adds agent-oriented inspection, safe write flows, and a PowerPoint comment workflow built around `@Agent`.

## Current capabilities

- Inspect full decks and individual slides.
- Extract compact outlines for agent planning.
- Create decks from JSON specs.
- Append slides and replace generated slide text.
- Add or replace speaker notes.
- Scan PowerPoint comments for `@Agent` follow-up instructions.
- Add agent comments back into a deck.
- Mark handled comments as processed while preserving provenance.
- Run as an MCP stdio server with stable tool schemas.

## Build

```bash
cargo build --release
```

## CLI usage

Inspect a presentation:

```bash
./target/release/zeroslide inspect-presentation ./deck.pptx --pretty
./target/release/zeroslide inspect-slide ./deck.pptx 2 --pretty
./target/release/zeroslide extract-outline ./deck.pptx --pretty
```

Create or edit a deck:

```bash
./target/release/zeroslide create-presentation ./examples/presentation_spec.json ./out.pptx --pretty
./target/release/zeroslide add-slide ./out.pptx ./examples/slide_spec.json ./out-v2.pptx --pretty
./target/release/zeroslide replace-slide-text ./out-v2.pptx 2 ./examples/slide_replace.json ./out-v3.pptx --pretty
./target/release/zeroslide add-speaker-notes ./out-v3.pptx 1 "Rehearse this opener." ./out-v4.pptx --pretty
```

Use the `@Agent` comment workflow:

```bash
./target/release/zeroslide add-agent-comment ./out-v4.pptx 1 "@Agent tighten the headline." ./out-v5.pptx --author "Reviewer" --initials RV --pretty
./target/release/zeroslide scan-agent-comments ./out-v5.pptx --pretty
./target/release/zeroslide resolve-agent-comment ./out-v5.pptx 1 1 "Headline updated in the next revision." ./out-v6.pptx --pretty
```

Run as MCP:

```bash
./target/release/zeroslide mcp-stdio
```

## JSON presentation spec

```json
{
  "title": "Quarterly Review",
  "slides": [
    {
      "title": "Executive Summary",
      "bullets": [
        "Revenue grew 18% year over year",
        "Pipeline quality improved in enterprise accounts"
      ],
      "notes": "Lead with the margin expansion story.",
      "comments": [
        {
          "text": "@Agent check whether the title should reference FY26.",
          "author": "Michael",
          "initials": "MW",
          "x": 0,
          "y": 0
        }
      ]
    }
  ]
}
```

## MCP tools

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

## Validation in this iteration

- `cargo test`
- `cargo build --release`
- `cargo clippy --all-targets -- -D warnings`

## Distribution

- Skill metadata: [SKILL.md](/Users/michaelwong/Developer/ZeroSlide/SKILL.md)
- MCP template: [mcp.json](/Users/michaelwong/Developer/ZeroSlide/mcp.json)
- Contract docs: [docs/SKILL_API_CONTRACT.md](/Users/michaelwong/Developer/ZeroSlide/docs/SKILL_API_CONTRACT.md)
- Comment workflow: [docs/COMMENT_WORKFLOW.md](/Users/michaelwong/Developer/ZeroSlide/docs/COMMENT_WORKFLOW.md)
- Bundle guide: [distribution/README.md](/Users/michaelwong/Developer/ZeroSlide/distribution/README.md)
