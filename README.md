# ZeroSlide

Rust-first PPTX tooling for AI agents.

`ZeroSlide` is the PowerPoint member of the Zero family: a single binary that works as:

- a CLI for inspecting and editing `.pptx` decks
- an MCP stdio server for OpenCode and other agent runtimes
- a local skill bundle with installable metadata and config templates

It uses [`ppt-rs`](https://github.com/yingkitw/ppt-rs) as the base presentation engine, then adds agent-oriented inspection, safe write flows, and a PowerPoint comment workflow built around `@Agent`.

## Current capabilities

- Inspect full decks and individual slides.
- Extract compact outlines and combined text for agent planning.
- Produce an interoperability report for PowerPoint, Google Slides import, and LibreOffice.
- Create decks from JSON specs.
- Append slides, append bullets, remove slides, reorder slides, and replace generated slide text.
- Add or replace speaker notes.
- Scan PowerPoint comments for `@Agent` follow-up instructions.
- Optionally fall back to a speaker-notes or custom-metadata inbox when classic comment structures are absent.
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
./target/release/zeroslide extract-text ./deck.pptx --pretty
./target/release/zeroslide extract-outline ./deck.pptx --pretty
./target/release/zeroslide interop-report ./deck.pptx --pretty
```

Create or edit a deck:

```bash
./target/release/zeroslide create-presentation ./examples/presentation_spec.json ./out.pptx --pretty
./target/release/zeroslide add-slide ./out.pptx ./examples/slide_spec.json ./out-v2.pptx --pretty
./target/release/zeroslide append-bullets ./out-v2.pptx 2 ./examples/append_bullets.json ./out-v3.pptx --pretty
./target/release/zeroslide remove-slide ./out-v3.pptx 1 ./out-v4.pptx --pretty
./target/release/zeroslide reorder-slides ./out-v4.pptx ./examples/reorder_slides.json ./out-v5.pptx --pretty
./target/release/zeroslide replace-slide-text ./out-v5.pptx 1 ./examples/slide_replace.json ./out-v6.pptx --pretty
./target/release/zeroslide add-speaker-notes ./out-v6.pptx 1 "Rehearse this opener." ./out-v7.pptx --pretty
```

Use the `@Agent` comment workflow:

```bash
./target/release/zeroslide add-agent-comment ./out-v7.pptx 1 "@Agent tighten the headline." ./out-v8.pptx --author "Reviewer" --initials RV --pretty
./target/release/zeroslide scan-agent-comments ./out-v8.pptx --pretty
./target/release/zeroslide resolve-agent-comment ./out-v8.pptx 1 1 "Headline updated in the next revision." ./out-v9.pptx --pretty
```

If a deck must avoid classic PowerPoint comments, opt into one of the fallback inbox modes:

```bash
./target/release/zeroslide add-agent-comment ./out-v7.pptx 1 "@Agent tighten the headline." ./out-v8.pptx --author "Reviewer" --initials RV --fallback-mode notes --pretty
./target/release/zeroslide scan-agent-comments ./out-v8.pptx --fallback-mode notes --pretty
./target/release/zeroslide resolve-agent-comment ./out-v8.pptx 1 1 "Headline updated in the next revision." ./out-v9.pptx --fallback-mode notes --pretty
./target/release/zeroslide add-agent-comment ./out-v7.pptx 1 "@Agent tighten the headline." ./out-v8b.pptx --author "Reviewer" --initials RV --fallback-mode metadata --pretty
```

`--fallback-mode notes` and `--fallback-mode metadata` are only used when the source deck does not already contain classic comment structures. Existing native comments continue to use the PowerPoint comment system.

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
- `extract_text`
- `extract_outline`
- `interop_report`
- `create_presentation`
- `add_slide`
- `append_bullets`
- `remove_slide`
- `reorder_slides`
- `replace_slide_text`
- `add_speaker_notes`
- `scan_agent_comments`
- `add_agent_comment`
- `resolve_agent_comment`
- `schema_info`
- `skill_api_contract`

For OpenCode-class agents, prefer the MCP tools when the runtime exposes them. Use the CLI as a fallback when the agent cannot call MCP directly or when debugging local behavior.

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
