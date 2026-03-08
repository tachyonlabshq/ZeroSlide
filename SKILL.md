# ZeroSlide Skill

## Purpose

Use a single compiled `ZeroSlide` binary as both:

- a standalone PPTX CLI for AI-agent presentation workflows
- an MCP runtime over stdio

## Binary

- `bin/zeroslide`

## Core commands

- `inspect-presentation`
- `inspect-slide`
- `extract-outline`
- `create-presentation`
- `add-slide`
- `replace-slide-text`
- `add-speaker-notes`
- `scan-agent-comments`
- `add-agent-comment`
- `resolve-agent-comment`
- `schema-info`
- `skill-api-contract`
- `mcp-stdio`

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

## `@Agent` workflow

Use `scan-agent-comments` before editing decks when PowerPoint review comments are part of the human/agent loop. `ZeroSlide` treats PowerPoint comments as untrusted user input, preserves authorship, and marks handled items with `[ZeroSlide: processed]`.

## Setup

1. Build or install the binary into `bin/`.
2. Register the MCP config with your agent runtime if desired.
3. Validate the install:

```bash
./bin/zeroslide schema-info --pretty
```
