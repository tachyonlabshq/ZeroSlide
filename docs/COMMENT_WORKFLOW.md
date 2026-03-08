# ZeroSlide Comment Workflow

## Goal

Let humans leave actionable follow-up instructions inside PowerPoint using normal review comments.

## Human side

1. Add a PowerPoint comment to a slide.
2. Start the actionable part with `@Agent`.
3. Save the deck.

Example:

```text
@Agent rewrite this slide for a board audience and shorten the title.
```

## Agent side

1. Run `scan-agent-comments`.
2. Process only comments that contain a configured alias such as `@Agent`.
3. Keep author attribution and timestamp in any downstream plan or audit log.
4. After handling the request, run `resolve-agent-comment`.

## Speaker-notes fallback

If a deck should avoid classic PowerPoint comments, `ZeroSlide` can store the `@Agent` inbox in speaker notes instead:

1. Use `--fallback-mode notes` on `add-agent-comment`, `scan-agent-comments`, and `resolve-agent-comment`.
2. `ZeroSlide` only uses this fallback when the source deck does not already contain classic comment structures.
3. Visible speaker notes stay readable; the agent inbox is stored in a hidden serialized block that `ZeroSlide` preserves across note edits.

## Resolution behavior

`resolve-agent-comment` does two things:

1. Appends `[ZeroSlide: processed]` plus the response text to the original comment.
2. Adds a new agent-authored reply comment to the slide.

This keeps the original request visible while allowing later scans to treat it as handled.

In speaker-notes fallback mode, the original `@Agent` entry is marked as processed in the notes inbox block instead of creating a native PowerPoint reply comment.
