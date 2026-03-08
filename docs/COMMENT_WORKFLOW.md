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

## Resolution behavior

`resolve-agent-comment` does two things:

1. Appends `[ZeroSlide: processed]` plus the response text to the original comment.
2. Adds a new agent-authored reply comment to the slide.

This keeps the original request visible while allowing later scans to treat it as handled.
