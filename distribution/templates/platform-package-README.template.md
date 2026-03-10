# {{PROJECT_NAME}}

Install-ready skill bundle for `{{PLATFORM}}` (`{{VERSION}}`).

## Included files

- `README.md`
- `SKILL.md`
- `mcp.json`
- `bin/{{BINARY_NAME}}`

## Installation

1. Extract this zip.
2. Move the `{{PROJECT_NAME}}/` folder directly into your agent skills directory.
3. Keep the bundled layout intact so `mcp.json` can invoke `{{BINARY_RELPATH}}`.

## MCP entrypoint

`mcp.json` is preconfigured to launch the bundled binary locally:

```json
{
  "command": ["{{BINARY_RELPATH}}", "mcp-stdio"]
}
```

## Validation

From the extracted `{{PROJECT_NAME}}/` folder:

```bash
{{BINARY_RELPATH}} schema-info --pretty
```
