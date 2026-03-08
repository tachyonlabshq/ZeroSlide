use anyhow::{Context, Result, anyhow};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use std::io::{self, Write};

use crate::ops::{
    add_agent_comment, add_slide, add_speaker_notes, append_bullets, create_presentation,
    extract_outline, extract_text, inspect_presentation, inspect_slide, remove_slide,
    reorder_slides, resolve_agent_comment, scan_agent_comments, schema_info, skill_api_contract,
};
use crate::schema::{PresentationSpec, SlideSpec};

#[derive(Debug, Clone, Deserialize)]
pub struct McpRequest {
    pub id: Option<Value>,
    pub method: String,
    pub params: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
pub struct McpResponse {
    pub jsonrpc: String,
    pub id: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub result: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<McpError>,
}

#[derive(Debug, Clone, Serialize)]
pub struct McpError {
    pub code: i64,
    pub message: String,
}

#[derive(Debug, Clone, Serialize)]
struct ToolDescriptor {
    name: String,
    description: String,
    #[serde(rename = "inputSchema")]
    input_schema: Value,
}

#[derive(Debug, Clone, Deserialize)]
struct ToolCallParams {
    name: String,
    arguments: Option<Value>,
}

pub fn run_mcp_stdio(pretty: bool) -> Result<()> {
    let mut stdout = io::stdout();
    let stdin = io::stdin();
    let stream = serde_json::Deserializer::from_reader(stdin.lock()).into_iter::<Value>();

    for item in stream {
        let response = match item {
            Ok(raw) => match serde_json::from_value::<McpRequest>(raw) {
                Ok(request) => handle_request(request),
                Err(err) => error_response(None, -32700, format!("invalid request: {err}")),
            },
            Err(err) => error_response(None, -32700, format!("invalid request stream: {err}")),
        };
        let payload = if pretty {
            serde_json::to_string_pretty(&response)?
        } else {
            serde_json::to_string(&response)?
        };
        writeln!(stdout, "{payload}").context("failed to write MCP response")?;
        stdout.flush().context("failed to flush MCP stdout")?;
    }

    Ok(())
}

fn handle_request(request: McpRequest) -> McpResponse {
    let result = match request.method.as_str() {
        "initialize" => Ok(json!({
            "protocolVersion": "2024-11-05",
            "capabilities": { "tools": {} },
            "serverInfo": {
                "name": "zeroslide",
                "version": env!("CARGO_PKG_VERSION")
            }
        })),
        "notifications/initialized" | "notifications/cancelled" => Ok(json!(null)),
        "tools/list" => Ok(json!({ "tools": list_tools() })),
        "tools/call" => call_tool(request.params),
        other => Err(anyhow!("unsupported method '{other}'")),
    };

    match result {
        Ok(value) => McpResponse {
            jsonrpc: "2.0".to_string(),
            id: request.id,
            result: Some(value),
            error: None,
        },
        Err(err) => error_response(request.id, -32000, err.to_string()),
    }
}

fn error_response(id: Option<Value>, code: i64, message: String) -> McpResponse {
    McpResponse {
        jsonrpc: "2.0".to_string(),
        id,
        result: None,
        error: Some(McpError { code, message }),
    }
}

fn call_tool(params: Option<Value>) -> Result<Value> {
    let params = params.ok_or_else(|| anyhow!("tools/call requires params"))?;
    let parsed: ToolCallParams =
        serde_json::from_value(params).context("invalid tools/call payload")?;
    let args = parsed.arguments.unwrap_or_else(|| json!({}));
    let object = args
        .as_object()
        .ok_or_else(|| anyhow!("tool arguments must be an object"))?;

    let value = match parsed.name.as_str() {
        "inspect_presentation" => {
            serde_json::to_value(inspect_presentation(required_string(object, "path")?)?)?
        }
        "inspect_slide" => serde_json::to_value(inspect_slide(
            required_string(object, "path")?,
            required_usize(object, "slide_number")?,
        )?)?,
        "extract_text" => serde_json::to_value(extract_text(required_string(object, "path")?)?)?,
        "extract_outline" => {
            serde_json::to_value(extract_outline(required_string(object, "path")?)?)?
        }
        "create_presentation" => {
            let spec: PresentationSpec =
                serde_json::from_value(required_value(object, "spec")?.clone())
                    .context("invalid presentation spec")?;
            serde_json::to_value(create_presentation(
                &spec,
                required_string(object, "output_path")?,
            )?)?
        }
        "add_slide" => {
            let spec: SlideSpec = serde_json::from_value(required_value(object, "spec")?.clone())
                .context("invalid slide spec")?;
            serde_json::to_value(add_slide(
                required_string(object, "input_path")?,
                &spec,
                required_string(object, "output_path")?,
            )?)?
        }
        "append_bullets" => {
            let bullets: Vec<String> =
                serde_json::from_value(required_value(object, "bullets")?.clone())
                    .context("invalid bullets array")?;
            serde_json::to_value(append_bullets(
                required_string(object, "input_path")?,
                required_usize(object, "slide_number")?,
                &bullets,
                required_string(object, "output_path")?,
            )?)?
        }
        "remove_slide" => serde_json::to_value(remove_slide(
            required_string(object, "input_path")?,
            required_usize(object, "slide_number")?,
            required_string(object, "output_path")?,
        )?)?,
        "reorder_slides" => {
            let order: Vec<usize> =
                serde_json::from_value(required_value(object, "order")?.clone())
                    .context("invalid slide order array")?;
            serde_json::to_value(reorder_slides(
                required_string(object, "input_path")?,
                &order,
                required_string(object, "output_path")?,
            )?)?
        }
        "replace_slide_text" => {
            let spec: SlideSpec = serde_json::from_value(required_value(object, "spec")?.clone())
                .context("invalid slide spec")?;
            serde_json::to_value(crate::ops::replace_slide_text(
                required_string(object, "input_path")?,
                required_usize(object, "slide_number")?,
                &spec,
                required_string(object, "output_path")?,
            )?)?
        }
        "add_speaker_notes" => serde_json::to_value(add_speaker_notes(
            required_string(object, "input_path")?,
            required_usize(object, "slide_number")?,
            required_string(object, "notes")?,
            required_string(object, "output_path")?,
        )?)?,
        "scan_agent_comments" => serde_json::to_value(scan_agent_comments(
            required_string(object, "path")?,
            object
                .get("include_resolved")
                .and_then(Value::as_bool)
                .unwrap_or(false),
        )?)?,
        "add_agent_comment" => serde_json::to_value(add_agent_comment(
            required_string(object, "input_path")?,
            required_usize(object, "slide_number")?,
            required_string(object, "text")?,
            required_string(object, "output_path")?,
            optional_string(object, "author").unwrap_or("ZeroSlide"),
            optional_string(object, "initials").unwrap_or("ZS"),
            optional_u32(object, "x").unwrap_or(0),
            optional_u32(object, "y").unwrap_or(0),
        )?)?,
        "resolve_agent_comment" => serde_json::to_value(resolve_agent_comment(
            required_string(object, "input_path")?,
            required_usize(object, "slide_number")?,
            required_u32(object, "comment_index")?,
            required_string(object, "response")?,
            required_string(object, "output_path")?,
            optional_string(object, "author").unwrap_or("ZeroSlide"),
            optional_string(object, "initials").unwrap_or("ZS"),
        )?)?,
        "schema_info" => serde_json::to_value(schema_info())?,
        "skill_api_contract" => serde_json::to_value(skill_api_contract())?,
        other => return Err(anyhow!("unsupported tool '{other}'")),
    };

    Ok(json!({
        "content": [
            {
                "type": "text",
                "text": serde_json::to_string_pretty(&value)?
            }
        ],
        "structuredContent": value
    }))
}

fn required_value<'a>(object: &'a serde_json::Map<String, Value>, key: &str) -> Result<&'a Value> {
    object
        .get(key)
        .ok_or_else(|| anyhow!("missing required argument '{key}'"))
}

fn required_string<'a>(object: &'a serde_json::Map<String, Value>, key: &str) -> Result<&'a str> {
    object
        .get(key)
        .and_then(Value::as_str)
        .filter(|value| !value.trim().is_empty())
        .ok_or_else(|| anyhow!("missing required string argument '{key}'"))
}

fn optional_string<'a>(object: &'a serde_json::Map<String, Value>, key: &str) -> Option<&'a str> {
    object.get(key).and_then(Value::as_str)
}

fn required_usize(object: &serde_json::Map<String, Value>, key: &str) -> Result<usize> {
    object
        .get(key)
        .and_then(Value::as_u64)
        .map(|value| value as usize)
        .ok_or_else(|| anyhow!("missing required integer argument '{key}'"))
}

fn required_u32(object: &serde_json::Map<String, Value>, key: &str) -> Result<u32> {
    object
        .get(key)
        .and_then(Value::as_u64)
        .map(|value| value as u32)
        .ok_or_else(|| anyhow!("missing required integer argument '{key}'"))
}

fn optional_u32(object: &serde_json::Map<String, Value>, key: &str) -> Option<u32> {
    object
        .get(key)
        .and_then(Value::as_u64)
        .map(|value| value as u32)
}

fn list_tools() -> Vec<ToolDescriptor> {
    vec![
        ToolDescriptor {
            name: "inspect_presentation".to_string(),
            description:
                "Inspect a PowerPoint deck and return per-slide text, notes, and comment counts."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "path": { "type": "string", "description": "Path to a .pptx file." }
                },
                "required": ["path"]
            }),
        },
        ToolDescriptor {
            name: "inspect_slide".to_string(),
            description: "Inspect one slide by 1-based slide number.".to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "path": { "type": "string" },
                    "slide_number": { "type": "integer", "minimum": 1 }
                },
                "required": ["path", "slide_number"]
            }),
        },
        ToolDescriptor {
            name: "extract_outline".to_string(),
            description:
                "Extract deck titles and bullet outlines in a compact agent-friendly format."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "path": { "type": "string" }
                },
                "required": ["path"]
            }),
        },
        ToolDescriptor {
            name: "extract_text".to_string(),
            description: "Extract combined slide text and notes from a deck.".to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "path": { "type": "string" }
                },
                "required": ["path"]
            }),
        },
        ToolDescriptor {
            name: "create_presentation".to_string(),
            description: "Create a new PowerPoint deck from a JSON presentation spec.".to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "output_path": { "type": "string" },
                    "spec": { "type": "object" }
                },
                "required": ["output_path", "spec"]
            }),
        },
        ToolDescriptor {
            name: "add_slide".to_string(),
            description: "Append a new slide to an existing deck and write a new output file."
                .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "spec": { "type": "object" }
                },
                "required": ["input_path", "output_path", "spec"]
            }),
        },
        ToolDescriptor {
            name: "append_bullets".to_string(),
            description: "Append bullet points to an existing slide and write a new output file."
                .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "slide_number": { "type": "integer", "minimum": 1 },
                    "bullets": {
                        "type": "array",
                        "items": { "type": "string" }
                    }
                },
                "required": ["input_path", "output_path", "slide_number", "bullets"]
            }),
        },
        ToolDescriptor {
            name: "remove_slide".to_string(),
            description:
                "Remove one slide and write a new output file while preserving surviving metadata."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "slide_number": { "type": "integer", "minimum": 1 }
                },
                "required": ["input_path", "output_path", "slide_number"]
            }),
        },
        ToolDescriptor {
            name: "reorder_slides".to_string(),
            description:
                "Reorder all slides according to a 1-based permutation and write a new output file."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "order": {
                        "type": "array",
                        "items": { "type": "integer", "minimum": 1 }
                    }
                },
                "required": ["input_path", "output_path", "order"]
            }),
        },
        ToolDescriptor {
            name: "replace_slide_text".to_string(),
            description: "Replace one slide's generated text content in a deck.".to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "slide_number": { "type": "integer", "minimum": 1 },
                    "spec": { "type": "object" }
                },
                "required": ["input_path", "output_path", "slide_number", "spec"]
            }),
        },
        ToolDescriptor {
            name: "add_speaker_notes".to_string(),
            description: "Add or replace speaker notes for a slide.".to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "slide_number": { "type": "integer", "minimum": 1 },
                    "notes": { "type": "string" }
                },
                "required": ["input_path", "output_path", "slide_number", "notes"]
            }),
        },
        ToolDescriptor {
            name: "scan_agent_comments".to_string(),
            description: "Scan classic PowerPoint comments for @Agent follow-up requests."
                .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "path": { "type": "string" },
                    "include_resolved": { "type": "boolean" }
                },
                "required": ["path"]
            }),
        },
        ToolDescriptor {
            name: "add_agent_comment".to_string(),
            description: "Append a PowerPoint comment to a slide using the classic comment format."
                .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "slide_number": { "type": "integer", "minimum": 1 },
                    "text": { "type": "string" },
                    "author": { "type": "string" },
                    "initials": { "type": "string" },
                    "x": { "type": "integer", "minimum": 0 },
                    "y": { "type": "integer", "minimum": 0 }
                },
                "required": ["input_path", "output_path", "slide_number", "text"]
            }),
        },
        ToolDescriptor {
            name: "resolve_agent_comment".to_string(),
            description: "Mark an @Agent comment as processed and append an agent reply comment."
                .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "slide_number": { "type": "integer", "minimum": 1 },
                    "comment_index": { "type": "integer", "minimum": 1 },
                    "response": { "type": "string" },
                    "author": { "type": "string" },
                    "initials": { "type": "string" }
                },
                "required": ["input_path", "output_path", "slide_number", "comment_index", "response"]
            }),
        },
        ToolDescriptor {
            name: "schema_info".to_string(),
            description: "Return the ZeroSlide command and tool inventory.".to_string(),
            input_schema: json!({ "type": "object", "properties": {} }),
        },
        ToolDescriptor {
            name: "skill_api_contract".to_string(),
            description: "Return the stable skill and MCP compatibility contract.".to_string(),
            input_schema: json!({ "type": "object", "properties": {} }),
        },
    ]
}
