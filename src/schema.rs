use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideSpec {
    pub title: String,
    #[serde(default)]
    pub bullets: Vec<String>,
    #[serde(default)]
    pub notes: Option<String>,
    #[serde(default)]
    pub layout: Option<String>,
    #[serde(default)]
    pub comments: Vec<CommentInput>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresentationSpec {
    pub title: String,
    #[serde(default)]
    pub slides: Vec<SlideSpec>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentInput {
    pub text: String,
    #[serde(default)]
    pub author: Option<String>,
    #[serde(default)]
    pub initials: Option<String>,
    #[serde(default)]
    pub x: Option<u32>,
    #[serde(default)]
    pub y: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SchemaInfo {
    pub name: String,
    pub version: String,
    pub contract_version: String,
    pub transport: Vec<String>,
    pub commands: Vec<String>,
    pub mcp_tools: Vec<String>,
    pub comment_marker: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SkillApiContract {
    pub contract_version: String,
    pub schema_version: String,
    pub minimum_compatible_schema_version: String,
    pub stable_commands: Vec<String>,
    pub stable_mcp_tools: Vec<String>,
    pub followup_comment_rules: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresentationInspection {
    pub path: String,
    pub title: Option<String>,
    pub creator: Option<String>,
    pub slide_count: usize,
    pub total_comments: usize,
    pub total_agent_comments: usize,
    pub slides: Vec<SlideInspection>,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideInspection {
    pub slide_number: usize,
    pub title: Option<String>,
    pub body_text: Vec<String>,
    pub notes: Option<String>,
    pub shape_count: usize,
    pub table_count: usize,
    pub comment_count: usize,
    pub agent_comment_count: usize,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresentationOutline {
    pub title: Option<String>,
    pub slides: Vec<OutlineSlide>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresentationText {
    pub path: String,
    pub title: Option<String>,
    pub slide_text: Vec<SlideText>,
    pub combined_text: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideText {
    pub slide_number: usize,
    pub title: Option<String>,
    pub text: Vec<String>,
    pub notes: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OutlineSlide {
    pub slide_number: usize,
    pub title: Option<String>,
    pub bullets: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MutationSummary {
    pub input_path: Option<String>,
    pub output_path: String,
    pub action: String,
    pub slide_number: Option<usize>,
    pub details: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AgentCommentScan {
    pub path: String,
    pub aliases: Vec<String>,
    pub total_comments: usize,
    pub pending: Vec<AgentCommentRecord>,
    pub resolved: Vec<AgentCommentRecord>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AgentCommentRecord {
    pub slide_number: usize,
    pub comment_index: u32,
    pub author: Option<String>,
    pub initials: Option<String>,
    pub text: String,
    pub instruction: String,
    pub timestamp: Option<String>,
    pub x: u32,
    pub y: u32,
    pub resolved: bool,
}
