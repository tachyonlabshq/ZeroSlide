pub mod mcp;
pub mod ops;
pub mod schema;

pub use mcp::run_mcp_stdio;
pub use ops::{
    add_agent_comment, add_slide, add_speaker_notes, append_bullets, create_presentation,
    extract_outline, extract_text, inspect_presentation, inspect_slide, interop_report,
    read_json_file, remove_slide, reorder_slides, replace_slide_text, resolve_agent_comment,
    scan_agent_comments, schema_info, skill_api_contract,
};
