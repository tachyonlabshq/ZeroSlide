use anyhow::Result;
use std::env;
use std::path::Path;
use zeroslide::schema::{PresentationSpec, SlideSpec};
use zeroslide::{
    add_agent_comment, add_slide, add_speaker_notes, create_presentation, extract_outline,
    inspect_presentation, inspect_slide, read_json_file, replace_slide_text, resolve_agent_comment,
    run_mcp_stdio, scan_agent_comments, schema_info, skill_api_contract,
};

fn main() {
    if let Err(err) = run() {
        eprintln!("zeroslide: {err}");
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        print_usage();
        return Ok(());
    }

    let pretty = args.iter().any(|arg| arg == "--pretty");

    match args[1].as_str() {
        "inspect-presentation" => {
            print_json(&inspect_presentation(required_arg(&args, 2)?)?, pretty)
        }
        "inspect-slide" => print_json(
            &inspect_slide(
                required_arg(&args, 2)?,
                parse_usize(required_arg(&args, 3)?, "slide_number")?,
            )?,
            pretty,
        ),
        "extract-outline" => print_json(&extract_outline(required_arg(&args, 2)?)?, pretty),
        "create-presentation" => {
            let spec: PresentationSpec = read_json_file(required_arg(&args, 2)?)?;
            print_json(
                &create_presentation(&spec, required_arg(&args, 3)?)?,
                pretty,
            )
        }
        "add-slide" => {
            let spec: SlideSpec = read_json_file(required_arg(&args, 3)?)?;
            print_json(
                &add_slide(required_arg(&args, 2)?, &spec, required_arg(&args, 4)?)?,
                pretty,
            )
        }
        "replace-slide-text" => {
            let spec: SlideSpec = read_json_file(required_arg(&args, 4)?)?;
            print_json(
                &replace_slide_text(
                    required_arg(&args, 2)?,
                    parse_usize(required_arg(&args, 3)?, "slide_number")?,
                    &spec,
                    required_arg(&args, 5)?,
                )?,
                pretty,
            )
        }
        "add-speaker-notes" => {
            let notes_arg = required_arg(&args, 4)?;
            let notes = if Path::new(notes_arg).exists() {
                std::fs::read_to_string(notes_arg)?
            } else {
                notes_arg.to_string()
            };
            print_json(
                &add_speaker_notes(
                    required_arg(&args, 2)?,
                    parse_usize(required_arg(&args, 3)?, "slide_number")?,
                    &notes,
                    required_arg(&args, 5)?,
                )?,
                pretty,
            )
        }
        "scan-agent-comments" => print_json(
            &scan_agent_comments(
                required_arg(&args, 2)?,
                args.iter().any(|arg| arg == "--include-resolved"),
            )?,
            pretty,
        ),
        "add-agent-comment" => {
            let author = flag_value(&args, "--author").unwrap_or("ZeroSlide");
            let initials = flag_value(&args, "--initials").unwrap_or("ZS");
            let x = flag_value(&args, "--x")
                .map(|value| parse_u32(value, "x"))
                .transpose()?
                .unwrap_or(0);
            let y = flag_value(&args, "--y")
                .map(|value| parse_u32(value, "y"))
                .transpose()?
                .unwrap_or(0);
            print_json(
                &add_agent_comment(
                    required_arg(&args, 2)?,
                    parse_usize(required_arg(&args, 3)?, "slide_number")?,
                    required_arg(&args, 4)?,
                    required_arg(&args, 5)?,
                    author,
                    initials,
                    x,
                    y,
                )?,
                pretty,
            )
        }
        "resolve-agent-comment" => {
            let author = flag_value(&args, "--author").unwrap_or("ZeroSlide");
            let initials = flag_value(&args, "--initials").unwrap_or("ZS");
            print_json(
                &resolve_agent_comment(
                    required_arg(&args, 2)?,
                    parse_usize(required_arg(&args, 3)?, "slide_number")?,
                    parse_u32(required_arg(&args, 4)?, "comment_index")?,
                    required_arg(&args, 5)?,
                    required_arg(&args, 6)?,
                    author,
                    initials,
                )?,
                pretty,
            )
        }
        "schema-info" => print_json(&schema_info(), pretty),
        "skill-api-contract" => print_json(&skill_api_contract(), pretty),
        "mcp-stdio" => run_mcp_stdio(pretty),
        _ => {
            print_usage();
            Ok(())
        }
    }
}

fn print_json<T: serde::Serialize>(value: &T, pretty: bool) -> Result<()> {
    if pretty {
        println!("{}", serde_json::to_string_pretty(value)?);
    } else {
        println!("{}", serde_json::to_string(value)?);
    }
    Ok(())
}

fn required_arg(args: &[String], index: usize) -> Result<&str> {
    args.get(index)
        .map(|value| value.as_str())
        .ok_or_else(|| anyhow::anyhow!("missing argument at position {}", index))
}

fn flag_value<'a>(args: &'a [String], flag: &str) -> Option<&'a str> {
    args.iter()
        .position(|arg| arg == flag)
        .and_then(|idx| args.get(idx + 1))
        .map(|value| value.as_str())
}

fn parse_usize(value: &str, field: &str) -> Result<usize> {
    value
        .parse::<usize>()
        .map_err(|_| anyhow::anyhow!("invalid integer for {field}: {value}"))
}

fn parse_u32(value: &str, field: &str) -> Result<u32> {
    value
        .parse::<u32>()
        .map_err(|_| anyhow::anyhow!("invalid integer for {field}: {value}"))
}

fn print_usage() {
    eprintln!(
        "Usage:
  zeroslide inspect-presentation <deck.pptx> [--pretty]
  zeroslide inspect-slide <deck.pptx> <slide_number> [--pretty]
  zeroslide extract-outline <deck.pptx> [--pretty]
  zeroslide create-presentation <spec.json> <output.pptx> [--pretty]
  zeroslide add-slide <input.pptx> <slide.json> <output.pptx> [--pretty]
  zeroslide replace-slide-text <input.pptx> <slide_number> <slide.json> <output.pptx> [--pretty]
  zeroslide add-speaker-notes <input.pptx> <slide_number> <notes-or-path> <output.pptx> [--pretty]
  zeroslide scan-agent-comments <deck.pptx> [--include-resolved] [--pretty]
  zeroslide add-agent-comment <input.pptx> <slide_number> <text> <output.pptx> [--author NAME] [--initials ZS] [--x 0] [--y 0] [--pretty]
  zeroslide resolve-agent-comment <input.pptx> <slide_number> <comment_index> <response> <output.pptx> [--author NAME] [--initials ZS] [--pretty]
  zeroslide schema-info [--pretty]
  zeroslide skill-api-contract [--pretty]
  zeroslide mcp-stdio [--pretty]"
    );
}
