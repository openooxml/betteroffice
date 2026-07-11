//! the `xlsx` binary: a thin arg parser over `xlsx_cli`'s library path, with
//! two subcommands — `render` (a region to png) and `info` (sheet summaries).

use std::path::{Path, PathBuf};
use std::process::ExitCode;

use xlsx_cli::{RenderOptions, load_workbook, render, resolve_sheet, sheet_summaries};
use xlsx_model::CellRange;

const USAGE: &str = "\
xlsx — headless renderer for OpenOOXML spreadsheets

usage:
  xlsx render <file.xlsx> [options]
  xlsx info <file.xlsx>

render options:
  -o, --output <path>   png output path (default: <input>.png)
      --sheet <s>       sheet name or 0-based index (default: first sheet)
      --range <A1:H40>  cell range to render (default: the sheet's used range)
      --scale <n>       output scale factor, e.g. 2 for hidpi (default: 1)
      --width <px>      cap output width; crops the region to fit
      --height <px>     cap output height; crops the region to fit
  -h, --help            show this help";

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("render") => run(cmd_render(&args[1..])),
        Some("info") => run(cmd_info(&args[1..])),
        Some("-h") | Some("--help") | None => {
            println!("{USAGE}");
            ExitCode::SUCCESS
        }
        Some(other) => {
            eprintln!("error: unknown command {other:?}\n\n{USAGE}");
            ExitCode::from(2)
        }
    }
}

/// map a command result to stderr + exit code.
fn run(result: Result<(), String>) -> ExitCode {
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("error: {e}");
            ExitCode::from(1)
        }
    }
}

fn cmd_render(args: &[String]) -> Result<(), String> {
    let mut input: Option<PathBuf> = None;
    let mut output: Option<PathBuf> = None;
    let mut sheet: Option<String> = None;
    let mut range: Option<String> = None;
    let mut opts = RenderOptions::default();

    let mut it = args.iter();
    while let Some(arg) = it.next() {
        match arg.as_str() {
            "-o" | "--output" => output = Some(PathBuf::from(next_value(&mut it, arg)?)),
            "--sheet" => sheet = Some(next_value(&mut it, arg)?),
            "--range" => range = Some(next_value(&mut it, arg)?),
            "--scale" => {
                let v = next_value(&mut it, arg)?;
                opts.scale = v.parse().map_err(|_| format!("invalid --scale: {v:?}"))?;
            }
            "--width" => {
                let v = next_value(&mut it, arg)?;
                opts.max_width = Some(v.parse().map_err(|_| format!("invalid --width: {v:?}"))?);
            }
            "--height" => {
                let v = next_value(&mut it, arg)?;
                opts.max_height = Some(v.parse().map_err(|_| format!("invalid --height: {v:?}"))?);
            }
            "-h" | "--help" => {
                println!("{USAGE}");
                return Ok(());
            }
            other if other.starts_with('-') => return Err(format!("unknown flag {other:?}")),
            _ => {
                if input.replace(PathBuf::from(arg)).is_some() {
                    return Err("render takes a single input file".to_string());
                }
            }
        }
    }

    let input = input.ok_or("render needs an input .xlsx file")?;
    if let Some(r) = &range {
        opts.range =
            Some(CellRange::parse_a1(r).map_err(|e| format!("invalid --range {r:?}: {e}"))?);
    }

    let bytes = std::fs::read(&input).map_err(|e| format!("reading {}: {e}", input.display()))?;
    let wb = load_workbook(&bytes)?;
    let sheet_id = resolve_sheet(&wb, sheet.as_deref())?;
    let png = render(&wb, sheet_id, &opts)?;

    let output = output.unwrap_or_else(|| default_output(&input));
    std::fs::write(&output, &png.bytes)
        .map_err(|e| format!("writing {}: {e}", output.display()))?;
    println!(
        "wrote {} ({}x{}px, {} bytes)",
        output.display(),
        png.width,
        png.height,
        png.bytes.len()
    );
    Ok(())
}

fn cmd_info(args: &[String]) -> Result<(), String> {
    let mut input: Option<PathBuf> = None;
    for arg in args {
        match arg.as_str() {
            "-h" | "--help" => {
                println!("{USAGE}");
                return Ok(());
            }
            other if other.starts_with('-') => return Err(format!("unknown flag {other:?}")),
            _ => {
                if input.replace(PathBuf::from(arg)).is_some() {
                    return Err("info takes a single input file".to_string());
                }
            }
        }
    }
    let input = input.ok_or("info needs an input .xlsx file")?;
    let bytes = std::fs::read(&input).map_err(|e| format!("reading {}: {e}", input.display()))?;
    let wb = load_workbook(&bytes)?;

    println!("{}: {} sheet(s)", input.display(), wb.sheets.len());
    for s in sheet_summaries(&wb) {
        let range = s.used_range.as_deref().unwrap_or("(empty)");
        println!(
            "  [{}] {:?}  used={}  cells={}",
            s.index, s.name, range, s.cell_count
        );
    }
    Ok(())
}

/// pull the value following a flag, erroring if it's missing.
fn next_value(it: &mut std::slice::Iter<'_, String>, flag: &str) -> Result<String, String> {
    it.next()
        .cloned()
        .ok_or_else(|| format!("{flag} needs a value"))
}

/// default png path: the input with its extension swapped to `.png`.
fn default_output(input: &Path) -> PathBuf {
    input.with_extension("png")
}
