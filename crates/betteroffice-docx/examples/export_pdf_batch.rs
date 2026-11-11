//! Export every DOCX fixture under a directory, recording a JSON manifest.
//!
//! ```text
//! cargo run -p betteroffice-docx --example export_pdf_batch -- <input_dir> <output_dir>
//! ```
//!
//! Set `BO_PDF_FALLBACK_FONT=<path>` to register one extra face (for example
//! Noto Color Emoji or a CJK face) in every document before export. Registered
//! fallback faces cover characters the document's own fonts lack.
//!
//! The scan is recursive and case-insensitive for the `.docx` extension. Word
//! lock files (`~$*.docx`) and the usual build directories are skipped. Each
//! document is exported to the same relative path under `output_dir` with a
//! `.pdf` extension, and `<output_dir>/manifest.json` records one entry per
//! fixture:
//!
//! ```json
//! {
//!   "path": "tables/merged.docx",
//!   "status": "ok",
//!   "pages": 2,
//!   "pdf_bytes": 34812,
//!   "fonts_embedded": 2,
//!   "text_chars": 517,
//!   "error": null
//! }
//! ```
//!
//! `fonts_embedded` counts embedded font-program streams (`/FontFile*`);
//! `text_chars` counts authored characters in the body, headers, footers and
//! notes, not characters recovered from the PDF. `pages` comes from the PDF
//! page tree. The command exits 1 when any fixture fails and still writes the
//! manifest for the fixtures that succeeded.

use std::fs;
use std::io;
use std::path::{Path, PathBuf};

use betteroffice_docx::{BlockContent, Document, get_paragraph_text};
use serde_json::{Map, Value, json};

const SKIP_DIRS: &[&str] = &[".git", "target", "node_modules", ".source"];

fn main() {
    if let Err(error) = run() {
        eprintln!("export_pdf_batch: {error}");
        std::process::exit(2);
    }
}

fn run() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let input_dir = args
        .next()
        .ok_or("usage: export_pdf_batch <input_dir> <output_dir>")?;
    let output_dir = args
        .next()
        .ok_or("usage: export_pdf_batch <input_dir> <output_dir>")?;
    if args.next().is_some() {
        return Err("usage: export_pdf_batch <input_dir> <output_dir>".into());
    }

    let input_dir = PathBuf::from(input_dir);
    let output_dir = PathBuf::from(output_dir);
    let fallback_font = std::env::var_os("BO_PDF_FALLBACK_FONT")
        .map(|path| fs::read(PathBuf::from(path)))
        .transpose()?;
    if !input_dir.is_dir() {
        return Err(format!("{} is not a directory", input_dir.display()).into());
    }
    fs::create_dir_all(&output_dir)?;

    let mut fixtures = Vec::new();
    collect_docx(&input_dir, &input_dir, &mut fixtures)?;
    fixtures.sort_by(|left, right| left.0.cmp(&right.0));
    if fixtures.is_empty() {
        return Err(format!("no .docx fixtures under {}", input_dir.display()).into());
    }

    let total = fixtures.len();
    let mut records = Vec::with_capacity(total);
    let mut failures = 0;
    for (index, (relative, source)) in fixtures.into_iter().enumerate() {
        let mut record = Map::new();
        record.insert("path".into(), json!(portable(&relative)));

        let result = export_one(&source, &output_dir, &relative, fallback_font.as_deref());
        match result {
            Ok((pages, bytes, fonts, text_chars)) => {
                eprintln!(
                    "[{}/{total}] ok   {} ({} page(s), {} bytes)",
                    index + 1,
                    portable(&relative),
                    pages.map_or_else(|| "?".to_owned(), |pages| pages.to_string()),
                    bytes
                );
                record.insert("status".into(), json!("ok"));
                record.insert("pages".into(), json!(pages));
                record.insert("pdf_bytes".into(), json!(bytes));
                record.insert("fonts_embedded".into(), json!(fonts));
                record.insert("text_chars".into(), json!(text_chars));
                record.insert("error".into(), Value::Null);
            }
            Err(error) => {
                failures += 1;
                eprintln!(
                    "[{}/{total}] FAIL {}: {error}",
                    index + 1,
                    portable(&relative)
                );
                record.insert("status".into(), json!("error"));
                record.insert("pages".into(), Value::Null);
                record.insert("pdf_bytes".into(), Value::Null);
                record.insert("fonts_embedded".into(), Value::Null);
                record.insert("text_chars".into(), Value::Null);
                record.insert("error".into(), json!(error));
            }
        }
        records.push(Value::Object(record));
    }

    let manifest = json!({
        "tool": "export_pdf_batch",
        "input_dir": absolute(&input_dir),
        "output_dir": absolute(&output_dir),
        "fixtures": records,
    });
    let manifest_path = output_dir.join("manifest.json");
    fs::write(&manifest_path, serde_json::to_vec_pretty(&manifest)?)?;
    println!(
        "exported {} fixture(s) to {}: {} ok, {} failed; manifest {}",
        total,
        output_dir.display(),
        total - failures,
        failures,
        manifest_path.display()
    );

    if failures > 0 {
        std::process::exit(1);
    }
    Ok(())
}

type Exported = (Option<usize>, usize, Option<usize>, Option<usize>);

fn export_one(
    source: &Path,
    output_dir: &Path,
    relative: &Path,
    fallback_font: Option<&[u8]>,
) -> Result<Exported, String> {
    let bytes = fs::read(source).map_err(|error| format!("read: {error}"))?;
    let mut document = Document::open(&bytes).map_err(|error| format!("open: {error}"))?;
    if let Some(font) = fallback_font {
        document
            .register_font("PDF Fallback", false, false, font)
            .map_err(|error| format!("register fallback font: {error}"))?;
    }
    let text_chars = authored_text_chars(&document);
    // Compute the PDF even when layout fails so the manifest records which
    // stage failed rather than collapsing every failure into "open".
    let pdf = document
        .export_pdf()
        .map_err(|error| format!("export: {error}"))?;
    if !pdf.starts_with(b"%PDF-") {
        return Err("export produced a document without a %PDF- header".into());
    }
    let pdf_path = output_dir.join(relative).with_extension("pdf");
    if let Some(parent) = pdf_path.parent() {
        fs::create_dir_all(parent).map_err(|error| format!("mkdir: {error}"))?;
    }
    fs::write(&pdf_path, &pdf).map_err(|error| format!("write: {error}"))?;
    Ok((
        pdf_page_count(&pdf),
        pdf.len(),
        Some(embedded_font_count(&pdf)),
        Some(text_chars),
    ))
}

fn collect_docx(
    root: &Path,
    directory: &Path,
    fixtures: &mut Vec<(PathBuf, PathBuf)>,
) -> io::Result<()> {
    let mut entries: Vec<_> = fs::read_dir(directory)?.collect::<Result<_, _>>()?;
    entries.sort_by_key(|entry| entry.file_name());
    for entry in entries {
        let path = entry.path();
        let metadata = fs::metadata(&path)?;
        if metadata.is_dir() {
            let name = entry.file_name();
            if SKIP_DIRS.iter().any(|skip| name == *skip) {
                continue;
            }
            collect_docx(root, &path, fixtures)?;
        } else if metadata.is_file() && is_fixture(&path) {
            let relative = path
                .strip_prefix(root)
                .map_err(|error| io::Error::other(error.to_string()))?
                .to_path_buf();
            fixtures.push((relative, path));
        }
    }
    Ok(())
}

fn is_fixture(path: &Path) -> bool {
    let name = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("");
    if name.starts_with("~$") {
        return false;
    }
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
}

/// Characters of authored story text: body, headers, footers, footnotes,
/// endnotes. This is a stable source-side count; PDF extraction is checked by
/// `scripts/office-quality/pdf_matrix.py`.
fn authored_text_chars(document: &Document) -> usize {
    let model = document.model();
    let mut chars = block_text_chars(&model.body.content);
    for (_, header) in &model.headers {
        chars += block_text_chars(&header.content);
    }
    for (_, footer) in &model.footers {
        chars += block_text_chars(&footer.content);
    }
    for note in &model.footnotes {
        if !note.is_separator() {
            chars += block_text_chars(&note.content);
        }
    }
    for note in &model.endnotes {
        if !note.is_separator() {
            chars += block_text_chars(&note.content);
        }
    }
    chars
}

fn block_text_chars(blocks: &[BlockContent]) -> usize {
    let mut chars = 0;
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => {
                chars += get_paragraph_text(paragraph).chars().count();
            }
            BlockContent::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        chars += block_text_chars(&cell.content);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => {
                chars += block_text_chars(&sdt.content);
            }
            BlockContent::RawXml(_) => {}
        }
    }
    chars
}

/// The page tree's `/Count` value. The writer emits a single page tree, so the
/// largest `/Count` in the file is its page count.
fn pdf_page_count(pdf: &[u8]) -> Option<usize> {
    let mut count = None;
    let mut index = 0;
    while let Some(found) = find(&pdf[index..], b"/Count") {
        let mut cursor = index + found + b"/Count".len();
        while cursor < pdf.len() && pdf[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let start = cursor;
        while cursor < pdf.len() && pdf[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if let Ok(value) = std::str::from_utf8(&pdf[start..cursor])
            .unwrap_or_default()
            .parse::<usize>()
        {
            count = Some(count.map_or(value, |current: usize| current.max(value)));
        }
        index += found + b"/Count".len();
    }
    count
}

/// Embedded font programs: `/FontFile2` (TrueType), `/FontFile3` (CFF/OpenType)
/// and the older bare `/FontFile` key.
fn embedded_font_count(pdf: &[u8]) -> usize {
    let mut count = count_token(pdf, b"/FontFile2") + count_token(pdf, b"/FontFile3");
    let mut index = 0;
    while let Some(found) = find(&pdf[index..], b"/FontFile") {
        let after = index + found + b"/FontFile".len();
        let is_variant = pdf.get(after).is_some_and(|byte| byte.is_ascii_digit());
        if !is_variant {
            count += 1;
        }
        index = after;
    }
    count
}

fn count_token(haystack: &[u8], needle: &[u8]) -> usize {
    let mut count = 0;
    let mut index = 0;
    while let Some(found) = find(&haystack[index..], needle) {
        count += 1;
        index += found + needle.len();
    }
    count
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn portable(path: &Path) -> String {
    path.components()
        .map(|component| component.as_os_str().to_string_lossy())
        .collect::<Vec<_>>()
        .join("/")
}

fn absolute(path: &Path) -> String {
    fs::canonicalize(path)
        .map(|path| path.display().to_string())
        .unwrap_or_else(|_| path.display().to_string())
}
