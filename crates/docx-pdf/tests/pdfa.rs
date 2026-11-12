//! PDF/A-2u acceptance for the public exporter.
//!
//! The structural half is always exercised: header version, XMP identification,
//! output intent, document information, embedded font programs, Unicode glyph
//! mappings and link annotation flags. The external half runs veraPDF only when
//! `VERAPDF` names a veraPDF executable, so the default test run stays offline
//! and deterministic:
//!
//! ```sh
//! VERAPDF=/path/to/verapdf cargo test -p betteroffice-docx-pdf --test pdfa
//! ```
//!
//! `scripts/office-quality/pdfa_matrix.sh` sets `VERAPDF` in CI.

use std::collections::BTreeMap;
use std::io::Read;
use std::process::Command;

use betteroffice_docx_pdf::{
    PdfConformance, PdfExportOptions, PdfMetadata, export_pdf_with_options,
};
use docx_layout::display_list::DisplayList;
use flate2::read::ZlibDecoder;
use ooxml_text::FontStore;
use serde_json::{Value, json};
use ttf_parser::Face;

/// Text covering letters, digits and punctuation. Every mapped scalar must
/// resolve to a real glyph in the fixture font.
const TEXT: &str = "PDF/A-2u";
const LINK: &str = "https://example.com/pdfa-2u";
const PAGE_WIDTH: f64 = 612.0;
const PAGE_HEIGHT: f64 = 792.0;

fn font_bytes() -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/fonts/assets/Roboto-Regular.ttf");
    match std::fs::read(&path) {
        Ok(bytes) => bytes,
        Err(error) => panic!("read font fixture {}: {error}", path.display()),
    }
}

/// Position one glyph per source character at a fixed pen advance.
fn placed_glyphs(bytes: &[u8]) -> Vec<Value> {
    let face = Face::parse(bytes, 0).expect("parse font fixture");
    let mut glyphs = Vec::new();
    for (cluster, character) in TEXT.char_indices() {
        let Some(glyph) = face.glyph_index(character) else {
            panic!("font fixture has no glyph for {character:?}");
        };
        assert_ne!(glyph.0, 0, "fixture text mapped {character:?} to .notdef");
        glyphs.push(json!({
            "id": u32::from(glyph.0),
            "x": 72.0 + cluster as f64 * 8.0,
            "y": 96.0,
            "cluster": cluster as u32,
            "advance": 8.0,
        }));
    }
    glyphs
}

/// A one-page display list with a linked, shaped text run.
fn sample() -> Vec<u8> {
    let bytes = font_bytes();
    let mut store = FontStore::default();
    let font_id = store
        .register(bytes.clone())
        .expect("register font")
        .to_u32();
    let display_list: DisplayList = serde_json::from_value(json!({
        "pages": [{
            "pageIndex": 0,
            "width": PAGE_WIDTH,
            "height": PAGE_HEIGHT,
            "primitives": [{
                "kind": "glyphRun",
                "fontId": font_id,
                "size": 16.0,
                "color": "#112233",
                "text": TEXT,
                "glyphs": placed_glyphs(&bytes),
                "href": LINK,
                "linkTitle": "PDF/A-2u link",
            }],
        }],
    }))
    .expect("display list JSON");

    let options = PdfExportOptions {
        conformance: PdfConformance::PdfA2u,
        metadata: PdfMetadata {
            title: Some("PDF/A-2u acceptance".to_owned()),
            author: Some("BetterOffice tests".to_owned()),
            subject: Some("archival export".to_owned()),
            keywords: Some("pdfa, acceptance".to_owned()),
            creator: Some("docx-pdf tests".to_owned()),
            creation_date: Some("D:20260101120000+00'00'".to_owned()),
            modification_date: Some("D:20260101120000Z".to_owned()),
        },
    };
    export_pdf_with_options(
        &display_list,
        Some(&store),
        &BTreeMap::new(),
        None,
        &options,
    )
    .expect("export PDF/A-2u")
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

/// Inflate every FlateDecode stream. That includes page content, font programs
/// and CID-to-GID maps; callers select the stream they need.
fn inflated_streams(pdf: &[u8]) -> Vec<Vec<u8>> {
    let mut streams = Vec::new();
    let mut rest = pdf;
    while let Some(position) = find(rest, b"stream\n") {
        let start = position + b"stream\n".len();
        let Some(end) = find(&rest[start..], b"\nendstream") else {
            break;
        };
        let mut decoded = Vec::new();
        if ZlibDecoder::new(&rest[start..start + end])
            .read_to_end(&mut decoded)
            .is_ok()
        {
            streams.push(decoded);
        }
        rest = &rest[start + end + b"\nendstream".len()..];
    }
    streams
}

fn push_hex_codes(hex: &[u8], codes: &mut Vec<u16>) {
    let digits: Vec<u8> = hex
        .iter()
        .copied()
        .filter(|byte| !byte.is_ascii_whitespace())
        .collect();
    assert!(
        digits.len().is_multiple_of(4),
        "glyph codes must be 2-byte values"
    );
    for pair in digits.chunks(4) {
        let value =
            u16::from_str_radix(std::str::from_utf8(pair).expect("hex digits"), 16).expect("hex");
        codes.push(value);
    }
}

fn push_literal_codes(bytes: &[u8], codes: &mut Vec<u16>) {
    assert!(
        bytes.len().is_multiple_of(2),
        "Identity-H glyph codes must be 2 bytes"
    );
    for pair in bytes.chunks_exact(2) {
        codes.push(u16::from_be_bytes([pair[0], pair[1]]));
    }
}

/// Decode one PDF literal string starting at the opening parenthesis.
fn literal_string(content: &[u8], start: usize) -> (Vec<u8>, usize) {
    let mut bytes = Vec::new();
    let mut index = start + 1;
    let mut depth = 0usize;
    while index < content.len() {
        match content[index] {
            b'\\' => {
                index += 1;
                let Some(&escaped) = content.get(index) else {
                    break;
                };
                match escaped {
                    b'n' => {
                        bytes.push(b'\n');
                        index += 1;
                    }
                    b'r' => {
                        bytes.push(b'\r');
                        index += 1;
                    }
                    b't' => {
                        bytes.push(b'\t');
                        index += 1;
                    }
                    b'b' => {
                        bytes.push(0x08);
                        index += 1;
                    }
                    b'f' => {
                        bytes.push(0x0c);
                        index += 1;
                    }
                    b'(' | b')' | b'\\' => {
                        bytes.push(escaped);
                        index += 1;
                    }
                    b'\r' => {
                        index += 1;
                        if content.get(index) == Some(&b'\n') {
                            index += 1;
                        }
                    }
                    b'\n' => index += 1,
                    b'0'..=b'7' => {
                        let mut value = 0u16;
                        let mut digits = 0;
                        while digits < 3 {
                            let Some(&digit) = content.get(index) else {
                                break;
                            };
                            if !(b'0'..=b'7').contains(&digit) {
                                break;
                            }
                            value = value * 8 + u16::from(digit - b'0');
                            index += 1;
                            digits += 1;
                        }
                        bytes.push((value & 0xff) as u8);
                    }
                    other => {
                        bytes.push(other);
                        index += 1;
                    }
                }
            }
            b'(' => {
                depth += 1;
                bytes.push(b'(');
                index += 1;
            }
            b')' if depth == 0 => return (bytes, index + 1),
            b')' => {
                depth -= 1;
                bytes.push(b')');
                index += 1;
            }
            byte => {
                bytes.push(byte);
                index += 1;
            }
        }
    }
    (bytes, index)
}

/// Codes actually shown by `Tj`/`TJ` operators in a page content stream.
fn shown_codes(content: &[u8]) -> Vec<u16> {
    #[derive(Debug)]
    enum Token {
        /// `<...>` hex operand.
        Hex(Vec<u8>),
        /// `(...)` literal operand, already unescaped.
        Literal(Vec<u8>),
        ArrayStart,
        Word(String),
    }

    let mut tokens = Vec::new();
    let mut index = 0;
    while index < content.len() {
        match content[index] {
            b'%' => {
                while index < content.len() && content[index] != b'\n' {
                    index += 1;
                }
            }
            b'<' if content.get(index + 1) == Some(&b'<') => {
                // Inline dictionaries (marked-content properties) carry no
                // glyph operands.
                index += 2;
                while index + 1 < content.len()
                    && !(content[index] == b'>' && content[index + 1] == b'>')
                {
                    index += 1;
                }
                index = (index + 2).min(content.len());
            }
            b'<' => {
                let start = index + 1;
                index += 1;
                while index < content.len() && content[index] != b'>' {
                    index += 1;
                }
                tokens.push(Token::Hex(content[start..index].to_vec()));
                index += 1;
            }
            b'(' => {
                let (bytes, next) = literal_string(content, index);
                tokens.push(Token::Literal(bytes));
                index = next;
            }
            b'[' => {
                tokens.push(Token::ArrayStart);
                index += 1;
            }
            byte if byte.is_ascii_whitespace() => index += 1,
            _ => {
                let start = index;
                index += 1;
                while index < content.len()
                    && !content[index].is_ascii_whitespace()
                    && !b"[]<>()/%".contains(&content[index])
                {
                    index += 1;
                }
                tokens.push(Token::Word(
                    String::from_utf8_lossy(&content[start..index]).into_owned(),
                ));
            }
        }
    }

    let mut codes = Vec::new();
    for (position, token) in tokens.iter().enumerate() {
        let Token::Word(word) = token else {
            continue;
        };
        if word == "Tj" {
            match position.checked_sub(1).and_then(|i| tokens.get(i)) {
                Some(Token::Hex(hex)) => push_hex_codes(hex, &mut codes),
                Some(Token::Literal(bytes)) => push_literal_codes(bytes, &mut codes),
                _ => {}
            }
        } else if word == "TJ" {
            // Walk back through the array: string operands and numeric adjust
            // items until the opening bracket.
            let mut cursor = position;
            while let Some(previous) = cursor.checked_sub(1) {
                match &tokens[previous] {
                    Token::Hex(hex) => push_hex_codes(hex, &mut codes),
                    Token::Literal(bytes) => push_literal_codes(bytes, &mut codes),
                    Token::ArrayStart => break,
                    Token::Word(_) => {}
                }
                cursor = previous;
            }
        }
    }
    codes
}

/// Hex targets of every `<cid> <utf16>` pair in the emitted `bfchar` blocks.
fn to_unicode_map(text: &str) -> BTreeMap<u16, String> {
    fn entries(line: &str) -> Vec<&str> {
        let mut values = Vec::new();
        let mut rest = line;
        while let Some(start) = rest.find('<') {
            let after = &rest[start + 1..];
            let Some(end) = after.find('>') else {
                break;
            };
            values.push(&after[..end]);
            rest = &after[end + 1..];
        }
        values
    }

    let mut map = BTreeMap::new();
    for block in text.split("beginbfchar").skip(1) {
        let block = block.split("endbfchar").next().unwrap_or_default();
        for line in block.lines() {
            let values = entries(line);
            if values.len() < 2 {
                continue;
            }
            let Ok(cid) = u16::from_str_radix(values[0].trim(), 16) else {
                continue;
            };
            let units: Vec<u16> = values[1..]
                .iter()
                .filter_map(|value| u16::from_str_radix(value.trim(), 16).ok())
                .collect();
            let target = String::from_utf16_lossy(&units);
            map.insert(cid, target);
        }
    }
    map
}

/// One object dictionary containing `marker`, without the surrounding bytes.
fn object_with<'a>(pdf: &'a str, marker: &str) -> &'a str {
    let Some(position) = pdf.find(marker) else {
        panic!("PDF has no {marker}");
    };
    let start = pdf[..position]
        .rfind("\nendobj")
        .map_or(0, |end| end + "\nendobj".len());
    let end = pdf[position..]
        .find("\nendobj")
        .map_or(pdf.len(), |offset| position + offset);
    &pdf[start..end]
}

#[test]
fn pdfa2u_export_has_the_required_structure() {
    let pdf = sample();
    assert!(
        pdf.starts_with(b"%PDF-1.7"),
        "PDF/A-2u output must declare PDF 1.7: {:?}",
        &pdf[..pdf.len().min(16)]
    );
    assert!(
        pdf.windows(5).any(|window| window == b"%%EOF"),
        "missing %%EOF trailer"
    );
    let text = String::from_utf8_lossy(&pdf).into_owned();

    // Catalog metadata, archival identification and color intent.
    assert!(text.contains("/Metadata"), "catalog has no /Metadata");
    assert!(
        text.contains("<pdfaid:part>2</pdfaid:part>"),
        "XMP does not declare pdfaid:part 2"
    );
    assert!(
        text.contains("<pdfaid:conformance>U</pdfaid:conformance>"),
        "XMP does not declare pdfaid:conformance U"
    );
    assert!(
        text.contains("/OutputIntents"),
        "catalog has no /OutputIntents"
    );
    assert!(
        text.contains("/S /GTS_PDFA1"),
        "output intent is not GTS_PDFA1"
    );
    assert!(
        text.contains("/DestOutputProfile"),
        "output intent has no ICC profile"
    );

    // Trailer /Info and its supplied fields.
    assert!(text.contains("/Info "), "trailer has no /Info reference");
    assert!(
        text.contains("/Title (PDF/A-2u acceptance)"),
        "/Info is missing the title"
    );
    assert!(
        text.contains("/Producer (BetterOffice "),
        "/Info is missing the fixed producer"
    );

    // Every font resource is a real embedded program; no Base-14 fallback.
    assert!(
        !text.contains("/BaseFont /Helvetica") && !text.contains("/BaseFont /Times"),
        "PDF/A-2u output fell back to a Base-14 font"
    );
    let descriptors = text.matches("/Type /FontDescriptor").count();
    assert!(descriptors > 0, "output has no font descriptors");
    let programs = text.matches("/FontFile2").count() + text.matches("/FontFile3").count();
    assert_eq!(
        programs, descriptors,
        "every font descriptor must embed a font program"
    );
    assert!(text.contains("/Identity-H"), "CID encoding missing");
    assert!(text.contains("/ToUnicode"), "ToUnicode stream missing");

    // Link annotations must carry the Print flag (/F 4), never Hidden or
    // Invisible, and must retain their URI action.
    let annotation = object_with(&text, "/Subtype /Link");
    assert!(
        annotation.contains("/F 4"),
        "link annotation is missing the Print flag: {annotation}"
    );
    assert!(
        annotation.contains("/URI ("),
        "link annotation has no URI action: {annotation}"
    );

    // Text codes: every shown code is non-zero and resolves through the
    // ToUnicode CMap to at least one non-NUL character.
    let content = inflated_streams(&pdf)
        .into_iter()
        .find(|stream| {
            find(stream, b"BT").is_some()
                && find(stream, b" Tf").is_some()
                && (find(stream, b"Tj").is_some() || find(stream, b"TJ").is_some())
        })
        .expect("page content stream");
    let codes = shown_codes(&content);
    assert!(!codes.is_empty(), "page content shows no text codes");
    assert_eq!(
        codes.len(),
        TEXT.chars().count(),
        "one non-zero glyph per source character expected"
    );
    assert!(
        codes.iter().all(|code| *code != 0),
        "text code 0 (.notdef) was painted: {codes:?}"
    );
    let cmap = to_unicode_map(&text);
    assert!(!cmap.is_empty(), "ToUnicode CMap has no bfchar entries");
    let mut recovered = String::new();
    for code in &codes {
        let Some(target) = cmap.get(code) else {
            panic!("text code {code} has no ToUnicode mapping: {cmap:?}");
        };
        assert!(
            target.chars().any(|character| character != '\0'),
            "text code {code} maps to an empty/NUL target"
        );
        recovered.extend(target.chars());
    }
    for character in TEXT.chars() {
        assert!(
            recovered.contains(character),
            "ToUnicode mappings do not recover {character:?}: {recovered:?}"
        );
    }
}

#[test]
fn verapdf_accepts_the_pdfa2u_export_when_configured() {
    let Ok(verapdf) = std::env::var("VERAPDF") else {
        eprintln!("VERAPDF is not set; skipping external PDF/A-2u validation");
        return;
    };
    let pdf = sample();
    let directory = std::env::temp_dir().join(format!(
        "betteroffice-pdfa2u-{}-{}",
        std::process::id(),
        std::thread::current().name().unwrap_or("test")
    ));
    std::fs::create_dir_all(&directory).expect("create temp directory");
    let path = directory.join("sample.pdf");
    std::fs::write(&path, &pdf).expect("write sample PDF");

    let output = match Command::new(&verapdf)
        .args(["--format", "json", "--flavour", "2u"])
        .arg(&path)
        .output()
    {
        Ok(output) => output,
        Err(error) => panic!("run {verapdf}: {error}"),
    };
    let _ = std::fs::remove_file(&path);
    let _ = std::fs::remove_dir(&directory);

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    let report: Value = match serde_json::from_slice(&output.stdout) {
        Ok(report) => report,
        Err(error) => {
            panic!("veraPDF did not emit JSON ({error}); stdout={stdout}; stderr={stderr}")
        }
    };
    let result = &report["report"]["jobs"][0]["validationResult"][0];
    let Some(compliant) = result["compliant"].as_bool() else {
        panic!("veraPDF report has no compliant flag; stdout={stdout}; stderr={stderr}")
    };
    assert!(
        compliant,
        "veraPDF rejected the PDF/A-2u export: {}",
        result["details"]
    );
}
