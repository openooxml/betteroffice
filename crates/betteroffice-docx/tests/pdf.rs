#![cfg(all(feature = "raster", feature = "fonts"))]

use betteroffice_docx::Document;

const CARLITO_BOLD: &[u8] = include_bytes!("../../../packages/fonts/assets/Carlito-Bold.ttf");

#[test]
fn exports_a_docx_as_a_paginated_pdf_with_embedded_fonts() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/corpus/fixtures/betteroffice-demo.docx");
    let bytes = std::fs::read(path).expect("demo docx");
    let document = Document::open(&bytes).expect("open");
    let pdf = document.export_pdf().expect("export pdf");
    let text = String::from_utf8_lossy(&pdf);

    assert!(text.starts_with("%PDF-1."));
    assert!(
        text.contains("/Count 2"),
        "demo document paginates to two pages"
    );
    assert!(text.contains("/MediaBox [0 0 612 792]"));
    assert!(text.contains("/FontFile2"), "outlines must be embedded");
    assert!(text.contains("/Identity-H"));
    assert!(text.contains("/ToUnicode"));
    assert!(
        !text.contains("/BaseFont /Helvetica"),
        "glyph runs need no Base-14 fallback"
    );
}

/// An application face registered under an alias family must keep its place
/// ahead of the bundled face the facade also resolved for that chain. The
/// all-caps run takes the fallback text path, where PDF painting used to
/// reconstruct family preference from declared names and lose the alias.
#[test]
fn application_alias_stays_ahead_of_the_bundled_face_in_pdf() {
    let aliased_pdf = export(
        &[
            run("Roboto", true, true, "BOLD"),
            run("Roboto", false, false, "Normal"),
        ],
        Some("Roboto"),
    );
    let aliased_streams = font_file_streams(&aliased_pdf);
    assert_eq!(aliased_streams.len(), 2, "one subset per selected face");

    // The bold all-caps run must embed the alias face, not bundled Roboto.
    let carlito_pdf = export(&[run("Carlito", true, true, "BOLD")], Some("Carlito"));
    let carlito_streams = font_file_streams(&carlito_pdf);
    assert_eq!(carlito_streams.len(), 1);
    assert!(
        aliased_streams
            .iter()
            .any(|stream| *stream == carlito_streams[0]),
        "the aliased bold run did not embed the application Carlito-Bold face"
    );

    // The normal run still resolves the bundled Roboto face.
    let normal_pdf = export(&[run("Roboto", false, false, "Normal")], None);
    let normal_streams = font_file_streams(&normal_pdf);
    assert_eq!(normal_streams.len(), 1);
    assert!(
        aliased_streams
            .iter()
            .any(|stream| *stream == normal_streams[0]),
        "the normal Roboto run did not keep the bundled Roboto subset"
    );

    // And it must not pick the bundled bold face the alias shadows.
    let roboto_bold_pdf = export(&[run("Roboto", true, true, "BOLD")], None);
    let roboto_bold_streams = font_file_streams(&roboto_bold_pdf);
    assert_eq!(roboto_bold_streams.len(), 1);
    assert!(
        !aliased_streams
            .iter()
            .any(|stream| *stream == roboto_bold_streams[0]),
        "the aliased bold run fell back to bundled Roboto-Bold"
    );
}

/// Exports `runs`, registering `CARLITO_BOLD` under `alias` bold first when
/// given, so the application face precedes the bundled face in the chain.
fn export(runs: &[String], alias: Option<&str>) -> Vec<u8> {
    let mut document = Document::open(&minimal_docx(runs)).expect("open");
    if let Some(family) = alias {
        document
            .register_font(family, true, false, CARLITO_BOLD)
            .expect("register alias");
    }
    document.export_pdf().expect("export pdf")
}

fn run(family: &str, bold: bool, all_caps: bool, text: &str) -> String {
    format!(
        "<w:r><w:rPr><w:rFonts w:ascii=\"{family}\" w:hAnsi=\"{family}\"/>{}{}</w:rPr><w:t>{text}</w:t></w:r>",
        if bold { "<w:b/>" } else { "" },
        if all_caps { "<w:caps/>" } else { "" },
    )
}

fn minimal_docx(runs: &[String]) -> Vec<u8> {
    let document = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body><w:p>{}</w:p></w:body></w:document>",
        runs.concat()
    );
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).expect("rezip docx")
}

/// Embedded `FontFile2` stream payloads in PDF object order. Subsetting is
/// deterministic, so equal payloads mean the same face supplied the glyphs.
fn font_file_streams(pdf: &[u8]) -> Vec<&[u8]> {
    let mut streams = Vec::new();
    let mut cursor = 0;
    while let Some(offset) = find(&pdf[cursor..], b"/FontFile2") {
        let after_ref = cursor + offset + b"/FontFile2".len();
        let stream = after_ref + find(&pdf[after_ref..], b"stream").expect("stream keyword");
        let mut start = stream + b"stream".len();
        if pdf.get(start) == Some(&b'\r') {
            start += 1;
        }
        if pdf.get(start) == Some(&b'\n') {
            start += 1;
        }
        let length = find(&pdf[start..], b"endstream").expect("endstream");
        streams.push(&pdf[start..start + length]);
        cursor = start + length;
    }
    streams
}

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}
