//! Focused tests for the PDF/A `/Info` and XMP metadata writers.
//!
//! The writers are library-private, so this test compiles the same source as a
//! local module (the aliases below satisfy the module's `crate::…` paths) and
//! drives it with a bare `Pdf`, where the reserved ids are the only ones used.

use pdf_writer::{Pdf, Ref};

pub use betteroffice_docx_pdf::{PdfConformance, PdfError, PdfMetadata};

#[path = "../src/pdfa_limits.rs"]
mod pdfa_limits;

#[path = "../src/metadata.rs"]
mod metadata;

/// Metadata with every field populated.
fn full_metadata() -> PdfMetadata {
    PdfMetadata {
        title: Some("Quarterly report".to_owned()),
        author: Some("Ada Lovelace".to_owned()),
        subject: Some("Q1 results".to_owned()),
        keywords: Some("finance, quarterly".to_owned()),
        creator: Some("BetterOffice Writer".to_owned()),
        creation_date: Some("D:20240102030405+02'00'".to_owned()),
        modification_date: Some("D:20231231235959Z".to_owned()),
    }
}

/// Render only the `/Info` object and return the PDF as lossy text.
fn info_text(metadata: &PdfMetadata) -> String {
    let mut pdf = Pdf::new();
    metadata::write_document_info(&mut pdf, Ref::new(1), metadata).expect("write document info");
    String::from_utf8_lossy(&pdf.finish()).into_owned()
}

/// Extract the XMP packet from a finished PDF.
fn xmp_xml(pdf: &[u8]) -> String {
    let text = String::from_utf8_lossy(pdf);
    let start = text.find("<?xml").expect("XMP declaration");
    let end = text.find("</x:xmpmeta>").expect("XMP root end") + "</x:xmpmeta>".len();
    text[start..end].to_owned()
}

/// Render only the XMP stream and return the packet.
fn xmp_text(
    conformance: PdfConformance,
    metadata: &PdfMetadata,
) -> Result<(Vec<u8>, String), PdfError> {
    let mut pdf = Pdf::new();
    metadata::write_xmp_metadata(&mut pdf, Ref::new(1), conformance, metadata)?;
    let bytes = pdf.finish();
    let xml = xmp_xml(&bytes);
    Ok((bytes, xml))
}

#[test]
fn info_writes_every_present_entry_with_the_fixed_producer() {
    let text = info_text(&full_metadata());

    for (key, value) in [
        ("Title", "Quarterly report"),
        ("Author", "Ada Lovelace"),
        ("Subject", "Q1 results"),
        ("Keywords", "finance, quarterly"),
        ("Creator", "BetterOffice Writer"),
    ] {
        assert!(
            text.contains(&format!("/{key} ({value})")),
            "/Info is missing /{key}: {text}"
        );
    }
    assert!(
        text.contains(&format!(
            "/Producer (BetterOffice {})",
            env!("CARGO_PKG_VERSION")
        )),
        "/Info is missing the fixed producer: {text}"
    );
    assert!(
        text.contains("/CreationDate (D:20240102030405+02'00')"),
        "creation date must be written verbatim: {text}"
    );
    assert!(
        text.contains("/ModDate (D:20231231235959Z)"),
        "modification date must be written verbatim: {text}"
    );
}

#[test]
fn info_keeps_present_entries_in_a_fixed_order() {
    let text = info_text(&full_metadata());
    let positions: Vec<usize> = [
        "/Title",
        "/Author",
        "/Subject",
        "/Keywords",
        "/Creator",
        "/Producer",
        "/CreationDate",
        "/ModDate",
    ]
    .iter()
    .map(|key| text.find(*key).unwrap_or_else(|| panic!("missing {key}")))
    .collect();
    assert!(
        positions.windows(2).all(|pair| pair[0] < pair[1]),
        "/Info entries are not in a deterministic order"
    );
}

#[test]
fn info_omits_absent_entries_but_always_writes_the_producer() {
    let text = info_text(&PdfMetadata::default());

    assert!(text.contains("/Producer ("));
    for key in [
        "/Title",
        "/Author",
        "/Subject",
        "/Keywords",
        "/Creator",
        "/CreationDate",
        "/ModDate",
    ] {
        assert!(!text.contains(key), "absent {key} must be omitted: {text}");
    }
}

#[test]
fn info_accepts_every_pdf_date_precision() {
    for date in [
        "D:2024",
        "D:202401",
        "D:20240102",
        "D:2024010203",
        "D:202401020304",
        "D:20240102030405",
        "D:20240102030405Z",
        "D:20240102030405+02'30'",
        "D:20240102030405-08'00'",
        // A typed `pdf_writer::Date` would rewrite this as `Z`; the writer
        // keeps the caller's bytes.
        "D:20240102030405+00'00'",
        "D:20240229",
        "D:20241231235959Z",
    ] {
        let metadata = PdfMetadata {
            creation_date: Some(date.to_owned()),
            ..PdfMetadata::default()
        };
        let text = info_text(&metadata);
        assert!(
            text.contains(&format!("/CreationDate ({date})")),
            "valid date {date} was not preserved verbatim: {text}"
        );
    }
}

#[test]
fn info_rejects_malformed_dates() {
    for date in [
        "2024-01-02",
        "d:20240102",
        "D:",
        "D:202",
        "D:20241",
        "D:2024-01-02",
        "D:20241301",
        "D:20240001",
        "D:20230229",
        "D:20240100",
        "D:20240132",
        "D:2024010224",
        "D:202401020360",
        "D:20240102030460",
        "D:20240102030405z",
        "D:20240102030405+0200",
        "D:20240102030405+02'00",
        "D:20240102030405+02'0'",
        "D:20240102030405+24'00'",
        "D:20240102030405+02'60'",
        "D:20240102030405 Z",
        "D:202401020304+02'00'",
    ] {
        let metadata = PdfMetadata {
            creation_date: Some(date.to_owned()),
            ..PdfMetadata::default()
        };
        let mut pdf = Pdf::new();
        let error = metadata::write_document_info(&mut pdf, Ref::new(1), &metadata)
            .expect_err(&format!("malformed date {date} was accepted"));
        assert!(
            matches!(error, PdfError::Unsupported(_) | PdfError::InvalidNumber(_)),
            "unexpected error for {date}: {error}"
        );
    }
}

#[test]
fn xmp_adds_minutes_to_an_hour_only_date() {
    let metadata = PdfMetadata {
        creation_date: Some("D:2024010203".to_owned()),
        ..PdfMetadata::default()
    };
    let (_bytes, xml) = xmp_text(PdfConformance::PdfA2u, &metadata).expect("write XMP");
    assert!(
        xml.contains("2024-01-02T03:00"),
        "hour-only PDF date must become XMP with minutes: {xml}"
    );
    assert!(
        !xml.contains("2024-01-02T03<"),
        "XMP date must not omit minutes: {xml}"
    );
}

#[test]
fn xmp_has_the_standard_namespaces_and_pdfa_identification() {
    let (bytes, xml) = xmp_text(PdfConformance::PdfA2u, &full_metadata()).expect("write XMP");
    let text = String::from_utf8_lossy(&bytes);

    assert!(text.contains("/Type /Metadata"), "missing metadata type");
    assert!(text.contains("/Subtype /XML"), "missing XML subtype");
    assert!(
        xml.starts_with("<?xml version=\"1.0\"?>"),
        "XMP declaration must omit the encoding attribute PDF/A-2 forbids: {xml}"
    );
    assert!(
        !xml.contains("encoding="),
        "XMP declaration must not carry an encoding attribute: {xml}"
    );
    for declaration in [
        "xmlns:x=\"adobe:ns:meta/\"",
        "xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"",
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\"",
        "xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\"",
        "xmlns:pdf=\"http://ns.adobe.com/pdf/1.3/\"",
        "xmlns:pdfaid=\"http://www.aiim.org/pdfa/ns/id/\"",
    ] {
        assert!(xml.contains(declaration), "missing {declaration}: {xml}");
    }
    assert!(xml.contains("<pdfaid:part>2</pdfaid:part>"));
    assert!(xml.contains("<pdfaid:conformance>U</pdfaid:conformance>"));
    assert_eq!(xml.matches("<pdfaid:part>").count(), 1);
    assert_eq!(xml.matches("<pdfaid:conformance>").count(), 1);
}

#[test]
fn xmp_mirrors_the_info_entries() {
    let (_, xml) = xmp_text(PdfConformance::PdfA2u, &full_metadata()).expect("write XMP");

    assert!(
        xml.contains(
            "<dc:title><rdf:Alt><rdf:li xml:lang=\"x-default\">Quarterly report</rdf:li></rdf:Alt></dc:title>"
        ),
        "missing dc:title: {xml}"
    );
    assert!(
        xml.contains("<dc:creator><rdf:Seq><rdf:li>Ada Lovelace</rdf:li></rdf:Seq></dc:creator>"),
        "missing dc:creator sequence: {xml}"
    );
    assert!(
        xml.contains(
            "<dc:description><rdf:Alt><rdf:li xml:lang=\"x-default\">Q1 results</rdf:li></rdf:Alt></dc:description>"
        ),
        "missing dc:description: {xml}"
    );
    assert!(xml.contains("<pdf:Keywords>finance, quarterly</pdf:Keywords>"));
    assert!(xml.contains("<xmp:CreatorTool>BetterOffice Writer</xmp:CreatorTool>"));
    assert!(xml.contains(&format!(
        "<pdf:Producer>BetterOffice {}</pdf:Producer>",
        env!("CARGO_PKG_VERSION")
    )));
    assert!(
        xml.contains("<xmp:CreateDate>2024-01-02T03:04:05+02:00</xmp:CreateDate>"),
        "PDF date was not converted to an XMP date: {xml}"
    );
    assert!(
        xml.contains("<xmp:ModifyDate>2023-12-31T23:59:59Z</xmp:ModifyDate>"),
        "PDF date was not converted to an XMP date: {xml}"
    );
}

#[test]
fn xmp_preserves_carriage_return_as_a_character_reference() {
    let metadata = PdfMetadata {
        title: Some("a\rb".to_owned()),
        ..PdfMetadata::default()
    };
    let (_, xml) = xmp_text(PdfConformance::PdfA2u, &metadata).expect("write XMP");
    assert!(
        xml.contains(">a&#13;b</rdf:li>"),
        "carriage return must survive XML line-ending normalization: {xml}"
    );
    let info = info_text(&metadata);
    assert!(
        info.contains("/Title <FEFF0061000D0062>"),
        "/Info retains the raw carriage return code unit: {info}"
    );
}

#[test]
fn xmp_omits_absent_fields() {
    let (_, xml) = xmp_text(PdfConformance::PdfA2u, &PdfMetadata::default()).expect("write XMP");

    for element in [
        "<dc:title>",
        "<dc:creator>",
        "<dc:description>",
        "<pdf:Keywords>",
        "<xmp:CreatorTool>",
        "<xmp:CreateDate>",
        "<xmp:ModifyDate>",
    ] {
        assert!(!xml.contains(element), "absent {element} must be omitted");
    }
    assert!(xml.contains("<pdf:Producer>"));
    assert!(xml.contains("<pdfaid:part>2</pdfaid:part>"));
    assert!(xml.contains("<pdfaid:conformance>U</pdfaid:conformance>"));
}

#[test]
fn xmp_escapes_xml_metacharacters_and_supports_unicode() {
    let title = "Résumé & <tag> \"quoted\" 'single' 日本語";
    let metadata = PdfMetadata {
        title: Some(title.to_owned()),
        ..PdfMetadata::default()
    };
    let (_, xml) = xmp_text(PdfConformance::PdfA2u, &metadata).expect("write XMP");

    assert!(
        xml.contains(
            "<dc:title><rdf:Alt><rdf:li xml:lang=\"x-default\">Résumé &amp; &lt;tag&gt; &quot;quoted&quot; &apos;single&apos; 日本語</rdf:li></rdf:Alt></dc:title>"
        ),
        "title was not escaped: {xml}"
    );

    let info = info_text(&metadata);
    let utf16: String = title
        .encode_utf16()
        .map(|unit| format!("{unit:04X}"))
        .collect();
    assert!(
        info.contains(&format!("/Title <FEFF{utf16}>")),
        "/Info must encode non-ASCII text as UTF-16BE: {info}"
    );
}

#[test]
fn xmp_rejects_malformed_dates() {
    let metadata = PdfMetadata {
        modification_date: Some("D:20230229".to_owned()),
        ..PdfMetadata::default()
    };
    let mut pdf = Pdf::new();
    let error =
        metadata::write_xmp_metadata(&mut pdf, Ref::new(1), PdfConformance::PdfA2u, &metadata)
            .expect_err("malformed date was accepted");
    assert!(matches!(
        error,
        PdfError::Unsupported(_) | PdfError::InvalidNumber(_)
    ));
}

#[test]
fn xmp_is_refused_for_standard_conformance() {
    let mut pdf = Pdf::new();
    let error = metadata::write_xmp_metadata(
        &mut pdf,
        Ref::new(1),
        PdfConformance::Standard,
        &PdfMetadata::default(),
    )
    .expect_err("Standard conformance must not claim PDF/A");
    assert!(matches!(error, PdfError::Unsupported(_)));
}

#[test]
fn identical_metadata_produces_identical_bytes() {
    let metadata = full_metadata();
    let render = || {
        let mut pdf = Pdf::new();
        metadata::write_document_info(&mut pdf, Ref::new(1), &metadata)
            .expect("write document info");
        metadata::write_xmp_metadata(&mut pdf, Ref::new(2), PdfConformance::PdfA2u, &metadata)
            .expect("write XMP");
        pdf.finish()
    };

    assert_eq!(render(), render(), "metadata output must be deterministic");
}

#[test]
fn info_rejects_a_pdfa_text_string_over_the_limit() {
    for title in ["A".repeat(32_768), "é".repeat(16_383)] {
        let metadata = PdfMetadata {
            title: Some(title),
            ..PdfMetadata::default()
        };
        let mut pdf = Pdf::new();
        let error = metadata::write_document_info(&mut pdf, Ref::new(1), &metadata)
            .expect_err("oversized PDF/A text string must be rejected");
        assert!(
            error.to_string().contains("32767"),
            "unexpected oversized-string error: {error}"
        );
    }
}

#[test]
fn info_accepts_the_largest_fitting_unicode_string() {
    let metadata = PdfMetadata {
        title: Some("é".repeat(16_382)),
        ..PdfMetadata::default()
    };
    let mut pdf = Pdf::new();
    metadata::write_document_info(&mut pdf, Ref::new(1), &metadata)
        .expect("32766-byte UTF-16 text string fits the PDF/A limit");
}
