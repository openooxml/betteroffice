use std::collections::BTreeMap;
use std::io::Cursor;

use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use image::{DynamicImage, ImageBuffer, ImageFormat, Rgb};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::*;

const DOCX_SECRETS: &[&str] = &[
    "DOCX_SECRET_TEXT",
    "DOCX_SECRET_COMMENT",
    "DOCX_SECRET_AUTHOR",
    "DOCX_SECRET_TITLE",
    "DOCX_SECRET_COMPANY",
    "https://secret.example/docx",
];
const XLSX_SECRETS: &[&str] = &[
    "XLSX_SECRET_TEXT",
    "XLSX_INLINE_SECRET",
    "XLSX_SECRET_SHEET",
    "XLSX_SECRET_AUTHOR",
    "XLSX_SECRET_COMPANY",
    "https://secret.example/xlsx",
];
const PPTX_SECRETS: &[&str] = &[
    "PPTX_SECRET_TEXT",
    "PPTX_SECRET_NOTES",
    "PPTX_SECRET_AUTHOR",
    "PPTX_SECRET_COMPANY",
    "https://secret.example/pptx",
];

#[test]
fn redacts_docx_without_changing_structure() {
    let source = docx_fixture();
    let (output, report) = redact_with_report(&source, Format::Auto).unwrap();
    assert_eq!(report.format, Format::Docx);
    assert_fixture_properties(&source, &output, DOCX_SECRETS, "word/media/image1.png");
    assert_text_lengths(&source, &output, "word/document.xml", "t");
    parse_docx_s9_wire(&output, S9ParseOptions::default()).unwrap();
}

#[test]
fn redacts_xlsx_without_changing_structure() {
    let source = xlsx_fixture();
    let (output, report) = redact_with_report(&source, Format::Xlsx).unwrap();
    assert_eq!(report.format, Format::Xlsx);
    assert_fixture_properties(&source, &output, XLSX_SECRETS, "xl/media/image1.png");
    assert_text_lengths(&source, &output, "xl/sharedStrings.xml", "t");
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    xlsx_parse::parse_workbook(&parts).unwrap();
}

#[test]
fn redacts_pptx_without_changing_structure() {
    let source = pptx_fixture();
    let (output, report) = redact_with_report(&source, Format::Pptx).unwrap();
    assert_eq!(report.format, Format::Pptx);
    assert_fixture_properties(&source, &output, PPTX_SECRETS, "ppt/media/image1.png");
    assert_text_lengths(&source, &output, "ppt/slides/slide1.xml", "t");
    pptx_parse::parse_pptx(&output).unwrap();
}

#[test]
fn jpeg_placeholder_is_fixed_size() {
    let source = placeholder_image(ImageFormat::Jpeg);
    let mut report = RedactionReport::default();
    let output = media::replace_media("word/media/photo.jpeg", &source, &mut report).unwrap();
    assert_ne!(source, output);
    assert_eq!(
        image_dimensions(&output),
        (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE)
    );
    assert_eq!(image::guess_format(&output).unwrap(), ImageFormat::Jpeg);
}

#[test]
fn rejects_explicit_format_mismatch() {
    let error = redact(&docx_fixture(), Format::Xlsx).unwrap_err();
    assert!(matches!(error, RedactError::FormatMismatch { .. }));
}

fn assert_fixture_properties(source: &[u8], output: &[u8], secrets: &[&str], media_path: &str) {
    let before = ooxml_opc::unzip_parts(source).unwrap();
    let after = ooxml_opc::unzip_parts(output).unwrap();
    assert_eq!(part_names(&before), part_names(&after));
    assert_eq!(element_counts(&before), element_counts(&after));

    for secret in secrets {
        assert!(
            after
                .iter()
                .all(|(_, bytes)| !String::from_utf8_lossy(bytes).contains(secret)),
            "secret survived: {secret}"
        );
    }

    let before_image = part(&before, media_path);
    let after_image = part(&after, media_path);
    assert_ne!(before_image, after_image);
    assert_eq!(
        image_dimensions(after_image),
        (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE)
    );
    assert_eq!(image::guess_format(after_image).unwrap(), ImageFormat::Png);
}

fn part_names(parts: &[(String, Vec<u8>)]) -> Vec<&str> {
    parts.iter().map(|(path, _)| path.as_str()).collect()
}

fn element_counts(parts: &[(String, Vec<u8>)]) -> BTreeMap<&str, usize> {
    parts
        .iter()
        .filter(|(path, _)| is_xml_part(&path.to_ascii_lowercase()))
        .map(|(path, bytes)| (path.as_str(), element_count(bytes)))
        .collect()
}

fn element_count(bytes: &[u8]) -> usize {
    let mut reader = Reader::from_reader(bytes);
    let mut count = 0;
    loop {
        match reader.read_event().unwrap() {
            Event::Start(_) | Event::Empty(_) => count += 1,
            Event::Eof => return count,
            _ => {}
        }
    }
}

fn assert_text_lengths(source: &[u8], output: &[u8], path: &str, element: &str) {
    let before = ooxml_opc::unzip_parts(source).unwrap();
    let after = ooxml_opc::unzip_parts(output).unwrap();
    assert_eq!(
        text_lengths(part(&before, path), element),
        text_lengths(part(&after, path), element)
    );
}

fn text_lengths(bytes: &[u8], target: &str) -> Vec<usize> {
    let mut reader = Reader::from_reader(bytes);
    let mut inside = false;
    let mut lengths = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) if start.name().local_name().as_ref() == target.as_bytes() => {
                inside = true;
            }
            Event::Text(text) if inside => lengths.push(text.decode().unwrap().chars().count()),
            Event::End(end) if end.name().local_name().as_ref() == target.as_bytes() => {
                inside = false;
            }
            Event::Eof => return lengths,
            _ => {}
        }
    }
}

fn part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> &'a [u8] {
    parts
        .iter()
        .find(|(candidate, _)| candidate == path)
        .map(|(_, bytes)| bytes.as_slice())
        .unwrap()
}

fn image_dimensions(bytes: &[u8]) -> (u32, u32) {
    image::ImageReader::new(Cursor::new(bytes))
        .with_guessed_format()
        .unwrap()
        .into_dimensions()
        .unwrap()
}

fn placeholder_png() -> Vec<u8> {
    placeholder_image(ImageFormat::Png)
}

fn placeholder_image(format: ImageFormat) -> Vec<u8> {
    let image = DynamicImage::ImageRgb8(ImageBuffer::from_fn(3, 2, |x, y| {
        Rgb([(x * 80) as u8, (y * 100) as u8, 40])
    }));
    let mut output = Cursor::new(Vec::new());
    image.write_to(&mut output, format).unwrap();
    output.into_inner()
}

fn package(mut parts: Vec<(&str, Vec<u8>)>) -> Vec<u8> {
    let owned: Vec<_> = parts
        .drain(..)
        .map(|(path, bytes)| (path.to_owned(), bytes))
        .collect();
    ooxml_opc::rezip_parts(&owned).unwrap()
}

fn xml(value: &str) -> Vec<u8> {
    value.as_bytes().to_vec()
}

fn docx_fixture() -> Vec<u8> {
    package(vec![
        (
            "[Content_Types].xml",
            xml(
                r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>"#,
            ),
        ),
        (
            "_rels/.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
        ),
        (
            "docProps/core.xml",
            xml(
                r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>DOCX_SECRET_TITLE</dc:title><dc:creator>DOCX_SECRET_AUTHOR</dc:creator></cp:coreProperties>"#,
            ),
        ),
        (
            "docProps/app.xml",
            xml(
                r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>DOCX_SECRET_COMPANY</Company><Pages>1</Pages></Properties>"#,
            ),
        ),
        (
            "word/document.xml",
            xml(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>DOCX_SECRET_TEXT</w:t></w:r><w:ins w:id="1" w:author="DOCX_SECRET_AUTHOR"><w:r><w:t>tracked secret</w:t></w:r></w:ins><w:hyperlink r:id="rId9"><w:r><w:t>private link</w:t></w:r></w:hyperlink></w:p><w:sectPr/></w:body></w:document>"#,
            ),
        ),
        (
            "word/comments.xml",
            xml(
                r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="DOCX_SECRET_AUTHOR"><w:p><w:r><w:t>DOCX_SECRET_COMMENT</w:t></w:r></w:p></w:comment></w:comments>"#,
            ),
        ),
        (
            "word/_rels/document.xml.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example/docx" TargetMode="External"/><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>"#,
            ),
        ),
        ("word/media/image1.png", placeholder_png()),
    ])
}

fn xlsx_fixture() -> Vec<u8> {
    package(vec![
        (
            "[Content_Types].xml",
            xml(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"#,
            ),
        ),
        (
            "_rels/.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
        ),
        (
            "xl/workbook.xml",
            xml(
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="XLSX_SECRET_SHEET" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
        ),
        (
            "xl/sharedStrings.xml",
            xml(
                r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>XLSX_SECRET_TEXT</t></si></sst>"#,
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            xml(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>XLSX_INLINE_SECRET</t></is></c><c r="C1"><f>SUM(1,2)</f><v>3</v></c></row></sheetData></worksheet>"#,
            ),
        ),
        (
            "xl/comments1.xml",
            xml(
                r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>XLSX_SECRET_AUTHOR</author></authors><commentList/></comments>"#,
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example/xlsx" TargetMode="External"/></Relationships>"#,
            ),
        ),
        (
            "docProps/app.xml",
            xml(
                r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>XLSX_SECRET_COMPANY</Company></Properties>"#,
            ),
        ),
        ("xl/media/image1.png", placeholder_png()),
    ])
}

fn pptx_fixture() -> Vec<u8> {
    package(vec![
        (
            "[Content_Types].xml",
            xml(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
            ),
        ),
        (
            "_rels/.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            ),
        ),
        (
            "ppt/presentation.xml",
            xml(
                r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#,
            ),
        ),
        (
            "ppt/_rels/presentation.xml.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
            ),
        ),
        (
            "ppt/slides/slide1.xml",
            xml(
                r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld name="Private slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="PPTX secret box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>PPTX_SECRET_TEXT</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
            ),
        ),
        (
            "ppt/notesSlides/notesSlide1.xml",
            xml(
                r#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>PPTX_SECRET_NOTES</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#,
            ),
        ),
        (
            "ppt/commentAuthors.xml",
            xml(
                r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="PPTX_SECRET_AUTHOR" initials="PSA"/></p:cmAuthorLst>"#,
            ),
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example/pptx" TargetMode="External"/></Relationships>"#,
            ),
        ),
        (
            "docProps/app.xml",
            xml(
                r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>PPTX_SECRET_COMPANY</Company></Properties>"#,
            ),
        ),
        ("ppt/media/image1.png", placeholder_png()),
    ])
}

#[test]
fn empty_shared_string_cell_does_not_leak_next_value() {
    // Greptile #68: a self-closing <c t="s"/> has no End event, so its cell
    // type must not bleed into the following untyped numeric cell's value.
    let sheet = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<sheetData><row r="1">"#,
        r#"<c r="A1" t="s"/>"#,
        r#"<c r="B1"><v>424242</v></c>"#,
        r#"</row></sheetData></worksheet>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Xlsx,
        "xl/worksheets/sheet1.xml",
        sheet.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(
        !text.contains("424242"),
        "numeric cell value leaked: {text}"
    );
}

#[test]
fn scheme_bearing_relationship_targets_redacted_without_target_mode() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example/docx" TargetMode="External"/>"#,
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a"/>"#,
        r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///C:/Users/jane/x.xlsx"/>"#,
        r#"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///\\server\share\x.xlsx"/>"#,
        r#"<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"Target="https://example.com""#));
    assert!(!text.contains("secret.example"));
    assert!(!text.contains(r"\server\share"));
    assert!(!text.contains("/Users/jane"));
    assert!(text.contains(r#"Target="media/image1.png""#));
    assert_eq!(report.attributes, 4);
}

#[test]
fn rfc3986_scheme_targets_redacted_without_target_mode() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="tel:+49123456789"/>"#,
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="MAILTO:jane@example.com"/>"#,
        r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="myapp://x"/>"#,
        r#"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../fonts/x.ttf"/>"#,
        r#"<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="media/image:1.png"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert_eq!(text.matches(r#"Target="https://example.com""#).count(), 3);
    assert!(!text.contains("tel:"));
    assert!(!text.contains("jane@example.com"));
    assert!(!text.contains("myapp:"));
    assert!(text.contains(r#"Target="../fonts/x.ttf""#));
    assert!(text.contains(r#"Target="media/image:1.png""#));
    assert_eq!(report.attributes, 3);
}

#[test]
fn uri_in_fragment_keeps_relationship_internal() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="worksheet.xml#ref=https://example.com/x"/>"#,
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/x"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"Target="worksheet.xml#ref=https://example.com/x""#));
    assert!(!text.contains(r#"Target="https://example.com/x""#));
    assert!(text.contains(r#"Target="https://example.com""#));
    assert_eq!(report.attributes, 1);
}

#[test]
fn unc_and_protocol_relative_targets_redacted_without_target_mode() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="\\fileserver01\finance\q3.xlsx"/>"#,
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="\\?\UNC\fileserver02\finance\q4.xlsx"/>"#,
        r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="//fileserver03.example/finance/q1.xlsx"/>"#,
        r#"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="\word\media\image1.png"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert_eq!(text.matches(r#"Target="https://example.com""#).count(), 3);
    assert!(!text.contains("fileserver01"));
    assert!(!text.contains("fileserver02"));
    assert!(!text.contains("fileserver03"));
    assert!(text.contains(r#"Target="\word\media\image1.png""#));
    assert_eq!(report.attributes, 3);
}

#[test]
fn percent_encoded_and_dotted_relative_targets_stay_internal() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../word/media/im%7Eage1.png"/>"#,
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../word/./media/image2.png"/>"#,
        r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="..\..\word\media\image3.png"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"Target="../../word/media/im%7Eage1.png""#));
    assert!(text.contains(r#"Target="../../word/./media/image2.png""#));
    assert!(text.contains(r#"Target="..\..\word\media\image3.png""#));
    assert!(!text.contains("TargetMode"));
    assert_eq!(report.attributes, 0);
}

#[test]
fn padded_target_mode_still_marks_relationship_external() {
    let source = pptx_fixture_with_slide_relationship(
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="PPTX_SECRET_SHARE/finance/q1.xlsx" TargetMode=" External "/>"#,
    );
    let (output, _) = redact_with_report(&source, Format::Pptx).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let rels =
        String::from_utf8_lossy(part(&parts, "ppt/slides/_rels/slide1.xml.rels")).into_owned();
    assert!(rels.contains(r#"Target="https://example.com""#));
    assert!(rels.contains(r#"TargetMode="External""#));
    assert!(!rels.contains("PPTX_SECRET_SHARE"));
    pptx_parse::parse_pptx(&output).unwrap();
}

#[test]
fn lowercase_target_mode_attribute_is_written_back_canonically() {
    let source = pptx_fixture_with_slide_relationship(
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://PPTX_SECRET_HOST/x" targetmode="external"/>"#,
    );
    let (output, _) = redact_with_report(&source, Format::Pptx).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let rels =
        String::from_utf8_lossy(part(&parts, "ppt/slides/_rels/slide1.xml.rels")).into_owned();
    assert!(rels.contains(r#"TargetMode="External""#));
    assert!(!rels.contains("targetmode"));
    assert!(!rels.contains("PPTX_SECRET_HOST"));
    pptx_parse::parse_pptx(&output).unwrap();
}

#[test]
fn a_case_variant_target_does_not_externalize_the_real_one() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml" target="mailto:jane@example.com"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Pptx,
        "ppt/_rels/presentation.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"Target="slides/slide1.xml""#));
    assert!(!text.contains("TargetMode"));
    assert_eq!(report.attributes, 0);
}

#[test]
fn repeated_target_mode_spellings_collapse_to_one_attribute() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example/x" TargetMode="External" targetmode="external"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert_eq!(text.matches("TargetMode").count(), 1);
    assert!(!text.contains("secret.example"));
    let mut again = RedactionReport::default();
    xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        text.as_bytes(),
        &mut again,
    )
    .unwrap();
}

#[test]
fn internal_target_mode_keeps_the_producer_spelling() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png" targetmode="internal"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"targetmode="internal""#));
    assert_eq!(report.attributes, 0);
}

#[test]
fn inferred_external_relationship_declares_target_mode() {
    let source = pptx_fixture_with_slide_relationship(
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="tel:+49PPTX_SECRET_PHONE"/>"#,
    );
    pptx_parse::parse_pptx(&source).unwrap();

    let (output, _) = redact_with_report(&source, Format::Pptx).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let rels =
        String::from_utf8_lossy(part(&parts, "ppt/slides/_rels/slide1.xml.rels")).into_owned();
    assert!(rels.contains(r#"Target="https://example.com""#));
    assert!(rels.contains(r#"TargetMode="External""#));
    assert!(
        parts
            .iter()
            .all(|(_, bytes)| !String::from_utf8_lossy(bytes).contains("PPTX_SECRET_PHONE")),
        "secret survived: PPTX_SECRET_PHONE"
    );
    pptx_parse::parse_pptx(&output).unwrap();
}

#[test]
fn declared_target_mode_is_not_duplicated() {
    let source = pptx_fixture_with_slide_relationship(
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://secret.example/pptx" TargetMode="External"/>"#,
    );
    let (output, _) = redact_with_report(&source, Format::Pptx).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let rels =
        String::from_utf8_lossy(part(&parts, "ppt/slides/_rels/slide1.xml.rels")).into_owned();
    assert_eq!(rels.matches("TargetMode").count(), 1);
    pptx_parse::parse_pptx(&output).unwrap();
}

#[test]
fn declared_internal_mode_is_corrected_when_the_target_is_external() {
    let source = pptx_fixture_with_slide_relationship(
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="tel:+49PPTX_SECRET_PHONE" TargetMode="Internal"/>"#,
    );
    pptx_parse::parse_pptx(&source).unwrap();

    let (output, _) = redact_with_report(&source, Format::Pptx).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let rels =
        String::from_utf8_lossy(part(&parts, "ppt/slides/_rels/slide1.xml.rels")).into_owned();
    assert!(rels.contains(r#"Target="https://example.com""#));
    assert!(rels.contains(r#"TargetMode="External""#));
    assert!(!rels.contains("PPTX_SECRET_PHONE"));
    pptx_parse::parse_pptx(&output).unwrap();
}

#[test]
fn relative_target_with_parent_segments_stays_internal() {
    let target = "../../xl/nested/../worksheets/sheet1.xml";
    let source = fixture_with_part(
        xlsx_fixture(),
        "xl/_rels/workbook.xml.rels",
        xml(&format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{target}"/></Relationships>"#
        )),
    );
    let (output, _) = redact_with_report(&source, Format::Xlsx).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let rels = String::from_utf8_lossy(part(&parts, "xl/_rels/workbook.xml.rels")).into_owned();
    assert!(rels.contains(&format!(r#"Target="{target}""#)));
    assert!(!rels.contains("TargetMode"));
    xlsx_parse::parse_workbook(&parts).unwrap();
}

#[test]
fn foreign_attributes_do_not_drive_relationship_mode() {
    let rels = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:q="urn:qa">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml" q:Target="https://foreign.example/x"/>"#,
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml" q:TargetMode="External"/>"#,
        r#"</Relationships>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/_rels/document.xml.rels",
        rels.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"Target="footnotes.xml""#));
    assert!(text.contains(r#"Target="endnotes.xml""#));
    assert!(text.contains(r#"q:TargetMode="External""#));
    assert!(!text.contains(r#" TargetMode="External""#));
    assert_eq!(report.attributes, 0);
}

#[test]
fn target_inspection_is_limited_to_rels_parts() {
    let body = concat!(
        r#"<?xml version="1.0"?>"#,
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:q="urn:qa">"#,
        r#"<q:Relationship Target="tel:+49123456789"/>"#,
        r#"<q:Relationship Target="https://secret.example/x" TargetMode="External"/>"#,
        r#"</w:document>"#,
    );
    let mut report = RedactionReport::default();
    let output = xml::redact_xml(
        Format::Docx,
        "word/document.xml",
        body.as_bytes(),
        &mut report,
    )
    .unwrap();
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains(r#"Target="tel:+49123456789""#));
    assert!(!text.contains("secret.example"));
    assert_eq!(text.matches("TargetMode").count(), 1);
    assert_eq!(report.attributes, 1);
}

fn pptx_fixture_with_slide_relationship(relationship: &str) -> Vec<u8> {
    fixture_with_part(
        pptx_fixture(),
        "ppt/slides/_rels/slide1.xml.rels",
        xml(&format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationship}</Relationships>"#
        )),
    )
}

fn fixture_with_part(source: Vec<u8>, path: &str, data: Vec<u8>) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    for (candidate, bytes) in &mut parts {
        if candidate == path {
            *bytes = data.clone();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn media_placeholder_keeps_each_format() {
    for (format, ext) in [
        (ImageFormat::Gif, "gif"),
        (ImageFormat::Bmp, "bmp"),
        (ImageFormat::Tiff, "tiff"),
    ] {
        let source = placeholder_image(format);
        let mut report = RedactionReport::default();
        let part = format!("word/media/image1.{ext}");
        let output = media::replace_media(&part, &source, &mut report).unwrap();
        assert_ne!(source, output, "{ext} not redacted");
        assert_eq!(
            image::guess_format(&output).unwrap(),
            format,
            "{ext} format changed"
        );
        assert_eq!(
            image_dimensions(&output),
            (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE),
            "{ext} dims changed"
        );
    }
}

#[test]
fn rejects_unknown_extension_media() {
    let mut report = RedactionReport::default();
    let error = media::replace_media("word/media/blob.dat", b"not an image", &mut report);
    assert!(matches!(error, Err(RedactError::Image { .. })));
}

#[test]
fn replaces_wmf_with_valid_stub() {
    let mut report = RedactionReport::default();
    let output = media::replace_media(
        "docProps/thumbnail.wmf",
        b"\xd7\xcd\xc6\x9a metafile",
        &mut report,
    )
    .unwrap();
    assert_eq!(report.media_parts, 1);

    assert_eq!(&output[..4], &0x9AC6_CDD7u32.to_le_bytes());
    assert_eq!(le_u16(&output, 20), 0x52B1);
    assert_eq!(le_u16(&output, 10), 2000);
    assert_eq!(le_u16(&output, 12), 2000);
    assert_eq!(le_u16(&output, 14), 1440);

    let content = &output[22..];
    assert_eq!(le_u16(content, 0), 1);
    assert_eq!(le_u16(content, 2), 9);
    assert_eq!(le_u16(content, 4), 0x0300);
    let mt_size = le_u32(content, 6) as usize;
    assert_eq!(mt_size * 2, content.len());
    assert_eq!(le_u32(content, 12), 3);

    let eof = &content[18..];
    assert_eq!(le_u32(eof, 0), 3);
    assert_eq!(le_u16(eof, 4), 0);
    assert_eq!(eof.len(), 6);
}

#[test]
fn replaces_emf_with_valid_stub() {
    let mut report = RedactionReport::default();
    let output = media::replace_media("word/media/image1.emf", b"not an emf", &mut report).unwrap();
    assert_eq!(report.media_parts, 1);

    assert_eq!(le_u32(&output, 0), 1);
    assert_eq!(le_u32(&output, 4), 88);
    assert_eq!(le_u32(&output, 40), 0x464D_4520);
    assert_eq!(le_u32(&output, 44), 0x0001_0000);
    assert_eq!(le_u32(&output, 48) as usize, output.len());
    assert_eq!(le_u32(&output, 52), 2);
    assert_eq!(le_u16(&output, 56), 1);

    let eof = &output[88..];
    assert_eq!(eof.len(), 20);
    assert_eq!(le_u32(eof, 0), 14);
    assert_eq!(le_u32(eof, 4), 20);
    assert_eq!(le_u32(eof, 8), 0);
    assert_eq!(le_u32(eof, 12), 16);
    assert_eq!(le_u32(eof, 16), 20);
}

fn le_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(bytes[offset..offset + 2].try_into().unwrap())
}

fn le_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

/// PNG whose IHDR declares `width` x `height` with a matching chunk CRC, so a
/// dimension-preserving encoder really would allocate that many pixels.
fn png_declaring(width: u32, height: u32) -> Vec<u8> {
    let mut out = placeholder_png();
    out[16..20].copy_from_slice(&width.to_be_bytes());
    out[20..24].copy_from_slice(&height.to_be_bytes());
    let crc = png_crc(&out[12..29]);
    out[29..33].copy_from_slice(&crc.to_be_bytes());
    out
}

fn png_crc(bytes: &[u8]) -> u32 {
    let mut crc = u32::MAX;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            crc = if crc & 1 == 0 {
                crc >> 1
            } else {
                (crc >> 1) ^ 0xEDB8_8320
            };
        }
    }
    !crc
}

#[test]
fn oversized_declared_dimensions_emit_fixed_placeholder() {
    let hostile = png_declaring(8000, 8000);
    let mut report = RedactionReport::default();
    let output = media::replace_media("word/media/huge.png", &hostile, &mut report).unwrap();
    assert_eq!(
        image_dimensions(&output),
        (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE)
    );
    assert!(output.len() < 1024);
}

fn media_package(media: &[(String, Vec<u8>)]) -> Vec<u8> {
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".to_owned(),
            xml(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/><Default Extension="tiff" ContentType="image/tiff"/><Default Extension="svg" ContentType="image/svg+xml"/><Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="emf" ContentType="image/x-emf"/></Types>"#,
            ),
        ),
        (
            "_rels/.rels".to_owned(),
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
        ),
        (
            "word/document.xml".to_owned(),
            xml(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            ),
        ),
    ];
    parts.extend(media.iter().cloned());
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn hostile_package_media_budget_is_bounded() {
    let hostile = png_declaring(8000, 8000);
    let mut media: Vec<(String, Vec<u8>)> = vec![
        ("docProps/thumbnail.wmf".to_owned(), b"metafile".to_vec()),
        ("ppt/media/pic.emf".to_owned(), b"metafile".to_vec()),
    ];
    for index in 0..32 {
        media.push((format!("word/media/hostile{index}.png"), hostile.clone()));
    }

    let (output, report) = redact_with_report(&media_package(&media), Format::Auto).unwrap();
    assert_eq!(report.media_parts, media.len());
    let after = ooxml_opc::unzip_parts(&output).unwrap();
    let mut media_total = 0;
    for (path, bytes) in &after {
        if !path.contains("/media/") && !path.ends_with("thumbnail.wmf") {
            continue;
        }
        if path.ends_with(".png") {
            assert_eq!(
                image_dimensions(bytes),
                (media::PLACEHOLDER_SIZE, media::PLACEHOLDER_SIZE)
            );
        }
        media_total += bytes.len();
    }
    assert!(media_total < 256 * 1024);
}

const MEDIA_MARKER: &str = "MEDIA_SOURCE_MARKER";

/// Every replaceable media shape. `mask` is XORed over each wrapped payload and
/// its marker run, so two masks share no wrapped byte; mask 0 leaves the marker
/// and the source images verbatim. `mask` also picks the sniffable PNGs' width
/// and pixels, and those two encodings still share their signature, framing,
/// IEND tail and some IDAT bytes.
fn marked_media(mask: u8) -> Vec<(String, Vec<u8>)> {
    let wrap = |bytes: &[u8]| {
        let mut out = vec![mask; 8];
        out.extend(MEDIA_MARKER.bytes().map(|byte| byte ^ mask));
        out.extend(bytes.iter().map(|byte| byte ^ mask));
        out.extend(MEDIA_MARKER.bytes().map(|byte| byte ^ mask));
        out.extend_from_slice(&[mask; 8]);
        out
    };
    let mut encoded = Cursor::new(Vec::new());
    DynamicImage::ImageRgb8(ImageBuffer::from_pixel(
        3 + u32::from(mask % 2),
        2,
        Rgb([mask, mask, mask]),
    ))
    .write_to(&mut encoded, ImageFormat::Png)
    .unwrap();
    let sniffable = encoded.into_inner();
    vec![
        ("word/media/image1.png".to_owned(), wrap(&placeholder_png())),
        (
            "word/media/photo.jpg".to_owned(),
            wrap(&placeholder_image(ImageFormat::Jpeg)),
        ),
        (
            "word/media/anim.gif".to_owned(),
            wrap(&placeholder_image(ImageFormat::Gif)),
        ),
        (
            "word/media/raster.bmp".to_owned(),
            wrap(&placeholder_image(ImageFormat::Bmp)),
        ),
        (
            "word/media/scan.tiff".to_owned(),
            wrap(&placeholder_image(ImageFormat::Tiff)),
        ),
        (
            "word/media/vector.svg".to_owned(),
            wrap(
                format!(
                    r#"<svg xmlns="http://www.w3.org/2000/svg"><desc>{MEDIA_MARKER}</desc></svg>"#
                )
                .as_bytes(),
            ),
        ),
        (
            "word/media/legacy.wmf".to_owned(),
            wrap(&[0xd7, 0xcd, 0xc6, 0x9a]),
        ),
        (
            "word/media/legacy.emf".to_owned(),
            wrap(&[0x01, 0x00, 0x00, 0x00]),
        ),
        (
            "docProps/thumbnail.wmf".to_owned(),
            wrap(&[0xd7, 0xcd, 0xc6, 0x9a]),
        ),
        (
            "word/media/MixedCase.PNG".to_owned(),
            wrap(&placeholder_png()),
        ),
        ("word/media/sniffed".to_owned(), sniffable.clone()),
        ("docProps/thumbnail".to_owned(), sniffable),
        (
            "word/media/mislabelled.png".to_owned(),
            wrap(&placeholder_image(ImageFormat::Jpeg)),
        ),
        (
            "word/media/mislabelled.emf".to_owned(),
            wrap(&placeholder_png()),
        ),
        (
            "word/media/oversized.png".to_owned(),
            wrap(&png_declaring(8000, 8000)),
        ),
        ("word/media/empty.png".to_owned(), Vec::new()),
        ("word/media/empty.wmf".to_owned(), Vec::new()),
    ]
}

#[test]
fn media_replacement_never_copies_source_bytes() {
    let first = marked_media(0);
    let second = marked_media(0x5f);
    for ((path, wrapped), (_, other)) in first.iter().zip(&second) {
        if matches!(path.as_str(), "word/media/sniffed" | "docProps/thumbnail") {
            continue;
        }
        assert!(
            wrapped.iter().zip(other).all(|(one, two)| one != two),
            "fixtures share a byte in {path}, so a copy of it would go unnoticed"
        );
    }

    let (output, report) = redact_with_report(&media_package(&first), Format::Auto).unwrap();
    assert_eq!(report.media_parts, first.len());
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    for (path, bytes) in &parts {
        assert!(
            !bytes
                .windows(MEDIA_MARKER.len())
                .any(|window| window == MEDIA_MARKER.as_bytes()),
            "source bytes survived in {path}"
        );
    }

    let (other, _) = redact_with_report(&media_package(&second), Format::Auto).unwrap();
    assert_eq!(parts, ooxml_opc::unzip_parts(&other).unwrap());
}
