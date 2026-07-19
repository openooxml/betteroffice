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
fn preserves_jpeg_dimensions_and_format() {
    let source = placeholder_image(ImageFormat::Jpeg);
    let mut report = RedactionReport::default();
    let output = media::replace_media("word/media/photo.jpeg", &source, &mut report).unwrap();
    assert_ne!(source, output);
    assert_eq!(image_dimensions(&source), image_dimensions(&output));
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
        image_dimensions(before_image),
        image_dimensions(after_image)
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
            image_dimensions(&source),
            image_dimensions(&output),
            "{ext} dims changed"
        );
    }
}

#[test]
fn media_rejects_unencodable_formats() {
    let mut report = RedactionReport::default();
    let error = media::replace_media("word/media/image1.emf", b"not an image", &mut report);
    assert!(matches!(error, Err(RedactError::Image { .. })));
}
