use std::io::Cursor;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

use image::{DynamicImage, ImageFormat, Rgba, RgbaImage};

static FIXTURE_ID: AtomicU64 = AtomicU64::new(0);

const PACKAGE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

pub fn complex_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>"#;
    let relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHyperlink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/><Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p w14:paraId="11111111"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:hyperlink r:id="rIdHyperlink"><w:r><w:t xml:space="preserve"> link</w:t></w:r></w:hyperlink><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>Plain paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>"#;
    let footnotes = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Footnote body</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
    package(vec![
        ("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ("_rels/.rels", PACKAGE_RELS.as_bytes().to_vec()),
        ("word/document.xml", document.as_bytes().to_vec()),
        (
            "word/_rels/document.xml.rels",
            relationships.as_bytes().to_vec(),
        ),
        ("word/footnotes.xml", footnotes.as_bytes().to_vec()),
    ])
}

pub fn editing_docx(body: &str) -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>"#
    );
    package(vec![
        ("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ("_rels/.rels", PACKAGE_RELS.as_bytes().to_vec()),
        ("word/document.xml", document.into_bytes()),
    ])
}

pub fn inline_embed_docx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>"#;
    let relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>"#;
    let document = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>Bold link</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t>a😀</w:t></w:r></w:p><w:p w14:paraId="33333333"><w:r><w:t>Bold</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p><w:p w14:paraId="44444444"><w:r><w:t>Hello 😀</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>"#;
    let footnotes = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Footnote body</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
    package(vec![
        ("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ("_rels/.rels", PACKAGE_RELS.as_bytes().to_vec()),
        ("word/document.xml", document.as_bytes().to_vec()),
        (
            "word/_rels/document.xml.rels",
            relationships.as_bytes().to_vec(),
        ),
        ("word/footnotes.xml", footnotes.as_bytes().to_vec()),
    ])
}

pub fn image_docx(rotation_degrees: Option<f64>) -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    let relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>"#;
    let rotation = rotation_degrees
        .map(|value| format!(r#" rot="{}""#, (value * 60_000.0).round() as i64))
        .unwrap_or_default();
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p w14:paraId="11111111"><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Green fixture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Green fixture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm{rotation}><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>"#
    );
    let pixels = RgbaImage::from_pixel(160, 80, Rgba([0, 255, 0, 255]));
    let mut png = Cursor::new(Vec::new());
    DynamicImage::ImageRgba8(pixels)
        .write_to(&mut png, ImageFormat::Png)
        .unwrap();
    package(vec![
        ("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ("_rels/.rels", PACKAGE_RELS.as_bytes().to_vec()),
        ("word/document.xml", document.into_bytes()),
        (
            "word/_rels/document.xml.rels",
            relationships.as_bytes().to_vec(),
        ),
        ("word/media/image1.png", png.into_inner()),
    ])
}

pub fn large_xlsx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    let package_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
    let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Large" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let mut rows = String::new();
    for row in 1..=410 {
        rows.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>A</t></is></c><c r="B{row}" t="inlineStr"><is><t>B</t></is></c></row>"#
        ));
    }
    let sheet = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B410"/><sheetData>{rows}</sheetData></worksheet>"#
    );
    package(vec![
        ("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ("_rels/.rels", package_rels.as_bytes().to_vec()),
        ("xl/workbook.xml", workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels",
            workbook_rels.as_bytes().to_vec(),
        ),
        ("xl/worksheets/sheet1.xml", sheet.into_bytes()),
    ])
}

pub fn shared_formula_xlsx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    let package_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
    let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Shared" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#;
    let sheet = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row></sheetData></worksheet>"#;
    package(vec![
        ("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ("_rels/.rels", package_rels.as_bytes().to_vec()),
        ("xl/workbook.xml", workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels",
            workbook_rels.as_bytes().to_vec(),
        ),
        ("xl/worksheets/sheet1.xml", sheet.as_bytes().to_vec()),
    ])
}

pub fn write_docx(label: &str, bytes: &[u8]) -> PathBuf {
    write_fixture(label, "docx", bytes)
}

pub fn write_xlsx(label: &str, bytes: &[u8]) -> PathBuf {
    write_fixture(label, "xlsx", bytes)
}

pub fn write_pptx(label: &str, bytes: &[u8]) -> PathBuf {
    write_fixture(label, "pptx", bytes)
}

pub fn signed_pptx() -> Vec<u8> {
    let mut parts = demo_pptx_parts();
    parts.push((
        "_xmlsignatures/sig1.xml".to_owned(),
        br#"<?xml version="1.0" encoding="UTF-8"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>"#
            .to_vec(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

pub fn transformed_pptx() -> Vec<u8> {
    let mut parts = demo_pptx_parts();
    let slide = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    let mut xml = String::from_utf8(std::mem::take(&mut slide.1)).unwrap();
    let marker = r#"<p:cNvPr id="3" name="Deck label"/>"#;
    let marker_start = xml.find(marker).unwrap();
    let transform_start = marker_start + xml[marker_start..].find("<a:xfrm").unwrap();
    xml.insert_str(
        transform_start + "<a:xfrm".len(),
        r#" rot="1800000" flipH="1""#,
    );
    slide.1 = xml.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

pub fn theme_linked_pptx() -> Vec<u8> {
    let mut parts = demo_pptx_parts();
    let slide = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    let xml = String::from_utf8(std::mem::take(&mut slide.1)).unwrap();
    let mut rewritten = String::with_capacity(xml.len());
    let mut cursor = 0;
    while let Some(offset) = xml[cursor..].find("<a:srgbClr") {
        let start = cursor + offset;
        let end = start + xml[start..].find("/>").unwrap() + 2;
        rewritten.push_str(&xml[cursor..start]);
        rewritten.push_str(r#"<a:schemeClr val="accent1"/>"#);
        cursor = end;
    }
    rewritten.push_str(&xml[cursor..]);
    slide.1 = rewritten.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

pub fn language_neutral_pptx() -> Vec<u8> {
    let mut parts = demo_pptx_parts();
    let slide = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    let xml = String::from_utf8(std::mem::take(&mut slide.1)).unwrap();
    slide.1 = xml.replace(r#" lang="en-US""#, "").into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn demo_pptx_parts() -> Vec<(String, Vec<u8>)> {
    ooxml_opc::unzip_parts(include_bytes!("../../demo/public/betteroffice-demo.pptx")).unwrap()
}

fn write_fixture(label: &str, extension: &str, bytes: &[u8]) -> PathBuf {
    let path = std::env::temp_dir().join(format!(
        "betteroffice-native-viewer-{label}-{}-{}.{extension}",
        std::process::id(),
        FIXTURE_ID.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::write(&path, bytes).unwrap();
    path
}

pub fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path.eq_ignore_ascii_case(name))
        .map(|(_, bytes)| bytes)
        .unwrap()
}

pub fn paragraph(bytes: &[u8], para_id: &str) -> Vec<u8> {
    let xml = String::from_utf8(part(bytes, "word/document.xml")).unwrap();
    let marker = format!(r#"w14:paraId="{para_id}""#);
    let marker_start = xml.find(&marker).unwrap();
    let start = xml[..marker_start].rfind("<w:p").unwrap();
    let end = marker_start + xml[marker_start..].find("</w:p>").unwrap() + "</w:p>".len();
    xml.as_bytes()[start..end].to_vec()
}

fn package(parts: Vec<(&str, Vec<u8>)>) -> Vec<u8> {
    let parts = parts
        .into_iter()
        .map(|(path, bytes)| (path.to_owned(), bytes))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}
