use betteroffice_docx::{
    BlockContent, Document, InlineNode, LayoutInput, ParagraphContent, ParseLimits, RunContent,
    get_paragraph_text,
};

fn sample_docx() -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p w14:paraId="11111111"><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:pPr><w:r><w:t>Hello DOCX</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="22222222"><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="33333333"><w:r><w:t>Second section</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#.to_vec(),
        ),
        (
            "word/header1.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p w14:paraId="44444444"><w:r><w:t>Native header</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn opens_edits_saves_and_reopens_typed_structure() {
    let mut document = Document::open(&sample_docx()).unwrap();
    let structure = document.structure();
    assert_eq!(structure.body_paragraphs, 3);
    assert_eq!(structure.body_tables, 1);
    assert_eq!(structure.sections, 2);
    assert_eq!(structure.headers, 1);
    assert_eq!(document.headers()[0].1.content.len(), 1);
    assert_eq!(
        get_paragraph_text(document.paragraph("11111111").unwrap()),
        "Hello DOCX"
    );

    let receipt = document
        .replace_paragraph_text("11111111", "Edited natively")
        .unwrap();
    assert_eq!(receipt.range.unwrap().start.para, "11111111");

    let saved = document.save().unwrap();
    let reopened = Document::open(&saved).unwrap();
    assert_eq!(reopened.structure(), structure);
    assert_eq!(
        get_paragraph_text(reopened.paragraph("11111111").unwrap()),
        "Edited natively"
    );
    assert_eq!(reopened.tables().len(), 1);
    assert_eq!(reopened.sections().len(), 2);
    assert_eq!(reopened.headers().len(), 1);
}

fn assert_send_sync<T: Send + Sync>() {}

/// `Document` backs language bindings (`#[pyclass]` requires `Send`), so this
/// gate is permanent under every feature combination.
#[test]
fn document_is_send_and_sync() {
    assert_send_sync::<Document>();
}

#[test]
fn open_with_limits_rejects_a_document_over_budget() {
    let bytes = sample_docx();
    let limits = ParseLimits {
        max_paragraphs: 2,
        ..ParseLimits::default()
    };
    let Err(error) = Document::open_with_limits(&bytes, &limits) else {
        panic!("a 3-paragraph document parsed under a 2-paragraph budget");
    };
    assert!(
        error.to_string().contains("paragraph"),
        "unexpected error: {error}"
    );
    assert_eq!(
        Document::open_with_limits(&bytes, &ParseLimits::default())
            .unwrap()
            .structure()
            .body_paragraphs,
        3
    );
}

#[test]
fn lays_out_typed_input_and_builds_a_display_list() {
    let document = Document::open(&sample_docx()).unwrap();
    let input: LayoutInput = serde_json::from_str(include_str!(
        "../../docx-layout/tests/fixtures/single-page-multi-paragraph.input.json"
    ))
    .unwrap();
    let result = document.layout(input).unwrap();
    assert_eq!(result.layout.pages.len(), 1);
    assert_eq!(result.display_list.pages.len(), 1);
    assert!(!result.display_list.pages[0].primitives.is_empty());
}

const DAMAGED_CHART: &[u8] = b"<c:chartSpace><c:chart></c:chartSpace>";

/// A document whose only chart part cannot be read.
fn damaged_chart_docx() -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p w14:paraId="11111111"><w:r><w:t>Hello DOCX</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#.to_vec(),
        ),
        ("word/charts/chart1.xml".to_owned(), DAMAGED_CHART.to_vec()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Declining to read a chart part must not stop the package from carrying it:
/// an untouched save still writes the source bytes back.
#[test]
fn an_unreadable_chart_part_survives_a_save_byte_for_byte() {
    let bytes = damaged_chart_docx();
    let document = Document::open(&bytes).unwrap();
    assert_eq!(document.structure().body_paragraphs, 2);
    assert_eq!(
        get_paragraph_text(document.paragraph("11111111").unwrap()),
        "Hello DOCX"
    );
    assert!(document.model().charts.is_empty());

    let before = ooxml_opc::unzip_parts(&bytes).unwrap();
    let after = ooxml_opc::unzip_parts(&document.save().unwrap()).unwrap();
    let part_bytes = |parts: &[(String, Vec<u8>)]| {
        parts
            .iter()
            .find(|(path, _)| path == "word/charts/chart1.xml")
            .map(|(_, bytes)| bytes.clone())
            .unwrap()
    };
    assert_eq!(part_bytes(&after), DAMAGED_CHART);
    assert_eq!(part_bytes(&after), part_bytes(&before));
    assert_eq!(
        after.iter().map(|(path, _)| path).collect::<Vec<_>>(),
        before.iter().map(|(path, _)| path).collect::<Vec<_>>()
    );
}

const CHART_DRAWING: &str = r#"<w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId5"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#;

const STORY_NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart""#;

/// A Word-shaped package whose body, header and footnote each hold the same
/// chart drawing, and whose `rId1` is the styles part Word always puts there.
fn charted_story_docx(chart_part: Option<&[u8]>) -> Vec<u8> {
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/styles.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {STORY_NAMESPACES}><w:body><w:p w14:paraId="11111111"><w:r><w:t>Hello DOCX</w:t></w:r></w:p><w:p w14:paraId="33333333"><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p w14:paraId="22222222"><w:r>{CHART_DRAWING}</w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#
            )
            .into_bytes(),
        ),
        (
            "word/header1.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {STORY_NAMESPACES}><w:p w14:paraId="44444444"><w:r><w:t>Native header</w:t></w:r></w:p><w:p w14:paraId="55555555"><w:r>{CHART_DRAWING}</w:r></w:p></w:hdr>"#
            )
            .into_bytes(),
        ),
        (
            "word/footnotes.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {STORY_NAMESPACES}><w:footnote w:id="-1" w:type="separator"><w:p w14:paraId="77777777"><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p w14:paraId="66666666"><w:r><w:t>Note text</w:t></w:r><w:r>{CHART_DRAWING}</w:r></w:p></w:footnote></w:footnotes>"#
            )
            .into_bytes(),
        ),
    ];
    if let Some(bytes) = chart_part {
        parts.push(("word/charts/chart1.xml".to_owned(), bytes.to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The drawings in `content` no chart was read for.
fn opaque_drawings(content: &[BlockContent]) -> usize {
    content
        .iter()
        .filter_map(|block| match block {
            BlockContent::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .flat_map(|paragraph| &paragraph.content)
        .filter_map(|item| match item {
            ParagraphContent::Inline(InlineNode::Run(run)) => Some(run),
            _ => None,
        })
        .flat_map(|run| &run.content)
        .filter(|item| matches!(item, RunContent::OpaqueDrawing { .. }))
        .count()
}

fn saved_part(parts: &[(String, Vec<u8>)], path: &str) -> String {
    parts
        .iter()
        .find(|(name, _)| name == path)
        .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        .unwrap_or_else(|| panic!("{path} is not in the package"))
}

/// A drawing whose chart part the parser has no chart for must stay opaque:
/// reading it as a picture would invent a `<a:blip r:embed="rId1"/>`, and in a
/// Word-shaped package `rId1` is the styles part. A full save drops the
/// drawing, exactly as it drops a chart it did read.
fn assert_no_stray_picture(chart_part: Option<&[u8]>) {
    let bytes = charted_story_docx(chart_part);
    let mut document = Document::open(&bytes).unwrap();
    assert_eq!(document.structure().body_paragraphs, 3);
    assert!(document.model().charts.is_empty());
    for story in [
        document.model().body.content.clone(),
        document.headers()[0].1.content.clone(),
        document.model().footnotes[0].content.clone(),
    ] {
        assert_eq!(opaque_drawings(&story), 1);
    }

    document
        .replace_paragraph_text("11111111", "Edited natively")
        .unwrap();
    let saved = document.save().unwrap();

    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    for path in [
        "word/document.xml",
        "word/header1.xml",
        "word/footnotes.xml",
    ] {
        let xml = saved_part(&parts, path);
        assert!(!xml.contains("pic:pic"), "{path} gained a picture");
        assert!(!xml.contains("a:blip"), "{path} gained a picture");
        assert!(
            !xml.contains(r#"r:embed="rId1""#),
            "{path} embeds the styles part"
        );
    }
    assert!(saved_part(&parts, "word/document.xml").contains("Edited natively"));
    assert_eq!(
        chart_part,
        parts
            .iter()
            .find(|(path, _)| path == "word/charts/chart1.xml")
            .map(|(_, bytes)| bytes.as_slice())
    );

    let reopened = Document::open(&saved).unwrap();
    assert_eq!(reopened.structure(), document.structure());
    assert_eq!(opaque_drawings(&reopened.model().body.content), 0);
    assert_eq!(
        get_paragraph_text(reopened.paragraph("11111111").unwrap()),
        "Edited natively"
    );
}

#[test]
fn an_unreadable_chart_part_writes_no_picture_in_any_story() {
    assert_no_stray_picture(Some(DAMAGED_CHART));
}

#[test]
fn an_absent_chart_part_writes_no_picture_in_any_story() {
    assert_no_stray_picture(None);
}
