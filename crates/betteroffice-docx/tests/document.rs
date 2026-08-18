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

const TEXT_BOX_NAMESPACES: &str = concat!(
    r#" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
    r#" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
    r#" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""#,
    r#" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#,
    r#" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#,
    r#" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="wps v""#,
);

/// The GB/T callout shape from issue #202: an `mc:Choice` textbox with a VML
/// fallback, a bare textbox, and character-unit first-line indents.
fn text_box_docx() -> Vec<u8> {
    let body = concat!(
        r##"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>1000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>2000</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="11" name="Text Box 11"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:ind w:firstLineChars="200" w:firstLine="420"/></w:pPr><w:r><w:t>Choice callout</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="_x0000_s1026" type="#_x0000_t202"><v:textbox><w:txbxContent><w:p><w:r><w:t>Fallback callout</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>"##,
        r#"<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="21" name="Text Box 21"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Inline callout</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
        r#"<w:p><w:pPr><w:ind w:firstLineChars="200" w:firstLine="420"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>"#,
        r#"<w:p><w:pPr><w:ind w:firstLineChars="0" w:firstLine="0"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>"#,
    );
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document{TEXT_BOX_NAMESPACES}><w:body>{body}</w:body></w:document>"#
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
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn saving_keeps_shape_text_box_bodies_and_character_unit_indents() {
    let document = Document::open(&text_box_docx()).unwrap();
    let saved = document.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let xml = saved_part(&parts, "word/document.xml");

    assert_eq!(xml.matches("<w:txbxContent>").count(), 2);
    assert!(xml.contains("Choice callout"));
    assert!(xml.contains("Inline callout"));
    assert_eq!(xml.matches(r#"<wps:cNvSpPr txBox="1"/>"#).count(), 2);

    assert_eq!(xml.matches(r#"w:firstLineChars="200""#).count(), 2);
    assert_eq!(xml.matches(r#"w:firstLine="420""#).count(), 2);
    assert!(xml.contains(r#"<w:ind w:firstLine="0" w:firstLineChars="0"/>"#));

    let reopened = Document::open(&saved).unwrap();
    let resaved = saved_part(
        &ooxml_opc::unzip_parts(&reopened.save().unwrap()).unwrap(),
        "word/document.xml",
    );
    assert_eq!(resaved, xml);
}

/// A minimal package holding `body`, plus `word/numbering.xml` when given.
fn story_docx(body: &str, numbering: Option<&str>) -> Vec<u8> {
    let numbering_override = if numbering.is_some() {
        r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#
    } else {
        ""
    };
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document{TEXT_BOX_NAMESPACES}><w:body>{body}</w:body></w:document>"#
    );
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{numbering_override}</Types>"#
            )
            .into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ];
    if let Some(numbering) = numbering {
        parts.push((
            "word/_rels/document.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>"#.to_vec(),
        ));
        parts.push((
            "word/numbering.xml".to_owned(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{numbering}</w:numbering>"#
            )
            .into_bytes(),
        ));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The document part after `Document::open` → `save`.
fn saved_document(package: &[u8]) -> String {
    let saved = Document::open(package).unwrap().save().unwrap();
    saved_part(
        &ooxml_opc::unzip_parts(&saved).unwrap(),
        "word/document.xml",
    )
}

fn text_box_drawing(body: &str, body_properties: &str) -> String {
    format!(
        r#"<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:docPr id="31" name="Text Box 31"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>{body}</w:txbxContent></wps:txbx>{body_properties}</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
    )
}

#[test]
fn saving_keeps_a_table_inside_a_text_box() {
    let table = r#"<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>TABLE-CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let package = story_docx(
        &text_box_drawing(table, r#"<wps:bodyPr rot="0" vert="horz"/>"#),
        None,
    );

    let xml = saved_document(&package);

    assert!(xml.contains("TABLE-CELL"));
    let body = xml
        .split_once("<w:txbxContent>")
        .and_then(|(_, rest)| rest.split_once("</w:txbxContent>"))
        .unwrap()
        .0;
    assert!(body.starts_with("<w:tbl>"));
    assert_eq!(body.matches("<w:tc>").count(), 1);

    let resaved = saved_document(&Document::open(&package).unwrap().save().unwrap());
    assert_eq!(resaved, xml);
}

/// The parser reads `anchor` into a canonical name; every one of them has to
/// be written back as the schema token, or the next open reads it as none.
#[test]
fn saving_keeps_every_text_box_anchor() {
    for token in ["t", "ctr", "b", "dist", "just"] {
        let package = story_docx(
            &text_box_drawing(
                r#"<w:p><w:r><w:t>Anchored</w:t></w:r></w:p>"#,
                &format!(r#"<wps:bodyPr rot="0" vert="horz" anchor="{token}"/>"#),
            ),
            None,
        );
        let xml = saved_document(&package);
        assert!(
            xml.contains(&format!(r#"anchor="{token}""#)),
            "{token}: {xml}"
        );
        let resaved = saved_document(&Document::open(&package).unwrap().save().unwrap());
        assert_eq!(resaved, xml, "{token} second save");
    }
}

#[test]
fn saving_keeps_the_writing_direction_of_a_vertical_text_box() {
    let package = story_docx(
        &text_box_drawing(
            r#"<w:p><w:r><w:t>Vertical</w:t></w:r></w:p>"#,
            r#"<wps:bodyPr rot="0" vert="eaVert"/>"#,
        ),
        None,
    );

    let xml = saved_document(&package);

    assert!(xml.contains(r#"<wps:bodyPr rot="0" vert="eaVert"/>"#));
    assert!(!xml.contains(r#"vert="horz""#));

    let resaved = saved_document(&Document::open(&package).unwrap().save().unwrap());
    assert_eq!(resaved, xml);
}

const HANGING_NUMBERING: &str = r#"<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>"#;

#[test]
fn a_direct_character_first_line_indent_outranks_a_numbering_hanging_indent() {
    let package = story_docx(
        r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:firstLineChars="200"/></w:pPr><w:r><w:t>Numbered</w:t></w:r></w:p>"#,
        Some(HANGING_NUMBERING),
    );

    let xml = saved_document(&package);

    assert!(xml.contains(r#"<w:ind w:left="720" w:firstLineChars="200"/>"#));
    assert!(!xml.contains("hangingChars"));
    assert!(!xml.contains("w:hanging="));

    let resaved = saved_document(&Document::open(&package).unwrap().save().unwrap());
    assert_eq!(resaved, xml);
}

#[test]
fn mixed_unit_indents_keep_the_direction_each_unit_was_authored_with() {
    let package = story_docx(
        r#"<w:p><w:pPr><w:ind w:firstLine="420" w:hangingChars="200"/></w:pPr><w:r><w:t>Mixed</w:t></w:r></w:p>"#,
        None,
    );

    let xml = saved_document(&package);

    assert!(xml.contains(r#"<w:ind w:firstLine="420" w:hangingChars="200"/>"#));

    let resaved = saved_document(&Document::open(&package).unwrap().save().unwrap());
    assert_eq!(resaved, xml);
}
