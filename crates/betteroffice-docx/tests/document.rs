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
/// Word-shaped package `rId1` is the styles part. A full save replays the
/// drawing verbatim, exactly as it replays a chart it did read.
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
        assert!(xml.contains(CHART_DRAWING), "{path} lost its drawing");
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
    assert_eq!(opaque_drawings(&reopened.model().body.content), 1);
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

const BAR_CHART: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#;

/// The charted runs in `content`.
fn chart_runs(content: &[BlockContent]) -> usize {
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
        .filter(|item| matches!(item, RunContent::Chart { .. }))
        .count()
}

fn stories(document: &Document) -> [Vec<BlockContent>; 3] {
    [
        document.model().body.content.clone(),
        document.headers()[0].1.content.clone(),
        document.model().footnotes[0].content.clone(),
    ]
}

/// A chart the parser did read is placed by its `w:drawing`, which a save
/// replays verbatim: the chart part, its relationship and the drawing all
/// come back on reopen, in the body, a header and a footnote alike. The one
/// finding is the note writer moving the root's chart binding onto the
/// `w:footnote` element, as it does for any foreign markup in a note.
#[test]
fn a_chart_survives_a_save_in_every_story() {
    let bytes = charted_story_docx(Some(BAR_CHART));
    let document = Document::open(&bytes).unwrap();
    assert!(!document.model().charts.is_empty());
    for story in stories(&document) {
        assert_eq!(chart_runs(&story), 1);
    }

    let saved = document.save().unwrap();
    let before = ooxml_opc::unzip_parts(&bytes).unwrap();
    let after = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(
        ooxml_fidelity::roundtrip_findings(&before, &after).unwrap(),
        vec!["fingerprint differs: word/footnotes.xml"]
    );
    for path in [
        "word/document.xml",
        "word/header1.xml",
        "word/footnotes.xml",
    ] {
        assert!(
            saved_part(&after, path).contains(CHART_DRAWING),
            "{path} lost its chart drawing"
        );
    }

    let mut reopened = Document::open(&saved).unwrap();
    assert_eq!(reopened.structure(), document.structure());
    for story in stories(&reopened) {
        assert_eq!(chart_runs(&story), 1);
    }

    reopened
        .replace_paragraph_text("11111111", "Edited natively")
        .unwrap();
    let edited = ooxml_opc::unzip_parts(&reopened.save().unwrap()).unwrap();
    assert_eq!(
        ooxml_fidelity::losses(
            &ooxml_fidelity::element_census(&before).unwrap(),
            &ooxml_fidelity::element_census(&edited).unwrap()
        ),
        vec![]
    );
    assert!(saved_part(&edited, "word/document.xml").contains("Edited natively"));
    assert_eq!(
        saved_part(&edited, "word/charts/chart1.xml").as_bytes(),
        BAR_CHART
    );
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

    // Level indents ride list rendering; w:ind saves back as authored.
    assert!(xml.contains(r#"<w:ind w:firstLineChars="200"/>"#));
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

/// Written exactly as a save serializes it, so an unedited save reproduces it.
const DUPLICATE_ID_DOCUMENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">"#,
    r#"<w:body><w:p w14:paraId="1A2B3C4D"><w:r><w:t>first</w:t></w:r></w:p>"#,
    r#"<w:p w14:paraId="1A2B3C4D"><w:r><w:t>second</w:t></w:r></w:p>"#,
    r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>"#,
    r#"<w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>"#,
    r#"<w:p w14:paraId="1A2B3C4D"><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    r#"<w:p w14:paraId="0B000003"><w:r><w:t>third</w:t></w:r></w:p>"#,
    r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
);

/// Three body paragraphs, one in a table cell, share the source ID `1A2B3C4D`.
fn duplicate_id_docx() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&story_docx("", None)).unwrap();
    for (name, bytes) in &mut parts {
        if name == "word/document.xml" {
            *bytes = DUPLICATE_ID_DOCUMENT.as_bytes().to_vec();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn identified_texts(document: &Document) -> Vec<(String, String)> {
    document
        .paragraphs()
        .into_iter()
        .map(|paragraph| {
            (
                paragraph.para_id.clone().unwrap_or_default(),
                get_paragraph_text(paragraph),
            )
        })
        .collect()
}

fn paragraph_ids(document: &Document) -> Vec<String> {
    identified_texts(document)
        .into_iter()
        .map(|(id, _)| id)
        .collect()
}

#[test]
fn a_repeated_paragraph_id_addresses_its_own_paragraph() {
    let mut document = Document::open(&duplicate_id_docx()).unwrap();
    let ids = paragraph_ids(&document);
    assert_eq!((ids[0].as_str(), ids[3].as_str()), ("1A2B3C4D", "0B000003"));
    assert_eq!(
        ids.iter().collect::<std::collections::BTreeSet<_>>().len(),
        4,
        "{ids:?}"
    );
    assert_eq!(
        get_paragraph_text(document.paragraph(&ids[1]).unwrap()),
        "second"
    );

    let receipt = document
        .replace_paragraph_text(&ids[1], "replacement")
        .unwrap();
    assert_eq!(receipt.range.unwrap().start.para, ids[1]);
    let section_texts: Vec<String> = document.sections()[0]
        .content
        .iter()
        .filter_map(|block| match block {
            BlockContent::Paragraph(paragraph) => Some(get_paragraph_text(paragraph)),
            _ => None,
        })
        .collect();
    assert_eq!(section_texts, ["first", "replacement", "third"]);

    let saved = document.save().unwrap();
    let xml = saved_part(
        &ooxml_opc::unzip_parts(&saved).unwrap(),
        "word/document.xml",
    );
    assert_eq!(xml.matches(r#"w14:paraId="1A2B3C4D""#).count(), 2);
    assert_eq!(
        identified_texts(&Document::open(&saved).unwrap()),
        [
            ("1A2B3C4D", "first"),
            (ids[1].as_str(), "replacement"),
            (ids[2].as_str(), "cell"),
            ("0B000003", "third"),
        ]
        .map(|(id, text)| (id.to_owned(), text.to_owned()))
    );
}

#[test]
fn an_unedited_repeated_paragraph_id_saves_as_authored() {
    let document = Document::open(&duplicate_id_docx()).unwrap();
    let ids = paragraph_ids(&document);
    assert_eq!(
        get_paragraph_text(document.paragraph(&ids[2]).unwrap()),
        "cell"
    );

    let saved = document.save().unwrap();
    assert_eq!(
        saved_part(
            &ooxml_opc::unzip_parts(&saved).unwrap(),
            "word/document.xml"
        ),
        DUPLICATE_ID_DOCUMENT
    );
}

#[test]
fn an_ambiguous_paragraph_id_is_refused() {
    let mut document = Document::open(&duplicate_id_docx()).unwrap();
    let BlockContent::Paragraph(last) = &mut document.model_mut().body.content[3] else {
        panic!("paragraph")
    };
    std::sync::Arc::make_mut(last).para_id = Some("1A2B3C4D".to_owned());

    assert!(document.paragraph("1A2B3C4D").is_none());
    assert!(matches!(
        document.replace_paragraph_text("1A2B3C4D", "replacement"),
        Err(betteroffice_docx::Error::AmbiguousParagraph(_))
    ));
}

const W14_ID: &str = "1A2B3C4D";
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 w15""#;
const XML_DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

fn word_paragraph(id: &str, text: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}" w14:textId="77777777"><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

/// A package as Word writes it rather than as a save serializes it: `1A2B3C4D`
/// repeats in the body, a nested table cell and a content control, and the
/// last body paragraph's ID `0B000003` is written with a character reference.
fn word_docx(edit: impl FnOnce(&mut Vec<(String, String)>)) -> Vec<u8> {
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let cell = |content: String| {
        format!(
            r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr></w:tbl>"#
        )
    };
    let body = [
        word_paragraph(W14_ID, "first"),
        word_paragraph(W14_ID, "second"),
        cell(cell(word_paragraph(W14_ID, "nested")) + &word_paragraph("0B000002", "")),
        format!(
            r#"<w:sdt><w:sdtPr><w:alias w:val="Clause"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
            word_paragraph(W14_ID, "control")
        ),
        word_paragraph("0B00000&#x33;", "third"),
    ]
    .concat();
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"{XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#
            ),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"{XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#
            ),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"{XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="{REL}/header" Target="header1.xml"/></Relationships>"#
            ),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"{XML_DECLARATION}<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
            ),
        ),
        (
            "word/header1.xml".to_owned(),
            format!(
                "{XML_DECLARATION}<w:hdr {NAMESPACES}>{}</w:hdr>",
                word_paragraph("0C000001", "header")
            ),
        ),
        (
            "customXml/item1.xml".to_owned(),
            format!("{XML_DECLARATION}<data xmlns=\"urn:custom\">kept</data>"),
        ),
    ];
    edit(&mut parts);
    let parts: Vec<(String, Vec<u8>)> = parts
        .into_iter()
        .map(|(name, xml)| (name, xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The `w14:paraId` of each `w:p` start tag in a serialized part.
fn written_ids(xml: &str) -> Vec<String> {
    xml.match_indices("<w:p")
        .filter(|(at, _)| matches!(xml.as_bytes()[at + 4], b' ' | b'>' | b'/'))
        .map(|(at, _)| {
            let tag = &xml[at..at + xml[at..].find('>').unwrap()];
            tag.split(r#"w14:paraId=""#)
                .nth(1)
                .and_then(|rest| rest.split('"').next())
                .unwrap_or_default()
                .to_owned()
        })
        .collect()
}

fn package_part(package: &[u8], path: &str) -> String {
    saved_part(&ooxml_opc::unzip_parts(package).unwrap(), path)
}

/// The IDs `word_docx` authors in the parts a save serializes, in document order.
const AUTHORED_IDS: [(&str, &[&str]); 2] = [
    (
        "word/document.xml",
        &[W14_ID, W14_ID, W14_ID, "0B000002", W14_ID, "0B000003"],
    ),
    ("word/header1.xml", &["0C000001"]),
];

#[test]
fn an_unedited_save_keeps_every_other_part_and_every_authored_paragraph_id() {
    let input = word_docx(|_| {});
    let document = Document::open(&input).unwrap();
    let ids = paragraph_ids(&document);
    for id in &ids {
        assert_eq!(document.paragraph(id).unwrap().para_id.as_ref(), Some(id));
    }

    let saved = document.save().unwrap();
    let (before, after) = (
        ooxml_opc::unzip_parts(&input).unwrap(),
        ooxml_opc::unzip_parts(&saved).unwrap(),
    );
    assert_eq!(
        before.iter().map(|(name, _)| name).collect::<Vec<_>>(),
        after.iter().map(|(name, _)| name).collect::<Vec<_>>()
    );
    for ((name, source), (_, written)) in before.iter().zip(&after) {
        match AUTHORED_IDS.iter().find(|(part, _)| part == name) {
            Some((_, authored)) => assert_eq!(
                written_ids(&String::from_utf8_lossy(written)),
                *authored,
                "{name}"
            ),
            None => assert_eq!(written, source, "{name}"),
        }
    }
    assert!(!ids[1..].contains(&W14_ID.to_owned()));
}

#[test]
fn repeated_ids_stay_addressable_across_save_and_reopen_cycles() {
    let input = word_docx(|_| {});
    let ids = paragraph_ids(&Document::open(&input).unwrap());
    assert_eq!(
        ids.iter().collect::<std::collections::BTreeSet<_>>().len(),
        6
    );

    let first = Document::open(&input).unwrap().save().unwrap();
    let reopened = Document::open(&first).unwrap();
    assert_eq!(paragraph_ids(&reopened), ids);
    assert_eq!(reopened.save().unwrap(), first);

    let mut cycled = input;
    for (index, text) in [(2, "nested edit"), (4, "control edit"), (1, "second edit")] {
        let mut document = Document::open(&cycled).unwrap();
        assert_eq!(paragraph_ids(&document), ids);
        document.replace_paragraph_text(&ids[index], text).unwrap();
        cycled = document.save().unwrap();
    }
    let reopened = Document::open(&cycled).unwrap();
    assert_eq!(
        identified_texts(&reopened),
        [
            (W14_ID, "first"),
            (ids[1].as_str(), "second edit"),
            (ids[2].as_str(), "nested edit"),
            ("0B000002", ""),
            (ids[4].as_str(), "control edit"),
            ("0B000003", "third"),
        ]
        .map(|(id, text)| (id.to_owned(), text.to_owned()))
    );
    assert_eq!(reopened.save().unwrap(), cycled);
    assert_eq!(
        written_ids(&package_part(&cycled, "word/document.xml")),
        ids
    );
}

/// `id` with its last digit written as a character reference.
fn escaped(id: &str) -> String {
    let (head, last) = id.split_at(7);
    format!("{head}&#x{:X};", last.as_bytes()[0])
}

type PartsEdit<'a> = Box<dyn Fn(&mut Vec<(String, String)>) + 'a>;

fn replace_in(parts: &mut [(String, String)], name: &str, from: &str, to: &str) {
    let (_, xml) = parts.iter_mut().find(|(part, _)| part == name).unwrap();
    assert!(xml.contains(from), "{name} lacks {from}");
    *xml = xml.replacen(from, to, 1);
}

#[test]
fn fresh_ids_avoid_escaped_ids_in_every_part() {
    let taken = paragraph_ids(&Document::open(&word_docx(|_| {})).unwrap())[1].clone();
    let reference = escaped(&taken);
    let variants: [(&str, PartsEdit<'_>); 3] = [
        (
            "a later body paragraph",
            Box::new(|parts| replace_in(parts, "word/document.xml", "0B00000&#x33;", &reference)),
        ),
        (
            "a header paragraph",
            Box::new(|parts| replace_in(parts, "word/header1.xml", "0C000001", &reference)),
        ),
        (
            "a comment companion reference",
            Box::new(|parts| {
                parts.push((
                    "word/comments.xml".to_owned(),
                    format!(
                        r#"{XML_DECLARATION}<w:comments {NAMESPACES}><w:comment w:id="0" w:author="Ada">{}</w:comment></w:comments>"#,
                        word_paragraph("0D000001", "note")
                    ),
                ));
                parts.push((
                    "word/commentsExtended.xml".to_owned(),
                    format!(
                        r#"{XML_DECLARATION}<w15:commentsEx {NAMESPACES}><w15:commentEx w15:paraId="0D000001" w15:paraIdParent="{reference}" w15:done="0"/></w15:commentsEx>"#
                    ),
                ));
            }),
        ),
    ];
    for (holder, edit) in &variants {
        let mut document = Document::open(&word_docx(edit)).unwrap();
        let ids = paragraph_ids(&document);
        assert!(!ids[1..5].contains(&taken), "{holder}: {ids:?}");
        assert_eq!(
            ids.iter().collect::<std::collections::BTreeSet<_>>().len(),
            6,
            "{holder}"
        );
        if *holder == "a later body paragraph" {
            assert_eq!(ids[5], taken);
            let mut authored = AUTHORED_IDS[0].1.to_vec();
            authored[5] = &taken;
            assert_eq!(
                written_ids(&package_part(
                    &document.save().unwrap(),
                    "word/document.xml"
                )),
                authored
            );
        }
        document.replace_paragraph_text(&ids[1], "edited").unwrap();
        let reopened = Document::open(&document.save().unwrap()).unwrap();
        assert_eq!(paragraph_ids(&reopened), ids, "{holder}");
        assert_eq!(
            get_paragraph_text(reopened.paragraph(&ids[1]).unwrap()),
            "edited"
        );
    }
}

#[test]
fn a_model_change_saves_the_fresh_ids_it_reads_with() {
    let mut document = Document::open(&word_docx(|_| {})).unwrap();
    let mut ids = paragraph_ids(&document);
    let mut expected = identified_texts(&document);
    document.model_mut().body.content.swap(1, 3);
    ids.swap(1, 4);
    expected.swap(1, 4);

    let saved = document.save().unwrap();
    assert_eq!(written_ids(&package_part(&saved, "word/document.xml")), ids);
    let reopened = Document::open(&saved).unwrap();
    assert_eq!(identified_texts(&reopened), expected);
    assert_eq!(
        get_paragraph_text(reopened.paragraph(&ids[4]).unwrap()),
        "second"
    );
}

#[test]
fn an_editing_session_reads_a_fresh_id_only_once_saved() {
    use docx_edit::{
        AnchorResolution, EditingDoc, ParagraphAnchor, ParagraphRef, SourceStory, SourceStoryKind,
        seed_from_docx,
    };

    let mut document = Document::open(&word_docx(|_| {})).unwrap();
    let ids = paragraph_ids(&document);
    document
        .replace_paragraph_text(&ids[1], "second edit")
        .unwrap();
    let session = EditingDoc::new(7);
    seed_from_docx(&session, &document.save().unwrap()).unwrap();
    let resolve = |para_id: &str| {
        session.resolve_paragraph_anchor(&ParagraphAnchor::Persisted {
            story: SourceStory {
                part_uri: "/word/document.xml".to_owned(),
                kind: SourceStoryKind::Body,
                item_id: None,
            },
            para_id: para_id.to_owned(),
        })
    };

    let AnchorResolution::Found(ParagraphRef::Session { story, para_id }) = resolve(&ids[1]) else {
        panic!("{:?}", resolve(&ids[1]))
    };
    let edited = session
        .paragraphs(&story)
        .unwrap()
        .into_iter()
        .find(|paragraph| paragraph.para_id == para_id)
        .unwrap();
    assert_eq!(edited.text, "second edit");
    assert_eq!(resolve(&ids[4]), AnchorResolution::Missing);
    assert!(matches!(resolve(W14_ID), AnchorResolution::Ambiguous(_)));
}
