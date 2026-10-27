use docx_edit::structured::{
    Block, BlockKind, ExportOptions, Inline, InlineKind, RevisionView, StorySelection,
};
use docx_edit::{
    DrawingProperties, EditTextView, EditingDoc, ReadParagraphsRequest, ReadStoriesRequest,
    StoryText, TextAtom, seed_from_docx,
};
use docx_parse::image::{DrawingPlacement, DrawingType};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" "#,
    r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" "#,
    r#"xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" "#,
    r#"xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
    r#"xmlns:adec="http://schemas.microsoft.com/office/drawing/2015/06/main""#,
);

fn picture(container: &str, doc_pr: &str) -> String {
    format!(
        r#"<w:drawing><wp:{container}><wp:extent cx="914400" cy="457200"/>{doc_pr}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="inner"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:{container}></w:drawing>"#
    )
}

fn inline_picture(id: &str, extra: &str) -> String {
    picture(
        "inline",
        &format!(r#"<wp:docPr id="{id}" {extra}/>"#),
    )
}

fn anchored_picture(id: &str, extra: &str, wrap: &str, behind: &str) -> String {
    format!(
        r#"<w:drawing><wp:anchor behindDoc="{behind}" relativeHeight="1"><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrap{wrap}/><wp:docPr id="{id}" {extra}/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>"#
    )
}

fn paragraph(id: &str, run_content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}"><w:r>{run_content}</w:r></w:p>"#)
}

fn chart_part() -> String {
    concat!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>"#,
        r#"<c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:tx><c:v>Sales</c:v></c:tx>"#,
        r#"<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:cat>"#,
        r#"<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"#,
        r#"<c:axId val="1"/><c:axId val="2"/></c:barChart>"#,
        r#"<c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>"#,
        r#"<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>"#,
        r#"</c:plotArea></c:chart></c:chartSpace>"#,
    )
    .to_owned()
}

fn fixture() -> Vec<u8> {
    let inline = paragraph(
        "10000001",
        &inline_picture("1", r#"name="N-inline" title="T-inline" descr="D-inline""#),
    );
    let anchored = paragraph(
        "10000002",
        &anchored_picture("2", r#"name="N-anchor" title="T-anchor" descr="D-anchor""#, "Square", "0"),
    );
    let behind = paragraph(
        "10000003",
        &anchored_picture("3", r#"name="N-behind""#, "None", "1"),
    );
    let cell = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        paragraph(
            "10000004",
            &inline_picture("4", r#"name="N-cell""#),
        )
    );
    let text_box = paragraph("10000006", concat!(
        r#"<w:drawing><wp:anchor relativeHeight="2"><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>"#,
        r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapSquare/>"#,
        r#"<wp:docPr id="6" name="N-textbox" title="T-textbox" descr="D-textbox"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">"#,
        r#"<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>"#,
        r#"<wps:txbx><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>"#,
        r#"</a:graphicData></a:graphic></wp:anchor></w:drawing>"#,
    ));
    let shape = paragraph("10000007", concat!(
        r#"<w:drawing><wp:anchor relativeHeight="3"><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>"#,
        r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapSquare/>"#,
        r#"<wp:docPr id="7" name="N-shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">"#,
        r#"<wps:wsp><wps:cNvPr id="7" name="inner-shape"/><wps:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>"#,
        r#"</a:graphicData></a:graphic></wp:anchor></w:drawing>"#,
    ));
    let group = paragraph("10000008", concat!(
        r#"<w:drawing><wp:inline><wp:extent cx="2000" cy="1000"/><wp:docPr id="8" name="N-group"/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp>"#,
        r#"<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm></wpg:grpSpPr>"#,
        r#"<wps:wsp><wps:cNvPr id="81" name="child"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr></wps:wsp>"#,
        r#"</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
    ));
    let chart = paragraph("10000009", concat!(
        r#"<w:drawing><wp:inline><wp:extent cx="2743200" cy="1828800"/><wp:docPr id="9" name="N-chart" descr="D-chart"/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
    ));
    let hidden = paragraph(
        "1000000A",
        &inline_picture("10", r#"name="N-hidden" hidden="1""#),
    );
    let decorative = paragraph("1000000B", &picture("inline", concat!(
        r#"<wp:docPr id="11" name="N-decorative"><a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">"#,
        r#"<adec:decorative val="1"/></a:ext></a:extLst></wp:docPr>"#,
    )));
    let empty_name = paragraph(
        "1000000C",
        &inline_picture("12", r#"name="" title="T-empty""#),
    );
    let duplicate_a = paragraph(
        "1000000D",
        &inline_picture("99", r#"name="N-dup-a""#),
    );
    let duplicate_b = paragraph(
        "1000000E",
        &inline_picture("99", r#"name="N-dup-b""#),
    );
    let body = format!(
        "{inline}{anchored}{behind}{cell}{text_box}{shape}{group}{chart}{hidden}{decorative}{empty_name}{duplicate_a}{duplicate_b}<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1120\" w:right=\"900\" w:bottom=\"1440\" w:left=\"900\"/></w:sectPr>"
    );
    let header = format!(
        "<w:hdr {NS}>{}</w:hdr>",
        paragraph(
            "20000001",
            &inline_picture("5", r#"name="N-header" descr="D-header""#),
        )
    );
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            concat!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
                r#"<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>"#,
                r#"<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#,
                r#"<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>"#,
                r#"<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
                r#"</Types>"#,
            )
            .as_bytes()
            .to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="{REL}/image" Target="media/image.png"/><Relationship Id="rIdHeader" Type="{REL}/header" Target="header1.xml"/><Relationship Id="rIdChart" Type="{REL}/chart" Target="charts/chart1.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "word/_rels/header1.xml.rels".to_owned(),
            format!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="{REL}/image" Target="media/image.png"/></Relationships>"#
            )
            .into_bytes(),
        ),
        (
            "word/document.xml".to_owned(),
            format!("<w:document {NS}><w:body>{body}</w:body></w:document>").into_bytes(),
        ),
        ("word/header1.xml".to_owned(), header.into_bytes()),
        ("word/charts/chart1.xml".to_owned(), chart_part().into_bytes()),
        (
            "word/media/image.png".to_owned(),
            b"\x89PNG\r\n\x1a\n".to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(1837);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn drawings(atoms: &[TextAtom]) -> Vec<&DrawingProperties> {
    atoms.iter().filter_map(|atom| atom.drawing.as_ref()).collect()
}

fn by_name<'a>(
    atoms: &'a [TextAtom],
    name: &str,
) -> &'a DrawingProperties {
    drawings(atoms)
        .into_iter()
        .find(|drawing| drawing.name.as_deref() == Some(name))
        .unwrap_or_else(|| panic!("no drawing named {name}"))
}

#[test]
fn read_stories_carry_doc_pr_properties() {
    let doc = open(&fixture());
    let stories = doc
        .read_stories(&ReadStoriesRequest {
            stories: None,
            by_root: false,
            view: EditTextView::Accepted,
            expect_version: None,
        })
        .unwrap()
        .stories;
    let mut atoms: Vec<TextAtom> = Vec::new();
    let mut header_atoms: Vec<TextAtom> = Vec::new();
    for story in &stories {
        let StoryText::Read { story, paragraphs } = story else {
            panic!("story refused: {story:?}");
        };
        for paragraph in paragraphs {
            if story.starts_with("hf:") {
                header_atoms.extend(paragraph.atoms.iter().cloned());
            }
            atoms.extend(paragraph.atoms.iter().cloned());
        }
    }
    let inline = by_name(&atoms, "N-inline");
    assert_eq!(inline.doc_pr_id.as_deref(), Some("1"));
    assert_eq!(inline.title.as_deref(), Some("T-inline"));
    assert_eq!(inline.descr.as_deref(), Some("D-inline"));
    assert_eq!(inline.placement, DrawingPlacement::Inline);
    assert_eq!(inline.drawing_type, DrawingType::Picture);
    assert!(!inline.hidden && !inline.decorative);

    let anchored = by_name(&atoms, "N-anchor");
    assert_eq!(anchored.placement, DrawingPlacement::Anchor);
    assert_eq!(anchored.descr.as_deref(), Some("D-anchor"));

    assert_eq!(by_name(&atoms, "N-behind").placement, DrawingPlacement::Anchor);
    assert_eq!(by_name(&atoms, "N-cell").doc_pr_id.as_deref(), Some("4"));
    assert_eq!(by_name(&header_atoms, "N-header").descr.as_deref(), Some("D-header"));

    let text_box = by_name(&atoms, "N-textbox");
    assert_eq!(text_box.drawing_type, DrawingType::TextBox);
    assert_eq!(text_box.placement, DrawingPlacement::Anchor);
    assert_eq!(by_name(&atoms, "N-shape").drawing_type, DrawingType::Shape);
    assert_eq!(by_name(&atoms, "N-group").drawing_type, DrawingType::Group);
    let chart = by_name(&atoms, "N-chart");
    assert_eq!(chart.drawing_type, DrawingType::Chart);
    assert_eq!(chart.descr.as_deref(), Some("D-chart"));

    assert!(by_name(&atoms, "N-hidden").hidden);
    assert!(by_name(&atoms, "N-decorative").decorative);

    let empty = atoms
        .iter()
        .filter_map(|atom| atom.drawing.as_ref())
        .find(|drawing| drawing.doc_pr_id.as_deref() == Some("12"))
        .unwrap();
    assert_eq!(empty.name, None);
    assert_eq!(empty.title.as_deref(), Some("T-empty"));

    assert_eq!(by_name(&atoms, "N-dup-a").doc_pr_id.as_deref(), Some("99"));
    assert_eq!(by_name(&atoms, "N-dup-b").doc_pr_id.as_deref(), Some("99"));

    let body = doc
        .read_paragraphs(&ReadParagraphsRequest {
            story: Some("body".to_owned()),
            para_ids: Some(vec!["10000001".to_owned()]),
            view: EditTextView::Original,
        })
        .unwrap()
        .paragraphs;
    assert_eq!(drawings(&body[0].atoms).len(), 1);
    assert_eq!(body[0].atoms[0].offset, 0);
}

fn collect_inlines(blocks: &[Block], out: &mut Vec<Inline>) {
    for block in blocks {
        match &block.content {
            BlockKind::Paragraph { paragraph }
            | BlockKind::Heading { paragraph, .. }
            | BlockKind::ListItem { paragraph, .. } => out.extend(paragraph.inlines.iter().cloned()),
            BlockKind::Table { table } => {
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_inlines(&cell.blocks, out);
                    }
                }
            }
            BlockKind::ContentControl { blocks, .. } => collect_inlines(blocks, out),
            _ => {}
        }
    }
}

#[test]
fn structured_export_carries_doc_pr_properties() {
    let bytes = fixture();
    let content = docx_edit::structured::export_docx_structured(
        &bytes,
        &ExportOptions {
            stories: Some(vec![StorySelection::Body, StorySelection::Headers]),
            ..ExportOptions::new(RevisionView::Accepted)
        },
    )
    .unwrap();
    let mut inlines = Vec::new();
    for story in &content.stories {
        collect_inlines(&story.blocks, &mut inlines);
    }
    let image = inlines
        .iter()
        .find_map(|inline| match &inline.content {
            InlineKind::Image { drawing, alt_text, .. } if drawing.as_ref().is_some_and(|d| d.name.as_deref() == Some("N-inline")) => {
                Some((drawing.clone(), alt_text.clone()))
            }
            _ => None,
        })
        .expect("inline picture image inline");
    assert_eq!(image.0.as_ref().unwrap().descr.as_deref(), Some("D-inline"));
    assert_eq!(image.1.as_deref(), Some("D-inline"));

    for (name, element) in [
        ("N-shape", "wps:wsp"),
        ("N-textbox", "wps:txbx"),
        ("N-chart", "c:chart"),
    ] {
        let found = inlines.iter().any(|inline| match &inline.content {
            InlineKind::Unsupported { element: found, drawing, .. } => {
                found == element
                    && drawing.as_ref().is_some_and(|d| d.name.as_deref() == Some(name))
            }
            InlineKind::Image { drawing, .. } => {
                drawing.as_ref().is_some_and(|d| d.name.as_deref() == Some(name))
            }
            _ => false,
        });
        assert!(found, "structured export carries {name} as {element}");
    }
    assert!(inlines.iter().any(|inline| matches!(
        &inline.content,
        InlineKind::Image { drawing: Some(drawing), .. } if drawing.name.as_deref() == Some("N-header")
    )));
}

#[test]
fn save_round_trip_keeps_document_xml_unchanged() {
    let bytes = fixture();
    let _doc = open(&bytes);
    let wire = docx_parse::parse_docx_s9_wire(&bytes, Default::default()).unwrap();
    let projection = docx_parse::parse_docx_s8_projection(&bytes).unwrap();
    let body = serde_json::to_value(projection.body).unwrap();
    let request: docx_parse::serializer::S13SaveRequest = serde_json::from_value(serde_json::json!({
        "determinism": {
            "seed": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
            "now": "2030-01-02T03:04:05.006Z"
        },
        "document": body,
        "relationshipEntries": wire.document.package.relationship_entries,
        "headerEntries": wire.document.package.header_entries,
        "options": {"updateModifiedDate": false},
        "selective": serde_json::Value::Null
    }))
    .unwrap();
    let saved = docx_parse::serializer::write_docx_s13(request, &bytes).unwrap();
    let original = ooxml_opc::unzip_parts(&bytes).unwrap();
    let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let part = |parts: &[(String, Vec<u8>)], name: &str| {
        parts.iter().find(|(path, _)| path == name).unwrap().1.clone()
    };
    assert_eq!(
        part(&saved_parts, "word/document.xml"),
        part(&original, "word/document.xml")
    );
}
