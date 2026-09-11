use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const SHAPE: &str = r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="66675" distR="123825" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>inside</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Inside shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#;

fn document(body: &str, styles: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned()),
        ("word/styles.xml", format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{styles}</w:styles>"#)),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>{body}</w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn paragraph(text: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="480" w:lineRule="exact"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn layout(body: &str, columns: u32) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(74201);
    seed_from_docx(engine.doc(), &document(body, "")).unwrap();
    serde_json::from_str(&engine.layout_document_with_regions_json(&json!({
        "bodyStory": "body", "renderEnv": {},
        "options": {"pageSize": {"w":816,"h":1056},
            "margins":{"top":96,"bottom":96,"left":96,"right":96},
            "columns":{"count":columns,"gap":48}},
        "measurement":{"fontChains":{"calibri|0|0":[font]},"defaults":{"fontFamily":"Calibri","fontSize":12}}
    }).to_string()).unwrap()).unwrap()
}

#[test]
fn break_only_paragraphs_match_word_page_and_column_flow() {
    for (kinds, columns, page_count, x, y) in [
        (vec!["column"], 2, 1, 432.0, 160.0),
        (vec!["page"], 1, 2, 96.0, 96.0),
        (vec!["column", "column"], 2, 2, 96.0, 160.0),
        (vec!["page", "column"], 2, 2, 432.0, 160.0),
        (vec!["column", "page"], 2, 2, 96.0, 96.0),
    ] {
        let breaks = kinds
            .iter()
            .map(|kind| format!(r#"<w:br w:type="{kind}"/>"#))
            .collect::<String>();
        let body = format!(
            r#"{}<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="960" w:lineRule="exact"/></w:pPr><w:r>{breaks}</w:r></w:p>{}"#,
            paragraph("Before"),
            paragraph("After")
        );
        let output = layout(&body, columns);
        let pages = output["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), page_count, "{kinds:?}");
        let after = pages.last().unwrap()["fragments"]
            .as_array()
            .unwrap()
            .last()
            .unwrap();
        assert_eq!(after["y"].as_f64().unwrap(), y, "{kinds:?}");
        assert_eq!(after["x"].as_f64().unwrap(), x, "{kinds:?}");
    }
}

#[test]
fn inside_shape_wrap_matches_actual_page_and_asymmetric_distances() {
    for (prefix, expected_page) in [
        (String::new(), 1),
        (
            format!(
                r#"{}<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#,
                paragraph("First page")
            ),
            2,
        ),
        (paragraph("Automatic pagination").repeat(30), 2),
    ] {
        let body = format!(
            "{prefix}<w:p>{SHAPE}<w:r><w:t>{}</w:t></w:r></w:p>",
            "Body text wraps around the shape. ".repeat(20)
        );
        let output = layout(&body, 1);
        let pages = output["layout"]["pages"].as_array().unwrap();
        let page = &pages[expected_page - 1];
        let fragments = page["fragments"].as_array().unwrap();
        let shape = fragments
            .iter()
            .find(|fragment| fragment["kind"] == "shape")
            .unwrap();
        let text = fragments
            .iter()
            .find(|fragment| {
                fragment["kind"] == "paragraph" && fragment["pmStart"] == shape["pmStart"]
            })
            .unwrap();
        let measure = output["measured"]
            .as_array()
            .unwrap()
            .iter()
            .find(|measured| measured["block"]["id"] == text["blockId"])
            .unwrap();
        let line = &measure["measure"]["lines"][0];
        let start = text["x"].as_f64().unwrap() + line["leftOffset"].as_f64().unwrap_or(0.0);
        let end = start + line["width"].as_f64().unwrap();
        let shape_x = shape["x"].as_f64().unwrap();
        if expected_page == 1 {
            assert_eq!(shape_x, 96.0);
            assert!(start >= shape_x + 192.0 + 13.0);
        } else {
            assert_eq!(shape_x, 528.0);
            assert_eq!(start, 96.0);
            assert!(
                end <= shape_x - 7.0,
                "text ends at {end}, shape starts at {shape_x}"
            );
        }
    }
}

#[test]
fn dark_cells_preserve_colors_inherited_from_styles_and_defaults() {
    let styles = r#"<w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="FF0000"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Green"><w:rPr><w:color w:val="00FF00"/></w:rPr></w:style><w:style w:type="character" w:styleId="Blue"><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>"#;
    for (ppr, rpr, expected) in [
        ("", "", "#FF0000"),
        (r#"<w:pStyle w:val="Green"/>"#, "", "#00FF00"),
        ("", r#"<w:rStyle w:val="Blue"/>"#, "#0000FF"),
    ] {
        let body = format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="222222"/></w:tcPr><w:p><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>Inherited color</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#
        );
        let engine = EngineSession::new(74202);
        seed_from_docx(engine.doc(), &document(&body, styles)).unwrap();
        let blocks = engine
            .with_lowered_story("body", &docx_edit::bridge::RenderEnv::default(), |blocks| {
                serde_json::to_value(blocks).unwrap()
            })
            .unwrap();
        assert_eq!(
            blocks[0]["rows"][0]["cells"][0]["blocks"][0]["runs"][0]["color"],
            expected
        );
    }
}
