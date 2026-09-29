use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const PAGES: usize = 12;
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;

/// A complex field whose cached result is `result`.
fn field_with_result(instruction: &str, result: &str) -> String {
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}<w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

fn field(instruction: &str, cached: &str) -> String {
    field_with_result(instruction, &format!("<w:r><w:t>{cached}</w:t></w:r>"))
}

/// A right-aligned text box anchored at the margin's right edge.
fn text_box(paragraph: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1143000" cy="228600"/><wp:wrapNone/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>{paragraph}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#
    )
}

/// Twelve pages whose footer is a text box holding `Page {PAGE}`, the field
/// cached as "2". The first page has a text box holding a TIME field, a PAGE
/// field in a hyperlink, a REF field whose result holds a PAGE field, and a
/// field whose instruction is only digits, which shows nothing in body text.
fn document() -> Vec<u8> {
    let body = (1..=PAGES)
        .map(|page| {
            let properties = if page > 1 {
                "<w:pPr><w:pageBreakBefore/></w:pPr>"
            } else {
                ""
            };
            let boxed = if page == 1 {
                let page_9 = r#"<w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple>"#;
                text_box(&format!(
                    r#"<w:p><w:r><w:t xml:space="preserve">At </w:t></w:r>{}<w:hyperlink w:anchor="top">{page_9}</w:hyperlink>{}{}</w:p>"#,
                    field("TIME", "10:30"),
                    field_with_result("REF top \\h", page_9),
                    field("12345", "HIDDEN"),
                ))
            } else {
                String::new()
            };
            format!(r#"<w:p>{properties}{boxed}<w:r><w:t>Body page {page}</w:t></w:r></w:p>"#)
        })
        .collect::<String>();
    let footer = format!(
        r#"<w:ftr {NAMESPACES}><w:p>{}</w:p></w:ftr>"#,
        text_box(&format!(
            r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>{}</w:p>"#,
            field("PAGE", "2")
        ))
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
        ("word/footer1.xml", footer),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn display_list(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(75300);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": {},
                "regions": {"sections": [{"properties": package.document.final_section_properties}]},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap()
}

/// The texts of the painted field primitives under `value`.
fn fields(value: &Value) -> Vec<String> {
    fn collect(value: &Value, out: &mut Vec<String>) {
        match value {
            Value::Object(map) if map.contains_key("field") && map.contains_key("text") => {
                out.push(map["text"].as_str().unwrap().to_owned())
            }
            Value::Object(map) => map.values().for_each(|value| collect(value, out)),
            Value::Array(items) => items.iter().for_each(|value| collect(value, out)),
            _ => {}
        }
    }
    let mut out = Vec::new();
    collect(value, &mut out);
    out
}

#[test]
fn fields_in_text_boxes_paint() {
    docx_layout::clear_measure_fonts();
    let display = display_list(&document());
    let footer = |page: usize| fields(&display["pages"][page]["footer"]);
    assert_eq!(footer(0), ["1"]);
    assert_eq!(footer(1), ["2"]);
    assert_eq!(footer(PAGES - 1), ["12"]);
    let body = &display["pages"][0]["primitives"];
    assert_eq!(fields(body), ["10:30", "1", "1", ""]);
    assert!(!body.to_string().contains("HIDDEN"));
}
