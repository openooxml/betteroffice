use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const PAGES: usize = 12;

/// Twelve pages whose footer table ends in a right-aligned `Page {PAGE}` cell,
/// the field cached as "2".
fn document() -> Vec<u8> {
    let body = (1..=PAGES)
        .map(|page| {
            let properties = if page > 1 {
                "<w:pPr><w:pageBreakBefore/></w:pPr>"
            } else {
                ""
            };
            format!(r#"<w:p>{properties}<w:r><w:t>Body page {page}</w:t></w:r></w:p>"#)
        })
        .collect::<String>();
    let cell = |content: &str| {
        format!(r#"<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr>{content}</w:tc>"#)
    };
    let footer = format!(
        r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl><w:p/></w:ftr>"#,
        cell("<w:p><w:r><w:t>Footer</w:t></w:r></w:p>"),
        cell(
            r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#
        ),
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
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

/// The painted text of the page's footer field and its left and right edges.
fn footer_field(display: &Value, page: usize) -> (String, f64, f64) {
    fn find(value: &Value) -> Option<&Value> {
        match value {
            Value::Object(map) if map.contains_key("field") && map.contains_key("width") => {
                Some(value)
            }
            Value::Object(map) => map.values().find_map(find),
            Value::Array(items) => items.iter().find_map(find),
            _ => None,
        }
    }
    let field = find(&display["pages"][page]["footer"]).expect("footer field paints");
    let x = field["x"].as_f64().unwrap();
    let width = field["width"].as_f64().unwrap();
    (field["text"].as_str().unwrap().to_owned(), x, x + width)
}

#[test]
fn a_right_aligned_page_field_in_a_footer_table_stays_right_aligned() {
    docx_layout::clear_measure_fonts();
    let display = display_list(&document());
    assert_eq!(display["pages"].as_array().unwrap().len(), PAGES);
    let (two, two_left, two_right) = footer_field(&display, 1);
    let (twelve, twelve_left, twelve_right) = footer_field(&display, PAGES - 1);
    assert_eq!((two.as_str(), twelve.as_str()), ("2", "12"));
    // Word ends "Page 2" and "Page 12" at the same right edge, the cell's.
    assert!(
        (two_right - twelve_right).abs() < 0.01,
        "{two_right} vs {twelve_right}"
    );
    assert!(twelve_left < two_left - 3.0, "{twelve_left} vs {two_left}");
}
