use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const PAGES: usize = 12;

fn field(kind: &str, cached: &str, run_properties: &str) -> String {
    let run = |content: &str| format!("<w:r>{run_properties}{content}</w:r>");
    [
        run(r#"<w:fldChar w:fldCharType="begin"/>"#),
        run(&format!(
            r#"<w:instrText xml:space="preserve"> {kind} </w:instrText>"#
        )),
        run(r#"<w:fldChar w:fldCharType="separate"/>"#),
        run(&format!("<w:t>{cached}</w:t>")),
        run(r#"<w:fldChar w:fldCharType="end"/>"#),
    ]
    .concat()
}

/// A one-row table of two `column`-twip cells.
fn table(column: u32, cells: [&str; 2]) -> String {
    let cells = cells
        .iter()
        .map(|content| {
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="{column}" w:type="dxa"/></w:tcPr>{content}</w:tc>"#
            )
        })
        .collect::<String>();
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="{}" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="{column}"/><w:gridCol w:w="{column}"/></w:tblGrid><w:tr>{cells}</w:tr></w:tbl>"#,
        2 * column
    )
}

/// Twelve pages whose footer table ends in a right-aligned `paragraph` cell.
fn document(paragraph: &str) -> Vec<u8> {
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
    let footer = format!(
        r#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}<w:p/></w:ftr>"#,
        table(4513, ["<w:p><w:r><w:t>Footer</w:t></w:r></w:p>", paragraph]),
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

struct Painted {
    text: String,
    left: f64,
    right: f64,
    clip_right: f64,
}

/// The page's painted footer fields in order.
fn footer_fields(display: &Value, page: usize) -> Vec<Painted> {
    fn collect<'a>(value: &'a Value, out: &mut Vec<&'a Value>) {
        match value {
            Value::Object(map) if map.contains_key("field") && map.contains_key("width") => {
                out.push(value)
            }
            Value::Object(map) => map.values().for_each(|value| collect(value, out)),
            Value::Array(items) => items.iter().for_each(|value| collect(value, out)),
            _ => {}
        }
    }
    let mut fields = Vec::new();
    collect(&display["pages"][page]["footer"], &mut fields);
    fields
        .into_iter()
        .map(|field| {
            let left = field["x"].as_f64().unwrap();
            let clip = &field["clipGroup"]["clip"];
            Painted {
                text: field["text"].as_str().unwrap().to_owned(),
                left,
                right: left + field["width"].as_f64().unwrap(),
                clip_right: clip["x"].as_f64().unwrap() + clip["w"].as_f64().unwrap(),
            }
        })
        .collect()
}

const SUPERSCRIPT: &str = r#"<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>"#;

/// A right-aligned `Page {PAGE}` paragraph, the field cached as "2".
fn page_paragraph(field_properties: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>{}</w:p>"#,
        field("PAGE", "2", field_properties)
    )
}

/// The page's one painted footer field.
fn footer_field(display: &Value, page: usize) -> Painted {
    let mut fields = footer_fields(display, page);
    assert_eq!(fields.len(), 1);
    fields.remove(0)
}

fn assert_close(a: f64, b: f64) {
    assert!((a - b).abs() < 0.01, "{a} vs {b}");
}

#[test]
fn a_right_aligned_page_field_in_a_footer_table_stays_right_aligned() {
    docx_layout::clear_measure_fonts();
    let display = display_list(&document(&page_paragraph("")));
    assert_eq!(display["pages"].as_array().unwrap().len(), PAGES);
    let two = footer_field(&display, 1);
    let twelve = footer_field(&display, PAGES - 1);
    assert_eq!((two.text.as_str(), twelve.text.as_str()), ("2", "12"));
    // Word ends "Page 2" and "Page 12" at the same right edge, the cell's.
    assert_close(two.right, twelve.right);
    assert_close(twelve.right, twelve.clip_right);
    assert!(
        twelve.left < two.left - 3.0,
        "{} vs {}",
        twelve.left,
        two.left
    );
}

#[test]
fn a_superscript_page_field_in_a_footer_table_keeps_its_script_size() {
    docx_layout::clear_measure_fonts();
    let plain = display_list(&document(&page_paragraph("")));
    let superscript = display_list(&document(&page_paragraph(SUPERSCRIPT)));
    for page in [1, PAGES - 1] {
        let plain = footer_field(&plain, page);
        let superscript = footer_field(&superscript, page);
        assert_close(superscript.right, plain.right);
        assert_close(
            (superscript.right - superscript.left) / (plain.right - plain.left),
            0.75,
        );
    }
}

#[test]
fn page_and_numpages_fields_in_a_nested_footer_table_stay_right_aligned() {
    docx_layout::clear_measure_fonts();
    let paragraph = format!(
        r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr>{}<w:r><w:t xml:space="preserve"> of </w:t></w:r>{}</w:p>"#,
        field("PAGE", "2", ""),
        field("NUMPAGES", "9", "")
    );
    let display = display_list(&document(&(table(2000, ["<w:p/>", &paragraph]) + "<w:p/>")));
    let [two, total_two] = footer_fields(&display, 1).try_into().ok().unwrap();
    let [twelve, total_twelve] = footer_fields(&display, PAGES - 1).try_into().ok().unwrap();
    assert_eq!(
        [&two.text, &total_two.text, &twelve.text, &total_twelve.text],
        ["2", "12", "12", "12"]
    );
    assert_close(two.right, twelve.right);
    assert!(
        twelve.left < two.left - 3.0,
        "{} vs {}",
        twelve.left,
        two.left
    );
    // Every field shows "12" in the same formatting, so each is as wide as PAGE's.
    let width = |field: &Painted| field.right - field.left;
    assert_close(width(&total_two), width(&twelve));
    assert_close(width(&total_twelve), width(&twelve));
    assert_close(total_two.right, total_two.clip_right);
    assert_close(total_twelve.right, total_twelve.clip_right);
}
