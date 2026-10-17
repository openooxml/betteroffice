use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// A 2x2 table whose outer edges carry `outer` borders and inner edges `inner`.
fn document(outer: &str, inner: &str) -> Vec<u8> {
    let edge = |side: &str, value: &str| {
        format!(r#"<w:{side} w:val="{value}" w:sz="4" w:space="0" w:color="000000"/>"#)
    };
    let borders = ["top", "left", "bottom", "right"]
        .map(|side| edge(side, outer))
        .concat()
        + &["insideH", "insideV"]
            .map(|side| edge(side, inner))
            .concat();
    let cell = r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>"#;
    let row = format!("<w:tr>{cell}{cell}</w:tr>");
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>{row}{row}</w:tbl><w:p/>"#
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn layout(bytes: &[u8], render_env: Value) -> (String, Value) {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75400);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "options": {}, "renderEnv": render_env,
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    let display = serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    (output, display)
}

fn gridlines(display: &Value) -> Vec<[f64; 4]> {
    display["pages"][0]["tableGridlines"]
        .as_array()
        .map(|lines| {
            lines
                .iter()
                .map(|line| ["x1", "y1", "x2", "y2"].map(|key| line[key].as_f64().unwrap()))
                .collect()
        })
        .unwrap_or_default()
}

fn table_borders(display: &Value) -> usize {
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|primitive| primitive["role"] == "table-border")
        .count()
}

#[test]
fn a_borderless_table_reports_every_cell_edge_as_a_gridline() {
    let (_, display) = layout(&document("nil", "nil"), json!({"showTableGridlines": true}));
    let lines = gridlines(&display);
    assert_eq!(lines.len(), 12);
    assert_eq!(table_borders(&display), 0);
    let mut columns: Vec<f64> = lines
        .iter()
        .filter(|[x1, _, x2, _]| x1 == x2)
        .map(|[x, ..]| *x)
        .collect();
    columns.sort_by(f64::total_cmp);
    columns.dedup();
    let mut rows: Vec<f64> = lines
        .iter()
        .filter(|[_, y1, _, y2]| y1 == y2)
        .map(|[_, y, ..]| *y)
        .collect();
    rows.sort_by(f64::total_cmp);
    rows.dedup();
    assert_eq!((columns.len(), rows.len()), (3, 3));
}

#[test]
fn outside_borders_leave_only_the_inner_edges_as_gridlines() {
    let (_, display) = layout(
        &document("single", "nil"),
        json!({"showTableGridlines": true}),
    );
    assert_eq!(gridlines(&display).len(), 4);
    assert_eq!(table_borders(&display), 8);
}

#[test]
fn gridlines_are_off_by_default_and_change_nothing_but_the_guides() {
    for bytes in [document("nil", "nil"), document("single", "nil")] {
        let (layout_off, off) = layout(&bytes, json!({}));
        let (layout_false, off_explicit) = layout(&bytes, json!({"showTableGridlines": false}));
        assert_eq!(layout_off, layout_false);
        assert_eq!(off, off_explicit);
        assert!(off["pages"][0].get("tableGridlines").is_none());
        let (_, mut on) = layout(&bytes, json!({"showTableGridlines": true}));
        assert!(!gridlines(&on).is_empty());
        on["pages"][0]
            .as_object_mut()
            .unwrap()
            .remove("tableGridlines");
        assert_eq!(on, off);
    }
}
