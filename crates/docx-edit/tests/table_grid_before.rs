use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditingDoc, seed_from_docx};
use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{BlockExtent, LayoutBlock};
use serde_json::json;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn lower(properties: &str, span: u32) -> Vec<LayoutBlock> {
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="300"/><w:gridCol w:w="1200"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr><w:trPr>{properties}</w:trPr><w:tc><w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>ECM Description Title</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>"#
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", xml),
    ];
    let bytes = ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap();
    let doc = EditingDoc::new(75140);
    seed_from_docx(&doc, &bytes).unwrap();
    let before = doc.encode_state_as_update_v1();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    assert_eq!(doc.encode_state_as_update_v1(), before);
    blocks
}

#[test]
fn bridge_preserves_leading_grid_columns_and_drops_legacy_synthetic_cell_widths() {
    docx_layout::with_private_measure_fonts(|| {
        let mut blocks = lower(r#"<w:gridBefore w:val="1"/>"#, 2);
        let LayoutBlock::Table(table) = &blocks[0] else {
            panic!("table");
        };
        assert_eq!(table.rows[0].grid_before, Some(1));
        assert_eq!(table.rows[0].cells[0].width_value, None);
        let font = docx_layout::register_measure_font(FONT).unwrap();
        let config = MeasurementConfig {
            font_chains: [("arial|0|0".to_owned(), vec![font])].into(),
            defaults: json!({"fontFamily": "Arial", "fontSize": 12}),
            authoritative_shaping: true,
            ..Default::default()
        };
        let measures = measure_blocks(&mut blocks, 220.0, &config).unwrap();
        let BlockExtent::Table(table) = &measures[0] else {
            panic!("table extent");
        };
        assert_eq!(table.column_widths, [20.0, 80.0, 120.0]);
        assert_eq!(table.rows[0].cells[0].width, 200.0);
        let BlockExtent::Paragraph(paragraph) = &table.rows[0].cells[0].blocks[0] else {
            panic!("paragraph extent");
        };
        assert_eq!(paragraph.lines.len(), 1);
        assert!(
            paragraph
                .lines
                .iter()
                .all(|line| line.synthetic_fallback != Some(true))
        );
    });
}

#[test]
fn bridge_preserves_trailing_grid_columns() {
    let blocks = lower(r#"<w:gridAfter w:val="1"/>"#, 2);
    let LayoutBlock::Table(table) = &blocks[0] else {
        panic!("table");
    };
    assert_eq!(table.rows[0].grid_before, None);
    assert_eq!(table.rows[0].grid_after, Some(1));
    assert_eq!(docx_layout::table_grid::count_table_columns(table), 3);
}

#[test]
fn bridge_preserves_both_skips_and_their_preferred_width_units() {
    let blocks = lower(
        concat!(
            r#"<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>"#,
            r#"<w:wBefore w:w="300" w:type="dxa"/><w:wAfter w:w="500" w:type="pct"/>"#,
        ),
        1,
    );
    let LayoutBlock::Table(table) = &blocks[0] else {
        panic!("table");
    };
    let row = &table.rows[0];
    assert_eq!((row.grid_before, row.grid_after), (Some(1), Some(1)));
    let before = row.width_before.as_ref().unwrap();
    let after = row.width_after.as_ref().unwrap();
    assert_eq!(
        (before.value, before.r#type.as_deref()),
        (Some(300.0), Some("dxa"))
    );
    assert_eq!(
        (after.value, after.r#type.as_deref()),
        (Some(500.0), Some("pct"))
    );
}

#[test]
fn bridge_leaves_unskipped_rows_without_new_layout_properties() {
    for properties in [
        "",
        r#"<w:wBefore w:w="300" w:type="dxa"/><w:wAfter w:w="500" w:type="pct"/>"#,
    ] {
        let blocks = lower(properties, 3);
        let LayoutBlock::Table(table) = &blocks[0] else {
            panic!("table");
        };
        let row = &table.rows[0];
        assert_eq!((row.grid_before, row.grid_after), (None, None));
        assert_eq!((&row.width_before, &row.width_after), (&None, &None));
        assert_eq!(
            table.column_widths.as_deref(),
            Some([20.0, 80.0, 120.0].as_slice())
        );
    }
}
