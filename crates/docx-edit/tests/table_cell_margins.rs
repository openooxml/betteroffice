use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditCtx, EditingDoc, EngineSession, Position, seed_from_docx};
use docx_layout::types::LayoutBlock;
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const ROWS: [[&str; 4]; 3] = [
    ["Item width:", "9,1", "cm", "( reference A )"],
    ["Item height:", "31,3", "cm", "( reference B )"],
    ["Item weight:", "2614", "g", "(± 10%)"],
];

/// The 4-column table, its third column 24 px (360 twips) wide.
fn document(table_margins: &str, cell_margins: &str) -> Vec<u8> {
    let grid = [2400, 1200, 360, 2400];
    let rows: String = ROWS
        .iter()
        .map(|cells| {
            let cells: String = cells
                .iter()
                .zip(grid)
                .map(|(text, width)| {
                    format!(
                        r#"<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{cell_margins}</w:tcPr><w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>"#
                    )
                })
                .collect();
            format!("<w:tr>{cells}</w:tr>")
        })
        .collect();
    let grid_cols: String = grid
        .map(|width| format!(r#"<w:gridCol w:w="{width}"/>"#))
        .concat();
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="6360" w:type="dxa"/><w:tblLayout w:type="fixed"/>{table_margins}<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid>{grid_cols}</w:tblGrid>{rows}</w:tbl><w:p/>"#
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

fn margins(element: &str, left: u32, right: u32) -> String {
    format!(
        r#"<w:{element}><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:{element}>"#
    )
}

fn display(bytes: &[u8]) -> Value {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75401);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "options": {}, "renderEnv": {},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap()
}

fn primitives(display: &Value) -> &Vec<Value> {
    display["pages"][0]["primitives"].as_array().unwrap()
}

fn text_lines(display: &Value) -> Vec<(String, f64, f64)> {
    primitives(display)
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
        .filter_map(|primitive| {
            let text = primitive["text"].as_str()?.to_owned();
            let (x, y) = match primitive["kind"].as_str() {
                Some("text") => (primitive["x"].as_f64()?, primitive["baselineY"].as_f64()?),
                _ => {
                    let glyph = &primitive["glyphs"][0];
                    (glyph["x"].as_f64()?, glyph["y"].as_f64()?)
                }
            };
            Some((text, x, y))
        })
        .collect()
}

fn border_extent(display: &Value, keys: [&str; 3], pick: fn(f64, f64) -> f64) -> f64 {
    primitives(display)
        .iter()
        .filter(|primitive| primitive["role"] == "table-border")
        .flat_map(|primitive| keys.iter().filter_map(|key| primitive[*key].as_f64()))
        .reduce(pick)
        .unwrap()
}

fn table_bottom(display: &Value) -> f64 {
    border_extent(display, ["y", "y1", "y2"], f64::max)
}

fn table_left(display: &Value) -> f64 {
    border_extent(display, ["x", "x1", "x2"], f64::min)
}

/// Text lines painted in the grid column spanning `from..to` px from the table's left edge.
fn column_lines(display: &Value, from: f64, to: f64) -> usize {
    let left = table_left(display);
    let mut baselines: Vec<f64> = text_lines(display)
        .into_iter()
        .filter(|(text, x, _)| !text.trim().is_empty() && *x >= left + from && *x < left + to)
        .map(|(_, _, y)| y)
        .collect();
    baselines.sort_by(f64::total_cmp);
    baselines.dedup_by(|a, b| (*a - *b).abs() < 0.5);
    baselines.len()
}

/// Where the first cell's text starts, measured from the table's left edge.
fn text_inset(display: &Value) -> f64 {
    let first = text_lines(display)
        .into_iter()
        .filter(|(text, ..)| !text.trim().is_empty())
        .map(|(_, x, _)| x)
        .reduce(f64::min)
        .unwrap();
    first - table_left(display)
}

#[test]
fn a_table_without_cell_margins_lays_out_with_words_default_108_twips() {
    let unset = display(&document("", ""));
    let word_default = display(&document(&margins("tblCellMar", 108, 108), ""));
    assert!(!text_lines(&unset).is_empty());
    assert_eq!(text_lines(&unset), text_lines(&word_default));
    assert_eq!(table_bottom(&unset), table_bottom(&word_default));
}

#[test]
fn the_24px_column_wraps_once_the_cell_margins_apply() {
    let unset = display(&document("", ""));
    let flush = display(&document(&margins("tblCellMar", 0, 0), ""));
    assert_eq!(column_lines(&flush, 240.0, 264.0), 3);
    assert_eq!(column_lines(&unset, 240.0, 264.0), 5);
    assert!(table_bottom(&unset) > table_bottom(&flush) + 10.0);
}

#[test]
fn text_starts_at_the_table_margin_and_a_cell_margin_overrides_it() {
    let flush = display(&document(&margins("tblCellMar", 0, 0), ""));
    let table = display(&document(&margins("tblCellMar", 288, 0), ""));
    let cell = display(&document(
        &margins("tblCellMar", 288, 0),
        &margins("tcMar", 0, 0),
    ));
    let unset = display(&document("", ""));
    assert!(text_inset(&flush).abs() < 0.01);
    assert!((text_inset(&table) - 19.2).abs() < 0.01);
    assert!(text_inset(&cell).abs() < 0.01);
    assert!((text_inset(&unset) - 7.2).abs() < 0.01);
}

#[test]
fn an_inserted_table_pads_its_cells_with_words_default_margins() {
    let doc = EditingDoc::new(301);
    doc.create_story("body", "ab", "Normal", "left").unwrap();
    doc.insert_table(
        &EditCtx::local("Owner", "2026-11-06T12:00:00Z"),
        Position::new("body", 2),
        1,
        2,
    )
    .unwrap();
    let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
    let table = blocks
        .iter()
        .find_map(|block| match block {
            LayoutBlock::Table(table) => Some(table),
            _ => None,
        })
        .unwrap();
    for cell in &table.rows[0].cells {
        let padding = cell.padding.as_ref().unwrap();
        assert!((padding.left - 7.2).abs() < 1e-9);
        assert!((padding.right - 7.2).abs() < 1e-9);
        assert_eq!((padding.top, padding.bottom), (0.0, 0.0));
    }
    for margin in [table.cell_margin_left, table.cell_margin_right] {
        assert!((margin.unwrap() - 7.2).abs() < 1e-9);
    }
}

#[test]
fn an_inserted_table_follows_the_document_compatibility_mode() {
    let doc = EditingDoc::new(302);
    doc.create_story("body", "ab", "Normal", "left").unwrap();
    doc.insert_table(
        &EditCtx::local("Owner", "2026-11-06T12:00:00Z"),
        Position::new("body", 2),
        1,
        1,
    )
    .unwrap();
    let mode = |document_mode: u8| {
        let mut env = RenderEnv::default();
        env.compatibility_flags.compatibility_mode = document_mode;
        yrs_doc_to_layout_blocks(&doc, "body", &env)
            .unwrap()
            .iter()
            .find_map(|block| match block {
                LayoutBlock::Table(table) => Some(table.compatibility_mode),
                _ => None,
            })
            .unwrap()
    };
    assert_eq!(mode(15), Some(15));
    assert_eq!(mode(14), Some(14));
    assert_eq!(mode(0), None);
}
