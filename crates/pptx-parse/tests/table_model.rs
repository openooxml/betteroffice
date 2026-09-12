use pptx_parse::{GraphicFrameData, ShapeNode, Table, parse_pptx};

const FIXTURE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");

fn table() -> Table {
    let package = parse_pptx(FIXTURE).unwrap();
    let ShapeNode::GraphicFrame(frame) = &package.slides[0].shapes[0] else {
        panic!("expected a graphic frame");
    };
    let GraphicFrameData::Table(table) = &frame.data else {
        panic!("expected a table");
    };
    table.clone()
}

#[test]
fn parses_the_grid_row_heights_and_table_properties() {
    let table = table();
    assert_eq!(table.grid, [1_828_800, 1_219_200, 2_667_000]);
    assert_eq!(
        table.rows.iter().map(|row| row.height).collect::<Vec<_>>(),
        [370_840, 370_840, 370_840]
    );
    assert!(table.properties.first_row);
    assert!(table.properties.band_row);
    assert!(!table.properties.last_row);
    assert!(!table.properties.first_col);
    assert!(!table.properties.band_col);
    assert_eq!(
        table.properties.style_id.as_deref(),
        Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    );
}

#[test]
fn parses_spans_and_their_continuation_cells_in_place() {
    let table = table();
    let text = |row: usize, cell: usize| {
        table.rows[row].cells[cell].text.paragraphs[0].runs[0]
            .text
            .clone()
    };
    assert_eq!(
        table
            .rows
            .iter()
            .map(|row| row.cells.len())
            .collect::<Vec<_>>(),
        [3, 3, 3]
    );

    assert_eq!(text(0, 0), "Header spans two");
    assert_eq!(table.rows[0].cells[0].grid_span, 2);
    assert!(!table.rows[0].cells[0].merged);
    assert!(table.rows[0].cells[1].merged);
    assert_eq!(table.rows[0].cells[1].grid_span, 1);
    assert_eq!(text(0, 2), "Third");

    assert_eq!(text(1, 0), "Tall");
    assert_eq!(table.rows[1].cells[0].row_span, 2);
    assert!(table.rows[2].cells[0].merged);
    assert_eq!(text(2, 1), "Below");
}

#[test]
fn parses_direct_cell_fills_and_borders() {
    let table = table();
    let fill = table.rows[1].cells[0].fill.as_ref().unwrap();
    assert_eq!(fill.fill_type, "solid");
    assert_eq!(fill.color.as_ref().unwrap().rgb.as_deref(), Some("DDEBF7"));
    assert!(table.rows[1].cells[1].fill.is_none());

    let bottom = table.rows[1].cells[2].borders.bottom.as_ref().unwrap();
    assert_eq!(bottom.width, Some(19_050.0));
    assert_eq!(
        bottom.color.as_ref().unwrap().rgb.as_deref(),
        Some("C00000")
    );
    assert!(table.rows[1].cells[2].borders.left.is_none());
    assert!(table.rows[1].cells[1].borders.is_empty());
}

#[test]
fn folds_cell_properties_into_the_cell_text_body() {
    let table = table();
    let text = &table.rows[0].cells[0].text;
    assert_eq!(text.anchor.as_deref(), Some("ctr"));
    assert_eq!(text.inset_left, Some(91_440));
    assert_eq!(text.inset_right, Some(91_440));
    assert_eq!(text.inset_top, None);
    assert_eq!(text.vertical, None);
}

#[test]
fn a_text_only_table_round_trips_through_the_released_json() {
    let legacy = r#"{"type":"table","rows":[[{"anchor":"ctr","vertical":null,"autofit":null,"insetLeft":null,"insetTop":null,"insetRight":null,"insetBottom":null,"paragraphs":[]}]]}"#;
    let data: GraphicFrameData = serde_json::from_str(legacy).unwrap();
    let GraphicFrameData::Table(stored) = &data else {
        panic!("expected a table");
    };
    assert!(stored.grid.is_empty());
    assert_eq!(stored.rows[0].cells[0].grid_span, 1);
    assert_eq!(stored.rows[0].cells[0].text.anchor.as_deref(), Some("ctr"));
    assert_eq!(serde_json::to_string(&data).unwrap(), legacy);

    let parsed = serde_json::to_string(&GraphicFrameData::Table(table())).unwrap();
    assert!(parsed.contains("\"grid\":[1828800,1219200,2667000]"));
    assert!(parsed.contains("\"gridSpan\":2"));
}

#[test]
fn a_row_or_cell_carrying_anything_beyond_text_keeps_the_full_encoding() {
    let text = pptx_parse::TextBody::default();
    let borders = pptx_parse::TableCellBorders {
        bottom: Some(ooxml_drawingml::ShapeOutline::default()),
        ..pptx_parse::TableCellBorders::default()
    };
    for (height, cell) in [
        (370_840, pptx_parse::TableCell::from_text(text.clone())),
        (
            0,
            pptx_parse::TableCell {
                grid_span: 2,
                ..pptx_parse::TableCell::from_text(text.clone())
            },
        ),
        (
            0,
            pptx_parse::TableCell {
                row_span: 2,
                ..pptx_parse::TableCell::from_text(text.clone())
            },
        ),
        (
            0,
            pptx_parse::TableCell {
                merged: true,
                ..pptx_parse::TableCell::from_text(text.clone())
            },
        ),
        (
            0,
            pptx_parse::TableCell {
                fill: Some(ooxml_drawingml::ShapeFill::named("solid")),
                ..pptx_parse::TableCell::from_text(text.clone())
            },
        ),
        (
            0,
            pptx_parse::TableCell {
                borders,
                ..pptx_parse::TableCell::from_text(text.clone())
            },
        ),
    ] {
        let data = GraphicFrameData::Table(pptx_parse::Table {
            rows: vec![pptx_parse::TableRow {
                height,
                cells: vec![cell],
            }],
            ..pptx_parse::Table::default()
        });
        let json = serde_json::to_string(&data).unwrap();
        assert!(
            json.contains("\"cells\""),
            "a row or cell with formatting must not serialize as a bare text list: {json}"
        );
        assert_eq!(
            serde_json::from_str::<GraphicFrameData>(&json).unwrap(),
            data
        );
    }
}
