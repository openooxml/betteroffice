use ooxml_drawingml::{
    ColorValue, ResolvedCellStyle, ShapeFill, TableCellPosition, TableCellStyle, TableStyleFlags,
};
use pptx_parse::parse_pptx;

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");
const STYLE_ID: &str = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

fn flags() -> TableStyleFlags {
    TableStyleFlags {
        first_row: true,
        band_row: true,
        ..Default::default()
    }
}

fn position(row: usize, column: usize) -> TableCellPosition {
    TableCellPosition {
        row,
        column,
        row_count: 3,
        column_count: 3,
    }
}

fn solid(fill: &Option<ShapeFill>) -> Option<&str> {
    fill.as_ref()
        .and_then(|fill| fill.color.as_ref())
        .and_then(|color| color.rgb.as_deref())
}

fn resolve(row: usize, column: usize) -> ResolvedCellStyle {
    let package = parse_pptx(DECK).unwrap();
    package
        .table_styles
        .resolve_cell(Some(STYLE_ID), flags(), position(row, column))
}

#[test]
fn the_table_styles_part_is_parsed_into_the_package() {
    let package = parse_pptx(DECK).unwrap();

    assert_eq!(
        package.table_styles.default_style_id.as_deref(),
        Some(STYLE_ID)
    );
    assert_eq!(package.table_styles.styles.len(), 1);
    assert_eq!(
        package
            .table_styles
            .style(Some(STYLE_ID))
            .unwrap()
            .style_name
            .as_deref(),
        Some("Fixture Style")
    );
}

#[test]
fn the_header_row_resolves_to_bold_white_text_on_the_first_row_fill() {
    let resolved = resolve(0, 0);

    assert_eq!(solid(&resolved.fill), Some("4472C4"));
    assert_eq!(resolved.bold, Some(true));
    assert_eq!(
        resolved.color.and_then(|color| color.rgb),
        Some("FFFFFF".to_owned())
    );
}

#[test]
fn data_rows_alternate_between_the_band_fill_and_the_whole_table_fill() {
    assert_eq!(solid(&resolve(1, 0).fill), Some("F2F2F2"));
    assert_eq!(solid(&resolve(2, 0).fill), Some("FFFFFF"));
}

#[test]
fn the_whole_table_border_reaches_the_outer_edges_only() {
    let resolved = resolve(2, 0);

    let left = resolved.left.expect("left edge");
    assert_eq!(left.width, Some(12700.0));
    assert_eq!(
        left.color.as_ref().and_then(|color| color.rgb.as_deref()),
        Some("7F7F7F")
    );
    assert_eq!(resolved.bottom.map(|edge| edge.width), Some(Some(12700.0)));
    assert!(resolved.right.is_none());
    assert!(resolved.top.is_none());
}

#[test]
fn an_unknown_style_id_leaves_the_cell_unstyled() {
    let package = parse_pptx(DECK).unwrap();

    let resolved =
        package
            .table_styles
            .resolve_cell(Some("{NOT-IN-THIS-DECK}"), flags(), position(0, 0));

    assert_eq!(solid(&resolved.fill), None);
}

#[test]
fn a_cells_own_fill_wins_over_the_resolved_style() {
    let package = parse_pptx(DECK).unwrap();
    let mut resolved = package
        .table_styles
        .resolve_cell(Some(STYLE_ID), flags(), position(1, 0));
    resolved.apply_cell_style(&TableCellStyle {
        fill: Some(ShapeFill {
            fill_type: "solid".to_owned(),
            color: Some(ColorValue {
                rgb: Some("DDEBF7".to_owned()),
                ..Default::default()
            }),
            gradient: None,
        }),
        ..Default::default()
    });

    assert_eq!(solid(&resolved.fill), Some("DDEBF7"));
}
