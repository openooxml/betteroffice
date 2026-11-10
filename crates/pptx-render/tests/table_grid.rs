#[allow(dead_code)]
#[path = "../../pptx-parse/tests/support/header_footer.rs"]
mod support;

use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_render::{Primitive, SlideRenderer, TableGrid, TableGridCell, TableHeader};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn cell(attributes: &str, text: &str) -> String {
    let paragraph = if text.is_empty() {
        "<a:p/>".to_owned()
    } else {
        format!("<a:p><a:r><a:t>{text}</a:t></a:r></a:p>")
    };
    format!(
        r#"<a:tc{attributes}><a:txBody><a:bodyPr/><a:lstStyle/>{paragraph}</a:txBody><a:tcPr/></a:tc>"#
    )
}

fn frame(properties: &str, rows: &[&[String]]) -> String {
    let columns = rows.iter().map(|row| row.len()).max().unwrap_or_default();
    let grid = r#"<a:gridCol w="2000000"/>"#.repeat(columns);
    let rows = rows
        .iter()
        .map(|cells| format!(r#"<a:tr h="400000">{}</a:tr>"#, cells.concat()))
        .collect::<String>();
    format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="1200000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr{properties}/><a:tblGrid>{grid}</a:tblGrid>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

/// Three by three, header row and first column on, with a column span in the
/// header row, a row span down the first column and an empty cell:
///
/// | Region (2 columns)  | Total |
/// | North (2 rows) | Q1 | (empty) |
/// |                | Q2 | 9     |
fn merged() -> String {
    frame(
        r#" firstRow="1" firstCol="1""#,
        &[
            &[
                cell(r#" gridSpan="2""#, "Region"),
                cell(r#" hMerge="1""#, ""),
                cell("", "Total"),
            ],
            &[
                cell(r#" rowSpan="2""#, "North"),
                cell("", "Q1"),
                cell("", ""),
            ],
            &[cell(r#" vMerge="1""#, ""), cell("", "Q2"), cell("", "9")],
        ],
    )
}

fn open(shapes: &str, id: u64) -> DeckSession {
    let bytes = support::deck(&[(shapes, false)], "", "", "", None);
    DeckSession::open(&bytes, id).unwrap()
}

fn tables(session: &DeckSession) -> Vec<Primitive> {
    fn collect(primitive: Primitive, out: &mut Vec<Primitive>) {
        match primitive {
            Primitive::Table { .. } => out.push(primitive),
            Primitive::Chart { primitives, .. } => {
                for child in primitives {
                    collect(child, out);
                }
            }
            _ => {}
        }
    }
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let mut out = Vec::new();
    for primitive in renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
        .primitives
    {
        collect(primitive, &mut out);
    }
    out
}

fn only_table(session: &DeckSession) -> (TableGrid, Vec<Primitive>) {
    let mut tables = tables(session);
    assert_eq!(tables.len(), 1, "one table on the slide");
    let Primitive::Table {
        grid, primitives, ..
    } = tables.remove(0)
    else {
        unreachable!()
    };
    (grid.expect("a laid-out table carries its grid"), primitives)
}

/// The text a cell's primitive range paints, paragraphs joined by newlines.
fn cell_text(primitives: &[Primitive], cell: &TableGridCell) -> String {
    let [start, end] = cell.primitives;
    primitives[start as usize..end as usize]
        .iter()
        .flat_map(|primitive| match primitive {
            Primitive::TextBox { paragraphs, .. } => paragraphs
                .iter()
                .map(|paragraph| {
                    paragraph
                        .runs
                        .iter()
                        .map(|run| run.text.as_str())
                        .collect::<String>()
                })
                .collect::<Vec<_>>(),
            other => panic!("a cell range holds only text boxes, found {other:?}"),
        })
        .filter(|text| !text.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

type Shape = (u32, u32, u32, u32, Option<TableHeader>);

fn shape(cell: &TableGridCell) -> Shape {
    (
        cell.row,
        cell.column,
        cell.row_span,
        cell.column_span,
        cell.header,
    )
}

#[test]
fn a_table_reports_origin_cells_with_spans_and_headers() {
    let (grid, primitives) = only_table(&open(&merged(), 9101));
    assert_eq!((grid.rows, grid.columns), (3, 3));
    let column = Some(TableHeader::Column);
    let row = Some(TableHeader::Row);
    assert_eq!(
        grid.cells.iter().map(shape).collect::<Vec<_>>(),
        [
            (0, 0, 1, 2, column),
            (0, 2, 1, 1, column),
            (1, 0, 2, 1, row),
            (1, 1, 1, 1, None),
            (1, 2, 1, 1, None),
            (2, 1, 1, 1, None),
            (2, 2, 1, 1, None),
        ],
        "merged continuations are left out and the empty cell stays"
    );
    assert_eq!(
        grid.cells
            .iter()
            .map(|cell| cell_text(&primitives, cell))
            .collect::<Vec<_>>(),
        ["Region", "Total", "North", "Q1", "", "Q2", "9"]
    );
}

#[test]
fn cell_ranges_are_ordered_disjoint_and_cover_every_text_box() {
    let (grid, primitives) = only_table(&open(&merged(), 9102));
    let mut previous = 0;
    for cell in &grid.cells {
        let [start, end] = cell.primitives;
        assert!(start <= end, "{cell:?}");
        assert!(start >= previous, "ranges do not overlap: {cell:?}");
        previous = end;
    }
    assert!(previous as usize <= primitives.len());
    let covered = grid
        .cells
        .iter()
        .map(|cell| (cell.primitives[1] - cell.primitives[0]) as usize)
        .sum::<usize>();
    let text_boxes = primitives
        .iter()
        .filter(|primitive| matches!(primitive, Primitive::TextBox { .. }))
        .count();
    assert_eq!(covered, text_boxes, "every text box belongs to one cell");
}

#[test]
fn a_table_without_header_flags_has_no_header_cells() {
    let table = frame(
        "",
        &[
            &[cell("", "a"), cell("", "b")],
            &[cell("", "c"), cell("", "d")],
        ],
    );
    let (grid, primitives) = only_table(&open(&table, 9103));
    assert_eq!((grid.rows, grid.columns), (2, 2));
    assert_eq!(
        grid.cells.iter().map(shape).collect::<Vec<_>>(),
        [
            (0, 0, 1, 1, None),
            (0, 1, 1, 1, None),
            (1, 0, 1, 1, None),
            (1, 1, 1, 1, None),
        ]
    );
    assert_eq!(
        grid.cells
            .iter()
            .map(|cell| cell_text(&primitives, cell))
            .collect::<Vec<_>>(),
        ["a", "b", "c", "d"]
    );
}

#[test]
fn only_the_first_column_heads_rows_when_the_header_row_is_off() {
    let table = frame(
        r#" firstCol="1""#,
        &[
            &[cell("", "Name"), cell("", "Value")],
            &[cell("", "Width"), cell("", "4")],
        ],
    );
    let (grid, _) = only_table(&open(&table, 9104));
    assert_eq!(
        grid.cells
            .iter()
            .map(|cell| cell.header)
            .collect::<Vec<_>>(),
        [Some(TableHeader::Row), None, Some(TableHeader::Row), None]
    );
}

#[test]
fn a_table_inside_a_group_keeps_its_grid() {
    let (grid, primitives) = only_table(&open(&support::group(&merged()), 9105));
    assert_eq!((grid.rows, grid.columns, grid.cells.len()), (3, 3, 7));
    assert_eq!(cell_text(&primitives, &grid.cells[0]), "Region");
}

#[test]
fn edited_cell_text_lands_in_its_cell_range() {
    let session = open(&merged(), 9106);
    let story = session.snapshot().unwrap().slides[0].shapes[0]
        .text_stories
        .iter()
        .map(|story| story.id.clone())
        .find(|id| id.ends_with(":table:1:2"))
        .expect("the empty cell has a story");
    session
        .insert_text(
            &EditCtx::local("test"),
            &story,
            0,
            "7",
            &TextStyle::default(),
        )
        .unwrap();
    let (grid, primitives) = only_table(&session);
    let edited = grid
        .cells
        .iter()
        .find(|cell| (cell.row, cell.column) == (1, 2))
        .unwrap();
    assert_eq!(cell_text(&primitives, edited), "7");
}

#[test]
fn the_grid_serializes_as_camel_case_and_omits_plain_headers() {
    let table = tables(&open(&merged(), 9107)).remove(0);
    let json = serde_json::to_value(&table).unwrap();
    let grid = &json["grid"];
    assert_eq!(grid["rows"], 3);
    assert_eq!(grid["columns"], 3);
    let first = &grid["cells"][0];
    assert_eq!(first["row"], 0);
    assert_eq!(first["column"], 0);
    assert_eq!(first["rowSpan"], 1);
    assert_eq!(first["columnSpan"], 2);
    assert_eq!(first["header"], "column");
    assert!(first["primitives"].is_array());
    assert_eq!(grid["cells"][2]["header"], "row");
    assert!(grid["cells"][3].get("header").is_none());
}

#[test]
fn a_table_without_a_grid_still_deserializes() {
    let table = tables(&open(&merged(), 9108)).remove(0);
    let mut json = serde_json::to_value(&table).unwrap();
    json.as_object_mut().unwrap().remove("grid");
    let Primitive::Table { grid, .. } = serde_json::from_value::<Primitive>(json).unwrap() else {
        unreachable!()
    };
    assert!(grid.is_none());
}
