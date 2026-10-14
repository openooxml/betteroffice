use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_render::{Paint, Primitive, SlideRenderer, SurfaceDisplayList};

const DECK: &[u8] = include_bytes!("fixtures/table-basic.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn session() -> DeckSession {
    DeckSession::open(DECK, 4101).unwrap()
}

fn display_list(session: &DeckSession) -> SurfaceDisplayList {
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
}

fn round(value: f32) -> f32 {
    (value * 100.0).round() / 100.0
}

struct RenderedTable {
    rect: (f32, f32, f32, f32),
    label: String,
    cells: Vec<Primitive>,
}

fn table(list: &SurfaceDisplayList) -> RenderedTable {
    list.primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Table {
                x,
                y,
                w,
                h,
                label,
                primitives,
                ..
            } => Some(RenderedTable {
                rect: (round(*x), round(*y), round(*w), round(*h)),
                label: label.clone(),
                cells: primitives.clone(),
            }),
            _ => None,
        })
        .expect("the frame paints a table")
}

fn text_of(primitive: &Primitive) -> Option<String> {
    match primitive {
        Primitive::TextBox { lines, .. } => Some(
            lines
                .iter()
                .flat_map(|line| line.runs.iter().map(|run| run.text.as_str()))
                .collect(),
        ),
        _ => None,
    }
}

fn boxes(cells: &[Primitive]) -> Vec<(String, f32, f32, f32, f32)> {
    cells
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::TextBox { x, y, w, h, .. } => Some((
                text_of(primitive)?,
                round(*x),
                round(*y),
                round(*w),
                round(*h),
            )),
            _ => None,
        })
        .collect()
}

fn fills(cells: &[Primitive]) -> Vec<(String, f32, f32, f32, f32)> {
    cells
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                geometry,
                x,
                y,
                w,
                h,
                fill: Some(Paint::Solid { color }),
                ..
            } if geometry == "rect" => {
                Some((color.clone(), round(*x), round(*y), round(*w), round(*h)))
            }
            _ => None,
        })
        .collect()
}

#[test]
fn a_table_frame_does_not_spill_its_first_cell_onto_the_slide() {
    let list = display_list(&session());
    let stray: Vec<String> = list.primitives.iter().filter_map(text_of).collect();
    assert!(
        stray.is_empty(),
        "a table frame must not paint a cell as loose slide text, but painted {stray:?}"
    );
}

#[test]
fn the_frame_paints_a_table_instead_of_a_placeholder() {
    let list = display_list(&session());
    assert!(
        !list
            .primitives
            .iter()
            .any(|primitive| matches!(primitive, Primitive::Placeholder { .. })),
        "the frame still paints a placeholder"
    );
    let table = table(&list);
    assert_eq!(table.rect, (64.0, 64.0, 600.0, 116.8));
    assert_eq!(table.label, "Table, 3 rows, 3 columns");
    let texts: Vec<String> = boxes(&table.cells).into_iter().map(|cell| cell.0).collect();
    assert_eq!(
        texts,
        [
            "Header spans two",
            "Third",
            "Tall",
            "Middle",
            "Edge",
            "Below",
            "Last"
        ],
        "a merged continuation must not paint"
    );
}

#[test]
fn cells_sit_on_the_declared_grid_and_a_span_covers_what_it_merges() {
    let table = table(&display_list(&session()));
    let placed = boxes(&table.cells);

    assert_eq!(
        placed[0],
        ("Header spans two".to_owned(), 64.0, 64.0, 320.0, 38.93)
    );
    assert_eq!(placed[1], ("Third".to_owned(), 384.0, 64.0, 280.0, 38.93));
    assert_eq!(
        placed[3],
        ("Middle".to_owned(), 256.0, 102.93, 128.0, 38.93)
    );
    assert_eq!(
        placed[2],
        ("Tall".to_owned(), 64.0, 102.93, 192.0, 77.87),
        "the rowSpan origin covers both of its rows"
    );
}

#[test]
fn the_style_cascade_and_the_cells_own_properties_both_reach_the_paint() {
    let table = table(&display_list(&session()));
    let painted = fills(&table.cells);

    assert_eq!(painted[0].0, "#4472C4", "the header row takes firstRow");
    assert_eq!(painted[0].3, 320.0, "the header fill spans both columns");
    assert_eq!(
        painted[2],
        ("#DDEBF7".to_owned(), 64.0, 102.93, 192.0, 77.87),
        "the cell's own solidFill wins over the band"
    );
    assert_eq!(painted[3].0, "#F2F2F2", "the first data row is banded");
    assert_eq!(painted[5].0, "#FFFFFF", "the second data row is not");

    let borders: Vec<(String, f32)> = table
        .cells
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                geometry,
                stroke: Some(stroke),
                ..
            } if geometry == "line" => Some((stroke.color.clone(), stroke.width)),
            _ => None,
        })
        .collect();
    assert!(
        borders.contains(&("#C00000".to_owned(), 2.0)),
        "the cell's own a:lnB is missing from {borders:?}"
    );
}

#[test]
fn every_fill_paints_before_every_border() {
    let table = table(&display_list(&session()));
    let order: Vec<&str> = table
        .cells
        .iter()
        .map(|primitive| match primitive {
            Primitive::Shape { geometry, .. } => geometry.as_str(),
            _ => "text",
        })
        .collect();
    let last_fill = order.iter().rposition(|kind| *kind == "rect").unwrap();
    let first_border = order.iter().position(|kind| *kind == "line").unwrap();
    assert!(
        last_fill < first_border,
        "a shared edge would be overpainted: {order:?}"
    );
}

#[test]
fn a_row_grows_to_hold_text_that_overruns_its_declared_height() {
    let session = session();
    let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[4]
        .id
        .clone();
    let before = table(&display_list(&session));
    let declared = boxes(&before.cells)[3].4;

    session
        .insert_text(
            &EditCtx::local("test"),
            &story_id,
            0,
            "a middle cell holding far more text than its declared row height leaves room for",
            &TextStyle::default(),
        )
        .unwrap();
    let after = table(&display_list(&session));
    let placed = boxes(&after.cells);

    assert!(
        placed[3].4 > declared,
        "the row kept its declared height of {declared} instead of growing"
    );
    assert!(
        after.rect.3 > before.rect.3,
        "the table did not grow past the frame it was declared with"
    );
    assert!(
        (placed[5].2 - (placed[3].2 + placed[3].4)).abs() < 0.05,
        "the row below did not move down with it: {} against {}",
        placed[5].2,
        placed[3].2 + placed[3].4
    );
}

#[test]
fn a_turned_table_keeps_the_frame_height_so_its_pivot_does_not_move() {
    const TURNED: &[u8] = include_bytes!("fixtures/table-rotated.pptx");
    const OVERRUN: &str =
        "a middle cell holding far more text than its declared row height leaves room for";

    let mut heights = Vec::new();
    for deck in [DECK, TURNED] {
        let session = DeckSession::open(deck, 4102).unwrap();
        let story_id = session.snapshot().unwrap().slides[0].shapes[0].text_stories[4]
            .id
            .clone();
        session
            .insert_text(
                &EditCtx::local("test"),
                &story_id,
                0,
                OVERRUN,
                &TextStyle::default(),
            )
            .unwrap();
        heights.push(table(&display_list(&session)).rect.3);
    }

    assert!(
        heights[0] > heights[1],
        "an upright table grows past its frame ({}) while a turned one keeps the frame height \
         ({}) so the rotation pivot stays on the frame centre",
        heights[0],
        heights[1]
    );
}
