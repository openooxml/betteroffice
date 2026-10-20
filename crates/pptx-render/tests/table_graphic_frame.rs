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

#[test]
fn normal_autofit_table_cell_keeps_its_scale_and_grows_its_row_after_typing() {
    for cell_index in [0, 4] {
        assert_normal_autofit_cell(cell_index);
    }
}

fn assert_normal_autofit_cell(cell_index: usize) {
    let mut package = pptx_parse::parse_pptx(DECK).unwrap();
    let path = "ppt/slides/slide1.xml";
    let mut source = std::str::from_utf8(package.part_bytes(path).unwrap())
        .unwrap()
        .replace("<a:tr h=\"370840\">", "<a:tr h=\"952500\">");
    let body_start = source
        .match_indices("<a:bodyPr/>")
        .nth(cell_index)
        .unwrap()
        .0;
    source.replace_range(
        body_start..body_start + "<a:bodyPr/>".len(),
        "<a:bodyPr><a:normAutofit fontScale=\"50000\" lnSpcReduction=\"10000\"/></a:bodyPr>",
    );
    assert!(package.replace_part(path, source.as_bytes().to_vec()));
    let bytes = pptx_parse::write_pptx(&package).unwrap();
    let session = DeckSession::open(&bytes, 4102).unwrap();
    session.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
    let story = session.snapshot().unwrap().slides[0].shapes[0].text_stories[cell_index].clone();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let before = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    let untouched = pptx_parse::parse_pptx(&renderer.save(&session).unwrap()).unwrap();
    assert_eq!(untouched.part_bytes(path).unwrap(), source.as_bytes());
    session
        .insert_text(
            &EditCtx::local("test"),
            &story.id,
            story.length - 1,
            &" more words to fit in this cell".repeat(8),
            &TextStyle::default(),
        )
        .unwrap();
    let after = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    let before_table = table(&before.display_list);
    let after_table = table(&after.display_list);
    let rendered_index = if cell_index == 0 { 0 } else { 3 };
    assert!(
        boxes(&after_table.cells)[rendered_index].4 > boxes(&before_table.cells)[rendered_index].4
    );
    assert!(after_table.rect.3 > before_table.rect.3);
    let (y, height, width, lines) = after_table
        .cells
        .iter()
        .filter_map(|cell| match cell {
            Primitive::TextBox { y, h, w, lines, .. } => Some((*y, *h, *w, lines)),
            _ => None,
        })
        .nth(rendered_index)
        .unwrap();
    assert!(
        lines
            .iter()
            .all(|line| line.y + line.height <= y + height + 0.01)
    );
    assert!(lines.iter().all(|line| line.width <= width + 0.01));
    let saved = renderer.save(&session).unwrap();
    let reopened = DeckSession::open(&saved, 4103).unwrap();
    let pptx_parse::ShapeNode::GraphicFrame(frame) = &reopened.package().slides[0].shapes[0] else {
        panic!("expected table frame");
    };
    let pptx_parse::GraphicFrameData::Table(table) = &frame.data else {
        panic!("expected table");
    };
    let Some(pptx_parse::TextAutofit::Normal {
        font_scale: Some(scale),
        line_space_reduction: Some(reduction),
    }) = table.rows[cell_index / 3].cells[cell_index % 3]
        .text
        .autofit
    else {
        panic!("expected persisted normal autofit");
    };
    assert_eq!(scale, 0.5, "typing must preserve the stored cell scale");
    assert_eq!(reduction, 0.1);
    let expected_size = ((18.0 * scale).round().max(1.0) * 96.0 / 72.0) as f32;
    assert!((lines[0].runs[0].font_size_px - expected_size).abs() < 0.01);
    assert_eq!(table.rows[0].cells[2].text.autofit, None);
    let mut reopened_renderer = SlideRenderer::new();
    reopened_renderer
        .register_font("Arial", false, false, FONT)
        .unwrap();
    assert_eq!(
        reopened_renderer
            .layout_slide(reopened.package(), &reopened.snapshot().unwrap(), 0)
            .unwrap()
            .display_list,
        after.display_list,
    );
    session.add_undo_barrier();
    assert!(session.undo());
    assert_eq!(
        renderer
            .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
            .unwrap()
            .display_list,
        before.display_list,
    );
    let restored = pptx_parse::parse_pptx(&renderer.save(&session).unwrap()).unwrap();
    assert_eq!(restored.part_bytes(path).unwrap(), source.as_bytes());
}

#[test]
fn text_columns_measure_table_rows_with_the_rendered_flow() {
    let session = session();
    let mut package = session.package().clone();
    let pptx_parse::ShapeNode::GraphicFrame(frame) = &mut package.slides[0].shapes[0] else {
        panic!("expected a table frame");
    };
    let pptx_parse::GraphicFrameData::Table(table_data) = &mut frame.data else {
        panic!("expected table data");
    };
    table_data.rows.truncate(1);
    table_data.grid.truncate(1);
    table_data.rows[0].height = 428_625;
    table_data.rows[0].cells.truncate(1);
    let cell = &mut table_data.rows[0].cells[0];
    cell.grid_span = 1;
    cell.row_span = 1;
    cell.text.num_col = Some(2);
    cell.text.spc_col = Some(95_250);
    cell.text.anchor = Some("t".to_owned());
    cell.text.paragraphs.truncate(1);
    cell.text.paragraphs[0].runs.truncate(1);
    cell.text.paragraphs[0].runs[0].text = "one\ntwo\nthree\nfour\nfive".to_owned();
    cell.text.paragraphs[0].properties.line_spacing =
        Some(pptx_parse::LineSpacing::Points { value: 12.0 });
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, FONT).unwrap();
    let mut snapshot = session.snapshot().unwrap();
    snapshot.slides[0].shapes.clear();
    package.masters[0].shapes = package.slides[0].shapes.clone();
    package.layouts[0].shapes.clear();
    let render = |package: &pptx_parse::PptxPackage| {
        table(
            &renderer
                .layout_slide(package, &snapshot, 0)
                .unwrap()
                .display_list,
        )
    };
    let columns = render(&package);
    let pptx_parse::ShapeNode::GraphicFrame(frame) = &mut package.masters[0].shapes[0] else {
        unreachable!()
    };
    let pptx_parse::GraphicFrameData::Table(data) = &mut frame.data else {
        unreachable!()
    };
    data.rows[0].cells[0].text.num_col = None;
    let single = render(&package);
    let columns_row = &boxes(&columns.cells)[0];
    let single_row = &boxes(&single.cells)[0];
    assert_eq!(columns.rect.3, 116.8);
    assert_eq!(single.rect.3, 116.8);
    assert_eq!(columns_row.4, 57.6);
    assert_eq!(single_row.4, 89.6);
    assert!(
        columns_row.4 < single_row.4,
        "{} >= {}",
        columns_row.4,
        single_row.4
    );
    let lines = columns
        .cells
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => Some(lines),
            _ => None,
        })
        .unwrap();
    assert_eq!(lines.len(), 5);
    assert!(lines.iter().any(|line| line.x > lines[0].x + 50.0));
    assert!(
        lines
            .iter()
            .all(|line| line.y + line.height <= columns_row.2 + columns_row.4 + 0.05)
    );
}
