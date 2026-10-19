use pptx_edit::{DeckSession, EditCtx, TextStyle};
use pptx_parse::{GraphicFrameData, PictureCrop, ShapeNode};
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

#[test]
fn table_cell_picture_fills_are_clipped_to_spans_below_borders_and_text() {
    let session =
        DeckSession::open(include_bytes!("fixtures/table-picture-fill.pptx"), 4102).unwrap();
    let rendered = table(&display_list(&session));
    let images: Vec<_> = rendered
        .cells
        .iter()
        .enumerate()
        .filter_map(|(index, primitive)| {
            if let Primitive::Image {
                x,
                y,
                w,
                h,
                asset_id,
                crop,
                tile,
                path,
                effects,
                ..
            } = primitive
            {
                Some((index, (*x, *y, *w, *h), asset_id, crop, tile, path, effects))
            } else {
                None
            }
        })
        .collect();
    assert_eq!(images.len(), 2, "both origin cells paint their pictures");
    let rounded =
        |rect: (f32, f32, f32, f32)| (round(rect.0), round(rect.1), round(rect.2), round(rect.3));
    assert_eq!(rounded(images[0].1), (64.0, 64.0, 320.0, 38.93));
    assert_eq!(rounded(images[1].1), (64.0, 102.93, 192.0, 77.87));
    for image in &images {
        assert_eq!(image.2.as_deref(), Some("ppt/media/image1.png"));
        assert!(image.5.is_none(), "rectangular images clip to their bounds");
    }
    let crop = images[0].3;
    assert!((crop.left - 1.42558 / 3.44418).abs() < 1e-6);
    assert!((crop.top - 2.14634 / 22.60975).abs() < 1e-6);
    assert!((crop.right - 1.01860 / 3.44418).abs() < 1e-6);
    assert!((crop.bottom - 19.46341 / 22.60975).abs() < 1e-6);
    assert_eq!(images[0].6, &[pptx_render::ImageEffect::Grayscale]);
    let tile = images[1].4.as_ref().expect("the row span uses a tile");
    assert_eq!((tile.scale_x, tile.scale_y), (0.5, 2.0));
    let last_image = images.last().unwrap().0;
    assert!(rendered.cells.iter().enumerate().all(|(index, primitive)| {
        !matches!(
            primitive,
            Primitive::TextBox { .. }
                | Primitive::Shape {
                    stroke: Some(_),
                    ..
                }
        ) || index > last_image
    }));
    let restored =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 4103).unwrap();
    assert_eq!(display_list(&session), display_list(&restored));
}

fn cell_picture(fill_rect: PictureCrop, crop: PictureCrop) -> Primitive {
    const FIXTURE: &[u8] = include_bytes!("fixtures/table-picture-fill.pptx");
    let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    let ShapeNode::GraphicFrame(frame) = &mut package.slides[0].shapes[0] else {
        panic!("expected a table frame")
    };
    let GraphicFrameData::Table(model) = &mut frame.data else {
        panic!("expected a table")
    };
    let picture = model.rows[0].cells[0].picture_fill.as_mut().unwrap();
    picture.fill_rect = fill_rect;
    picture.crop = crop;
    let session = DeckSession::from_package_with_source(package, FIXTURE, 4104).unwrap();
    table(&display_list(&session))
        .cells
        .into_iter()
        .find(|primitive| matches!(primitive, Primitive::Image { .. }))
        .unwrap()
}

#[test]
fn table_cell_picture_positive_fill_insets_letterbox_the_destination() {
    let Primitive::Image {
        x, y, w, h, crop, ..
    } = cell_picture(
        PictureCrop {
            left: 25_000,
            right: 25_000,
            ..PictureCrop::default()
        },
        PictureCrop::default(),
    )
    else {
        panic!("expected a cell image")
    };
    assert_eq!(
        (round(x), round(y), round(w), round(h)),
        (144.0, 64.0, 160.0, 38.93)
    );
    assert!(crop.is_whole());

    let Primitive::Image {
        x, y, w, h, crop, ..
    } = cell_picture(
        PictureCrop {
            top: 25_000,
            bottom: 25_000,
            ..PictureCrop::default()
        },
        PictureCrop::default(),
    )
    else {
        panic!("expected a cell image")
    };
    assert_eq!(
        (round(x), round(y), round(w), round(h)),
        (64.0, 73.73, 320.0, 19.47)
    );
    assert!(crop.is_whole());
}

#[test]
fn table_cell_picture_mixed_fill_insets_letterbox_and_crop_the_source() {
    for (fill_rect, expected_rect, expected_crop) in [
        (
            PictureCrop {
                left: 25_000,
                top: -50_000,
                right: -25_000,
                bottom: 25_000,
            },
            (144.0, 64.0, 240.0, 29.2),
            (0.1, 0.41, 0.375, 0.05),
        ),
        (
            PictureCrop {
                left: -25_000,
                top: 25_000,
                right: 25_000,
                bottom: -50_000,
            },
            (64.0, 73.73, 240.0, 29.2),
            (0.275, 0.05, 0.2, 0.41),
        ),
    ] {
        let Primitive::Image {
            x, y, w, h, crop, ..
        } = cell_picture(
            fill_rect,
            PictureCrop {
                left: 10_000,
                top: 5_000,
                right: 20_000,
                bottom: 5_000,
            },
        )
        else {
            panic!("expected a cell image")
        };
        assert_eq!((round(x), round(y), round(w), round(h)), expected_rect);
        for (actual, expected) in [crop.left, crop.top, crop.right, crop.bottom]
            .into_iter()
            .zip([
                expected_crop.0,
                expected_crop.1,
                expected_crop.2,
                expected_crop.3,
            ])
        {
            assert!((actual - expected).abs() < 1e-6, "{actual} != {expected}");
        }
    }
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
