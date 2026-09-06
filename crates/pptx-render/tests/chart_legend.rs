use pptx_edit::DeckSession;
use pptx_render::{Paint, PositionedTextLine, Primitive, SlideRenderer, TextAlign};

const FIXTURE: &[u8] = include_bytes!("fixtures/chart-legend.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn chart(slide: usize) -> Vec<Primitive> {
    let session = DeckSession::open(FIXTURE, 316).unwrap();
    let mut renderer = SlideRenderer::new();
    for bold in [false, true] {
        renderer.register_font("Arial", bold, false, FONT).unwrap();
    }
    renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), slide)
        .unwrap()
        .display_list
        .primitives
        .into_iter()
        .find_map(|primitive| match primitive {
            Primitive::Chart { primitives, .. } => Some(primitives),
            _ => None,
        })
        .unwrap()
}

fn swatches(parts: &[Primitive]) -> Vec<(f32, f32, &str)> {
    parts
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Shape {
                x,
                y,
                w,
                h,
                fill: Some(Paint::Solid { color }),
                ..
            } if *w == 8.0 && *h == 8.0 => Some((*x, *y, color.as_str())),
            _ => None,
        })
        .collect()
}

fn text<'a>(parts: &'a [Primitive], value: &str) -> &'a PositionedTextLine {
    parts
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => lines
                .iter()
                .find(|line| line.runs.iter().any(|run| run.text == value)),
            _ => None,
        })
        .unwrap()
}

#[test]
fn top_and_bottom_legends_reserve_their_own_rows() {
    for (slide, y, plot_y) in [(0, 417.0, 124.0), (1, 131.0, 146.0)] {
        let parts = chart(slide);
        let swatches = swatches(&parts);
        assert_eq!(swatches.len(), 2);
        assert_eq!(swatches[0].1, y);
        assert_eq!(swatches[1].1, y);
        assert!(swatches[0].0 < swatches[1].0);
        assert_eq!(swatches[0].2, "#6254E7");
        assert_eq!(swatches[1].2, "#1FA97A");
        assert!(parts.iter().any(|primitive| matches!(primitive,
            Primitive::Shape { x, y, w, h, .. }
            if *x == 138.0 && *y == plot_y && *w == 0.0 && *h == 252.0
        )));
        let title = text(&parts, "Revenue");
        assert!((title.x + title.width / 2.0 - 384.0).abs() < 0.001);
    }
}

#[test]
fn long_legend_entries_wrap_without_overlapping_or_leaving_the_frame() {
    let parts = chart(4);
    let swatches = swatches(&parts);
    assert_eq!(swatches.len(), 2);
    assert!(swatches[0].1 < swatches[1].1);
    let north = text(&parts, "Northwestern international division");
    let south = text(&parts, "Southeastern international division");
    for line in [north, south] {
        assert!(line.x >= 96.0);
        assert!(line.x + line.width <= 296.0);
    }
    assert!(north.y + north.height <= south.y);
    assert!(south.y + south.height <= 432.0);
}

#[test]
fn a_large_legend_font_stays_between_the_title_and_plot() {
    let parts = chart(5);
    let title = text(&parts, "Revenue");
    let legend = text(&parts, "North");
    assert_eq!(legend.runs[0].font_size_px, 40.0);
    assert!(legend.y >= title.y + title.height);
    let plot_top = parts
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Shape { x, y, w, h, .. } if *x == 138.0 && *w == 0.0 && *h > 8.0 => Some(*y),
            _ => None,
        })
        .unwrap();
    assert!(legend.y + legend.height <= plot_top);
}

#[test]
fn the_composed_chart_title_preserves_its_alignment() {
    let output = pptx_render::compile_json(
        r##"{
        "widthPx":320,"heightPx":180,
        "shapes":[{"kind":"chart","id":4,"name":"Revenue chart",
        "rect":{"x":10,"y":20,"w":300,"h":150},"rotationDeg":0,
        "chart":{"chartType":"column","title":"Revenue","plotGroups":[],
        "series":[{"name":"North","categories":["Q1"],"values":[3],"color":"#6254E7"}]}}]
    }"##,
    )
    .unwrap();
    let output: pptx_render::SurfaceDisplayList = serde_json::from_str(&output).unwrap();
    let Primitive::Chart { primitives, .. } = &output.primitives[0] else {
        panic!()
    };
    for primitive in primitives {
        if let Primitive::TextBox {
            paragraphs, x, w, ..
        } = primitive
        {
            let is_title = paragraphs[0].runs[0].text == "Revenue";
            assert_eq!(
                paragraphs[0].align,
                Some(if is_title {
                    TextAlign::Center
                } else {
                    TextAlign::Left
                })
            );
            if is_title {
                assert_eq!(x + w / 2.0, 160.0);
            }
        }
    }
}

#[test]
fn legend_rows_use_shaped_widths_for_wide_glyphs() {
    let parts = chart(8);
    let swatches = swatches(&parts);
    assert_eq!(swatches.len(), 2);
    assert!(swatches[0].1 < swatches[1].1);
    for value in ["WWWWWWWWWW", "MMMMMMMMMM"] {
        let line = text(&parts, value);
        assert!(line.x + line.width <= 296.0);
    }
}

#[test]
fn a_single_long_legend_label_wraps_without_losing_text() {
    let parts = chart(9);
    let mut labels = Vec::<Vec<&PositionedTextLine>>::new();
    for part in &parts {
        match part {
            Primitive::Shape { w, h, .. } if *w == 8.0 && *h == 8.0 => labels.push(Vec::new()),
            Primitive::TextBox { lines, .. } if !labels.is_empty() => {
                labels.last_mut().unwrap().extend(lines)
            }
            _ => {}
        }
    }
    assert_eq!(labels.len(), 2);
    for (lines, value) in labels.iter().zip([
        "Northwestern international division with offices across the entire continent",
        "Southeastern international division with offices across the entire continent",
    ]) {
        let combined: String = lines
            .iter()
            .flat_map(|line| line.runs.iter().map(|run| run.text.as_str()))
            .collect();
        assert_eq!(combined, value);
        assert!(lines.len() > 1);
        for line in lines {
            assert!(line.x >= 96.0 && line.x + line.width <= 296.0);
            assert!(line.y + line.height <= 432.0);
        }
    }
    let all: Vec<_> = labels.into_iter().flatten().collect();
    assert!(
        all.windows(2)
            .all(|pair| pair[0].y + pair[0].height <= pair[1].y)
    );
}
