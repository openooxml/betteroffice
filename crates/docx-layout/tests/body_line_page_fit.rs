use docx_layout::types::{BlockId, Fragment, Input, Layout};
use serde_json::{Value, json};

fn paragraph(
    id: &str,
    lines: usize,
    height: f64,
    ascent: f64,
    descent: f64,
    attrs: Value,
) -> Value {
    json!({
        "block": {
            "kind": "paragraph", "id": id,
            "runs": [{"kind": "text", "text": "x".repeat(lines)}],
            "attrs": attrs,
        },
        "measure": {
            "kind": "paragraph", "totalHeight": lines as f64 * height,
            "lines": (0..lines).map(|index| json!({
                "headRun": 0, "headChar": index, "tailRun": 0, "tailChar": index + 1,
                "width": 10, "ascent": ascent, "descent": descent, "lineHeight": height,
            })).collect::<Vec<_>>(),
        },
    })
}

fn text(lines: usize, height: f64, attrs: Value) -> Value {
    paragraph("text", lines, height, 14.0, 4.0, attrs)
}

fn prefix(height: f64) -> Value {
    paragraph("prefix", 1, height, height, 0.0, json!({}))
}

fn input(measured: Vec<Value>, body_height: f64) -> Input {
    serde_json::from_value(json!({
        "measured": measured,
        "options": {
            "pageSize": {"w": 200, "h": body_height + 20.0},
            "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10},
        },
    }))
    .unwrap()
}

fn layout(measured: Vec<Value>, body_height: f64) -> Layout {
    docx_layout::compute_layout_input(&mut input(measured, body_height)).unwrap()
}

fn slices(layout: &Layout, id: &str) -> Vec<(usize, usize, usize)> {
    layout
        .pages
        .iter()
        .enumerate()
        .flat_map(|(page, value)| {
            value
                .fragments
                .iter()
                .filter_map(move |fragment| match fragment {
                    Fragment::Paragraph(fragment)
                        if fragment.block_id == BlockId::Str(id.to_owned()) =>
                    {
                        Some((page, fragment.from_line, fragment.to_line))
                    }
                    _ => None,
                })
        })
        .collect()
}

#[test]
fn single_spaced_font_box_fits_without_trailing_leading() {
    let attrs = json!({"widowControl": false, "spacing": {
        "line": 1, "lineUnit": "multiplier", "lineRule": "auto"
    }});
    let result = layout(vec![text(3, 20.0, attrs)], 58.0);
    assert_eq!(slices(&result, "text"), vec![(0, 0, 3)]);
    let Fragment::Paragraph(fragment) = &result.pages[0].fragments[0] else {
        panic!()
    };
    assert_eq!(fragment.height, 60.0);
    assert_eq!(fragment.y + fragment.height, 70.0);
    assert_eq!(result.pages[0].size.h - result.pages[0].margins.bottom, 68.0);
}

#[test]
fn font_box_crossing_body_bottom_moves_the_line() {
    let result = layout(vec![text(3, 20.0, json!({"widowControl": false}))], 57.0);
    assert_eq!(slices(&result, "text"), vec![(0, 0, 2), (1, 2, 3)]);
}

#[test]
fn auto_276_keeps_the_font_box_and_advances_by_the_full_pitch() {
    let attrs = json!({"widowControl": false, "spacing": {
        "line": 1.15, "lineUnit": "multiplier", "lineRule": "auto"
    }});
    let result = layout(vec![text(3, 23.0, attrs)], 64.0);
    assert_eq!(slices(&result, "text"), vec![(0, 0, 3)]);
    let Fragment::Paragraph(fragment) = &result.pages[0].fragments[0] else {
        panic!()
    };
    assert_eq!(fragment.height, 69.0);
}

#[test]
fn exact_line_boxes_still_require_their_full_height() {
    let attrs = json!({"widowControl": false, "spacing": {
        "line": 20, "lineUnit": "px", "lineRule": "exact"
    }});
    let result = layout(vec![paragraph("text", 3, 20.0, 16.0, 4.0, attrs)], 58.0);
    assert_eq!(slices(&result, "text"), vec![(0, 0, 2), (1, 2, 3)]);
}

#[test]
fn at_least_preserves_its_floor_and_can_use_natural_trailing_leading() {
    let attrs = json!({"widowControl": false, "spacing": {
        "line": 24, "lineUnit": "px", "lineRule": "atLeast"
    }});
    let active = layout(vec![paragraph("text", 3, 24.0, 20.0, 4.0, attrs)], 70.0);
    assert_eq!(slices(&active, "text"), vec![(0, 0, 2), (1, 2, 3)]);
    let attrs = json!({"widowControl": false, "spacing": {
        "line": 16, "lineUnit": "px", "lineRule": "atLeast"
    }});
    let natural = layout(vec![text(3, 20.0, attrs)], 58.0);
    assert_eq!(slices(&natural, "text"), vec![(0, 0, 3)]);
}

#[test]
fn widow_control_uses_font_box_fit_for_the_first_two_lines() {
    let result = layout(vec![prefix(62.0), text(4, 20.0, json!({}))], 100.0);
    assert_eq!(slices(&result, "text"), vec![(0, 0, 2), (1, 2, 4)]);
}

#[test]
fn widow_control_can_move_two_lines_whose_leading_exceeds_a_fresh_page() {
    let result = layout(vec![prefix(1.0), text(4, 20.0, json!({}))], 38.0);
    assert_eq!(slices(&result, "text"), vec![(1, 0, 2), (2, 2, 4)]);
}

#[test]
fn keep_lines_uses_only_the_terminal_lines_trailing_leading() {
    let fitting = layout(
        vec![prefix(42.0), text(3, 20.0, json!({"keepLines": true}))],
        100.0,
    );
    assert_eq!(slices(&fitting, "text"), vec![(0, 0, 3)]);
    let crossing = layout(
        vec![prefix(43.0), text(3, 20.0, json!({"keepLines": true}))],
        100.0,
    );
    assert_eq!(slices(&crossing, "text"), vec![(1, 0, 3)]);
}

#[test]
fn keep_next_discounts_only_the_followers_terminal_leading() {
    for (height, page) in [(62.0, 0), (64.0, 1)] {
        let result = layout(
            vec![
                prefix(height),
                paragraph("heading", 1, 20.0, 14.0, 4.0, json!({"keepNext": true})),
                text(1, 20.0, json!({})),
            ],
            100.0,
        );
        assert_eq!(slices(&result, "heading"), vec![(page, 0, 1)]);
        assert_eq!(slices(&result, "text"), vec![(page, 0, 1)]);
    }
}

#[test]
fn split_keep_next_retains_full_pitch_before_its_follower() {
    let result = layout(
        vec![
            prefix(42.0),
            paragraph(
                "head",
                4,
                20.0,
                14.0,
                4.0,
                json!({"widowControl": false, "keepNext": true}),
            ),
            text(1, 20.0, json!({})),
        ],
        100.0,
    );
    assert_eq!(slices(&result, "head"), vec![(0, 0, 3), (1, 3, 4)]);
    assert_eq!(slices(&result, "text"), vec![(1, 0, 1)]);
}

#[test]
fn paragraph_after_spacing_is_deferred_and_full_pitch_separates_paragraphs() {
    let result = layout(
        vec![
            paragraph("first", 1, 20.0, 14.0, 4.0, json!({"spacing": {"after": 8}})),
            text(1, 20.0, json!({"spacing": {"after": 50}})),
        ],
        46.0,
    );
    assert_eq!(slices(&result, "text"), vec![(0, 0, 1)]);
    let Fragment::Paragraph(fragment) = &result.pages[0].fragments[1] else {
        panic!()
    };
    assert_eq!(fragment.y, 38.0);
    assert_eq!(fragment.height, 20.0);
}

#[test]
fn footnote_reservation_uses_the_same_font_box_boundary() {
    let mut value = input(vec![text(3, 20.0, json!({"widowControl": false}))], 100.0);
    value.options.footnote_reserved_heights = Some([("1".to_owned(), 42.0)].into_iter().collect());
    let result = docx_layout::compute_layout_input(&mut value).unwrap();
    assert_eq!(slices(&result, "text"), vec![(0, 0, 3)]);
}

#[test]
fn column_bottom_uses_the_same_font_box_boundary() {
    let mut value = input(vec![text(6, 20.0, json!({"widowControl": false}))], 58.0);
    value.options.columns = Some(serde_json::from_value(json!({"count": 2, "gap": 10})).unwrap());
    let result = docx_layout::compute_layout_input(&mut value).unwrap();
    assert_eq!(result.pages.len(), 1);
    assert_eq!(slices(&result, "text"), vec![(0, 0, 3), (0, 3, 6)]);
}

#[test]
fn trailing_leading_does_not_overhang_a_float_band() {
    let mut value = input(vec![text(3, 20.0, json!({"widowControl": false}))], 100.0);
    value.options.section_page_float_bands = Some(
        serde_json::from_value(json!([{"default": [{"top": 68, "bottom": 80}]}])).unwrap(),
    );
    let result = docx_layout::compute_layout_input(&mut value).unwrap();
    let fragments = &result.pages[0].fragments;
    let Fragment::Paragraph(before) = &fragments[0] else {
        panic!()
    };
    let Fragment::Paragraph(after) = &fragments[1] else {
        panic!()
    };
    assert_eq!(before.to_line, 2);
    assert_eq!(after.y, 80.0);
}
