use docx_layout::display_list::Primitive;
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

fn footnote_boundary_input(target: Value) -> Input {
    let mut value = input(vec![prefix(62.0), target], 100.0);
    value.options.footnote_reserved_heights = Some([("1".to_owned(), 20.0)].into_iter().collect());
    value
}

fn assert_full_pitch_fit_above_footnotes(target: Value) {
    let mut value = footnote_boundary_input(target);
    let result = docx_layout::compute_layout_input(&mut value).unwrap();
    assert_eq!(slices(&result, "text"), vec![(1, 0, 1)]);
    let display = docx_layout::build_display_list(&value, &result).unwrap();
    let mut decoration_count = 0;
    for (index, page) in display.pages.iter().enumerate() {
        let layout_page = &result.pages[index];
        let reserved = if index == 0 { 20.0 } else { 0.0 };
        let bottom = layout_page.size.h - layout_page.margins.bottom - reserved;
        for primitive in &page.primitives {
            let (y, height) = match primitive {
                Primitive::Decoration(decoration) => (&decoration.y, &decoration.h),
                Primitive::Rect(rect) => (&rect.y, &rect.h),
                _ => continue,
            };
            if height.as_f64() != Some(60.0) {
                continue;
            }
            decoration_count += 1;
            assert!(y.as_f64().unwrap() + height.as_f64().unwrap() <= bottom);
        }
    }
    assert!(decoration_count > 0);
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
    assert_eq!(
        result.pages[0].size.h - result.pages[0].margins.bottom,
        68.0
    );
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
            paragraph(
                "first",
                1,
                20.0,
                14.0,
                4.0,
                json!({"spacing": {"after": 8}}),
            ),
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
fn text_effects_and_emphasis_require_full_pitch_above_footnotes() {
    for (key, effect) in [
        ("modernEffects", json!({"reflection": {"distance": 2}})),
        ("modernEffects", json!({"glow": {"radius": 4}})),
        ("modernEffects", json!({"shadow": {"distance": 4}})),
        ("modernEffects", json!({"textOutline": {"width": 2}})),
        ("emphasisMark", json!("underDot")),
        ("textShadow", json!(true)),
        ("textOutline", json!(true)),
        ("emboss", json!(true)),
        ("imprint", json!(true)),
        ("textEffect", json!("sparkle")),
    ] {
        let mut target = text(
            1,
            60.0,
            json!({"spacing": {"line": 3, "lineUnit": "multiplier", "lineRule": "auto"}}),
        );
        target["block"]["runs"][0][key] = effect;
        let mut value = footnote_boundary_input(target);
        let result = docx_layout::compute_layout_input(&mut value).unwrap();
        assert_eq!(slices(&result, "text"), vec![(1, 0, 1)], "{key}");
    }
}

#[test]
fn independently_sized_list_markers_require_full_pitch_above_footnotes() {
    for (hidden, page) in [(false, 1), (true, 0)] {
        let target = text(
            1,
            60.0,
            json!({
                "spacing": {"line": 3, "lineUnit": "multiplier", "lineRule": "auto"},
                "listMarker": "g", "listMarkerFontSize": 72,
                "listMarkerFontFamily": "Independent Marker Face", "listMarkerHidden": hidden,
            }),
        );
        let mut value = footnote_boundary_input(target);
        let result = docx_layout::compute_layout_input(&mut value).unwrap();
        assert_eq!(slices(&result, "text"), vec![(page, 0, 1)]);
    }
}

#[test]
fn underline_and_tab_leader_bounds_extend_the_fit_above_footnotes() {
    for decoration in ["underline", "underscore", "heavy"] {
        let mut target = paragraph(
            "text",
            1,
            120.0,
            76.8,
            19.2,
            json!({"spacing": {"line": 1.25, "lineUnit": "multiplier", "lineRule": "auto"}}),
        );
        if decoration == "underline" {
            target["block"]["runs"][0]["fontSize"] = json!(72);
            target["block"]["runs"][0]["underline"] = json!(true);
        } else {
            target["block"]["runs"] = json!([{"kind": "tab", "width": 10, "fontSize": 72}]);
            target["block"]["attrs"]["tabs"] =
                json!([{"val": "left", "pos": 300, "leader": decoration}]);
        }
        for (available, page) in [(96.0, 1), (101.0, 0)] {
            let mut value = input(vec![prefix(180.0 - available), target.clone()], 200.0);
            value.options.footnote_reserved_heights =
                Some([("1".to_owned(), 20.0)].into_iter().collect());
            let result = docx_layout::compute_layout_input(&mut value).unwrap();
            assert_eq!(slices(&result, "text"), vec![(page, 0, 1)], "{decoration}");
            let display = docx_layout::build_display_list(&value, &result).unwrap();
            let mut decoration_count = 0;
            for (index, painted_page) in display.pages.iter().enumerate() {
                let layout_page = &result.pages[index];
                let reserved = if index == 0 { 20.0 } else { 0.0 };
                let bottom = layout_page.size.h - layout_page.margins.bottom - reserved;
                for primitive in &painted_page.primitives {
                    if let Primitive::Decoration(rule) = primitive
                        && rule.attrs.block_key.as_deref() == Some("text")
                    {
                        decoration_count += 1;
                        assert_eq!(rule.h.as_f64(), Some(6.0));
                        assert!(rule.y.as_f64().unwrap() + rule.h.as_f64().unwrap() <= bottom);
                    }
                }
            }
            assert!(decoration_count > 0);
        }
    }
}

#[test]
fn normally_underlined_links_keep_font_box_fit_above_footnotes() {
    let mut target = text(
        1,
        60.0,
        json!({"spacing": {"line": 3, "lineUnit": "multiplier", "lineRule": "auto"}}),
    );
    target["block"]["runs"][0]["fontSize"] = json!(12);
    target["block"]["runs"][0]["underline"] = json!(true);
    target["block"]["runs"][0]["hyperlink"] = json!({"href": "https://example.test"});
    let mut value = footnote_boundary_input(target);
    let result = docx_layout::compute_layout_input(&mut value).unwrap();
    assert_eq!(slices(&result, "text"), vec![(0, 0, 1)]);
    let display = docx_layout::build_display_list(&value, &result).unwrap();
    let rule = display.pages[0]
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Decoration(rule) if rule.attrs.block_key.as_deref() == Some("text") => {
                Some(rule)
            }
            _ => None,
        })
        .unwrap();
    assert!(rule.y.as_f64().unwrap() + rule.h.as_f64().unwrap() <= 90.0);
}

#[test]
fn independently_sized_tab_leader_glyphs_require_full_pitch() {
    let mut target = text(1, 60.0, json!({}));
    target["block"]["runs"] = json!([{
        "kind": "tab", "width": 10,
        "leaderGlyphs": {"glyph": "g", "advance": 10, "count": 1,
            "font": "Independent Leader Face", "fontSize": 72},
    }]);
    let mut value = footnote_boundary_input(target);
    let result = docx_layout::compute_layout_input(&mut value).unwrap();
    assert_eq!(slices(&result, "text"), vec![(1, 0, 1)]);
}

#[test]
fn comment_highlights_require_full_pitch_above_footnotes() {
    for run in [
        json!({"kind": "text", "text": "x", "commentIds": [1]}),
        json!({"kind": "field", "fieldType": "MERGEFIELD", "fallback": "x", "commentIds": [1]}),
    ] {
        let mut target = text(
            1,
            60.0,
            json!({"spacing": {"line": 3, "lineUnit": "multiplier", "lineRule": "auto"}}),
        );
        target["block"]["runs"] = json!([run]);
        assert_full_pitch_fit_above_footnotes(target);
    }
}

#[test]
fn tracked_insertions_and_deletions_require_full_pitch_above_footnotes() {
    for mark in ["isInsertion", "isDeletion"] {
        for mut run in [
            json!({"kind": "text", "text": "x"}),
            json!({"kind": "field", "fieldType": "MERGEFIELD", "fallback": "x"}),
        ] {
            run[mark] = json!(true);
            let mut target = text(
                1,
                60.0,
                json!({"spacing": {"line": 3, "lineUnit": "multiplier", "lineRule": "auto"}}),
            );
            target["block"]["runs"] = json!([run]);
            assert_full_pitch_fit_above_footnotes(target);
        }
    }
}

#[test]
fn paragraph_revision_marks_require_full_pitch_above_footnotes() {
    for mark in ["pPrIns", "pPrDel"] {
        let mut attrs = json!({
            "spacing": {"line": 3, "lineUnit": "multiplier", "lineRule": "auto"},
        });
        attrs[mark] = json!({"revisionId": 1, "author": "Reviewer", "date": "2026-01-01"});
        assert_full_pitch_fit_above_footnotes(text(1, 60.0, attrs));
    }
}

#[test]
fn undecorated_runs_preserve_font_box_fit_above_footnotes() {
    for run in [
        json!({"kind": "text", "text": "x"}),
        json!({"kind": "text", "text": "x", "commentIds": [],
            "isInsertion": false, "isDeletion": false}),
        json!({"kind": "text", "text": "x", "highlight": "#ffff00"}),
    ] {
        let mut target = text(
            1,
            60.0,
            json!({"spacing": {"line": 3, "lineUnit": "multiplier", "lineRule": "auto"}}),
        );
        target["block"]["runs"] = json!([run]);
        let mut value = footnote_boundary_input(target);
        let result = docx_layout::compute_layout_input(&mut value).unwrap();
        assert_eq!(slices(&result, "text"), vec![(0, 0, 1)]);
    }
}

#[test]
fn inline_atoms_require_full_pitch_at_body_and_footnote_boundaries() {
    let image = json!({"kind": "image", "src": "synthetic-image", "width": 30, "height": 30});
    let mut shape = image.clone();
    shape["inlineShape"] = json!({
        "id": "shape", "shapeType": "rect", "geometryPath": [],
        "width": 30, "height": 30, "children": [],
    });
    let mut chart = image.clone();
    chart["inlineChart"] = json!({
        "id": "chart", "chart": {}, "width": 30, "height": 30,
    });
    let reserved = json!({
        "kind": "reserved", "reservation": {"owner": "test", "id": "atom"},
        "width": 30, "height": 30,
    });
    for run in [image, shape, chart, reserved] {
        for reserved_height in [0.0, 20.0] {
            let mut target = text(
                1,
                40.0,
                json!({"spacing": {"line": 2, "lineUnit": "multiplier", "lineRule": "auto"}}),
            );
            target["block"]["runs"] = json!([run.clone()]);
            let mut value = input(vec![prefix(82.0 - reserved_height), target], 100.0);
            value.options.footnote_reserved_heights =
                Some([("1".to_owned(), reserved_height)].into_iter().collect());
            let result = docx_layout::compute_layout_input(&mut value).unwrap();
            assert_eq!(
                slices(&result, "text"),
                vec![(1, 0, 1)],
                "run: {run}, reserved height: {reserved_height}"
            );
        }
    }
}

#[test]
fn inline_atoms_do_not_disable_trailing_leading_on_other_lines() {
    let mut target = text(2, 40.0, json!({"widowControl": false}));
    target["block"]["runs"] = json!([
        {"kind": "image", "src": "synthetic-image", "width": 30, "height": 30},
        {"kind": "text", "text": "x"},
    ]);
    let line = &mut target["measure"]["lines"][1];
    line["headRun"] = json!(1);
    line["headChar"] = json!(0);
    line["tailRun"] = json!(1);
    line["tailChar"] = json!(1);
    let result = layout(vec![target], 58.0);
    assert_eq!(slices(&result, "text"), vec![(0, 0, 2)]);
}

#[test]
fn paragraph_borders_and_shading_require_full_pitch_at_body_and_footnote_boundaries() {
    let mut decorations = vec![json!({"shading": "#eeeeee"})];
    for side in ["bottom", "between", "left", "right", "bar"] {
        let mut attrs = json!({"borders": {}});
        attrs["borders"][side] = json!({"style": "single", "width": 1});
        decorations.push(attrs);
    }
    decorations.push(json!({"borders": {"bottom": {}}}));
    for mut attrs in decorations {
        attrs["spacing"] = json!({"line": 3, "lineUnit": "multiplier", "lineRule": "auto"});
        for reserved_height in [0.0, 20.0] {
            let mut value = input(
                vec![prefix(82.0 - reserved_height), text(1, 60.0, attrs.clone())],
                100.0,
            );
            value.options.footnote_reserved_heights =
                Some([("1".to_owned(), reserved_height)].into_iter().collect());
            let result = docx_layout::compute_layout_input(&mut value).unwrap();
            assert_eq!(
                slices(&result, "text"),
                vec![(1, 0, 1)],
                "attrs: {attrs}, reserved height: {reserved_height}"
            );
        }
    }
}

#[test]
fn invisible_paragraph_borders_preserve_font_box_fit() {
    for border in [
        json!({"style": "none", "width": 1}),
        json!({"style": "nil", "width": 1}),
        json!({"style": "single", "width": 0}),
    ] {
        let result = layout(
            vec![
                prefix(82.0),
                text(
                    1,
                    60.0,
                    json!({"borders": {"bottom": border, "between": border}}),
                ),
            ],
            100.0,
        );
        assert_eq!(slices(&result, "text"), vec![(0, 0, 1)]);
    }
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
    value.options.section_page_float_bands =
        Some(serde_json::from_value(json!([{"default": [{"top": 68, "bottom": 80}]}])).unwrap());
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
