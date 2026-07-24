//! Display-list gates:
//! 1. contract round-trip — the committed TS demo fixture survives the serde
//!    types with no field loss and serializes idempotently (byte-stable).
//! 2. determinism snapshots — the paragraph golden scenarios build to
//!    byte-identical JSON run-over-run, pinned by committed snapshot files.
//!    Regenerate deliberately: DL_SNAPSHOT_UPDATE=1 cargo test -p docx-layout
//! 3. hit-testing — point -> PM position and range -> rects behave like the
//!    painted-DOM resolvers on the same scenarios.

use docx_layout::display_list::{
    DecoKind, DecorationPrimitive, DisplayList, DocAttrs, MAX_ALT_TEXT_CHARS, Primitive,
    RevisionKind, ShapePathCommand, StructuralRevisionKind, StructuralRevisionScope, TableCellRef,
    build_display_list_json,
};
use docx_layout::hit::{VerticalDirection, caret_rect, hit_test, range_rects, vertical_move};

const DEMO_FIXTURE: &str =
    include_str!("../../../packages/docx/src/layout/render/__fixtures__/displayList.demo.json");

const SCENARIOS: &[&str] = &[
    "single-page-multi-paragraph",
    "multi-page-paragraph-overflow",
];

fn fixture_path(name: &str, suffix: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(format!("{name}.{suffix}.json"))
}

// numeric-tolerant deep equality: serde_json distinguishes 96 from 96.0, but
// the contract doesn't — compare numbers as f64
fn value_eq(a: &serde_json::Value, b: &serde_json::Value) -> bool {
    use serde_json::Value::*;
    match (a, b) {
        (Number(x), Number(y)) => {
            (x.as_f64().unwrap_or(f64::NAN) - y.as_f64().unwrap_or(f64::NAN)).abs() < 1e-9
        }
        (Array(x), Array(y)) => x.len() == y.len() && x.iter().zip(y).all(|(a, b)| value_eq(a, b)),
        (Object(x), Object(y)) => {
            x.len() == y.len()
                && x.iter()
                    .all(|(k, v)| y.get(k).is_some_and(|w| value_eq(v, w)))
        }
        _ => a == b,
    }
}

#[test]
fn demo_fixture_round_trips_through_serde_types() {
    let typed: DisplayList = serde_json::from_str(DEMO_FIXTURE).expect("fixture parses");

    // no field loss: re-serialized output is value-identical to the fixture
    let out = serde_json::to_string(&typed).expect("serializes");
    let original: serde_json::Value = serde_json::from_str(DEMO_FIXTURE).unwrap();
    let round_tripped: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert!(
        value_eq(&original, &round_tripped),
        "round-trip dropped or renamed fields:\n{out}"
    );

    // byte-stable: parse -> serialize is idempotent and deterministic
    let typed2: DisplayList = serde_json::from_str(&out).expect("own output parses");
    let out2 = serde_json::to_string(&typed2).unwrap();
    assert_eq!(out, out2, "serialization is not byte-stable");
    assert_eq!(typed, typed2);
}

#[test]
fn paragraph_scenarios_snapshot_and_determinism() {
    let update = std::env::var("DL_SNAPSHOT_UPDATE").as_deref() == Ok("1");

    for name in SCENARIOS {
        let input = std::fs::read_to_string(fixture_path(name, "input")).unwrap_or_else(|_| {
            panic!("missing input fixture for {name}; run scripts/export-displaylist-fixtures.ts")
        });

        // determinism gate: two runs are byte-identical
        let a = build_display_list_json(&input).expect("builds");
        let b = build_display_list_json(&input).expect("builds");
        assert_eq!(a, b, "{name}: display list build is not deterministic");

        let snapshot_file = fixture_path(name, "displaylist");
        if update {
            let pretty: serde_json::Value = serde_json::from_str(&a).unwrap();
            std::fs::write(
                &snapshot_file,
                format!("{}\n", serde_json::to_string_pretty(&pretty).unwrap()),
            )
            .expect("write snapshot");
            continue;
        }

        let expected = std::fs::read_to_string(&snapshot_file).unwrap_or_else(|_| {
            panic!(
                "missing snapshot for {name}; run DL_SNAPSHOT_UPDATE=1 cargo test -p docx-layout"
            )
        });
        let pretty: serde_json::Value = serde_json::from_str(&a).unwrap();
        let actual = format!("{}\n", serde_json::to_string_pretty(&pretty).unwrap());
        assert_eq!(
            actual, expected,
            "{name}: display list drifted from snapshot"
        );
    }
}

/// Determinism gate over the WHOLE golden corpus, not just the two paragraph
/// scenarios in `SCENARIOS`. Every fixture the display-list builder currently
/// accepts must build byte-identically run-over-run AND match its committed
/// `<name>.displaylist.json` snapshot. This is the standing replacement for the
/// mirror-vs-painter watchdog: once the DOM painter is retired there is no
/// second renderer to diff against, so display-list drift must be caught here.
///
/// Scenarios the builder does not yet accept (composing the golden layout still
/// hits a not-yet-ported path) are REPORTED, never failed, so the gate widens
/// automatically as the builder gains coverage — add nothing here when a new
/// primitive lands, just regenerate:
///   DL_SNAPSHOT_UPDATE=1 cargo test -p docx-layout --test display_list
#[test]
fn golden_corpus_display_list_determinism_and_snapshot() {
    let update = std::env::var("DL_SNAPSHOT_UPDATE").as_deref() == Ok("1");
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");

    let mut names: Vec<String> = std::fs::read_dir(&fixtures)
        .expect("fixtures dir exists — run `bun scripts/export-golden-fixtures.ts`")
        .filter_map(|e| {
            let n = e.ok()?.file_name().into_string().ok()?;
            n.strip_suffix(".input.json").map(str::to_string)
        })
        .collect();
    names.sort();
    assert!(!names.is_empty(), "no fixtures in {fixtures:?}");

    let mut snapshotted: Vec<String> = Vec::new();
    let mut unsupported: Vec<String> = Vec::new();
    let mut failures: Vec<String> = Vec::new();

    for name in &names {
        let input = std::fs::read_to_string(fixture_path(name, "input")).unwrap();
        // Paginator-only fixtures pin `{ measured, options }`; splice the golden
        // canonical layout in as the `layout` the display-list builder consumes.
        // Fixtures that already carry their own `layout` (the paragraph pair)
        // are used verbatim, so their snapshot matches `SCENARIOS`.
        let composed = {
            let mut v: serde_json::Value = serde_json::from_str(&input).unwrap();
            if v.get("layout").is_none() {
                let golden = std::fs::read_to_string(fixture_path(name, "golden")).unwrap();
                v["layout"] = serde_json::from_str(&golden).unwrap();
            }
            v.to_string()
        };

        let a = match build_display_list_json(&composed) {
            Ok(json) => json,
            Err(_) => {
                unsupported.push(name.clone());
                continue;
            }
        };
        let b = build_display_list_json(&composed).expect("second build");
        if a != b {
            failures.push(format!("{name}: display-list build is not deterministic"));
            continue;
        }

        let snapshot_file = fixture_path(name, "displaylist");
        let pretty: serde_json::Value = serde_json::from_str(&a).unwrap();
        let actual = format!("{}\n", serde_json::to_string_pretty(&pretty).unwrap());
        if update {
            std::fs::write(&snapshot_file, &actual).expect("write snapshot");
            snapshotted.push(name.clone());
            continue;
        }
        match std::fs::read_to_string(&snapshot_file) {
            Ok(expected) if expected == actual => snapshotted.push(name.clone()),
            Ok(_) => failures.push(format!(
                "{name}: display list drifted from committed snapshot \
                 (regenerate deliberately: DL_SNAPSHOT_UPDATE=1)"
            )),
            Err(_) => failures.push(format!(
                "{name}: missing snapshot; run DL_SNAPSHOT_UPDATE=1 cargo test -p docx-layout"
            )),
        }
    }

    println!(
        "\n display-list determinism: {} snapshotted [{}]\n  unsupported (builder gap, not a failure): {} [{}]",
        snapshotted.len(),
        snapshotted.join(", "),
        unsupported.len(),
        unsupported.join(", "),
    );
    assert!(failures.is_empty(), "\n{}", failures.join("\n"));
}

fn build(name: &str) -> DisplayList {
    let input = std::fs::read_to_string(fixture_path(name, "input")).unwrap();
    let json = build_display_list_json(&input).unwrap();
    serde_json::from_str(&json).unwrap()
}

fn snapshot(name: &str) -> DisplayList {
    let json = std::fs::read_to_string(fixture_path(name, "displaylist")).unwrap();
    serde_json::from_str(&json).unwrap()
}

// single-page-multi-paragraph: three one-line paragraphs at y 96/120/144,
// x 96, widths 120/130/120, pm spans [1,16] [18,34] [36,51]
#[test]
fn hit_test_resolves_clicks_like_the_dom_resolver() {
    let dl = build("single-page-multi-paragraph");
    assert_eq!(dl.pages.len(), 1);

    // click at the very start of the first line lands on its first position
    assert_eq!(hit_test(&dl, 0, 96.0, 110.0), Some(1));

    // click inside the first line resolves proportionally within its pm span
    let mid = hit_test(&dl, 0, 156.0, 110.0).unwrap();
    assert!((1..=16).contains(&mid), "mid-line hit out of span: {mid}");
    assert!(mid > 1, "mid-line hit should advance past the run start");

    // click right of the line's end snaps to the line end (nearest-span rule)
    assert_eq!(hit_test(&dl, 0, 400.0, 110.0), Some(16));

    // click in the second paragraph's band resolves into its span
    let second = hit_test(&dl, 0, 100.0, 133.0).unwrap();
    assert!(
        (18..=34).contains(&second),
        "second-paragraph hit: {second}"
    );

    // out-of-range page yields nothing
    assert_eq!(hit_test(&dl, 5, 100.0, 100.0), None);
}

#[test]
fn hit_test_reaches_content_on_later_pages() {
    let dl = build("multi-page-paragraph-overflow");
    assert_eq!(dl.pages.len(), 2);

    // page 2's first paragraph is block 8 (pm 121..133), painted at y 96
    let pos = hit_test(&dl, 1, 100.0, 150.0).unwrap();
    assert!(pos >= 121, "page-2 hit resolved into page-1 content: {pos}");
}

#[test]
fn vertical_move_follows_wrapped_lines_and_preserves_goal_x() {
    let dl = snapshot("mixed-run-line-segments");
    let down = vertical_move(&dl, 4, VerticalDirection::Down, None).unwrap();
    assert!((10..=26).contains(&down.position));

    let up = vertical_move(&dl, down.position, VerticalDirection::Up, Some(down.goal_x)).unwrap();
    assert_eq!(up.position, 4);
    assert_eq!(up.goal_x, down.goal_x);
}

#[test]
fn vertical_move_crosses_paragraphs_columns_and_pages() {
    let paragraphs = snapshot("single-page-multi-paragraph");
    let second = vertical_move(&paragraphs, 8, VerticalDirection::Down, None).unwrap();
    assert!((18..=34).contains(&second.position));

    let columns = snapshot("column-break-in-multi-column");
    let right_column = vertical_move(&columns, 5, VerticalDirection::Down, None).unwrap();
    assert!((14..=26).contains(&right_column.position));

    let pages = snapshot("multi-page-paragraph-overflow");
    let next_page = vertical_move(&pages, 112, VerticalDirection::Down, None).unwrap();
    assert!((121..=132).contains(&next_page.position));
}

#[test]
fn vertical_move_uses_table_rows_as_visual_lines() {
    let text = |value: &str,
                x: f64,
                baseline: f64,
                doc_start: i64,
                doc_end: i64,
                cell: Option<(u64, u64)>| {
        let mut primitive = serde_json::json!({
            "kind": "text",
            "text": value,
            "x": x,
            "baselineY": baseline,
            "width": 80,
            "font": "400 16px Calibri",
            "color": "#000000",
            "docStart": doc_start,
            "docEnd": doc_end,
            "blockId": doc_start
        });
        if let Some((row, col)) = cell {
            primitive["cell"] = serde_json::json!({
                "row": row,
                "col": col,
                "rowSpan": 1,
                "colSpan": 1
            });
            primitive["table"] = serde_json::json!({
                "tableId": "table-1",
                "rowStart": 0,
                "rowEnd": 2,
                "rowCount": 2,
                "columnCount": 2
            });
        }
        primitive
    };
    let dl: DisplayList = serde_json::from_value(serde_json::json!({
        "pages": [{
            "pageIndex": 0,
            "width": 500,
            "height": 500,
            "columnBounds": [{"x": 80, "y": 80, "width": 340, "height": 340}],
            "primitives": [
                text("left", 100.0, 120.0, 1, 5, Some((0, 0))),
                text("below", 100.0, 150.0, 20, 25, Some((1, 0))),
                text("right", 220.0, 120.0, 10, 15, Some((0, 1))),
                text("below", 220.0, 150.0, 30, 35, Some((1, 1))),
                text("after", 100.0, 190.0, 40, 45, None)
            ]
        }]
    }))
    .unwrap();

    let next_row = vertical_move(&dl, 2, VerticalDirection::Down, None).unwrap();
    assert!((20..=25).contains(&next_row.position));

    let after_table = vertical_move(&dl, 32, VerticalDirection::Down, None).unwrap();
    assert!((40..=45).contains(&after_table.position));

    let into_table = vertical_move(
        &dl,
        after_table.position,
        VerticalDirection::Up,
        Some(240.0),
    )
    .unwrap();
    assert!((30..=35).contains(&into_table.position));
}

#[test]
fn range_rects_cover_selected_lines() {
    let dl = build("single-page-multi-paragraph");

    // selecting the whole first paragraph yields one rect over its line
    let rects = range_rects(&dl, 1, 16);
    assert_eq!(rects.len(), 1);
    let r = &rects[0];
    assert_eq!(r.page_index, 0);
    assert!((r.x - 96.0).abs() < 0.001, "rect x: {}", r.x);
    assert!((r.width - 120.0).abs() < 0.001, "rect width: {}", r.width);
    assert!(r.y > 90.0 && r.y < 120.0, "rect y: {}", r.y);

    // a sub-range shrinks the rect proportionally and keeps it inside the line
    let sub = range_rects(&dl, 4, 10);
    assert_eq!(sub.len(), 1);
    assert!(sub[0].x > 96.0 && sub[0].x + sub[0].width < 96.0 + 120.0 + 0.001);

    // a cross-paragraph range emits one rect per affected line
    let cross = range_rects(&dl, 1, 40);
    assert_eq!(cross.len(), 3);

    // collapsed range selects nothing
    assert!(range_rects(&dl, 5, 5).is_empty());
}

#[test]
fn caret_rect_uses_forward_and_trailing_edges() {
    let dl = build("single-page-multi-paragraph");
    let start = caret_rect(&dl, 1).unwrap();
    assert_eq!(start.page_index, 0);
    assert!((start.x - 96.0).abs() < 0.001);

    let end = caret_rect(&dl, 16).unwrap();
    assert_eq!(end.page_index, 0);
    assert!((end.x - 216.0).abs() < 0.001);
    assert_eq!(start.y, end.y);
    assert_eq!(start.height, end.height);
}

#[test]
fn range_rects_span_pages() {
    let dl = build("multi-page-paragraph-overflow");
    // blocks 0..9 at pm i*15+1..i*15+13; select across the page break
    let rects = range_rects(&dl, 1, 148);
    let pages: std::collections::BTreeSet<usize> = rects.iter().map(|r| r.page_index).collect();
    assert_eq!(pages.into_iter().collect::<Vec<_>>(), vec![0, 1]);
    assert_eq!(rects.len(), 10);
}

#[test]
fn page_geometry_metadata_uses_authored_margins_and_columns() {
    let input = serde_json::json!({
        "measured": [],
        "options": {},
        "layout": { "pages": [{
            "size": { "w": 500.0, "h": 700.0 },
            "margins": { "top": 60.0, "right": 40.0, "bottom": 80.0, "left": 50.0 },
            "columns": {
                "count": 3,
                "gap": 12.0,
                "equalWidth": false,
                "columns": [
                    { "width": 100.0, "space": 10.0 },
                    { "width": 140.0, "space": 20.0 },
                    {}
                ]
            },
            "fragments": []
        }] }
    });
    let json = build_display_list_json(&input.to_string()).expect("display list builds");
    let dl: DisplayList = serde_json::from_str(&json).expect("display list parses");
    let page = &dl.pages[0];
    let content = page
        .content_bounds
        .as_ref()
        .expect("content bounds emitted");
    assert_eq!(content.x.as_f64(), Some(50.0));
    assert_eq!(content.y.as_f64(), Some(60.0));
    assert_eq!(content.width.as_f64(), Some(410.0));
    assert_eq!(content.height.as_f64(), Some(560.0));

    let columns: Vec<(f64, f64)> = page
        .column_bounds
        .iter()
        .map(|bounds| (bounds.x.as_f64().unwrap(), bounds.width.as_f64().unwrap()))
        .collect();
    // The missing third width receives the remaining authored content budget:
    // 410 - gaps(10+20) - known widths(100+140) = 140.
    assert_eq!(columns, vec![(50.0, 100.0), (160.0, 140.0), (320.0, 140.0)]);
}

// ---------------------------------------------------------------------------
// block identity for string ids (blockKey), table cell structure, alt text
// ---------------------------------------------------------------------------

fn doc_attrs(p: &Primitive) -> Option<&DocAttrs> {
    match p {
        Primitive::Text(t) => Some(&t.attrs),
        Primitive::GlyphRun(g) => Some(&g.attrs),
        Primitive::Rect(r) => Some(&r.attrs),
        Primitive::Image(i) => Some(&i.attrs),
        Primitive::Shape(s) => Some(&s.attrs),
        Primitive::Decoration(d) => Some(&d.attrs),
        Primitive::Line(_) => None,
    }
}

fn chart_text_prims(prims: &[Primitive]) -> Vec<&docx_layout::display_list::TextRunPrimitive> {
    prims
        .iter()
        .filter_map(|p| match p {
            Primitive::Text(t) => Some(t),
            _ => None,
        })
        .collect()
}

fn line_prims(prims: &[Primitive]) -> Vec<&docx_layout::display_list::LinePrimitive> {
    prims
        .iter()
        .filter_map(|p| match p {
            Primitive::Line(l) => Some(l),
            _ => None,
        })
        .collect()
}

fn chart_display_list(chart_type: &str) -> DisplayList {
    let input = serde_json::json!({
        "measured": [{
            "block": {
                "kind": "chart",
                "id": 42,
                "width": 260.0,
                "height": 180.0,
                "docStart": 4,
                "docEnd": 5,
                "chart": {
                    "type": "chart",
                    "chartType": chart_type,
                    "title": if chart_type == "pie" { "Share" } else { "Revenue" },
                    "legend": { "position": "right", "visible": true },
                    "series": [{
                        "name": "North",
                        "categories": ["Q1", "Q2"],
                        "values": [10.0, 20.0],
                        "color": "#4472C4"
                    }],
                    "axes": { "value": { "min": 0.0, "max": 25.0 } }
                }
            },
            "measure": { "kind": "chart", "width": 260.0, "height": 180.0 }
        }],
        "options": {},
        "layout": { "pages": [{
            "size": { "w": 400.0, "h": 300.0 },
            "margins": {},
            "fragments": [{
                "kind": "chart",
                "blockId": 42,
                "x": 50.0,
                "y": 40.0,
                "width": 260.0,
                "height": 180.0,
                "docStart": 4,
                "docEnd": 5
            }]
        }] }
    });
    let json = build_display_list_json(&input.to_string()).expect("chart display list builds");
    serde_json::from_str(&json).unwrap()
}

#[test]
fn chart_column_block_emits_axes_bars_title_and_a11y_label() {
    let dl = chart_display_list("column");
    let prims = &dl.pages[0].primitives;
    let rects = rect_prims(prims);
    let texts = chart_text_prims(prims);
    let lines = line_prims(prims);

    assert!(
        rects.len() >= 4,
        "background, bars, and legend swatch should be rect primitives"
    );
    assert!(
        lines.len() >= 6,
        "column charts should emit axes and gridline primitives"
    );
    assert!(
        texts.iter().any(|t| t.text == "Revenue"),
        "chart title missing"
    );
    assert!(
        texts.iter().any(|t| t.text == "Q1") && texts.iter().any(|t| t.text == "Q2"),
        "category labels missing"
    );

    let chart_attrs = doc_attrs(&prims[0])
        .and_then(|a| a.chart.as_ref())
        .expect("chart primitives carry an accessibility label");
    assert_eq!(
        chart_attrs.label,
        "Revenue, column chart, 1 series, 2 categories"
    );
}

#[test]
fn chart_pie_block_emits_wedge_shapes_and_category_legend() {
    let dl = chart_display_list("pie");
    let prims = &dl.pages[0].primitives;
    let shapes = shape_prims(prims);
    let texts = chart_text_prims(prims);

    assert!(
        shapes.len() >= 2,
        "pie chart should emit one shape per wedge"
    );
    assert!(
        shapes.iter().all(|s| s
            .geometry_path
            .iter()
            .any(|cmd| matches!(cmd, ShapePathCommand::Close))),
        "pie wedges should be closed shape paths"
    );
    assert!(
        texts.iter().any(|t| t.text == "Share"),
        "pie chart title missing"
    );
    assert!(
        texts.iter().any(|t| t.text == "Q1") && texts.iter().any(|t| t.text == "Q2"),
        "pie legend should use category labels"
    );

    let label = shapes[0].attrs.chart.as_ref().map(|c| c.label.as_str());
    assert_eq!(label, Some("Share, pie chart, 1 series, 2 categories"));
}

/// live-pipeline inputs use compound STRING block ids (`block-N`); the
/// contract carries them as `blockKey` while numeric ids keep `blockId` —
/// identical geometry, only the identity field changes
#[test]
fn string_block_ids_carry_block_key() {
    let input = std::fs::read_to_string(fixture_path("single-page-multi-paragraph", "input"))
        .expect("input fixture");
    let mut v: serde_json::Value = serde_json::from_str(&input).unwrap();

    // rewrite every numeric block id to the live pipeline's `block-N` shape
    for mb in v["measured"].as_array_mut().unwrap() {
        let id = mb["block"]["id"].clone();
        mb["block"]["id"] = serde_json::json!(format!("block-{id}"));
    }
    for page in v["layout"]["pages"].as_array_mut().unwrap() {
        for frag in page["fragments"].as_array_mut().unwrap() {
            let id = frag["blockId"].clone();
            frag["blockId"] = serde_json::json!(format!("block-{id}"));
        }
    }

    let numeric_out = build_display_list_json(&input).unwrap();
    let string_out = build_display_list_json(&v.to_string()).unwrap();

    let numeric: DisplayList = serde_json::from_str(&numeric_out).unwrap();
    let string: DisplayList = serde_json::from_str(&string_out).unwrap();

    let mut checked = 0;
    for (np, sp) in numeric.pages[0]
        .primitives
        .iter()
        .zip(&string.pages[0].primitives)
    {
        let (Some(na), Some(sa)) = (doc_attrs(np), doc_attrs(sp)) else {
            continue;
        };
        let id = na.block_id.as_ref().expect("numeric build carries blockId");
        assert_eq!(na.block_key, None, "numeric build must not emit blockKey");
        assert_eq!(
            sa.block_key.as_deref(),
            Some(format!("block-{id}").as_str()),
            "string build carries the raw id as blockKey"
        );
        assert_eq!(sa.block_id, None, "string ids have no numeric identity");
        checked += 1;
    }
    assert!(checked > 0, "no attributed primitives compared");

    // identity fields are the ONLY delta: strip them and the builds match
    let strip = |json: &str| -> serde_json::Value {
        let mut v: serde_json::Value = serde_json::from_str(json).unwrap();
        fn walk(v: &mut serde_json::Value) {
            match v {
                serde_json::Value::Object(o) => {
                    o.remove("blockId");
                    o.remove("blockKey");
                    for x in o.values_mut() {
                        walk(x);
                    }
                }
                serde_json::Value::Array(a) => {
                    for x in a {
                        walk(x);
                    }
                }
                _ => {}
            }
        }
        walk(&mut v);
        v
    };
    assert_eq!(strip(&numeric_out), strip(&string_out));
}

/// compose the `{ measured, options, layout }` build input for a golden
/// scenario whose fixture pins only the paginator pair
fn compose_input(name: &str) -> String {
    let input = std::fs::read_to_string(fixture_path(name, "input")).expect("input fixture");
    let golden = std::fs::read_to_string(fixture_path(name, "golden")).expect("golden fixture");
    let mut v: serde_json::Value = serde_json::from_str(&input).unwrap();
    v["layout"] = serde_json::from_str(&golden).unwrap();
    v.to_string()
}

// vertically-merged-cell-continuation: table 60, colw [100,100]; cell (0,0)
// spans 3 rows in column 0, cells (0,1)/(1,0)/(2,0) land in column 1
#[test]
fn cell_structure_rides_on_table_cell_primitives() {
    let json = build_display_list_json(&compose_input("vertically-merged-cell-continuation"))
        .expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    assert_eq!(dl.pages.len(), 2);

    let cells_on = |page: usize| -> Vec<TableCellRef> {
        dl.pages[page]
            .primitives
            .iter()
            .filter_map(doc_attrs)
            .filter_map(|a| a.cell.clone())
            .collect()
    };

    let page0 = cells_on(0);
    assert!(!page0.is_empty(), "page 0 emits cell-attributed primitives");
    // the vmerge anchor: row 0, column 0, spanning 3 rows
    assert!(
        page0.iter().any(|c| c.row == 0
            && c.col == 0
            && c.row_span == 3
            && c.col_span == 1
            && c.continuation.is_none()),
        "anchor cell (0,0) rowSpan 3 missing: {page0:?}"
    );
    // vmerge pushes the row-1 and row-2 cells onto the column-1 grid slot
    for (row, col) in [(0, 1), (1, 1), (2, 1)] {
        assert!(
            page0
                .iter()
                .any(|c| c.row == row && c.col == col && c.row_span == 1),
            "cell ({row},{col}) missing on page 0: {page0:?}"
        );
    }

    // continuation page: the visible content is the vmerge cell's clipped
    // slice — every ref keeps the anchor's (0,0) grid slot, spans, and the
    // continuation flag (row 2's own one-line cell sits above the clip window)
    let page1 = cells_on(1);
    assert!(!page1.is_empty(), "continuation page emits cell refs");
    for c in &page1 {
        assert_eq!(
            (c.row, c.col, c.row_span, c.col_span, c.continuation),
            (0, 0, 3, 1, Some(true)),
            "continuation slice must keep the anchor cell's grid slot"
        );
    }
}

/// a vertically-merged cell whose span crosses the page break re-paints on
/// the continuation page flagged `continuation: true` with doc positions
/// stripped — the data-vmerge-continuation analogue
#[test]
fn vmerge_continuation_slice_is_flagged_and_unselectable() {
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "table", "id": 7, "rows": [
                { "cells": [
                    { "rowSpan": 2, "background": "#ffcc00",
                      "blocks": [{ "kind": "paragraph", "id": 700, "pmStart": 2, "pmEnd": 6,
                                   "runs": [{ "kind": "text", "text": "span", "pmStart": 2 }] }] },
                    { "blocks": [{ "kind": "paragraph", "id": 701, "pmStart": 8, "pmEnd": 10,
                                   "runs": [{ "kind": "text", "text": "b", "pmStart": 8 }] }] }
                ] },
                { "cells": [
                    { "blocks": [{ "kind": "paragraph", "id": 702, "pmStart": 12, "pmEnd": 14,
                                   "runs": [{ "kind": "text", "text": "c", "pmStart": 12 }] }] }
                ] }
            ] },
            "measure": { "kind": "table", "columnWidths": [100.0, 100.0], "totalHeight": 48.0,
                "rows": [
                    { "height": 24.0, "cells": [
                        { "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                            "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                                        "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] },
                        { "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                            "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                                        "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }
                    ] },
                    { "height": 24.0, "cells": [
                        { "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                            "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                                        "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }
                    ] }
                ] }
        }],
        "options": {},
        "layout": { "pages": [
            { "size": { "w": 400.0, "h": 200.0 }, "margins": {}, "fragments": [
                { "kind": "table", "blockId": 7, "x": 50.0, "y": 50.0, "width": 200.0,
                  "height": 24.0, "rowStart": 0, "rowEnd": 1, "carriedToNext": true } ] },
            { "size": { "w": 400.0, "h": 200.0 }, "margins": {}, "fragments": [
                { "kind": "table", "blockId": 7, "x": 50.0, "y": 50.0, "width": 200.0,
                  "height": 24.0, "rowStart": 1, "rowEnd": 2, "carriedFromPrev": true } ] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();

    // page 1: the anchor paints normally, no continuation flag anywhere
    let page0_flagged = dl.pages[0]
        .primitives
        .iter()
        .filter_map(doc_attrs)
        .filter_map(|a| a.cell.as_ref())
        .any(|c| c.continuation == Some(true));
    assert!(!page0_flagged, "anchor page must not flag continuation");

    // page 2: the re-painted vmerge slice keeps the anchor's row/col, is
    // flagged, and carries no doc positions (not selectable)
    let cont: Vec<&DocAttrs> = dl.pages[1]
        .primitives
        .iter()
        .filter_map(doc_attrs)
        .filter(|a| {
            a.cell
                .as_ref()
                .is_some_and(|c| c.continuation == Some(true))
        })
        .collect();
    assert!(!cont.is_empty(), "continuation slice emitted no primitives");
    for a in &cont {
        let c = a.cell.as_ref().unwrap();
        assert_eq!((c.row, c.col, c.row_span, c.col_span), (0, 0, 2, 1));
        assert_eq!(a.doc_start, None, "continuation slices are not selectable");
        assert_eq!(a.doc_end, None);
    }

    // the row the continuation page actually owns is selectable and unflagged
    assert!(
        dl.pages[1]
            .primitives
            .iter()
            .filter_map(doc_attrs)
            .any(|a| a.doc_start.is_some()
                && a.cell
                    .as_ref()
                    .is_some_and(|c| c.row == 1 && c.col == 1 && c.continuation.is_none())),
        "row-1 cell content missing on page 2"
    );
}

/// #188 — a floating image anchored INSIDE a table cell paints on the cell's
/// float layer: an Image primitive at its cell-content-relative geometry,
/// carrying the cell grid ref, with behind-doc floats under the cell text and
/// front floats over it (matching renderCellContent's layer order).
#[test]
fn cell_anchored_floating_images_paint_at_cell_relative_geometry() {
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "table", "id": 9, "rows": [
                { "cells": [
                    { "blocks": [{ "kind": "paragraph", "id": 900, "pmStart": 2, "pmEnd": 5, "runs": [
                        { "kind": "text", "text": "hi", "pmStart": 2, "pmEnd": 4 },
                        // flush-right square wrap (front layer)
                        { "kind": "image", "src": "rIdFront", "width": 40.0, "height": 30.0,
                          "wrapType": "square", "cssFloat": "right", "alt": "front logo",
                          "pmStart": 4, "pmEnd": 5 },
                        // top-left behind-doc float (behind layer)
                        { "kind": "image", "src": "rIdBehind", "width": 30.0, "height": 20.0,
                          "wrapType": "behind",
                          "position": { "horizontal": { "align": "left" },
                                        "vertical": { "align": "top" } },
                          "pmStart": 5, "pmEnd": 6 }
                    ] }] }
                ] }
            ] },
            "measure": { "kind": "table", "columnWidths": [200.0], "totalHeight": 100.0,
                "rows": [
                    { "height": 100.0, "cells": [
                        { "width": 200.0, "height": 24.0, "blocks": [{ "kind": "paragraph",
                            "totalHeight": 24.0,
                            "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 2,
                                        "width": 14.0, "ascent": 12.0, "descent": 4.0,
                                        "lineHeight": 24.0 }] }] }
                    ] }
                ] }
        }],
        "options": {},
        "layout": { "pages": [
            { "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
                { "kind": "table", "blockId": 9, "x": 50.0, "y": 50.0, "width": 200.0,
                  "height": 100.0, "rowStart": 0, "rowEnd": 1 } ] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    let prims = &dl.pages[0].primitives;

    let image = |rel: &str| {
        prims
            .iter()
            .find_map(|p| match p {
                Primitive::Image(i) if i.rel_id == rel => Some(i),
                _ => None,
            })
            .unwrap_or_else(|| panic!("cell float {rel} not emitted"))
    };
    let nf = |n: &serde_json::Number| n.as_f64().unwrap();

    // cell content box origin: cx(50) + padLeft(7) = 57 x, cy(50) + padTop(1) = 51 y
    // front float is flush-right: x = contentX + (contentWidth 186 - width 40) = 203
    let front = image("rIdFront");
    assert!(
        (nf(&front.x) - 203.0).abs() < 0.001,
        "front x {:?}",
        front.x
    );
    assert!((nf(&front.y) - 51.0).abs() < 0.001, "front y {:?}", front.y);
    assert!((nf(&front.w) - 40.0).abs() < 0.001);
    assert!((nf(&front.h) - 30.0).abs() < 0.001);
    assert_eq!(front.alt_text.as_deref(), Some("front logo"));
    let fc = front
        .attrs
        .cell
        .as_ref()
        .expect("front float carries a cell ref");
    assert_eq!((fc.row, fc.col, fc.col_span), (0, 0, 1));
    // a selectable cell keeps the run's doc positions
    assert_eq!(front.attrs.doc_start, Some(4));
    assert_eq!(front.attrs.doc_end, Some(5));

    // behind float sits at the cell content-box top-left
    let behind = image("rIdBehind");
    assert!(
        (nf(&behind.x) - 57.0).abs() < 0.001,
        "behind x {:?}",
        behind.x
    );
    assert!(
        (nf(&behind.y) - 51.0).abs() < 0.001,
        "behind y {:?}",
        behind.y
    );
    assert!(
        behind.attrs.cell.is_some(),
        "behind float carries a cell ref"
    );

    // paint order: behind float under the cell text, front float over it
    let pos = |pred: &dyn Fn(&Primitive) -> bool| prims.iter().position(pred).unwrap();
    let behind_idx = pos(&|p| matches!(p, Primitive::Image(i) if i.rel_id == "rIdBehind"));
    let front_idx = pos(&|p| matches!(p, Primitive::Image(i) if i.rel_id == "rIdFront"));
    let text_idx = pos(&|p| matches!(p, Primitive::Text(t) if t.text == "hi"));
    assert!(
        behind_idx < text_idx,
        "behind float must paint under the cell text"
    );
    assert!(
        front_idx > text_idx,
        "front float must paint over the cell text"
    );

    // inline text is unaffected — the floating runs never paint in the line flow
    assert_eq!(
        prims
            .iter()
            .filter(|p| matches!(p, Primitive::Text(_)))
            .count(),
        1,
        "the two floating image runs must not paint inline"
    );
}

/// the measured line window rides on paragraph fragments the mirror surfaces as
/// a paragraph wrapper (body/HF/text box) via `fromLine`/`toLine`; table-cell
/// paragraphs omit it (rendered as ARIA cells, no fragment node)
#[test]
fn paragraph_line_range_stamps_body_but_not_cells() {
    let input = serde_json::json!({
        "measured": [
            { "block": { "kind": "paragraph", "id": 1, "pmStart": 1, "pmEnd": 4,
                  "runs": [{ "kind": "text", "text": "body", "pmStart": 1, "pmEnd": 5 }] },
              "measure": { "kind": "paragraph", "totalHeight": 24.0,
                  "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                              "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] } },
            { "block": { "kind": "table", "id": 2, "rows": [
                  { "cells": [{ "blocks": [{ "kind": "paragraph", "id": 200, "pmStart": 8, "pmEnd": 10,
                      "runs": [{ "kind": "text", "text": "c", "pmStart": 8, "pmEnd": 9 }] }] }] } ] },
              "measure": { "kind": "table", "columnWidths": [100.0], "totalHeight": 24.0, "rows": [
                  { "height": 24.0, "cells": [{ "width": 100.0, "height": 24.0, "blocks": [
                      { "kind": "paragraph", "totalHeight": 24.0, "lines": [
                          { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                            "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }] } ] } }
        ],
        "options": {},
        "layout": { "pages": [
            { "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
                { "kind": "paragraph", "blockId": 1, "x": 50.0, "y": 50.0, "width": 200.0,
                  "height": 24.0, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 4 },
                { "kind": "table", "blockId": 2, "x": 50.0, "y": 80.0, "width": 100.0,
                  "height": 24.0, "rowStart": 0, "rowEnd": 1 }
            ] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();

    // body paragraph text carries the fragment's line window
    let body = dl.pages[0]
        .primitives
        .iter()
        .find_map(|p| match p {
            Primitive::Text(t) if t.text == "body" => Some(&t.attrs),
            _ => None,
        })
        .expect("body text emitted");
    assert_eq!(body.from_line, Some(0), "body fragment carries fromLine");
    assert_eq!(body.to_line, Some(1), "body fragment carries toLine");

    // table-cell paragraph text does NOT (surfaces as an ARIA cell)
    let cell = dl.pages[0]
        .primitives
        .iter()
        .find_map(|p| match p {
            Primitive::Text(t) if t.text == "c" => Some(&t.attrs),
            _ => None,
        })
        .expect("cell text emitted");
    assert_eq!(cell.from_line, None, "cell paragraphs omit the line range");
    assert_eq!(cell.to_line, None);
    assert!(cell.cell.is_some(), "cell text still carries its grid ref");
}

/// image alt text (`wp:docPr` descr → ImageRun/ImageBlock `alt`) threads onto
/// image primitives as `altText`, empty values drop, oversized values cap
#[test]
fn image_alt_text_threads_and_caps() {
    let long_alt: String = "a".repeat(MAX_ALT_TEXT_CHARS + 500);
    let input = serde_json::json!({
        "measured": [
            { "block": { "kind": "paragraph", "id": 1, "pmStart": 1, "pmEnd": 3, "runs": [
                  { "kind": "image", "src": "rId9", "width": 40.0, "height": 20.0,
                    "alt": "Chart of quarterly revenue", "pmStart": 1, "pmEnd": 2 } ] },
              "measure": { "kind": "paragraph", "totalHeight": 24.0,
                  "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                              "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] } },
            { "block": { "kind": "image", "id": 2, "src": "rId10", "alt": long_alt },
              "measure": { "kind": "image", "width": 60.0, "height": 30.0 } },
            { "block": { "kind": "image", "id": 3, "src": "rId11", "alt": "" },
              "measure": { "kind": "image", "width": 10.0, "height": 10.0 } }
        ],
        "options": {},
        "layout": { "pages": [
            { "size": { "w": 400.0, "h": 200.0 }, "margins": {}, "fragments": [
                { "kind": "paragraph", "blockId": 1, "x": 50.0, "y": 50.0, "width": 200.0,
                  "height": 24.0, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 3 },
                { "kind": "image", "blockId": 2, "x": 50.0, "y": 80.0, "width": 60.0, "height": 30.0 },
                { "kind": "image", "blockId": 3, "x": 50.0, "y": 120.0, "width": 10.0, "height": 10.0 }
            ] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();

    let images: Vec<_> = dl.pages[0]
        .primitives
        .iter()
        .filter_map(|p| match p {
            Primitive::Image(i) => Some(i),
            _ => None,
        })
        .collect();
    assert_eq!(images.len(), 3);

    let inline = images.iter().find(|i| i.rel_id == "rId9").unwrap();
    assert_eq!(
        inline.alt_text.as_deref(),
        Some("Chart of quarterly revenue")
    );

    let capped = images.iter().find(|i| i.rel_id == "rId10").unwrap();
    assert_eq!(
        capped.alt_text.as_ref().map(|s| s.chars().count()),
        Some(MAX_ALT_TEXT_CHARS),
        "oversized alt text must cap at MAX_ALT_TEXT_CHARS"
    );

    let empty = images.iter().find(|i| i.rel_id == "rId11").unwrap();
    assert_eq!(empty.alt_text, None, "empty alt drops like the DOM painter");
}

#[test]
fn text_watermark_emits_rotated_translucent_text_primitive() {
    let input = serde_json::json!({
        "measured": [],
        "options": {},
        "headersFooters": {
            "variants": [],
            "watermark": {
                "kind": "text",
                "text": "DRAFT",
                "font": "Calibri",
                "color": "#C0C0C0",
                "semitransparent": true,
                "layout": "diagonal"
            }
        },
        "layout": { "pages": [
            { "number": 1, "size": { "w": 816.0, "h": 1056.0 },
              "margins": { "top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0 },
              "fragments": [] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    assert_eq!(dl.pages[0].primitives.len(), 1);
    let Primitive::Text(t) = &dl.pages[0].primitives[0] else {
        panic!("watermark should emit text");
    };
    assert_eq!(t.text, "DRAFT");
    assert_eq!(t.color, "#C0C0C0");
    assert_eq!(t.font, "700 180px Calibri, sans-serif");
    assert_eq!(t.opacity.as_ref().and_then(|n| n.as_f64()), Some(0.5));
    assert_eq!(
        t.rotation_deg.as_ref().and_then(|n| n.as_f64()),
        Some(-45.0)
    );
    assert_eq!(t.x.as_f64(), Some(129.0));
    assert_eq!(t.baseline_y.as_f64(), Some(564.0));
    assert_eq!(t.width.as_f64(), Some(558.0));
    assert_eq!(t.attrs.doc_start, None);
}

#[test]
fn picture_watermark_emits_decorative_washout_image_primitive() {
    let input = serde_json::json!({
        "measured": [],
        "options": {},
        "headersFooters": {
            "variants": [],
            "watermark": {
                "kind": "picture",
                "dataUrl": "data:image/png;base64,abc",
                "relId": "rId9",
                "scale": 0.5,
                "washout": true,
                "widthEmu": 1828800,
                "heightEmu": 914400
            }
        },
        "layout": { "pages": [
            { "number": 1, "size": { "w": 816.0, "h": 1056.0 },
              "margins": { "top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0 },
              "fragments": [] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    assert_eq!(dl.pages[0].primitives.len(), 1);
    let Primitive::Image(img) = &dl.pages[0].primitives[0] else {
        panic!("watermark should emit image");
    };
    assert_eq!(img.rel_id, "data:image/png;base64,abc");
    assert_eq!(img.x.as_f64(), Some(360.0));
    assert_eq!(img.y.as_f64(), Some(504.0));
    assert_eq!(img.w.as_f64(), Some(96.0));
    assert_eq!(img.h.as_f64(), Some(48.0));
    assert_eq!(img.opacity.as_ref().and_then(|n| n.as_f64()), Some(0.5));
    assert_eq!(img.filter.as_deref(), Some("brightness(1.4) contrast(0.4)"));
    assert!(img.decorative);
    assert_eq!(img.alt_text, None);
}

/// tracked-change runs paint like the DOM painter (renderParagraph/runs.ts):
/// an insertion gets a green wash behind the glyphs plus a green *dashed*
/// underline, a deletion gets a red wash plus a strike (the red text itself
/// comes from run_color). Guards the canvas revision styling that the display
/// list forgot for insertions.
#[test]
fn tracked_change_runs_emit_revision_washes_and_rules() {
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "paragraph", "id": 1, "pmStart": 1, "pmEnd": 18, "runs": [
                { "kind": "text", "text": "inserted", "pmStart": 1, "pmEnd": 9,
                  "isInsertion": true, "changeAuthor": "A", "changeDate": "2026", "changeRevisionId": 5 },
                { "kind": "text", "text": "deleted", "pmStart": 9, "pmEnd": 16,
                  "isDeletion": true, "changeAuthor": "A", "changeDate": "2026", "changeRevisionId": 6 }
            ] },
            "measure": { "kind": "paragraph", "totalHeight": 24.0, "lines": [
                { "headRun": 0, "headChar": 0, "tailRun": 1, "tailChar": 7,
                  "width": 120.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }
            ] }
        }],
        "options": {},
        "layout": { "pages": [
            { "size": { "w": 400.0, "h": 200.0 }, "margins": {}, "fragments": [
                { "kind": "paragraph", "blockId": 1, "x": 50.0, "y": 50.0, "width": 200.0,
                  "height": 24.0, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 18 }
            ] }
        ] }
    });

    // determinism: the same input builds byte-identically
    let a = build_display_list_json(&input.to_string()).expect("builds");
    let b = build_display_list_json(&input.to_string()).expect("builds");
    assert_eq!(a, b, "tracked-change build is not deterministic");

    let dl: DisplayList = serde_json::from_str(&a).unwrap();
    let decos: Vec<&DecorationPrimitive> = dl.pages[0]
        .primitives
        .iter()
        .filter_map(|p| match p {
            Primitive::Decoration(d) => Some(d),
            _ => None,
        })
        .collect();

    // insertion: green wash (behind glyphs) + green dashed underline
    let ins_wash = decos
        .iter()
        .find(|d| d.deco == DecoKind::Highlight && d.color == "rgba(52, 168, 83, 0.08)")
        .expect("insertion green wash decoration missing");
    assert!(!ins_wash.dashed, "a wash is never dashed");
    assert_eq!(
        ins_wash.attrs.revision.as_ref().map(|r| r.kind),
        Some(RevisionKind::Ins),
        "the wash must carry the run's insertion revision for hit-testing"
    );

    let ins_rule = decos
        .iter()
        .find(|d| d.deco == DecoKind::Underline && d.color == "#2e7d32")
        .expect("insertion green underline decoration missing");
    assert!(ins_rule.dashed, "insertion underline must be dashed");

    // deletion: red wash (the red text + strike carry the deletion otherwise)
    assert!(
        decos
            .iter()
            .any(|d| d.deco == DecoKind::Highlight && d.color == "rgba(211, 47, 47, 0.08)"),
        "deletion red wash decoration missing"
    );
    assert!(
        decos
            .iter()
            .any(|d| d.deco == DecoKind::Strike && d.color == "#c62828"),
        "deletion strike decoration missing"
    );

    // the washes sit behind their text run: each wash precedes the text
    // primitive carrying the same revision id
    let order: Vec<&Primitive> = dl.pages[0].primitives.iter().collect();
    let wash_idx = order
        .iter()
        .position(|p| matches!(p, Primitive::Decoration(d) if d.color == "rgba(52, 168, 83, 0.08)"))
        .unwrap();
    let text_idx = order
        .iter()
        .position(|p| matches!(p, Primitive::Text(t) if t.text == "inserted"))
        .unwrap();
    assert!(
        wash_idx < text_idx,
        "insertion wash must paint behind the glyphs"
    );
}

#[derive(Debug, PartialEq)]
struct StructuralPin {
    visual: &'static str,
    scope: &'static str,
    kind: &'static str,
    revision_id: String,
    row_index: Option<u64>,
    col_index: Option<u64>,
    text: Option<String>,
    fill: Option<String>,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
}

fn scope_name(scope: StructuralRevisionScope) -> &'static str {
    match scope {
        StructuralRevisionScope::Pmark => "pmark",
        StructuralRevisionScope::Table => "table",
        StructuralRevisionScope::Row => "row",
        StructuralRevisionScope::Cell => "cell",
    }
}

fn structural_kind_name(kind: StructuralRevisionKind) -> &'static str {
    match kind {
        StructuralRevisionKind::Ins => "ins",
        StructuralRevisionKind::Del => "del",
        StructuralRevisionKind::Merge => "merge",
    }
}

fn structural_pin(p: &Primitive) -> Option<StructuralPin> {
    let attrs = doc_attrs(p)?;
    let rev = attrs.structural_revision.as_ref()?;
    match p {
        Primitive::Rect(r) => Some(StructuralPin {
            visual: "rect",
            scope: scope_name(rev.scope),
            kind: structural_kind_name(rev.kind),
            revision_id: rev.revision_id.clone(),
            row_index: rev.row_index,
            col_index: rev.col_index,
            text: None,
            fill: Some(r.fill.clone()),
            x: r.x.as_f64().unwrap(),
            y: r.y.as_f64().unwrap(),
            w: r.w.as_f64().unwrap(),
            h: r.h.as_f64().unwrap(),
        }),
        Primitive::Text(t) => Some(StructuralPin {
            visual: "text",
            scope: scope_name(rev.scope),
            kind: structural_kind_name(rev.kind),
            revision_id: rev.revision_id.clone(),
            row_index: rev.row_index,
            col_index: rev.col_index,
            text: Some(t.text.clone()),
            fill: Some(t.color.clone()),
            x: t.x.as_f64().unwrap(),
            y: t.baseline_y.as_f64().unwrap(),
            w: t.width.as_f64().unwrap(),
            h: 0.0,
        }),
        _ => None,
    }
}

/// Structural tracked changes are display-list primitives, not CSS-only DOM
/// side effects: paragraph-mark changes emit a margin bar plus pilcrow, row and
/// whole-table changes emit the painter's left-margin bar, and cell changes
/// emit the top inset marker. This pins the primitives that canvas and mirror
/// consume for review/sidebar workflows.
#[test]
fn structural_revisions_emit_pinned_primitives() {
    let para = serde_json::json!({
        "block": { "kind": "paragraph", "id": 1, "pmStart": 1, "pmEnd": 2,
            "attrs": { "pPrIns": { "revisionId": 10, "author": "Ada", "date": "2026-01-01T00:00:00Z" } },
            "runs": [{ "kind": "text", "text": "A", "pmStart": 1, "pmEnd": 2 }] },
        "measure": { "kind": "paragraph", "totalHeight": 24.0,
            "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                        "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }
    });
    let structural_table = serde_json::json!({
        "block": { "kind": "table", "id": 2, "rows": [
            { "trackedIns": { "revisionId": 20, "author": "Ada", "date": "2026-01-02T00:00:00Z" },
              "cells": [{ "blocks": [{ "kind": "paragraph", "id": 20, "pmStart": 10, "pmEnd": 11,
                  "runs": [{ "kind": "text", "text": "R", "pmStart": 10, "pmEnd": 11 }] }] }] },
            { "cells": [{ "trackedMarker": { "kind": "del",
                    "info": { "revisionId": 21, "author": "Ada", "date": "2026-01-03T00:00:00Z" } },
                "blocks": [{ "kind": "paragraph", "id": 21, "pmStart": 12, "pmEnd": 13,
                  "runs": [{ "kind": "text", "text": "C", "pmStart": 12, "pmEnd": 13 }] }] }] }
        ] },
        "measure": { "kind": "table", "columnWidths": [80.0], "totalHeight": 48.0,
            "rows": [
                { "height": 24.0, "cells": [{ "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                    "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                                "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }] },
                { "height": 24.0, "cells": [{ "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                    "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                                "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }] }
            ] }
    });
    let whole_table = serde_json::json!({
        "block": { "kind": "table", "id": 3, "rows": [
            { "trackedDel": { "revisionId": 30, "author": "Ada", "date": "2026-01-04T00:00:00Z" },
              "cells": [{ "blocks": [{ "kind": "paragraph", "id": 30, "pmStart": 20, "pmEnd": 21,
                  "runs": [{ "kind": "text", "text": "T", "pmStart": 20, "pmEnd": 21 }] }] }] },
            { "trackedDel": { "revisionId": 31, "author": "Ada", "date": "2026-01-04T00:00:00Z" },
              "cells": [{ "blocks": [{ "kind": "paragraph", "id": 31, "pmStart": 22, "pmEnd": 23,
                  "runs": [{ "kind": "text", "text": "U", "pmStart": 22, "pmEnd": 23 }] }] }] }
        ] },
        "measure": { "kind": "table", "columnWidths": [80.0], "totalHeight": 48.0,
            "rows": [
                { "height": 24.0, "cells": [{ "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                    "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                                "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }] },
                { "height": 24.0, "cells": [{ "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                    "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                                "width": 10.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }] }
            ] }
    });
    let input = serde_json::json!({
        "measured": [para, structural_table, whole_table],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
            { "kind": "paragraph", "blockId": 1, "x": 50.0, "y": 50.0, "width": 100.0,
              "height": 24.0, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 2 },
            { "kind": "table", "blockId": 2, "x": 50.0, "y": 90.0, "width": 80.0,
              "height": 48.0, "rowStart": 0, "rowEnd": 2 },
            { "kind": "table", "blockId": 3, "x": 50.0, "y": 150.0, "width": 80.0,
              "height": 48.0, "rowStart": 0, "rowEnd": 2 }
        ] }] }
    });

    let a = build_display_list_json(&input.to_string()).expect("builds");
    let b = build_display_list_json(&input.to_string()).expect("builds");
    assert_eq!(
        a, b,
        "structural revision display list is not deterministic"
    );

    let dl: DisplayList = serde_json::from_str(&a).unwrap();
    let pins: Vec<StructuralPin> = dl.pages[0]
        .primitives
        .iter()
        .filter_map(structural_pin)
        .collect();
    assert_eq!(
        pins,
        vec![
            StructuralPin {
                visual: "rect",
                scope: "pmark",
                kind: "ins",
                revision_id: "10".to_string(),
                row_index: None,
                col_index: None,
                text: None,
                fill: Some("#2e7d32".to_string()),
                x: 40.0,
                y: 50.0,
                w: 2.0,
                h: 24.0,
            },
            StructuralPin {
                visual: "text",
                scope: "pmark",
                kind: "ins",
                revision_id: "10".to_string(),
                row_index: None,
                col_index: None,
                text: Some("¶".to_string()),
                fill: Some("#2e7d32".to_string()),
                x: 62.0,
                y: 66.0,
                w: 8.0,
                h: 0.0,
            },
            StructuralPin {
                visual: "rect",
                scope: "row",
                kind: "ins",
                revision_id: "20".to_string(),
                row_index: Some(0),
                col_index: None,
                text: None,
                fill: Some("#2e7d32".to_string()),
                x: 40.0,
                y: 90.0,
                w: 2.0,
                h: 24.0,
            },
            StructuralPin {
                visual: "rect",
                scope: "cell",
                kind: "del",
                revision_id: "21".to_string(),
                row_index: Some(1),
                col_index: Some(0),
                text: None,
                fill: Some("#c62828".to_string()),
                x: 50.0,
                y: 114.0,
                w: 80.0,
                h: 3.0,
            },
            StructuralPin {
                visual: "rect",
                scope: "table",
                kind: "del",
                revision_id: "30".to_string(),
                row_index: None,
                col_index: None,
                text: None,
                fill: Some("#c62828".to_string()),
                x: 40.0,
                y: 150.0,
                w: 2.0,
                h: 48.0,
            },
        ]
    );
}

// ---------------------------------------------------------------------------
// watchdog-finding geometry (F1/F4/F5/F6): hand-computed origins pinning each
// fix against the DOM painter's placement. Numbers are derived in the comments
// so a regression names the exact term that drifted.
// ---------------------------------------------------------------------------

fn build_dl(input: &str) -> DisplayList {
    let json = build_display_list_json(input).expect("builds");
    serde_json::from_str(&json).unwrap()
}

/// (text, x, width, baselineY, wordSpacing) for every ink-bearing text primitive
fn text_prims(prims: &[Primitive]) -> Vec<(String, f64, f64, f64, Option<f64>)> {
    prims
        .iter()
        .filter_map(|p| match p {
            Primitive::Text(t) if !t.text.is_empty() => Some((
                t.text.clone(),
                t.x.as_f64().unwrap(),
                t.width.as_f64().unwrap(),
                t.baseline_y.as_f64().unwrap(),
                t.word_spacing.as_ref().and_then(|n| n.as_f64()),
            )),
            _ => None,
        })
        .collect()
}

/// F4: a hanging-indent list paragraph's marker fills the hang, so BOTH the
/// first line and the body lines start their body text at the text indent
/// (max(indent_left, hanging)) — NOT `indent_left - hanging` like a plain
/// hanging paragraph. Mirrors renderParagraph.ts's marker inline-block of
/// min-width = hanging.
#[test]
fn hanging_list_first_line_body_sits_at_the_text_indent() {
    // indent_left 40, hanging 20, frag.x 100 → body text at 100 + 40 = 140 on
    // every line (the 20px hang holds the — unemitted — marker)
    let with_marker = |marker: serde_json::Value| {
        serde_json::json!({
            "measured": [{
                "block": { "kind": "paragraph", "id": 1, "pmStart": 1, "pmEnd": 30,
                    "attrs": { "indent": { "left": 40.0, "hanging": 20.0 }, "listMarker": marker },
                    "runs": [{ "kind": "text", "text": "alpha bravo", "pmStart": 1, "pmEnd": 12 }] },
                "measure": { "kind": "paragraph", "totalHeight": 48.0, "lines": [
                    { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 6,
                      "width": 60.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 },
                    { "headRun": 0, "headChar": 6, "tailRun": 0, "tailChar": 11,
                      "width": 50.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }
                ] }
            }],
            "options": {},
            "layout": { "pages": [{ "size": { "w": 700.0, "h": 400.0 }, "margins": {}, "fragments": [
                { "kind": "paragraph", "blockId": 1, "x": 100.0, "y": 50.0, "width": 500.0,
                  "height": 48.0, "fromLine": 0, "toLine": 2, "pmStart": 1, "pmEnd": 30 }
            ] }] }
        })
        .to_string()
    };

    let dl = build_dl(&with_marker(serde_json::json!("\u{2022}")));
    let xs: Vec<f64> = dl.pages[0]
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Text(text)
                if !text.text.is_empty() && text.attrs.list_marker != Some(true) =>
            {
                Some(text.x.as_f64().unwrap())
            }
            _ => None,
        })
        .collect();
    assert_eq!(xs.len(), 2, "two body lines emitted");
    assert!(
        (xs[0] - 140.0).abs() < 0.01,
        "first line x {} (want 140)",
        xs[0]
    );
    assert!(
        (xs[1] - 140.0).abs() < 0.01,
        "body line x {} (want 140)",
        xs[1]
    );

    // no marker: a plain hanging paragraph pulls the first line left by the hang
    // (100 + 40 - 20 = 120), body lines still at 140 — unchanged v0 behavior
    let dl2 = build_dl(&with_marker(serde_json::Value::Null));
    let xs2: Vec<f64> = text_prims(&dl2.pages[0].primitives)
        .iter()
        .map(|t| t.1)
        .collect();
    assert!(
        (xs2[0] - 120.0).abs() < 0.01,
        "plain first line x {} (want 120)",
        xs2[0]
    );
    assert!(
        (xs2[1] - 140.0).abs() < 0.01,
        "plain body line x {} (want 140)",
        xs2[1]
    );
}

/// F6: `jc=both` stretches expandable spaces to fill the usable width. A
/// non-last line reaches the right margin and carries the per-space add as
/// `wordSpacing`; the paragraph's closing line stays at natural width.
#[test]
fn justified_lines_stretch_to_the_usable_width() {
    // frag.x 96, width 600, no indent → usable width 600. Line 0 "a b c"
    // (natural 40, 2 spaces): slack 560, share 280, stretched width 600 →
    // right edge 96 + 600 = 696. Line 1 "d e f" is the last line → no stretch.
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "paragraph", "id": 1, "pmStart": 1, "pmEnd": 13,
                "attrs": { "alignment": "justify" },
                "runs": [{ "kind": "text", "text": "a b c d e f", "pmStart": 1, "pmEnd": 12 }] },
            "measure": { "kind": "paragraph", "totalHeight": 48.0, "lines": [
                { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 5,
                  "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 },
                { "headRun": 0, "headChar": 6, "tailRun": 0, "tailChar": 11,
                  "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }
            ] }
        }],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 800.0, "h": 400.0 }, "margins": {}, "fragments": [
            { "kind": "paragraph", "blockId": 1, "x": 96.0, "y": 50.0, "width": 600.0,
              "height": 48.0, "fromLine": 0, "toLine": 2, "pmStart": 1, "pmEnd": 13 }
        ] }] }
    })
    .to_string();

    let dl = build_dl(&input);
    let t = text_prims(&dl.pages[0].primitives);
    assert_eq!(t.len(), 2);

    let (l0_text, l0_x, l0_w, _, l0_ws) = &t[0];
    assert_eq!(l0_text, "a b c");
    assert!((l0_x - 96.0).abs() < 0.01, "line0 x {}", l0_x);
    assert!(
        (l0_w - 600.0).abs() < 0.01,
        "line0 stretched width {} (want 600)",
        l0_w
    );
    assert!(
        (l0_x + l0_w - 696.0).abs() < 0.01,
        "line0 right edge {} (want 696)",
        l0_x + l0_w
    );
    assert_eq!(*l0_ws, Some(280.0), "line0 per-space wordSpacing");

    let (l1_text, l1_x, l1_w, _, l1_ws) = &t[1];
    assert_eq!(l1_text, "d e f");
    assert!((l1_x - 96.0).abs() < 0.01, "line1 x {}", l1_x);
    assert!(
        (l1_w - 40.0).abs() < 0.01,
        "last line stays natural width {} (want 40)",
        l1_w
    );
    assert_eq!(*l1_ws, None, "the last line is not justified");
}

/// F6b: a paragraph whose final run is a `<w:br>` justifies its closing line
/// too (renderParagraph/line.ts `paragraphEndsWithLineBreak`).
#[test]
fn soft_return_justifies_the_final_line() {
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "paragraph", "id": 1, "pmStart": 1, "pmEnd": 13,
                "attrs": { "alignment": "justify" },
                "runs": [
                    { "kind": "text", "text": "a b c", "pmStart": 1, "pmEnd": 6 },
                    { "kind": "lineBreak", "pmStart": 6, "pmEnd": 7 }
                ] },
            "measure": { "kind": "paragraph", "totalHeight": 24.0, "lines": [
                { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 5,
                  "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }
            ] }
        }],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 800.0, "h": 400.0 }, "margins": {}, "fragments": [
            { "kind": "paragraph", "blockId": 1, "x": 96.0, "y": 50.0, "width": 600.0,
              "height": 24.0, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 13 }
        ] }] }
    })
    .to_string();

    let dl = build_dl(&input);
    let t = text_prims(&dl.pages[0].primitives);
    assert_eq!(t.len(), 1);
    // the only line is the last, but the trailing <w:br> makes it justify
    assert!(
        (t[0].2 - 600.0).abs() < 0.01,
        "soft-return line width {} (want 600)",
        t[0].2
    );
    assert_eq!(t[0].4, Some(280.0), "soft-return line carries wordSpacing");
}

/// F1 + F5: a table cell insets its content by the box-sizing left-border width
/// (leftmost column only) and offsets it vertically per w:vAlign. Constructed
/// so the two effects are isolable: the left cell has a 1px left border and
/// bottom vAlign, the right cell has neither.
#[test]
fn table_cell_content_insets_border_and_honors_valign() {
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "table", "id": 7, "rows": [{ "cells": [
                { "verticalAlign": "bottom",
                  "padding": { "top": 1.0, "bottom": 1.0, "left": 7.0, "right": 7.0 },
                  "borders": { "left": { "style": "single", "width": 1.0, "color": "#000000" } },
                  "blocks": [{ "kind": "paragraph", "id": 70, "pmStart": 2, "pmEnd": 5,
                               "runs": [{ "kind": "text", "text": "sig", "pmStart": 2 }] }] },
                { "padding": { "top": 1.0, "bottom": 1.0, "left": 7.0, "right": 7.0 },
                  "blocks": [{ "kind": "paragraph", "id": 71, "pmStart": 7, "pmEnd": 11,
                               "runs": [{ "kind": "text", "text": "date", "pmStart": 7 }] }] }
            ] }] },
            "measure": { "kind": "table", "columnWidths": [100.0, 100.0], "totalHeight": 60.0,
                "rows": [{ "height": 60.0, "cells": [
                    { "height": 20.0, "blocks": [{ "kind": "paragraph", "totalHeight": 20.0,
                        "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 3,
                                    "width": 20.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 20.0 }] }] },
                    { "height": 20.0, "blocks": [{ "kind": "paragraph", "totalHeight": 20.0,
                        "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                                    "width": 24.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 20.0 }] }] }
                ] }] }
        }],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 400.0, "h": 200.0 }, "margins": {}, "fragments": [
            { "kind": "table", "blockId": 7, "x": 50.0, "y": 50.0, "width": 200.0,
              "height": 60.0, "rowStart": 0, "rowEnd": 1 }
        ] }] }
    })
    .to_string();

    let dl = build_dl(&input);
    let t = text_prims(&dl.pages[0].primitives);
    let sig = t.iter().find(|x| x.0 == "sig").expect("sig cell text");
    let date = t.iter().find(|x| x.0 == "date").expect("date cell text");

    // left cell: cx 50 + left-border 1 + padLeft 7 = 58 (F1). vAlign bottom:
    // avail = 60 - 1 - 1 = 58, content 20 → offset 38; content top =
    // 50 + padTop 1 + 38 = 89; baseline = 89 + half-leading 2 + ascent 12 = 103.
    assert!(
        (sig.1 - 58.0).abs() < 0.01,
        "sig x {} (want 58: cx+border+pad)",
        sig.1
    );
    assert!(
        (sig.3 - 103.0).abs() < 0.01,
        "sig baseline {} (want 103: bottom vAlign)",
        sig.3
    );

    // right cell: cx 150 + no border + padLeft 7 = 157; top-anchored →
    // content top 50 + 1 = 51; baseline = 51 + 2 + 12 = 65.
    assert!(
        (date.1 - 157.0).abs() < 0.01,
        "date x {} (want 157: no left border)",
        date.1
    );
    assert!(
        (date.3 - 65.0).abs() < 0.01,
        "date baseline {} (want 65: top-anchored)",
        date.3
    );
}

/// F5 (demo residual): cell paragraphs stack with Word's max-collapse spacing
/// (the gap above a paragraph is `max(prev.after, this.before)`), exactly like
/// the DOM painter's renderCellContent. The demo's signature/date table carries
/// NO w:vAlign — its "bottom-aligned" look is purely the 320-twip after-spacing
/// on the second paragraph. Summing bare line heights (the old path) dropped
/// every inter-paragraph gap and painted the last line one after-spacing block
/// too high. Three single-line paragraphs, no vAlign/borders, isolate the two
/// collapse directions.
#[test]
fn cell_paragraphs_stack_with_collapsed_spacing() {
    // a: after 10; b: before 25 (wins over prev after 10) + after 30; c: before
    // 5 (loses to prev after 30). Line box 20 (ascent 12, descent 4 → leading 2,
    // baseline = lineTop + 14). content top = cy 50 + padTop 1 = 51.
    let para = |id: u64, ch: &str, before: f64, after: f64, pm: i64| {
        serde_json::json!({
            "kind": "paragraph", "id": id, "pmStart": pm, "pmEnd": pm + 3,
            "attrs": { "spacing": { "before": before, "after": after } },
            "runs": [{ "kind": "text", "text": ch, "pmStart": pm }]
        })
    };
    let pmeasure = || {
        serde_json::json!({ "kind": "paragraph", "totalHeight": 20.0, "lines": [
            { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
              "width": 8.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 20.0 }] })
    };
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "table", "id": 9, "rows": [{ "cells": [
                { "padding": { "top": 1.0, "bottom": 1.0, "left": 7.0, "right": 7.0 },
                  "blocks": [
                    para(90, "a", 0.0, 10.0, 2),
                    para(91, "b", 25.0, 30.0, 6),
                    para(92, "c", 5.0, 0.0, 10)
                  ] }
            ] }] },
            "measure": { "kind": "table", "columnWidths": [100.0], "totalHeight": 200.0,
                "rows": [{ "height": 200.0, "cells": [
                    { "height": 117.0, "blocks": [pmeasure(), pmeasure(), pmeasure()] }
                ] }] }
        }],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
            { "kind": "table", "blockId": 9, "x": 50.0, "y": 50.0, "width": 100.0,
              "height": 200.0, "rowStart": 0, "rowEnd": 1 }
        ] }] }
    })
    .to_string();

    let dl = build_dl(&input);
    let t = text_prims(&dl.pages[0].primitives);
    let base = |ch: &str| {
        t.iter()
            .find(|x| x.0 == ch)
            .unwrap_or_else(|| panic!("{ch} cell text"))
            .3
    };

    // a: top 0 → baseline 51 + 14 = 65.
    assert!(
        (base("a") - 65.0).abs() < 0.01,
        "a baseline {} (want 65)",
        base("a")
    );
    // b: gap max(after 10, before 25) = 25 → top 20 + 25 = 45 → baseline
    // 51 + 45 + 14 = 110 (before wins the collapse).
    assert!(
        (base("b") - 110.0).abs() < 0.01,
        "b baseline {} (want 110)",
        base("b")
    );
    // c: gap max(after 30, before 5) = 30 → top 45 + 20 + 30 = 95 → baseline
    // 51 + 95 + 14 = 160 (after wins). The old line-height-only stack put it at
    // 51 + 40 + 14 = 105, one after-spacing block too high (the F5 residual).
    assert!(
        (base("c") - 160.0).abs() < 0.01,
        "c baseline {} (want 160)",
        base("c")
    );
}

/// F2: a centered footer line that carries a PAGE field re-centers per page from
/// the supplied per-page field widths, instead of holding the once-measured
/// (fallback "1") position on every page. Covers both a per-digit glyph-width
/// change (page 8 → 9) and a char-count jump (page 9 → 10), and pins that the
/// field renders at its resolved width, not the char-distributed fallback.
#[test]
fn centered_hf_field_line_recenters_per_page() {
    // widths of the PAGE field's resolved text per page index (0 = page 1):
    // pages 1..8 identical, page 9 a wider single glyph, page 10 two digits.
    let per_page = [10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 14.0, 26.0];

    let pages: Vec<serde_json::Value> = (0..per_page.len())
        .map(|i| {
            serde_json::json!({
                "size": { "w": 200.0, "h": 300.0 },
                "margins": { "left": 0.0, "right": 0.0, "footer": 20.0 },
                "number": i + 1,
                "fragments": []
            })
        })
        .collect();

    // footer paragraph: "Page " + PAGE field, centered. Measured line.width 50 =
    // text "Page " + fallback "1" (10); pool = 50 - 10 = 40 over the 5 text chars.
    let footer_variant = serde_json::json!({
        "rId": "rId9", "kind": "footer", "type": "default", "height": 20.0,
        "measured": [{
            "block": { "kind": "paragraph", "id": 1, "pmStart": 0, "pmEnd": 20,
                "attrs": { "alignment": "center" },
                "runs": [
                    { "kind": "text", "text": "Page ", "pmStart": 1 },
                    { "kind": "field", "fieldType": "PAGE", "fallback": "1", "pmStart": 6 }
                ] },
            "measure": { "kind": "paragraph", "totalHeight": 20.0, "lines": [
                { "headRun": 0, "headChar": 0, "tailRun": 1, "tailChar": 1,
                  "width": 50.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 20.0 }] }
        }],
        "fieldWidths": [
            { "pmStart": 6, "fallbackWidth": 10.0, "perPage": per_page.to_vec() }
        ]
    });

    let input = serde_json::json!({
        "measured": [], "options": {},
        "layout": { "pages": pages },
        "headersFooters": { "variants": [footer_variant] }
    })
    .to_string();

    let dl = build_dl(&input);

    // left edge of "Page " = align_shift = (usable 200 - (pool 40 + resolved))/2.
    // The line re-centers as the resolved digit widens: pages 8/9/10 → 75/73/67.
    // (Without the fix align_shift is (200 - line.width 50)/2 = 75 on every page.)
    let left_of = |page_idx: usize| -> f64 {
        let footer = dl.pages[page_idx].footer.as_ref().expect("footer region");
        text_prims(&footer.primitives)
            .iter()
            .find(|x| x.0 == "Page ")
            .expect("Page text")
            .1
    };
    assert!(
        (left_of(7) - 75.0).abs() < 0.01,
        "page 8 left {} (want 75)",
        left_of(7)
    );
    assert!(
        (left_of(8) - 73.0).abs() < 0.01,
        "page 9 left {} (want 73: wider glyph)",
        left_of(8)
    );
    assert!(
        (left_of(9) - 67.0).abs() < 0.01,
        "page 10 left {} (want 67: two digits)",
        left_of(9)
    );

    // the field renders at its supplied resolved width (char-count-independent):
    // page 10's "10" is 26 wide, page 8's "8" is 10.
    let field_w = |page_idx: usize, text: &str| -> f64 {
        let footer = dl.pages[page_idx].footer.as_ref().unwrap();
        text_prims(&footer.primitives)
            .iter()
            .find(|x| x.0 == text)
            .unwrap_or_else(|| panic!("field text {text}"))
            .2
    };
    assert!(
        (field_w(7, "8") - 10.0).abs() < 0.01,
        "page 8 field width {}",
        field_w(7, "8")
    );
    assert!(
        (field_w(9, "10") - 26.0).abs() < 0.01,
        "page 10 field width {}",
        field_w(9, "10")
    );
}

/// a paragraph's stable `paraId` is stamped on every primitive it emits (the
/// a11y mirror reads it to expose `data-para-id`); a paragraph without one emits
/// no `paraId` field, keeping the wire form byte-identical to before.
#[test]
fn paragraph_para_id_stamps_on_primitives() {
    let input = |para_id: serde_json::Value| {
        serde_json::json!({
            "measured": [{
                "block": { "kind": "paragraph", "id": 1, "paraId": para_id,
                    "pmStart": 1, "pmEnd": 8,
                    "runs": [{ "kind": "text", "text": "hello", "pmStart": 1, "pmEnd": 6 }] },
                "measure": { "kind": "paragraph", "totalHeight": 24.0, "lines": [
                    { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 5,
                      "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }
            }],
            "options": {},
            "layout": { "pages": [{ "size": { "w": 400.0, "h": 200.0 }, "margins": {}, "fragments": [
                { "kind": "paragraph", "blockId": 1, "x": 50.0, "y": 50.0, "width": 200.0,
                  "height": 24.0, "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 8 }
            ] }] }
        })
        .to_string()
    };

    let dl: DisplayList =
        serde_json::from_str(&build_display_list_json(&input(serde_json::json!("4F2A"))).unwrap())
            .unwrap();
    let text = dl.pages[0]
        .primitives
        .iter()
        .find_map(|p| match p {
            Primitive::Text(t) if t.text == "hello" => Some(&t.attrs),
            _ => None,
        })
        .expect("text primitive");
    assert_eq!(text.para_id.as_deref(), Some("4F2A"));

    // no paraId in the source ⇒ the field is absent from the wire form
    let json = build_display_list_json(&input(serde_json::Value::Null)).unwrap();
    assert!(!json.contains("paraId"), "absent paraId must not serialize");
}

// ---------------------------------------------------------------------------
// text boxes + floating image runs (SILENT-CONTENT-LOSS gaps)
// ---------------------------------------------------------------------------

fn rect_prims(prims: &[Primitive]) -> Vec<&docx_layout::display_list::RectPrimitive> {
    prims
        .iter()
        .filter_map(|p| match p {
            Primitive::Rect(r) => Some(r),
            _ => None,
        })
        .collect()
}

fn image_prims(prims: &[Primitive]) -> Vec<&docx_layout::display_list::ImagePrimitive> {
    prims
        .iter()
        .filter_map(|p| match p {
            Primitive::Image(i) => Some(i),
            _ => None,
        })
        .collect()
}

fn shape_prims(prims: &[Primitive]) -> Vec<&docx_layout::display_list::ShapePrimitive> {
    prims
        .iter()
        .filter_map(|p| match p {
            Primitive::Shape(s) => Some(s),
            _ => None,
        })
        .collect()
}

fn border_line_count(prims: &[Primitive]) -> usize {
    prims
        .iter()
        .filter(|p| {
            matches!(p, Primitive::Line(l) if l.role == Some(docx_layout::display_list::LineRole::Border))
        })
        .count()
}

/// A DrawingML `shape` block emits one page-placed Shape primitive carrying the
/// scaled path, fill, stroke, transform, and doc range. This is the final
/// display-list arm for basic autoshapes, so the canvas path no longer
/// deserializes them as unsupported/no-op content. Pins determinism too.
#[test]
fn shape_fragment_emits_scaled_path_fill_stroke_and_transform() {
    let input = serde_json::json!({
        "measured": [{
            "block": {
                "kind": "shape",
                "id": 8,
                "shapeType": "rect",
                "geometryPath": [
                    { "type": "move", "x": 0.0, "y": 0.0 },
                    { "type": "line", "x": 1.0, "y": 0.0 },
                    { "type": "line", "x": 1.0, "y": 1.0 },
                    { "type": "line", "x": 0.0, "y": 1.0 },
                    { "type": "close" }
                ],
                "fill": { "type": "solid", "color": "#00B0F0" },
                "stroke": { "color": "#FF0000", "width": 2.0, "dash": "dash" },
                "transform": { "rotation": 15.0, "flipH": true },
                "width": 200.0,
                "height": 80.0,
                "docStart": 10,
                "docEnd": 11
            },
            "measure": { "kind": "shape", "width": 200.0, "height": 80.0 }
        }],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
            { "kind": "shape", "blockId": 8, "x": 100.0, "y": 50.0, "width": 200.0,
              "height": 80.0, "docStart": 10, "docEnd": 11 }
        ] }] }
    })
    .to_string();

    let a = build_display_list_json(&input).expect("builds");
    let b = build_display_list_json(&input).expect("builds");
    assert_eq!(a, b, "shape build is not deterministic");

    let dl: DisplayList = serde_json::from_str(&a).unwrap();
    let shapes = shape_prims(&dl.pages[0].primitives);
    assert_eq!(shapes.len(), 1);
    let shape = shapes[0];
    assert_eq!(shape.x.as_f64(), Some(100.0));
    assert_eq!(shape.y.as_f64(), Some(50.0));
    assert_eq!(shape.w.as_f64(), Some(200.0));
    assert_eq!(shape.h.as_f64(), Some(80.0));
    assert_eq!(shape.fill.as_deref(), Some("#00B0F0"));
    assert_eq!(shape.attrs.doc_start, Some(10));
    assert_eq!(shape.attrs.doc_end, Some(11));
    assert!(shape.decorative);

    let stroke = shape.stroke.as_ref().expect("stroke");
    assert_eq!(stroke.color, "#FF0000");
    assert_eq!(stroke.width.as_f64(), Some(2.0));
    assert_eq!(stroke.dash.as_deref(), Some("dash"));
    let transform = shape.transform.as_ref().expect("transform");
    assert_eq!(
        transform.rotation.as_ref().and_then(|n| n.as_f64()),
        Some(15.0)
    );
    assert!(transform.flip_h);
    assert!(!transform.flip_v);

    assert_eq!(shape.geometry_path.len(), 5);
    let point = |cmd: &ShapePathCommand| match cmd {
        ShapePathCommand::Move { x, y } | ShapePathCommand::Line { x, y } => {
            (x.as_f64().unwrap(), y.as_f64().unwrap())
        }
        other => panic!("expected point command, got {other:?}"),
    };
    assert_eq!(point(&shape.geometry_path[0]), (100.0, 50.0));
    assert_eq!(point(&shape.geometry_path[1]), (300.0, 50.0));
    assert_eq!(point(&shape.geometry_path[2]), (300.0, 130.0));
    assert_eq!(point(&shape.geometry_path[3]), (100.0, 130.0));
    assert!(matches!(shape.geometry_path[4], ShapePathCommand::Close));
}

/// A `textBox` block emits its container chrome (fill rect + four border edges)
/// and its inner paragraphs as real text at the content origin — the box's
/// content no longer silently vanishes on the canvas path. The content origin is
/// `frag.(x|y) + outlineWidth + margin` (box-sizing: border-box), and inner
/// paragraphs stack by their measured totalHeight. Also pins determinism.
#[test]
fn text_box_fragment_emits_container_and_inner_text() {
    let input = serde_json::json!({
        "measured": [{
            "block": {
                "kind": "textBox", "id": 5,
                "fillColor": "#eeeeff", "outlineWidth": 2.0, "outlineColor": "#334455",
                "outlineStyle": "solid",
                "margins": { "top": 4.0, "bottom": 4.0, "left": 7.0, "right": 7.0 },
                "content": [
                    { "kind": "paragraph", "id": 6, "pmStart": 2, "pmEnd": 8,
                      "runs": [{ "kind": "text", "text": "Alpha", "pmStart": 2, "pmEnd": 7 }] },
                    { "kind": "paragraph", "id": 7, "pmStart": 9, "pmEnd": 15,
                      "runs": [{ "kind": "text", "text": "Bravo", "pmStart": 9, "pmEnd": 14 }] }
                ]
            },
            "measure": {
                "kind": "textBox", "width": 200.0, "height": 60.0,
                "innerMeasures": [
                    { "kind": "paragraph", "totalHeight": 24.0, "lines": [
                        { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 5,
                          "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] },
                    { "kind": "paragraph", "totalHeight": 24.0, "lines": [
                        { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 5,
                          "width": 42.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }
                ]
            }
        }],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
            { "kind": "textBox", "blockId": 5, "x": 100.0, "y": 50.0, "width": 200.0,
              "height": 60.0, "pmStart": 1, "pmEnd": 16 }
        ] }] }
    })
    .to_string();

    // determinism
    let a = build_display_list_json(&input).expect("builds");
    let b = build_display_list_json(&input).expect("builds");
    assert_eq!(a, b, "text-box build is not deterministic");

    let dl: DisplayList = serde_json::from_str(&a).unwrap();
    let prims = &dl.pages[0].primitives;

    // container fill rect at the box's page rect, carrying the box's doc range
    let fills = rect_prims(prims);
    let fill = fills
        .iter()
        .find(|r| r.fill == "#eeeeff")
        .expect("container fill rect missing");
    assert_eq!(fill.x.as_f64(), Some(100.0));
    assert_eq!(fill.y.as_f64(), Some(50.0));
    assert_eq!(fill.w.as_f64(), Some(200.0));
    assert_eq!(fill.h.as_f64(), Some(60.0));
    assert_eq!(fill.attrs.doc_start, Some(1));
    assert_eq!(fill.attrs.doc_end, Some(16));

    // four border edges
    assert_eq!(
        border_line_count(prims),
        4,
        "text box draws four border edges"
    );

    // inner paragraphs at content origin: x = 100 + border 2 + padLeft 7 = 109;
    // first para top = 50 + 2 + 4 = 56 → baseline 56 + half-leading 4 + ascent 12
    // = 72; second para stacks by totalHeight 24 → baseline 96.
    let t = text_prims(prims);
    let alpha = t.iter().find(|x| x.0 == "Alpha").expect("Alpha inner text");
    let bravo = t.iter().find(|x| x.0 == "Bravo").expect("Bravo inner text");
    assert!(
        (alpha.1 - 109.0).abs() < 0.01,
        "Alpha x {} (want 109)",
        alpha.1
    );
    assert!(
        (alpha.3 - 72.0).abs() < 0.01,
        "Alpha baseline {} (want 72)",
        alpha.3
    );
    assert!(
        (bravo.1 - 109.0).abs() < 0.01,
        "Bravo x {} (want 109)",
        bravo.1
    );
    assert!(
        (bravo.3 - 96.0).abs() < 0.01,
        "Bravo baseline {} (want 96)",
        bravo.3
    );

    // inner text carries selectable doc positions (a11y + hit-test contract)
    let alpha_attr = prims
        .iter()
        .find_map(|p| match p {
            Primitive::Text(t) if t.text == "Alpha" => Some(&t.attrs),
            _ => None,
        })
        .unwrap();
    assert_eq!(alpha_attr.doc_start, Some(2));
}

/// A floating (square-wrap) image RUN inside a paragraph is emitted as an Image
/// primitive at its resolved page rect — it no longer leaves a hole where the
/// wrap zone was reserved. The paragraph's inline text still paints. A `behind`
/// float paints before the body text; a front float after it. Pins determinism.
#[test]
fn floating_image_run_emits_at_resolved_float_position() {
    let make = |wrap: &str| {
        serde_json::json!({
            "measured": [{
                "block": {
                    "kind": "paragraph", "id": 3, "pmStart": 1, "pmEnd": 40,
                    "runs": [
                        { "kind": "image", "src": "rId5", "width": 120.0, "height": 80.0,
                          "wrapType": wrap, "cssFloat": "left", "alt": "pic",
                          "pmStart": 1, "pmEnd": 2,
                          "position": {
                            "horizontal": { "relativeTo": "column", "posOffset": 0 },
                            "vertical": { "relativeTo": "paragraph", "posOffset": 0 } } },
                        { "kind": "text", "text": "wrapped text", "pmStart": 2, "pmEnd": 14 }
                    ]
                },
                "measure": { "kind": "paragraph", "totalHeight": 72.0, "lines": [
                    { "headRun": 1, "headChar": 0, "tailRun": 1, "tailChar": 12,
                      "width": 76.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0, "leftOffset": 132.0 },
                    { "headRun": 1, "headChar": 0, "tailRun": 1, "tailChar": 12,
                      "width": 76.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0, "leftOffset": 132.0 },
                    { "headRun": 1, "headChar": 0, "tailRun": 1, "tailChar": 12,
                      "width": 76.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }
                ] }
            }],
            "options": {},
            "layout": { "pages": [{ "size": { "w": 400.0, "h": 300.0 },
                "margins": { "top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0 },
                "fragments": [
                { "kind": "paragraph", "blockId": 3, "x": 96.0, "y": 96.0, "width": 208.0,
                  "height": 72.0, "fromLine": 0, "toLine": 3, "pmStart": 1, "pmEnd": 40 }
            ] }] }
        })
        .to_string()
    };

    // front float (square wrap)
    let input = make("square");
    let a = build_display_list_json(&input).expect("builds");
    let b = build_display_list_json(&input).expect("builds");
    assert_eq!(a, b, "floating-image build is not deterministic");
    let dl: DisplayList = serde_json::from_str(&a).unwrap();
    let prims = &dl.pages[0].primitives;

    // exactly one image, at the resolved page rect: content-relative (0,0) +
    // page margins (96,96) → (96,96); size = the run's extent 120×80.
    let imgs = image_prims(prims);
    assert_eq!(imgs.len(), 1, "the floating image run must paint");
    let img = imgs[0];
    assert_eq!(img.rel_id, "rId5");
    assert_eq!(img.x.as_f64(), Some(96.0));
    assert_eq!(img.y.as_f64(), Some(96.0));
    assert_eq!(img.w.as_f64(), Some(120.0));
    assert_eq!(img.h.as_f64(), Some(80.0));
    assert_eq!(img.alt_text.as_deref(), Some("pic"));
    assert_eq!(img.attrs.doc_start, Some(1));

    // the wrapped inline text still paints (three lines)
    let text_lines = text_prims(prims).len();
    assert_eq!(text_lines, 3, "wrapped inline text must still paint");

    // a square-wrap float is a FRONT float → it paints after the body text
    let img_idx = prims
        .iter()
        .position(|p| matches!(p, Primitive::Image(_)))
        .unwrap();
    let first_text_idx = prims
        .iter()
        .position(|p| matches!(p, Primitive::Text(t) if !t.text.is_empty()))
        .unwrap();
    assert!(
        img_idx > first_text_idx,
        "front float paints after body text"
    );

    // a behind-doc float paints BEFORE the body text (renderPage phase 3)
    let dl_behind: DisplayList =
        serde_json::from_str(&build_display_list_json(&make("behind")).unwrap()).unwrap();
    let bp = &dl_behind.pages[0].primitives;
    assert_eq!(image_prims(bp).len(), 1, "behind float still paints");
    let b_img_idx = bp
        .iter()
        .position(|p| matches!(p, Primitive::Image(_)))
        .unwrap();
    let b_text_idx = bp
        .iter()
        .position(|p| matches!(p, Primitive::Text(t) if !t.text.is_empty()))
        .unwrap();
    assert!(
        b_img_idx < b_text_idx,
        "behind float paints before body text"
    );
}

/// Table/cell border lines carry explicit ownership metadata: the owning grid
/// cell (with its border-ownership flags) and the enclosing table-fragment
/// identity ride on every table-border line, replacing the consumer-side
/// geometric fallback association.
#[test]
fn table_border_lines_carry_cell_and_table_ownership() {
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "table", "id": 7, "rows": [
                { "cells": [
                    { "borders": {
                        "top": { "color": "#111111", "width": 1.0, "style": "solid" },
                        "right": { "color": "#222222", "width": 1.0, "style": "solid" },
                        "bottom": { "color": "#333333", "width": 1.0, "style": "solid" },
                        "left": { "color": "#444444", "width": 1.0, "style": "solid" }
                      },
                      "blocks": [{ "kind": "paragraph", "id": 700, "pmStart": 2, "pmEnd": 6,
                                   "runs": [{ "kind": "text", "text": "cell", "pmStart": 2 }] }] }
                ] }
            ] },
            "measure": { "kind": "table", "columnWidths": [100.0], "totalHeight": 24.0,
                "rows": [
                    { "height": 24.0, "cells": [
                        { "blocks": [{ "kind": "paragraph", "totalHeight": 24.0,
                            "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                                        "width": 40.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }] }
                    ] }
                ] }
        }],
        "options": {},
        "layout": { "pages": [
            { "size": { "w": 400.0, "h": 200.0 }, "margins": {}, "fragments": [
                { "kind": "table", "blockId": 7, "x": 50.0, "y": 50.0, "width": 100.0,
                  "height": 24.0, "rowStart": 0, "rowEnd": 1 } ] }
        ] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();

    let borders: Vec<&docx_layout::display_list::LinePrimitive> = dl.pages[0]
        .primitives
        .iter()
        .filter_map(|p| match p {
            Primitive::Line(line)
                if line.role == Some(docx_layout::display_list::LineRole::TableBorder) =>
            {
                Some(line)
            }
            _ => None,
        })
        .collect();
    assert_eq!(borders.len(), 4, "four bordered edges expected: {json}");
    for line in &borders {
        let cell = line.attrs.cell.as_ref().expect("border line carries cell");
        assert_eq!(
            (cell.row, cell.col, cell.row_span, cell.col_span),
            (0, 0, 1, 1)
        );
        assert_eq!(cell.cell_id.as_deref(), Some("7-r0-c0"));
        // every authored edge is owned by this single cell
        assert_eq!(cell.owns_top_border, Some(true));
        assert_eq!(cell.owns_left_border, Some(true));
        let table = line
            .attrs
            .table
            .as_ref()
            .expect("border line carries table");
        assert_eq!(table.table_id, "7");
    }

    // ownership serializes to the wire (camelCase, additive)
    assert!(json.contains(r#""borderOwner":"cell""#), "{json}");
    assert!(json.contains(r#""cellId":"7-r0-c0""#), "{json}");
}

/// Shape picture fills thread the full resolved payload — safe embedded
/// source, source crop, fill mode, tile parameters, and alpha — into the
/// primitive's `fillPaint`, and refuse to pass through a non-embedded source.
#[test]
fn shape_picture_fill_payload_threads_to_fill_paint() {
    let build = |src: &str| -> serde_json::Value {
        let input = serde_json::json!({
            "measured": [{
                "block": {
                    "kind": "shape",
                    "id": 8,
                    "shapeType": "rect",
                    "geometryPath": [
                        { "type": "move", "x": 0.0, "y": 0.0 },
                        { "type": "line", "x": 1.0, "y": 0.0 },
                        { "type": "close" }
                    ],
                    "fill": {
                        "type": "picture",
                        "pictureRelId": "rId5",
                        "pictureSrc": src,
                        "pictureSrcRect": { "left": 0.1, "top": 0.2 },
                        "pictureFillMode": "tile",
                        "pictureTile": { "scaleX": 0.5, "scaleY": 0.5, "alignment": "ctr", "flip": "xy" },
                        "pictureStretchRect": { "left": 0.05 },
                        "pictureOpacity": 0.75
                    },
                    "width": 100.0,
                    "height": 50.0
                },
                "measure": { "kind": "shape", "width": 100.0, "height": 50.0 }
            }],
            "options": {},
            "layout": { "pages": [{ "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
                { "kind": "shape", "blockId": 8, "x": 10.0, "y": 10.0, "width": 100.0, "height": 50.0 }
            ] }] }
        });
        let json = build_display_list_json(&input.to_string()).expect("builds");
        let dl: serde_json::Value = serde_json::from_str(&json).unwrap();
        dl["pages"][0]["primitives"][0]["fillPaint"].clone()
    };

    let paint = build("data:image/png;base64,AAA");
    assert_eq!(paint["kind"], "picture");
    assert_eq!(paint["pictureRelId"], "rId5");
    assert_eq!(paint["pictureSrc"], "data:image/png;base64,AAA");
    assert_eq!(paint["pictureSrcRect"]["left"], 0.1);
    assert_eq!(paint["pictureFillMode"], "tile");
    assert_eq!(paint["pictureTile"]["alignment"], "ctr");
    assert_eq!(paint["pictureTile"]["flip"], "xy");
    assert_eq!(paint["pictureStretchRect"]["left"], 0.05);
    assert_eq!(paint["pictureOpacity"], 0.75);

    // a non-embedded scheme is dropped at the emission gate (defense in depth
    // behind the parser-side embedded-only resolution)
    let external = build("https://example.com/tracker.png");
    assert!(
        external.get("pictureSrc").is_none(),
        "external pictureSrc must not pass through: {external}"
    );
}

/// Modern w14 text effects ride the run formatting losslessly onto the
/// text primitive so the canvas backend can replay glow/textFill/outline.
#[test]
fn modern_text_effects_thread_to_text_primitives() {
    let effects = serde_json::json!({
        "glow": { "color": "#00ff00", "radius": 4 },
        "textFill": { "kind": "solid", "color": "#ff00aa" }
    });
    let input = serde_json::json!({
        "measured": [{
            "block": { "kind": "paragraph", "id": 1, "pmStart": 2, "pmEnd": 8,
                "runs": [{ "kind": "text", "text": "hello", "pmStart": 2,
                           "modernEffects": effects }] },
            "measure": { "kind": "paragraph", "totalHeight": 24.0,
                "lines": [{ "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 5,
                            "width": 50.0, "ascent": 12.0, "descent": 4.0, "lineHeight": 24.0 }] }
        }],
        "options": {},
        "layout": { "pages": [{ "size": { "w": 400.0, "h": 300.0 }, "margins": {}, "fragments": [
            { "kind": "paragraph", "blockId": 1, "x": 50.0, "y": 50.0, "width": 300.0,
              "height": 24.0, "fromLine": 0, "toLine": 1, "pmStart": 2, "pmEnd": 8 }
        ] }] }
    });
    let json = build_display_list_json(&input.to_string()).expect("builds");
    let dl: DisplayList = serde_json::from_str(&json).unwrap();
    let text = dl.pages[0]
        .primitives
        .iter()
        .find_map(|p| match p {
            Primitive::Text(t) if t.text == "hello" => Some(t),
            _ => None,
        })
        .expect("text primitive");
    assert_eq!(text.attrs.modern_effects.as_ref(), Some(&effects));
}
