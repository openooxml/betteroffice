//! Note-area fixture and interaction gates.
//!
//! `two-footnotes` stacks two footnotes in one area at y 860..960 of a single
//! page, each note's story starting at position 1 — the same numbers, two
//! unrelated documents. Everything here turns on the two never being confused
//! with each other or with the body. The second note runs to a bordered
//! paragraph, so a border line sits between two of its own primitives.
//!
//! `two-column-footnotes` puts one note in each of two note columns, which the
//! emitter starts at the same vertical position: telling those two apart takes
//! more than the pointer's `y`.

use docx_layout::display_list::{
    DisplayList, Primitive, TextRunPrimitive, build_display_list_json,
};
use docx_layout::hit::{
    HitRegion, HoverTarget, RegionScope, hit_test_regions, range_rects, range_rects_in_region,
};

const NOTE_AREA_TOP: f64 = 860.0;
const NOTE_AREA_BOTTOM: f64 = 960.0;

fn read(name: &str) -> String {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join(format!("tests/fixtures/notes/{name}.input.json"));
    std::fs::read_to_string(path).unwrap_or_else(|_| panic!("missing {name}.input.json"))
}

fn input() -> String {
    read("two-footnotes")
}

fn build_from(input: &str) -> DisplayList {
    serde_json::from_str(&build_display_list_json(input).expect("builds")).expect("parses")
}

fn build() -> DisplayList {
    build_from(&input())
}

/// The run painting `text` inside the page's only note area.
fn note_run<'a>(dl: &'a DisplayList, text: &str) -> &'a TextRunPrimitive {
    dl.pages[0].note_areas[0]
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Text(run) if run.text == text => Some(run),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no note run painting {text:?}"))
}

/// Horizontal centre of a run, where a pointer probe lands on its glyphs.
fn run_center_x(run: &TextRunPrimitive) -> f64 {
    run.x.as_f64().unwrap() + run.width.as_f64().unwrap() / 2.0
}

#[test]
fn note_primitives_carry_their_own_story_range() {
    let dl = build();

    for (text, group, doc_start, doc_end) in [
        ("First note", "footnote-1", 1, 11),
        ("Second note", "footnote-2", 1, 12),
        ("and more", "footnote-2", 14, 22),
    ] {
        let attrs = &note_run(&dl, text).attrs;
        assert_eq!(attrs.group_id.as_deref(), Some(group));
        // the note story's own range, not the body anchor the mark sits at
        assert_eq!(
            (attrs.doc_start, attrs.doc_end),
            (Some(doc_start), Some(doc_end))
        );
    }

    // the body anchor stays reachable as backlink metadata
    let anchors: Vec<_> = dl.pages[0].note_areas[0]
        .notes
        .iter()
        .map(|note| (note.id, note.anchor_doc_start, note.anchor_doc_end))
        .collect();
    assert_eq!(
        anchors,
        vec![(Some(1), Some(9), Some(10)), (Some(2), Some(21), Some(22))]
    );
}

#[test]
fn a_point_in_a_note_resolves_against_that_note_story() {
    let dl = build();

    for (text, note_id, doc_end) in [
        ("First note", 1, 11),
        ("Second note", 2, 12),
        ("and more", 2, 22),
    ] {
        let run = note_run(&dl, text);
        let baseline = run.baseline_y.as_f64().unwrap();
        let hit = hit_test_regions(&dl, 0, run_center_x(run), baseline - 4.0).expect("page 0");
        assert_eq!(hit.region, HitRegion::Footnote, "{text}");
        assert_eq!(hit.note_id, Some(note_id), "{text}");
        assert_eq!(hit.target, HoverTarget::Text, "{text}");
        let pos = hit
            .pos
            .unwrap_or_else(|| panic!("{text} resolved no position"));
        assert!(
            (1..=doc_end).contains(&pos),
            "{text} resolved {pos} outside its story"
        );
    }

    // the area owns every point in its band, and the body owns the line above
    let below = hit_test_regions(&dl, 0, 400.0, NOTE_AREA_BOTTOM - 1.0).expect("page 0");
    assert_eq!(below.region, HitRegion::Footnote);
    let body = hit_test_regions(&dl, 0, 120.0, NOTE_AREA_TOP - 40.0).expect("page 0");
    assert_eq!(body.region, HitRegion::Body);
    assert_eq!(body.note_id, None);
}

#[test]
fn range_rects_scoped_to_a_note_cover_that_note_only() {
    let dl = build();

    // 1..11 covers the first line of either story and stops short of the
    // second note's own second paragraph, so each yields exactly one rect
    let first = range_rects_in_region(&dl, RegionScope::Footnote(1), 1, 11);
    let second = range_rects_in_region(&dl, RegionScope::Footnote(2), 1, 11);
    assert_eq!(first.len(), 1);
    assert_eq!(second.len(), 1);
    for rect in first.iter().chain(&second) {
        assert_eq!(rect.page_index, 0);
        assert!(rect.width > 0.0 && rect.height > 0.0);
        assert!(
            rect.y >= NOTE_AREA_TOP - 1.0 && rect.y <= NOTE_AREA_BOTTOM,
            "note rect y {} outside the area band",
            rect.y
        );
    }
    assert_ne!(first[0].y, second[0].y, "both notes highlighted one line");

    // the same positions in the body are a different document entirely
    for rect in range_rects(&dl, 1, 11) {
        assert!(
            rect.y < NOTE_AREA_TOP,
            "a body rect landed in the note area"
        );
    }
    // and a note the area does not carry highlights nothing
    assert!(range_rects_in_region(&dl, RegionScope::Footnote(9), 1, 11).is_empty());
    assert!(range_rects_in_region(&dl, RegionScope::Endnote(1), 1, 11).is_empty());
}

/// A paragraph or table border paints a line that carries no attrs, so it
/// names no paint group. It must not end the note's span of primitives:
/// everything after it belongs to the same story, and dropping it would leave
/// a selection over the whole note highlighting only its first lines.
#[test]
fn a_border_line_inside_a_note_does_not_split_its_story() {
    let dl = build();
    let primitives = &dl.pages[0].note_areas[0].primitives;
    let line = primitives
        .iter()
        .position(|primitive| matches!(primitive, Primitive::Line(_)))
        .expect("the fixture note paints a border line");
    assert!(
        primitives[line + 1..]
            .iter()
            .any(|primitive| matches!(primitive, Primitive::Text(_))),
        "the border line no longer sits before any of the note's text"
    );

    let rects = range_rects_in_region(&dl, RegionScope::Footnote(2), 1, 22);
    assert_eq!(
        rects.len(),
        2,
        "the note lines painted after its border line lost their rects"
    );
    assert!(rects[1].y > rects[0].y);
}

/// A note area lays its notes out in columns the emitter starts at the same
/// vertical position, so two stories can share a band and only the pointer's
/// `x` tells them apart. Resolving on `y` alone hands a point in one column the
/// other column's note and a position from its story.
#[test]
fn notes_in_parallel_columns_resolve_by_column() {
    let dl = build_from(&read("two-column-footnotes"));

    let left = note_run(&dl, "Left column");
    let right = note_run(&dl, "Right column");
    assert!(
        right.x.as_f64().unwrap() > run_center_x(left),
        "the fixture did not lay its notes out side by side"
    );
    assert_eq!(
        left.baseline_y, right.baseline_y,
        "the columns no longer share a vertical band, which is what this pins"
    );

    for (run, note_id) in [(left, 1), (right, 2)] {
        let baseline = run.baseline_y.as_f64().unwrap();
        let hit = hit_test_regions(&dl, 0, run_center_x(run), baseline - 4.0).expect("page 0");
        assert_eq!(hit.note_id, Some(note_id), "over {:?}", run.text);
        assert_eq!(hit.target, HoverTarget::Text, "over {:?}", run.text);
    }
}

#[test]
fn note_range_rects_match_the_json_and_handle_paths() {
    use docx_layout::hit::range_rects_region_json;
    use docx_layout::session::{
        close_display_list, open_display_list, range_rects_region_by_handle,
    };

    let dl_json = build_display_list_json(&input()).expect("builds");
    let native = range_rects_in_region(&build(), RegionScope::Footnote(2), 1, 12);
    assert!(!native.is_empty());

    let via_json = range_rects_region_json(&dl_json, "footnote", "2", 1, 12).expect("json ok");
    assert_eq!(via_json, serde_json::to_string(&native).unwrap());

    let handle = open_display_list(&dl_json).expect("opens");
    let via_handle =
        range_rects_region_by_handle(handle, "footnote", "2", 1, 12).expect("handle ok");
    close_display_list(handle);
    assert_eq!(via_handle, via_json);

    // a note region cannot be addressed without naming the note
    assert!(range_rects_region_json(&dl_json, "footnote", "", 1, 12).is_err());
}
