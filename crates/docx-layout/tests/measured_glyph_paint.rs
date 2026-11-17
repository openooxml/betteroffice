//! Joining and Indic lines paint the glyphs measurement shaped them with.

use docx_layout::display_list::{
    DisplayList, GlyphRunPrimitive, Primitive, build_display_list_json_with_fonts,
};
use docx_layout::hit::{caret_rect, range_rects};
use ooxml_text::{FontId, FontStore, ShapeDirection, ShapeFeature, shape_with_direction};
use serde_json::{Value, json};

const LATIN: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NASKH: &[u8] = include_bytes!("../../../packages/fonts/assets/NotoNaskhArabic-Regular.ttf");
const DEVANAGARI: &[u8] =
    include_bytes!("../../../packages/fonts-intl/assets/NotoSansDevanagari-Regular.ttf.zst");
const SYRIAC: &[u8] =
    include_bytes!("../../../packages/fonts-intl/assets/NotoSansSyriac-Regular.ttf.zst");
const KERN_OFF: [ShapeFeature; 1] = [ShapeFeature {
    tag: *b"kern",
    value: 0,
}];
const SIZE_PX: f32 = 16.0;
const FRAG_X: f64 = 96.0;
const FRAG_WIDTH: f64 = 624.0;
const ARABIC: &str = "مرحبا بالعالم";

struct Fixture {
    store: FontStore,
    latin: FontId,
    face: FontId,
}

fn fixture(face: &[u8]) -> Fixture {
    let mut store = FontStore::new();
    let latin = store.register(LATIN.to_vec()).unwrap();
    let face = store.register(face.to_vec()).unwrap();
    Fixture { store, latin, face }
}

fn compressed(bytes: &[u8]) -> Vec<u8> {
    betteroffice_fonts::decode(bytes, None, 64 << 20).unwrap()
}

fn text_run(text: &str, pm_start: i64) -> Value {
    json!({
        "kind": "text", "text": text, "fontFamily": "Face", "fontSize": 12.0,
        "pmStart": pm_start,
        "pmEnd": pm_start + text.encode_utf16().count() as i64
    })
}

struct Built {
    measure: Value,
    input: Value,
}

fn build(fixture: &Fixture, runs: Value, attrs: Value, width: f64) -> Built {
    let chains = json!({ "face|0|0": [fixture.face.to_u32(), fixture.latin.to_u32()] });
    let block = json!({ "kind": "paragraph", "id": 0, "runs": runs, "attrs": attrs, "pmStart": 1, "pmEnd": 200 });
    let mut measure_block = block.clone();
    for run in measure_block["runs"].as_array_mut().unwrap() {
        run["wordFeatures"] = json!(ooxml_text::ligature_features(
            run["ligatures"].as_str(),
            run["contextualAlternates"].as_bool(),
        ));
    }
    let measure: Value = serde_json::from_str(
        &ooxml_text::measure_paragraph_json(
            &fixture.store,
            &json!({
                "block": measure_block,
                "maxWidth": width,
                "fontChains": chains,
                "authoritativeShaping": true,
                "defaults": { "fontSize": 12.0, "fontFamily": "Face" }
            })
            .to_string(),
        )
        .unwrap(),
    )
    .unwrap();
    let line_count = measure["lines"].as_array().unwrap().len();
    let input = json!({
        "measured": [{ "block": block, "measure": measure }],
        "fontChains": chains,
        "options": {},
        "layout": { "pages": [{
            "size": { "w": 816.0, "h": 1056.0 },
            "margins": { "top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0 },
            "number": 1,
            "fragments": [{
                "kind": "paragraph", "blockId": 0,
                "x": FRAG_X, "y": 96.0, "width": FRAG_WIDTH, "height": 20.0 * line_count as f64,
                "fromLine": 0, "toLine": line_count, "pmStart": 1, "pmEnd": 200
            }]
        }] }
    });
    Built { measure, input }
}

fn list_of(fixture: &Fixture, input: &Value) -> DisplayList {
    serde_json::from_str(
        &build_display_list_json_with_fonts(&input.to_string(), &fixture.store).unwrap(),
    )
    .unwrap()
}

fn painted(fixture: &Fixture, built: &Built) -> DisplayList {
    list_of(fixture, &built.input)
}

#[test]
fn carlito_paints_the_measured_optional_ligatures() {
    let fixture = fixture(include_bytes!(
        "../../../packages/fonts/assets/Carlito-Regular.ttf"
    ));
    let mut counts = Vec::new();
    for setting in [None, Some("none"), Some("standardContextual")] {
        let mut run = text_run("fi ti", 1);
        if let Some(setting) = setting {
            run["ligatures"] = json!(setting);
        }
        let built = build(&fixture, json!([run]), json!({}), FRAG_WIDTH);
        let measured =
            ooxml_text::unpack_glyph_runs(built.measure["lines"][0]["glyphRuns"].as_str().unwrap())
                .unwrap();
        let expected: Vec<_> = measured
            .iter()
            .flat_map(|run| run.glyphs.iter().map(|glyph| glyph.id))
            .collect();
        for list in [painted(&fixture, &built), reshaped(&fixture, &built)] {
            let actual: Vec<_> = glyph_runs(&list).iter().flat_map(|run| ids(run)).collect();
            assert_eq!(actual, expected);
            for run in glyph_runs(&list) {
                let encoded = serde_json::to_value(run).unwrap();
                if setting == Some("standardContextual") {
                    assert_eq!(encoded["fontFeatures"], json!(["liga", "clig"]));
                } else {
                    assert!(encoded.get("fontFeatures").is_none());
                }
            }
        }
        counts.push(expected.len());
    }
    assert_eq!(counts[0], 5);
    assert_eq!(counts[1], counts[0]);
    assert!(counts[2] < counts[0]);
}

fn reshaped(fixture: &Fixture, built: &Built) -> DisplayList {
    let mut input = built.input.clone();
    for line in input["measured"][0]["measure"]["lines"]
        .as_array_mut()
        .unwrap()
    {
        line.as_object_mut().unwrap().remove("glyphRuns");
    }
    list_of(fixture, &input)
}

fn glyph_runs(list: &DisplayList) -> Vec<&GlyphRunPrimitive> {
    list.pages[0]
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::GlyphRun(run) => Some(run),
            _ => None,
        })
        .collect()
}

fn ids(run: &GlyphRunPrimitive) -> Vec<u32> {
    run.glyphs.iter().map(|glyph| glyph.id).collect()
}

fn whole_ids(fixture: &Fixture, text: &str, direction: ShapeDirection) -> Vec<u32> {
    shape_with_direction(
        &fixture.store,
        fixture.face,
        text,
        SIZE_PX,
        &KERN_OFF,
        direction,
    )
    .unwrap()
    .iter()
    .map(|glyph| glyph.glyph_id)
    .collect()
}

fn right_edge(runs: &[&GlyphRunPrimitive]) -> f64 {
    runs.iter()
        .flat_map(|run| run.glyphs.iter())
        .map(|glyph| glyph.x + glyph.advance)
        .fold(f64::MIN, f64::max)
}

/// Total paint width per UTF-16 offset in the document, whichever primitive holds it.
fn cluster_widths(runs: &[&GlyphRunPrimitive]) -> Vec<(i64, f64)> {
    let mut widths = std::collections::BTreeMap::new();
    for run in runs {
        for glyph in &run.glyphs {
            let offset = run.text[..glyph.cluster as usize].encode_utf16().count() as i64;
            *widths
                .entry(run.attrs.doc_start.unwrap() + offset)
                .or_insert(0.0) += glyph.advance;
        }
    }
    widths.into_iter().collect()
}

fn assert_same_widths(left: &[(i64, f64)], right: &[(i64, f64)]) {
    assert_eq!(left.len(), right.len(), "{left:?} {right:?}");
    for (left, right) in left.iter().zip(right) {
        assert_eq!(left.0, right.0);
        assert!((left.1 - right.1).abs() < 0.01, "{left:?} {right:?}");
    }
}

#[test]
fn an_arabic_line_paints_one_run_of_contextual_glyphs() {
    let fixture = fixture(NASKH);
    let built = build(
        &fixture,
        json!([text_run(ARABIC, 1)]),
        json!({ "bidi": true }),
        FRAG_WIDTH,
    );
    let list = painted(&fixture, &built);
    let runs = glyph_runs(&list);
    assert_eq!(runs.len(), 1, "{runs:?}");
    let run = runs[0];
    assert_eq!(run.text, ARABIC);
    assert_eq!(run.font_id, fixture.face.to_u32());
    assert_eq!(run.rtl, Some(true));
    assert_eq!(
        (run.attrs.doc_start, run.attrs.doc_end),
        (Some(1), Some(1 + ARABIC.encode_utf16().count() as i64))
    );
    assert_eq!(ids(run), whole_ids(&fixture, ARABIC, ShapeDirection::Rtl));
    let before = reshaped(&fixture, &built);
    let isolated: Vec<u32> = glyph_runs(&before)
        .iter()
        .flat_map(|run| ids(run))
        .collect();
    assert_ne!(
        ids(run),
        isolated,
        "cluster-by-cluster shaping loses the joins"
    );
    let clusters: Vec<u32> = run.glyphs.iter().map(|glyph| glyph.cluster).collect();
    assert!(
        clusters.windows(2).all(|pair| pair[0] >= pair[1]),
        "{clusters:?}"
    );
    assert!(run.glyphs.iter().all(|glyph| glyph.bidi_level == Some(1)));
    assert_eq!(
        run.attrs.fallback_font,
        glyph_runs(&before)[0].attrs.fallback_font
    );
}

#[test]
fn devanagari_and_syriac_lines_paint_their_measured_glyphs() {
    for (font, text, direction, attrs) in [
        (DEVANAGARI, "क्षत्रिय नमस्ते", ShapeDirection::Ltr, json!({})),
        (
            SYRIAC,
            "ܫܠܡܐ ܥܠܡܐ",
            ShapeDirection::Rtl,
            json!({ "bidi": true }),
        ),
    ] {
        let fixture = fixture(&compressed(font));
        let built = build(&fixture, json!([text_run(text, 1)]), attrs, FRAG_WIDTH);
        let list = painted(&fixture, &built);
        let runs = glyph_runs(&list);
        assert_eq!(runs.len(), 1, "{text}: {runs:?}");
        assert_eq!(runs[0].text, text);
        assert_eq!(ids(runs[0]), whole_ids(&fixture, text, direction), "{text}");
        let before = reshaped(&fixture, &built);
        assert_same_widths(
            &cluster_widths(&runs),
            &cluster_widths(&glyph_runs(&before)),
        );
    }
}

#[test]
fn a_justified_arabic_line_keeps_its_widths_and_end() {
    let fixture = fixture(NASKH);
    let text = "مرحبا بالعالم الجميل";
    let built = build(
        &fixture,
        json!([text_run(text, 1), { "kind": "lineBreak", "pmStart": 21 }]),
        json!({ "bidi": true, "alignment": "justify" }),
        FRAG_WIDTH,
    );
    let list = painted(&fixture, &built);
    let before = reshaped(&fixture, &built);
    let (runs, old) = (glyph_runs(&list), glyph_runs(&before));
    assert_eq!(runs.len(), 1);
    assert!(runs[0].word_spacing.is_some(), "the line stretches");
    assert!((right_edge(&runs) - right_edge(&old)).abs() < 0.01);
    assert_same_widths(&cluster_widths(&runs), &cluster_widths(&old));
}

#[test]
fn letter_spacing_lands_on_each_cluster() {
    let fixture = fixture(NASKH);
    let mut run = text_run(ARABIC, 1);
    run["letterSpacing"] = json!(2.0);
    let built = build(&fixture, json!([run]), json!({ "bidi": true }), FRAG_WIDTH);
    let list = painted(&fixture, &built);
    let before = reshaped(&fixture, &built);
    let (runs, old) = (glyph_runs(&list), glyph_runs(&before));
    assert_eq!(runs.len(), 1);
    assert_eq!(
        ids(runs[0]),
        whole_ids(&fixture, ARABIC, ShapeDirection::Rtl)
    );
    assert_same_widths(&cluster_widths(&runs), &cluster_widths(&old));
    let measured: f64 = built.measure["lines"][0]["clusterAdvances"]
        .as_array()
        .unwrap()
        .iter()
        .map(|cluster| cluster["advance"].as_f64().unwrap())
        .sum();
    let painted: f64 = runs[0].glyphs.iter().map(|glyph| glyph.advance).sum();
    assert!((measured - painted).abs() < 0.02, "{measured} {painted}");
}

#[test]
fn a_mixed_bidi_line_merges_only_the_arabic() {
    let fixture = fixture(NASKH);
    let mut run = text_run("abc سلام def", 1);
    run["fontFamily"] = json!("Face");
    let chains = json!({ "face|0|0": [fixture.latin.to_u32(), fixture.face.to_u32()] });
    let mut built = build(&fixture, json!([run]), json!({}), FRAG_WIDTH);
    let block = built.input["measured"][0]["block"].clone();
    built.measure = serde_json::from_str(
        &ooxml_text::measure_paragraph_json(
            &fixture.store,
            &json!({
                "block": block, "maxWidth": FRAG_WIDTH, "fontChains": chains,
                "authoritativeShaping": true,
                "defaults": { "fontSize": 12.0, "fontFamily": "Face" }
            })
            .to_string(),
        )
        .unwrap(),
    )
    .unwrap();
    built.input["measured"][0]["measure"] = built.measure.clone();
    built.input["fontChains"] = chains;
    let list = painted(&fixture, &built);
    let before = reshaped(&fixture, &built);
    let runs = glyph_runs(&list);
    let arabic: Vec<_> = runs
        .iter()
        .filter(|run| run.text.chars().any(ooxml_text::is_joining_or_indic))
        .collect();
    assert_eq!(arabic.len(), 1);
    assert_eq!(arabic[0].text, "سلام");
    assert_eq!(arabic[0].rtl, Some(true));
    assert_eq!(
        ids(arabic[0]),
        whole_ids(&fixture, "سلام", ShapeDirection::Rtl)
    );
    let latin = |runs: &[&GlyphRunPrimitive]| -> Vec<GlyphRunPrimitive> {
        runs.iter()
            .filter(|run| !run.text.chars().any(ooxml_text::is_joining_or_indic))
            .map(|run| (*run).clone())
            .collect()
    };
    assert_eq!(latin(&runs), latin(&glyph_runs(&before)));
    let starts: Vec<f64> = runs.iter().map(|run| run.glyphs[0].x).collect();
    assert!(
        starts.windows(2).all(|pair| pair[0] < pair[1]),
        "{starts:?}"
    );
}

#[test]
fn a_word_cut_across_lines_keeps_its_glyphs() {
    let fixture = fixture(NASKH);
    let text = "مستقبل".repeat(4);
    let built = build(
        &fixture,
        json!([text_run(&text, 1)]),
        json!({ "bidi": true }),
        60.0,
    );
    assert!(built.measure["lines"].as_array().unwrap().len() > 1);
    let list = painted(&fixture, &built);
    let runs = glyph_runs(&list);
    assert_eq!(runs.len(), built.measure["lines"].as_array().unwrap().len());
    let cut: Vec<u32> = runs.iter().rev().flat_map(|run| ids(run)).collect();
    assert_eq!(cut, whole_ids(&fixture, &text, ShapeDirection::Rtl));
}

#[test]
fn a_field_and_a_tab_beside_arabic_keep_their_own_paint() {
    let fixture = fixture(NASKH);
    let built = build(
        &fixture,
        json!([
            text_run("سلام", 1),
            { "kind": "tab", "fontFamily": "Face", "fontSize": 12.0, "pmStart": 5, "pmEnd": 6 },
            text_run("عالم", 6),
            { "kind": "field", "fieldType": "MERGEFIELD", "fallback": "مرحبا", "fontFamily": "Face", "fontSize": 12.0, "pmStart": 10, "pmEnd": 11 }
        ]),
        json!({ "bidi": true }),
        FRAG_WIDTH,
    );
    let list = painted(&fixture, &built);
    let before = reshaped(&fixture, &built);
    let runs = glyph_runs(&list);
    for word in ["سلام", "عالم"] {
        let run = runs.iter().find(|run| run.text == word).expect(word);
        assert_eq!(ids(run), whole_ids(&fixture, word, ShapeDirection::Rtl));
    }
    let field = |runs: &[&GlyphRunPrimitive]| -> Vec<GlyphRunPrimitive> {
        runs.iter()
            .filter(|run| run.attrs.field.is_some())
            .map(|run| (*run).clone())
            .collect()
    };
    assert!(!field(&runs).is_empty());
    assert_eq!(field(&runs), field(&glyph_runs(&before)));
    assert!((right_edge(&runs) - right_edge(&glyph_runs(&before))).abs() < 0.01);
}

#[test]
fn a_line_without_glyph_runs_reshapes_its_clusters() {
    let fixture = fixture(NASKH);
    let built = build(
        &fixture,
        json!([text_run(ARABIC, 1)]),
        json!({ "bidi": true }),
        FRAG_WIDTH,
    );
    let before = reshaped(&fixture, &built);
    let runs = glyph_runs(&before);
    assert_eq!(runs.len(), ARABIC.chars().count());
    for run in runs {
        assert_eq!(
            ids(run),
            whole_ids(&fixture, &run.text, ShapeDirection::Rtl)
        );
    }
}

#[test]
fn horizontal_scale_keeps_measured_glyphs_per_cluster() {
    let fixture = fixture(NASKH);
    let mut run = text_run(ARABIC, 1);
    run["horizontalScale"] = json!(80.0);
    let built = build(&fixture, json!([run]), json!({ "bidi": true }), FRAG_WIDTH);
    let list = painted(&fixture, &built);
    let before = reshaped(&fixture, &built);
    let (runs, old) = (glyph_runs(&list), glyph_runs(&before));
    assert_eq!(runs.len(), old.len());
    let painted_ids: Vec<u32> = runs.iter().flat_map(|run| ids(run)).collect();
    assert_eq!(
        painted_ids,
        whole_ids(&fixture, ARABIC, ShapeDirection::Rtl)
    );
    for (run, old) in runs.iter().zip(&old) {
        assert_eq!(run.horizontal_scale, old.horizontal_scale);
        assert_eq!(run.attrs.doc_start, old.attrs.doc_start);
        let last = run.glyphs.last().unwrap();
        let old_last = old.glyphs.last().unwrap();
        assert!((last.x + last.advance - old_last.x - old_last.advance).abs() < 0.01);
    }
}

#[test]
fn a_complex_script_size_paints_at_that_size() {
    let fixture = fixture(NASKH);
    let mut run = text_run("مرحبا", 1);
    run["fontSizeCs"] = json!(18.0);
    let built = build(&fixture, json!([run]), json!({ "bidi": true }), FRAG_WIDTH);
    let list = painted(&fixture, &built);
    assert_eq!(glyph_runs(&list)[0].size, 24.0);
}

#[test]
fn carets_and_ranges_resolve_inside_a_merged_run() {
    let fixture = fixture(NASKH);
    let built = build(
        &fixture,
        json!([text_run(ARABIC, 1)]),
        json!({ "bidi": true }),
        FRAG_WIDTH,
    );
    let list = painted(&fixture, &built);
    let before = reshaped(&fixture, &built);
    for position in 1..=ARABIC.chars().count() as i64 {
        let (now, then) = (
            caret_rect(&list, position).unwrap(),
            caret_rect(&before, position).unwrap(),
        );
        assert!(
            (now.x - then.x).abs() < 1.0,
            "{position}: {} {}",
            now.x,
            then.x
        );
    }
    let width =
        |list: &DisplayList| -> f64 { range_rects(list, 2, 5).iter().map(|rect| rect.width).sum() };
    assert!((width(&list) - width(&before)).abs() < 1.0);
}

#[test]
fn joining_and_indic_runs_never_echo() {
    for (font, text) in [(NASKH.to_vec(), "مرحبا"), (compressed(DEVANAGARI), "नमस्ते")]
    {
        let fixture = fixture(&font);
        let built = build(&fixture, json!([text_run(text, 1)]), json!({}), FRAG_WIDTH);
        for list in [painted(&fixture, &built), reshaped(&fixture, &built)] {
            let runs = glyph_runs(&list);
            assert!(!runs.is_empty());
            assert!(runs.iter().all(|run| run.echo.is_none()), "{text}");
        }
    }
}

#[test]
fn latin_lines_paint_as_before_and_still_echo() {
    let fixture = fixture(LATIN);
    let built = build(
        &fixture,
        json!([text_run("office AV Wave", 1)]),
        json!({}),
        FRAG_WIDTH,
    );
    let (list, before) = (painted(&fixture, &built), reshaped(&fixture, &built));
    assert_eq!(glyph_runs(&list), glyph_runs(&before));
    assert_eq!(glyph_runs(&list).len(), "office AV Wave".chars().count());
    assert!(glyph_runs(&list).iter().all(|run| run.echo.is_some()));
}
