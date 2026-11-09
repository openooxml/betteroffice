//! Under `authoritativeShaping` every line carries the glyphs measurement
//! shaped it with, so paint can draw them instead of reshaping clusters.

use ooxml_text::{
    FontId, FontStore, MeasuredGlyphRun, ShapeDirection, ShapeFeature, measure_paragraph_json,
    shape_with_direction, unpack_glyph_runs,
};
use serde_json::{Value, json};

const LATIN: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");
const NASKH: &[u8] = include_bytes!("../../../packages/fonts/assets/NotoNaskhArabic-Regular.ttf");
const SIZE_PX: f32 = 16.0;
const KERN_OFF: [ShapeFeature; 1] = [ShapeFeature {
    tag: *b"kern",
    value: 0,
}];

fn store() -> (FontStore, FontId, FontId) {
    let mut store = FontStore::new();
    let latin = store.register(LATIN.to_vec()).unwrap();
    let naskh = store.register(NASKH.to_vec()).unwrap();
    (store, latin, naskh)
}

fn measure(store: &FontStore, text: &str, chain: &[FontId], bidi: bool, width: f64) -> Value {
    measure_with(store, text, chain, bidi, width, true)
}

fn measure_with(
    store: &FontStore,
    text: &str,
    chain: &[FontId],
    bidi: bool,
    width: f64,
    authoritative: bool,
) -> Value {
    let chain: Vec<u32> = chain.iter().map(|font| font.to_u32()).collect();
    let input = json!({
        "block": {
            "kind": "paragraph",
            "runs": [{ "kind": "text", "text": text }],
            "attrs": { "bidi": bidi }
        },
        "maxWidth": width,
        "fontChains": { "liberation sans|0|0": chain },
        "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" },
        "authoritativeShaping": authoritative
    });
    let out = measure_paragraph_json(store, &input.to_string()).unwrap();
    serde_json::from_str(&out).unwrap()
}

fn shaped_ids(store: &FontStore, font: FontId, text: &str) -> Vec<u64> {
    shape_with_direction(store, font, text, SIZE_PX, &KERN_OFF, ShapeDirection::Rtl)
        .unwrap()
        .iter()
        .map(|glyph| u64::from(glyph.glyph_id))
        .collect()
}

fn glyph_runs(line: &Value) -> Vec<MeasuredGlyphRun> {
    let packed = line["glyphRuns"]
        .as_str()
        .unwrap_or_else(|| panic!("line carries glyphRuns: {line}"));
    unpack_glyph_runs(packed).unwrap()
}

fn glyph_field(run: &MeasuredGlyphRun, field: &str) -> Vec<f64> {
    run.glyphs
        .iter()
        .map(|glyph| match field {
            "cluster" => f64::from(glyph.cluster),
            "advance" => f64::from(glyph.advance),
            other => panic!("no glyph field {other}"),
        })
        .collect()
}

fn run_ids(run: &MeasuredGlyphRun) -> Vec<u64> {
    run.glyphs.iter().map(|glyph| u64::from(glyph.id)).collect()
}

fn char_range(run: &MeasuredGlyphRun) -> (u64, u64) {
    (u64::from(run.head.start_char), u64::from(run.head.end_char))
}

#[test]
fn an_arabic_line_carries_the_whole_run_shaping() {
    let (store, _, naskh) = store();
    let text = "مرحبا بالعالم";
    let measured = measure(&store, text, &[naskh], true, 1000.0);
    let lines = measured["lines"].as_array().unwrap();
    assert_eq!(lines.len(), 1);
    let runs = glyph_runs(&lines[0]);
    assert_eq!(runs.len(), 1, "{runs:?}");
    let run = &runs[0];
    assert_eq!(run.head.font_id, naskh.to_u32());
    assert_eq!(run.head.bidi_level, 1);
    assert_eq!(char_range(run), (0, 13));
    assert_eq!(run.head.x_offset, 0.0);

    let whole =
        shape_with_direction(&store, naskh, text, SIZE_PX, &KERN_OFF, ShapeDirection::Rtl).unwrap();
    assert_eq!(run_ids(run), shaped_ids(&store, naskh, text));
    let advances: Vec<f32> = glyph_field(run, "advance")
        .into_iter()
        .map(|advance| advance as f32)
        .collect();
    let shaped: Vec<f32> = whole.iter().map(|glyph| glyph.x_advance).collect();
    assert_eq!(advances, shaped);
    let isolated: Vec<u64> = text
        .chars()
        .rev()
        .flat_map(|ch| shaped_ids(&store, naskh, &ch.to_string()))
        .collect();
    assert_ne!(run_ids(run), isolated, "contextual forms are kept");
    let clusters = glyph_field(run, "cluster");
    assert!(
        clusters.windows(2).all(|pair| pair[0] >= pair[1]),
        "{clusters:?}"
    );
    let width = lines[0]["width"].as_f64().unwrap();
    assert!((f64::from(run.head.advance) - width).abs() < 1e-3);
}

#[test]
fn a_word_cut_across_lines_keeps_its_joined_glyphs() {
    let (store, _, naskh) = store();
    let text = "مستقبل".repeat(4);
    let units: Vec<u16> = text.encode_utf16().collect();
    let measured = measure(&store, &text, &[naskh], true, 60.0);
    let lines = measured["lines"].as_array().unwrap();
    assert!(lines.len() > 1, "the word is chopped: {measured}");
    let mut cut = Vec::new();
    let mut reshaped_differs = false;
    for line in lines.iter().rev() {
        let runs = glyph_runs(line);
        assert_eq!(runs.len(), 1);
        let (start, end) = char_range(&runs[0]);
        assert_eq!(start, line["headChar"].as_u64().unwrap());
        assert_eq!(end, line["tailChar"].as_u64().unwrap());
        for cluster in glyph_field(&runs[0], "cluster") {
            assert!((start as f64..end as f64).contains(&cluster));
        }
        let ids = run_ids(&runs[0]);
        let piece = String::from_utf16(&units[start as usize..end as usize]).unwrap();
        reshaped_differs |= ids != shaped_ids(&store, naskh, &piece);
        cut.extend(ids);
    }
    assert_eq!(cut, shaped_ids(&store, naskh, &text));
    assert!(reshaped_differs, "a cut line is not reshaped on its own");
}

#[test]
fn mixed_latin_and_arabic_lines_carry_one_run_per_face_and_direction() {
    let (store, latin, naskh) = store();
    let measured = measure(&store, "abc سلام def", &[latin, naskh], false, 1000.0);
    let line = &measured["lines"][0];
    let runs = glyph_runs(line);
    let summary: Vec<(u64, u64, (u64, u64))> = runs
        .iter()
        .map(|run| {
            (
                u64::from(run.head.font_id),
                u64::from(run.head.bidi_level),
                char_range(run),
            )
        })
        .collect();
    assert_eq!(run_ids(&runs[1]), shaped_ids(&store, naskh, "سلام"));
    let (latin, naskh) = (u64::from(latin.to_u32()), u64::from(naskh.to_u32()));
    assert_eq!(
        summary,
        vec![(latin, 0, (0, 4)), (naskh, 1, (4, 8)), (latin, 0, (8, 12))]
    );
    let mut x = 0.0;
    for run in &runs {
        assert!((f64::from(run.head.x_offset) - x).abs() < 1e-3, "{run:?}");
        x += f64::from(run.head.advance);
    }
    assert!((x - line["width"].as_f64().unwrap()).abs() < 1e-3);
    let arabic_x = line["clusterAdvances"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|cluster| cluster["bidiLevel"] == 1)
        .map(|cluster| cluster["xOffset"].as_f64().unwrap())
        .fold(f64::INFINITY, f64::min);
    assert!((f64::from(runs[1].head.x_offset) - arabic_x).abs() < 1e-3);
}

#[test]
fn lines_without_authoritative_shaping_carry_no_glyph_runs() {
    let (store, latin, naskh) = store();
    let measured = measure_with(&store, "abc سلام def", &[latin, naskh], false, 1000.0, false);
    assert!(measured["lines"][0].get("glyphRuns").is_none());
    assert!(measured["lines"][0].get("clusterAdvances").is_none());
}

#[test]
fn plain_text_lines_cost_little_on_the_wire() {
    let (store, latin, _) = store();
    let text = "plain words on a line of ordinary text";
    let measured = measure(&store, text, &[latin], false, 1000.0);
    let line = &measured["lines"][0];
    let runs = glyph_runs(line);
    assert_eq!(runs.len(), 1);
    let packed = line["glyphRuns"].as_str().unwrap();
    let limit = (4 + 19 + runs[0].glyphs.len() * 6).div_ceil(3) * 4;
    assert!(packed.len() <= limit, "{packed}");
}
