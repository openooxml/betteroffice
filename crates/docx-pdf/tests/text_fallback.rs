//! Fallback-text and rotated-glyph-run coverage for the PDF text painter.
//!
//! The crate keeps its font planner private, so this integration test compiles
//! the real `font`, `text` and `units` modules directly. That lets it exercise
//! the same code paths the exporter uses without re-exporting internals.

#![allow(dead_code)]

use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};

use docx_layout::display_list::{
    DisplayList, DisplayPage, DocAttrs, GlyphRunPrimitive, PlacedGlyph, Primitive, TextRunPrimitive,
};
use ooxml_text::FontStore;
use pdf_writer::Content;
use serde_json::Number;
use ttf_parser::Face;

#[path = "../src/font.rs"]
mod font;
#[path = "../src/text.rs"]
mod text;
#[path = "../src/units.rs"]
mod units;

/// Minimal stand-in for the exporter error type used by the included modules.
#[derive(Debug)]
pub enum PdfError {
    Unsupported(&'static str),
    InvalidNumber(&'static str),
    NoPages,
    Font(font::FontError),
    Image(String),
}

impl From<font::FontError> for PdfError {
    fn from(error: font::FontError) -> Self {
        Self::Font(error)
    }
}

impl std::fmt::Display for PdfError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{self:?}")
    }
}

impl std::error::Error for PdfError {}

fn font_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../packages/fonts/assets")
}

fn font_bytes(name: &str) -> Vec<u8> {
    std::fs::read(font_dir().join(name)).expect("font fixture")
}

/// A font under the repository fixtures that covers CJK, when one exists.
fn cjk_font() -> Option<Vec<u8>> {
    let roots = [
        font_dir(),
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../scripts/base-latin/sources"),
    ];
    for root in roots {
        let Ok(entries) = std::fs::read_dir(root) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|ext| ext.to_str()) != Some("ttf") {
                continue;
            }
            let Ok(bytes) = std::fs::read(&path) else {
                continue;
            };
            let Ok(face) = Face::parse(&bytes, 0) else {
                continue;
            };
            if face.glyph_index('你').is_some() {
                return Some(bytes);
            }
        }
    }
    None
}

fn page(primitives: Vec<Primitive>) -> DisplayPage {
    DisplayPage {
        page_index: 0,
        width: Number::from(816),
        height: Number::from(1056),
        content_bounds: None,
        column_bounds: Vec::new(),
        section_id: None,
        section_index: None,
        section_page_index: None,
        section_page_number: None,
        page_label: None,
        primitives,
        watermark_primitive_count: None,
        background: None,
        page_borders: Vec::new(),
        header: None,
        footer: None,
        note_areas: Vec::new(),
        reservations: Vec::new(),
        table_gridlines: Vec::new(),
        unbuilt: false,
        position_span: None,
        hf_parts: None,
    }
}

fn text_run(value: &str) -> TextRunPrimitive {
    TextRunPrimitive {
        text: value.to_owned(),
        paint_text: None,
        x: Number::from(72),
        baseline_y: Number::from(96),
        width: Number::from(200),
        paint_clip: None,
        font: "16px Roboto".to_owned(),
        scripts: None,
        color: "#112233".to_owned(),
        letter_spacing: None,
        word_spacing: None,
        rtl: None,
        opacity: None,
        rotation_deg: None,
        horizontal_scale: None,
        substitute_scale: None,
        all_caps: false,
        small_caps: false,
        hidden: false,
        text_shadow: None,
        text_outline: false,
        emphasis_mark: None,
        text_effect: None,
        attrs: DocAttrs::default(),
    }
}

fn glyph_run(font_id: u32, text: &str, bytes: &[u8]) -> GlyphRunPrimitive {
    let face = Face::parse(bytes, 0).expect("parse font");
    let mut glyphs = Vec::new();
    for (cluster, ch) in text.char_indices() {
        let id = face.glyph_index(ch).expect("glyph").0;
        glyphs.push(PlacedGlyph {
            id: u32::from(id),
            x: 72.0 + cluster as f64 * 8.0,
            y: 96.0,
            cluster: cluster as u32,
            advance: 8.0,
            logical_order: None,
            bidi_level: None,
        });
    }
    GlyphRunPrimitive {
        font_id,
        size: 16.0,
        color: "#112233".to_owned(),
        text: text.to_owned(),
        paint_text: None,
        glyphs,
        echo: None,
        paint_clip: None,
        word_spacing: None,
        rtl: None,
        opacity: None,
        rotation_deg: None,
        horizontal_scale: None,
        all_caps: false,
        small_caps: false,
        hidden: false,
        text_shadow: None,
        text_outline: false,
        emphasis_mark: None,
        text_effect: None,
        attrs: DocAttrs::default(),
    }
}

/// Build embedded fonts exactly as the exporter does.
fn embed_all(
    store: &FontStore,
    uses: &BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>>,
) -> BTreeMap<u32, font::EmbeddedFont> {
    let mut fonts = BTreeMap::new();
    for (index, (font_id, glyphs)) in uses.iter().enumerate() {
        let embedded = font::EmbeddedFont::new(format!("F{}", index + 1), store, *font_id, glyphs)
            .expect("embed font");
        assert!(!embedded.tounicode.is_empty(), "ToUnicode cmap is empty");
        fonts.insert(*font_id, embedded);
    }
    fonts
}

fn assert_strings(glyphs: &BTreeMap<u16, BTreeSet<String>>) {
    assert!(!glyphs.is_empty());
    for values in glyphs.values() {
        assert!(!values.is_empty(), "empty ToUnicode source string");
    }
}

fn collect_for(store: &FontStore, value: &str) -> BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> {
    let display_list = DisplayList {
        contract_version: None,
        pages: vec![page(vec![Primitive::Text(text_run(value))])],
    };
    font::collect_with_text(&display_list, store)
}

#[test]
fn fallback_collection_covers_cjk_emoji_and_hebrew() {
    let mut store = FontStore::default();
    let roboto = store
        .register(font_bytes("Roboto-Regular.ttf"))
        .expect("roboto");
    let hebrew = store
        .register(font_bytes("NotoSansHebrew-Regular.ttf"))
        .expect("hebrew");
    let cjk = cjk_font().map(|bytes| store.register(bytes).expect("cjk"));

    let cjk_out = collect_for(&store, "你好🙂");
    assert!(!cjk_out.is_empty());
    for glyphs in cjk_out.values() {
        assert_strings(glyphs);
    }
    if let Some(cjk) = cjk {
        assert!(
            cjk_out.contains_key(&cjk.to_u32()),
            "CJK fixture was registered but did not supply 你好"
        );
    } else {
        // No fixture covers these scalars: the first font's .notdef must still
        // be collectable so the painter can address it.
        assert!(cjk_out.contains_key(&roboto.to_u32()));
        assert!(cjk_out[&roboto.to_u32()].contains_key(&0));
    }
    let _ = embed_all(&store, &cjk_out);

    let hebrew_out = collect_for(&store, "שלום");
    assert!(hebrew_out.contains_key(&hebrew.to_u32()));
    for glyphs in hebrew_out.values() {
        assert_strings(glyphs);
    }
    let _ = embed_all(&store, &hebrew_out);

    let mixed = collect_for(&store, "A你🙂B");
    assert!(!mixed.is_empty());
    for glyphs in mixed.values() {
        assert_strings(glyphs);
    }

    // Every collected entry must yield a subset with a ToUnicode map.
    let _ = embed_all(&store, &mixed);
}

#[test]
fn missing_glyph_text_paints_without_unsupported() {
    let mut store = FontStore::default();
    let roboto = store
        .register(font_bytes("Roboto-Regular.ttf"))
        .expect("roboto");
    let text = text_run("A你🙂שלוםB");
    let page = page(vec![Primitive::Text(text.clone())]);
    let display_list = DisplayList {
        contract_version: None,
        pages: vec![page.clone()],
    };
    let uses = font::collect_with_text(&display_list, &store);
    assert!(uses.contains_key(&roboto.to_u32()));
    let fonts = embed_all(&store, &uses);
    let mut content = Content::new();
    text::paint_text(&mut content, &page, &text, &fonts, &store).expect("paint text");
    let stream = content.finish();
    assert!(!stream.is_empty());
    assert!(stream.windows(2).any(|window| window == b"Tj" as &[u8]));
}

#[test]
fn rotated_glyph_run_emits_rotated_text_matrix() {
    let mut store = FontStore::default();
    let bytes = font_bytes("Roboto-Regular.ttf");
    let roboto = store.register(bytes.clone()).expect("roboto");
    let mut run = glyph_run(roboto.to_u32(), "Hi", &bytes);
    run.rotation_deg = Number::from_f64(45.0);
    let mut uses: BTreeMap<u16, BTreeSet<String>> = BTreeMap::new();
    for glyph in &run.glyphs {
        uses.entry(glyph.id as u16)
            .or_default()
            .insert("H".to_owned());
    }
    let fonts = embed_all(&store, &BTreeMap::from([(roboto.to_u32(), uses)]));
    let page = page(vec![Primitive::GlyphRun(run.clone())]);
    let mut content = Content::new();
    text::paint_glyph_run(&mut content, &page, &run, &fonts).expect("rotated glyph run");
    let stream = String::from_utf8_lossy(&content.finish()).into_owned();
    assert!(stream.contains("Tm"));
    assert!(
        stream.contains("0.7071"),
        "rotated text matrix missing: {stream}"
    );
}

#[test]
fn missing_glyph_clusters_use_store_fallback_fonts() {
    let emoji_path = Path::new("/usr/share/fonts/noto/NotoColorEmoji.ttf");
    let Ok(emoji_bytes) = std::fs::read(emoji_path) else {
        eprintln!("skipping: {} is unavailable", emoji_path.display());
        return;
    };
    let mut store = FontStore::default();
    let bytes = font_bytes("Roboto-Regular.ttf");
    let roboto = store.register(bytes.clone()).expect("roboto");
    let emoji = store.register(emoji_bytes).expect("emoji");
    let face = Face::parse(&bytes, 0).expect("parse roboto");

    let mut glyphs = Vec::new();
    for (index, (cluster, _)) in "😀🎉🚀".char_indices().enumerate() {
        glyphs.push(PlacedGlyph {
            id: 0,
            x: 72.0 + index as f64 * 8.0,
            y: 96.0,
            cluster: cluster as u32,
            advance: 8.0,
            logical_order: None,
            bidi_level: None,
        });
    }
    glyphs.push(PlacedGlyph {
        id: u32::from(face.glyph_index('A').expect("A glyph").0),
        x: 96.0,
        y: 96.0,
        cluster: 12,
        advance: 8.0,
        logical_order: None,
        bidi_level: None,
    });
    let mut run = glyph_run(roboto.to_u32(), "A", &bytes);
    run.text = "😀🎉🚀A".to_owned();
    run.glyphs = glyphs;

    let page = page(vec![Primitive::GlyphRun(run.clone())]);
    let display_list = DisplayList {
        contract_version: None,
        pages: vec![page.clone()],
    };
    let uses = font::collect_with_text(&display_list, &store);
    let fallback = uses.get(&emoji.to_u32()).expect("emoji fallback font");
    for glyph in [883u16, 381, 963] {
        assert!(
            fallback.contains_key(&glyph),
            "fallback gid {glyph} missing: {fallback:?}"
        );
        assert!(!fallback[&glyph].is_empty());
    }

    let fonts = embed_all(&store, &uses);
    let mut content = Content::new();
    text::paint_glyph_run_with_store(&mut content, &page, &run, &fonts, &store)
        .expect("paint fallback glyphs");
    let stream = String::from_utf8_lossy(&content.finish()).into_owned();
    assert!(
        stream.contains("/F1"),
        "primary font not referenced: {stream}"
    );
    assert!(
        stream.contains("/F2"),
        "fallback font not referenced: {stream}"
    );

    let pdf = betteroffice_docx_pdf::export_pdf_with_fonts(&display_list, &store)
        .expect("whole export with emoji fallback");
    assert!(pdf.starts_with(b"%PDF-"));
}

#[test]
fn public_export_accepts_fallback_text_and_rotated_runs() {
    let mut store = FontStore::default();
    let bytes = font_bytes("Roboto-Regular.ttf");
    let roboto = store.register(bytes.clone()).expect("roboto");
    store
        .register(font_bytes("NotoSansHebrew-Regular.ttf"))
        .expect("hebrew");

    let mut run = glyph_run(roboto.to_u32(), "Hi", &bytes);
    run.rotation_deg = Number::from_f64(45.0);
    let display_list = DisplayList {
        contract_version: None,
        pages: vec![page(vec![
            Primitive::Text(text_run("A你🙂שלוםB")),
            Primitive::GlyphRun(run),
        ])],
    };

    let pdf = betteroffice_docx_pdf::export_pdf_with_fonts(&display_list, &store)
        .expect("export fallback text and rotated glyphs");
    assert!(pdf.starts_with(b"%PDF-"));
    let pdf = String::from_utf8_lossy(&pdf);
    assert!(pdf.contains("/FontFile2"));
    assert!(pdf.contains("/ToUnicode"));
    assert!(pdf.contains("/Identity-H"));
}

#[test]
fn one_glyph_with_two_texts_gets_two_cids() {
    let mut store = FontStore::default();
    let bytes = font_bytes("Roboto-Regular.ttf");
    let roboto = store.register(bytes).expect("roboto");
    let face_bytes = font_bytes("Roboto-Regular.ttf");
    let face = Face::parse(&face_bytes, 0).expect("face");
    let gid = u16::from(face.glyph_index('A').expect("A").0);
    let uses = BTreeMap::from([(gid, BTreeSet::from(["A".to_owned(), "B".to_owned()]))]);
    let font =
        font::EmbeddedFont::new("F1".to_owned(), &store, roboto.to_u32(), &uses).expect("embed");
    let a = font.cid(u32::from(gid), "A").expect("A cid");
    let b = font.cid(u32::from(gid), "B").expect("B cid");
    assert_ne!(a, b, "one glyph cannot carry two ToUnicode meanings");
    assert!(
        font.cid_to_gid.is_some(),
        "duplicate CIDs need a CIDToGIDMap"
    );
    let cmap = String::from_utf8_lossy(&font.tounicode);
    assert!(cmap.contains("0041") && cmap.contains("0042"), "{cmap}");
}

#[test]
fn cff_duplicate_texts_fall_back_to_one_cid() {
    let path =
        std::path::Path::new("/usr/share/fonts/adobe-source-han-sans/SourceHanSansCN-Normal.otf");
    let Ok(bytes) = std::fs::read(path) else {
        return;
    };
    let mut store = FontStore::default();
    let id = store.register(bytes).expect("cff font");
    let uses = BTreeMap::from([(1u16, BTreeSet::from(["A".to_owned(), "B".to_owned()]))]);
    let font =
        font::EmbeddedFont::new("F1".to_owned(), &store, id.to_u32(), &uses).expect("embed cff");
    assert!(font.is_cff);
    assert!(font.cid_to_gid.is_none(), "CFF has no CIDToGIDMap");
    assert_eq!(font.cid(1, "A"), font.cid(1, "B"));
}
