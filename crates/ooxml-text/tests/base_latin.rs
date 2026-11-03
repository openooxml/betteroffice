#![cfg(feature = "base-latin")]

use ooxml_text::{FontId, FontStore, measure::measure_paragraph_json, shape};
use serde_json::{Value, json};

const REAL_FACE: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");

fn measure(store: &FontStore, chain: &[u32], text: &str) -> Value {
    let input = json!({
        "block": { "kind": "paragraph", "runs": [{ "kind": "text", "text": text }] },
        "maxWidth": 600,
        "fontChains": { "missing|0|0": chain },
        "defaults": { "fontSize": 11, "fontFamily": "Missing" }
    });
    serde_json::from_str(&measure_paragraph_json(store, &input.to_string()).unwrap()).unwrap()
}

#[test]
fn explicit_backstop_measures_and_shapes_without_registered_faces() {
    let store = FontStore::new();
    let backstop = FontId::base_latin(false, false);
    let text = "Latin Ελληνικά Кириллица ✓ ₹";
    for ch in text.chars() {
        assert_eq!(store.resolve(&[backstop], ch), Some(backstop));
    }
    let glyphs = shape(&store, backstop, text, 11.0, &[]).unwrap();
    assert!(!glyphs.is_empty());
    assert!(glyphs.iter().all(|glyph| glyph.glyph_id != 0));
    assert!(
        measure(&store, &[backstop.to_u32()], text)["lines"][0]["width"]
            .as_f64()
            .unwrap()
            > 0.0
    );
    assert_eq!(store.font_count(), 0);
}

#[test]
fn chain_style_selects_bold_latin_and_regular_greek() {
    let mut store = FontStore::new();
    let bold = store
        .register(include_bytes!("../../../packages/fonts/assets/NotoSansHebrew-Bold.ttf").to_vec())
        .unwrap();
    let latin = ('A'..='Z')
        .chain('\u{00a0}'..='\u{017f}')
        .find(|&ch| {
            store.covers(FontId::base_latin(true, false), ch).unwrap()
                && !store.covers(bold, ch).unwrap()
        })
        .unwrap();
    assert!(!store.covers(bold, 'Ω').unwrap());
    assert_eq!(
        store.resolve_with_style(&[bold], latin, true, false),
        Some(FontId::base_latin(true, false))
    );
    assert_eq!(
        store.resolve_with_style(&[bold], 'Ω', true, false),
        Some(FontId::base_latin(false, false))
    );
    assert_eq!(
        store.resolve_with_style(&[bold], latin, true, true),
        Some(FontId::base_latin(true, true))
    );
}

#[test]
fn reserved_faces_share_bytes_and_support_views_without_shifting_host_ids() {
    let mut store = FontStore::new();
    let other = FontStore::new();
    for (bold, italic, bytes) in base_latin::faces() {
        let id = FontId::base_latin(bold, italic);
        assert_eq!(id.base_latin_style(), Some((bold, italic)));
        assert!(std::ptr::eq(store.font_bytes(id).unwrap(), bytes));
        assert!(std::ptr::eq(
            store.font_bytes(id).unwrap(),
            other.font_bytes(id).unwrap()
        ));
        assert_eq!(store.byte_owner(id).unwrap(), id);
        let flags = store.metrics(id).unwrap().os2_fs_selection;
        assert_eq!(flags & (1 << 5) != 0, bold);
        assert_eq!(flags & 1 != 0, italic);
        let glyph = store.glyph_id(id, 'A').unwrap().unwrap();
        assert!(!store.outline_glyph(id, glyph).unwrap().cmds.is_empty());
        assert!(
            shape(&store, id, "Ab", 11.0, &[])
                .unwrap()
                .iter()
                .all(|g| g.glyph_id != 0)
        );
    }
    assert_eq!(store.register(REAL_FACE.to_vec()).unwrap().to_u32(), 0);
    let base = FontId::base_latin(false, false);
    let view = store
        .register_named_substitute(base, "Arial Narrow")
        .unwrap();
    assert_eq!(view.to_u32(), 1);
    assert_eq!(store.byte_owner(view).unwrap(), base);
    let glyph = store.glyph_id(view, 'A').unwrap().unwrap();
    assert!(!store.outline_glyph(view, glyph).unwrap().cmds.is_empty());
    assert!(
        shape(&store, view, "Ab", 11.0, &[])
            .unwrap()
            .iter()
            .all(|g| g.x_advance > 0.0)
    );
    assert_eq!(store.register(REAL_FACE.to_vec()).unwrap().to_u32(), 2);
}

#[test]
fn appending_backstop_preserves_a_covering_faces_measurement_and_shaping() {
    let mut store = FontStore::new();
    let real = store.register(REAL_FACE.to_vec()).unwrap();
    let text = "AV Latin Ελληνικά Кириллица";
    let before = measure(&store, &[real.to_u32()], text);
    let shaped = shape(&store, real, text, 11.0, &[]).unwrap();
    let backstop = store
        .register(base_latin::face(false, false).to_vec())
        .unwrap();
    for ch in text.chars() {
        assert_eq!(store.resolve(&[real, backstop], ch), Some(real));
    }
    assert_eq!(
        before,
        measure(&store, &[real.to_u32(), backstop.to_u32()], text)
    );
    assert_eq!(shaped, shape(&store, real, text, 11.0, &[]).unwrap());
}

#[test]
fn characters_no_face_covers_keep_the_terminal_faces_notdef() {
    let mut store = FontStore::new();
    let real = store.register(REAL_FACE.to_vec()).unwrap();
    assert_eq!(
        store.resolve_with_style(&[real], '中', false, false),
        Some(real)
    );
    assert_eq!(store.resolve_with_style(&[], '中', true, false), None);
}

#[test]
fn empty_chains_do_not_resolve_to_backstop_faces() {
    let store = FontStore::new();
    assert_eq!(store.resolve(&[], 'A'), None);
    assert_eq!(store.resolve_with_style(&[], 'A', true, true), None);
    for chains in [json!({}), json!({"missing|0|0": []})] {
        let input = json!({
            "block": {"kind": "paragraph", "runs": [{"kind": "text", "text": "Latin"}]},
            "maxWidth": 600,
            "fontChains": chains,
            "defaults": {"fontFamily": "Missing", "fontSize": 11}
        });
        assert!(measure_paragraph_json(&store, &input.to_string()).is_err());
    }
}
