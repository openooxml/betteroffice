//! A sliced CJK face registered from its skeleton and core gains outline
//! chunks in place: shaping is identical throughout, and after
//! `update_outlines` every shaped glyph draws as in the whole face.

use std::collections::BTreeSet;
use std::fs;
use std::path::PathBuf;

use betteroffice_fonts::Codec;
use betteroffice_fonts::slices::Skeleton;
use ooxml_text::{FontId, FontStore, OutlineUpdateError, shape};

fn assets() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../packages/fonts-cjk/assets")
}

fn decode(file: &str) -> Vec<u8> {
    let frame = fs::read(assets().join(file)).unwrap();
    betteroffice_fonts::decode_face(&frame, None, Codec::Zstd, 64 << 20)
        .unwrap()
        .font
}

struct Sliced {
    skeleton: Vec<u8>,
    core: Vec<u8>,
    chunks: Vec<(u16, Vec<u8>)>,
    whole: Vec<u8>,
}

fn sliced(stem: &str) -> Sliced {
    let index: serde_json::Value =
        serde_json::from_slice(&fs::read(assets().join(format!("{stem}/slices.json"))).unwrap())
            .unwrap();
    Sliced {
        skeleton: decode(index["skeleton"]["file"].as_str().unwrap()),
        core: decode(index["core"]["file"].as_str().unwrap()),
        chunks: index["chunks"]
            .as_array()
            .unwrap()
            .iter()
            .map(|chunk| {
                (
                    chunk["id"].as_u64().unwrap() as u16,
                    decode(chunk["file"].as_str().unwrap()),
                )
            })
            .collect(),
        whole: decode(&format!("{stem}.otf.zst")),
    }
}

impl Sliced {
    fn reassemble(&self, wanted: &BTreeSet<u16>) -> Vec<u8> {
        let skeleton = Skeleton::parse(&self.skeleton).unwrap();
        let mut chunks: Vec<&[u8]> = vec![&self.core];
        chunks.extend(
            self.chunks
                .iter()
                .filter(|(id, _)| wanted.contains(id))
                .map(|(_, bytes)| bytes.as_slice()),
        );
        skeleton.reassemble(&chunks).unwrap()
    }

    fn chunks_of(&self, glyphs: &[u32]) -> BTreeSet<u16> {
        let skeleton = Skeleton::parse(&self.skeleton).unwrap();
        glyphs
            .iter()
            .filter_map(|&glyph| skeleton.chunk_of_glyph(glyph))
            .filter(|&chunk| chunk != 0)
            .collect()
    }
}

fn glyph_ids(store: &FontStore, font: FontId, text: &str) -> Vec<u32> {
    shape(store, font, text, 16.0, &[])
        .unwrap()
        .iter()
        .map(|glyph| glyph.glyph_id)
        .collect()
}

fn assert_progressive(stem: &str, text: &str) {
    let face = sliced(stem);
    let mut whole_store = FontStore::new();
    let whole = whole_store.register(face.whole.clone()).unwrap();
    let mut store = FontStore::new();
    let font = store.register(face.reassemble(&BTreeSet::new())).unwrap();
    let substitute = store.register_named_substitute(font, "MS Mincho").unwrap();

    let before = shape(&store, font, text, 16.0, &[]).unwrap();
    assert_eq!(before, shape(&whole_store, whole, text, 16.0, &[]).unwrap());
    let glyphs = glyph_ids(&store, font, text);
    let wanted = face.chunks_of(&glyphs);
    assert!(
        !wanted.is_empty(),
        "{text} should need chunks beyond the core"
    );
    let missing = glyphs
        .iter()
        .filter(|&&glyph| {
            store.outline_glyph_json(font, glyph as u16).unwrap()
                != whole_store.outline_glyph_json(whole, glyph as u16).unwrap()
        })
        .count();
    assert!(missing > 0);

    let revision = store.outline_revision();
    store
        .update_outlines(substitute, face.reassemble(&wanted))
        .unwrap();
    assert_eq!(store.outline_revision(), revision + 1);
    assert_eq!(shape(&store, font, text, 16.0, &[]).unwrap(), before);
    for glyph in glyphs {
        let expected = whole_store.outline_glyph_json(whole, glyph as u16).unwrap();
        assert_eq!(
            store.outline_glyph_json(font, glyph as u16).unwrap(),
            expected
        );
        assert_eq!(
            store.outline_glyph_json(substitute, glyph as u16).unwrap(),
            store.outline_glyph_json(font, glyph as u16).unwrap()
        );
    }
}

#[test]
fn japanese_text_draws_as_the_whole_face_once_its_chunks_are_installed() {
    assert_progressive(
        "NotoSansJP-Regular",
        "鬱蒼とした森の奥に、齟齬と矛盾が潜む。",
    );
}

#[test]
fn composed_hangul_requests_the_chunk_of_the_shaped_syllable() {
    let face = sliced("NotoSansKR-Regular");
    let mut store = FontStore::new();
    let font = store.register(face.reassemble(&BTreeSet::new())).unwrap();
    let skeleton = Skeleton::parse(&face.skeleton).unwrap();
    let text = "\u{1112}\u{1175}\u{11C2}";
    let glyphs = glyph_ids(&store, font, text);
    let syllable = skeleton.chunk_of_codepoint(0xD7A3);
    assert!(
        glyphs
            .iter()
            .any(|&glyph| skeleton.chunk_of_glyph(glyph) == syllable),
        "the shaper composes the jamo, so the syllable's chunk is requested"
    );
    assert_progressive("NotoSansKR-Regular", "\u{1112}\u{1175}\u{11C2} 똠방각하 뷁");
}

#[test]
fn refuses_bytes_that_change_layout_tables_and_unknown_ids() {
    let face = sliced("NotoSansJP-Regular");
    let mut store = FontStore::new();
    let font = store.register(face.reassemble(&BTreeSet::new())).unwrap();
    let revision = store.outline_revision();
    let other = sliced("NotoSansKR-Regular").reassemble(&BTreeSet::new());
    assert_eq!(
        store.update_outlines(font, other),
        Err(OutlineUpdateError::TablesDiffer)
    );
    assert_eq!(
        store.update_outlines(FontId::from_u32(99), face.whole.clone()),
        Err(OutlineUpdateError::UnknownFont)
    );
    assert!(matches!(
        store.update_outlines(font, vec![0; 8]),
        Err(OutlineUpdateError::Parse(_))
    ));
    assert_eq!(store.outline_revision(), revision);
    let glyphs = glyph_ids(&store, font, "日本");
    assert_eq!(glyphs, glyph_ids(&store, font, "日本"));
}
