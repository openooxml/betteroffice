#![cfg(feature = "base-latin")]

use docx_layout::display_list::Primitive;
use ooxml_text::{FontId, FontStore};

#[path = "fixtures/base_latin.rs"]
mod fixture;

#[test]
fn fontless_paragraph_emits_backstop_glyphs() {
    let fonts = FontStore::new();
    for bold in [false, true] {
        let scene = fixture::measured_document(
            &fonts,
            &[FontId::base_latin(bold, false).to_u32()],
            "Latin Ελληνικά Кириллица ✓ ₹",
            bold,
        );
        let runs: Vec<_> = scene
            .pages
            .iter()
            .flat_map(|page| &page.primitives)
            .filter_map(|p| match p {
                Primitive::GlyphRun(run) => Some(run),
                _ => None,
            })
            .collect();
        assert!(!runs.is_empty());
        for run in &runs {
            assert!(FontId::from_u32(run.font_id).base_latin_style().is_some());
            assert!(
                run.attrs
                    .fallback_font
                    .as_ref()
                    .unwrap()
                    .contains(ooxml_text::BASE_LATIN_FAMILY)
            );
            assert!(!run.glyphs.is_empty());
            assert!(
                run.glyphs
                    .iter()
                    .all(|glyph| glyph.id != 0 && glyph.advance > 0.0)
            );
        }
        if bold {
            assert!(
                runs.iter()
                    .any(|run| run.font_id == FontId::base_latin(true, false).to_u32())
            );
            assert!(
                runs.iter()
                    .any(|run| run.font_id == FontId::base_latin(false, false).to_u32())
            );
        }
    }
    assert_eq!(fonts.font_count(), 0);
}

#[test]
fn nonempty_chain_uses_regular_backstop_for_uncovered_greek() {
    let mut fonts = FontStore::new();
    let real = fonts
        .register(include_bytes!("../../../packages/fonts/assets/NotoSansHebrew-Bold.ttf").to_vec())
        .unwrap();
    assert!(!fonts.covers(real, 'Ω').unwrap());
    let backstop = FontId::base_latin(false, false);
    assert_eq!(
        fonts.resolve_with_style(&[real], 'Ω', true, false),
        Some(backstop)
    );
    let scene = fixture::measured_document(&fonts, &[real.to_u32()], "Ω", true);
    let runs: Vec<_> = scene
        .pages
        .iter()
        .flat_map(|page| &page.primitives)
        .filter_map(|p| match p {
            Primitive::GlyphRun(run) => Some(run),
            _ => None,
        })
        .collect();
    assert!(!runs.is_empty());
    assert!(runs.iter().all(|run| run.font_id == backstop.to_u32()));
    assert!(
        runs.iter()
            .flat_map(|run| &run.glyphs)
            .all(|glyph| glyph.id != 0 && glyph.advance > 0.0)
    );
}

#[test]
fn a_covering_face_keeps_its_display_list_with_the_backstop_reachable() {
    let mut fonts = FontStore::new();
    let real = fonts
        .register(
            include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf").to_vec(),
        )
        .unwrap();
    let text = "AV Latin Ελληνικά Кириллица";
    let before = fixture::measured_document(&fonts, &[real.to_u32()], text, false);
    let after = fixture::measured_document(
        &fonts,
        &[real.to_u32(), FontId::base_latin(false, false).to_u32()],
        text,
        false,
    );
    assert_eq!(before, after);
    assert!(
        before
            .pages
            .iter()
            .flat_map(|page| &page.primitives)
            .any(|p| matches!(p, Primitive::GlyphRun(_)))
    );
}
