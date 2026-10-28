//! OpenType shaping via rustybuzz.
//!
//! One shaping call covers one run of uniform font/size/features. Script is
//! inferred from the text (rustybuzz's segment-property guess); callers that
//! already resolved bidi levels should pass the run direction explicitly via
//! [`shape_with_direction`]. Bidi reordering happens above this layer using
//! [`crate::bidi`] runs, so text handed in here is a single directional run.

use crate::family_metrics::{FamilyMetrics, family_advance};
use crate::font_store::{FontError, FontId, FontStore, MAX_CACHED_SHAPE_TEXT_BYTES, ShapeCacheKey};
use std::collections::BTreeMap;
use std::str::FromStr;

/// One positioned glyph produced by shaping.
///
/// `cluster` is the byte index into the input text of the character (cluster)
/// this glyph belongs to — the glyph↔char mapping the display-list
/// `cluster_map` and hit-testing are built from.
///
/// Advances and offsets are scaled to `size` (same unit the caller passed,
/// e.g. CSS px or twips): `value_font_units * size / units_per_em`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapedGlyph {
    pub glyph_id: u32,
    pub cluster: u32,
    pub x_advance: f32,
    pub x_offset: f32,
    pub y_offset: f32,
}

/// An OpenType feature setting, e.g. `(*b"liga", 0)` to disable ligatures or
/// `(*b"smcp", 1)` for small caps. Applied over the whole run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeFeature {
    pub tag: [u8; 4],
    pub value: u32,
}

/// Direction to use when shaping a single bidi-level run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShapeDirection {
    /// Let rustybuzz infer from script, preserving the historical `shape` API.
    #[default]
    Auto,
    Ltr,
    Rtl,
}

/// Shape `text` with the given registered font, returning glyphs in visual
/// order with advances/offsets scaled to `size`.
///
/// Kerning and default OpenType features (GSUB ligatures, GPOS kerning) are
/// applied by rustybuzz; `features` overrides/extends them.
pub fn shape(
    store: &FontStore,
    font: FontId,
    text: &str,
    size: f32,
    features: &[ShapeFeature],
) -> Result<Vec<ShapedGlyph>, FontError> {
    shape_with_direction(store, font, text, size, features, ShapeDirection::Auto)
}

/// Shape `text` with an explicit resolved direction.
///
/// Use this for UBA level runs. It keeps neutral-only or script-mixed slices
/// from depending on rustybuzz's script-based direction guess.
pub fn shape_with_direction(
    store: &FontStore,
    font: FontId,
    text: &str,
    size: f32,
    features: &[ShapeFeature],
    direction: ShapeDirection,
) -> Result<Vec<ShapedGlyph>, FontError> {
    shape_with_properties(store, font, text, size, features, direction, None)
}

/// Shape with explicit direction and an optional BCP-47 language tag.
///
/// Language-sensitive OpenType substitutions (localized forms, for example)
/// must receive the same run language during measurement and display-list
/// shaping. Invalid/empty tags are ignored rather than reaching a font parser.
pub fn shape_with_properties(
    store: &FontStore,
    font: FontId,
    text: &str,
    size: f32,
    features: &[ShapeFeature],
    direction: ShapeDirection,
    language: Option<&str>,
) -> Result<Vec<ShapedGlyph>, FontError> {
    let cache_key = (text.len() <= MAX_CACHED_SHAPE_TEXT_BYTES).then(|| ShapeCacheKey {
        font,
        text: text.to_owned(),
        size_bits: size.to_bits(),
        features: features
            .iter()
            .map(|feature| (feature.tag, feature.value))
            .collect(),
        direction: match direction {
            ShapeDirection::Auto => 0,
            ShapeDirection::Ltr => 1,
            ShapeDirection::Rtl => 2,
        },
        // Over-long tags are ignored by shaping below, so they are keyed as
        // absent rather than retained verbatim in cache keys.
        language: language
            .filter(|tag| tag.len() <= 128)
            .map(ToOwned::to_owned),
    });
    if let Some(glyphs) = cache_key.as_ref().and_then(|key| store.cached_shape(key)) {
        return Ok(glyphs);
    }
    let face = store
        .shaping_face(font)?
        .ok_or_else(|| FontError::Parse("rustybuzz rejected font bytes".to_string()))?;

    let mut buffer = rustybuzz::UnicodeBuffer::new();
    buffer.push_str(text);
    match direction {
        ShapeDirection::Auto => {}
        ShapeDirection::Ltr => buffer.set_direction(rustybuzz::Direction::LeftToRight),
        ShapeDirection::Rtl => buffer.set_direction(rustybuzz::Direction::RightToLeft),
    }
    if let Some(language) = language
        && language.len() <= 128
        && let Ok(language) = rustybuzz::Language::from_str(language)
    {
        buffer.set_language(language);
    }

    let features: Vec<rustybuzz::Feature> = features
        .iter()
        .map(|f| {
            rustybuzz::Feature::new(rustybuzz::ttf_parser::Tag::from_bytes(&f.tag), f.value, ..)
        })
        .collect();

    let glyphs = rustybuzz::shape(face, &features, buffer);

    let upem = store.metrics(font)?.units_per_em as f32;
    let scale = size / upem;
    let horizontal = scale * store.advance_scale(font)?;

    let mut glyphs: Vec<_> = glyphs
        .glyph_infos()
        .iter()
        .zip(glyphs.glyph_positions())
        .map(|(info, pos)| ShapedGlyph {
            glyph_id: info.glyph_id,
            cluster: info.cluster,
            x_advance: pos.x_advance as f32 * horizontal,
            x_offset: pos.x_offset as f32 * horizontal,
            y_offset: pos.y_offset as f32 * scale,
        })
        .collect();
    // rustybuzz 0.20.1 reverses a backward buffer for the legacy `kern` table
    // and then skips the un-reverse when kerning was turned off, so a `kern=0`
    // RTL run can come back in logical order. Visual order is the contract.
    if direction == ShapeDirection::Rtl
        && let (Some(first), Some(last)) = (glyphs.first(), glyphs.last())
        && first.cluster < last.cluster
    {
        glyphs.reverse();
    }
    if let Some(widths) = store.family_widths(font)? {
        apply_family_widths(&mut glyphs, text, size, widths);
    }
    if let Some(cache_key) = cache_key {
        store.cache_shape(cache_key, &glyphs);
    }
    Ok(glyphs)
}

fn apply_family_widths(glyphs: &mut [ShapedGlyph], text: &str, size: f32, widths: &FamilyMetrics) {
    let mut clusters: BTreeMap<u32, (bool, Option<f32>)> = BTreeMap::new();
    for glyph in glyphs.iter() {
        let (has_advance, _) = clusters.entry(glyph.cluster).or_insert((false, None));
        *has_advance |= glyph.x_advance != 0.0;
    }
    let offsets: Vec<_> = clusters
        .keys()
        .map(|&offset| offset as usize)
        .chain(std::iter::once(text.len()))
        .collect();
    for bounds in offsets.windows(2) {
        if !clusters[&(bounds[0] as u32)].0 {
            continue;
        }
        let advance = text.get(bounds[0]..bounds[1]).and_then(|cluster| {
            cluster
                .chars()
                .filter(|&ch| !is_default_ignorable(ch))
                .try_fold(0.0, |total, ch| {
                    family_advance(widths, ch).map(|advance| total + (advance * 1000.0).round())
                })
        });
        clusters.insert(
            bounds[0] as u32,
            (true, advance.map(|advance| advance * (size / 1000.0))),
        );
    }
    for glyph in glyphs {
        if let Some((_, Some(advance))) = clusters.get_mut(&glyph.cluster) {
            glyph.x_advance = *advance;
            *advance = 0.0;
        }
    }
}

fn is_default_ignorable(ch: char) -> bool {
    matches!(
        ch as u32,
        0x00AD
            | 0x034F
            | 0x061C
            | 0x115F..=0x1160
            | 0x17B4..=0x17B5
            | 0x180B..=0x180F
            | 0x200B..=0x200F
            | 0x202A..=0x202E
            | 0x2060..=0x206F
            | 0x3164
            | 0xFE00..=0xFE0F
            | 0xFEFF
            | 0xFFA0
            | 0xFFF0..=0xFFF8
            | 0x1BCA0..=0x1BCA3
            | 0x1D173..=0x1D17A
            | 0xE0000..=0xE0FFF
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    const LIBERATION_SANS: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn mixed_clusters_skip_default_ignorables_but_keep_visible_hyphens() {
        let widths = crate::family_metrics::family_metrics("arial narrow", false, false).unwrap();
        let expected =
            (family_advance(widths, 'f').unwrap() + family_advance(widths, 'i').unwrap()) * 32.0;
        for ch in [
            '\u{00AD}',
            '\u{034F}',
            '\u{061C}',
            '\u{115F}',
            '\u{1160}',
            '\u{17B4}',
            '\u{17B5}',
            '\u{180B}',
            '\u{180F}',
            '\u{200B}',
            '\u{200F}',
            '\u{202A}',
            '\u{202E}',
            '\u{2060}',
            '\u{206F}',
            '\u{3164}',
            '\u{FE00}',
            '\u{FE0F}',
            '\u{FEFF}',
            '\u{FFA0}',
            '\u{FFF0}',
            '\u{FFF8}',
            '\u{1BCA0}',
            '\u{1BCA3}',
            '\u{1D173}',
            '\u{1D17A}',
            '\u{E0000}',
            '\u{E0FFF}',
            '-',
        ] {
            let mut glyphs = [
                ShapedGlyph {
                    glyph_id: 1,
                    cluster: 0,
                    x_advance: 10.0,
                    x_offset: 0.0,
                    y_offset: 0.0,
                },
                ShapedGlyph {
                    glyph_id: 2,
                    cluster: 0,
                    x_advance: 0.0,
                    x_offset: 0.0,
                    y_offset: 0.0,
                },
            ];
            apply_family_widths(&mut glyphs, &format!("f{ch}i"), 32.0, widths);
            let expected = expected
                + if ch == '-' {
                    family_advance(widths, ch).unwrap() * 32.0
                } else {
                    0.0
                };
            assert!((glyphs[0].x_advance - expected).abs() < 1e-4, "{ch:?}");
            assert_eq!(glyphs[1].x_advance, 0.0);
        }
    }

    #[test]
    fn table_clusters_assign_the_total_to_the_first_visual_glyph() {
        let widths = crate::family_metrics::family_metrics("arial narrow", false, false).unwrap();
        for rtl in [false, true] {
            let mut glyphs: Vec<_> = [0, 0, 2, 3]
                .into_iter()
                .map(|cluster| ShapedGlyph {
                    glyph_id: 1,
                    cluster,
                    x_advance: 7.0,
                    x_offset: 2.0,
                    y_offset: 3.0,
                })
                .collect();
            if rtl {
                glyphs.reverse();
            }
            apply_family_widths(&mut glyphs, "fiA\u{03A9}", 32.0, widths);
            let cluster: Vec<_> = glyphs.iter().filter(|glyph| glyph.cluster == 0).collect();
            let expected = (family_advance(widths, 'f').unwrap()
                + family_advance(widths, 'i').unwrap())
                * 32.0;
            assert!((cluster[0].x_advance - expected).abs() < 1e-4);
            assert_eq!(cluster[1].x_advance, 0.0);
            assert_eq!(
                glyphs
                    .iter()
                    .find(|glyph| glyph.cluster == 3)
                    .unwrap()
                    .x_advance,
                7.0
            );
            assert!(
                glyphs
                    .iter()
                    .all(|glyph| glyph.x_offset == 2.0 && glyph.y_offset == 3.0)
            );
        }
    }

    #[test]
    fn first_shape_caches_and_repeats_reuse_the_resident_entry() {
        let mut store = FontStore::new();
        let font = store.register(LIBERATION_SANS.to_vec()).unwrap();
        let first = shape_with_properties(
            &store,
            font,
            "resident shaping",
            16.0,
            &[],
            ShapeDirection::Ltr,
            Some("en-US"),
        )
        .unwrap();
        assert_eq!(
            store.shape_cache_len(),
            1,
            "a run caches on the first shape"
        );
        let second = shape_with_properties(
            &store,
            font,
            "resident shaping",
            16.0,
            &[],
            ShapeDirection::Ltr,
            Some("en-US"),
        )
        .unwrap();
        assert_eq!(second, first);
        assert_eq!(store.shape_cache_len(), 1);

        shape_with_direction(
            &store,
            font,
            "resident shaping",
            16.0,
            &[],
            ShapeDirection::Rtl,
        )
        .unwrap();
        assert_eq!(
            store.shape_cache_len(),
            2,
            "a different direction is a distinct entry"
        );
    }
}
