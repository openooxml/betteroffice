//! OpenType shaping via rustybuzz.
//!
//! One shaping call covers one run of uniform font/size/features. Script is
//! inferred from the text (rustybuzz's segment-property guess); callers that
//! already resolved bidi levels should pass the run direction explicitly via
//! [`shape_with_direction`]. Bidi reordering happens above this layer using
//! [`crate::bidi`] runs, so text handed in here is a single directional run.

use crate::font_store::{FontError, FontId, FontStore, MAX_CACHED_SHAPE_TEXT_BYTES, ShapeCacheKey};
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
        language: language.map(ToOwned::to_owned),
    });
    if let Some(glyphs) = cache_key.as_ref().and_then(|key| store.cached_shape(key)) {
        return Ok(glyphs);
    }
    let bytes = store.font_bytes(font)?;
    let face = rustybuzz::Face::from_slice(bytes, 0)
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

    let glyphs = rustybuzz::shape(&face, &features, buffer);

    let upem = store.metrics(font)?.units_per_em as f32;
    let scale = size / upem;

    let glyphs: Vec<_> = glyphs
        .glyph_infos()
        .iter()
        .zip(glyphs.glyph_positions())
        .map(|(info, pos)| ShapedGlyph {
            glyph_id: info.glyph_id,
            cluster: info.cluster,
            x_advance: pos.x_advance as f32 * scale,
            x_offset: pos.x_offset as f32 * scale,
            y_offset: pos.y_offset as f32 * scale,
        })
        .collect();
    if let Some(cache_key) = cache_key {
        store.cache_shape(cache_key, &glyphs);
    }
    Ok(glyphs)
}

#[cfg(test)]
mod tests {
    use super::*;

    const LIBERATION_SANS: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn repeated_exact_shapes_reuse_a_bounded_resident_entry() {
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
        assert_eq!(store.shape_cache_len(), 0);
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
        let third = shape_with_properties(
            &store,
            font,
            "resident shaping",
            16.0,
            &[],
            ShapeDirection::Ltr,
            Some("en-US"),
        )
        .unwrap();
        assert_eq!(third, first);
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
        assert_eq!(store.shape_cache_len(), 2);
    }
}
