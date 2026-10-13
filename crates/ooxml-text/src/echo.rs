//! Typing echo model: the shaping facts a host needs to place glyphs exactly
//! where [`crate::shape`] would put them for simple left-to-right text, so a
//! keystroke can paint before the engine relays out.
//!
//! For one font at one shaping size, each alphabet character gets its isolated
//! glyph and advance in font units, and every ordered pair of characters gets
//! the advance change shaping applies to the left one (pair kerning). Pairs
//! whose shaping is not two independent glyphs (ligatures, contextual forms,
//! glyph offsets) are reported unsafe. A host's advance in pixels is then
//! `(advance + pair) as f32 * horizontal`, the same product shaping computes.

use std::collections::HashMap;

use serde::Serialize;

use crate::font_store::{FontError, FontId, FontStore};
use crate::shape::{ShapeDirection, ShapedGlyph, shape_with_direction};
use crate::word_metrics::kern_features;

/// The echo model of one font at one shaping size. Indices refer to the
/// alphabet's characters in order.
#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EchoGlyphSet {
    /// Pixels per advance unit, as shaping scales advances.
    pub horizontal: f32,
    /// Isolated glyph per character; 0 when the font cannot echo it.
    pub glyph_ids: Vec<u32>,
    /// Isolated advance per character, in font units.
    pub advances: Vec<i32>,
    /// `[left, right, units]` for every pair whose left advance changes.
    pub kerning: Vec<[i32; 3]>,
    /// `[left, right]` pairs that do not shape as two independent glyphs.
    pub unsafe_pairs: Vec<[u32; 2]>,
}

/// Builds the echo model of `font` at `size` for `alphabet`, shaped with the
/// kerning gate `kern` exactly as a run with that gate shapes.
pub fn echo_glyph_set(
    store: &FontStore,
    font: FontId,
    size: f32,
    kern: bool,
    alphabet: &[char],
) -> Result<EchoGlyphSet, FontError> {
    let features = kern_features(kern);
    let upem = store.metrics(font)?.units_per_em as f32;
    let scale = size / upem;
    let horizontal = scale * store.advance_scale(font)?;
    let shape =
        |text: &str| shape_with_direction(store, font, text, size, &features, ShapeDirection::Ltr);
    let units = |glyph: &ShapedGlyph| {
        let value = (glyph.x_advance / horizontal).round() as i32;
        (horizontal > 0.0
            && glyph.x_offset == 0.0
            && glyph.y_offset == 0.0
            && value as f32 * horizontal == glyph.x_advance)
            .then_some(value)
    };

    let mut glyph_ids = vec![0; alphabet.len()];
    let mut advances = vec![0; alphabet.len()];
    for (index, ch) in alphabet.iter().enumerate() {
        if !store.covers(font, *ch)? {
            continue;
        }
        let glyphs = shape(&ch.to_string())?;
        if let [glyph] = glyphs.as_slice()
            && glyph.glyph_id != 0
            && let Some(value) = units(glyph)
        {
            glyph_ids[index] = glyph.glyph_id;
            advances[index] = value;
        }
    }
    let ligature_starts = substitution_barriers(store, font, alphabet, &mut glyph_ids, kern)?;
    let supported: Vec<usize> = (0..alphabet.len())
        .filter(|&index| glyph_ids[index] != 0)
        .collect();

    let mut pairs: HashMap<(usize, usize), Option<i32>> = HashMap::new();
    if !supported.is_empty() {
        let sequence: Vec<usize> = pair_sequence(supported.len())
            .into_iter()
            .map(|index| supported[index])
            .collect();
        let mut text = String::with_capacity(sequence.len() * 2);
        let mut offsets = Vec::with_capacity(sequence.len());
        for &index in &sequence {
            offsets.push(text.len() as u32);
            text.push(alphabet[index]);
        }
        let glyphs = shape(&text)?;
        let mut clusters: HashMap<u32, Vec<&ShapedGlyph>> = HashMap::new();
        for glyph in &glyphs {
            clusters.entry(glyph.cluster).or_default().push(glyph);
        }
        let shaped: Vec<Option<i32>> = sequence
            .iter()
            .zip(&offsets)
            .map(
                |(&index, offset)| match clusters.get(offset).map(Vec::as_slice) {
                    Some([glyph]) if glyph.glyph_id == glyph_ids[index] => units(glyph),
                    _ => None,
                },
            )
            .collect();
        for position in 0..sequence.len() - 1 {
            let delta = shaped[position]
                .zip(shaped[position + 1])
                .map(|(value, _)| value - advances[sequence[position]]);
            record(
                &mut pairs,
                (sequence[position], sequence[position + 1]),
                delta,
            );
        }
    }

    let mut kerning = Vec::new();
    let mut unsafe_pairs = Vec::new();
    let mut keys: Vec<_> = pairs.keys().copied().collect();
    keys.sort_unstable();
    for key in keys {
        match pairs[&key] {
            Some(0) => {}
            Some(delta) => kerning.push([key.0 as i32, key.1 as i32, delta]),
            None => unsafe_pairs.push([key.0 as u32, key.1 as u32]),
        }
    }
    for (left, &unsafe_start) in ligature_starts.iter().enumerate() {
        if unsafe_start {
            for right in 0..alphabet.len() {
                unsafe_pairs.push([left as u32, right as u32]);
            }
        }
    }
    unsafe_pairs.sort_unstable();
    unsafe_pairs.dedup();
    Ok(EchoGlyphSet {
        horizontal,
        glyph_ids,
        advances,
        kerning,
        unsafe_pairs,
    })
}

fn substitution_barriers(
    store: &FontStore,
    font: FontId,
    alphabet: &[char],
    glyph_ids: &mut [u32],
    kern: bool,
) -> Result<Vec<bool>, FontError> {
    use rustybuzz::ttf_parser::{GlyphId, gsub::SubstitutionSubtable};
    use std::collections::BTreeSet;

    let mut ligatures = vec![false; alphabet.len()];
    let Some(face) = store.shaping_face(font)? else {
        glyph_ids.fill(0);
        return Ok(ligatures);
    };
    if face.tables().morx.is_some() || face.tables().trak.is_some() {
        glyph_ids.fill(0);
        return Ok(ligatures);
    }
    let Some(table) = face.tables().gsub else {
        return Ok(ligatures);
    };
    if table.variations.is_some() {
        glyph_ids.fill(0);
        return Ok(ligatures);
    }
    let nominal: Vec<_> = alphabet
        .iter()
        .map(|&ch| store.glyph_id(font, ch))
        .collect::<Result<_, _>>()?;
    let mut lookups = BTreeSet::new();
    let mut required = BTreeSet::new();
    for index in 0..table.scripts.len() {
        if let Some(script) = table.scripts.get(index) {
            if let Some(feature) = script
                .default_language
                .and_then(|language| language.required_feature)
            {
                required.insert(feature);
            }
            for index in 0..script.languages.len() {
                if let Some(feature) = script
                    .languages
                    .get(index)
                    .and_then(|language| language.required_feature)
                {
                    required.insert(feature);
                }
            }
        }
    }
    for index in 0..table.features.len() {
        let Some(feature) = table.features.get(index) else {
            continue;
        };
        let active = matches!(
            &feature.tag.to_bytes(),
            b"ccmp"
                | b"locl"
                | b"rlig"
                | b"calt"
                | b"clig"
                | b"liga"
                | b"rclt"
                | b"rvrn"
                | b"abvm"
                | b"blwm"
                | b"mark"
                | b"mkmk"
                | b"curs"
                | b"dist"
                | b"ltra"
                | b"ltrm"
                | b"rand"
                | b"trak"
                | b"Harf"
                | b"HARF"
                | b"Buzz"
                | b"BUZZ"
        ) || alphabet.contains(&'\u{2044}')
            && matches!(&feature.tag.to_bytes(), b"frac" | b"numr" | b"dnom")
            || kern && feature.tag.to_bytes() == *b"kern"
            || required.contains(&index);
        if active {
            lookups.extend(feature.lookup_indices);
        }
    }
    for index in lookups {
        let Some(lookup) = table.lookups.get(index) else {
            glyph_ids.fill(0);
            return Ok(ligatures);
        };
        for index in 0..lookup.subtables.len() {
            let Some(subtable) = lookup.subtables.get::<SubstitutionSubtable<'_>>(index) else {
                glyph_ids.fill(0);
                return Ok(ligatures);
            };
            for (index, nominal) in nominal.iter().enumerate() {
                let covered = nominal.is_some_and(|id| subtable.coverage().contains(GlyphId(id)))
                    || glyph_ids[index] != 0
                        && subtable
                            .coverage()
                            .contains(GlyphId(glyph_ids[index] as u16));
                if !covered {
                    continue;
                }
                match subtable {
                    SubstitutionSubtable::Ligature(table) => {
                        if table.ligature_sets.into_iter().any(|set| {
                            set.into_iter()
                                .any(|ligature| ligature.components.is_empty())
                        }) {
                            glyph_ids[index] = 0;
                        } else {
                            ligatures[index] = true;
                        }
                    }
                    SubstitutionSubtable::Context(_)
                    | SubstitutionSubtable::ChainContext(_)
                    | SubstitutionSubtable::ReverseChainSingle(_) => {
                        glyph_ids.fill(0);
                        return Ok(ligatures);
                    }
                    _ => glyph_ids[index] = 0,
                }
            }
        }
    }
    Ok(ligatures)
}

/// A sequence over `count` symbols in which every ordered pair occurs exactly
/// once: the Lyndon words of length one and two in order, closed by its first
/// symbol (a linear de Bruijn sequence of order two).
fn pair_sequence(count: usize) -> Vec<usize> {
    let mut sequence = Vec::with_capacity(count * count + 1);
    for left in 0..count {
        sequence.push(left);
        for right in left + 1..count {
            sequence.push(left);
            sequence.push(right);
        }
    }
    if let Some(&first) = sequence.first() {
        sequence.push(first);
    }
    sequence
}

/// Keeps a pair's advance change only while every occurrence agrees.
fn record(
    pairs: &mut HashMap<(usize, usize), Option<i32>>,
    key: (usize, usize),
    value: Option<i32>,
) {
    let entry = pairs.entry(key).or_insert(value);
    if *entry != value {
        *entry = None;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const LIBERATION_SANS: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

    fn store() -> (FontStore, FontId) {
        let mut store = FontStore::new();
        let font = store.register(LIBERATION_SANS.to_vec()).unwrap();
        (store, font)
    }

    fn advance(set: &EchoGlyphSet, alphabet: &[char], text: &str) -> Vec<f32> {
        let chars: Vec<usize> = text
            .chars()
            .map(|ch| alphabet.iter().position(|&a| a == ch).unwrap())
            .collect();
        chars
            .iter()
            .enumerate()
            .map(|(position, &index)| {
                let pair = chars.get(position + 1).map_or(0, |&right| {
                    set.kerning
                        .iter()
                        .find(|entry| entry[0] == index as i32 && entry[1] == right as i32)
                        .map_or(0, |entry| entry[2])
                });
                (set.advances[index] + pair) as f32 * set.horizontal
            })
            .collect()
    }

    #[test]
    fn echo_advances_equal_shaped_advances_with_and_without_kerning() {
        let (store, font) = store();
        let alphabet: Vec<char> = "AVTRo.ywe xu".chars().collect();
        for kern in [false, true] {
            let set = echo_glyph_set(&store, font, 13.333_333, kern, &alphabet).unwrap();
            assert!(set.glyph_ids.iter().all(|&id| id != 0));
            assert_eq!(set.kerning.is_empty(), !kern);
            for text in ["AVATAR", "To. you", "wey xAV", "oVoToy"] {
                let shaped: Vec<f32> = shape_with_direction(
                    &store,
                    font,
                    text,
                    13.333_333,
                    &kern_features(kern),
                    ShapeDirection::Ltr,
                )
                .unwrap()
                .iter()
                .map(|glyph| glyph.x_advance)
                .collect();
                assert_eq!(advance(&set, &alphabet, text), shaped, "{text} kern={kern}");
            }
        }
    }

    #[test]
    fn pair_sequence_holds_every_ordered_pair_once() {
        for count in [1, 2, 3, 7] {
            let sequence = pair_sequence(count);
            assert_eq!(sequence.len(), count * count + 1);
            let mut seen: Vec<_> = sequence.windows(2).map(|pair| (pair[0], pair[1])).collect();
            seen.sort_unstable();
            seen.dedup();
            assert_eq!(seen.len(), count * count);
        }
    }

    #[test]
    fn ordinary_digits_do_not_enable_fraction_substitutions() {
        let (store, font) = store();
        let set = echo_glyph_set(&store, font, 16.0, false, &['0', '1', '2', '/', '3']).unwrap();
        assert!(set.glyph_ids.iter().all(|&id| id != 0));
    }

    #[test]
    fn uncovered_characters_and_ligatures_are_not_echoed() {
        let (store, font) = store();
        let alphabet: Vec<char> = "fi\u{4e00}".chars().collect();
        let set = echo_glyph_set(&store, font, 16.0, false, &alphabet).unwrap();
        assert_eq!(set.glyph_ids[2], 0);
        let ligated = shape_with_direction(
            &store,
            font,
            "fi",
            16.0,
            &kern_features(false),
            ShapeDirection::Ltr,
        )
        .unwrap()
        .len()
            == 1;
        assert_eq!(set.unsafe_pairs.contains(&[0, 1]), ligated);
    }

    #[test]
    fn three_component_ligature_without_pair_ligatures_is_not_echoed() {
        let (store, font) = store();
        let f = store.glyph_id(font, 'f').unwrap().unwrap();
        let i = store.glyph_id(font, 'i').unwrap().unwrap();
        let ligature = store.glyph_id(font, 'A').unwrap().unwrap();
        let mut gsub = Vec::new();
        for value in [1u16, 0, 10, 30, 44, 1] {
            gsub.extend(value.to_be_bytes());
        }
        gsub.extend(b"DFLT");
        for value in [8u16, 4, 0, 0, 0xffff, 1, 0, 1] {
            gsub.extend(value.to_be_bytes());
        }
        gsub.extend(b"liga");
        for value in [
            8u16, 0, 1, 0, 1, 4, 4, 0, 1, 8, 1, 20, 1, 8, 1, 4, ligature, 3, f, i, 1, 1, f,
        ] {
            gsub.extend(value.to_be_bytes());
        }
        let mut bytes = LIBERATION_SANS.to_vec();
        let count = u16::from_be_bytes(bytes[4..6].try_into().unwrap()) as usize;
        let entry = (0..count)
            .map(|index| 12 + index * 16)
            .find(|&offset| &bytes[offset..offset + 4] == b"GSUB")
            .unwrap();
        let offset = bytes.len() as u32;
        bytes[entry + 8..entry + 12].copy_from_slice(&offset.to_be_bytes());
        bytes[entry + 12..entry + 16].copy_from_slice(&(gsub.len() as u32).to_be_bytes());
        bytes.extend(gsub);
        let mut store = FontStore::new();
        let font = store.register(bytes).unwrap();
        let shape = |text| {
            shape_with_direction(
                &store,
                font,
                text,
                16.0,
                &kern_features(false),
                ShapeDirection::Ltr,
            )
            .unwrap()
        };
        assert_eq!(shape("ff").len(), 2);
        assert_eq!(shape("fi").len(), 2);
        assert_eq!(shape("ffi").len(), 1);
        let set = echo_glyph_set(&store, font, 16.0, false, &['f', 'a', 'i']).unwrap();
        assert!(
            set.glyph_ids[0] == 0
                || set.unsafe_pairs.contains(&[0, 0])
                || set.unsafe_pairs.contains(&[0, 2]),
            "ffi must decline even when ff and fi shape independently"
        );
    }
}
