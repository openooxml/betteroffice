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

use std::collections::{BTreeSet, HashMap};

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
    let ligature_pairs = shaping_barriers(store, font, alphabet, &mut glyph_ids, kern)?;
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
    unsafe_pairs.extend(ligature_pairs);
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

fn shaping_barriers(
    store: &FontStore,
    font: FontId,
    alphabet: &[char],
    glyph_ids: &mut [u32],
    kern: bool,
) -> Result<Vec<[u32; 2]>, FontError> {
    use rustybuzz::ttf_parser::{GlyphId, gsub::SubstitutionSubtable};

    // pairs whose left glyph starts a ligature that continues with the right glyph
    let mut ligatures = Vec::new();
    let Some(face) = store.shaping_face(font)? else {
        glyph_ids.fill(0);
        return Ok(ligatures);
    };
    if face.tables().morx.is_some() || face.tables().trak.is_some() {
        glyph_ids.fill(0);
        return Ok(ligatures);
    }
    if face
        .tables()
        .gpos
        .is_some_and(|table| contextual_positioning(&table, alphabet, glyph_ids, kern))
    {
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
        let active = default_echo_feature(feature.tag.to_bytes(), alphabet, kern)
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
            for (index, nominal_glyph) in nominal.iter().enumerate() {
                let coverage = subtable.coverage();
                let covered = nominal_glyph
                    .and_then(|id| coverage.get(GlyphId(id)))
                    .or_else(|| {
                        (glyph_ids[index] != 0)
                            .then(|| coverage.get(GlyphId(glyph_ids[index] as u16)))
                            .flatten()
                    });
                let Some(covered) = covered else {
                    continue;
                };
                match subtable {
                    SubstitutionSubtable::Ligature(table) => {
                        let Some(set) = table.ligature_sets.get(covered) else {
                            glyph_ids[index] = 0;
                            continue;
                        };
                        for ligature in set {
                            let Some(next) = ligature.components.get(0) else {
                                glyph_ids[index] = 0;
                                break;
                            };
                            for (right, right_glyph) in nominal.iter().enumerate() {
                                if *right_glyph == Some(next.0)
                                    || glyph_ids[right] != 0
                                        && glyph_ids[right] == u32::from(next.0)
                                {
                                    ligatures.push([index as u32, right as u32]);
                                }
                            }
                        }
                    }
                    SubstitutionSubtable::Context(_)
                    | SubstitutionSubtable::ChainContext(_)
                    | SubstitutionSubtable::ReverseChainSingle(_) => {
                        if context_may_match(subtable, &nominal) {
                            glyph_ids.fill(0);
                            return Ok(ligatures);
                        }
                    }
                    _ => glyph_ids[index] = 0,
                }
            }
        }
    }
    Ok(ligatures)
}

fn default_echo_feature(tag: [u8; 4], alphabet: &[char], kern: bool) -> bool {
    matches!(
        &tag,
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
    ) || alphabet.contains(&'\u{2044}') && matches!(&tag, b"frac" | b"numr" | b"dnom")
        || kern && tag == *b"kern"
}

fn contextual_positioning(
    table: &rustybuzz::ttf_parser::opentype_layout::LayoutTable<'_>,
    alphabet: &[char],
    glyph_ids: &[u32],
    kern: bool,
) -> bool {
    use rustybuzz::{UnicodeBuffer, script, ttf_parser::gpos::PositioningSubtable};

    if table.variations.is_some() {
        return true;
    }
    let mut scripts = BTreeSet::new();
    for (&ch, &id) in alphabet.iter().zip(glyph_ids) {
        if id == 0 {
            continue;
        }
        let mut buffer = UnicodeBuffer::new();
        buffer.push_str(&ch.to_string());
        buffer.set_direction(rustybuzz::Direction::LeftToRight);
        buffer.guess_segment_properties();
        let tags = match buffer.script() {
            script::LATIN => [*b"latn", *b"DFLT", *b"dflt"],
            script::COMMON | script::INHERITED | script::UNKNOWN => [*b"DFLT", *b"dflt", *b"latn"],
            _ => return true,
        };
        if let Some(selected) = tags.into_iter().find_map(|tag| {
            (0..table.scripts.len()).find_map(|index| {
                table
                    .scripts
                    .get(index)
                    .filter(|script| script.tag.to_bytes() == tag)
            })
        }) {
            scripts.insert(selected.tag.to_bytes());
        }
    }
    let mut lookups = BTreeSet::new();
    for index in 0..table.scripts.len() {
        let Some(script) = table.scripts.get(index) else {
            return true;
        };
        if !scripts.contains(&script.tag.to_bytes()) {
            continue;
        }
        let mut languages = Vec::new();
        if let Some(language) = script.default_language {
            languages.push(language);
        }
        for index in 0..script.languages.len() {
            let Some(language) = script.languages.get(index) else {
                return true;
            };
            languages.push(language);
        }
        for language in languages {
            for index in language
                .feature_indices
                .into_iter()
                .chain(language.required_feature)
            {
                let Some(feature) = table.features.get(index) else {
                    return true;
                };
                if default_echo_feature(feature.tag.to_bytes(), alphabet, kern)
                    || language.required_feature == Some(index)
                {
                    lookups.extend(feature.lookup_indices);
                }
            }
        }
    }
    for index in lookups {
        let Some(lookup) = table.lookups.get(index) else {
            return true;
        };
        for index in 0..lookup.subtables.len() {
            if matches!(
                lookup.subtables.get::<PositioningSubtable<'_>>(index),
                None | Some(PositioningSubtable::Context(_) | PositioningSubtable::ChainContext(_))
            ) {
                return true;
            }
        }
    }
    false
}

fn context_may_match(
    subtable: rustybuzz::ttf_parser::gsub::SubstitutionSubtable<'_>,
    nominal: &[Option<u16>],
) -> bool {
    use rustybuzz::ttf_parser::{
        GlyphId,
        gsub::SubstitutionSubtable,
        opentype_layout::{ChainedContextLookup, ContextLookup, Coverage},
    };
    let matches = |coverage: Coverage<'_>| {
        nominal
            .iter()
            .any(|id| id.is_some_and(|id| coverage.contains(GlyphId(id))))
    };
    match subtable {
        SubstitutionSubtable::Context(ContextLookup::Format3 {
            coverage,
            coverages,
            ..
        }) => matches(coverage) && coverages.into_iter().all(matches),
        SubstitutionSubtable::ChainContext(ChainedContextLookup::Format3 {
            coverage,
            backtrack_coverages,
            input_coverages,
            lookahead_coverages,
            ..
        }) => {
            matches(coverage)
                && backtrack_coverages.into_iter().all(matches)
                && input_coverages.into_iter().all(matches)
                && lookahead_coverages.into_iter().all(matches)
        }
        SubstitutionSubtable::ReverseChainSingle(table) => {
            matches(table.coverage)
                && table.backtrack_coverages.into_iter().all(matches)
                && table.lookahead_coverages.into_iter().all(matches)
        }
        _ => true,
    }
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
        assert_ne!(set.glyph_ids[0], 0);
        assert!(set.unsafe_pairs.contains(&[0, 0]));
        assert!(
            !set.unsafe_pairs.contains(&[0, 1]) && !set.unsafe_pairs.contains(&[0, 2]),
            "a ligature start leaves the pairs no ligature continues with echoable"
        );
    }

    fn contextual_positioning_font(kind: u16, extension: bool, feature: &[u8; 4]) -> Vec<u8> {
        let (store, font) = store();
        let a = store.glyph_id(font, 'A').unwrap().unwrap();
        let context: Vec<u16> = if kind == 7 {
            vec![3, 3, 1, 16, 16, 16, 1, 1, 1, 1, a]
        } else {
            vec![3, 1, 20, 1, 20, 1, 20, 1, 0, 1, 1, 1, a]
        };
        let mut gpos = Vec::new();
        for value in [1u16, 0, 10, 30, 44, 1] {
            gpos.extend(value.to_be_bytes());
        }
        gpos.extend(b"DFLT");
        for value in [8u16, 4, 0, 0, 0xffff, 1, 0, 1] {
            gpos.extend(value.to_be_bytes());
        }
        gpos.extend(feature);
        let second = 6 + 8 + context.len() as u16 * 2 + if extension { 8 } else { 0 };
        for value in [
            8u16,
            0,
            1,
            0,
            2,
            6,
            second,
            if extension { 9 } else { kind },
            0,
            1,
            8,
        ] {
            gpos.extend(value.to_be_bytes());
        }
        if extension {
            gpos.extend(1u16.to_be_bytes());
            gpos.extend(kind.to_be_bytes());
            gpos.extend(8u32.to_be_bytes());
        }
        for value in context {
            gpos.extend(value.to_be_bytes());
        }
        for value in [1u16, 0, 1, 8, 1, 8, 4, (-128i16) as u16, 1, 1, a] {
            gpos.extend(value.to_be_bytes());
        }
        let mut bytes = LIBERATION_SANS.to_vec();
        let count = u16::from_be_bytes(bytes[4..6].try_into().unwrap()) as usize;
        let entry = (0..count)
            .map(|index| 12 + index * 16)
            .find(|&offset| &bytes[offset..offset + 4] == b"GPOS")
            .unwrap();
        let offset = bytes.len() as u32;
        bytes[entry + 8..entry + 12].copy_from_slice(&offset.to_be_bytes());
        bytes[entry + 12..entry + 16].copy_from_slice(&(gpos.len() as u32).to_be_bytes());
        bytes.extend(gpos);
        bytes
    }

    fn assert_contextual_positioning_declines(kind: u16, extension: bool, feature: &[u8; 4]) {
        let mut store = FontStore::new();
        let font = store
            .register(contextual_positioning_font(kind, extension, feature))
            .unwrap();
        let alphabet = ['A', 'B'];
        let size = 16.0;
        let shape = |text: &str| {
            shape_with_direction(
                &store,
                font,
                text,
                size,
                &kern_features(true),
                ShapeDirection::Ltr,
            )
            .unwrap()
        };
        let isolated = shape("A")[0].x_advance;
        for text in ["AA", "AABBA"] {
            assert_eq!(
                shape(text)
                    .iter()
                    .map(|glyph| glyph.x_advance)
                    .collect::<Vec<_>>(),
                text.chars()
                    .map(|ch| shape(&ch.to_string())[0].x_advance)
                    .collect::<Vec<_>>()
            );
        }
        let triple = shape("AAA");
        assert_eq!(triple.len(), 3);
        assert_eq!(triple[0].x_advance, isolated);
        assert_ne!(triple[1].x_advance, isolated);
        assert_eq!(triple[2].x_advance, isolated);
        let set = echo_glyph_set(&store, font, size, true, &alphabet).unwrap();
        assert!(
            set.glyph_ids.iter().all(|&id| id == 0),
            "contextual positioning must decline the font"
        );
    }

    #[test]
    fn contextual_kern_positioning_declines_echo() {
        assert_contextual_positioning_declines(7, false, b"kern");
    }

    #[test]
    fn chained_contextual_kern_positioning_declines_echo() {
        assert_contextual_positioning_declines(8, false, b"kern");
    }

    #[test]
    fn extension_contextual_positioning_declines_echo() {
        for kind in [7, 8] {
            assert_contextual_positioning_declines(kind, true, b"kern");
        }
    }

    #[test]
    fn contextual_mark_positioning_declines_echo_without_kerning() {
        assert_contextual_positioning_declines(7, false, b"mark");
        let mut store = FontStore::new();
        let font = store
            .register(contextual_positioning_font(7, false, b"mark"))
            .unwrap();
        let set = echo_glyph_set(&store, font, 16.0, false, &['A', 'B']).unwrap();
        assert!(set.glyph_ids.iter().all(|&id| id == 0));
    }

    #[test]
    fn disabled_contextual_kern_positioning_keeps_echo() {
        for kind in [7, 8] {
            let mut store = FontStore::new();
            let font = store
                .register(contextual_positioning_font(kind, true, b"kern"))
                .unwrap();
            let set = echo_glyph_set(&store, font, 16.0, false, &['A', 'B']).unwrap();
            assert!(set.glyph_ids.iter().all(|&id| id != 0));
            let shaped: Vec<f32> = shape_with_direction(
                &store,
                font,
                "AAA",
                16.0,
                &kern_features(false),
                ShapeDirection::Ltr,
            )
            .unwrap()
            .iter()
            .map(|glyph| glyph.x_advance)
            .collect();
            assert_eq!(advance(&set, &['A', 'B'], "AAA"), shaped);
        }
    }

    fn rewrite_gpos(mut bytes: Vec<u8>, rewrite: impl FnOnce(&[u8]) -> Vec<u8>) -> Vec<u8> {
        let count = u16::from_be_bytes(bytes[4..6].try_into().unwrap()) as usize;
        let entry = (0..count)
            .map(|index| 12 + index * 16)
            .find(|&offset| &bytes[offset..offset + 4] == b"GPOS")
            .unwrap();
        let start = u32::from_be_bytes(bytes[entry + 8..entry + 12].try_into().unwrap()) as usize;
        let len = u32::from_be_bytes(bytes[entry + 12..entry + 16].try_into().unwrap()) as usize;
        let gpos = rewrite(&bytes[start..start + len]);
        let offset = bytes.len() as u32;
        bytes[entry + 8..entry + 12].copy_from_slice(&offset.to_be_bytes());
        bytes[entry + 12..entry + 16].copy_from_slice(&(gpos.len() as u32).to_be_bytes());
        bytes.extend(gpos);
        bytes
    }

    fn assert_latin_positioning_keeps_echo(bytes: Vec<u8>) {
        let mut store = FontStore::new();
        let font = store.register(bytes).unwrap();
        let alphabet = ['A', 'B'];
        let set = echo_glyph_set(&store, font, 16.0, true, &alphabet).unwrap();
        assert!(set.glyph_ids.iter().all(|&id| id != 0));
        for text in ["A", "AA", "AAA", "AABBA"] {
            let shaped: Vec<f32> = shape_with_direction(
                &store,
                font,
                text,
                16.0,
                &kern_features(true),
                ShapeDirection::Ltr,
            )
            .unwrap()
            .iter()
            .map(|glyph| glyph.x_advance)
            .collect();
            assert_eq!(advance(&set, &alphabet, text), shaped);
        }
    }

    #[test]
    fn contextual_positioning_in_another_script_keeps_latin_echo() {
        for kind in [7, 8] {
            let bytes = rewrite_gpos(contextual_positioning_font(kind, true, b"mark"), |table| {
                let mut table = table.to_vec();
                table[12..16].copy_from_slice(b"hebr");
                table
            });
            assert_latin_positioning_keeps_echo(bytes);
        }
    }

    #[test]
    fn disabled_optional_positioning_features_keep_echo() {
        for feature in [b"ss01", b"salt", b"dlig", b"vkrn"] {
            assert_latin_positioning_keeps_echo(contextual_positioning_font(7, true, feature));
        }
    }

    #[test]
    fn latin_script_precedes_default_contextual_positioning() {
        let bytes = rewrite_gpos(contextual_positioning_font(8, true, b"kern"), |table| {
            let mut gpos = table[..10].to_vec();
            gpos[6..8].copy_from_slice(&46u16.to_be_bytes());
            gpos[8..10].copy_from_slice(&60u16.to_be_bytes());
            gpos.extend(2u16.to_be_bytes());
            gpos.extend(b"DFLT");
            gpos.extend(14u16.to_be_bytes());
            gpos.extend(b"latn");
            gpos.extend(26u16.to_be_bytes());
            gpos.extend(&table[18..30]);
            for value in [4u16, 0, 0, 0xffff, 0] {
                gpos.extend(value.to_be_bytes());
            }
            gpos.extend(&table[30..]);
            gpos
        });
        assert_latin_positioning_keeps_echo(bytes);
    }
}
