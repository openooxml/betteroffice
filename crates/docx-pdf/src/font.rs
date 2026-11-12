//! Font planning and subsetting for PDF export.

use std::collections::{BTreeMap, BTreeSet, HashMap};

use docx_layout::display_list::{GlyphRunPrimitive, Primitive};
use ooxml_text::{FontId, FontStore};
use pdf_writer::types::{FontFlags, SystemInfo, UnicodeCmap};
use pdf_writer::{Name, Str};
use subsetter::GlyphRemapper;
use ttf_parser::{Face, GlyphId};

use crate::PdfError;
use crate::options::PdfConformance;

/// PDF/A-2u rejects any text code that addresses `.notdef` (glyph/CID 0).
pub(crate) const PDFA_MULTI_SCALAR_CLUSTER: &str =
    "PDF/A-2u cannot fall back for a multi-scalar glyph cluster";
pub(crate) const PDFA_NOTDEF: &str = "PDF/A-2u text cannot address .notdef";
/// No registered, usable face maps the painted scalar to a real glyph.
pub(crate) const PDFA_MISSING_GLYPH: &str =
    "PDF/A-2u text needs a glyph the registered fonts do not cover";
/// PDF/A-2u cannot embed a bitmap-only face as a valid CID font program.
pub(crate) const PDFA_NO_OUTLINES: &str = "PDF/A-2u text font has no embeddable outlines";
/// The collection and the painter disagree about the embedded font set.
pub(crate) const PDFA_EMBEDDED_FONT: &str = "PDF/A-2u text needs an embedded font";
/// A glyph is painted from a cluster with no Unicode source text.
pub(crate) const PDFA_NO_SOURCE_TEXT: &str = "PDF/A-2u glyph has no Unicode source text";
/// A painted code has no `ToUnicode` entry to extract it back to Unicode.
pub(crate) const PDFA_UNMAPPED_CODE: &str = "PDF/A-2u painted code has no ToUnicode mapping";
/// One CFF glyph would carry several distinct Unicode source strings.
pub(crate) const PDFA_CFF_AMBIGUOUS: &str =
    "PDF/A-2u cannot map one CFF glyph to several source strings";
/// A source string has no character a `ToUnicode` map may legally carry.
pub(crate) const PDFA_UNMAPPABLE_TEXT: &str = "PDF/A-2u source text has no valid Unicode mapping";

/// The part of a source string a PDF/A `ToUnicode` destination may carry.
///
/// U+0000, U+FEFF and U+FFFE carry no extractable text for PDF/A and must not
/// appear in a destination string. An empty result means the code cannot be
/// mapped conformantly; callers reject it instead of inventing U+FFFD.
fn pdfa_unicode_text(text: &str) -> String {
    text.chars()
        .filter(|ch| !matches!(ch, '\u{0000}' | '\u{FEFF}' | '\u{FFFE}'))
        .collect()
}

/// A font could not be embedded, separating policy rejections from load and
/// subset failures so [`EmbeddedFont::new`] keeps its historical error type.
#[derive(Debug)]
enum EmbedFailure {
    Font(FontError),
    Policy(&'static str),
}

impl From<FontError> for EmbedFailure {
    fn from(error: FontError) -> Self {
        Self::Font(error)
    }
}

#[derive(Debug)]
pub enum FontError {
    UnknownFont(u32),
    Font(String),
    Subset(String),
}

impl std::fmt::Display for FontError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::UnknownFont(id) => write!(f, "unknown font id {id}"),
            Self::Font(error) => write!(f, "font error: {error}"),
            Self::Subset(error) => write!(f, "font subset error: {error}"),
        }
    }
}

impl std::error::Error for FontError {}

/// One embedded font, already subset and ready to write.
pub struct EmbeddedFont {
    /// PDF resource name, e.g. `F1`.
    pub resource: String,
    pub base_name: String,
    pub subset: Vec<u8>,
    pub remap: BTreeMap<u16, u16>,
    /// CID for one `(glyph id, source text)` occurrence. A glyph shared by
    /// several source characters (Arabic dots, `.notdef` emoji) gets one CID
    /// per text so ToUnicode stays unambiguous.
    pub cid_map: BTreeMap<(u16, String), u16>,
    /// CID-to-subset-glyph map when duplicate CIDs exist.
    pub cid_to_gid: Option<Vec<u16>>,
    pub widths: Vec<f32>,
    pub tounicode: Vec<u8>,
    pub bbox: [f32; 4],
    pub italic_angle: f32,
    pub ascent: f32,
    pub descent: f32,
    pub cap_height: f32,
    pub flags: u32,
    pub is_cff: bool,
}

fn subset_tag(mut id: u32) -> String {
    let mut tag = String::with_capacity(6);
    for _ in 0..6 {
        tag.push(char::from(b'A' + (id % 26) as u8));
        id /= 26;
    }
    tag
}

pub(crate) fn face_has_outlines(face: &Face) -> bool {
    face.tables().glyf.is_some() || face.tables().cff.is_some() || face.tables().cff2.is_some()
}

fn has_outlines(face: &Face) -> bool {
    face_has_outlines(face)
}

fn push_u16(out: &mut Vec<u8>, value: u16) {
    out.extend_from_slice(&value.to_be_bytes());
}

fn push_i16(out: &mut Vec<u8>, value: i16) {
    out.extend_from_slice(&value.to_be_bytes());
}

fn push_u32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_be_bytes());
}

fn push_i64(out: &mut Vec<u8>, value: i64) {
    out.extend_from_slice(&value.to_be_bytes());
}

fn table_checksum(data: &[u8]) -> u32 {
    let mut sum = 0u32;
    for chunk in data.chunks(4) {
        let mut word = [0u8; 4];
        word[..chunk.len()].copy_from_slice(chunk);
        sum = sum.wrapping_add(u32::from_be_bytes(word));
    }
    sum
}

/// Build a minimal valid TrueType program with empty glyphs for `uses`.
///
/// Bitmap-only faces (CBDT/CBLC, sbix) have no `glyf`/CFF outlines, so neither
/// `subsetter` nor a PDF CID font can carry their pixels. A valid empty-glyph
/// program keeps CID addressing and ToUnicode extraction working instead of
/// failing the whole export. `uses` keys are remapped to compact glyph ids.
fn synth_outline_less_font(
    face: &Face,
    uses: &BTreeMap<u16, BTreeSet<String>>,
) -> Result<(Vec<u8>, BTreeMap<u16, u16>, Vec<f32>), FontError> {
    let mut remap = BTreeMap::new();
    let mut next_gid = 1u16;
    if uses.contains_key(&0) {
        remap.insert(0, 0);
    }
    for old in uses.keys().copied() {
        if old == 0 {
            continue;
        }
        remap.insert(old, next_gid);
        next_gid = next_gid
            .checked_add(1)
            .ok_or_else(|| FontError::Subset("too many glyphs to synthesize".into()))?;
    }
    let num_glyphs = next_gid;
    let upem = face.units_per_em().max(1);
    let notdef_advance = face.glyph_hor_advance(GlyphId(0)).unwrap_or(upem);
    let mut advances = vec![notdef_advance; usize::from(num_glyphs)];
    for (&old, &new) in &remap {
        if let Some(advance) = face.glyph_hor_advance(GlyphId(old)) {
            advances[usize::from(new)] = advance;
        }
    }
    let widths: Vec<f32> = advances
        .iter()
        .map(|&advance| f32::from(advance) * 1000.0 / f32::from(upem))
        .collect();

    // The program only needs metrics and the glyph slots; every outline is
    // empty, so `glyf` is zero-length and every `loca` entry is zero.
    let glyf: Vec<u8> = Vec::new();
    let loca = vec![0u8; (usize::from(num_glyphs) + 1) * 2];

    let bbox = face.global_bounding_box();
    let mut head = Vec::with_capacity(54);
    push_u32(&mut head, 0x0001_0000);
    push_u32(&mut head, 0x0001_0000);
    push_u32(&mut head, 0);
    push_u32(&mut head, 0x5F0F_3CF5);
    push_u16(&mut head, 0x000B);
    push_u16(&mut head, upem);
    push_i64(&mut head, 0);
    push_i64(&mut head, 0);
    push_i16(&mut head, bbox.x_min);
    push_i16(&mut head, bbox.y_min);
    push_i16(&mut head, bbox.x_max);
    push_i16(&mut head, bbox.y_max);
    let mac_style = u16::from(face.is_bold()) | (u16::from(face.is_italic()) << 1);
    push_u16(&mut head, mac_style);
    push_u16(&mut head, 8);
    push_i16(&mut head, 2);
    push_i16(&mut head, 0);
    push_i16(&mut head, 0);

    let advance_max = advances.iter().copied().max().unwrap_or(upem);
    let mut hhea = Vec::with_capacity(36);
    push_u32(&mut hhea, 0x0001_0000);
    push_i16(&mut hhea, face.ascender());
    push_i16(&mut hhea, face.descender());
    push_i16(&mut hhea, face.line_gap());
    push_u16(&mut hhea, advance_max);
    push_i16(&mut hhea, 0);
    push_i16(&mut hhea, 0);
    push_i16(&mut hhea, 0);
    push_i16(&mut hhea, 1);
    push_i16(&mut hhea, 0);
    push_i16(&mut hhea, 0);
    for _ in 0..4 {
        push_i16(&mut hhea, 0);
    }
    push_i16(&mut hhea, 0);
    push_u16(&mut hhea, num_glyphs);

    let mut hmtx = Vec::with_capacity(usize::from(num_glyphs) * 4);
    for &advance in &advances {
        push_u16(&mut hmtx, advance);
        push_i16(&mut hmtx, 0);
    }

    let mut maxp = Vec::with_capacity(32);
    push_u32(&mut maxp, 0x0001_0000);
    push_u16(&mut maxp, num_glyphs);
    for _ in 0..13 {
        push_u16(&mut maxp, 0);
    }

    let mut cmap = Vec::with_capacity(36);
    push_u16(&mut cmap, 0);
    push_u16(&mut cmap, 1);
    push_u16(&mut cmap, 3);
    push_u16(&mut cmap, 1);
    push_u32(&mut cmap, 12);
    push_u16(&mut cmap, 4);
    push_u16(&mut cmap, 24);
    push_u16(&mut cmap, 0);
    push_u16(&mut cmap, 2);
    push_u16(&mut cmap, 2);
    push_u16(&mut cmap, 0);
    push_u16(&mut cmap, 0);
    push_u16(&mut cmap, 0xFFFF);
    push_u16(&mut cmap, 0);
    push_u16(&mut cmap, 0xFFFF);
    push_u16(&mut cmap, 1);
    push_u16(&mut cmap, 0);

    let mut name = Vec::with_capacity(6);
    push_u16(&mut name, 0);
    push_u16(&mut name, 0);
    push_u16(&mut name, 6);

    let mut post = Vec::with_capacity(32);
    push_u32(&mut post, 0x0003_0000);
    push_u32(&mut post, (face.italic_angle() * 65_536.0) as i32 as u32);
    push_i16(&mut post, -100);
    push_i16(&mut post, 50);
    for _ in 0..5 {
        push_u32(&mut post, 0);
    }

    let mut tables: Vec<(&[u8; 4], Vec<u8>)> = vec![
        (b"cmap", cmap),
        (b"glyf", glyf),
        (b"head", head),
        (b"hhea", hhea),
        (b"hmtx", hmtx),
        (b"loca", loca),
        (b"maxp", maxp),
        (b"name", name),
        (b"post", post),
    ];
    tables.sort_by(|left, right| left.0.cmp(right.0));

    let count = tables.len() as u16;
    let mut entry_selector = 0u16;
    while (1u16 << (entry_selector + 1)) <= count {
        entry_selector += 1;
    }
    let search_range = 16u16 << entry_selector;
    let range_shift = count * 16 - search_range;

    let mut offset = 12 + 16 * usize::from(count);
    let mut records = Vec::with_capacity(tables.len());
    for (tag, data) in &tables {
        offset = (offset + 3) & !3;
        records.push((*tag, table_checksum(data), offset as u32, data.len() as u32));
        offset += data.len();
    }

    let mut font = Vec::with_capacity(offset);
    push_u32(&mut font, 0x0001_0000);
    push_u16(&mut font, count);
    push_u16(&mut font, search_range);
    push_u16(&mut font, entry_selector);
    push_u16(&mut font, range_shift);
    for (tag, checksum, table_offset, length) in &records {
        font.extend_from_slice(*tag);
        push_u32(&mut font, *checksum);
        push_u32(&mut font, *table_offset);
        push_u32(&mut font, *length);
    }
    for (index, (_, data)) in tables.iter().enumerate() {
        let target = records[index].2 as usize;
        while font.len() < target {
            font.push(0);
        }
        font.extend_from_slice(data);
    }

    let adjustment = 0xB1B0_AFBAu32.wrapping_sub(table_checksum(&font));
    if let Some(head_offset) = records
        .iter()
        .find(|record| record.0 == b"head")
        .map(|record| record.2 as usize)
    {
        font[head_offset + 8..head_offset + 12].copy_from_slice(&adjustment.to_be_bytes());
    }

    Face::parse(&font, 0).map_err(|error| FontError::Font(format!("synthesized font: {error}")))?;
    Ok((font, remap, widths))
}

/// Reject uses that can never produce conforming PDF/A-2u text.
///
/// Runs in `BTreeMap`/`BTreeSet` order, so the first rejection is deterministic.
fn validate_pdfa_uses(
    uses: &BTreeMap<u16, BTreeSet<String>>,
    is_cff: bool,
) -> Result<(), EmbedFailure> {
    for (old, texts) in uses {
        if *old == 0 {
            return Err(EmbedFailure::Policy(PDFA_NOTDEF));
        }
        if texts.is_empty() {
            return Err(EmbedFailure::Policy(PDFA_NO_SOURCE_TEXT));
        }
        // A CFF CID *is* its charstring index, and `/CIDToGIDMap` is only
        // legal for CIDFontType2, so one CFF glyph cannot back several codes.
        if is_cff && texts.len() > 1 {
            return Err(EmbedFailure::Policy(PDFA_CFF_AMBIGUOUS));
        }
        for text in texts {
            if text.is_empty() {
                return Err(EmbedFailure::Policy(PDFA_NO_SOURCE_TEXT));
            }
            if pdfa_unicode_text(text).is_empty() {
                return Err(EmbedFailure::Policy(PDFA_UNMAPPABLE_TEXT));
            }
        }
    }
    Ok(())
}

impl EmbeddedFont {
    #[allow(dead_code)] // Standard-path constructor used by the text fallback tests.
    pub fn new(
        resource: String,
        store: &FontStore,
        id: u32,
        uses: &BTreeMap<u16, BTreeSet<String>>,
    ) -> Result<Self, FontError> {
        Self::build(resource, store, id, uses, PdfConformance::Standard).map_err(|failure| {
            match failure {
                EmbedFailure::Font(error) => error,
                // Standard applies no PDF/A policy; keep the conversion total.
                EmbedFailure::Policy(message) => FontError::Subset(message.to_owned()),
            }
        })
    }

    /// Embed a subset under `conformance`.
    ///
    /// [`PdfConformance::Standard`] is [`EmbeddedFont::new`] with the error
    /// mapped onto [`PdfError`]. [`PdfConformance::PdfA2u`] also rejects
    /// `.notdef`, glyphs without Unicode source text, ambiguous CFF glyphs and
    /// unmappable source text, and never builds a `ToUnicode` destination from
    /// U+0000, U+FEFF or U+FFFE.
    pub(crate) fn new_with_conformance(
        resource: String,
        store: &FontStore,
        id: u32,
        uses: &BTreeMap<u16, BTreeSet<String>>,
        conformance: PdfConformance,
    ) -> Result<Self, PdfError> {
        Self::build(resource, store, id, uses, conformance).map_err(|failure| match failure {
            EmbedFailure::Font(error) => PdfError::Font(error),
            EmbedFailure::Policy(message) => PdfError::Unsupported(message),
        })
    }

    fn build(
        resource: String,
        store: &FontStore,
        id: u32,
        uses: &BTreeMap<u16, BTreeSet<String>>,
        conformance: PdfConformance,
    ) -> Result<Self, EmbedFailure> {
        let font_id = FontId::from_u32(id);
        let bytes = store
            .font_bytes(font_id)
            .map_err(|error| FontError::Font(error.to_string()))?;
        let original = Face::parse(bytes, 0).map_err(|error| FontError::Font(error.to_string()))?;
        let upem = f32::from(original.units_per_em().max(1));
        let scale = 1000.0 / upem;
        let is_cff = original.tables().cff.is_some();

        if uses.len() > usize::from(u16::MAX) {
            return Err(FontError::Subset("too many glyphs".into()).into());
        }
        if conformance.is_pdfa() {
            validate_pdfa_uses(uses, is_cff)?;
            if !has_outlines(&original) {
                return Err(EmbedFailure::Policy(PDFA_NO_OUTLINES));
            }
        }
        let mut remapper = GlyphRemapper::new();
        for old in uses.keys().copied() {
            remapper.remap(old);
        }
        let subset_result = if original.tables().cff2.is_some() {
            // CFF2 has no direct PDF embedding path; the subsetter converts it
            // to a TrueType program so its outlines survive as FontFile2.
            subsetter::subset_with_variations(bytes, 0, &[], &remapper)
        } else {
            subsetter::subset(bytes, 0, &remapper)
        };
        let (subset, remap, widths) = match subset_result {
            Ok(subset) => {
                let face =
                    Face::parse(&subset, 0).map_err(|error| FontError::Font(error.to_string()))?;
                let mut remap = BTreeMap::new();
                let mut widths = vec![1000.0; face.number_of_glyphs() as usize];
                for (old, new) in uses
                    .keys()
                    .copied()
                    .filter_map(|old| remapper.get(old).map(|new| (old, new)))
                {
                    remap.insert(old, new);
                    if let Some(advance) = face.glyph_hor_advance(GlyphId(new))
                        && let Some(width) = widths.get_mut(usize::from(new))
                    {
                        *width = f32::from(advance) * scale;
                    }
                }
                (subset, remap, widths)
            }
            Err(_) if !has_outlines(&original) => {
                // Bitmap-only faces (CBDT/CBLC, sbix) cannot be subset by
                // `subsetter`, and PDF has no CID font path for their bitmaps.
                // Emit a valid outline-less TrueType program instead: text
                // extraction stays correct and export never fails.
                synth_outline_less_font(&original, uses)?
            }
            Err(error) => return Err(EmbedFailure::Font(FontError::Subset(error.to_string()))),
        };

        let mut widths = widths;
        let max_glyph = remap.values().copied().max().unwrap_or(0);
        let mut next_cid = max_glyph
            .checked_add(1)
            .ok_or_else(|| FontError::Subset("too many glyphs".into()))?;
        let mut cid_map = BTreeMap::new();
        // Identity for every subset glyph: non-base dot components still
        // paint, so their CID must resolve to their own glyph. Duplicate CIDs
        // are appended after this range and redirected by CIDToGIDMap.
        let mut cid_to_gid: Vec<u16> = (0..=max_glyph).collect();
        for (old, texts) in uses {
            let Some(glyph) = remap.get(old).copied() else {
                if conformance.is_pdfa() {
                    return Err(EmbedFailure::Font(FontError::Subset(
                        "subset lost a requested glyph".into(),
                    )));
                }
                continue;
            };
            if conformance.is_pdfa() && glyph == 0 {
                return Err(EmbedFailure::Policy(PDFA_NOTDEF));
            }
            for (index, text) in texts.iter().enumerate() {
                if text.is_empty() || (is_cff && index > 0) {
                    continue;
                }
                let cid = if index == 0 || is_cff {
                    glyph
                } else {
                    let cid = next_cid;
                    next_cid = next_cid
                        .checked_add(1)
                        .ok_or_else(|| FontError::Subset("too many CIDs".into()))?;
                    cid
                };
                // CID 0 is `.notdef`: a PDF/A-2u code must never address it.
                if conformance.is_pdfa() && cid == 0 {
                    return Err(EmbedFailure::Policy(PDFA_NOTDEF));
                }
                cid_map.insert((*old, text.clone()), cid);
                if usize::from(cid) >= widths.len() {
                    widths.resize(usize::from(cid) + 1, widths[usize::from(glyph)]);
                } else {
                    widths[usize::from(cid)] = widths[usize::from(glyph)];
                }
                if usize::from(cid) >= cid_to_gid.len() {
                    cid_to_gid.resize(usize::from(cid) + 1, 0);
                }
                cid_to_gid[usize::from(cid)] = glyph;
            }
        }
        let cid_to_gid = cid_to_gid
            .iter()
            .enumerate()
            .any(|(cid, gid)| u16::try_from(cid).ok() != Some(*gid))
            .then_some(cid_to_gid);

        let mut cmap = UnicodeCmap::<u16>::new(
            Name(b"Custom"),
            SystemInfo {
                registry: Str(b"Adobe"),
                ordering: Str(b"Identity"),
                supplement: 0,
            },
        );
        if conformance.is_pdfa() {
            // One destination per painted code. Pre-validation already
            // rejected empty or unmappable sources; re-check here because the
            // mapping must be total for every collected code.
            let mut destinations: BTreeMap<u16, String> = BTreeMap::new();
            for ((_, text), cid) in &cid_map {
                let destination = pdfa_unicode_text(text);
                if destination.is_empty() {
                    return Err(EmbedFailure::Policy(PDFA_UNMAPPABLE_TEXT));
                }
                match destinations.get(cid) {
                    Some(existing) if existing != &destination => {
                        return Err(EmbedFailure::Policy(PDFA_CFF_AMBIGUOUS));
                    }
                    _ => {
                        destinations.insert(*cid, destination);
                    }
                }
            }
            for (cid, destination) in &destinations {
                cmap.pair_with_multiple(*cid, destination.chars());
            }
        } else {
            for ((_, text), cid) in &cid_map {
                cmap.pair_with_multiple(*cid, text.chars());
            }
        }
        let tounicode = cmap.finish().into_vec();

        let bbox = original.global_bounding_box();
        let bbox = [
            f32::from(bbox.x_min) * scale,
            f32::from(bbox.y_min) * scale,
            f32::from(bbox.x_max) * scale,
            f32::from(bbox.y_max) * scale,
        ];

        let mut flags = FontFlags::NON_SYMBOLIC;
        if original.is_italic() {
            flags |= FontFlags::ITALIC;
        }
        if original.is_bold() {
            flags |= FontFlags::FORCE_BOLD;
        }
        if original.is_monospaced() {
            flags |= FontFlags::FIXED_PITCH;
        }

        Ok(Self {
            resource,
            base_name: format!("{}+BO{id}", subset_tag(id)),
            remap,
            cid_map,
            cid_to_gid,
            widths,
            tounicode,
            bbox,
            italic_angle: original.italic_angle(),
            ascent: f32::from(original.ascender()) * scale,
            descent: f32::from(original.descender()) * scale,
            cap_height: original
                .capital_height()
                .map(|value| f32::from(value) * scale)
                .unwrap_or(700.0),
            flags: flags.bits(),
            is_cff,
            subset,
        })
    }

    pub fn glyph(&self, old: u32) -> Option<u16> {
        self.remap.get(&(old as u16)).copied()
    }

    /// CID for a glyph painted as `text`; falls back to the first CID of the
    /// glyph when no duplicate was allocated (CFF, or text absent).
    pub fn cid(&self, old: u32, text: &str) -> Option<u16> {
        let old = old as u16;
        self.cid_map
            .get(&(old, text.to_owned()))
            .copied()
            .or_else(|| self.remap.get(&old).copied())
    }

    /// Exact CID for the collected `(glyph id, source text)` pair.
    ///
    /// Unlike [`EmbeddedFont::cid`], this never falls back to the glyph's
    /// first CID: PDF/A-2u needs the code whose `ToUnicode` entry is the
    /// painted source string. Returns `None` for `.notdef` and for pairs the
    /// subset does not address.
    pub(crate) fn cid_exact(&self, old: u32, text: &str) -> Option<u16> {
        let old = u16::try_from(old).ok()?;
        let cid = self.cid_map.get(&(old, text.to_owned())).copied()?;
        if cid == 0 {
            return None;
        }
        if let Some(map) = &self.cid_to_gid
            && map.get(usize::from(cid)).copied().unwrap_or(0) == 0
        {
            return None;
        }
        Some(cid)
    }
}

/// Visit every primitive a page paints, in paint order, stopping at the first
/// policy error.
fn try_each_primitive(
    display_list: &docx_layout::display_list::DisplayList,
    mut visit: impl FnMut(&Primitive) -> Result<(), PdfError>,
) -> Result<(), PdfError> {
    for page in &display_list.pages {
        for primitive in page.primitives.iter().chain(
            page.header
                .iter()
                .flat_map(|region| region.primitives.iter())
                .chain(
                    page.footer
                        .iter()
                        .flat_map(|region| region.primitives.iter()),
                )
                .chain(
                    page.note_areas
                        .iter()
                        .flat_map(|region| region.primitives.iter())
                        .chain(
                            page.note_areas
                                .iter()
                                .flat_map(|region| region.separator_primitives.iter()),
                        ),
                ),
        ) {
            visit(primitive)?;
        }
    }
    Ok(())
}

/// Visit every primitive a page paints, in paint order.
fn each_primitive(
    display_list: &docx_layout::display_list::DisplayList,
    mut visit: impl FnMut(&Primitive),
) {
    try_each_primitive(display_list, |primitive| {
        visit(primitive);
        Ok(())
    })
    .expect("an infallible primitive visitor cannot fail");
}

/// Collect every glyph a document paints, keyed by font id and original glyph id.
pub fn collect(
    display_list: &docx_layout::display_list::DisplayList,
) -> BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> {
    let mut fonts: BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> = BTreeMap::new();
    each_primitive(display_list, |primitive| {
        if let Primitive::GlyphRun(run) = primitive {
            collect_run(&mut fonts, run);
        }
    });
    fonts
}

/// Collect every glyph the document paints, including fallback text.
///
/// `GlyphRunPrimitive`s already name their glyphs. `TextRunPrimitive`s carry
/// Unicode text instead, so each scalar is resolved against `store` in
/// registration order; the subset therefore contains every glyph the painter
/// can emit, including `.notdef` for uncovered scalars.
pub fn collect_with_text(
    display_list: &docx_layout::display_list::DisplayList,
    store: &FontStore,
) -> BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> {
    let mut fonts: BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> = BTreeMap::new();
    each_primitive(display_list, |primitive| match primitive {
        Primitive::Text(text) if !text.hidden && !text.text.is_empty() => {
            crate::text::collect_text_glyphs(store, text, &mut fonts);
        }
        Primitive::GlyphRun(run) if !run.hidden => {
            collect_run(&mut fonts, run);
            crate::text::collect_glyph_run_fallbacks(store, run, &mut fonts);
        }
        _ => {}
    });
    fonts
}

/// Collect every glyph under `conformance`, resolving fallback text exactly
/// like [`collect_with_text`].
///
/// [`PdfConformance::Standard`] is [`collect_with_text`] unchanged.
/// [`PdfConformance::PdfA2u`] never records glyph 0: a run whose scalars no
/// registered face covers fails with [`PdfError::Unsupported`] instead of
/// embedding `.notdef`, and hidden content stays uncollected.
pub(crate) fn collect_with_text_policy(
    display_list: &docx_layout::display_list::DisplayList,
    store: &FontStore,
    conformance: PdfConformance,
) -> Result<BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>>, PdfError> {
    if !conformance.is_pdfa() {
        return Ok(collect_with_text(display_list, store));
    }
    let mut fonts: BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> = BTreeMap::new();
    try_each_primitive(display_list, |primitive| match primitive {
        Primitive::Text(text) if !text.hidden && !text.text.is_empty() => {
            crate::text::collect_text_glyphs_with_chains_policy(
                store,
                text,
                None,
                conformance,
                &mut fonts,
            )
        }
        Primitive::GlyphRun(run) if !run.hidden => {
            collect_run_policy(&mut fonts, run)?;
            crate::text::collect_glyph_run_fallbacks_policy(store, run, conformance, &mut fonts)
        }
        _ => Ok(()),
    })?;
    Ok(fonts)
}

/// Collect every glyph for a chain-resolved document under `conformance`.
///
/// [`PdfConformance::Standard`] reproduces the writer's historical chain
/// collector exactly. [`PdfConformance::PdfA2u`] collects only nonzero glyphs,
/// so a run font whose every cluster is re-resolved to a fallback font is not
/// embedded at all; the painter must likewise never reference it.
pub(crate) fn collect_with_chains_policy(
    display_list: &docx_layout::display_list::DisplayList,
    store: &FontStore,
    chains: &HashMap<String, Vec<FontId>>,
    conformance: PdfConformance,
) -> Result<BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>>, PdfError> {
    if !conformance.is_pdfa() {
        return Ok(collect_with_chains_standard(display_list, store, chains));
    }
    let mut fonts: BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> = BTreeMap::new();
    try_each_primitive(display_list, |primitive| match primitive {
        Primitive::Text(text) if !text.hidden && !text.text.is_empty() => {
            crate::text::collect_text_glyphs_with_chains_policy(
                store,
                text,
                Some(chains),
                conformance,
                &mut fonts,
            )
        }
        Primitive::GlyphRun(run) if !run.hidden => {
            collect_run_policy(&mut fonts, run)?;
            crate::text::collect_glyph_run_fallbacks_policy(store, run, conformance, &mut fonts)
        }
        _ => Ok(()),
    })?;
    Ok(fonts)
}

/// The historical chain collection, kept byte-identical for
/// [`PdfConformance::Standard`].
fn collect_with_chains_standard(
    display_list: &docx_layout::display_list::DisplayList,
    store: &FontStore,
    chains: &HashMap<String, Vec<FontId>>,
) -> BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>> {
    let mut fonts = collect(display_list);
    let mut visible_glyph_fonts = BTreeSet::new();
    let mut resolved_fonts = BTreeMap::new();
    each_primitive(display_list, |primitive| match primitive {
        Primitive::Text(text) if !text.hidden && !text.text.is_empty() => {
            crate::text::collect_text_glyphs_with_chains(
                store,
                text,
                Some(chains),
                &mut resolved_fonts,
            );
        }
        Primitive::GlyphRun(run) if !run.hidden => {
            visible_glyph_fonts.insert(run.font_id);
            crate::text::collect_glyph_run_fallbacks(store, run, &mut resolved_fonts);
        }
        _ => {}
    });
    fonts.retain(|font_id, _| visible_glyph_fonts.contains(font_id));
    for (font_id, glyphs) in resolved_fonts {
        let collected = fonts.entry(font_id).or_default();
        for (glyph, sources) in glyphs {
            collected.entry(glyph).or_default().extend(sources);
        }
    }
    fonts
}

/// Sorted, deduplicated source byte offsets of a run's clusters.
pub(crate) fn cluster_starts(run: &GlyphRunPrimitive) -> Vec<u32> {
    // Glyph order is visual for RTL runs, so cluster values decrease. Build
    // the source boundaries from sorted clusters, not from the next painted
    // glyph.
    let mut starts: Vec<u32> = run.glyphs.iter().map(|glyph| glyph.cluster).collect();
    starts.sort_unstable();
    starts.dedup();
    starts
}

/// Source text of the cluster starting at `cluster` (empty when unaddressable).
pub(crate) fn cluster_source_text(run: &GlyphRunPrimitive, starts: &[u32], cluster: u32) -> String {
    let start = cluster as usize;
    let end = starts
        .iter()
        .find(|next| **next > cluster)
        .map_or(run.text.len(), |next| *next as usize);
    run.text
        .get(start..end)
        .filter(|text| !text.is_empty())
        .map(str::to_owned)
        .or_else(|| {
            run.text
                .get(start..)
                .and_then(|text| text.chars().next())
                .map(String::from)
        })
        .unwrap_or_default()
}

fn collect_run(
    fonts: &mut BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>>,
    run: &GlyphRunPrimitive,
) {
    let entry = fonts.entry(run.font_id).or_default();
    let starts = cluster_starts(run);
    // A cluster can paint a base plus dot/mark components. Only the widest
    // glyph carries the source text; shared dot glyphs get a zero-width space
    // so extraction does not repeat the base letter.
    let mut bases: BTreeMap<u32, (f64, usize)> = BTreeMap::new();
    for (index, glyph) in run.glyphs.iter().enumerate() {
        if bases
            .get(&glyph.cluster)
            .is_none_or(|(advance, _)| glyph.advance > *advance)
        {
            bases.insert(glyph.cluster, (glyph.advance, index));
        }
    }
    for (index, glyph) in run.glyphs.iter().enumerate() {
        let is_base = bases
            .get(&glyph.cluster)
            .is_some_and(|(_, base)| *base == index);
        if !is_base {
            entry
                .entry(glyph.id as u16)
                .or_default()
                .insert("\u{200B}".to_owned());
            continue;
        }
        let text = cluster_source_text(run, &starts, glyph.cluster);
        // `.notdef` still belongs to the subset: a run of only missing glyphs
        // must embed a font the painter can address. Keep its ToUnicode entry
        // non-empty so extraction never yields a zero-width nothing.
        let text = if text.is_empty() && glyph.id == 0 {
            "\u{FFFD}".to_owned()
        } else {
            text
        };
        if !text.is_empty() {
            entry.entry(glyph.id as u16).or_default().insert(text);
        }
    }
}

/// Collect a glyph run's real glyphs for PDF/A-2u.
///
/// `.notdef` is never recorded: a base gid-0 cluster is left for
/// [`crate::text::collect_glyph_run_fallbacks_policy`] to re-resolve, and a
/// missing cluster fails there. A non-base gid-0 component cannot be
/// re-resolved (it carries no source scalar of its own), so it rejects here.
fn collect_run_policy(
    fonts: &mut BTreeMap<u32, BTreeMap<u16, BTreeSet<String>>>,
    run: &GlyphRunPrimitive,
) -> Result<(), PdfError> {
    let entry = fonts.entry(run.font_id).or_default();
    let starts = cluster_starts(run);
    let mut bases: BTreeMap<u32, (f64, usize)> = BTreeMap::new();
    for (index, glyph) in run.glyphs.iter().enumerate() {
        if bases
            .get(&glyph.cluster)
            .is_none_or(|(advance, _)| glyph.advance > *advance)
        {
            bases.insert(glyph.cluster, (glyph.advance, index));
        }
    }
    for (index, glyph) in run.glyphs.iter().enumerate() {
        let is_base = bases
            .get(&glyph.cluster)
            .is_some_and(|(_, base)| *base == index);
        if !is_base {
            if glyph.id == 0 {
                return Err(PdfError::Unsupported(PDFA_NOTDEF));
            }
            entry
                .entry(glyph.id as u16)
                .or_default()
                .insert("\u{200B}".to_owned());
            continue;
        }
        if glyph.id == 0 {
            continue;
        }
        let text = cluster_source_text(run, &starts, glyph.cluster);
        if text.is_empty() {
            return Err(PdfError::Unsupported(PDFA_NO_SOURCE_TEXT));
        }
        entry.entry(glyph.id as u16).or_default().insert(text);
    }
    Ok(())
}
