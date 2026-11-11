//! Font planning and subsetting for PDF export.

use std::collections::BTreeMap;

use docx_layout::display_list::{GlyphRunPrimitive, Primitive};
use ooxml_text::{FontId, FontStore};
use pdf_writer::types::{FontFlags, SystemInfo, UnicodeCmap};
use pdf_writer::{Name, Str};
use subsetter::GlyphRemapper;
use ttf_parser::{Face, GlyphId};

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

fn has_outlines(face: &Face) -> bool {
    face.tables().glyf.is_some() || face.tables().cff.is_some()
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
    uses: &BTreeMap<u16, String>,
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

impl EmbeddedFont {
    pub fn new(
        resource: String,
        store: &FontStore,
        id: u32,
        uses: &BTreeMap<u16, String>,
    ) -> Result<Self, FontError> {
        let font_id = FontId::from_u32(id);
        let bytes = store
            .font_bytes(font_id)
            .map_err(|error| FontError::Font(error.to_string()))?;
        let original = Face::parse(bytes, 0).map_err(|error| FontError::Font(error.to_string()))?;
        let upem = f32::from(original.units_per_em().max(1));
        let scale = 1000.0 / upem;

        let mut remapper = GlyphRemapper::new();
        for old in uses.keys().copied() {
            remapper.remap(old);
        }
        let (subset, remap, widths) = match subsetter::subset(bytes, 0, &remapper) {
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
                    if let Some(advance) = face.glyph_hor_advance(GlyphId(new)) {
                        if let Some(width) = widths.get_mut(usize::from(new)) {
                            *width = f32::from(advance) * scale;
                        }
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
            Err(error) => return Err(FontError::Subset(error.to_string())),
        };

        let mut cmap = UnicodeCmap::<u16>::new(
            Name(b"Custom"),
            SystemInfo {
                registry: Str(b"Adobe"),
                ordering: Str(b"Identity"),
                supplement: 0,
            },
        );
        for (old, text) in uses {
            if let Some(new) = remap.get(old) {
                cmap.pair_with_multiple(*new, text.chars());
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
            is_cff: original.tables().cff.is_some(),
            subset,
        })
    }

    pub fn glyph(&self, old: u32) -> Option<u16> {
        self.remap.get(&(old as u16)).copied()
    }
}

/// Visit every primitive a page paints, in paint order.
fn each_primitive(
    display_list: &docx_layout::display_list::DisplayList,
    mut visit: impl FnMut(&Primitive),
) {
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
            visit(primitive);
        }
    }
}

/// Collect every glyph a document paints, keyed by font id and original glyph id.
pub fn collect(
    display_list: &docx_layout::display_list::DisplayList,
) -> BTreeMap<u32, BTreeMap<u16, String>> {
    let mut fonts: BTreeMap<u32, BTreeMap<u16, String>> = BTreeMap::new();
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
) -> BTreeMap<u32, BTreeMap<u16, String>> {
    let mut fonts: BTreeMap<u32, BTreeMap<u16, String>> = BTreeMap::new();
    each_primitive(display_list, |primitive| match primitive {
        Primitive::Text(text) if !text.hidden && !text.text.is_empty() => {
            crate::text::collect_text_glyphs(store, &text.text, &mut fonts);
        }
        Primitive::GlyphRun(run) if !run.hidden => {
            collect_run(&mut fonts, run);
            crate::text::collect_glyph_run_fallbacks(store, run, &mut fonts);
        }
        _ => {}
    });
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

fn collect_run(fonts: &mut BTreeMap<u32, BTreeMap<u16, String>>, run: &GlyphRunPrimitive) {
    let entry = fonts.entry(run.font_id).or_default();
    let starts = cluster_starts(run);
    for glyph in &run.glyphs {
        let text = cluster_source_text(run, &starts, glyph.cluster);
        // `.notdef` still belongs to the subset: a run of only missing glyphs
        // must embed a font the painter can address. Keep its ToUnicode entry
        // non-empty so extraction never yields a zero-width nothing.
        let text = if text.is_empty() && glyph.id == 0 {
            "\u{FFFD}".to_owned()
        } else {
            text
        };
        entry.entry(glyph.id as u16).or_insert(text);
    }
}
