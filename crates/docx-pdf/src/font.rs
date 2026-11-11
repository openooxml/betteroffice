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
}

fn subset_tag(mut id: u32) -> String {
    let mut tag = String::with_capacity(6);
    for _ in 0..6 {
        tag.push(char::from(b'A' + (id % 26) as u8));
        id /= 26;
    }
    tag
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

        let mut remapper = GlyphRemapper::new();
        for old in uses.keys().copied() {
            remapper.remap(old);
        }
        let subset = subsetter::subset(bytes, 0, &remapper)
            .map_err(|error| FontError::Subset(error.to_string()))?;
        let face = Face::parse(&subset, 0).map_err(|error| FontError::Font(error.to_string()))?;
        let upem = f32::from(face.units_per_em().max(1));
        let scale = 1000.0 / upem;

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

        let mut cmap = UnicodeCmap::<u16>::new(
            Name(b"Custom"),
            SystemInfo {
                registry: Str(b"Adobe"),
                ordering: Str(b"Identity"),
                supplement: 0,
            },
        );
        for (old, text) in uses {
            if let Some(new) = remapper.get(*old) {
                cmap.pair_with_multiple(new, text.chars());
            }
        }
        let tounicode = cmap.finish().into_vec();

        let bbox = face.global_bounding_box();
        let bbox = [
            f32::from(bbox.x_min) * scale,
            f32::from(bbox.y_min) * scale,
            f32::from(bbox.x_max) * scale,
            f32::from(bbox.y_max) * scale,
        ];

        let mut flags = FontFlags::NON_SYMBOLIC;
        if face.is_italic() {
            flags |= FontFlags::ITALIC;
        }
        if face.is_bold() {
            flags |= FontFlags::FORCE_BOLD;
        }
        if face.is_monospaced() {
            flags |= FontFlags::FIXED_PITCH;
        }

        Ok(Self {
            resource,
            base_name: format!("{}+BO{id}", subset_tag(id)),
            remap,
            widths,
            tounicode,
            bbox,
            italic_angle: face.italic_angle(),
            ascent: f32::from(face.ascender()) * scale,
            descent: f32::from(face.descender()) * scale,
            cap_height: face
                .capital_height()
                .map(|value| f32::from(value) * scale)
                .unwrap_or(700.0),
            flags: flags.bits(),
            subset,
        })
    }

    pub fn glyph(&self, old: u32) -> Option<u16> {
        self.remap.get(&(old as u16)).copied()
    }
}

/// Collect every glyph a document paints, keyed by font id and original glyph id.
pub fn collect(
    display_list: &docx_layout::display_list::DisplayList,
) -> BTreeMap<u32, BTreeMap<u16, String>> {
    let mut fonts: BTreeMap<u32, BTreeMap<u16, String>> = BTreeMap::new();
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
            if let Primitive::GlyphRun(run) = primitive {
                collect_run(&mut fonts, run);
            }
        }
    }
    fonts
}

fn collect_run(fonts: &mut BTreeMap<u32, BTreeMap<u16, String>>, run: &GlyphRunPrimitive) {
    let entry = fonts.entry(run.font_id).or_default();
    for (index, glyph) in run.glyphs.iter().enumerate() {
        if glyph.id == 0 {
            continue;
        }
        let start = glyph.cluster as usize;
        let end = run.glyphs[index + 1..]
            .iter()
            .find(|next| next.cluster > glyph.cluster)
            .map(|next| next.cluster as usize)
            .unwrap_or(run.text.len());
        let text = run
            .text
            .get(start..end)
            .filter(|text| !text.is_empty())
            .map(str::to_owned)
            .or_else(|| {
                run.text
                    .get(start..)
                    .and_then(|text| text.chars().next())
                    .map(String::from)
            })
            .unwrap_or_default();
        entry.entry(glyph.id as u16).or_insert(text);
    }
}
