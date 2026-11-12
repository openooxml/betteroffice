//! Native PDF export for DOCX display lists.
//!
//! This crate deliberately does not depend on the VSDX writer. It consumes the
//! renderer-agnostic [`DisplayList`] produced by `docx-layout` and emits a
//! searchable PDF with embedded TrueType subsets. It covers page geometry,
//! positioned glyph runs, fallback text, rectangles, lines, decorations,
//! images (RGB plus soft masks), shape paths, headers/footers, note areas and
//! page borders. Links and tagged PDF are follow-up work.
//!
//! Plain PDF 1.7 is the default. [`PdfConformance::PdfA2u`] opts into archival
//! output through [`export_pdf_with_options`]; it requires embedded fonts and
//! rejects input the writer cannot represent conformantly instead of silently
//! emitting a non-conforming file.

use std::collections::{BTreeMap, HashMap};
use std::io::Write;

use docx_layout::display_list::{
    DecorationPrimitive, DisplayBorderStyle, DisplayList, DisplayPage, LinePrimitive,
    PageBorderPrimitive, PageBorderZOrder, Primitive, RectPrimitive, ShapePathCommand,
    ShapePrimitive, TextRunPrimitive,
};
use flate2::Compression;
use flate2::write::ZlibEncoder;
use ooxml_text::{FontId, FontStore};
use pdf_writer::types::{CidFontType, FontFlags, LineCapStyle, LineJoinStyle, SystemInfo};
use pdf_writer::{Content, Filter, Finish, Name, Pdf, Rect, Ref, Str};
use serde_json::Number;

use crate::font::{EmbeddedFont, FontError};

mod font;
mod image;
mod link;
// PDF/A-2u seam: `/Info` + XMP generation is owned by the metadata slice.
mod metadata;
pub(crate) mod options;
mod pdfa_limits;
// PDF/A-2u seam: sRGB ICC profile + `/OutputIntents` is owned by the color
// slice.
mod pdfa_color;
mod text;
mod units;

pub use options::{PdfConformance, PdfExportOptions, PdfMetadata};

/// PDF export failed.
#[derive(Debug)]
pub enum PdfError {
    /// The display list contains content this slice cannot represent yet.
    Unsupported(&'static str),
    /// A numeric field was not finite or outside the PDF range.
    InvalidNumber(&'static str),
    /// No pages were supplied.
    NoPages,
    /// A font could not be loaded or subset.
    Font(FontError),
    /// An image could not be decoded.
    Image(String),
}

impl From<FontError> for PdfError {
    fn from(error: FontError) -> Self {
        Self::Font(error)
    }
}

impl std::fmt::Display for PdfError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Unsupported(what) => write!(f, "unsupported PDF content: {what}"),
            Self::InvalidNumber(what) => write!(f, "invalid PDF number: {what}"),
            Self::NoPages => f.write_str("display list has no pages"),
            Self::Font(error) => write!(f, "{error}"),
            Self::Image(error) => write!(f, "image error: {error}"),
        }
    }
}

impl std::error::Error for PdfError {}

/// CSS pixels per PDF point (96 dpi / 72 dpi).
const PX_TO_PT: f32 = 72.0 / 96.0;
const FONT_NAME: Name<'static> = Name(b"F0");

#[derive(Clone)]
pub(crate) struct ImageRefs {
    pub(crate) resource: String,
    pub(crate) id: Ref,
    pub(crate) smask: Option<Ref>,
}

#[derive(Clone, Copy)]
enum ImagePart<'a> {
    Body,
    HeaderFooter(&'a str),
    Notes(Option<&'a str>),
}

fn scoped_image_key(part: ImagePart<'_>, rel_id: &str) -> String {
    match part {
        ImagePart::Body => format!("\u{1f}{rel_id}"),
        ImagePart::HeaderFooter(region_rel_id) => format!("hf:{region_rel_id}\u{1f}{rel_id}"),
        ImagePart::Notes(kind) => {
            let part = if kind == Some("endnote") {
                "endnotes"
            } else {
                "footnotes"
            };
            format!("{part}\u{1f}{rel_id}")
        }
    }
}

fn image_refs<'a>(
    images: &'a BTreeMap<String, ImageRefs>,
    part: ImagePart<'_>,
    rel_id: &str,
) -> Option<&'a ImageRefs> {
    images
        .get(&scoped_image_key(part, rel_id))
        .or_else(|| images.get(rel_id))
}

#[derive(Clone, Copy)]
struct FontRefs {
    type0: Ref,
    cid: Ref,
    descriptor: Ref,
    file: Ref,
    to_unicode: Ref,
    cid_to_gid: Ref,
    /// PDF/A-2u `/CIDSet` stream; `None` for Standard output.
    cid_set: Option<Ref>,
}

/// Object ids reserved for the PDF/A structures the catalog and trailer must
/// reference while they are written.
///
/// The output intent itself is a direct dictionary in the catalog's
/// `/OutputIntents` array, so only its `/DestOutputProfile` stream needs an
/// id. All three ids come from the single object counter that also allocates
/// fonts, images, opacity states, pages and link annotations.
struct PdfAObjects {
    /// Trailer `/Info` dictionary.
    info: Ref,
    /// Catalog `/Metadata` XMP stream.
    xmp: Ref,
    /// `/OutputIntents[0] /DestOutputProfile` ICC stream.
    profile: Ref,
}

impl PdfAObjects {
    /// Reserve the ids in a fixed order and advance the caller's counter past
    /// them.
    fn reserve(next_id: &mut i32) -> Self {
        let info = Ref::new(*next_id);
        let xmp = Ref::new(*next_id + 1);
        let profile = Ref::new(*next_id + 2);
        *next_id += 3;
        Self { info, xmp, profile }
    }
}

/// Deterministic 128-bit file identifier accumulated from emitted content.
///
/// PDF/A requires `/ID` in the trailer. Deriving it from the document content
/// makes repeated exports of the same input reproduce the same identifier and
/// different input produce a different one. FNV-1a is used because its output
/// is stable across Rust releases and it needs no additional dependency.
struct DocumentId {
    low: u64,
    high: u64,
}

impl DocumentId {
    const FNV_OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
    const FNV_PRIME: u64 = 0x0000_0100_0000_01b3;

    fn new() -> Self {
        Self {
            low: Self::FNV_OFFSET,
            high: Self::FNV_OFFSET.rotate_left(32) ^ 0xa5a5_5a5a_1234_5678,
        }
    }

    /// Hash one chunk with a length prefix so concatenations stay unambiguous.
    fn update(&mut self, bytes: &[u8]) {
        self.mix(b'\x1e');
        for byte in (bytes.len() as u64).to_le_bytes() {
            self.mix(byte);
        }
        for &byte in bytes {
            self.mix(byte);
        }
    }

    /// Hash a field discriminator.
    fn update_tag(&mut self, tag: u8) {
        self.mix(tag);
    }

    fn mix(&mut self, byte: u8) {
        let byte = u64::from(byte);
        self.low = (self.low ^ byte).wrapping_mul(Self::FNV_PRIME);
        self.high = (self.high ^ byte)
            .wrapping_mul(Self::FNV_PRIME)
            .rotate_left(13);
    }

    /// Hash every metadata field, distinguishing absent from empty values.
    fn update_metadata(&mut self, metadata: &PdfMetadata) {
        for (tag, value) in [
            (1, &metadata.title),
            (2, &metadata.author),
            (3, &metadata.subject),
            (4, &metadata.keywords),
            (5, &metadata.creator),
            (6, &metadata.creation_date),
            (7, &metadata.modification_date),
        ] {
            self.update_tag(tag);
            match value {
                Some(value) => {
                    self.update_tag(1);
                    self.update(value.as_bytes());
                }
                None => self.update_tag(0),
            }
        }
    }

    fn finish(&self) -> Vec<u8> {
        let mut id = Vec::with_capacity(16);
        id.extend_from_slice(&self.low.to_be_bytes());
        id.extend_from_slice(&self.high.to_be_bytes());
        id
    }
}

/// Export a display list as a plain, untagged PDF.
///
/// Text primitives use a Base-14 font. Glyph runs require
/// [`export_pdf_with_fonts`].
pub fn export_pdf(display_list: &DisplayList) -> Result<Vec<u8>, PdfError> {
    export_pdf_with_options(
        display_list,
        None,
        &BTreeMap::new(),
        None,
        &PdfExportOptions::default(),
    )
}

/// Export a display list with embedded, subset TrueType fonts.
///
/// `store` must be the same font registry that produced the display list's
/// `font_id`s.
pub fn export_pdf_with_fonts(
    display_list: &DisplayList,
    store: &FontStore,
) -> Result<Vec<u8>, PdfError> {
    export_pdf_with_options(
        display_list,
        Some(store),
        &BTreeMap::new(),
        None,
        &PdfExportOptions::default(),
    )
}

/// Export with embedded fonts and registered image bytes.
///
/// Keys are the writer's scoped relationship keys (`"\u{1f}rId5"` for body,
/// `"hf:<part>\u{1f}rId5"` for header/footer, `"footnotes\u{1f}rId5"` or
/// `"endnotes\u{1f}rId5"` for notes). An unscoped `rel_id` and a `data:` URL
/// in the primitive are accepted as fallbacks.
pub fn export_pdf_with_fonts_and_images(
    display_list: &DisplayList,
    store: &FontStore,
    images: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<u8>, PdfError> {
    export_pdf_with_options(
        display_list,
        Some(store),
        images,
        None,
        &PdfExportOptions::default(),
    )
}

/// Export with embedded fonts, registered image bytes, and the facade's
/// ordered font chains.
///
/// `chains` is keyed `family|bold|italic` (trimmed, lowercased family) and
/// orders the faces layout resolved for each run. Painting and subsetting both
/// use that order, keeping an application or embedded alias ahead of the
/// bundled face it shadows. Scalars no chain face covers still fall back to
/// registry order.
pub fn export_pdf_with_fonts_images_and_chains(
    display_list: &DisplayList,
    store: &FontStore,
    images: &BTreeMap<String, Vec<u8>>,
    chains: &HashMap<String, Vec<FontId>>,
) -> Result<Vec<u8>, PdfError> {
    export_pdf_with_options(
        display_list,
        Some(store),
        images,
        Some(chains),
        &PdfExportOptions::default(),
    )
}

/// Export with explicit conformance and document metadata.
///
/// `store` must be the same font registry that produced the display list's
/// `font_id`s. [`PdfConformance::Standard`] keeps the historical behavior and
/// accepts `None` (Base-14 fallback text). [`PdfConformance::PdfA2u`] requires
/// a store and rejects the Base-14 fallback whenever the document contains
/// text, returning [`PdfError::Unsupported`] for input it cannot represent
/// conformantly; it never silently downgrades.
///
/// `images` and `chains` have the same meaning as in
/// [`export_pdf_with_fonts_images_and_chains`].
pub fn export_pdf_with_options(
    display_list: &DisplayList,
    store: Option<&FontStore>,
    images: &BTreeMap<String, Vec<u8>>,
    chains: Option<&HashMap<String, Vec<FontId>>>,
    options: &PdfExportOptions,
) -> Result<Vec<u8>, PdfError> {
    export_pdf_inner(display_list, store, images, chains, options)
}

fn export_pdf_inner(
    display_list: &DisplayList,
    store: Option<&FontStore>,
    images: &BTreeMap<String, Vec<u8>>,
    chains: Option<&HashMap<String, Vec<FontId>>>,
    options: &PdfExportOptions,
) -> Result<Vec<u8>, PdfError> {
    if display_list.pages.is_empty() {
        return Err(PdfError::NoPages);
    }

    let mut pdf = Pdf::new();
    // PDF/A-2 is based on PDF 1.7. pdf-writer defaults to a 1.7 header, but
    // archival output states it explicitly instead of depending on that
    // default; Standard output keeps the default bytes.
    if options.conformance.is_pdfa() {
        pdf.set_version(1, 7);
    }
    let catalog_id = Ref::new(1);
    let page_tree_id = Ref::new(2);
    let text_font_id = Ref::new(3);
    let has_builtin_text = store.is_none_or(|store| !text::store_has_usable_face(store));
    let has_text = has_builtin_text
        && display_list.pages.iter().any(|page| {
            page.primitives
                .iter()
                .chain(
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
                )
                .any(|primitive| matches!(primitive, Primitive::Text(_)))
        });
    // PDF/A-2u seam: enforce archival preconditions before any object is
    // written. A `PdfA2u` request must not fall back to Base-14 and must not
    // be downgraded to `Standard`; unsupported input is reported instead.
    // Empty metadata is accepted: ISO 19005-2 does not require a title (that
    // is a PDF/UA requirement), and the metadata writer omits absent fields.
    if options.conformance.is_pdfa() {
        if store.is_none() {
            return Err(PdfError::Unsupported(
                "PDF/A-2u requires embedded fonts (a FontStore)",
            ));
        }
        if has_text {
            return Err(PdfError::Unsupported(
                "PDF/A-2u forbids the Base-14 text fallback",
            ));
        }
    }

    // PDF/A-2u seam: the precondition above guarantees `has_text` is false for
    // `PdfA2u`, so no Base-14 object or `/F0` resource can appear in archival
    // output. The explicit gate keeps that invariant if the precondition moves.
    let has_text = has_text && !options.conformance.is_pdfa();

    let mut next_id = if has_text { 4 } else { 3 };
    // PDF/A-2u seam: reserved ids draw from the same counter as fonts,
    // images, opacity states, pages and link annotations, so they stay
    // deterministic and collision-free.
    let pdfa = options
        .conformance
        .is_pdfa()
        .then(|| PdfAObjects::reserve(&mut next_id));
    // PDF/A-2u seam: `/ID` is a deterministic content hash set on the trailer
    // before `finish`. Standard exports keep pdf-writer's default trailer
    // without `/ID`, so their bytes do not change.
    let mut file_id = pdfa.is_some().then(DocumentId::new);

    {
        let mut catalog = pdf.catalog(catalog_id);
        catalog.pages(page_tree_id);
        if let Some(pdfa) = &pdfa {
            catalog.metadata(pdfa.xmp);
            // PDF/A-2u seam: the output intent is a direct dictionary in the
            // catalog's `/OutputIntents` array; `profile` is the reserved ICC
            // stream referenced by `/DestOutputProfile`.
            pdfa_color::attach_output_intent(&mut catalog, pdfa.profile)?;
        }
    }
    if has_text {
        pdf.type1_font(text_font_id).base_font(Name(b"Helvetica"));
    }
    if let Some(pdfa) = &pdfa {
        if let Some(file_id) = file_id.as_mut() {
            file_id.update(b"PDF/A-2u");
            file_id.update_metadata(&options.metadata);
        }
        // PDF/A-2u seam: `/Info` and the XMP stream are owned by the metadata
        // slice; this slice reserves their ids and registers them with the
        // trailer and catalog above.
        metadata::write_document_info(&mut pdf, pdfa.info, &options.metadata)?;
        metadata::write_xmp_metadata(&mut pdf, pdfa.xmp, options.conformance, &options.metadata)?;
        pdfa_color::write_output_profile(&mut pdf, pdfa.profile)?;
    }

    let font_uses = match (store, chains) {
        (Some(store), Some(chains)) if !chains.is_empty() => {
            font::collect_with_chains_policy(display_list, store, chains, options.conformance)?
        }
        (Some(store), _) => {
            font::collect_with_text_policy(display_list, store, options.conformance)?
        }
        (None, _) => font::collect(display_list),
    };
    if store.is_none() && !font_uses.is_empty() {
        return Err(PdfError::Unsupported("glyph runs need embedded fonts"));
    }

    let mut font_map = BTreeMap::new();
    let mut font_objects = Vec::new();
    if let Some(store) = store {
        for (index, (font_id, glyphs)) in font_uses.iter().enumerate() {
            if glyphs.is_empty() {
                continue;
            }
            let resource = format!("F{}", index + 1);
            let embedded = EmbeddedFont::new_with_conformance(
                resource,
                store,
                *font_id,
                glyphs,
                options.conformance,
            )?;
            if let Some(file_id) = file_id.as_mut() {
                file_id.update(embedded.base_name.as_bytes());
                file_id.update(&embedded.subset);
                file_id.update(&embedded.tounicode);
                if let Some(map) = &embedded.cid_to_gid {
                    for gid in map {
                        file_id.update(&gid.to_be_bytes());
                    }
                }
            }
            let refs = FontRefs {
                type0: Ref::new(next_id),
                cid: Ref::new(next_id + 1),
                descriptor: Ref::new(next_id + 2),
                file: Ref::new(next_id + 3),
                to_unicode: Ref::new(next_id + 4),
                cid_to_gid: Ref::new(next_id + 5),
                // PDF/A-2u seam: `/CIDSet` needs its own stream. Reserve it
                // only for archival output so Standard ref numbering (and
                // therefore bytes) stay unchanged.
                cid_set: options.conformance.is_pdfa().then(|| Ref::new(next_id + 6)),
            };
            next_id += if options.conformance.is_pdfa() { 7 } else { 6 };
            font_objects.push((*font_id, refs));
            font_map.insert(*font_id, embedded);
        }
    }

    for (font_id, refs) in &font_objects {
        write_font(&mut pdf, &font_map[font_id], *refs, options.conformance);
    }

    let mut image_map: BTreeMap<String, ImageRefs> = BTreeMap::new();
    let mut image_objects = Vec::new();
    if store.is_some() {
        for page in &display_list.pages {
            for primitive in &page.primitives {
                collect_image(
                    primitive,
                    ImagePart::Body,
                    images,
                    store,
                    &mut image_map,
                    &mut image_objects,
                    &mut next_id,
                );
            }
            if let Some(header) = &page.header {
                for primitive in &header.primitives {
                    collect_image(
                        primitive,
                        ImagePart::HeaderFooter(&header.r_id),
                        images,
                        store,
                        &mut image_map,
                        &mut image_objects,
                        &mut next_id,
                    );
                }
            }
            if let Some(footer) = &page.footer {
                for primitive in &footer.primitives {
                    collect_image(
                        primitive,
                        ImagePart::HeaderFooter(&footer.r_id),
                        images,
                        store,
                        &mut image_map,
                        &mut image_objects,
                        &mut next_id,
                    );
                }
            }
            for note in &page.note_areas {
                for primitive in note
                    .separator_primitives
                    .iter()
                    .chain(note.primitives.iter())
                {
                    collect_image(
                        primitive,
                        ImagePart::Notes(note.kind.as_deref()),
                        images,
                        store,
                        &mut image_map,
                        &mut image_objects,
                        &mut next_id,
                    );
                }
            }
        }
    }

    for (refs, decoded) in &image_objects {
        if let Some(file_id) = file_id.as_mut() {
            file_id.update(&(decoded.width as u32).to_be_bytes());
            file_id.update(&(decoded.height as u32).to_be_bytes());
            file_id.update(&decoded.rgb);
            if let Some(alpha) = &decoded.alpha {
                file_id.update(alpha);
            }
        }
        write_image(&mut pdf, refs, decoded);
    }

    let mut opacity_map: BTreeMap<u32, (String, Ref)> = BTreeMap::new();
    for page in &display_list.pages {
        for primitive in all_primitives(page) {
            let Some(key) = primitive_opacity(primitive) else {
                continue;
            };
            if opacity_map.contains_key(&key) {
                continue;
            }
            let name = format!("GS{}", opacity_map.len() + 1);
            let id = Ref::new(next_id);
            next_id += 1;
            let alpha = key as f32 / 1000.0;
            pdf.ext_graphics(id)
                .non_stroking_alpha(alpha)
                .stroking_alpha(alpha);
            opacity_map.insert(key, (name, id));
        }
    }

    let mut page_ids = Vec::with_capacity(display_list.pages.len());
    for page in &display_list.pages {
        let page_id = Ref::new(next_id);
        let content_id = Ref::new(next_id + 1);
        next_id += 2;

        let links = link::collect_links(page);
        let annotation_ids: Vec<Ref> = (0..links.len())
            .map(|index| Ref::new(next_id + index as i32))
            .collect();
        next_id += links.len() as i32;

        let width = points(&page.width, "page width")?;
        let height = points(&page.height, "page height")?;
        let content = page_content(
            page,
            width,
            height,
            &font_map,
            &image_map,
            store,
            chains,
            &opacity_map,
            options.conformance,
        )?;
        if let Some(file_id) = file_id.as_mut() {
            file_id.update(&width.to_le_bytes());
            file_id.update(&height.to_le_bytes());
            // Hash the uncompressed content so `/ID` does not depend on the
            // deflate implementation.
            file_id.update(&content);
            for link in &links {
                for value in link.rect {
                    file_id.update(&value.to_le_bytes());
                }
                if let link::LinkAction::Uri(uri) = &link.action {
                    file_id.update(uri.as_bytes());
                }
                if let Some(title) = &link.title {
                    file_id.update(title.as_bytes());
                }
            }
        }
        let content = deflate(&content);
        pdf.stream(content_id, &content).filter(Filter::FlateDecode);
        let mut writer = pdf.page(page_id);
        writer
            .parent(page_tree_id)
            .media_box(Rect::new(0.0, 0.0, width, height));
        writer.contents(content_id);
        if !annotation_ids.is_empty() {
            writer.annotations(annotation_ids.iter().copied());
        }
        let mut resources = writer.resources();
        let mut fonts = resources.fonts();
        if has_text {
            fonts.pair(FONT_NAME, text_font_id);
        }
        for (font_id, refs) in &font_objects {
            let resource = font_map[font_id].resource.as_bytes();
            fonts.pair(Name(resource), refs.type0);
        }
        fonts.finish();
        if !image_map.is_empty() {
            let mut xobjects = resources.x_objects();
            for refs in image_map.values() {
                xobjects.pair(Name(refs.resource.as_bytes()), refs.id);
            }
            xobjects.finish();
        }
        if !opacity_map.is_empty() {
            let mut states = resources.ext_g_states();
            for (name, id) in opacity_map.values() {
                states.pair(Name(name.as_bytes()), *id);
            }
            states.finish();
        }
        resources.finish();
        writer.finish();
        for (id, annotation) in annotation_ids.iter().zip(&links) {
            // PDF/A-2u seam: the link slice adds `/F 4` (print flag) for
            // archival annotations; Standard output stays byte-identical.
            link::write_annotation(&mut pdf, *id, page_id, annotation, options.conformance)?;
        }
        page_ids.push(page_id);
    }

    pdf.pages(page_tree_id)
        .kids(page_ids.iter().copied())
        .count(page_ids.len() as i32);
    if let Some(file_id) = file_id {
        // PDF/A-2u seam: both `/ID` strings are identical for the first
        // revision; a later incremental update would change the second only.
        let id = DocumentId::finish(&file_id);
        pdf.set_file_id((id.clone(), id));
    }
    Ok(pdf.finish())
}

fn deflate(data: &[u8]) -> Vec<u8> {
    let mut encoder = ZlibEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(data).expect("in-memory deflate");
    encoder.finish().expect("in-memory deflate")
}

fn write_font(pdf: &mut Pdf, font: &EmbeddedFont, refs: FontRefs, conformance: PdfConformance) {
    let base = Name(font.base_name.as_bytes());
    pdf.type0_font(refs.type0)
        .base_font(base)
        .encoding_predefined(Name(b"Identity-H"))
        .descendant_font(refs.cid)
        .to_unicode(refs.to_unicode);

    {
        let mut cid = pdf.cid_font(refs.cid);
        let subtype = if font.is_cff {
            CidFontType::Type0
        } else {
            CidFontType::Type2
        };
        cid.subtype(subtype)
            .base_font(base)
            .system_info(SystemInfo {
                registry: Str(b"Adobe"),
                ordering: Str(b"Identity"),
                supplement: 0,
            })
            .font_descriptor(refs.descriptor)
            .default_width(1000.0);
        if font.cid_to_gid.is_some() {
            cid.cid_to_gid_map_stream(refs.cid_to_gid);
        } else if !font.is_cff {
            cid.cid_to_gid_map_predefined(Name(b"Identity"));
        }
        // PDF/A-2u seam: `/CIDSet` is allocated with the font refs and
        // written below from the final subset plan; Standard has no ref and
        // writes no descriptor entry.
        cid.widths().consecutive(0, font.widths.iter().copied());
    }

    {
        let mut descriptor = pdf.font_descriptor(refs.descriptor);
        descriptor
            .name(base)
            .flags(FontFlags::from_bits_retain(font.flags))
            .bbox(Rect::new(
                font.bbox[0],
                font.bbox[1],
                font.bbox[2],
                font.bbox[3],
            ))
            .italic_angle(font.italic_angle)
            .ascent(font.ascent)
            .descent(font.descent)
            .cap_height(font.cap_height)
            .stem_v(80.0);
        if conformance.is_pdfa()
            && let Some(cid_set) = refs.cid_set
        {
            descriptor.cid_set(cid_set);
        }
        if font.is_cff {
            descriptor.font_file3(refs.file);
        } else {
            descriptor.font_file2(refs.file);
        }
    }

    let subset = deflate(&font.subset);
    let mut stream = pdf.stream(refs.file, &subset);
    stream.filter(Filter::FlateDecode);
    if font.is_cff {
        stream.pair(Name(b"Subtype"), Name(b"OpenType"));
    } else {
        stream.pair(Name(b"Length1"), font.subset.len() as i32);
    }
    drop(stream);
    if let Some(map) = &font.cid_to_gid {
        let mut bytes = Vec::with_capacity(map.len() * 2);
        for gid in map {
            bytes.extend_from_slice(&gid.to_be_bytes());
        }
        let bytes = deflate(&bytes);
        let mut stream = pdf.stream(refs.cid_to_gid, &bytes);
        stream.filter(Filter::FlateDecode);
    }
    pdf.cmap(refs.to_unicode, &font.tounicode);
    if conformance.is_pdfa()
        && let Some(cid_set) = refs.cid_set
    {
        // PDF/A-2u seam: deterministic one-bit-per-CID payload computed from
        // the final subset, including duplicate CIDs for distinct sources.
        pdf.stream(cid_set, &font.cid_set());
    }
}

fn page_content(
    page: &DisplayPage,
    width: f32,
    height: f32,
    fonts: &BTreeMap<u32, EmbeddedFont>,
    images: &BTreeMap<String, ImageRefs>,
    store: Option<&FontStore>,
    chains: Option<&HashMap<String, Vec<FontId>>>,
    opacities: &BTreeMap<u32, (String, Ref)>,
    conformance: PdfConformance,
) -> Result<Vec<u8>, PdfError> {
    let mut content = Content::new();
    if let Some(rgb) = page.background.as_deref().and_then(parse_color) {
        content.set_fill_rgb(rgb[0], rgb[1], rgb[2]);
        content.rect(0.0, 0.0, width, height);
        content.fill_nonzero();
    }

    for border in page
        .page_borders
        .iter()
        .filter(|border| border.z_order == Some(PageBorderZOrder::Back))
    {
        paint_page_border(&mut content, page, border)?;
    }

    let watermark_count = page
        .watermark_primitive_count
        .unwrap_or(0)
        .min(page.primitives.len());
    let (watermark, body) = page.primitives.split_at(watermark_count);
    paint_primitives(
        &mut content,
        page,
        watermark,
        fonts,
        images,
        store,
        chains,
        opacities,
        conformance,
        ImagePart::Body,
    )?;
    if let Some(header) = &page.header {
        paint_primitives(
            &mut content,
            page,
            &header.primitives,
            fonts,
            images,
            store,
            chains,
            opacities,
            conformance,
            ImagePart::HeaderFooter(&header.r_id),
        )?;
    }
    if let Some(footer) = &page.footer {
        paint_primitives(
            &mut content,
            page,
            &footer.primitives,
            fonts,
            images,
            store,
            chains,
            opacities,
            conformance,
            ImagePart::HeaderFooter(&footer.r_id),
        )?;
    }
    paint_primitives(
        &mut content,
        page,
        body,
        fonts,
        images,
        store,
        chains,
        opacities,
        conformance,
        ImagePart::Body,
    )?;
    for note in &page.note_areas {
        paint_primitives(
            &mut content,
            page,
            &note.separator_primitives,
            fonts,
            images,
            store,
            chains,
            opacities,
            conformance,
            ImagePart::Notes(note.kind.as_deref()),
        )?;
        paint_primitives(
            &mut content,
            page,
            &note.primitives,
            fonts,
            images,
            store,
            chains,
            opacities,
            conformance,
            ImagePart::Notes(note.kind.as_deref()),
        )?;
    }
    for border in page
        .page_borders
        .iter()
        .filter(|border| border.z_order != Some(PageBorderZOrder::Back))
    {
        paint_page_border(&mut content, page, border)?;
    }

    Ok(content.finish().into_vec())
}

fn all_primitives(page: &DisplayPage) -> impl Iterator<Item = &Primitive> {
    page.primitives.iter().chain(
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
    )
}

fn primitive_opacity(primitive: &Primitive) -> Option<u32> {
    let value = match primitive {
        Primitive::Text(text) => text.opacity.as_ref(),
        Primitive::GlyphRun(run) => run.opacity.as_ref(),
        Primitive::Line(line) => line.opacity.as_ref(),
        Primitive::Image(image) => {
            return image::opacity_alpha(image)
                .map(|alpha| (alpha * 1000.0).round() as u32)
                .filter(|key| *key < 1000);
        }
        _ => None,
    }?;
    let alpha = value.as_f64()?;
    if !alpha.is_finite() {
        return None;
    }
    let alpha = alpha.clamp(0.0, 1.0);
    if alpha >= 1.0 {
        return None;
    }
    Some((alpha * 1000.0).round() as u32)
}

fn paint_primitives(
    content: &mut Content,
    page: &DisplayPage,
    primitives: &[Primitive],
    fonts: &BTreeMap<u32, EmbeddedFont>,
    images: &BTreeMap<String, ImageRefs>,
    store: Option<&FontStore>,
    chains: Option<&HashMap<String, Vec<FontId>>>,
    opacities: &BTreeMap<u32, (String, Ref)>,
    conformance: PdfConformance,
    part: ImagePart<'_>,
) -> Result<(), PdfError> {
    for primitive in primitives {
        let opacity =
            primitive_opacity(primitive).and_then(|key| opacities.get(&key).map(|(name, _)| name));
        if let Some(name) = opacity {
            content.save_state();
            content.set_parameters(Name(name.as_bytes()));
        }
        match primitive {
            Primitive::Text(text) => match store {
                Some(store) => match chains {
                    // PDF/A-2u seam: policy painters reject `.notdef`, missing
                    // coverage and unmappable source text. Standard keeps the
                    // historical painters untouched.
                    Some(chains) if conformance.is_pdfa() => text::paint_text_with_chains_policy(
                        content,
                        page,
                        text,
                        fonts,
                        store,
                        Some(chains),
                        conformance,
                    )?,
                    Some(chains) => text::paint_text_with_chains(
                        content,
                        page,
                        text,
                        fonts,
                        store,
                        Some(chains),
                    )?,
                    None if conformance.is_pdfa() => {
                        text::paint_text_policy(content, page, text, fonts, store, conformance)?
                    }
                    None => text::paint_text(content, page, text, fonts, store)?,
                },
                None => paint_text(content, page, text, conformance)?,
            },
            Primitive::Rect(rect) => paint_rect(content, page, rect)?,
            Primitive::Line(line) => paint_line(content, page, line)?,
            // PDF/A-2u seam: policy painting rejects a cluster that would
            // still paint `.notdef` after fallback resolution.
            Primitive::GlyphRun(run) => match store {
                Some(store) if conformance.is_pdfa() => text::paint_glyph_run_with_store_policy(
                    content,
                    page,
                    run,
                    fonts,
                    store,
                    conformance,
                )?,
                Some(store) => text::paint_glyph_run_with_store(content, page, run, fonts, store)?,
                None => text::paint_glyph_run(content, page, run, fonts)?,
            },
            Primitive::Image(image) => {
                if let Some(refs) = image_refs(images, part, &image.rel_id) {
                    let page_height = points(&page.height, "page height")?;
                    image::paint_image(content, page_height, image, refs)?;
                }
            }
            Primitive::Shape(shape) => paint_shape(content, page, shape)?,
            Primitive::Decoration(deco) => paint_decoration(content, page, deco)?,
        }
        if opacity.is_some() {
            content.restore_state();
        }
    }
    Ok(())
}

fn paint_page_border(
    content: &mut Content,
    page: &DisplayPage,
    border: &PageBorderPrimitive,
) -> Result<(), PdfError> {
    let x = points(&border.x, "border x")?;
    let y = points(&border.y, "border y")?;
    let w = points(&border.w, "border width")?;
    let h = points(&border.h, "border height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(());
    }
    let page_height = points(&page.height, "page height")?;
    let sides = [
        (
            border.top.as_ref(),
            BorderSegment {
                x1: x,
                y1: page_height - y,
                x2: x + w,
                y2: page_height - y,
            },
        ),
        (
            border.bottom.as_ref(),
            BorderSegment {
                x1: x,
                y1: page_height - y - h,
                x2: x + w,
                y2: page_height - y - h,
            },
        ),
        (
            border.left.as_ref(),
            BorderSegment {
                x1: x,
                y1: page_height - y,
                x2: x,
                y2: page_height - y - h,
            },
        ),
        (
            border.right.as_ref(),
            BorderSegment {
                x1: x + w,
                y1: page_height - y,
                x2: x + w,
                y2: page_height - y - h,
            },
        ),
    ];
    for (side, segment) in sides {
        let Some(side) = side else {
            continue;
        };
        let Some(color) = parse_color(&side.color) else {
            continue;
        };
        let Some(style) = page_border_style(&side.style) else {
            return Err(PdfError::Unsupported("unsupported page border style"));
        };
        let width_px = raw_px(&side.width, "border side width")?;
        content.save_state();
        paint_border_recipe(content, segment, width_px, color, style, None, None)?;
        content.restore_state();
    }
    Ok(())
}

/// The page-border styles the display list can carry (see
/// `docx_layout::display_list::page_border_style`); everything else is
/// rejected rather than silently painted as a solid rule.
fn page_border_style(style: &str) -> Option<DisplayBorderStyle> {
    Some(match style {
        "solid" => DisplayBorderStyle::Solid,
        "double" => DisplayBorderStyle::Double,
        "dotted" => DisplayBorderStyle::Dotted,
        "dashed" => DisplayBorderStyle::Dashed,
        "groove" => DisplayBorderStyle::Groove,
        "ridge" => DisplayBorderStyle::Ridge,
        "inset" => DisplayBorderStyle::Inset,
        "outset" => DisplayBorderStyle::Outset,
        _ => return None,
    })
}

fn paint_text(
    content: &mut Content,
    page: &DisplayPage,
    text: &TextRunPrimitive,
    conformance: PdfConformance,
) -> Result<(), PdfError> {
    // PDF/A-2u seam: defense in depth behind the top-level precondition.
    // Base-14 fonts are not embedded, so archival output must never reach
    // this painter.
    if conformance.is_pdfa() {
        return Err(PdfError::Unsupported(
            "PDF/A-2u forbids the Base-14 text fallback",
        ));
    }
    if text.hidden || text.text.is_empty() {
        return Ok(());
    }
    if !text.text.is_ascii() {
        return Err(PdfError::Unsupported(
            "non-ASCII text needs an embedded font",
        ));
    }
    let size = font_size_px(&text.font) * PX_TO_PT;
    let x = points(&text.x, "text x")?;
    let baseline = points(&text.baseline_y, "text baseline")?;
    let page_height = points(&page.height, "page height")?;
    let width = points(&text.width, "text width")?;
    let color = parse_color(&text.color).unwrap_or([0.0, 0.0, 0.0]);
    let rotation = optional_rotation(text.rotation_deg.as_ref());
    let horizontal_scale = checked_scale(text.horizontal_scale.as_ref())?;
    let origin_y = page_height - baseline;
    // The same pivot `text.rs` uses: the measured run box centred on x and
    // 0.2em above the baseline, which is `baseline - 0.2em` in y-down CSS.
    let center_x = x + width / 2.0;
    let center_y = origin_y + size * 0.2;

    content.begin_text();
    content.set_font(FONT_NAME, size);
    content.set_fill_rgb(color[0], color[1], color[2]);
    // Tz is text state and persists across BT/ET; always write it so one
    // scaled run cannot stretch every later run on the page.
    content.set_horizontal_scaling(horizontal_scale);
    if text.rtl == Some(true) {
        // `width` already includes the horizontal scale; `Tz` never applies
        // to a text matrix translation, so the run starts at its measured
        // right edge and each byte takes an unscaled share of `width`.
        let advance = f64::from(width) / text.text.len().max(1) as f64;
        let mut cursor = f64::from(x) + f64::from(width);
        for byte in text.text.bytes() {
            cursor -= advance;
            content.set_text_matrix(base14_text_matrix(
                x,
                origin_y,
                cursor - f64::from(x),
                rotation,
                center_x,
                center_y,
            ));
            content.show(Str(&[byte]));
        }
    } else {
        content.set_text_matrix(base14_text_matrix(
            x, origin_y, 0.0, rotation, center_x, center_y,
        ));
        content.show(Str(text.text.as_bytes()));
    }
    content.end_text();
    Ok(())
}

/// Absent or non-finite rotation means an unrotated run, like `text.rs`.
fn optional_rotation(value: Option<&Number>) -> f64 {
    value
        .and_then(Number::as_f64)
        .filter(|value| value.is_finite())
        .unwrap_or(0.0)
}

/// Validate a CSS `horizontal_scale` percentage.
///
/// Absent means the CSS default. A present value must be positive and finite:
/// `0` would collapse every advance and NaN has no visual meaning, so both are
/// refused instead of silently defaulted, matching the embedded-font path.
fn checked_scale(value: Option<&Number>) -> Result<f32, PdfError> {
    let Some(value) = value else {
        return Ok(100.0);
    };
    let scale = value
        .as_f64()
        .ok_or(PdfError::InvalidNumber("horizontal scale"))? as f32;
    if scale.is_finite() && scale > 0.0 {
        Ok(scale)
    } else {
        Err(PdfError::InvalidNumber("horizontal scale"))
    }
}

/// Text matrix for a Base-14 run: the pen origin with `rotation_deg` applied
/// about the run centre.
///
/// CSS rotates clockwise on a y-down page; PDF's y-up space turns the same
/// visual angle counterclockwise, so the angle is negated. `offset` is the
/// pen advance along the baseline from `origin_x`; `Tz` never scales a text
/// matrix translation, so the caller keeps `offset` unscaled.
fn base14_text_matrix(
    origin_x: f32,
    origin_y: f32,
    offset: f64,
    rotation_deg: f64,
    center_x: f32,
    center_y: f32,
) -> [f32; 6] {
    let point_x = origin_x + offset as f32;
    if rotation_deg.abs() <= f64::EPSILON || !rotation_deg.is_finite() {
        return [1.0, 0.0, 0.0, 1.0, point_x, origin_y];
    }
    let theta = -rotation_deg.to_radians();
    let (sin, cos) = theta.sin_cos();
    let (a, b, c, d) = (cos as f32, sin as f32, -sin as f32, cos as f32);
    let dx = point_x - center_x;
    let dy = origin_y - center_y;
    [
        a,
        b,
        c,
        d,
        center_x + a * dx + c * dy,
        center_y + b * dx + d * dy,
    ]
}

fn collect_image(
    primitive: &Primitive,
    part: ImagePart<'_>,
    images: &BTreeMap<String, Vec<u8>>,
    store: Option<&FontStore>,
    image_map: &mut BTreeMap<String, ImageRefs>,
    image_objects: &mut Vec<(ImageRefs, image::DecodedImage)>,
    next_id: &mut i32,
) {
    let Primitive::Image(image) = primitive else {
        return;
    };
    let key = scoped_image_key(part, &image.rel_id);
    if image_map.contains_key(&key) {
        return;
    }
    let Some(bytes) = image_bytes(images, part, &image.rel_id) else {
        return;
    };
    let Some(decoded) = image::try_decode_with_fonts(bytes.as_ref(), store) else {
        return;
    };
    let smask = decoded.alpha.as_ref().map(|_| Ref::new(*next_id + 1));
    let refs = ImageRefs {
        resource: format!("Im{}", image_map.len() + 1),
        id: Ref::new(*next_id),
        smask,
    };
    *next_id += if smask.is_some() { 2 } else { 1 };
    image_map.insert(key, refs.clone());
    image_objects.push((refs, decoded));
}

fn image_bytes<'a>(
    images: &'a BTreeMap<String, Vec<u8>>,
    part: ImagePart<'_>,
    rel_id: &str,
) -> Option<std::borrow::Cow<'a, [u8]>> {
    if let Some(bytes) = images.get(&scoped_image_key(part, rel_id)) {
        return Some(std::borrow::Cow::Borrowed(bytes));
    }
    if let Some(bytes) = images.get(rel_id) {
        return Some(std::borrow::Cow::Borrowed(bytes));
    }
    image::data_url_bytes(rel_id).map(std::borrow::Cow::Owned)
}

fn write_image(pdf: &mut Pdf, refs: &ImageRefs, image: &image::DecodedImage) {
    let rgb = deflate(&image.rgb);
    {
        let mut xobject = pdf.image_xobject(refs.id, &rgb);
        xobject
            .width(image.width)
            .height(image.height)
            .bits_per_component(8);
        xobject.color_space().device_rgb();
        xobject.filter(Filter::FlateDecode);
        if let Some(smask) = refs.smask {
            xobject.s_mask(smask);
        }
    }
    if let (Some(smask), Some(alpha)) = (refs.smask, image.alpha.as_ref()) {
        let alpha = deflate(alpha);
        let mut mask = pdf.image_xobject(smask, &alpha);
        mask.width(image.width)
            .height(image.height)
            .bits_per_component(8);
        mask.color_space().device_gray();
        mask.filter(Filter::FlateDecode);
    }
}

fn paint_decoration(
    content: &mut Content,
    page: &DisplayPage,
    deco: &DecorationPrimitive,
) -> Result<(), PdfError> {
    let x = points(&deco.x, "decoration x")?;
    let y = points(&deco.y, "decoration y")?;
    let w = points(&deco.w, "decoration width")?;
    let h = points(&deco.h, "decoration height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(());
    }
    let Some(color) = parse_color(&deco.color) else {
        return Ok(());
    };
    let page_height = points(&page.height, "page height")?;
    if deco.dashed || deco.dotted {
        let thickness = h.max(0.75);
        content.set_stroke_rgb(color[0], color[1], color[2]);
        content.set_line_width(thickness);
        let unit = if deco.dotted {
            thickness
        } else {
            thickness * 2.0
        };
        content.save_state();
        content.set_dash_pattern([unit, unit], 0.0);
        content.move_to(x, page_height - (y + h / 2.0));
        content.line_to(x + w, page_height - (y + h / 2.0));
        content.stroke();
        content.restore_state();
    } else {
        content.set_fill_rgb(color[0], color[1], color[2]);
        content.rect(x, page_height - y - h, w, h);
        content.fill_nonzero();
    }
    Ok(())
}

fn paint_shape(
    content: &mut Content,
    page: &DisplayPage,
    shape: &ShapePrimitive,
) -> Result<(), PdfError> {
    if shape.geometry_path.is_empty() {
        return Ok(());
    }
    let page_height = points(&page.height, "page height")?;
    content.save_state();
    if let Some(transform) = &shape.transform {
        let rotation = transform
            .rotation
            .as_ref()
            .and_then(Number::as_f64)
            .unwrap_or(0.0);
        if !rotation.is_finite() {
            return Err(PdfError::InvalidNumber("shape rotation"));
        }
        let rotation = rotation as f32;
        if rotation.abs() > f32::EPSILON || transform.flip_h || transform.flip_v {
            let cx = points(&shape.x, "shape x")? + points(&shape.w, "shape width")? / 2.0;
            let cy = page_height
                - points(&shape.y, "shape y")?
                - points(&shape.h, "shape height")? / 2.0;
            let sx = if transform.flip_h { -1.0 } else { 1.0 };
            let sy = if transform.flip_v { -1.0 } else { 1.0 };
            let (sin, cos) = (-rotation).to_radians().sin_cos();
            // translate(cx,cy) * rotate * scale * translate(-cx,-cy)
            let a = cos * sx;
            let b = sin * sx;
            let c = -sin * sy;
            let d = cos * sy;
            let e = cx - a * cx - c * cy;
            let f = cy - b * cx - d * cy;
            content.transform([a, b, c, d, e, f]);
        }
    }

    let mut current = (0.0_f32, 0.0_f32);
    for command in &shape.geometry_path {
        match command {
            ShapePathCommand::Move { x, y } => {
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                content.move_to(x, y);
                current = (x, y);
            }
            ShapePathCommand::Line { x, y } => {
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                content.line_to(x, y);
                current = (x, y);
            }
            ShapePathCommand::Quad { cpx, cpy, x, y } => {
                let cx = points(cpx, "shape control x")?;
                let cy = page_height - points(cpy, "shape control y")?;
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                // pdf-writer has no quadratic operator; convert exactly to a
                // cubic: controls at 2/3 from each endpoint toward the quad
                // control point.
                let c1x = current.0 + 2.0 / 3.0 * (cx - current.0);
                let c1y = current.1 + 2.0 / 3.0 * (cy - current.1);
                let c2x = x + 2.0 / 3.0 * (cx - x);
                let c2y = y + 2.0 / 3.0 * (cy - y);
                content.cubic_to(c1x, c1y, c2x, c2y, x, y);
                current = (x, y);
            }
            ShapePathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                let x = points(x, "shape x")?;
                let y = page_height - points(y, "shape y")?;
                content.cubic_to(
                    points(cp1x, "shape control x")?,
                    page_height - points(cp1y, "shape control y")?,
                    points(cp2x, "shape control x")?,
                    page_height - points(cp2y, "shape control y")?,
                    x,
                    y,
                );
                current = (x, y);
            }
            ShapePathCommand::Close => {
                content.close_path();
            }
        }
    }

    let fill = shape.fill.as_deref().and_then(parse_color);
    let stroke = shape
        .stroke
        .as_ref()
        .and_then(|stroke| parse_color(&stroke.color).map(|color| (color, stroke)));
    if let Some(color) = fill {
        content.set_fill_rgb(color[0], color[1], color[2]);
    }
    if let Some((color, stroke)) = stroke {
        content.set_stroke_rgb(color[0], color[1], color[2]);
        content.set_line_width(points(&stroke.width, "stroke width")?);
    }
    match (fill.is_some(), stroke.is_some()) {
        (true, true) => {
            content.fill_nonzero_and_stroke();
        }
        (true, false) => {
            content.fill_nonzero();
        }
        (false, true) => {
            content.stroke();
        }
        (false, false) => {
            content.end_path();
        }
    }
    content.restore_state();
    Ok(())
}

fn paint_rect(
    content: &mut Content,
    page: &DisplayPage,
    rect: &RectPrimitive,
) -> Result<(), PdfError> {
    let x = points(&rect.x, "rect x")?;
    let y = points(&rect.y, "rect y")?;
    let w = points(&rect.w, "rect width")?;
    let h = points(&rect.h, "rect height")?;
    if w <= 0.0 || h <= 0.0 {
        return Ok(());
    }
    let Some(color) = parse_color(&rect.fill) else {
        return Ok(());
    };
    let page_height = points(&page.height, "page height")?;
    content.set_fill_rgb(color[0], color[1], color[2]);
    content.rect(x, page_height - y - h, w, h);
    content.fill_nonzero();
    Ok(())
}

fn paint_line(
    content: &mut Content,
    page: &DisplayPage,
    line: &LinePrimitive,
) -> Result<(), PdfError> {
    let x1 = points(&line.x1, "line x1")?;
    let y1 = points(&line.y1, "line y1")?;
    let x2 = points(&line.x2, "line x2")?;
    let y2 = points(&line.y2, "line y2")?;
    let width_px = raw_px(&line.stroke_width, "line width")?;
    if width_px <= 0.0 {
        return Ok(());
    }
    let Some(color) = parse_color(&line.color) else {
        return Ok(());
    };
    let secondary = line.secondary_color.as_deref().and_then(parse_color);
    let dash = line
        .dash
        .as_deref()
        .map(|values| {
            values
                .iter()
                .map(|value| raw_px(value, "line dash"))
                .collect::<Result<Vec<_>, _>>()
        })
        .transpose()?;
    let page_height = points(&page.height, "page height")?;
    let segment = BorderSegment {
        x1,
        y1: page_height - y1,
        x2,
        y2: page_height - y2,
    };
    content.save_state();
    paint_border_recipe(
        content,
        segment,
        width_px,
        color,
        line.border_style.unwrap_or(DisplayBorderStyle::Solid),
        secondary,
        dash.as_deref(),
    )?;
    content.restore_state();
    Ok(())
}

/// One straight border segment in PDF points (y up).
#[derive(Clone, Copy)]
struct BorderSegment {
    x1: f32,
    y1: f32,
    x2: f32,
    y2: f32,
}

/// Paint one border segment the way `docx-raster::paint_border_recipe` does:
/// dash styles for plain rules, parallel stroked lanes for compound styles,
/// quadratic waves, and the two-tone lanes of the 3-D styles.
fn paint_border_recipe(
    content: &mut Content,
    segment: BorderSegment,
    width_px: f32,
    color: [f32; 3],
    style: DisplayBorderStyle,
    secondary: Option<[f32; 3]>,
    dash_px: Option<&[f32]>,
) -> Result<(), PdfError> {
    let width_px = width_px.max(0.5);
    match style {
        DisplayBorderStyle::Wave | DisplayBorderStyle::DoubleWave => paint_wave(
            content,
            segment,
            width_px,
            color,
            style == DisplayBorderStyle::DoubleWave,
        ),
        DisplayBorderStyle::Double
        | DisplayBorderStyle::Triple
        | DisplayBorderStyle::ThinThick
        | DisplayBorderStyle::ThickThin => {
            paint_compound_border(content, segment, width_px, color, style);
            Ok(())
        }
        DisplayBorderStyle::Groove
        | DisplayBorderStyle::Ridge
        | DisplayBorderStyle::Inset
        | DisplayBorderStyle::Outset => {
            paint_three_d_border(content, segment, width_px, color, secondary, style);
            Ok(())
        }
        DisplayBorderStyle::Solid
        | DisplayBorderStyle::Dotted
        | DisplayBorderStyle::Dashed
        | DisplayBorderStyle::DashDot
        | DisplayBorderStyle::DashDotDot => {
            let dash = match dash_px {
                Some(values) => values.to_vec(),
                None => style_dash_px(style, width_px).unwrap_or_default(),
            };
            paint_border_lane(content, segment, width_px, color, &dash);
            Ok(())
        }
    }
}

/// The dash arrays the native renderer derives from a border style, in CSS
/// pixels (the display list's unit); empty means a solid rule.
fn style_dash_px(style: DisplayBorderStyle, width_px: f32) -> Option<Vec<f32>> {
    match style {
        DisplayBorderStyle::Dotted => Some(vec![width_px.max(1.0), (width_px * 1.5).max(1.5)]),
        DisplayBorderStyle::Dashed => {
            Some(vec![(width_px * 4.0).max(2.0), (width_px * 2.0).max(2.0)])
        }
        DisplayBorderStyle::DashDot => Some(vec![
            width_px * 4.0,
            width_px * 1.5,
            width_px,
            width_px * 1.5,
        ]),
        DisplayBorderStyle::DashDotDot => Some(vec![
            width_px * 4.0,
            width_px * 1.5,
            width_px,
            width_px * 1.5,
            width_px,
            width_px * 1.5,
        ]),
        _ => None,
    }
}

/// Convert a dash pattern from CSS pixels to PDF points. PDF requires an even,
/// non-negative array with a positive period; anything else means solid, which
/// is also how the native renderer treats an invalid pattern.
fn pdf_dash(values: &[f32]) -> Option<Vec<f32>> {
    if values.len() < 2 || values.len() % 2 != 0 {
        return None;
    }
    let mut total = 0.0_f32;
    for value in values {
        if !value.is_finite() || *value < 0.0 {
            return None;
        }
        total += value;
    }
    if !(total > 0.0) || !total.is_finite() {
        return None;
    }
    let converted: Vec<f32> = values.iter().map(|value| value * PX_TO_PT).collect();
    converted
        .iter()
        .all(|value| value.is_finite() && *value >= 0.0)
        .then_some(converted)
}

/// Set the stroke parameters shared by every border lane. The dash pattern is
/// always written (an empty array for solid) so a dashed neighbour cannot leak
/// its pattern into this stroke.
fn set_border_stroke_state(content: &mut Content, width_px: f32, color: [f32; 3], dash_px: &[f32]) {
    content.set_stroke_rgb(color[0], color[1], color[2]);
    content.set_line_width(width_px * PX_TO_PT);
    content.set_line_cap(LineCapStyle::ButtCap);
    content.set_line_join(LineJoinStyle::MiterJoin);
    content.set_miter_limit(4.0);
    match pdf_dash(dash_px) {
        Some(dash) => {
            content.set_dash_pattern(dash, 0.0);
        }
        None => {
            content.set_dash_pattern([0.0_f32; 0], 0.0);
        }
    }
}

fn paint_border_lane(
    content: &mut Content,
    segment: BorderSegment,
    width_px: f32,
    color: [f32; 3],
    dash_px: &[f32],
) {
    set_border_stroke_state(content, width_px, color, dash_px);
    content.move_to(segment.x1, segment.y1);
    content.line_to(segment.x2, segment.y2);
    content.stroke();
}

/// A normal pointing toward the side the native renderer offsets its lanes to.
/// The display list is y-down while PDF user space is y-up, so the vertical
/// sign is inverted; otherwise `thinThick`, groove and ridge would render
/// mirrored.
fn border_normal(segment: BorderSegment) -> (f32, f32) {
    let dx = segment.x2 - segment.x1;
    let dy = segment.y2 - segment.y1;
    let length = dx.hypot(dy).max(f32::EPSILON);
    (dy / length, -dx / length)
}

fn offset_border_segment(
    segment: BorderSegment,
    normal: (f32, f32),
    offset_px: f32,
) -> BorderSegment {
    let dx = normal.0 * offset_px * PX_TO_PT;
    let dy = normal.1 * offset_px * PX_TO_PT;
    BorderSegment {
        x1: segment.x1 + dx,
        y1: segment.y1 + dy,
        x2: segment.x2 + dx,
        y2: segment.y2 + dy,
    }
}

fn paint_compound_border(
    content: &mut Content,
    segment: BorderSegment,
    width_px: f32,
    color: [f32; 3],
    style: DisplayBorderStyle,
) {
    let normal = border_normal(segment);
    let strokes = match style {
        DisplayBorderStyle::Triple => vec![
            (-width_px, (width_px / 3.0).max(0.5)),
            (0.0, (width_px / 3.0).max(0.5)),
            (width_px, (width_px / 3.0).max(0.5)),
        ],
        DisplayBorderStyle::ThinThick => vec![
            (-width_px * 0.45, (width_px * 0.25).max(0.5)),
            (width_px * 0.3, (width_px * 0.55).max(0.75)),
        ],
        DisplayBorderStyle::ThickThin => vec![
            (-width_px * 0.3, (width_px * 0.55).max(0.75)),
            (width_px * 0.45, (width_px * 0.25).max(0.5)),
        ],
        DisplayBorderStyle::Double => vec![
            (-width_px / 2.0, (width_px / 3.0).max(0.5)),
            (width_px / 2.0, (width_px / 3.0).max(0.5)),
        ],
        _ => Vec::new(),
    };
    for (offset, lane_width) in strokes {
        paint_border_lane(
            content,
            offset_border_segment(segment, normal, offset),
            lane_width,
            color,
            &[],
        );
    }
}

fn paint_three_d_border(
    content: &mut Content,
    segment: BorderSegment,
    width_px: f32,
    color: [f32; 3],
    secondary: Option<[f32; 3]>,
    style: DisplayBorderStyle,
) {
    let light = secondary.unwrap_or_else(|| mix_rgb(color, [1.0, 1.0, 1.0], 0.55));
    let dark = mix_rgb(color, [0.0, 0.0, 0.0], 0.45);
    let reverse = matches!(
        style,
        DisplayBorderStyle::Ridge | DisplayBorderStyle::Outset
    );
    let (first, second) = if reverse {
        (light, dark)
    } else {
        (dark, light)
    };
    let normal = border_normal(segment);
    let offset = width_px / 4.0;
    for (lane, lane_color) in [(-offset, first), (offset, second)] {
        paint_border_lane(
            content,
            offset_border_segment(segment, normal, lane),
            (width_px / 2.0).max(0.5),
            lane_color,
            &[],
        );
    }
}

/// Cubic segments one wave may expand into before export refuses the border.
///
/// PDF 1.6 caps a page side at 14,400 units, so even the diagonal of the
/// largest legal page is ~20,365 pt. At the smallest half wavelength (1.5 pt)
/// that is ~13,577 cubics per lane, or ~27,154 for a DoubleWave. 65,536 keeps
/// generous headroom for content that overhangs the page while bounding one
/// wave to roughly 4 MB of content stream (a cubic is about 60 bytes), far
/// below the 2 GiB a PDF stream can address and the native rasterizer's page
/// budget.
const MAX_WAVE_CUBICS: u64 = 65_536;

/// Cubics one lane of a wave expands into: one per half wavelength, rounding
/// a partial half wave up, plus one for the drift of walking the lane in f32.
/// Computed in f64 so a long segment cannot overflow before the budget check.
fn wave_cubics(length: f32, wavelength: f32) -> u64 {
    let half = f64::from(wavelength) / 2.0;
    let cubics = (f64::from(length) / half).ceil();
    if !cubics.is_finite() || cubics <= 0.0 {
        // Unusable geometry, over any budget the caller applies.
        return u64::MAX;
    }
    (cubics as u64).saturating_add(1)
}

fn paint_wave(
    content: &mut Content,
    segment: BorderSegment,
    width_px: f32,
    color: [f32; 3],
    double: bool,
) -> Result<(), PdfError> {
    let dx = segment.x2 - segment.x1;
    let dy = segment.y2 - segment.y1;
    let length = dx.hypot(dy);
    if length == 0.0 {
        return Ok(());
    }
    if !length.is_finite() {
        return Err(PdfError::InvalidNumber("border wave length"));
    }
    let tangent_x = dx / length;
    let tangent_y = dy / length;
    let normal_x = tangent_y;
    let normal_y = -tangent_x;
    let wavelength = (width_px * 4.0).max(4.0) * PX_TO_PT;
    let amplitude = width_px.max(1.0) * PX_TO_PT;
    if !wavelength.is_finite() || !(wavelength > 0.0) || !amplitude.is_finite() {
        return Err(PdfError::InvalidNumber("border wave geometry"));
    }
    // Divide only after the wavelength is finite and nonzero, so the walk
    // below cannot lose its step and spin until the float stalls.
    let half_wavelength = wavelength / 2.0;
    if !half_wavelength.is_finite() || !(half_wavelength > 0.0) {
        return Err(PdfError::InvalidNumber("border wave geometry"));
    }
    let lanes: &[f32] = if double { &[-1.0, 1.0] } else { &[0.0] };
    // Refuse an over-budget wave before the first operator is written: the
    // walk below then has a point count bounded by the validated budget.
    let lane_cubics = wave_cubics(length, wavelength);
    let cubics = lane_cubics.saturating_mul(lanes.len() as u64);
    if cubics > MAX_WAVE_CUBICS {
        return Err(PdfError::Unsupported(
            "border wave exceeds the cubic budget",
        ));
    }
    let lane_width = (width_px / if double { 2.0 } else { 1.0 }).max(0.5);
    for lane_factor in lanes {
        let lane = lane_factor * amplitude;
        content.move_to(segment.x1 + normal_x * lane, segment.y1 + normal_y * lane);
        let mut distance = 0.0_f32;
        let mut sign = 1.0_f32;
        for _ in 0..lane_cubics {
            if !(distance < length) {
                break;
            }
            let next = length.min(distance + half_wavelength);
            if !(next > distance) {
                break;
            }
            let midpoint = (distance + next) / 2.0;
            let start_x = segment.x1 + tangent_x * distance + normal_x * lane;
            let start_y = segment.y1 + tangent_y * distance + normal_y * lane;
            let control_x =
                segment.x1 + tangent_x * midpoint + normal_x * (lane + amplitude * sign);
            let control_y =
                segment.y1 + tangent_y * midpoint + normal_y * (lane + amplitude * sign);
            let end_x = segment.x1 + tangent_x * next + normal_x * lane;
            let end_y = segment.y1 + tangent_y * next + normal_y * lane;
            // pdf-writer has no quadratic operator; convert exactly to a cubic
            // with controls 2/3 of the way from each endpoint to the quad
            // control point.
            let c1x = start_x + (control_x - start_x) * 2.0 / 3.0;
            let c1y = start_y + (control_y - start_y) * 2.0 / 3.0;
            let c2x = end_x + (control_x - end_x) * 2.0 / 3.0;
            let c2y = end_y + (control_y - end_y) * 2.0 / 3.0;
            content.cubic_to(c1x, c1y, c2x, c2y, end_x, end_y);
            distance = next;
            sign = -sign;
        }
        set_border_stroke_state(content, lane_width, color, &[]);
        content.stroke();
    }
    Ok(())
}

fn mix_rgb(left: [f32; 3], right: [f32; 3], amount: f32) -> [f32; 3] {
    let amount = amount.clamp(0.0, 1.0);
    [
        left[0] + (right[0] - left[0]) * amount,
        left[1] + (right[1] - left[1]) * amount,
        left[2] + (right[2] - left[2]) * amount,
    ]
}

fn points(value: &Number, what: &'static str) -> Result<f32, PdfError> {
    let value = value.as_f64().ok_or(PdfError::InvalidNumber(what))?;
    let value = (value * f64::from(PX_TO_PT)) as f32;
    if !value.is_finite() {
        return Err(PdfError::InvalidNumber(what));
    }
    Ok(value)
}

/// Parse a display-list length without the CSS-to-point conversion, for
/// quantities the border recipe derives from the CSS-pixel stroke width.
fn raw_px(value: &Number, what: &'static str) -> Result<f32, PdfError> {
    let value = value.as_f64().ok_or(PdfError::InvalidNumber(what))? as f32;
    if !value.is_finite() {
        return Err(PdfError::InvalidNumber(what));
    }
    Ok(value)
}

/// Best-effort CSS font-size parser for the fallback `text` primitive.
fn font_size_px(css: &str) -> f32 {
    css.split_whitespace()
        .find_map(|part| part.strip_suffix("px"))
        .and_then(|value| value.parse::<f32>().ok())
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(12.0)
}

/// Parse `#rgb`, `#rrggbb`, `rgb(...)` or `rgba(...)`; alpha is ignored.
fn parse_color(value: &str) -> Option<[f32; 3]> {
    let value = value.trim();
    if let Some(hex) = value.strip_prefix('#') {
        return match hex.len() {
            3 => Some([
                hex_channel(hex.get(0..1)?)?,
                hex_channel(hex.get(1..2)?)?,
                hex_channel(hex.get(2..3)?)?,
            ]),
            6 => Some([
                hex_channel(hex.get(0..2)?)?,
                hex_channel(hex.get(2..4)?)?,
                hex_channel(hex.get(4..6)?)?,
            ]),
            8 => Some([
                hex_channel(hex.get(0..2)?)?,
                hex_channel(hex.get(2..4)?)?,
                hex_channel(hex.get(4..6)?)?,
            ]),
            _ => None,
        };
    }
    let inner = value
        .strip_prefix("rgb(")
        .or_else(|| value.strip_prefix("rgba("))
        .and_then(|value| value.strip_suffix(')'))?;
    let mut parts = inner.split(',').map(str::trim);
    let r = channel(parts.next()?)?;
    let g = channel(parts.next()?)?;
    let b = channel(parts.next()?)?;
    Some([r, g, b])
}

fn hex_channel(value: &str) -> Option<f32> {
    let value = u8::from_str_radix(value, 16).ok()?;
    Some(f32::from(value) / 255.0)
}

fn channel(value: &str) -> Option<f32> {
    let value = value.parse::<f32>().ok()?;
    value.is_finite().then(|| (value / 255.0).clamp(0.0, 1.0))
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::display_list::{
        DecoKind, DocAttrs, GlyphRunPrimitive, HfKind, HfRegion, ImagePrimitive,
        PageBorderPrimitive, PageBorderSide, PlacedGlyph,
    };

    fn number(value: f64) -> Number {
        Number::from_f64(value).unwrap()
    }

    /// One page carrying `primitives`, with `watermark_primitive_count` and a
    /// single header (footer stays empty).
    fn one_page(
        primitives: Vec<Primitive>,
        watermark_primitive_count: Option<usize>,
        header: Option<HfRegion>,
    ) -> DisplayList {
        DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
                page_index: 0,
                width: Number::from(816),
                height: Number::from(1056),
                content_bounds: None,
                column_bounds: Vec::new(),
                section_id: None,
                section_index: None,
                section_page_index: None,
                section_page_number: None,
                page_label: None,
                primitives,
                watermark_primitive_count,
                background: None,
                page_borders: Vec::new(),
                header,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        }
    }

    fn rect(x: f64, y: f64, w: f64, h: f64, fill: &str) -> Primitive {
        Primitive::Rect(RectPrimitive {
            x: number(x),
            y: number(y),
            w: number(w),
            h: number(h),
            fill: fill.to_owned(),
            attrs: DocAttrs::default(),
        })
    }

    fn line(x1: f64, y1: f64, x2: f64, y2: f64, width: f64, color: &str) -> Primitive {
        Primitive::Line(LinePrimitive {
            x1: number(x1),
            y1: number(y1),
            x2: number(x2),
            y2: number(y2),
            stroke_width: number(width),
            color: color.to_owned(),
            dash: None,
            role: None,
            border_style: None,
            secondary_color: None,
            opacity: None,
            border_owner: None,
            attrs: DocAttrs::default(),
        })
    }

    /// A border line with an explicit style and optional dash override (CSS
    /// pixels), the way `docx-layout` emits table and paragraph borders.
    fn border_line(
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        width: f64,
        color: &str,
        style: DisplayBorderStyle,
        dash: Option<Vec<f64>>,
    ) -> Primitive {
        let mut primitive = line(x1, y1, x2, y2, width, color);
        if let Primitive::Line(line) = &mut primitive {
            line.border_style = Some(style);
            line.dash = dash.map(|values| values.into_iter().map(number).collect());
        }
        primitive
    }

    /// A shape with nonempty geometry and neither fill nor stroke.
    fn paintless_shape(x: f64, y: f64) -> Primitive {
        Primitive::Shape(ShapePrimitive {
            x: number(x),
            y: number(y),
            w: number(100.0),
            h: number(50.0),
            geometry_path: vec![
                ShapePathCommand::Move {
                    x: number(10.0),
                    y: number(10.0),
                },
                ShapePathCommand::Line {
                    x: number(90.0),
                    y: number(10.0),
                },
                ShapePathCommand::Line {
                    x: number(90.0),
                    y: number(40.0),
                },
                ShapePathCommand::Close,
            ],
            fill: None,
            stroke: None,
            transform: None,
            decorative: false,
            attrs: DocAttrs::default(),
        })
    }

    fn text(x: f64, y: f64, value: &str) -> Primitive {
        Primitive::Text(TextRunPrimitive {
            text: value.to_owned(),
            paint_text: None,
            x: Number::from_f64(x).unwrap(),
            baseline_y: Number::from_f64(y).unwrap(),
            width: Number::from(100),
            paint_clip: None,
            font: "12px Helvetica".to_owned(),
            scripts: None,
            color: "#112233".to_owned(),
            letter_spacing: None,
            word_spacing: None,
            rtl: None,
            opacity: None,
            rotation_deg: None,
            horizontal_scale: None,
            substitute_scale: None,
            all_caps: false,
            small_caps: false,
            hidden: false,
            text_shadow: None,
            text_outline: false,
            emphasis_mark: None,
            text_effect: None,
            attrs: DocAttrs::default(),
        })
    }

    #[test]
    fn writes_images_with_soft_masks() {
        let mut png = std::io::Cursor::new(Vec::new());
        let source = ::image::RgbaImage::from_pixel(2, 2, ::image::Rgba([200, 20, 10, 128]));
        ::image::DynamicImage::ImageRgba8(source)
            .write_to(&mut png, ::image::ImageFormat::Png)
            .expect("encode png");
        let mut images = BTreeMap::new();
        images.insert("rId7".to_owned(), png.into_inner());

        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
                page_index: 0,
                width: Number::from(816),
                height: Number::from(1056),
                content_bounds: None,
                column_bounds: Vec::new(),
                section_id: None,
                section_index: None,
                section_page_index: None,
                section_page_number: None,
                page_label: None,
                primitives: vec![Primitive::Image(ImagePrimitive {
                    rel_id: "rId7".to_owned(),
                    x: Number::from(72),
                    y: Number::from(72),
                    w: Number::from(48),
                    h: Number::from(48),
                    rotation_deg: None,
                    opacity: Some(Number::from_f64(0.5).unwrap()),
                    filter: None,
                    decorative: false,
                    crop: None,
                    alt_text: None,
                    relationship_id: None,
                    attrs: DocAttrs::default(),
                })],
                watermark_primitive_count: None,
                background: None,
                page_borders: Vec::new(),
                header: None,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        };

        let mut store = FontStore::default();
        let bytes =
            export_pdf_with_fonts_and_images(&display_list, &mut store, &images).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.contains("/Subtype /Image"));
        assert!(pdf.contains("/SMask"));
        assert!(pdf.contains("/DeviceRGB"));
        assert!(pdf.contains("/DeviceGray"));
        assert!(pdf.contains("/ExtGState"));
        assert!(pdf.contains("/ca 0.5"));
    }

    #[test]
    fn zero_opacity_emits_a_transparent_graphics_state() {
        let mut page = text(72.0, 96.0, "transparent");
        if let Primitive::Text(run) = &mut page {
            run.opacity = Some(Number::from_f64(0.0).unwrap());
        }
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
                page_index: 0,
                width: Number::from(816),
                height: Number::from(1056),
                content_bounds: None,
                column_bounds: Vec::new(),
                section_id: None,
                section_index: None,
                section_page_index: None,
                section_page_number: None,
                page_label: None,
                primitives: vec![page],
                watermark_primitive_count: None,
                background: None,
                page_borders: Vec::new(),
                header: None,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        };
        let mut store = FontStore::default();
        let bytes = export_pdf_with_fonts_and_images(&display_list, &mut store, &BTreeMap::new())
            .expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.contains("/ca 0"));
    }

    #[test]
    fn embeds_and_subsets_glyph_runs() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../packages/fonts/assets/Roboto-Bold.ttf");
        let data = std::fs::read(path).expect("font fixture");
        let mut store = FontStore::default();
        let font_id = store.register(data.clone()).expect("register font");

        let face = ttf_parser::Face::parse(&data, 0).expect("parse font");
        let upem = f64::from(face.units_per_em());
        let size = 16.0;
        let text = "Hello";
        let mut x = 72.0;
        let mut glyphs = Vec::new();
        for (cluster, ch) in text.char_indices() {
            let glyph = face.glyph_index(ch).expect("glyph").0;
            let advance = f64::from(face.glyph_hor_advance(ttf_parser::GlyphId(glyph)).unwrap())
                * size
                / upem;
            glyphs.push(PlacedGlyph {
                id: u32::from(glyph),
                x,
                y: 96.0,
                cluster: cluster as u32,
                advance,
                logical_order: None,
                bidi_level: None,
            });
            x += advance;
        }

        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
                page_index: 0,
                width: Number::from(816),
                height: Number::from(1056),
                content_bounds: None,
                column_bounds: Vec::new(),
                section_id: None,
                section_index: None,
                section_page_index: None,
                section_page_number: None,
                page_label: None,
                primitives: vec![Primitive::GlyphRun(GlyphRunPrimitive {
                    font_id: font_id.to_u32(),
                    size,
                    color: "#112233".to_owned(),
                    text: text.to_owned(),
                    paint_text: None,
                    glyphs,
                    echo: None,
                    paint_clip: None,
                    word_spacing: None,
                    rtl: None,
                    opacity: None,
                    rotation_deg: None,
                    horizontal_scale: None,
                    all_caps: false,
                    small_caps: false,
                    hidden: false,
                    text_shadow: None,
                    text_outline: false,
                    emphasis_mark: None,
                    text_effect: None,
                    attrs: DocAttrs::default(),
                })],
                watermark_primitive_count: None,
                background: None,
                page_borders: Vec::new(),
                header: None,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        };

        let bytes = export_pdf_with_fonts(&display_list, &store).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.contains("/FontFile2"));
        assert!(pdf.contains("/Identity-H"));
        assert!(pdf.contains("/ToUnicode"));
        assert!(pdf.contains("/BaseFont /AAAAAA+BO"));
    }

    fn inflate_first_stream(pdf: &[u8]) -> String {
        use std::io::Read;
        let start = pdf
            .windows(b"stream\n".len())
            .position(|window| window == b"stream\n")
            .expect("stream")
            + b"stream\n".len();
        let end = pdf[start..]
            .windows(b"endstream".len())
            .position(|window| window == b"endstream")
            .expect("endstream")
            + start;
        let mut decoded = String::new();
        flate2::read::ZlibDecoder::new(&pdf[start..end])
            .read_to_string(&mut decoded)
            .expect("inflate content");
        decoded
    }

    fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
        haystack
            .windows(needle.len())
            .position(|window| window == needle)
    }

    /// Decompress the content stream the page dictionary's `/Contents` names.
    ///
    /// Unlike [`inflate_first_stream`], this finds the page stream even when an
    /// image or font stream precedes it.
    fn inflate_page_content(pdf: &[u8]) -> String {
        use std::io::Read;
        let contents = find_bytes(pdf, b"/Contents ").expect("/Contents");
        let id_start = contents + b"/Contents ".len();
        let id_len = pdf[id_start..]
            .iter()
            .take_while(|byte| byte.is_ascii_digit())
            .count();
        assert!(id_len > 0, "content stream reference");
        let id = std::str::from_utf8(&pdf[id_start..id_start + id_len]).expect("content id");
        let object = format!("\n{id} 0 obj");
        let object_start = find_bytes(pdf, object.as_bytes()).expect("content object");
        let stream_start = object_start
            + find_bytes(&pdf[object_start..], b"stream\n").expect("content stream")
            + b"stream\n".len();
        let stream_end = stream_start
            + find_bytes(&pdf[stream_start..], b"endstream").expect("content stream end");
        let mut decoded = String::new();
        flate2::read::ZlibDecoder::new(&pdf[stream_start..stream_end])
            .read_to_string(&mut decoded)
            .expect("inflate content");
        decoded
    }

    /// `Tm` matrices in a content stream, in painting order.
    fn tm_matrices(stream: &str) -> Vec<[f64; 6]> {
        stream
            .lines()
            .filter_map(|line| {
                let tokens: Vec<&str> = line.split_whitespace().collect();
                if tokens.last() != Some(&"Tm") {
                    return None;
                }
                let start = tokens.len().checked_sub(7)?;
                let mut matrix = [0.0; 6];
                for (slot, token) in matrix.iter_mut().zip(&tokens[start..start + 6]) {
                    *slot = token.parse().ok()?;
                }
                Some(matrix)
            })
            .collect()
    }

    /// The four numbers of the first `/Rect [...]` in the PDF objects.
    fn first_annotation_rect(pdf: &str) -> Option<[f64; 4]> {
        let start = pdf.find("/Rect [")? + "/Rect [".len();
        let end = start + pdf[start..].find(']')?;
        let values: Vec<f64> = pdf[start..end]
            .split_whitespace()
            .map(str::parse)
            .collect::<Result<_, _>>()
            .ok()?;
        (values.len() == 4).then(|| [values[0], values[1], values[2], values[3]])
    }

    /// Value of one `/ca` or `/CA` entry in the ExtGState dictionaries.
    fn ext_g_state_alpha(pdf: &str, key: &str) -> Option<f32> {
        let start = pdf.find(key)? + key.len();
        let rest = pdf[start..].trim_start();
        let end = rest
            .find(|character: char| {
                !(character.is_ascii_digit() || character == '.' || character == '-')
            })
            .unwrap_or(rest.len());
        rest[..end].parse().ok()
    }

    #[derive(Debug, Clone, Default)]
    struct GState {
        fill: Option<[f32; 3]>,
        stroke: Option<[f32; 3]>,
        dash: Vec<f32>,
    }

    enum Operand {
        Num(f32),
        Array(Vec<f32>),
    }

    /// One path-painting operation together with the graphics state and the
    /// path it actually painted.
    #[derive(Debug, Clone, PartialEq)]
    struct PaintOp {
        /// `n`, `S`/`s`, `f`/`F`/`f*`, `B`/`B*`/`b`/`b*`.
        op: String,
        fill: Option<[f32; 3]>,
        stroke: Option<[f32; 3]>,
        dashed: bool,
        /// The dash array in effect, in points; empty means solid.
        dash: Vec<f32>,
        /// Endpoints of `m`/`l` operands seen since the previous paint.
        points: Vec<(f32, f32)>,
        /// Subpaths started since the previous paint; a leaked shape path
        /// shows up as an extra subpath on the next fill.
        subpaths: usize,
    }

    fn operands_rgb(operands: &[Operand]) -> Option<[f32; 3]> {
        match operands {
            [Operand::Num(r), Operand::Num(g), Operand::Num(b)] => Some([*r, *g, *b]),
            _ => None,
        }
    }

    /// Interpret the path, color, dash and `q`/`Q` operators of a page content
    /// stream and return its painting operations in order.
    ///
    /// Text and image operators are ignored: their operands never change the
    /// state this tracks (colors are RGB, dash arrays are the only arrays).
    fn paint_ops(content: &str) -> Vec<PaintOp> {
        let normalized = content.replace('[', " [ ").replace(']', " ] ");
        let mut tokens = normalized.split_whitespace();
        let mut operands: Vec<Operand> = Vec::new();
        let mut state = GState::default();
        let mut stack: Vec<GState> = Vec::new();
        let mut subpaths = 0usize;
        let mut path_points: Vec<(f32, f32)> = Vec::new();
        let mut ops = Vec::new();
        while let Some(token) = tokens.next() {
            if token == "[" {
                let mut values = Vec::new();
                for item in tokens.by_ref() {
                    if item == "]" {
                        break;
                    }
                    if let Ok(value) = item.parse::<f32>() {
                        values.push(value);
                    }
                }
                operands.push(Operand::Array(values));
                continue;
            }
            if let Ok(value) = token.parse::<f32>() {
                operands.push(Operand::Num(value));
                continue;
            }
            match token {
                "q" => stack.push(state.clone()),
                "Q" => state = stack.pop().unwrap_or_default(),
                "rg" => state.fill = operands_rgb(&operands),
                "RG" => state.stroke = operands_rgb(&operands),
                "d" => {
                    state.dash = match operands.first() {
                        Some(Operand::Array(values)) => values.clone(),
                        _ => Vec::new(),
                    };
                }
                "m" | "l" => {
                    if let (Some(Operand::Num(x)), Some(Operand::Num(y))) =
                        (operands.first(), operands.get(1))
                    {
                        path_points.push((*x, *y));
                    }
                    subpaths += 1;
                }
                "re" => subpaths += 1,
                "n" | "S" | "s" | "f" | "F" | "f*" | "B" | "B*" | "b" | "b*" => {
                    ops.push(PaintOp {
                        op: token.to_owned(),
                        fill: state.fill,
                        stroke: state.stroke,
                        dashed: state.dash.iter().any(|value| *value > 0.0),
                        dash: state.dash.clone(),
                        points: std::mem::take(&mut path_points),
                        subpaths,
                    });
                    subpaths = 0;
                }
                _ => {}
            }
            operands.clear();
        }
        ops
    }

    #[test]
    fn paintless_shape_ends_its_path_before_the_next_fill() {
        let display_list = one_page(
            vec![
                paintless_shape(72.0, 120.0),
                rect(100.0, 120.0, 40.0, 30.0, "#ff0000"),
            ],
            None,
            None,
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let first = ops.first().expect("a painting operation");
        assert_eq!(
            first.op, "n",
            "a paint-less shape must end its path before the next primitive"
        );
        let fill = ops
            .iter()
            .find(|op| op.op == "f")
            .expect("rect fill operation");
        assert_eq!(fill.fill, Some([1.0, 0.0, 0.0]));
        assert_eq!(
            fill.subpaths, 1,
            "the rect fill must paint only the rect, not the leaked shape path"
        );
    }

    #[test]
    fn watermark_paints_below_header_and_body_paints_above() {
        let header = HfRegion {
            r_id: "rIdHeader".to_owned(),
            kind: HfKind::Header,
            y: Number::from(0),
            height: Number::from(72),
            primitives: vec![rect(72.0, 36.0, 100.0, 10.0, "#00ff00")],
        };
        let display_list = one_page(
            vec![
                rect(0.0, 200.0, 300.0, 300.0, "#ff0000"),
                rect(72.0, 400.0, 100.0, 10.0, "#0000ff"),
            ],
            Some(1),
            Some(header),
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let fills: Vec<Option<[f32; 3]>> = ops
            .iter()
            .filter(|op| op.op == "f")
            .map(|op| op.fill)
            .collect();
        assert_eq!(
            fills,
            vec![
                Some([1.0, 0.0, 0.0]),
                Some([0.0, 1.0, 0.0]),
                Some([0.0, 0.0, 1.0]),
            ],
            "watermark, header, then body must paint in that order"
        );
    }

    #[test]
    fn dashed_rule_does_not_leak_its_dash_into_a_solid_line() {
        let decoration = Primitive::Decoration(DecorationPrimitive {
            deco: DecoKind::Underline,
            x: Number::from(72),
            y: Number::from(200),
            w: Number::from(100),
            h: Number::from(2),
            color: "#000000".to_owned(),
            dashed: true,
            dotted: false,
            attrs: DocAttrs::default(),
        });
        let display_list = one_page(
            vec![decoration, line(72.0, 240.0, 172.0, 240.0, 1.0, "#ff0000")],
            None,
            None,
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let strokes: Vec<&PaintOp> = ops
            .iter()
            .filter(|op| op.op == "S" || op.op == "s")
            .collect();
        assert_eq!(strokes.len(), 2, "one dashed rule and one solid line");
        assert!(
            strokes[0].dashed,
            "the dashed decoration must set a dash pattern"
        );
        assert!(
            !strokes[1].dashed,
            "the solid line must not inherit the decoration dash pattern"
        );
        assert_eq!(strokes[1].stroke, Some([1.0, 0.0, 0.0]));
    }

    #[test]
    fn dotted_and_dashed_lines_emit_their_dash_patterns() {
        let display_list = one_page(
            vec![
                // Dotted without an explicit array: derive it from the style,
                // like the native renderer's `border_dash`.
                border_line(
                    72.0,
                    200.0,
                    172.0,
                    200.0,
                    2.0,
                    "#ff0000",
                    DisplayBorderStyle::Dotted,
                    None,
                ),
                // An explicit override always wins over the style, in CSS px.
                border_line(
                    72.0,
                    240.0,
                    172.0,
                    240.0,
                    1.0,
                    "#0000ff",
                    DisplayBorderStyle::Dashed,
                    Some(vec![12.0, 6.0]),
                ),
                line(72.0, 280.0, 172.0, 280.0, 1.0, "#00ff00"),
            ],
            None,
            None,
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let strokes: Vec<&PaintOp> = ops
            .iter()
            .filter(|op| op.op == "S" || op.op == "s")
            .collect();
        assert_eq!(strokes.len(), 3, "two dashed lines and one solid line");
        assert!(strokes[0].dashed, "a dotted line must set a dash pattern");
        assert_eq!(
            strokes[0].dash,
            vec![1.5, 2.25],
            "2px dotted dashes convert to 1.5/2.25pt"
        );
        assert_eq!(strokes[0].stroke, Some([1.0, 0.0, 0.0]));
        assert!(strokes[1].dashed, "a dashed line must set a dash pattern");
        assert_eq!(
            strokes[1].dash,
            vec![9.0, 4.5],
            "the dash override is used and converted from CSS px to points"
        );
        assert_eq!(strokes[1].stroke, Some([0.0, 0.0, 1.0]));
        assert!(
            !strokes[2].dashed,
            "the solid line must not inherit a neighbour's dash pattern"
        );
        assert!(strokes[2].dash.is_empty());
        assert_eq!(strokes[2].stroke, Some([0.0, 1.0, 0.0]));
    }

    #[test]
    fn double_border_paints_two_parallel_solid_lanes() {
        let display_list = one_page(
            vec![border_line(
                72.0,
                200.0,
                172.0,
                200.0,
                4.0,
                "#ff0000",
                DisplayBorderStyle::Double,
                None,
            )],
            None,
            None,
        );
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let strokes: Vec<&PaintOp> = ops
            .iter()
            .filter(|op| op.op == "S" || op.op == "s")
            .collect();
        assert_eq!(strokes.len(), 2, "a double border paints two lanes");
        for stroke in &strokes {
            assert!(!stroke.dashed);
            assert!(stroke.dash.is_empty());
            assert_eq!(stroke.stroke, Some([1.0, 0.0, 0.0]));
            assert_eq!(stroke.points.len(), 2, "each lane is one straight segment");
        }
        // The page is 1056px (792pt) tall, so y=200px sits at 642pt; a 4px
        // border offsets each lane by 2px = 1.5pt along the normal.
        let ys: Vec<f32> = strokes.iter().map(|stroke| stroke.points[0].1).collect();
        assert!((ys[0] - 643.5).abs() < 0.01, "first lane y: {ys:?}");
        assert!((ys[1] - 640.5).abs() < 0.01, "second lane y: {ys:?}");
        for stroke in &strokes {
            assert_eq!(stroke.points[0].0, 54.0);
            assert_eq!(stroke.points[1].0, 129.0);
        }
    }

    /// Cubics a 300pt wave writes at the minimum 3pt wavelength, and the size
    /// of the content stream carrying them.
    fn wave_cubics_and_bytes(double: bool) -> (usize, usize) {
        let mut content = Content::new();
        paint_wave(
            &mut content,
            BorderSegment {
                x1: 0.0,
                y1: 0.0,
                x2: 300.0,
                y2: 0.0,
            },
            1.0,
            [0.0, 0.0, 0.0],
            double,
        )
        .expect("a short wave paints");
        let bytes = content.finish().into_vec();
        let cubics = String::from_utf8_lossy(&bytes)
            .lines()
            .filter(|line| line.split_whitespace().last() == Some("c"))
            .count();
        (cubics, bytes.len())
    }

    #[test]
    fn short_wave_emits_cubics_and_double_wave_doubles_the_lanes() {
        let (single, single_bytes) = wave_cubics_and_bytes(false);
        let (double, double_bytes) = wave_cubics_and_bytes(true);
        assert_eq!(single, 200, "300pt at a 3pt wavelength is 200 half waves");
        assert_eq!(double, single * 2, "DoubleWave must paint two lanes");
        assert!(
            single_bytes < 64 * 1024 && double_bytes < 64 * 1024,
            "short waves must stay small: {single_bytes}/{double_bytes} bytes"
        );
    }

    #[test]
    fn over_budget_wave_is_rejected_before_emitting_any_operator() {
        // 4,000,000 CSS px is 3,000,000pt: finite, but at the 3pt minimum
        // wavelength one lane alone would need 2,000,000 cubics.
        let mut content = Content::new();
        let result = paint_wave(
            &mut content,
            BorderSegment {
                x1: 0.0,
                y1: 0.0,
                x2: 3_000_000.0,
                y2: 0.0,
            },
            1.0,
            [0.0, 0.0, 0.0],
            true,
        );
        assert!(
            matches!(result, Err(PdfError::Unsupported(_))),
            "a 3,000,000pt wave must be refused: {result:?}"
        );
        let bytes = content.finish().into_vec();
        assert!(
            bytes.is_empty(),
            "the budget check must run before any path operator: {} bytes",
            bytes.len()
        );
    }

    #[test]
    fn wave_rejects_non_finite_geometry_before_dividing() {
        for (x2, width, expected) in [
            (f32::INFINITY, 1.0, "an infinite length"),
            (f32::NAN, 1.0, "a NaN length"),
            (100.0, f32::INFINITY, "an infinite width"),
            (100.0, f32::MAX, "a width whose wavelength overflows"),
        ] {
            let mut content = Content::new();
            let result = paint_wave(
                &mut content,
                BorderSegment {
                    x1: 0.0,
                    y1: 0.0,
                    x2,
                    y2: 0.0,
                },
                width,
                [0.0, 0.0, 0.0],
                false,
            );
            assert!(
                matches!(result, Err(PdfError::InvalidNumber(_))),
                "{expected} must be refused: {result:?}"
            );
            assert!(
                content.finish().into_vec().is_empty(),
                "{expected} must be refused before any operator"
            );
        }
    }

    #[test]
    fn page_border_side_styles_are_honored() {
        let mut display_list = one_page(Vec::new(), None, None);
        display_list.pages[0]
            .page_borders
            .push(PageBorderPrimitive {
                x: Number::from(72),
                y: Number::from(72),
                w: Number::from(600),
                h: Number::from(800),
                z_order: None,
                top: Some(PageBorderSide {
                    width: Number::from(2),
                    color: "#ff0000".to_owned(),
                    style: "double".to_owned(),
                }),
                right: None,
                bottom: None,
                left: None,
            });
        let bytes = export_pdf(&display_list).expect("export");
        let ops = paint_ops(&inflate_page_content(&bytes));
        let strokes: Vec<&PaintOp> = ops
            .iter()
            .filter(|op| op.op == "S" || op.op == "s")
            .collect();
        assert_eq!(
            strokes.len(),
            2,
            "a double page border side paints two lanes"
        );
        for stroke in strokes {
            assert!(!stroke.dashed);
            assert_eq!(stroke.stroke, Some([1.0, 0.0, 0.0]));
        }
    }

    #[test]
    fn line_opacity_uses_a_stroke_alpha_state() {
        let mut primitive = line(72.0, 200.0, 172.0, 200.0, 1.0, "#ff0000");
        if let Primitive::Line(line) = &mut primitive {
            line.opacity = Some(number(0.25));
        }
        let bytes = export_pdf(&one_page(vec![primitive], None, None)).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert_eq!(ext_g_state_alpha(&pdf, "/CA "), Some(0.25));
        let content = inflate_page_content(&bytes);
        let state = content.find("/GS1 gs").expect("stroke alpha state");
        let stroke = content.find("S\nQ").expect("isolated stroke");
        assert!(
            state < stroke,
            "the line must paint inside the stroke-alpha graphics state"
        );
    }

    #[test]
    fn zero_opacity_image_emits_ca_zero() {
        let mut png = std::io::Cursor::new(Vec::new());
        let source = ::image::RgbaImage::from_pixel(2, 2, ::image::Rgba([10, 20, 30, 255]));
        ::image::DynamicImage::ImageRgba8(source)
            .write_to(&mut png, ::image::ImageFormat::Png)
            .expect("encode png");
        let mut images = BTreeMap::new();
        images.insert("rId9".to_owned(), png.into_inner());

        let display_list = one_page(
            vec![Primitive::Image(ImagePrimitive {
                rel_id: "rId9".to_owned(),
                x: Number::from(72),
                y: Number::from(72),
                w: Number::from(48),
                h: Number::from(48),
                rotation_deg: None,
                opacity: Some(number(0.0)),
                filter: None,
                decorative: false,
                crop: None,
                alt_text: None,
                relationship_id: None,
                attrs: DocAttrs::default(),
            })],
            None,
            None,
        );
        let mut store = FontStore::default();
        let bytes =
            export_pdf_with_fonts_and_images(&display_list, &mut store, &images).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert_eq!(ext_g_state_alpha(&pdf, "/ca "), Some(0.0));
        assert_eq!(ext_g_state_alpha(&pdf, "/CA "), Some(0.0));
        let content = inflate_page_content(&bytes);
        let state = content.find("/GS1 gs").expect("zero-alpha graphics state");
        let paint = content.find("/Im1 Do").expect("image paint");
        assert!(
            state < paint,
            "the image must paint inside the zero-alpha graphics state"
        );
    }

    #[test]
    fn rotated_image_link_annotation_covers_the_painted_frame() {
        let image = ImagePrimitive {
            rel_id: "rId7".to_owned(),
            x: Number::from(100),
            y: Number::from(100),
            w: Number::from(200),
            h: Number::from(20),
            rotation_deg: Some(Number::from(90)),
            opacity: None,
            filter: None,
            decorative: false,
            crop: None,
            alt_text: None,
            relationship_id: None,
            attrs: DocAttrs {
                href: Some("https://example.com/logo".to_owned()),
                ..DocAttrs::default()
            },
        };
        let display_list = one_page(vec![Primitive::Image(image.clone())], None, None);
        let page = &display_list.pages[0];
        let links = crate::link::collect_links(page);
        assert_eq!(links.len(), 1, "the rotated image must stay clickable");
        let rect = links[0].rect;
        // Hand-computed: a 150x15pt frame at (75, 702) rotated 90° about its
        // centre paints over [142.5, 634.5, 157.5, 784.5]; an unrotated link
        // box would be [75, 702, 225, 717] and miss the painted frame.
        for (actual, expected) in rect.iter().zip([142.5_f32, 634.5, 157.5, 784.5]) {
            assert!(
                (actual - expected).abs() < 0.01,
                "annotation rect {rect:?} should cover the rotated frame"
            );
        }
        // Cross-check against the painter: every corner of the image unit
        // square, mapped through the placement the exporter uses, must lie
        // inside the clickable rectangle.
        let page_height = points(&page.height, "page height").expect("page height");
        let transform = crate::image::image_transform(&image, page_height).expect("transform");
        let [a, b, c, d, e, f] = transform.matrix;
        for (u, v) in [(0.0_f32, 0.0_f32), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)] {
            let x = a * u + c * v + e;
            let y = b * u + d * v + f;
            assert!(
                x >= rect[0] - 0.01
                    && x <= rect[2] + 0.01
                    && y >= rect[1] - 0.01
                    && y <= rect[3] + 0.01,
                "painted corner ({x}, {y}) must lie inside link rect {rect:?}"
            );
        }
    }

    #[test]
    fn writes_a4_page_with_text() {
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
                page_index: 0,
                width: Number::from(816),
                height: Number::from(1056),
                content_bounds: None,
                column_bounds: Vec::new(),
                section_id: None,
                section_index: None,
                section_page_index: None,
                section_page_number: None,
                page_label: None,
                primitives: vec![text(72.0, 96.0, "Hello")],
                watermark_primitive_count: None,
                background: Some("#ffffff".to_owned()),
                page_borders: Vec::new(),
                header: None,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        };
        let bytes = export_pdf(&display_list).expect("export");
        let pdf = String::from_utf8_lossy(&bytes);
        assert!(pdf.starts_with("%PDF-1."));
        assert!(pdf.contains("/MediaBox [0 0 612 792]"));
        let content = inflate_first_stream(&bytes);
        assert!(content.contains("Hello"));
    }

    #[test]
    fn rejects_non_ascii_without_fonts() {
        let display_list = DisplayList {
            contract_version: None,
            pages: vec![DisplayPage {
                page_index: 0,
                width: Number::from(816),
                height: Number::from(1056),
                content_bounds: None,
                column_bounds: Vec::new(),
                section_id: None,
                section_index: None,
                section_page_index: None,
                section_page_number: None,
                page_label: None,
                primitives: vec![text(72.0, 96.0, "你好")],
                watermark_primitive_count: None,
                background: None,
                page_borders: Vec::new(),
                header: None,
                footer: None,
                note_areas: Vec::new(),
                reservations: Vec::new(),
                table_gridlines: Vec::new(),
                unbuilt: false,
                position_span: None,
                hf_parts: None,
            }],
        };
        assert!(matches!(
            export_pdf(&display_list),
            Err(PdfError::Unsupported(_))
        ));
    }

    /// One `TextRunPrimitive` with placement overrides, exported through the
    /// no-font Base-14 path.
    fn export_base14_text(
        x: f64,
        baseline: f64,
        value: &str,
        edit: impl FnOnce(&mut TextRunPrimitive),
    ) -> (Vec<u8>, String) {
        let mut primitive = text(x, baseline, value);
        if let Primitive::Text(run) = &mut primitive {
            edit(run);
        }
        let bytes = export_pdf(&one_page(vec![primitive], None, None)).expect("export");
        let content = inflate_page_content(&bytes);
        (bytes, content)
    }

    #[test]
    fn rotated_base14_text_matrix_pivots_about_the_run_centre() {
        // A 16px run at (100, 200)px with a 48px measured width, rotated 90°
        // clockwise in display space.
        let (_, content) = export_base14_text(100.0, 200.0, "Rotated", |run| {
            run.width = number(48.0);
            run.font = "16px Helvetica".to_owned();
            run.rotation_deg = Some(number(90.0));
        });
        let matrices = tm_matrices(&content);
        assert_eq!(matrices.len(), 1, "one text matrix per run: {content}");
        // The 48x16px box pivots about (124, 196.8)px = (93, 644.4)pt, so the
        // glyph origin lands 18pt above and 2.4pt left of that centre. A
        // non-pivoted matrix would keep the origin at (75, 642)pt.
        let expected = [0.0, -1.0, 1.0, 0.0, 90.6, 662.4];
        for (actual, expected) in matrices[0].iter().zip(expected) {
            assert!(
                (actual - expected).abs() < 1e-3,
                "matrix {:?} should be {expected:?}",
                matrices[0]
            );
        }
    }

    #[test]
    fn base14_text_honours_horizontal_scale_once() {
        let (_, content) = export_base14_text(72.0, 96.0, "AB", |run| {
            run.horizontal_scale = Some(number(200.0));
        });
        assert_eq!(
            content.matches(" Tz").count(),
            1,
            "horizontal scale must be written exactly once: {content}"
        );
        assert!(content.contains("200 Tz"), "{content}");
        assert!(
            !content.contains("400 Tz"),
            "the scale must not be squared: {content}"
        );
        let matrices = tm_matrices(&content);
        assert_eq!(matrices.len(), 1);
        // Tz never scales a text matrix translation, so 72px stays 54pt even
        // though every glyph advance is doubled.
        assert!((matrices[0][4] - 54.0).abs() < 1e-3, "{:?}", matrices[0]);
        assert!((matrices[0][5] - 720.0).abs() < 1e-3, "{:?}", matrices[0]);
        assert_eq!(
            [
                matrices[0][0],
                matrices[0][1],
                matrices[0][2],
                matrices[0][3]
            ],
            [1.0, 0.0, 0.0, 1.0]
        );
    }

    #[test]
    fn rtl_base14_text_starts_at_the_measured_right_edge() {
        let (_, content) = export_base14_text(72.0, 96.0, "abcd", |run| {
            run.rtl = Some(true);
        });
        let matrices = tm_matrices(&content);
        assert_eq!(matrices.len(), 4, "one matrix per byte: {content}");
        // The 100px-wide run has four 25px byte slots. `width` already
        // includes scale, so the first byte starts at x + width - 25px and
        // the last byte lands on the left edge, all in unscaled points.
        let xs: Vec<f64> = matrices.iter().map(|matrix| matrix[4]).collect();
        for (actual, expected) in xs.iter().zip([110.25, 91.5, 72.75, 54.0]) {
            assert!(
                (actual - expected).abs() < 1e-3,
                "byte origins {xs:?} should be {expected:?}"
            );
        }
    }

    #[test]
    fn base14_text_rejects_invalid_horizontal_scale() {
        for scale in [0.0, -50.0] {
            let mut primitive = text(72.0, 96.0, "AB");
            if let Primitive::Text(run) = &mut primitive {
                run.horizontal_scale = Some(number(scale));
            }
            assert!(
                matches!(
                    export_pdf(&one_page(vec![primitive], None, None)),
                    Err(PdfError::InvalidNumber("horizontal scale"))
                ),
                "scale {scale} must be refused like the font path"
            );
        }
    }

    #[test]
    fn rotated_base14_text_link_annotation_follows_the_painter() {
        let mut primitive = text(100.0, 200.0, "Rotated");
        if let Primitive::Text(run) = &mut primitive {
            run.width = number(48.0);
            run.font = "16px Helvetica".to_owned();
            run.rotation_deg = Some(number(90.0));
            run.attrs.href = Some("https://example.com/rotated".to_owned());
        }
        let display_list = one_page(vec![primitive], None, None);
        let links = crate::link::collect_links(&display_list.pages[0]);
        assert_eq!(links.len(), 1, "the rotated run must stay clickable");

        let bytes = export_pdf(&display_list).expect("export");
        let content = inflate_page_content(&bytes);
        let pdf = String::from_utf8_lossy(&bytes);
        let rect = first_annotation_rect(&pdf).expect("link annotation rect");
        for (actual, expected) in rect.iter().zip(links[0].rect) {
            assert!(
                (actual - f64::from(expected)).abs() < 0.01,
                "annotation {rect:?} should match the link box {:?}",
                links[0].rect
            );
        }
        let matrix = tm_matrices(&content)[0];
        // The painter pivots the run about (93, 644.4)pt, so both ends of the
        // rotated baseline (90.6, 662.4) and (90.6, 626.4) must stay inside
        // the clickable box; an unrotated annotation would miss them.
        for (x, y) in [
            (matrix[4], matrix[5]),
            (matrix[4] + matrix[0] * 36.0, matrix[5] + matrix[1] * 36.0),
        ] {
            assert!(
                x >= rect[0] - 0.01
                    && x <= rect[2] + 0.01
                    && y >= rect[1] - 0.01
                    && y <= rect[3] + 0.01,
                "painted baseline point ({x}, {y}) must lie inside link rect {rect:?}"
            );
        }
        assert!(
            (rect[0] - 87.6).abs() < 0.01
                && (rect[1] - 626.4).abs() < 0.01
                && (rect[2] - 100.2).abs() < 0.01
                && (rect[3] - 662.4).abs() < 0.01,
            "annotation {rect:?} should cover the rotated run"
        );
    }
}
