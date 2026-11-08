//! Font registry over raw font bytes.

use std::borrow::Cow;
use std::cell::{Cell, RefCell};
use std::collections::{HashMap, HashSet};
use std::hash::{Hash, Hasher};
use std::ops::Deref;
use std::sync::OnceLock;
use std::sync::atomic::{AtomicU64, Ordering};

use skrifa::raw::TableProvider;
use skrifa::raw::tables::os2::SelectionFlags;
use skrifa::{FontRef, MetadataProvider};

use crate::family_metrics::{FamilyMetrics, family_advance, family_metrics, runs_at_own_widths};
use crate::shape::ShapedGlyph;

const MAX_SHAPE_CACHE_ENTRIES: usize = 4_096;
pub(crate) const MAX_CACHED_SHAPE_TEXT_BYTES: usize = 64 * 1024;
const MAX_CACHED_SHAPE_GLYPHS: usize = 64 * 1024;
const MAX_SHAPE_CACHE_TEXT_BYTES: usize = 8 * 1024 * 1024;
const MAX_SHAPE_CACHE_GLYPHS: usize = 1_000_000;
const MAX_CHAR_CACHE_ENTRIES: usize = 8_192;

const MAX_CACHED_SHAPE_FEATURES: usize = 64;
const BASE_LATIN_ID_START: u32 = u32::MAX - 4;
const MAX_BASE_LATIN_MISSES: usize = 4_096;
/// Bytes from each end of a font that key the identical-bytes lookup; the
/// table directory up front carries every table's checksum.
const FACE_KEY_SPAN: usize = 4 * 1024;

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub(crate) struct ShapeCacheKey {
    pub font: FontId,
    pub text: String,
    pub size_bits: u32,
    pub features: Vec<([u8; 4], u32)>,
    pub direction: u8,
    pub language: Option<String>,
}

impl ShapeCacheKey {
    /// Retained bytes for cache accounting: the owned text plus the owned
    /// key fields callers control (features, language).
    fn weight(&self) -> usize {
        self.text.len()
            + self.language.as_ref().map_or(0, String::len)
            + self.features.len() * std::mem::size_of::<([u8; 4], u32)>()
    }
}

/// One generation of the segmented shape cache, with its size accounting.
#[derive(Default)]
struct ShapeCacheGeneration {
    entries: HashMap<ShapeCacheKey, Vec<ShapedGlyph>>,
    text_bytes: usize,
    glyphs: usize,
}

impl ShapeCacheGeneration {
    fn would_overflow(&self, key: &ShapeCacheKey, glyphs: usize) -> bool {
        self.entries.len() >= MAX_SHAPE_CACHE_ENTRIES
            || self.text_bytes.saturating_add(key.weight()) > MAX_SHAPE_CACHE_TEXT_BYTES
            || self.glyphs.saturating_add(glyphs) > MAX_SHAPE_CACHE_GLYPHS
    }

    fn insert(&mut self, key: ShapeCacheKey, glyphs: Vec<ShapedGlyph>) {
        if let Some(replaced) = self.entries.remove(&key) {
            self.text_bytes = self.text_bytes.saturating_sub(key.weight());
            self.glyphs = self.glyphs.saturating_sub(replaced.len());
        }
        self.text_bytes += key.weight();
        self.glyphs += glyphs.len();
        self.entries.insert(key, glyphs);
    }

    fn remove(&mut self, key: &ShapeCacheKey) -> Option<Vec<ShapedGlyph>> {
        let glyphs = self.entries.remove(key)?;
        self.text_bytes = self.text_bytes.saturating_sub(key.weight());
        self.glyphs = self.glyphs.saturating_sub(glyphs.len());
        Some(glyphs)
    }
}

/// Segmented (two-generation) shape cache: inserts and hits live in `hot`;
/// when `hot` reaches a cap it becomes `cold` and a fresh `hot` starts. A
/// `cold` hit promotes the entry back into `hot`. Entries not touched for a
/// full generation age out with `cold` — a gradual LRU approximation with
/// O(1) operations and no wholesale invalidation cliff.
#[derive(Default)]
struct ShapeCache {
    hot: ShapeCacheGeneration,
    cold: ShapeCacheGeneration,
}

impl ShapeCache {
    fn insert_hot(&mut self, key: ShapeCacheKey, glyphs: Vec<ShapedGlyph>) {
        if self.hot.would_overflow(&key, glyphs.len()) {
            self.cold = std::mem::take(&mut self.hot);
        }
        self.hot.insert(key, glyphs);
    }
}

/// Per-character cmap/advance lookup memo. `mapped` is the raw cmap result
/// (including glyph id 0); `advance` is the design-space advance for the
/// mapped glyph, when the font reports one.
#[derive(Clone, Copy, Debug)]
struct CharEntry {
    mapped: Option<u16>,
    advance: Option<f32>,
}

/// Opaque handle to a font registered in a [`FontStore`].
///
/// The display-list contract carries `font_id` (this type), never a font
/// name; hosts resolve names to bytes before registration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FontId(pub(crate) u32);

impl FontId {
    /// Reserved handle for an always-available backstop face.
    pub const fn base_latin(bold: bool, italic: bool) -> Self {
        Self(BASE_LATIN_ID_START + bold as u32 * 2 + italic as u32)
    }

    /// Style of a reserved backstop handle.
    pub fn base_latin_style(self) -> Option<(bool, bool)> {
        #[cfg(feature = "base-latin")]
        {
            self.0
                .checked_sub(BASE_LATIN_ID_START)
                .filter(|&style| style < 4)
                .map(|style| (style & 2 != 0, style & 1 != 0))
        }
        #[cfg(not(feature = "base-latin"))]
        None
    }

    /// Raw index value, for serialization into display-list primitives.
    pub fn to_u32(self) -> u32 {
        self.0
    }

    /// Reconstruct a handle from the raw value [`FontId::to_u32`] produced —
    /// the inverse the display-list builder (and the outline wasm export) need
    /// to turn `fontChains` u32 ids back into store handles. Validated against
    /// the store at query time, so an out-of-range id surfaces as
    /// [`FontError::UnknownFont`], never undefined behavior.
    pub fn from_u32(raw: u32) -> Self {
        FontId(raw)
    }
}

/// Why [`FontStore::update_outlines`] refused new bytes for a face.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OutlineUpdateError {
    /// The [`FontId`] does not belong to this store.
    UnknownFont,
    /// The new bytes are not a parseable sfnt font.
    Parse(String),
    /// A table other than `CFF ` (or `head`'s checksum adjustment) differs,
    /// so shaping, metrics or layout could change.
    TablesDiffer,
}

impl std::fmt::Display for OutlineUpdateError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::UnknownFont => f.write_str("unknown font id"),
            Self::Parse(message) => write!(f, "font parse error: {message}"),
            Self::TablesDiffer => f.write_str("new outlines change a table other than CFF"),
        }
    }
}

impl std::error::Error for OutlineUpdateError {}

/// Errors produced when registering or querying fonts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FontError {
    /// The bytes are not a parseable sfnt font.
    Parse(String),
    /// A table this engine requires is missing (`hhea`, `OS/2`, ...).
    /// Word's line-height rules need `OS/2` winAscent/winDescent, so fonts
    /// without it are rejected at the registration boundary rather than
    /// producing wrong layout later.
    MissingTable(&'static str),
    /// The [`FontId`] does not belong to this store.
    UnknownFont,
    /// The font has no glyph with this id (id ≥ the font's glyph count).
    /// Raised by [`FontStore::outline_glyph`].
    GlyphOutOfRange(u16),
    /// The glyph outline draw failed (a malformed outline table on an
    /// attacker-controlled font). Raised by [`FontStore::outline_glyph`].
    Outline(String),
}

impl std::fmt::Display for FontError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            FontError::Parse(msg) => write!(f, "font parse error: {msg}"),
            FontError::MissingTable(t) => write!(f, "font is missing required table {t}"),
            FontError::UnknownFont => write!(f, "unknown FontId for this FontStore"),
            FontError::GlyphOutOfRange(g) => write!(f, "glyph id {g} out of range for this font"),
            FontError::Outline(msg) => write!(f, "glyph outline error: {msg}"),
        }
    }
}

impl std::error::Error for FontError {}

/// `head.unitsPerEm`, the `hhea` line metrics and the advance pitch of a face
/// a document asked for but the host could not supply, carried by the
/// substitute that stands in for it. See [`FontStore::register_substitute`].
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RequestedLineMetrics {
    pub units_per_em: u16,
    pub hhea_ascender: i16,
    pub hhea_descender: i16,
    pub hhea_line_gap: i16,
    /// Whether Word measures the requested family with the East Asian pitch
    /// ([`crate::word_metrics::EAST_ASIAN_CODE_PAGES`]) rather than the Latin
    /// win-box rule. Decides which fields the view has to move.
    pub east_asian: bool,
    /// Fallback advance ratio when no named width table applies.
    /// `1.0` leaves the substitute's own pitch alone.
    pub advance_scale: f32,
}

/// Design-space metrics extracted at registration time, in font units.
///
/// Both the `hhea` and `OS/2` variants support the line-metric experiments in
/// [`crate::word_metrics`]. On an entry made by
/// [`FontStore::register_substitute`] the vertical fields describe the face the
/// document asked for rather than the bytes; `units_per_em` always describes
/// the bytes, so shaping and outlines are unaffected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FontMetrics {
    pub units_per_em: u16,
    pub hhea_ascender: i16,
    pub hhea_descender: i16,
    pub hhea_line_gap: i16,
    pub os2_typo_ascender: i16,
    pub os2_typo_descender: i16,
    pub os2_typo_line_gap: i16,
    pub os2_win_ascent: u16,
    pub os2_win_descent: u16,
    /// `OS/2` fsSelection, defined bits only (the parser drops reserved bits).
    pub os2_fs_selection: u16,
    /// `OS/2` table version, which decides how much of the table is defined.
    pub os2_version: u16,
    /// `OS/2` ulCodePageRange1, or 0 before table version 1 defined it.
    pub os2_code_page_range1: u32,
}

impl FontMetrics {
    /// `OS/2` fsSelection bit 7 (USE_TYPO_METRICS): the font asks for its
    /// sTypo values to drive default line spacing.
    ///
    /// Only defined from table version 4. Versions 0-3 reserve the bit, so a
    /// set bit there is meaningless and must not select the typo family.
    pub fn use_typo_metrics(&self) -> bool {
        self.os2_version >= 4
            && self.os2_fs_selection & SelectionFlags::USE_TYPO_METRICS.bits() != 0
    }

    /// Whether Word measures this face with the East Asian line pitch — see
    /// [`crate::word_metrics::EAST_ASIAN_CODE_PAGES`].
    pub fn east_asian_line_metrics(&self) -> bool {
        self.os2_code_page_range1 & crate::word_metrics::EAST_ASIAN_CODE_PAGES != 0
    }
}

struct FontEntry {
    backstop_style: OnceLock<Option<(bool, bool)>>,
    /// Entry owning the bytes, for a metrics-only view of another face;
    /// `None` when this entry owns its own.
    bytes_of: Option<usize>,
    // Declared before `data` so the borrowing parser views drop first.
    //
    // SAFETY invariants for the `'static` lifetimes below: both views borrow
    // `data`'s allocation. `data` is static bytes or an owned Vec that is never
    // mutated in place and dropped only when the whole entry drops (entries are
    // never removed individually — the store only grows or drops wholesale) or
    // when `FontStore::update_outlines` replaces it, which takes `&mut FontStore`
    // and drops `face` before `data`. Moving `FontEntry` (e.g. a `Vec` realloc)
    // moves the Vec pointer, not the heap bytes, so the views stay valid.
    face: Option<Box<rustybuzz::Face<'static>>>,
    data: Cow<'static, [u8]>,
    metrics: FontMetrics,
    /// Horizontal multiplier for fallback advances, offsets and outlines.
    advance_scale: f32,
    widths: Option<&'static FamilyMetrics>,
}

struct RegisteredFontEntry {
    entry: FontEntry,
    char_cache: RefCell<HashMap<char, CharEntry>>,
}

impl Deref for RegisteredFontEntry {
    type Target = FontEntry;

    fn deref(&self) -> &Self::Target {
        &self.entry
    }
}

impl From<FontEntry> for RegisteredFontEntry {
    fn from(entry: FontEntry) -> Self {
        Self {
            entry,
            char_cache: RefCell::new(HashMap::new()),
        }
    }
}

#[cfg(feature = "base-latin")]
static BASE_LATIN_ENTRIES: [OnceLock<FontEntry>; 4] = [const { OnceLock::new() }; 4];

/// Registry of fonts, keyed by [`FontId`], parsed from raw bytes.
///
/// The store owns the byte buffers borrowed by shaping and outline extraction.
/// Registering bytes it already holds adds an alias sharing them.
pub struct FontStore {
    id: u64,
    outline_revision: u64,
    fonts: Vec<RegisteredFontEntry>,
    faces: HashMap<u64, Vec<usize>>,
    face_count: usize,
    shape_cache: RefCell<ShapeCache>,
    missing_base_latin: RefCell<HashSet<u64>>,
    new_base_latin_miss: Cell<bool>,
    #[cfg(feature = "base-latin")]
    base_latin_char_cache: [RefCell<HashMap<char, CharEntry>>; 4],
}

static NEXT_FONT_STORE_ID: AtomicU64 = AtomicU64::new(0);

/// Covered glyph character, including nonbreaking-hyphen stand-ins.
pub fn covered_character(
    store: &FontStore,
    font: FontId,
    ch: char,
) -> Result<Option<char>, FontError> {
    if store.covers(font, ch)? {
        return Ok(Some(ch));
    }
    if ch == '\u{2011}' {
        for stand_in in ['\u{2010}', '-'] {
            if store.covers(font, stand_in)? {
                return Ok(Some(stand_in));
            }
        }
    }
    Ok(None)
}

pub fn base_latin_alias(ch: char) -> bool {
    matches!(
        ch,
        '\u{f0b7}' | '\u{f0a7}' | '\u{f06f}' | '\u{f0d8}' | '\u{f0fc}' | '\u{f076}' | '\u{f0a8}'
    )
}

pub const BASE_LATIN_RANGES: [(u32, u32); 7] = [
    (0x20, 0x24f),
    (0x370, 0x4ff),
    (0x2000, 0x206f),
    (0x20a0, 0x20cf),
    (0x2190, 0x22ff),
    (0x25a0, 0x27bf),
    (0xfb00, 0xfb06),
];
pub const BASE_LATIN_ALIASES: [char; 7] = [
    '\u{f06f}', '\u{f076}', '\u{f0a7}', '\u{f0a8}', '\u{f0b7}', '\u{f0d8}', '\u{f0fc}',
];

pub fn base_latin_character_index(ch: char) -> Option<usize> {
    let point = ch as u32;
    let mut offset = 0;
    for (start, end) in BASE_LATIN_RANGES {
        if (start..=end).contains(&point) {
            return Some((offset + point - start) as usize);
        }
        offset += end - start + 1;
    }
    BASE_LATIN_ALIASES
        .iter()
        .position(|&alias| alias == ch)
        .map(|index| offset as usize + index)
}

pub fn symbol_family(family: &str) -> bool {
    let family = family.trim();
    family.eq_ignore_ascii_case("Symbol")
        || family
            .get(..9)
            .is_some_and(|name| name.eq_ignore_ascii_case("Wingdings"))
            && matches!(&family[9..], "" | " 2" | " 3")
}

fn miss_key(chain: &[FontId], family: &str, bold: bool, italic: bool, ch: char) -> u64 {
    let mut hasher = std::hash::DefaultHasher::new();
    (chain, family, bold, italic, ch).hash(&mut hasher);
    hasher.finish()
}

impl Default for FontStore {
    fn default() -> Self {
        Self::new()
    }
}

impl FontStore {
    pub fn new() -> Self {
        Self {
            id: NEXT_FONT_STORE_ID.fetch_add(1, Ordering::Relaxed),
            outline_revision: 0,
            fonts: Vec::new(),
            faces: HashMap::new(),
            face_count: 0,
            shape_cache: RefCell::default(),
            missing_base_latin: RefCell::default(),
            new_base_latin_miss: Cell::new(false),
            #[cfg(feature = "base-latin")]
            base_latin_char_cache: std::array::from_fn(|_| RefCell::new(HashMap::new())),
        }
    }

    /// Identity unique to this store within the process. [`FontId`]s are
    /// store-local indexes, so anything keyed by them (glyph caches) binds to
    /// this value instead of silently reusing another store's outlines.
    pub fn id(&self) -> u64 {
        self.id
    }

    /// Advances on every [`FontStore::update_outlines`]: outline caches keyed by
    /// [`FontId`] drop what they hold when it changes. Metrics and shaping never
    /// change with it.
    pub fn outline_revision(&self) -> u64 {
        self.outline_revision
    }

    /// How many fonts the store holds. Fonts are only ever added, so a store with the same id
    /// and count holds the same fonts.
    pub fn font_count(&self) -> usize {
        self.fonts.len()
    }

    /// How many distinct font byte buffers the store holds. Aliases and
    /// substitute views share their owner's.
    pub fn face_count(&self) -> usize {
        self.face_count
    }

    /// The registered face holding exactly `bytes`, which a registration of
    /// the same bytes would alias rather than copy.
    pub fn find_face(&self, bytes: &[u8]) -> Option<FontId> {
        self.faces
            .get(&face_key(bytes))?
            .iter()
            .find(|&&index| *self.fonts[index].data == *bytes)
            .map(|&index| FontId(index as u32))
    }

    /// Parse and register a font from raw bytes, returning its handle.
    ///
    /// Metrics are extracted eagerly so later queries are infallible cheap
    /// lookups and malformed fonts are rejected at the trust boundary
    /// (embedded DOCX fonts are attacker-controlled). Bytes the store already
    /// holds register as a new id sharing them.
    pub fn register(&mut self, bytes: Vec<u8>) -> Result<FontId, FontError> {
        match self.find_face(&bytes) {
            Some(owner) => Ok(self.register_alias(owner)),
            None => self.register_face(bytes),
        }
    }

    /// [`FontStore::register`] for bytes held either way, copied only when
    /// borrowed and the store does not already hold them.
    pub fn register_cow(&mut self, bytes: std::borrow::Cow<'_, [u8]>) -> Result<FontId, FontError> {
        match bytes {
            std::borrow::Cow::Borrowed(bytes) => self.register_slice(bytes),
            std::borrow::Cow::Owned(bytes) => self.register(bytes),
        }
    }

    /// [`FontStore::register`] for borrowed bytes, copied only when the store
    /// does not already hold them.
    pub fn register_slice(&mut self, bytes: &[u8]) -> Result<FontId, FontError> {
        match self.find_face(bytes) {
            Some(owner) => Ok(self.register_alias(owner)),
            None => self.register_face(bytes.to_vec()),
        }
    }

    fn register_alias(&mut self, owner: FontId) -> FontId {
        let metrics = self.fonts[owner.0 as usize].metrics;
        let id = FontId(self.fonts.len() as u32);
        self.fonts.push(
            FontEntry {
                backstop_style: OnceLock::new(),
                bytes_of: Some(owner.0 as usize),
                face: None,
                data: Cow::Borrowed(&[]),
                metrics,
                advance_scale: 1.0,
                widths: None,
            }
            .into(),
        );
        id
    }

    fn register_face(&mut self, bytes: Vec<u8>) -> Result<FontId, FontError> {
        let entry = Self::parse_entry(Cow::Owned(bytes))?;
        let index = self.fonts.len();
        self.faces
            .entry(face_key(&entry.data))
            .or_default()
            .push(index);
        self.face_count += 1;
        self.fonts.push(entry.into());
        Ok(FontId(index as u32))
    }

    fn parse_entry(data: Cow<'static, [u8]>) -> Result<FontEntry, FontError> {
        let font = FontRef::new(&data).map_err(|e| FontError::Parse(e.to_string()))?;

        let head = font.head().map_err(|_| FontError::MissingTable("head"))?;
        let hhea = font.hhea().map_err(|_| FontError::MissingTable("hhea"))?;
        let os2 = font.os2().map_err(|_| FontError::MissingTable("OS/2"))?;

        let metrics = FontMetrics {
            units_per_em: head.units_per_em(),
            hhea_ascender: hhea.ascender().to_i16(),
            hhea_descender: hhea.descender().to_i16(),
            hhea_line_gap: hhea.line_gap().to_i16(),
            os2_typo_ascender: os2.s_typo_ascender(),
            os2_typo_descender: os2.s_typo_descender(),
            os2_typo_line_gap: os2.s_typo_line_gap(),
            os2_win_ascent: os2.us_win_ascent(),
            os2_win_descent: os2.us_win_descent(),
            os2_fs_selection: os2.fs_selection().bits(),
            os2_version: os2.version(),
            os2_code_page_range1: os2.ul_code_page_range_1().unwrap_or(0),
        };

        // SAFETY: see the `FontEntry` field invariants — the view borrows
        // `data`'s heap allocation, which outlives it and never moves.
        let pinned: &'static [u8] =
            unsafe { std::slice::from_raw_parts(data.as_ptr(), data.len()) };
        // A font rustybuzz rejects still registers (skrifa accepted it);
        // shaping then reports the parse failure per call, as before.
        let face = rustybuzz::Face::from_slice(pinned, 0).map(Box::new);

        Ok(FontEntry {
            backstop_style: OnceLock::new(),
            bytes_of: None,
            face,
            data,
            metrics,
            advance_scale: 1.0,
            widths: None,
        })
    }

    /// A measurement view of `base` carrying `requested`'s line metrics and
    /// advance pitch. Shares `base`'s bytes, so the glyph repertoire is
    /// unchanged; hosts put the returned id at the head of the fallback chain.
    /// A Latin entry mirrors the span into the win box and clears the East
    /// Asian code pages, so the view measures the same under either line rule
    /// whatever the substitute claims. A `requested.advance_scale` other than
    /// `1.0` widens every advance, offset and outline the view reports, so the
    /// same id measures and paints at the requested face's pitch.
    pub fn register_substitute(
        &mut self,
        base: FontId,
        requested: RequestedLineMetrics,
    ) -> Result<FontId, FontError> {
        self.register_substitute_view(base, Some(requested), None)
    }

    /// A substitute view with the named family's line metrics and glyph widths.
    /// Matching faces retain their own advances and kerning.
    pub fn register_named_substitute(
        &mut self,
        base: FontId,
        family: &str,
    ) -> Result<FontId, FontError> {
        let entry = self.entry(base)?;
        let owner = FontId(entry.bytes_of.unwrap_or(base.0 as usize) as u32);
        let selection = entry.metrics.os2_fs_selection;
        let family = family.trim().to_lowercase();
        let family = crate::word_fonts::requested_family_name(&family);
        let bold = selection & (1 << 5) != 0;
        let table = family_metrics(family, bold, selection & 1 != 0);
        let requested =
            crate::word_fonts::requested_line_metrics_styled(family, bold).map(|mut requested| {
                if table.is_some() {
                    requested.advance_scale = 1.0;
                }
                requested
            });
        let widths = table.filter(|metrics| !runs_at_own_widths(self, owner, metrics));
        if requested.is_none() && widths.is_none() {
            return Ok(base);
        }
        self.register_substitute_view(base, requested, widths)
    }

    fn register_substitute_view(
        &mut self,
        base: FontId,
        requested: Option<RequestedLineMetrics>,
        widths: Option<&'static FamilyMetrics>,
    ) -> Result<FontId, FontError> {
        let entry = self.entry(base)?;
        let owner = entry.bytes_of.unwrap_or(base.0 as usize);
        let mut metrics = entry.metrics;
        if metrics.units_per_em == 0 || requested.is_some_and(|metrics| metrics.units_per_em == 0) {
            return Ok(base);
        }
        let advance_scale = match widths {
            Some(_) if requested.is_some_and(|metrics| metrics.east_asian) => 1.0,
            Some(widths) => self.width_scale(FontId(owner as u32), widths)?,
            None => requested.map_or(1.0, |metrics| metrics.advance_scale),
        };
        let advance_scale = if advance_scale.is_finite() && advance_scale > 0.0 {
            advance_scale
        } else {
            1.0
        };
        if let Some(requested) = requested {
            let scale = f32::from(metrics.units_per_em) / f32::from(requested.units_per_em);
            let rescale = |design: i16| {
                (f32::from(design) * scale)
                    .round()
                    .clamp(f32::from(i16::MIN), f32::from(i16::MAX)) as i16
            };
            metrics.hhea_ascender = rescale(requested.hhea_ascender);
            metrics.hhea_descender = rescale(requested.hhea_descender);
            metrics.hhea_line_gap = rescale(requested.hhea_line_gap);
            if requested.east_asian {
                metrics.os2_code_page_range1 |= crate::word_metrics::EAST_ASIAN_CODE_PAGES;
            } else {
                metrics.os2_code_page_range1 &= !crate::word_metrics::EAST_ASIAN_CODE_PAGES;
                metrics.os2_win_ascent = metrics.hhea_ascender.max(0) as u16;
                metrics.os2_win_descent = metrics.hhea_descender.saturating_neg().max(0) as u16;
            }
        }

        let id = FontId(self.fonts.len() as u32);
        self.fonts.push(
            FontEntry {
                backstop_style: OnceLock::new(),
                bytes_of: Some(owner),
                face: None,
                data: Cow::Borrowed(&[]),
                metrics,
                advance_scale,
                widths,
            }
            .into(),
        );
        Ok(id)
    }

    /// Horizontal multiplier for fallback advances, offsets and outlines.
    /// Width tables take precedence for covered characters.
    pub fn advance_scale(&self, id: FontId) -> Result<f32, FontError> {
        self.entry(id).map(|entry| entry.advance_scale)
    }

    /// The bytes, metrics and fallback advance scale of `id`.
    pub fn measured_face(&self, id: FontId) -> Result<(&[u8], &FontMetrics, f32), FontError> {
        let entry = self.entry(id)?;
        Ok((
            &self.byte_entry(id)?.data,
            &entry.metrics,
            entry.advance_scale,
        ))
    }

    /// Width table carried by a substitute view, if its base widths differ.
    pub fn family_widths(&self, id: FontId) -> Result<Option<&'static FamilyMetrics>, FontError> {
        self.entry(id).map(|entry| entry.widths)
    }

    fn width_scale(&self, base: FontId, widths: &FamilyMetrics) -> Result<f32, FontError> {
        const FREQUENCIES: [f32; 26] = [
            8.167, 1.492, 2.782, 4.253, 12.702, 2.228, 2.015, 6.094, 6.966, 0.153, 0.772, 4.025,
            2.406, 6.749, 7.507, 1.929, 0.095, 5.987, 6.327, 9.056, 2.758, 0.978, 2.360, 0.150,
            1.974, 0.074,
        ];
        let upem = f32::from(self.metrics(base)?.units_per_em);
        let sample = ('a'..='z').zip(FREQUENCIES).chain(std::iter::once((
            ' ',
            FREQUENCIES.iter().sum::<f32>() / 5.1,
        )));
        let mut wanted = 0.0;
        let mut own = 0.0;
        for (ch, frequency) in sample {
            if let Some(width) = family_advance(widths, ch)
                && let Some(advance) = self.char_entry(base, ch)?.advance
            {
                wanted += width * frequency;
                own += advance / upem * frequency;
            }
        }
        Ok(if own > 0.0 { wanted / own } else { 1.0 })
    }

    /// Per-font design-space metrics captured at registration.
    pub fn metrics(&self, id: FontId) -> Result<&FontMetrics, FontError> {
        self.entry(id).map(|e| &e.metrics)
    }

    #[cfg(test)]
    pub(crate) fn replace_metrics_for_test(
        &mut self,
        id: FontId,
        metrics: FontMetrics,
    ) -> Result<(), FontError> {
        let entry = self
            .fonts
            .get_mut(id.0 as usize)
            .ok_or(FontError::UnknownFont)?;
        entry.entry.metrics = metrics;
        Ok(())
    }

    /// Raw bytes of a registered font (for shaping / outline extraction).
    pub fn font_bytes(&self, id: FontId) -> Result<&[u8], FontError> {
        self.byte_entry(id).map(|e| &*e.data)
    }

    /// Every family name the font's own name table declares, legacy and
    /// typographic, in every language, so a clone registered under another
    /// name reads as itself.
    pub fn declared_families(&self, id: FontId) -> Result<Vec<String>, FontError> {
        let font =
            FontRef::new(self.font_bytes(id)?).map_err(|e| FontError::Parse(e.to_string()))?;
        Ok([
            skrifa::string::StringId::FAMILY_NAME,
            skrifa::string::StringId::TYPOGRAPHIC_FAMILY_NAME,
        ]
        .into_iter()
        .flat_map(|name| font.localized_strings(name))
        .map(|name| name.to_string())
        .collect())
    }

    /// Replaces the outlines of the face holding `id`'s bytes with `bytes`,
    /// a face whose every table except `CFF ` is byte-identical (`head` may
    /// differ only in its checksum adjustment), such as a sliced face
    /// reassembled with more outline chunks. Every id sharing those bytes
    /// (aliases and measurement views) draws the new outlines; ids, metrics,
    /// cmap and shaping caches stay, since layout tables are unchanged.
    pub fn update_outlines(
        &mut self,
        id: FontId,
        bytes: Vec<u8>,
    ) -> Result<(), OutlineUpdateError> {
        let owner = self
            .fonts
            .get(id.0 as usize)
            .ok_or(OutlineUpdateError::UnknownFont)?
            .bytes_of
            .unwrap_or(id.0 as usize);
        let current = &self
            .fonts
            .get(owner)
            .ok_or(OutlineUpdateError::UnknownFont)?
            .data;
        if !same_layout_tables(current, &bytes)? {
            return Err(OutlineUpdateError::TablesDiffer);
        }
        let old_key = face_key(current);
        let new_key = face_key(&bytes);
        // SAFETY: see the `FontEntry` field invariants; the old face is
        // dropped before the bytes it borrows, and moving the Vec into the
        // entry keeps its heap allocation.
        let pinned: &'static [u8] =
            unsafe { std::slice::from_raw_parts(bytes.as_ptr(), bytes.len()) };
        let face = rustybuzz::Face::from_slice(pinned, 0).map(Box::new);
        let entry = &mut self.fonts[owner].entry;
        entry.face = None;
        entry.data = Cow::Owned(bytes);
        entry.face = face;
        if let Some(indices) = self.faces.get_mut(&old_key) {
            indices.retain(|&index| index != owner);
            if indices.is_empty() {
                self.faces.remove(&old_key);
            }
        }
        self.faces.entry(new_key).or_default().push(owner);
        self.outline_revision += 1;
        Ok(())
    }

    /// The registered face owning this id's unscaled font bytes.
    pub fn byte_owner(&self, id: FontId) -> Result<FontId, FontError> {
        let entry = self.entry(id)?;
        Ok(FontId(entry.bytes_of.unwrap_or(id.0 as usize) as u32))
    }

    /// Parsed shaping face memoized at registration. `None` when rustybuzz
    /// rejected bytes skrifa accepted.
    pub(crate) fn shaping_face(
        &self,
        id: FontId,
    ) -> Result<Option<&rustybuzz::Face<'_>>, FontError> {
        self.byte_entry(id).map(|e| e.face.as_deref())
    }

    pub(crate) fn cached_shape(&self, key: &ShapeCacheKey) -> Option<Vec<ShapedGlyph>> {
        let mut cache = self.shape_cache.borrow_mut();
        if let Some(glyphs) = cache.hot.entries.get(key) {
            return Some(glyphs.clone());
        }
        let glyphs = cache.cold.remove(key)?;
        let cloned = glyphs.clone();
        cache.insert_hot(key.clone(), glyphs);
        Some(cloned)
    }

    pub(crate) fn cache_shape(&self, key: ShapeCacheKey, glyphs: &[ShapedGlyph]) {
        if key.text.len() > MAX_CACHED_SHAPE_TEXT_BYTES
            || key.features.len() > MAX_CACHED_SHAPE_FEATURES
            || glyphs.len() > MAX_CACHED_SHAPE_GLYPHS
        {
            return;
        }
        let mut cache = self.shape_cache.borrow_mut();
        cache.cold.remove(&key);
        cache.insert_hot(key, glyphs.to_vec());
    }

    #[cfg(test)]
    pub(crate) fn shape_cache_len(&self) -> usize {
        let cache = self.shape_cache.borrow();
        cache.hot.entries.len() + cache.cold.entries.len()
    }

    /// Memoized cmap (+advance) lookup for one character of one font.
    fn char_entry(&self, id: FontId, ch: char) -> Result<CharEntry, FontError> {
        let id = self.byte_owner(id)?;
        let entry = self.byte_entry(id)?;
        #[cfg(feature = "base-latin")]
        let cache = if id.base_latin_style().is_some() {
            &self.base_latin_char_cache[(id.0 - BASE_LATIN_ID_START) as usize]
        } else {
            &self.fonts[id.0 as usize].char_cache
        };
        #[cfg(not(feature = "base-latin"))]
        let cache = &self.fonts[id.0 as usize].char_cache;
        if let Some(cached) = cache.borrow().get(&ch) {
            return Ok(*cached);
        }
        let font = Self::font_ref(entry);
        let mapped = font.charmap().map(ch);
        let advance = mapped.and_then(|gid| {
            font.glyph_metrics(
                skrifa::instance::Size::unscaled(),
                skrifa::instance::LocationRef::default(),
            )
            .advance_width(gid)
        });
        let computed = CharEntry {
            mapped: mapped.map(|g| g.to_u32() as u16),
            advance,
        };
        let mut cache = cache.borrow_mut();
        if cache.len() >= MAX_CHAR_CACHE_ENTRIES {
            cache.clear();
        }
        cache.insert(ch, computed);
        Ok(computed)
    }

    /// cmap lookup: glyph id for a character, `None` if the font does not
    /// cover it. Glyph id 0 (`.notdef`) counts as no coverage.
    pub fn glyph_id(&self, id: FontId, ch: char) -> Result<Option<u16>, FontError> {
        Ok(self.char_entry(id, ch)?.mapped.filter(|&g| g != 0))
    }

    /// Horizontal character advance in font units, using the view's widths.
    /// `None` if the character is not covered by this font's cmap.
    pub fn advance_width(&self, id: FontId, ch: char) -> Result<Option<f32>, FontError> {
        let entry = self.char_entry(id, ch)?;
        if entry.mapped.is_none() {
            return Ok(None);
        }
        let font = self.entry(id)?;
        if let Some(widths) = font.widths
            && let Some(advance) = family_advance(widths, ch)
        {
            return Ok(Some(advance * f32::from(font.metrics.units_per_em)));
        }
        Ok(entry.advance.map(|advance| advance * font.advance_scale))
    }

    /// Whether the font's cmap covers `ch`.
    pub fn covers(&self, id: FontId, ch: char) -> Result<bool, FontError> {
        Ok(self.glyph_id(id, ch)?.is_some())
    }

    /// Resolves the first covering font, skipping unknown IDs.
    pub fn resolve(&self, chain: &[FontId], ch: char) -> Option<FontId> {
        chain
            .iter()
            .copied()
            .find(|&id| covered_character(self, id, ch).ok().flatten().is_some())
    }

    /// Resolves a chain with the requested style for its implicit backstop.
    pub fn resolve_with_style(
        &self,
        chain: &[FontId],
        ch: char,
        bold: bool,
        italic: bool,
    ) -> Option<FontId> {
        self.resolve_with_family(chain, ch, bold, italic, "")
    }

    pub fn backstop_style(&self, id: FontId) -> Option<(bool, bool)> {
        #[cfg(feature = "base-latin")]
        if let Some(style) = id.base_latin_style() {
            return Some(style);
        }
        let entry = self.byte_entry(id).ok()?;
        *entry.backstop_style.get_or_init(|| {
            Self::font_ref(entry)
                .localized_strings(skrifa::string::StringId::FAMILY_NAME)
                .any(|name| name.to_string() == crate::BASE_LATIN_FAMILY)
                .then_some((
                    entry.metrics.os2_fs_selection & (1 << 5) != 0,
                    entry.metrics.os2_fs_selection & 1 != 0,
                ))
        })
    }

    pub fn resolve_with_family(
        &self,
        chain: &[FontId],
        ch: char,
        bold: bool,
        italic: bool,
        family: &str,
    ) -> Option<FontId> {
        self.resolve_with_family_tracking(chain, ch, bold, italic, family, false)
    }

    pub fn take_base_latin_miss(&self) -> bool {
        self.new_base_latin_miss.replace(false)
    }

    pub fn resolve_with_family_tracking(
        &self,
        chain: &[FontId],
        ch: char,
        bold: bool,
        italic: bool,
        family: &str,
        track: bool,
    ) -> Option<FontId> {
        if chain.is_empty() {
            return None;
        }
        if base_latin_alias(ch) && !symbol_family(family) {
            let allowed = |id: FontId| self.backstop_style(id).is_none();
            return chain
                .iter()
                .copied()
                .find(|&id| allowed(id) && covered_character(self, id, ch).ok().flatten().is_some())
                .or_else(|| chain.iter().rev().copied().find(|&id| allowed(id)));
        }
        let resolved = self.resolve(chain, ch);
        #[cfg(feature = "base-latin")]
        let resolved = resolved.or_else(|| {
            [
                FontId::base_latin(bold, italic),
                FontId::base_latin(false, false),
            ]
            .into_iter()
            .find(|&id| covered_character(self, id, ch).ok().flatten().is_some())
        });
        #[cfg(not(feature = "base-latin"))]
        let _ = (bold, italic);
        if track && resolved.is_none() {
            self.record_base_latin_miss(chain, ch, bold, italic, family);
        }
        resolved.or_else(|| chain.last().copied())
    }

    #[cold]
    #[inline(never)]
    fn record_base_latin_miss(
        &self,
        chain: &[FontId],
        ch: char,
        bold: bool,
        italic: bool,
        family: &str,
    ) {
        if !(base_latin_character_index(ch).is_some() && crate::measure::needs_glyph(ch)) {
            return;
        }
        let mut missing = self.missing_base_latin.borrow_mut();
        if missing.len() >= MAX_BASE_LATIN_MISSES {
            missing.clear();
        }
        if missing.insert(miss_key(chain, family, bold, italic, ch)) {
            self.new_base_latin_miss.set(true);
        }
    }

    fn entry(&self, id: FontId) -> Result<&FontEntry, FontError> {
        #[cfg(feature = "base-latin")]
        if let Some((bold, italic)) = id.base_latin_style() {
            return Ok(
                BASE_LATIN_ENTRIES[(id.0 - BASE_LATIN_ID_START) as usize].get_or_init(|| {
                    Self::parse_entry(Cow::Borrowed(base_latin::face(bold, italic)))
                        .expect("embedded backstop face is valid")
                }),
            );
        }
        self.fonts
            .get(id.0 as usize)
            .map(|entry| &entry.entry)
            .ok_or(FontError::UnknownFont)
    }

    /// Entry holding `id`'s bytes: itself, or the face a measurement view
    /// stands in for.
    fn byte_entry(&self, id: FontId) -> Result<&FontEntry, FontError> {
        match self.entry(id)?.bytes_of {
            Some(base) => self.entry(FontId(base as u32)),
            None => self.entry(id),
        }
    }

    // registration already proved the bytes parse, so this cannot fail
    fn font_ref(entry: &FontEntry) -> FontRef<'_> {
        FontRef::new(&entry.data).expect("bytes validated at register()")
    }
}

fn sfnt_tables(
    bytes: &[u8],
) -> Result<std::collections::BTreeMap<[u8; 4], &[u8]>, OutlineUpdateError> {
    let parse = || OutlineUpdateError::Parse("truncated sfnt table directory".to_owned());
    let count = usize::from(u16::from_be_bytes(
        bytes
            .get(4..6)
            .ok_or_else(parse)?
            .try_into()
            .map_err(|_| parse())?,
    ));
    let mut tables = std::collections::BTreeMap::new();
    for index in 0..count {
        let record = bytes
            .get(12 + index * 16..28 + index * 16)
            .ok_or_else(parse)?;
        let offset = u32::from_be_bytes(record[8..12].try_into().map_err(|_| parse())?) as usize;
        let length = u32::from_be_bytes(record[12..16].try_into().map_err(|_| parse())?) as usize;
        let end = offset.checked_add(length).ok_or_else(parse)?;
        let tag: [u8; 4] = record[..4].try_into().map_err(|_| parse())?;
        tables.insert(tag, bytes.get(offset..end).ok_or_else(parse)?);
    }
    Ok(tables)
}

/// Whether two faces differ at most in `CFF ` and `head`'s checksum adjustment.
fn same_layout_tables(current: &[u8], next: &[u8]) -> Result<bool, OutlineUpdateError> {
    FontRef::new(next).map_err(|error| OutlineUpdateError::Parse(error.to_string()))?;
    let (current, next) = (sfnt_tables(current)?, sfnt_tables(next)?);
    if current.keys().ne(next.keys()) || !current.contains_key(b"CFF ") {
        return Ok(false);
    }
    Ok(current.iter().all(|(tag, table)| match tag {
        b"CFF " => true,
        b"head" => {
            let other = next[tag];
            table.len() == other.len()
                && table.len() >= 12
                && table[..8] == other[..8]
                && table[12..] == other[12..]
        }
        _ => *table == next[tag],
    }))
}

fn face_key(bytes: &[u8]) -> u64 {
    let mut hasher = std::hash::DefaultHasher::new();
    bytes.len().hash(&mut hasher);
    bytes[..bytes.len().min(FACE_KEY_SPAN)].hash(&mut hasher);
    bytes[bytes.len().saturating_sub(FACE_KEY_SPAN)..].hash(&mut hasher);
    hasher.finish()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shape::ShapedGlyph;

    const LIBERATION_SANS: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn base_latin_miss_history_stays_bounded() {
        let store = FontStore::new();
        for index in 0..MAX_BASE_LATIN_MISSES * 2 {
            store.record_base_latin_miss(
                &[FontId(0)],
                'A',
                false,
                false,
                &format!("family {index}"),
            );
            assert!(store.missing_base_latin.borrow().len() <= MAX_BASE_LATIN_MISSES);
        }
        assert!(store.take_base_latin_miss());
        store.record_base_latin_miss(&[FontId(0)], 'A', false, false, "family 0");
        store.record_base_latin_miss(&[FontId(0)], 'A', false, false, "family 0");
        assert!(store.take_base_latin_miss());
        store.record_base_latin_miss(&[FontId(0)], 'A', false, false, "family 0");
        assert!(!store.take_base_latin_miss());
    }

    fn key(text: String) -> ShapeCacheKey {
        ShapeCacheKey {
            font: FontId(0),
            text,
            size_bits: 16f32.to_bits(),
            features: Vec::new(),
            direction: 1,
            language: None,
        }
    }

    const GLYPH: ShapedGlyph = ShapedGlyph {
        glyph_id: 1,
        cluster: 0,
        x_advance: 10.0,
        x_offset: 0.0,
        y_offset: 0.0,
    };

    #[test]
    fn shape_cache_rotation_retains_recent_entries_instead_of_clearing() {
        let store = FontStore::new();
        store.cache_shape(key("survivor".to_owned()), &[GLYPH]);
        for wave in 0..2 {
            for index in 0..MAX_SHAPE_CACHE_ENTRIES {
                store.cache_shape(key(format!("filler {wave} {index}")), &[GLYPH]);
            }
            // touching the survivor after each wave promotes it back to hot
            assert!(
                store.cached_shape(&key("survivor".to_owned())).is_some(),
                "a recently used entry survives generation rotation"
            );
        }
        assert!(
            store.shape_cache_len() <= 2 * MAX_SHAPE_CACHE_ENTRIES,
            "the segmented cache is bounded"
        );
        assert!(
            store.cached_shape(&key("filler 0 0".to_owned())).is_none(),
            "entries untouched for a full generation age out"
        );
    }

    #[test]
    fn char_lookups_are_memoized_per_font() {
        let mut store = FontStore::new();
        let font = store.register(LIBERATION_SANS.to_vec()).unwrap();
        let gid = store.glyph_id(font, 'A').unwrap();
        let advance = store.advance_width(font, 'A').unwrap();
        assert!(gid.is_some());
        assert!(advance.is_some());
        assert_eq!(store.glyph_id(font, 'A').unwrap(), gid);
        assert_eq!(store.advance_width(font, 'A').unwrap(), advance);
        assert_eq!(
            store.fonts[font.0 as usize].char_cache.borrow().len(),
            1,
            "repeat lookups reuse the memo"
        );
        assert!(!store.covers(font, '\u{10FFFF}').unwrap());
    }

    #[test]
    fn identical_bytes_register_as_aliases_sharing_one_buffer() {
        let mut store = FontStore::new();
        let ids: Vec<_> = (0..300)
            .map(|_| store.register_slice(LIBERATION_SANS).unwrap())
            .collect();
        let owned = store.register(LIBERATION_SANS.to_vec()).unwrap();
        assert_eq!(store.face_count(), 1);
        assert_eq!(store.font_count(), 301);
        assert_eq!(store.find_face(LIBERATION_SANS), Some(ids[0]));
        let held: usize = store.fonts.iter().map(|entry| entry.data.len()).sum();
        assert_eq!(held, LIBERATION_SANS.len());
        let shared = store.font_bytes(ids[0]).unwrap().as_ptr();
        for &id in ids.iter().chain([&owned]) {
            assert_eq!(store.byte_owner(id).unwrap(), ids[0]);
            assert_eq!(store.font_bytes(id).unwrap().as_ptr(), shared);
            assert_eq!(
                store.glyph_id(id, 'A').unwrap(),
                store.glyph_id(ids[0], 'A').unwrap()
            );
            assert_eq!(
                store.advance_width(id, 'A').unwrap(),
                store.advance_width(ids[0], 'A').unwrap()
            );
            assert!(store.shaping_face(id).unwrap().is_some());
        }
        let mut distinct = ids.clone();
        distinct.dedup();
        assert_eq!(distinct.len(), 300, "every registration keeps its own id");
    }

    #[test]
    fn owned_and_borrowed_bytes_share_one_face() {
        let mut store = FontStore::default();
        let owned = store
            .register_cow(std::borrow::Cow::Owned(LIBERATION_SANS.to_vec()))
            .unwrap();
        let borrowed = store
            .register_cow(std::borrow::Cow::Borrowed(LIBERATION_SANS))
            .unwrap();
        assert_ne!(owned, borrowed);
        assert_eq!(store.face_count(), 1);
    }

    #[test]
    fn different_bytes_register_as_their_own_faces() {
        let mut store = FontStore::new();
        let first = store.register_slice(LIBERATION_SANS).unwrap();
        let mut padded = LIBERATION_SANS.to_vec();
        padded.push(0);
        let second = store.register(padded.clone()).unwrap();
        assert_eq!(store.face_count(), 2);
        assert_eq!(store.byte_owner(second).unwrap(), second);
        assert_eq!(store.find_face(&padded), Some(second));
        assert_ne!(store.find_face(LIBERATION_SANS), Some(second));
        assert_eq!(store.find_face(LIBERATION_SANS), Some(first));
        assert!(store.register_slice(b"not a font").is_err());
        assert_eq!(store.face_count(), 2);
        assert_eq!(store.find_face(b"not a font"), None);
    }

    #[test]
    fn a_substitute_of_an_alias_shares_the_owner_bytes() {
        let mut store = FontStore::new();
        let owner = store.register_slice(LIBERATION_SANS).unwrap();
        let alias = store.register_slice(LIBERATION_SANS).unwrap();
        let view = store.register_named_substitute(alias, "Calibri").unwrap();
        assert_eq!(store.byte_owner(view).unwrap(), owner);
        assert_eq!(store.face_count(), 1);
    }

    #[test]
    fn shaping_face_is_memoized_at_registration() {
        let mut store = FontStore::new();
        let font = store.register(LIBERATION_SANS.to_vec()).unwrap();
        assert!(store.shaping_face(font).unwrap().is_some());
    }
}
