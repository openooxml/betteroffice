//! Font registry over raw font bytes.
//!
//! Parsing crate choice: **skrifa** (over `ttf-parser`). Rationale:
//!
//! - skrifa is the Google Fonts "oxidize" parser built for exactly this
//!   consumer profile — metrics + cmap + (later) glyph outlines for a
//!   renderer. Its outline API (`OutlineGlyph` → path pen) is what the
//!   canvas-engine glyph pipeline (`Path2D` outlines per design.md) will call
//!   next, so metrics and rasterization come from one parser by construction.
//! - First-class variable-font (`LocationRef`) and COLR support, which the
//!   display-list renderer needs for variable and color fonts.
//! - Actively developed and fuzzed as the shaping/metrics backend of the
//!   fontations stack (parley/vello); font bytes here are attacker-controlled
//!   (embedded DOCX fonts), so a hardened, panic-free parser is a requirement,
//!   not a nicety.
//!
//! Trade-off acknowledged: rustybuzz (used in [`crate::shape`]) embeds
//! `ttf-parser` internally, so both parsers end up in the dependency tree.
//! That duplication is confined to shaping; every metric this crate reports
//! comes from skrifa, so measurement can never disagree with the future
//! outline path over which parser read the table.

use std::cell::RefCell;
use std::collections::HashMap;

use skrifa::raw::TableProvider;
use skrifa::{FontRef, MetadataProvider};

use crate::shape::ShapedGlyph;

const MAX_SHAPE_CACHE_ENTRIES: usize = 4_096;
pub(crate) const MAX_CACHED_SHAPE_TEXT_BYTES: usize = 64 * 1024;
const MAX_CACHED_SHAPE_GLYPHS: usize = 64 * 1024;
const MAX_SHAPE_CACHE_TEXT_BYTES: usize = 8 * 1024 * 1024;
const MAX_SHAPE_CACHE_GLYPHS: usize = 1_000_000;
const MAX_CHAR_CACHE_ENTRIES: usize = 8_192;

const MAX_CACHED_SHAPE_FEATURES: usize = 64;

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

/// Design-space metrics extracted at registration time, in font units.
///
/// Both the `hhea` and `OS/2` variants are kept because Word derives line
/// height from `OS/2` usWinAscent/usWinDescent while typographic spacing
/// uses the sTypo values — see [`crate::word_metrics`].
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
}

struct FontEntry {
    // Declared before `data` so the borrowing parser views drop first.
    //
    // SAFETY invariants for the `'static` lifetimes below: both views borrow
    // `data`'s heap allocation. `data` is a `Box<[u8]>` that is never
    // reassigned, never mutated, and dropped only when the whole entry drops
    // (entries are never removed individually — the store only grows or drops
    // wholesale). Moving `FontEntry` (e.g. a `Vec` realloc) moves the box
    // pointer, not the heap bytes, so the views stay valid.
    face: Option<rustybuzz::Face<'static>>,
    data: Box<[u8]>,
    metrics: FontMetrics,
    char_cache: RefCell<HashMap<char, CharEntry>>,
}

/// Registry of fonts, keyed by [`FontId`], parsed from raw bytes.
///
/// The store owns the byte buffers; shaping and (later) outline extraction
/// borrow them via [`FontStore::font_bytes`].
#[derive(Default)]
pub struct FontStore {
    fonts: Vec<FontEntry>,
    shape_cache: RefCell<ShapeCache>,
}

impl FontStore {
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse and register a font from raw bytes, returning its handle.
    ///
    /// Metrics are extracted eagerly so later queries are infallible cheap
    /// lookups and malformed fonts are rejected at the trust boundary
    /// (embedded DOCX fonts are attacker-controlled).
    pub fn register(&mut self, bytes: Vec<u8>) -> Result<FontId, FontError> {
        let font = FontRef::new(&bytes).map_err(|e| FontError::Parse(e.to_string()))?;

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
        };

        let data: Box<[u8]> = bytes.into_boxed_slice();
        // SAFETY: see the `FontEntry` field invariants — the view borrows
        // `data`'s heap allocation, which outlives it and never moves.
        let pinned: &'static [u8] =
            unsafe { std::slice::from_raw_parts(data.as_ptr(), data.len()) };
        // A font rustybuzz rejects still registers (skrifa accepted it);
        // shaping then reports the parse failure per call, as before.
        let face = rustybuzz::Face::from_slice(pinned, 0);

        let id = FontId(self.fonts.len() as u32);
        self.fonts.push(FontEntry {
            face,
            data,
            metrics,
            char_cache: RefCell::new(HashMap::new()),
        });
        Ok(id)
    }

    /// Per-font design-space metrics captured at registration.
    pub fn metrics(&self, id: FontId) -> Result<&FontMetrics, FontError> {
        self.entry(id).map(|e| &e.metrics)
    }

    /// Raw bytes of a registered font (for shaping / outline extraction).
    pub fn font_bytes(&self, id: FontId) -> Result<&[u8], FontError> {
        self.entry(id).map(|e| &*e.data)
    }

    /// Parsed shaping face memoized at registration. `None` when rustybuzz
    /// rejected bytes skrifa accepted.
    pub(crate) fn shaping_face(
        &self,
        id: FontId,
    ) -> Result<Option<&rustybuzz::Face<'_>>, FontError> {
        self.entry(id).map(|e| e.face.as_ref())
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
        let entry = self.entry(id)?;
        if let Some(cached) = entry.char_cache.borrow().get(&ch) {
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
        let mut cache = entry.char_cache.borrow_mut();
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

    /// Horizontal advance width for a character, in font units.
    /// `None` if the character is not covered by this font's cmap.
    pub fn advance_width(&self, id: FontId, ch: char) -> Result<Option<f32>, FontError> {
        let entry = self.char_entry(id, ch)?;
        if entry.mapped.is_none() {
            return Ok(None);
        }
        Ok(entry.advance)
    }

    /// Whether the font's cmap covers `ch`.
    pub fn covers(&self, id: FontId, ch: char) -> Result<bool, FontError> {
        Ok(self.glyph_id(id, ch)?.is_some())
    }

    /// Fallback-chain resolution: the first font in `chain` whose cmap covers
    /// `ch` wins; `None` if no font in the chain covers it (the host then
    /// degrades that run to its browser-measured path per design.md).
    ///
    /// Unknown ids in the chain are skipped rather than failing the whole
    /// resolution: a chain is host-assembled config, and one stale entry must
    /// not take down measurement for a run another font can cover.
    pub fn resolve(&self, chain: &[FontId], ch: char) -> Option<FontId> {
        chain
            .iter()
            .copied()
            .find(|&id| self.covers(id, ch).unwrap_or(false))
    }

    fn entry(&self, id: FontId) -> Result<&FontEntry, FontError> {
        self.fonts.get(id.0 as usize).ok_or(FontError::UnknownFont)
    }

    // registration already proved the bytes parse, so this cannot fail
    fn font_ref(entry: &FontEntry) -> FontRef<'_> {
        FontRef::new(&entry.data).expect("bytes validated at register()")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shape::ShapedGlyph;

    const LIBERATION_SANS: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

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
    fn shaping_face_is_memoized_at_registration() {
        let mut store = FontStore::new();
        let font = store.register(LIBERATION_SANS.to_vec()).unwrap();
        assert!(store.shaping_face(font).unwrap().is_some());
    }
}
