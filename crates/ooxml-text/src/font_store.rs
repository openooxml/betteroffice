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
use std::collections::{HashMap, HashSet, hash_map::DefaultHasher};
use std::hash::{Hash, Hasher};

use skrifa::raw::TableProvider;
use skrifa::{FontRef, MetadataProvider};

use crate::shape::ShapedGlyph;

const MAX_SHAPE_CACHE_ENTRIES: usize = 4_096;
const MAX_SHAPE_CACHE_ADMISSIONS: usize = 4_096;
pub(crate) const MAX_CACHED_SHAPE_TEXT_BYTES: usize = 64 * 1024;
const MAX_CACHED_SHAPE_GLYPHS: usize = 64 * 1024;
const MAX_SHAPE_CACHE_TEXT_BYTES: usize = 8 * 1024 * 1024;
const MAX_SHAPE_CACHE_GLYPHS: usize = 1_000_000;

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
pub(crate) struct ShapeCacheKey {
    pub font: FontId,
    pub text: String,
    pub size_bits: u32,
    pub features: Vec<([u8; 4], u32)>,
    pub direction: u8,
    pub language: Option<String>,
}

#[derive(Default)]
struct ShapeCache {
    entries: HashMap<ShapeCacheKey, Vec<ShapedGlyph>>,
    admissions: HashSet<u64>,
    text_bytes: usize,
    glyphs: usize,
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
    data: Vec<u8>,
    metrics: FontMetrics,
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

        let id = FontId(self.fonts.len() as u32);
        self.fonts.push(FontEntry {
            data: bytes,
            metrics,
        });
        Ok(id)
    }

    /// Per-font design-space metrics captured at registration.
    pub fn metrics(&self, id: FontId) -> Result<&FontMetrics, FontError> {
        self.entry(id).map(|e| &e.metrics)
    }

    /// Raw bytes of a registered font (for shaping / outline extraction).
    pub fn font_bytes(&self, id: FontId) -> Result<&[u8], FontError> {
        self.entry(id).map(|e| e.data.as_slice())
    }

    pub(crate) fn cached_shape(&self, key: &ShapeCacheKey) -> Option<Vec<ShapedGlyph>> {
        self.shape_cache.borrow().entries.get(key).cloned()
    }

    pub(crate) fn cache_shape(&self, key: ShapeCacheKey, glyphs: &[ShapedGlyph]) {
        if key.text.len() > MAX_CACHED_SHAPE_TEXT_BYTES || glyphs.len() > MAX_CACHED_SHAPE_GLYPHS {
            return;
        }
        let mut cache = self.shape_cache.borrow_mut();
        let mut hasher = DefaultHasher::new();
        key.hash(&mut hasher);
        let admission = hasher.finish();
        if !cache.admissions.remove(&admission) {
            if cache.admissions.len() >= MAX_SHAPE_CACHE_ADMISSIONS {
                cache.admissions.clear();
            }
            cache.admissions.insert(admission);
            return;
        }
        if let Some(replaced) = cache.entries.remove(&key) {
            cache.text_bytes = cache.text_bytes.saturating_sub(key.text.len());
            cache.glyphs = cache.glyphs.saturating_sub(replaced.len());
        }
        if cache.entries.len() >= MAX_SHAPE_CACHE_ENTRIES
            || cache.text_bytes.saturating_add(key.text.len()) > MAX_SHAPE_CACHE_TEXT_BYTES
            || cache.glyphs.saturating_add(glyphs.len()) > MAX_SHAPE_CACHE_GLYPHS
        {
            cache.entries.clear();
            cache.text_bytes = 0;
            cache.glyphs = 0;
        }
        cache.text_bytes += key.text.len();
        cache.glyphs += glyphs.len();
        cache.entries.insert(key, glyphs.to_vec());
    }

    #[cfg(test)]
    pub(crate) fn shape_cache_len(&self) -> usize {
        self.shape_cache.borrow().entries.len()
    }

    /// cmap lookup: glyph id for a character, `None` if the font does not
    /// cover it. Glyph id 0 (`.notdef`) counts as no coverage.
    pub fn glyph_id(&self, id: FontId, ch: char) -> Result<Option<u16>, FontError> {
        let entry = self.entry(id)?;
        let font = Self::font_ref(entry);
        Ok(font
            .charmap()
            .map(ch)
            .map(|g| g.to_u32() as u16)
            .filter(|&g| g != 0))
    }

    /// Horizontal advance width for a character, in font units.
    /// `None` if the character is not covered by this font's cmap.
    pub fn advance_width(&self, id: FontId, ch: char) -> Result<Option<f32>, FontError> {
        let entry = self.entry(id)?;
        let font = Self::font_ref(entry);
        let Some(gid) = font.charmap().map(ch) else {
            return Ok(None);
        };
        let metrics = font.glyph_metrics(
            skrifa::instance::Size::unscaled(),
            skrifa::instance::LocationRef::default(),
        );
        Ok(metrics.advance_width(gid))
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
