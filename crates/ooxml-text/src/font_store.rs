use skrifa::raw::TableProvider;
use skrifa::{FontRef, MetadataProvider};

/// Font handle within a [`FontStore`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FontId(pub(crate) u32);

impl FontId {
    /// Return the raw index.
    pub fn to_u32(self) -> u32 {
        self.0
    }

    /// Construct a handle from a raw index.
    pub fn from_u32(raw: u32) -> Self {
        FontId(raw)
    }
}

/// Errors produced when registering or querying fonts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FontError {
    /// The bytes are not a parseable sfnt font.
    Parse(String),
    /// A required font table is missing.
    MissingTable(&'static str),
    /// The [`FontId`] does not belong to this store.
    UnknownFont,
    /// The glyph id is out of range.
    GlyphOutOfRange(u16),
    /// The glyph outline could not be read.
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

/// Font metrics in design units.
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

/// Registry of parsed fonts.
#[derive(Default)]
pub struct FontStore {
    fonts: Vec<FontEntry>,
}

impl FontStore {
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse and register a font.
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

    /// Raw bytes of a registered font.
    pub fn font_bytes(&self, id: FontId) -> Result<&[u8], FontError> {
        self.entry(id).map(|e| e.data.as_slice())
    }

    /// Look up a character's glyph id.
    pub fn glyph_id(&self, id: FontId, ch: char) -> Result<Option<u16>, FontError> {
        let entry = self.entry(id)?;
        let font = Self::font_ref(entry);
        Ok(font
            .charmap()
            .map(ch)
            .map(|g| g.to_u32() as u16)
            .filter(|&g| g != 0))
    }

    /// Return a character's horizontal advance.
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

    pub fn resolve(&self, chain: &[FontId], ch: char) -> Option<FontId> {
        chain
            .iter()
            .copied()
            .find(|&id| self.covers(id, ch).unwrap_or(false))
    }

    fn entry(&self, id: FontId) -> Result<&FontEntry, FontError> {
        self.fonts.get(id.0 as usize).ok_or(FontError::UnknownFont)
    }

    fn font_ref(entry: &FontEntry) -> FontRef<'_> {
        FontRef::new(&entry.data).expect("bytes validated at register()")
    }
}
