//! The device-independent drawing a full metafile replay produces.

use std::sync::Arc;

pub use ooxml_drawingml::GeometryPathCommand as PathCommand;

/// A replayed metafile in its own coordinate space: `0..width` by `0..height`
/// CSS pixels, the size the metafile's frame occupies at 96 DPI.
#[derive(Debug, Clone, PartialEq)]
pub struct Drawing {
    pub width: f64,
    pub height: f64,
    pub ops: Vec<Op>,
    /// Ink the replay could not reproduce, drawn without.
    pub omissions: Vec<Omission>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Op {
    Shape(Shape),
    Text(Text),
    Image(Image),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Shape {
    pub path: Vec<PathCommand>,
    pub fill: Option<Paint>,
    pub stroke: Option<Stroke>,
    pub even_odd: bool,
    pub clip: Clip,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Rgba {
    pub r: u8,
    pub g: u8,
    pub b: u8,
    pub a: u8,
}

impl Rgba {
    pub const BLACK: Self = Self::opaque(0, 0, 0);
    pub const WHITE: Self = Self::opaque(255, 255, 255);

    pub const fn opaque(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b, a: 255 }
    }

    /// A GDI `COLORREF` (`0x00BBGGRR`).
    pub const fn from_colorref(value: u32) -> Self {
        Self::opaque(value as u8, (value >> 8) as u8, (value >> 16) as u8)
    }

    /// A GDI+ `ARGB` (`0xAARRGGBB`).
    pub const fn from_argb(value: u32) -> Self {
        Self {
            r: (value >> 16) as u8,
            g: (value >> 8) as u8,
            b: value as u8,
            a: (value >> 24) as u8,
        }
    }

    /// `#rrggbb`, ignoring alpha.
    pub fn hex(self) -> String {
        format!("#{:02x}{:02x}{:02x}", self.r, self.g, self.b)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Paint {
    Solid(Rgba),
    /// A GDI hatch: lines of `color` on `background`, or on nothing, repeating
    /// every `cell` units.
    Hatch {
        style: Hatch,
        color: Rgba,
        background: Option<Rgba>,
        cell: f64,
    },
    /// A bitmap tiled from the origin, each tile `width` by `height` units.
    Pattern {
        tile: Arc<Bitmap>,
        width: f64,
        height: f64,
    },
    Linear(Arc<LinearGradient>),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Hatch {
    Horizontal,
    Vertical,
    ForwardDiagonal,
    BackwardDiagonal,
    Cross,
    DiagonalCross,
}

#[derive(Debug, Clone, PartialEq)]
pub struct LinearGradient {
    pub start: (f64, f64),
    pub end: (f64, f64),
    /// Offsets in `0..=1` along `start..end`, ascending.
    pub stops: Arc<[(f64, Rgba)]>,
    pub spread: Spread,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Spread {
    Pad,
    Repeat,
    Reflect,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Stroke {
    pub paint: Paint,
    pub width: f64,
    /// Alternating dash and gap lengths, in drawing units.
    pub dash: Option<Vec<f64>>,
    pub cap: LineCap,
    pub join: LineJoin,
    pub miter_limit: f64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineCap {
    Butt,
    Round,
    Square,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineJoin {
    Miter,
    Round,
    Bevel,
}

/// A run of characters laid out in its own text space, where the baseline
/// runs along +x through the origin and +y points down the glyphs.
#[derive(Debug, Clone, PartialEq)]
pub struct Text {
    /// Maps text space into drawing space: `[a, b, c, d, e, f]` as in SVG.
    pub transform: [f64; 6],
    pub text: String,
    /// Each character's origin in text space, when the metafile positions
    /// every character; otherwise the run starts at the origin.
    pub positions: Option<Vec<(f64, f64)>>,
    pub anchor: TextAnchor,
    pub font: Font,
    pub fill: Paint,
    pub clip: Clip,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAnchor {
    Start,
    Middle,
    End,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Font {
    pub family: String,
    /// The em size in text-space units.
    pub size: f64,
    pub weight: u16,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Image {
    /// Maps bitmap pixels (`0..width` by `0..height`, y down) into drawing space.
    pub transform: [f64; 6],
    pub bitmap: Arc<Bitmap>,
    pub opacity: f64,
    pub clip: Clip,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bitmap {
    pub width: u32,
    pub height: u32,
    pub pixels: Pixels,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Pixels {
    /// Straight (not premultiplied) RGBA, row by row from the top.
    Rgba(Vec<u8>),
    /// Encoded image bytes a browser decodes as they are.
    Encoded { mime: &'static str, bytes: Vec<u8> },
}

/// The clip an op draws under: every region of the chain intersected, or
/// nothing clipped when `None`.
pub type Clip = Option<Arc<ClipChain>>;

#[derive(Debug, Clone, PartialEq)]
pub struct ClipChain {
    pub parent: Clip,
    pub region: ClipRegion,
}

impl ClipChain {
    pub fn depth(clip: &Clip) -> usize {
        let mut depth = 0;
        let mut at = clip.as_deref();
        while let Some(chain) = at {
            depth += 1;
            at = chain.parent.as_deref();
        }
        depth
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ClipRegion {
    pub path: Vec<PathCommand>,
    pub even_odd: bool,
    /// Keeps everything outside `path` instead of inside it.
    pub exclude: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Omission {
    pub what: &'static str,
    pub count: usize,
}

/// Why a metafile could not be replayed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Refusal(pub String);

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl std::error::Error for Refusal {}
