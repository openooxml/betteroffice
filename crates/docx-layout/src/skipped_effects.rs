//! The effects a PNG engine leaves off a page, read from its display list.

use serde_json::Value;

use crate::display_list::{
    DisplayPage, DocAttrs, ImagePrimitive, Primitive, ShapePrimitive, TextRunPrimitive,
};

/// A kind of effect a PNG export can leave off a page.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SkippedEffect {
    /// A run's classic hollow outline.
    TextOutline,
    /// The marks over or under emphasised characters.
    TextEmphasisMark,
    /// A legacy animated text effect.
    TextEffect,
    /// Modern text effects: glow, shadow, reflection, fill and outline.
    TextModernEffects,
    /// A run's classic shadow, emboss or imprint.
    TextShadow,
    /// The glyphs a tab leader repeats.
    TextLeaderGlyphs,
    /// A shape effect.
    ShapeEffects,
    /// A picture's colour filter.
    ImageFilter,
    /// A picture's colour or blur effects.
    ImageEffects,
}

impl SkippedEffect {
    /// Every kind, in the order a page reports them.
    pub const ALL: [SkippedEffect; 9] = [
        SkippedEffect::TextOutline,
        SkippedEffect::TextEmphasisMark,
        SkippedEffect::TextEffect,
        SkippedEffect::TextModernEffects,
        SkippedEffect::TextShadow,
        SkippedEffect::TextLeaderGlyphs,
        SkippedEffect::ShapeEffects,
        SkippedEffect::ImageFilter,
        SkippedEffect::ImageEffects,
    ];

    /// The kind as the `effects-skipped` diagnostic names it.
    pub fn name(self) -> &'static str {
        match self {
            SkippedEffect::TextOutline => "text.outline",
            SkippedEffect::TextEmphasisMark => "text.emphasisMark",
            SkippedEffect::TextEffect => "text.effect",
            SkippedEffect::TextModernEffects => "text.modernEffects",
            SkippedEffect::TextShadow => "text.shadow",
            SkippedEffect::TextLeaderGlyphs => "text.leaderGlyphs",
            SkippedEffect::ShapeEffects => "shape.effects",
            SkippedEffect::ImageFilter => "image.filter",
            SkippedEffect::ImageEffects => "image.effects",
        }
    }
}

/// When an effect a run carries is drawn.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Draws {
    Never,
    Always,
    /// Unless the run also carries modern text effects.
    UnlessModernEffects,
    /// Unless the run also carries a modern glow or a modern shadow.
    UnlessModernGlowOrShadow,
}

/// What a run must paint for an effect that rides on its paint to show.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Paint {
    /// Nothing: the effect paints itself.
    Always,
    /// Its fill, unless `textFill` is `none`, or a modern outline without `noFill`.
    FillOrModernOutline,
    /// Its fill, unless the modern `textFill` is `none`, or a classic or modern outline.
    FillOrOutline,
    /// As [`Paint::FillOrOutline`], or an emphasis mark of one of the listed values.
    FillOutlineOrMarks(&'static [&'static str]),
}

/// An effect drawn for the listed values only.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Values {
    pub values: &'static [&'static str],
    pub draws: Draws,
    /// What the run must paint for the effect to show.
    pub painted: Paint,
}

/// How modern text effects are drawn.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Modern {
    pub drawn: bool,
    /// Whether a modern shadow is drawn when the run also carries a modern glow.
    pub shadow_beside_glow: bool,
    /// Whether a modern outline with no fill is drawn.
    pub outline_without_fill: bool,
    /// Whether a solid or gradient modern fill is drawn beside a classic outline.
    pub fill_beside_classic_outline: bool,
    /// What the run must paint for a modern reflection to show.
    pub reflection_painted: Paint,
    /// What the run must paint for a modern glow or shadow to show.
    pub glow_and_shadow_painted: Paint,
}

/// One rule for each way a run is painted.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RunPaths<T> {
    /// A text primitive painted as text.
    pub text: T,
    /// A text primitive that paints its tab leader in place of its text.
    pub leader: T,
    /// A glyph run painted from its glyph outlines.
    pub glyph_run: T,
}

/// Whether leader glyphs are drawn, by primitive kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct LeaderGlyphs {
    pub text: bool,
    pub glyph_run: bool,
}

/// Which picture filters are drawn.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ImageFilter {
    Never,
    /// A filter at most `max_len` UTF-16 units long, made only of `functions`.
    Accepted {
        max_len: usize,
        functions: &'static [FilterFunction],
    },
}

/// One CSS filter function: its name and the argument it takes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FilterFunction(pub &'static str, pub FilterArgument);

/// The argument of a CSS filter function.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FilterArgument {
    /// A number with no unit.
    Number,
    /// A number of `px`.
    Pixels,
    /// A number with an optional `%`.
    Percentage,
    /// An optionally negative number of `deg`.
    SignedDegrees,
}

/// What one engine draws on its normal path, one entry per [`SkippedEffect`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Capabilities {
    pub text_outline: RunPaths<Draws>,
    pub text_emphasis_mark: RunPaths<Values>,
    pub text_effect: RunPaths<Draws>,
    pub text_modern_effects: RunPaths<Modern>,
    pub text_shadow: RunPaths<Values>,
    pub text_leader_glyphs: LeaderGlyphs,
    /// The first effect of one of these kinds is drawn; every other effect is not.
    pub shape_effects: &'static [&'static str],
    pub image_filter: ImageFilter,
    /// The picture effect kinds drawn.
    pub image_effects: &'static [&'static str],
}

const NEVER: RunPaths<Draws> = RunPaths {
    text: Draws::Never,
    leader: Draws::Never,
    glyph_run: Draws::Never,
};

const NOT_DRAWN: Values = Values {
    values: &[],
    draws: Draws::Never,
    painted: Paint::Always,
};

const MODERN_NOT_DRAWN: Modern = Modern {
    drawn: false,
    shadow_beside_glow: false,
    outline_without_fill: false,
    fill_beside_classic_outline: false,
    reflection_painted: Paint::Always,
    glow_and_shadow_painted: Paint::Always,
};

const CLASSIC_SHADOW: Values = Values {
    values: &["shadow"],
    draws: Draws::Always,
    painted: Paint::Always,
};

const FIRST_SHAPE_EFFECT: &[&str] = &["shadow", "outerShadow", "glow"];

const CANVAS_MARKS: &[&str] = &["dot", "underDot", "comma", "circle"];

const CANVAS_EMPHASIS_MARKS: Values = Values {
    values: CANVAS_MARKS,
    draws: Draws::Always,
    painted: Paint::Always,
};

const CANVAS_FILTER_FUNCTIONS: &[FilterFunction] = &[
    FilterFunction("blur", FilterArgument::Pixels),
    FilterFunction("brightness", FilterArgument::Number),
    FilterFunction("contrast", FilterArgument::Number),
    FilterFunction("grayscale", FilterArgument::Percentage),
    FilterFunction("hue-rotate", FilterArgument::SignedDegrees),
    FilterFunction("invert", FilterArgument::Percentage),
    FilterFunction("opacity", FilterArgument::Percentage),
    FilterFunction("saturate", FilterArgument::Number),
    FilterFunction("sepia", FilterArgument::Percentage),
];

/// The native rasteriser, `docx-raster`.
pub const NATIVE: Capabilities = Capabilities {
    // docx-raster font.rs paint_text_core, paint_leader, paint_glyph_run_core
    text_outline: NEVER,
    // docx-raster font.rs paint_text_core, paint_leader, paint_glyph_run_core
    text_emphasis_mark: RunPaths {
        text: NOT_DRAWN,
        leader: NOT_DRAWN,
        glyph_run: NOT_DRAWN,
    },
    // docx-raster font.rs paint_text_core, paint_leader, paint_glyph_run_core
    text_effect: NEVER,
    // docx-raster font.rs paint_text_core, paint_leader, paint_glyph_run_core
    text_modern_effects: RunPaths {
        text: MODERN_NOT_DRAWN,
        leader: MODERN_NOT_DRAWN,
        glyph_run: MODERN_NOT_DRAWN,
    },
    // docx-raster font.rs drawn_text_shadow (shadow.rs text_shadow); paint_leader draws none
    text_shadow: RunPaths {
        text: CLASSIC_SHADOW,
        leader: NOT_DRAWN,
        glyph_run: CLASSIC_SHADOW,
    },
    // docx-raster font.rs paint_leader; paint_glyph_run_core draws no leader
    text_leader_glyphs: LeaderGlyphs {
        text: true,
        glyph_run: false,
    },
    // docx-raster lib.rs paint_shape with shape_shadow
    shape_effects: FIRST_SHAPE_EFFECT,
    // docx-raster lib.rs paint_image
    image_filter: ImageFilter::Never,
    // docx-raster lib.rs paint_image
    image_effects: &[],
};

/// The web canvas backend, `canvasBackend.ts`.
pub const CANVAS: Capabilities = Capabilities {
    // canvasBackend.ts drawCanvasTextWithEffects or drawModernRunText, drawLeaderText, drawGlyphRun
    text_outline: RunPaths {
        text: Draws::UnlessModernEffects,
        leader: Draws::Never,
        glyph_run: Draws::Always,
    },
    // canvasBackend.ts drawTextEmphasisMarks (leader runs too), drawGlyphEmphasisMarks
    text_emphasis_mark: RunPaths {
        text: CANVAS_EMPHASIS_MARKS,
        leader: CANVAS_EMPHASIS_MARKS,
        glyph_run: CANVAS_EMPHASIS_MARKS,
    },
    // canvasBackend.ts drawTextRun, drawGlyphRun
    text_effect: NEVER,
    // canvasBackend.ts drawModernRunText, drawLeaderText, drawGlyphRun, applyModernTextShadow
    text_modern_effects: RunPaths {
        text: Modern {
            drawn: true,
            shadow_beside_glow: false,
            outline_without_fill: false,
            fill_beside_classic_outline: true,
            reflection_painted: Paint::Always,
            glow_and_shadow_painted: Paint::FillOrModernOutline,
        },
        leader: MODERN_NOT_DRAWN,
        glyph_run: Modern {
            drawn: true,
            shadow_beside_glow: false,
            outline_without_fill: true,
            fill_beside_classic_outline: false,
            reflection_painted: Paint::FillOrOutline,
            glow_and_shadow_painted: Paint::FillOutlineOrMarks(CANVAS_MARKS),
        },
    },
    // canvasBackend.ts drawCanvasTextWithEffects, drawLeaderText, drawGlyphRun
    text_shadow: RunPaths {
        text: Values {
            values: &["shadow", "emboss", "imprint"],
            draws: Draws::UnlessModernEffects,
            painted: Paint::Always,
        },
        leader: NOT_DRAWN,
        glyph_run: Values {
            values: &["shadow"],
            draws: Draws::UnlessModernGlowOrShadow,
            painted: Paint::FillOutlineOrMarks(CANVAS_MARKS),
        },
    },
    // canvasBackend.ts drawLeaderText; drawGlyphRun draws no leader
    text_leader_glyphs: LeaderGlyphs {
        text: true,
        glyph_run: false,
    },
    // canvasBackend.ts applyShapeEffects
    shape_effects: FIRST_SHAPE_EFFECT,
    // canvasBackend.ts imageFilter with sanitizeCanvasFilter
    image_filter: ImageFilter::Accepted {
        max_len: 512,
        functions: CANVAS_FILTER_FUNCTIONS,
    },
    // canvasBackend.ts imageFilter
    image_effects: &[
        "brightness",
        "contrast",
        "saturate",
        "saturation",
        "grayscale",
        "biLevel",
        "blur",
    ],
};

/// The list of a page a primitive sits in.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Region {
    /// `primitives`, watermark included.
    Body,
    Header,
    Footer,
    /// The separator primitives of the note area at this index.
    NoteSeparator(usize),
    /// The primitives of the note area at this index.
    Note(usize),
}

/// A primitive's position on its page: its list and its order in that list.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PrimitiveAt {
    pub region: Region,
    pub index: usize,
}

/// An image, or a shape with a picture fill and a path.
#[derive(Clone, Copy, Debug)]
pub struct Picture<'a> {
    pub at: PrimitiveAt,
    pub primitive: &'a Primitive,
    /// The picture's full reference: a relationship id or a whole `data:` URL.
    pub source: &'a str,
}

/// Every picture on `page`, in walk order.
pub fn pictures(page: &DisplayPage) -> Vec<Picture<'_>> {
    page_primitives(page)
        .filter_map(|(at, primitive)| picture(at, primitive))
        .collect()
}

/// The positions of the pictures on `page` whose full reference is one of `sources`.
pub fn pictures_named<S: AsRef<str>>(page: &DisplayPage, sources: &[S]) -> Vec<PrimitiveAt> {
    pictures(page)
        .into_iter()
        .filter(|picture| named(sources, picture.source))
        .map(|picture| picture.at)
        .collect()
}

fn named<S: AsRef<str>>(sources: &[S], source: &str) -> bool {
    sources.iter().any(|item| item.as_ref() == source)
}

/// Positions packed as `(region, area, index)` triples, skipping malformed ones.
pub fn unpack_positions(packed: &[u32]) -> Vec<PrimitiveAt> {
    packed
        .as_chunks::<3>()
        .0
        .iter()
        .filter_map(|triple| {
            let area = triple[1] as usize;
            let region = match triple[0] {
                0 => Region::Body,
                1 => Region::Header,
                2 => Region::Footer,
                3 => Region::NoteSeparator(area),
                4 => Region::Note(area),
                _ => return None,
            };
            Some(PrimitiveAt {
                region,
                index: triple[2] as usize,
            })
        })
        .collect()
}

/// `effects` as bits in [`SkippedEffect::ALL`] order, the first kind the lowest bit.
pub fn effect_bits(effects: &[SkippedEffect]) -> u32 {
    SkippedEffect::ALL
        .iter()
        .enumerate()
        .filter(|(_, kind)| effects.contains(*kind))
        .fold(0, |bits, (bit, _)| bits | 1 << bit)
}

/// A picture reference as `picture-unavailable` names it.
pub fn reported_id(source: &str) -> &str {
    if !source.starts_with("data:") {
        return source;
    }
    source.split_once(',').map_or(source, |(header, _)| header)
}

/// The kinds of effect `page` carries that `engine` does not draw.
pub fn skipped_effects(
    page: &DisplayPage,
    unavailable: &[PrimitiveAt],
    engine: &Capabilities,
) -> Vec<SkippedEffect> {
    let mut found = [false; SkippedEffect::ALL.len()];
    for (at, primitive) in page_primitives(page) {
        let subject = Subject::of(primitive, unavailable.contains(&at));
        for (seen, kind) in found.iter_mut().zip(SkippedEffect::ALL) {
            if !*seen {
                *seen = subject.skips(kind, engine);
            }
        }
    }
    SkippedEffect::ALL
        .into_iter()
        .zip(found)
        .filter_map(|(kind, seen)| seen.then_some(kind))
        .collect()
}

/// Every primitive a page paints, with its position.
pub fn page_primitives(page: &DisplayPage) -> impl Iterator<Item = (PrimitiveAt, &Primitive)> {
    let header = page.header.iter();
    let header = header.flat_map(|region| located(Region::Header, &region.primitives));
    let footer = page.footer.iter();
    let footer = footer.flat_map(|region| located(Region::Footer, &region.primitives));
    let areas = page.note_areas.iter().enumerate();
    let notes = areas.flat_map(|(area, notes)| {
        let separators = located(Region::NoteSeparator(area), &notes.separator_primitives);
        separators.chain(located(Region::Note(area), &notes.primitives))
    });
    let body = located(Region::Body, &page.primitives);
    body.chain(header).chain(footer).chain(notes)
}

fn located(
    region: Region,
    primitives: &[Primitive],
) -> impl Iterator<Item = (PrimitiveAt, &Primitive)> {
    primitives
        .iter()
        .enumerate()
        .map(move |(index, primitive)| (PrimitiveAt { region, index }, primitive))
}

fn picture(at: PrimitiveAt, primitive: &Primitive) -> Option<Picture<'_>> {
    let source = match primitive {
        Primitive::Image(image) => image_source(image),
        Primitive::Shape(shape) if !shape.geometry_path.is_empty() => shape_source(shape)?,
        _ => return None,
    };
    Some(Picture {
        at,
        primitive,
        source,
    })
}

/// What a primitive offers the effect rules.
enum Subject<'a> {
    Run(Run<'a>),
    Shape(&'a ShapePrimitive),
    Image(&'a ImagePrimitive),
    /// No effect, or a picture reported unavailable.
    Silent,
}

impl<'a> Subject<'a> {
    fn of(primitive: &'a Primitive, unavailable: bool) -> Self {
        match primitive {
            Primitive::Text(run) => Subject::Run(Run {
                path: text_path(run),
                outline: run.text_outline,
                emphasis_mark: run.emphasis_mark.as_deref(),
                text_effect: run.text_effect.is_some(),
                shadow: run.text_shadow.as_deref(),
                modern: modern_effects(&run.attrs),
                leader_glyphs: run.attrs.leader_glyphs.is_some(),
            }),
            Primitive::GlyphRun(run) => Subject::Run(Run {
                path: RunPath::GlyphRun,
                outline: run.text_outline,
                emphasis_mark: run.emphasis_mark.as_deref(),
                text_effect: run.text_effect.is_some(),
                shadow: run.text_shadow.as_deref(),
                modern: modern_effects(&run.attrs),
                leader_glyphs: run.attrs.leader_glyphs.is_some(),
            }),
            Primitive::Shape(_) | Primitive::Image(_) if unavailable => Subject::Silent,
            Primitive::Shape(shape) => Subject::Shape(shape),
            Primitive::Image(image) => Subject::Image(image),
            Primitive::Rect(_) | Primitive::Line(_) | Primitive::Decoration(_) => Subject::Silent,
        }
    }

    fn skips(&self, kind: SkippedEffect, rules: &Capabilities) -> bool {
        match self {
            Subject::Run(run) => run.skips(kind, rules),
            Subject::Shape(shape) => shape_skips(kind, shape, rules),
            Subject::Image(image) => image_skips(kind, image, rules),
            Subject::Silent => false,
        }
    }
}

#[derive(Clone, Copy)]
enum RunPath {
    Text,
    Leader,
    GlyphRun,
}

impl<T> RunPaths<T> {
    fn on(&self, path: RunPath) -> &T {
        match path {
            RunPath::Text => &self.text,
            RunPath::Leader => &self.leader,
            RunPath::GlyphRun => &self.glyph_run,
        }
    }
}

/// The effects a text or glyph run carries, and the path it paints on.
struct Run<'a> {
    path: RunPath,
    outline: bool,
    emphasis_mark: Option<&'a str>,
    text_effect: bool,
    shadow: Option<&'a str>,
    modern: Option<&'a Value>,
    leader_glyphs: bool,
}

impl Run<'_> {
    fn skips(&self, kind: SkippedEffect, rules: &Capabilities) -> bool {
        match kind {
            SkippedEffect::TextOutline => self.flag_skipped(self.outline, rules.text_outline),
            SkippedEffect::TextEmphasisMark => {
                self.value_skipped(self.emphasis_mark, rules.text_emphasis_mark)
            }
            SkippedEffect::TextEffect => self.flag_skipped(self.text_effect, rules.text_effect),
            SkippedEffect::TextModernEffects => self.modern_skipped(rules.text_modern_effects),
            SkippedEffect::TextShadow => self.value_skipped(self.shadow, rules.text_shadow),
            SkippedEffect::TextLeaderGlyphs => self.leader_skipped(rules.text_leader_glyphs),
            SkippedEffect::ShapeEffects
            | SkippedEffect::ImageFilter
            | SkippedEffect::ImageEffects => false,
        }
    }

    fn flag_skipped(&self, carried: bool, rule: RunPaths<Draws>) -> bool {
        carried && !self.draws(*rule.on(self.path))
    }

    fn value_skipped(&self, value: Option<&str>, rule: RunPaths<Values>) -> bool {
        let rule = rule.on(self.path);
        let Some(value) = value else {
            return false;
        };
        let drawn = listed(rule.values, value) && self.draws(rule.draws);
        !drawn || !self.paints(rule.painted)
    }

    fn modern_skipped(&self, rule: RunPaths<Modern>) -> bool {
        let rule = rule.on(self.path);
        match self.modern {
            Some(modern) => !rule.drawn || rule.drops(modern, self),
            None => false,
        }
    }

    fn paints(&self, paint: Paint) -> bool {
        let modern = self.modern;
        let outline = modern.and_then(|modern| modern.get("textOutline"));
        let fills = modern.and_then(fill_kind) != Some("none");
        match paint {
            Paint::Always => true,
            Paint::FillOrModernOutline => {
                let hollow = truthy(outline.and_then(|outline| outline.get("noFill")));
                fills || (truthy(outline) && !hollow)
            }
            Paint::FillOrOutline => fills || self.outline || truthy(outline),
            Paint::FillOutlineOrMarks(marks) => {
                let marked = self.emphasis_mark.is_some_and(|mark| listed(marks, mark));
                self.paints(Paint::FillOrOutline) || marked
            }
        }
    }

    fn leader_skipped(&self, rule: LeaderGlyphs) -> bool {
        let drawn = match self.path {
            RunPath::Text | RunPath::Leader => rule.text,
            RunPath::GlyphRun => rule.glyph_run,
        };
        self.leader_glyphs && !drawn
    }

    fn draws(&self, draws: Draws) -> bool {
        match draws {
            Draws::Never => false,
            Draws::Always => true,
            Draws::UnlessModernEffects => self.modern.is_none(),
            Draws::UnlessModernGlowOrShadow => !self.modern.is_some_and(glow_or_shadow),
        }
    }
}

impl Modern {
    /// Whether this path leaves out part of the modern effects it draws.
    fn drops(&self, modern: &Value, run: &Run<'_>) -> bool {
        let glow = truthy(modern.get("glow"));
        let outline = modern.get("textOutline");
        let hollow = truthy(outline.and_then(|outline| outline.get("noFill")));
        let shadow_hidden = glow && truthy(modern.get("shadow")) && !self.shadow_beside_glow;
        let outline_dropped = hollow && !self.outline_without_fill;
        let painted_fill = matches!(fill_kind(modern), Some("solid" | "gradient"));
        let fill_hidden = painted_fill && run.outline && !self.fill_beside_classic_outline;
        let reflected = truthy(modern.get("reflection"));
        let reflection_hidden = reflected && !run.paints(self.reflection_painted);
        let unpainted = glow_or_shadow(modern) && !run.paints(self.glow_and_shadow_painted);
        shadow_hidden || outline_dropped || fill_hidden || reflection_hidden || unpainted
    }
}

fn fill_kind(modern: &Value) -> Option<&str> {
    modern
        .get("textFill")
        .and_then(|fill| fill.get("kind"))
        .and_then(Value::as_str)
}

fn glow_or_shadow(modern: &Value) -> bool {
    truthy(modern.get("glow")) || truthy(modern.get("shadow"))
}

/// The path a text run paints on: its tab leader in place of its text, or its text.
fn text_path(run: &TextRunPrimitive) -> RunPath {
    let Some(leader) = run.attrs.leader_glyphs.as_deref() else {
        return RunPath::Text;
    };
    let glyph = leader.glyph.as_deref().unwrap_or_default();
    if glyph.is_empty() || leader.count.unwrap_or(0) == 0 {
        RunPath::Text
    } else {
        RunPath::Leader
    }
}

fn modern_effects(attrs: &DocAttrs) -> Option<&Value> {
    let modern = attrs.modern_effects.as_deref()?;
    truthy(Some(modern)).then_some(modern)
}

/// A JSON value as a JavaScript condition reads it.
fn truthy(value: Option<&Value>) -> bool {
    match value {
        None | Some(Value::Null) => false,
        Some(Value::Bool(flag)) => *flag,
        Some(Value::Number(number)) => number.as_f64().is_some_and(|number| number != 0.0),
        Some(Value::String(text)) => !text.is_empty(),
        Some(Value::Array(_) | Value::Object(_)) => true,
    }
}

fn shape_skips(kind: SkippedEffect, shape: &ShapePrimitive, rules: &Capabilities) -> bool {
    match kind {
        SkippedEffect::ShapeEffects => extra_effects(shape, rules.shape_effects),
        SkippedEffect::TextOutline
        | SkippedEffect::TextEmphasisMark
        | SkippedEffect::TextEffect
        | SkippedEffect::TextModernEffects
        | SkippedEffect::TextShadow
        | SkippedEffect::TextLeaderGlyphs
        | SkippedEffect::ImageFilter
        | SkippedEffect::ImageEffects => false,
    }
}

/// Whether `shape` carries an effect besides the first one of the kinds `drawn` lists.
fn extra_effects(shape: &ShapePrimitive, drawn: &[&str]) -> bool {
    let effects = &shape.attrs.effects;
    let first = effects
        .iter()
        .any(|effect| listed(drawn, effect_kind(effect)));
    effects.len() > usize::from(first)
}

fn image_skips(kind: SkippedEffect, image: &ImagePrimitive, rules: &Capabilities) -> bool {
    match kind {
        SkippedEffect::ImageFilter => match image.filter.as_deref() {
            Some(filter) => !filter_drawn(rules.image_filter, filter),
            None => false,
        },
        SkippedEffect::ImageEffects => undrawn_effects(image, rules.image_effects),
        SkippedEffect::TextOutline
        | SkippedEffect::TextEmphasisMark
        | SkippedEffect::TextEffect
        | SkippedEffect::TextModernEffects
        | SkippedEffect::TextShadow
        | SkippedEffect::TextLeaderGlyphs
        | SkippedEffect::ShapeEffects => false,
    }
}

fn undrawn_effects(image: &ImagePrimitive, drawn: &[&str]) -> bool {
    let effects = &image.attrs.effects;
    effects
        .iter()
        .any(|effect| !listed(drawn, effect_kind(effect)))
}

fn effect_kind(effect: &Value) -> &str {
    effect
        .get("kind")
        .and_then(Value::as_str)
        .unwrap_or_default()
}

fn listed(list: &[&str], value: &str) -> bool {
    list.iter().any(|item| *item == value)
}

fn filter_drawn(rule: ImageFilter, filter: &str) -> bool {
    match rule {
        ImageFilter::Never => false,
        ImageFilter::Accepted { max_len, functions } => {
            filter.encode_utf16().count() <= max_len && accepted_filter(filter, functions)
        }
    }
}

fn accepted_filter(filter: &str, functions: &[FilterFunction]) -> bool {
    let mut rest = filter;
    let mut any = false;
    loop {
        rest = rest.trim_start_matches(js_space);
        if rest.is_empty() {
            return any;
        }
        match functions.iter().find_map(|function| function.accept(rest)) {
            Some(after) => {
                rest = after;
                any = true;
            }
            None => return false,
        }
    }
}

impl FilterFunction {
    /// What follows this function at the start of `input`, when it starts there.
    fn accept<'i>(&self, input: &'i str) -> Option<&'i str> {
        let FilterFunction(name, argument) = *self;
        let mut rest = input.strip_prefix(name)?.strip_prefix('(')?;
        if argument == FilterArgument::SignedDegrees {
            rest = rest.strip_prefix('-').unwrap_or(rest);
        }
        rest = number(rest)?;
        rest = match argument {
            FilterArgument::Number => rest,
            FilterArgument::Pixels => rest.strip_prefix("px")?,
            FilterArgument::Percentage => rest.strip_prefix('%').unwrap_or(rest),
            FilterArgument::SignedDegrees => rest.strip_prefix("deg")?,
        };
        rest.strip_prefix(')')
    }
}

/// What follows a number, ASCII digits with an optional fraction, at the start of `input`.
fn number(input: &str) -> Option<&str> {
    let rest = digits(input)?;
    match rest.strip_prefix('.') {
        Some(fraction) => digits(fraction),
        None => Some(rest),
    }
}

fn digits(input: &str) -> Option<&str> {
    let rest = input.trim_start_matches(|character: char| character.is_ascii_digit());
    (rest.len() < input.len()).then_some(rest)
}

/// White space as a JavaScript regular expression's `\s` matches it.
fn js_space(character: char) -> bool {
    character != '\u{85}' && (character.is_whitespace() || character == '\u{feff}')
}

fn image_source(image: &ImagePrimitive) -> &str {
    match image.relationship_id.as_deref() {
        Some(relationship) if image.rel_id.is_empty() => relationship,
        _ => image.rel_id.as_str(),
    }
}

fn shape_source(shape: &ShapePrimitive) -> Option<&str> {
    let paint = shape.attrs.fill_paint.as_deref()?;
    if paint.get("kind").and_then(Value::as_str) != Some("picture") {
        return None;
    }
    let source = match paint.get("pictureSrc") {
        None | Some(Value::Null) => paint.get("pictureRelId"),
        source => source,
    };
    Some(source.and_then(Value::as_str).unwrap_or_default())
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{Value, json};

    const NONE: Vec<&str> = Vec::new();

    fn with(mut base: Value, extra: Value) -> Value {
        for (key, value) in extra.as_object().unwrap() {
            base[key.as_str()] = value.clone();
        }
        base
    }

    fn text(extra: Value) -> Value {
        let base = json!({
            "kind": "text", "text": "Run", "x": 0, "baselineY": 9, "width": 20,
            "font": "10px Arial", "color": "#000000",
        });
        with(base, extra)
    }

    fn glyph_run(extra: Value) -> Value {
        let base = json!({
            "kind": "glyphRun", "fontId": 0, "size": 10, "color": "#000000", "text": "R",
            "glyphs": [{ "id": 1, "x": 0, "y": 9, "cluster": 0, "advance": 5 }],
        });
        with(base, extra)
    }

    fn image(extra: Value) -> Value {
        let base = json!({
            "kind": "image", "relId": "rIdImage", "x": 0, "y": 0, "w": 10, "h": 10,
        });
        with(base, extra)
    }

    fn shape(extra: Value) -> Value {
        let base = json!({
            "kind": "shape", "x": 0, "y": 0, "w": 10, "h": 10, "fill": "#ff0000",
            "geometryPath": [{ "type": "move", "x": 0, "y": 0 }, { "type": "close" }],
        });
        with(base, extra)
    }

    fn page(primitives: Vec<Value>) -> DisplayPage {
        let mut page = json!({ "pageIndex": 0, "width": 100, "height": 100 });
        page["primitives"] = Value::Array(primitives);
        serde_json::from_value(page).unwrap()
    }

    fn report(
        primitives: Vec<Value>,
        unavailable: &[&str],
        rules: &Capabilities,
    ) -> Vec<&'static str> {
        let page = page(primitives);
        let unavailable = pictures_named(&page, unavailable);
        let kinds = skipped_effects(&page, &unavailable, rules);
        kinds.into_iter().map(SkippedEffect::name).collect()
    }

    fn both(primitives: Vec<Value>) -> [Vec<&'static str>; 2] {
        [
            report(primitives.clone(), &[], &NATIVE),
            report(primitives, &[], &CANVAS),
        ]
    }

    fn entries_differ(kind: SkippedEffect) -> bool {
        match kind {
            SkippedEffect::TextOutline => NATIVE.text_outline != CANVAS.text_outline,
            SkippedEffect::TextEmphasisMark => {
                NATIVE.text_emphasis_mark != CANVAS.text_emphasis_mark
            }
            SkippedEffect::TextEffect => NATIVE.text_effect != CANVAS.text_effect,
            SkippedEffect::TextModernEffects => {
                NATIVE.text_modern_effects != CANVAS.text_modern_effects
            }
            SkippedEffect::TextShadow => NATIVE.text_shadow != CANVAS.text_shadow,
            SkippedEffect::TextLeaderGlyphs => {
                NATIVE.text_leader_glyphs != CANVAS.text_leader_glyphs
            }
            SkippedEffect::ShapeEffects => NATIVE.shape_effects != CANVAS.shape_effects,
            SkippedEffect::ImageFilter => NATIVE.image_filter != CANVAS.image_filter,
            SkippedEffect::ImageEffects => NATIVE.image_effects != CANVAS.image_effects,
        }
    }

    #[test]
    fn kinds_are_named_in_reporting_order() {
        let expected = [
            "text.outline",
            "text.emphasisMark",
            "text.effect",
            "text.modernEffects",
            "text.shadow",
            "text.leaderGlyphs",
            "shape.effects",
            "image.filter",
            "image.effects",
        ];
        assert_eq!(SkippedEffect::ALL.map(SkippedEffect::name), expected);
    }

    #[test]
    fn the_engines_differ_on_exactly_these_kinds() {
        let differing: Vec<_> = SkippedEffect::ALL
            .into_iter()
            .filter(|kind| entries_differ(*kind))
            .map(SkippedEffect::name)
            .collect();
        let expected = [
            "text.outline",
            "text.emphasisMark",
            "text.modernEffects",
            "text.shadow",
            "image.filter",
            "image.effects",
        ];
        assert_eq!(differing, expected);
    }

    #[test]
    fn a_page_without_effects_reports_nothing() {
        let primitives = vec![text(json!({})), glyph_run(json!({})), image(json!({}))];
        assert_eq!(both(primitives), [NONE, NONE]);
    }

    #[test]
    fn a_classic_shadow_both_engines_draw_reports_nothing() {
        let shadow = json!({ "textShadow": "shadow" });
        let primitives = vec![text(shadow.clone()), glyph_run(shadow)];
        assert_eq!(both(primitives), [NONE, NONE]);
    }

    #[test]
    fn a_glow_beside_a_modern_shadow_drops_the_modern_shadow_on_the_canvas() {
        let modern = json!({ "modernEffects": { "glow": { "radius": 4 }, "shadow": {} } });
        let skipped = [vec!["text.modernEffects"], vec!["text.modernEffects"]];
        assert_eq!(both(vec![text(modern.clone())]), skipped);
        assert_eq!(both(vec![glyph_run(modern)]), skipped);
        let glow = json!({ "modernEffects": { "glow": { "radius": 4 } } });
        let glowing = both(vec![glyph_run(glow)]);
        assert_eq!(glowing, [vec!["text.modernEffects"], NONE]);
    }

    #[test]
    fn modern_effects_drop_a_text_runs_classic_outline_on_the_canvas() {
        let outlined = json!({ "textOutline": true, "modernEffects": { "glow": {} } });
        let native = vec!["text.outline", "text.modernEffects"];
        let text_run = both(vec![text(outlined.clone())]);
        assert_eq!(text_run, [native.clone(), vec!["text.outline"]]);
        assert_eq!(both(vec![glyph_run(outlined)]), [native, NONE]);
    }

    #[test]
    fn a_glyph_run_draws_no_emboss() {
        let embossed = json!({ "textShadow": "emboss" });
        let glyphs = both(vec![glyph_run(embossed.clone())]);
        assert_eq!(glyphs, [vec!["text.shadow"], vec!["text.shadow"]]);
        assert_eq!(both(vec![text(embossed)]), [vec!["text.shadow"], NONE]);
    }

    #[test]
    fn a_modern_glow_or_shadow_overwrites_a_glyph_runs_classic_shadow_on_the_canvas() {
        let glowing = json!({ "textShadow": "shadow", "modernEffects": { "glow": {} } });
        let skipped = [vec!["text.modernEffects"], vec!["text.shadow"]];
        assert_eq!(both(vec![glyph_run(glowing.clone())]), skipped);
        assert_eq!(both(vec![text(glowing)]), skipped);
        let mirrored = json!({ "textShadow": "shadow", "modernEffects": { "reflection": {} } });
        let glyphs = both(vec![glyph_run(mirrored)]);
        assert_eq!(glyphs, [vec!["text.modernEffects"], NONE]);
    }

    #[test]
    fn a_modern_outline_without_fill_is_dropped_on_the_canvas_text_path() {
        let hollow = json!({ "modernEffects": { "textOutline": { "noFill": true } } });
        let skipped = [vec!["text.modernEffects"], vec!["text.modernEffects"]];
        assert_eq!(both(vec![text(hollow.clone())]), skipped);
        let glyphs = both(vec![glyph_run(hollow)]);
        assert_eq!(glyphs, [vec!["text.modernEffects"], NONE]);
    }

    fn unfilled(effect: &str) -> Value {
        let mut modern = json!({ "textFill": { "kind": "none" } });
        modern[effect] = json!({});
        modern
    }

    #[test]
    fn a_classic_outline_costs_a_glyph_runs_modern_fill_on_the_canvas() {
        let fill = json!({ "textFill": { "kind": "solid", "color": "#ff0000" } });
        let outlined = json!({ "textOutline": true, "modernEffects": fill });
        let native = vec!["text.outline", "text.modernEffects"];
        let glyphs = both(vec![glyph_run(outlined.clone())]);
        assert_eq!(glyphs, [native.clone(), vec!["text.modernEffects"]]);
        let text_run = both(vec![text(outlined)]);
        assert_eq!(text_run, [native, vec!["text.outline"]]);
        let filled = json!({ "modernEffects": { "textFill": { "kind": "gradient" } } });
        let glyphs = both(vec![glyph_run(filled)]);
        assert_eq!(glyphs, [vec!["text.modernEffects"], NONE]);
    }

    #[test]
    fn an_unfilled_text_run_shows_no_glow_on_the_canvas_without_a_drawn_outline() {
        let skipped = [vec!["text.modernEffects"], vec!["text.modernEffects"]];
        for outline in [json!(null), json!({ "noFill": true })] {
            let mut modern = unfilled("glow");
            modern["textOutline"] = outline;
            let effects = json!({ "modernEffects": modern });
            assert_eq!(both(vec![text(effects)]), skipped);
        }
        let mut modern = unfilled("shadow");
        modern["textOutline"] = json!({ "width": 1 });
        let outlined = both(vec![text(json!({ "modernEffects": modern }))]);
        assert_eq!(outlined, [vec!["text.modernEffects"], NONE]);
        let effects = json!({ "modernEffects": unfilled("reflection") });
        let reflected = both(vec![text(effects)]);
        assert_eq!(reflected, [vec!["text.modernEffects"], NONE]);
    }

    #[test]
    fn an_unpainted_glyph_run_shows_no_reflection_glow_or_shadow_on_the_canvas() {
        let skipped = [vec!["text.modernEffects"], vec!["text.modernEffects"]];
        let native = vec!["text.outline", "text.modernEffects"];
        for effect in ["reflection", "glow", "shadow"] {
            let effects = json!({ "modernEffects": unfilled(effect) });
            assert_eq!(both(vec![glyph_run(effects)]), skipped);
            let outlined = json!({ "textOutline": true, "modernEffects": unfilled(effect) });
            let glyphs = both(vec![glyph_run(outlined)]);
            assert_eq!(glyphs, [native.clone(), NONE]);
        }
        let marks = vec!["text.emphasisMark", "text.modernEffects"];
        let mirror = unfilled("reflection");
        let reflected = json!({ "emphasisMark": "dot", "modernEffects": mirror });
        let glyphs = both(vec![glyph_run(reflected)]);
        assert_eq!(glyphs, [marks.clone(), vec!["text.modernEffects"]]);
        let glowing = json!({ "emphasisMark": "dot", "modernEffects": unfilled("glow") });
        assert_eq!(both(vec![glyph_run(glowing)]), [marks, NONE]);
    }

    #[test]
    fn an_unpainted_glyph_run_shows_no_classic_shadow_on_the_canvas() {
        let bare = json!({ "textFill": { "kind": "none" } });
        let shadowed = json!({ "textShadow": "shadow", "modernEffects": bare });
        let glyphs = both(vec![glyph_run(shadowed.clone())]);
        assert_eq!(glyphs, [vec!["text.modernEffects"], vec!["text.shadow"]]);
        let marks = vec!["text.emphasisMark", "text.modernEffects"];
        let marked = with(shadowed.clone(), json!({ "emphasisMark": "dot" }));
        let glyphs = both(vec![glyph_run(marked)]);
        assert_eq!(glyphs, [marks, NONE]);
        let outlines = vec!["text.outline", "text.modernEffects"];
        let outlined = with(shadowed, json!({ "textOutline": true }));
        let glyphs = both(vec![glyph_run(outlined)]);
        assert_eq!(glyphs, [outlines, NONE]);
    }

    #[test]
    fn a_text_run_that_paints_its_leader_drops_its_effects() {
        let leader = json!({ "glyph": ".", "count": 3, "advance": 8 });
        let effects = json!({
            "leaderGlyphs": leader, "textShadow": "shadow", "textOutline": true,
            "emphasisMark": "dot", "modernEffects": { "glow": {} },
        });
        let native = vec![
            "text.outline",
            "text.emphasisMark",
            "text.modernEffects",
            "text.shadow",
        ];
        let canvas = vec!["text.outline", "text.modernEffects", "text.shadow"];
        assert_eq!(both(vec![text(effects)]), [native, canvas]);
        let idle = json!({
            "leaderGlyphs": { "glyph": ".", "count": 0 }, "textShadow": "shadow",
        });
        assert_eq!(both(vec![text(idle)]), [NONE, NONE]);
    }

    #[test]
    fn leader_glyphs_on_a_glyph_run_are_not_drawn() {
        let leader = json!({ "leaderGlyphs": { "glyph": ".", "count": 2 } });
        let skipped = [vec!["text.leaderGlyphs"], vec!["text.leaderGlyphs"]];
        assert_eq!(both(vec![glyph_run(leader)]), skipped);
    }

    #[test]
    fn a_legacy_text_effect_is_never_drawn() {
        let effect = json!({ "textEffect": "shimmer" });
        let primitives = vec![text(effect.clone()), glyph_run(effect)];
        let skipped = [vec!["text.effect"], vec!["text.effect"]];
        assert_eq!(both(primitives), skipped);
    }

    #[test]
    fn the_canvas_draws_the_emphasis_marks_it_knows() {
        let dot = json!({ "emphasisMark": "dot" });
        let primitives = vec![text(dot.clone()), glyph_run(dot)];
        assert_eq!(both(primitives), [vec!["text.emphasisMark"], NONE]);
        let unknown = json!({ "emphasisMark": "none" });
        let skipped = [vec!["text.emphasisMark"], vec!["text.emphasisMark"]];
        assert_eq!(both(vec![text(unknown)]), skipped);
    }

    #[test]
    fn a_shapes_first_shadow_or_glow_is_drawn_and_no_other_effect_is() {
        let shadow = json!({ "effects": [{ "kind": "outerShadow" }] });
        assert_eq!(both(vec![shape(shadow)]), [NONE, NONE]);
        let skipped = [vec!["shape.effects"], vec!["shape.effects"]];
        for effects in [
            json!([{ "kind": "outerShadow" }, { "kind": "softEdge" }]),
            json!([{ "kind": "glow" }, { "kind": "shadow" }]),
            json!([{ "kind": "reflection" }]),
        ] {
            let effects = json!({ "effects": effects });
            assert_eq!(both(vec![shape(effects)]), skipped);
        }
    }

    #[test]
    fn the_canvas_draws_the_picture_filters_and_effects_it_knows() {
        let filtered = json!({ "filter": "brightness(1.4) contrast(0.4)" });
        assert_eq!(both(vec![image(filtered)]), [vec!["image.filter"], NONE]);
        let hostile = json!({ "filter": "url(#washout)" });
        let skipped = [vec!["image.filter"], vec!["image.filter"]];
        assert_eq!(both(vec![image(hostile)]), skipped);
        let bright = json!({ "effects": [{ "kind": "brightness", "amount": 1.2 }] });
        assert_eq!(both(vec![image(bright)]), [vec!["image.effects"], NONE]);
        let duotone = json!({ "effects": [{ "kind": "brightness" }, { "kind": "duotone" }] });
        let skipped = [vec!["image.effects"], vec!["image.effects"]];
        assert_eq!(both(vec![image(duotone)]), skipped);
    }

    #[test]
    fn an_unavailable_picture_reports_nothing_else() {
        let carried = json!({ "filter": "url(#a)", "effects": [{ "kind": "duotone" }] });
        let filtered = image(carried);
        let source = "data:image/png;base64,iVBORw0KGgo=";
        let fill = json!({ "kind": "picture", "pictureSrc": source });
        let effects = json!([{ "kind": "outerShadow" }, { "kind": "softEdge" }]);
        let filled = shape(json!({ "fillPaint": fill, "effects": effects }));
        let unavailable = ["rIdImage", source];
        for rules in [&NATIVE, &CANVAS] {
            let primitives = vec![filtered.clone(), filled.clone()];
            assert_eq!(report(primitives, &unavailable, rules), NONE);
        }
        let skipped = vec!["shape.effects", "image.filter", "image.effects"];
        assert_eq!(report(vec![filtered, filled], &[], &CANVAS), skipped);
        let named = json!({ "relId": "", "relationshipId": "rId9", "filter": "url(#a)" });
        assert_eq!(report(vec![image(named)], &["rId9"], &NATIVE), NONE);
    }

    #[test]
    fn two_pictures_of_one_type_are_told_apart_by_position() {
        let broken = image(json!({ "relId": "data:image/png;base64,AAAA" }));
        let filter = json!({ "relId": "data:image/png;base64,iVBORw0KGgo=", "filter": "url(#a)" });
        let page = page(vec![broken, image(filter)]);
        let found = pictures(&page);
        assert_eq!(found.len(), 2);
        assert_eq!(reported_id(found[0].source), reported_id(found[1].source));
        let first = PrimitiveAt {
            region: Region::Body,
            index: 0,
        };
        let named = pictures_named(&page, &["data:image/png;base64,AAAA"]);
        assert_eq!(named, [first]);
        let kinds = skipped_effects(&page, &[first], &NATIVE);
        assert_eq!(kinds, [SkippedEffect::ImageFilter]);
        let neither = skipped_effects(&page, &[first, found[1].at], &NATIVE);
        assert!(neither.is_empty());
    }

    #[test]
    fn positions_unpack_from_triples_and_kinds_pack_into_bits() {
        let packed = [0, 9, 2, 4, 1, 0, 7, 0, 0, 1];
        let note = PrimitiveAt {
            region: Region::Note(1),
            index: 0,
        };
        let body = PrimitiveAt {
            region: Region::Body,
            index: 2,
        };
        assert_eq!(unpack_positions(&packed), [body, note]);
        let effects = [SkippedEffect::TextOutline, SkippedEffect::ImageEffects];
        assert_eq!(effect_bits(&effects), 1 | 1 << 8);
        assert_eq!(effect_bits(&SkippedEffect::ALL), (1 << 9) - 1);
    }

    #[test]
    fn a_position_with_no_primitive_is_ignored() {
        let page = page(vec![image(json!({ "filter": "url(#a)" }))]);
        let nowhere = [
            PrimitiveAt {
                region: Region::Body,
                index: 7,
            },
            PrimitiveAt {
                region: Region::Note(3),
                index: 0,
            },
        ];
        let kinds = skipped_effects(&page, &nowhere, &NATIVE);
        assert_eq!(kinds, [SkippedEffect::ImageFilter]);
    }

    #[test]
    fn pictures_are_numbered_within_their_list() {
        let mut page = json!({
            "pageIndex": 0, "width": 100, "height": 100, "primitives": [],
        });
        let header = json!({
            "rId": "rIdHeader", "kind": "header", "y": 0, "height": 10,
            "primitives": [text(json!({})), image(json!({}))],
        });
        page["header"] = header;
        page["noteAreas"] = json!([{ "primitives": [image(json!({}))] }]);
        let page: DisplayPage = serde_json::from_value(page).unwrap();
        let found: Vec<_> = pictures(&page).iter().map(|picture| picture.at).collect();
        let header = PrimitiveAt {
            region: Region::Header,
            index: 1,
        };
        let note = PrimitiveAt {
            region: Region::Note(0),
            index: 0,
        };
        assert_eq!(found, [header, note]);
    }

    #[test]
    fn a_degenerate_primitive_still_reports_its_effect() {
        let crop = json!({ "top": 0, "right": 0.5, "bottom": 0, "left": 0.5 });
        let cropped = json!({ "crop": crop, "filter": "brightness(1.4)" });
        assert_eq!(both(vec![image(cropped)]), [vec!["image.filter"], NONE]);
        let blank = json!({ "text": "   ", "emphasisMark": "dot" });
        assert_eq!(both(vec![text(blank)]), [vec!["text.emphasisMark"], NONE]);
        let leader = json!({ "glyph": ".", "count": 4, "advance": 0 });
        let flat = json!({ "leaderGlyphs": leader, "textShadow": "shadow" });
        let skipped = [vec!["text.shadow"], vec!["text.shadow"]];
        assert_eq!(both(vec![text(flat)]), skipped);
    }

    #[test]
    fn kinds_come_in_declared_order() {
        let picture = json!({ "filter": "url(#a)", "effects": [{ "kind": "duotone" }] });
        let leader = json!({ "leaderGlyphs": { "glyph": ".", "count": 1 } });
        let glyphs = json!({ "emphasisMark": "dot", "modernEffects": { "glow": {} } });
        let primitives = vec![
            image(picture),
            shape(json!({ "effects": [{ "kind": "softEdge" }] })),
            glyph_run(leader),
            text(json!({ "textShadow": "emboss", "textEffect": "lights" })),
            glyph_run(glyphs),
            text(json!({ "textOutline": true })),
        ];
        let names = SkippedEffect::ALL.map(SkippedEffect::name);
        assert_eq!(report(primitives.clone(), &[], &NATIVE), names);
        let canvas = vec![
            "text.effect",
            "text.leaderGlyphs",
            "shape.effects",
            "image.filter",
            "image.effects",
        ];
        assert_eq!(report(primitives, &[], &CANVAS), canvas);
    }

    #[test]
    fn header_footer_and_note_primitives_are_walked() {
        let mut page = json!({
            "pageIndex": 0, "width": 100, "height": 100, "primitives": [],
        });
        let effect = text(json!({ "textEffect": "lights" }));
        let outline = text(json!({ "textOutline": true }));
        let shadow = glyph_run(json!({ "textShadow": "imprint" }));
        page["footer"] = json!({
            "rId": "rIdFooter", "kind": "footer", "y": 90, "height": 10, "primitives": [effect],
        });
        page["noteAreas"] = json!([
            { "separatorPrimitives": [outline], "primitives": [shadow] },
        ]);
        let page: DisplayPage = serde_json::from_value(page).unwrap();
        let kinds = skipped_effects(&page, &[], &NATIVE);
        let names: Vec<_> = kinds.into_iter().map(SkippedEffect::name).collect();
        assert_eq!(names, ["text.outline", "text.effect", "text.shadow"]);
    }

    #[test]
    fn the_canvas_filter_rule_matches_the_canvas_sanitizer() {
        let drawn = |filter: &str| filter_drawn(CANVAS.image_filter, filter);
        let longest = "blur(1px) ".repeat(51);
        for accepted in [
            "blur(2px)",
            " brightness(1.4)  contrast(0.4) ",
            "hue-rotate(-30deg)",
            "grayscale(50%)",
            "grayscale(0.5)sepia(1)",
            "invert(10%) opacity(80%) saturate(2)",
            "\u{feff}blur(1px)",
            longest.as_str(),
        ] {
            assert!(drawn(accepted), "{accepted}");
        }
        let too_long = "blur(1px) ".repeat(52);
        for rejected in [
            "",
            "   ",
            "blur(2)",
            "brightness(-1)",
            "brightness(1.)",
            "Blur(2px)",
            "url(#washout)",
            "blur(1px) drop-shadow(1px 1px red)",
            "\u{85}blur(1px)",
            too_long.as_str(),
        ] {
            assert!(!drawn(rejected), "{rejected}");
        }
        assert!(!filter_drawn(NATIVE.image_filter, "blur(2px)"));
    }
}
