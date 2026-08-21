use std::collections::{BTreeMap, HashMap};

use anyhow::{Context, Result, bail};
use betteroffice_pptx::{
    GradientType, Paint, PositionedGlyph, PositionedTextLine, PositionedTextRun, PptxPackage,
    Presentation, Primitive, Stroke as DisplayStroke, SurfaceDisplayList, Transform,
};
use ooxml_drawingml::GeometryPathCommand;
use serde_json::{Value, json};
use vello::kurbo::{Affine, BezPath, Rect, Stroke};
use vello::peniko::{
    Blob, Color, Fill, FontData, Gradient, ImageAlphaType, ImageBrush, ImageData, ImageFormat,
};
use vello::{Glyph, Scene};

use crate::scene_shared::{PageScene, SkipStats, color, draw_transformed_placeholder, with_dashes};

const CARET_DRIFT_TOLERANCE_PX: f32 = 0.05;
const ARIAL_REGULAR: &[u8] =
    include_bytes!("../../../packages/fonts/assets/LiberationSans-Regular.ttf");
const ARIAL_BOLD: &[u8] = include_bytes!("../../../packages/fonts/assets/LiberationSans-Bold.ttf");
const ARIAL_ITALIC: &[u8] =
    include_bytes!("../../../packages/fonts/assets/LiberationSans-Italic.ttf");
const ARIAL_BOLD_ITALIC: &[u8] =
    include_bytes!("../../../packages/fonts/assets/LiberationSans-BoldItalic.ttf");
const FONT_FACES: &[(&str, bool, bool, &[u8])] = &[
    ("Arial", false, false, ARIAL_REGULAR),
    ("Arial", true, false, ARIAL_BOLD),
    ("Arial", false, true, ARIAL_ITALIC),
    ("Arial", true, true, ARIAL_BOLD_ITALIC),
];

pub struct PptxSceneResources {
    fonts: PptxFonts,
    images: PptxImages,
    pub font_faces: Vec<String>,
}

#[derive(Default)]
pub struct PptxSlideSummary {
    primitives: BTreeMap<String, PrimitiveCounts>,
    pub glyph_audit: GlyphAudit,
}

#[derive(Default)]
struct PrimitiveCounts {
    total: usize,
    translated: usize,
    skipped: usize,
}

#[derive(Default)]
pub struct GlyphAudit {
    pub glyph_runs: usize,
    pub glyphs: usize,
    pub caret_stops_checked: usize,
    pub missing_caret_stops: usize,
    pub drifted_caret_stops: usize,
    pub max_caret_drift_px: f32,
    pub width_checks: usize,
    pub drifted_widths: usize,
    pub max_width_drift_px: f32,
    drift_samples: Vec<String>,
}

impl PptxSlideSummary {
    pub fn structured(&self, skipped: &SkipStats) -> Value {
        let primitives = self
            .primitives
            .iter()
            .filter(|(kind, _)| kind.as_str() != "background")
            .map(|(kind, counts)| {
                (
                    kind.clone(),
                    json!({
                        "total": counts.total,
                        "translated": counts.translated,
                        "skipped": counts.skipped,
                    }),
                )
            })
            .collect::<BTreeMap<_, _>>();
        let background = self.primitives.get("background").map_or_else(
            || json!({ "present": false, "translated": false, "skipped": false }),
            |counts| {
                json!({
                    "present": true,
                    "translated": counts.translated == 1,
                    "skipped": counts.skipped == 1,
                })
            },
        );
        let total = self
            .primitives
            .iter()
            .filter(|(kind, _)| kind.as_str() != "background")
            .map(|(_, counts)| counts.total)
            .sum::<usize>();
        let translated = self
            .primitives
            .iter()
            .filter(|(kind, _)| kind.as_str() != "background")
            .map(|(_, counts)| counts.translated)
            .sum::<usize>();
        let skipped_total = self
            .primitives
            .iter()
            .filter(|(kind, _)| kind.as_str() != "background")
            .map(|(_, counts)| counts.skipped)
            .sum::<usize>();
        json!({
            "background": background,
            "primitives": primitives,
            "totals": {
                "total": total,
                "translated": translated,
                "skipped": skipped_total,
            },
            "skipReasons": &skipped.reasons,
            "glyphAudit": {
                "glyphRuns": self.glyph_audit.glyph_runs,
                "glyphs": self.glyph_audit.glyphs,
                "caretStopsChecked": self.glyph_audit.caret_stops_checked,
                "missingCaretStops": self.glyph_audit.missing_caret_stops,
                "driftedCaretStops": self.glyph_audit.drifted_caret_stops,
                "maxCaretDriftPx": self.glyph_audit.max_caret_drift_px,
                "widthChecks": self.glyph_audit.width_checks,
                "driftedWidths": self.glyph_audit.drifted_widths,
                "maxWidthDriftPx": self.glyph_audit.max_width_drift_px,
                "driftSamples": &self.glyph_audit.drift_samples,
                "tolerancePx": 0.05,
            },
        })
    }

    fn seen(&mut self, kind: &str) {
        self.primitives.entry(kind.to_owned()).or_default().total += 1;
    }

    fn translated(&mut self, kind: &str) {
        self.primitives
            .entry(kind.to_owned())
            .or_default()
            .translated += 1;
    }

    fn skipped(&mut self, kind: &str) {
        self.primitives.entry(kind.to_owned()).or_default().skipped += 1;
    }
}

impl PptxSceneResources {
    pub fn new(presentation: &mut Presentation) -> Result<Self> {
        let fonts = PptxFonts::load(|family, bold, italic, bytes| {
            presentation
                .register_font(family, bold, italic, bytes)
                .map_err(anyhow::Error::from)
        })?;
        let images = PptxImages::load(presentation.package());
        let font_faces = fonts.labels.clone();
        Ok(Self {
            fonts,
            images,
            font_faces,
        })
    }

    pub fn translate(
        &self,
        display_list: &SurfaceDisplayList,
        max_texture_dimension_2d: u32,
    ) -> Result<(PageScene, PptxSlideSummary)> {
        translate_slide(
            display_list,
            &self.fonts,
            &self.images,
            max_texture_dimension_2d,
        )
    }
}

struct PptxFonts {
    vello: HashMap<u32, FontData>,
    labels: Vec<String>,
}

impl PptxFonts {
    fn load(mut register: impl FnMut(&str, bool, bool, &[u8]) -> Result<u32>) -> Result<Self> {
        let mut vello = HashMap::new();
        let mut labels = Vec::new();
        for &(family, bold, italic, bytes) in FONT_FACES {
            let layout_id = register(family, bold, italic, bytes)
                .with_context(|| format!("register PPTX layout font {family}"))?;
            vello.insert(layout_id, FontData::new(Blob::from(bytes.to_vec()), 0));
            labels.push(format!(
                "{family}{}{}",
                if bold { " bold" } else { "" },
                if italic { " italic" } else { "" }
            ));
        }
        Ok(Self { vello, labels })
    }

    fn vello_face(&self, id: u32) -> Result<&FontData, String> {
        self.vello
            .get(&id)
            .ok_or_else(|| format!("font id {id} was not registered"))
    }
}

enum PptxImage {
    Decoded(ImageData),
    Failed(String),
}

struct PptxImages {
    assets: HashMap<String, PptxImage>,
}

impl PptxImages {
    fn load(package: &PptxPackage) -> Self {
        let assets = package
            .media
            .iter()
            .map(|part| {
                let image = image::load_from_memory(&part.bytes)
                    .map(|image| {
                        let rgba = image.into_rgba8();
                        let (width, height) = rgba.dimensions();
                        ImageData {
                            data: Blob::from(rgba.into_raw()),
                            format: ImageFormat::Rgba8,
                            alpha_type: ImageAlphaType::Alpha,
                            width,
                            height,
                        }
                    })
                    .map(PptxImage::Decoded)
                    .unwrap_or_else(|error| {
                        PptxImage::Failed(format!(
                            "image {} ({}) is undecodable: {error}",
                            part.part_path, part.content_type
                        ))
                    });
                (part.part_path.clone(), image)
            })
            .collect();
        Self { assets }
    }

    fn get(&self, asset_id: &str) -> Result<&ImageData, String> {
        match self.assets.get(asset_id) {
            Some(PptxImage::Decoded(image)) => Ok(image),
            Some(PptxImage::Failed(reason)) => Err(reason.clone()),
            None => Err(format!("image asset {asset_id} is missing")),
        }
    }
}

fn translate_slide(
    display_list: &SurfaceDisplayList,
    fonts: &PptxFonts,
    images: &PptxImages,
    max_texture_dimension_2d: u32,
) -> Result<(PageScene, PptxSlideSummary)> {
    let width = dimension(display_list.width, "slide width", max_texture_dimension_2d)?;
    let height = dimension(
        display_list.height,
        "slide height",
        max_texture_dimension_2d,
    )?;
    let mut translator = Translator {
        scene: Scene::new(),
        fonts,
        images,
        skipped: SkipStats::default(),
        summary: PptxSlideSummary::default(),
    };
    let slide_bounds = Rect::new(0.0, 0.0, width, height);
    match &display_list.background {
        Some(paint) => {
            translator.summary.seen("background");
            match prepare_paint(paint, slide_bounds) {
                Ok(paint) => {
                    paint.fill(&mut translator.scene, Affine::IDENTITY, &slide_bounds);
                    translator.summary.translated("background");
                }
                Err(reason) => {
                    translator.summary.skipped("background");
                    translator.skipped.record("background", reason);
                    draw_transformed_placeholder(
                        &mut translator.scene,
                        Some(slide_bounds),
                        Affine::IDENTITY,
                    );
                }
            }
        }
        None => translator.scene.fill(
            Fill::NonZero,
            Affine::IDENTITY,
            Color::WHITE,
            None,
            &slide_bounds,
        ),
    }
    for primitive in &display_list.primitives {
        translator.translate_primitive(primitive, Affine::IDENTITY);
    }
    Ok((
        PageScene {
            background: Scene::new(),
            scene: translator.scene,
            width,
            height,
            skipped: translator.skipped,
        },
        translator.summary,
    ))
}

struct Translator<'a> {
    scene: Scene,
    fonts: &'a PptxFonts,
    images: &'a PptxImages,
    skipped: SkipStats,
    summary: PptxSlideSummary,
}

impl Translator<'_> {
    fn translate_primitive(&mut self, primitive: &Primitive, parent: Affine) {
        if let Primitive::Chart {
            x,
            y,
            w,
            h,
            primitives,
            transform,
            ..
        } = primitive
        {
            self.translate_chart(Frame::new(*x, *y, *w, *h), primitives, *transform, parent);
            return;
        }
        let kind = primitive_kind(primitive);
        self.summary.seen(kind);
        let result = match primitive {
            Primitive::Shape {
                x,
                y,
                w,
                h,
                path,
                fill,
                stroke,
                transform,
                ..
            } => Frame::new(*x, *y, *w, *h).and_then(|frame| {
                self.draw_shape(
                    frame,
                    path,
                    fill.as_ref(),
                    stroke.as_ref(),
                    *transform,
                    parent,
                )
            }),
            Primitive::Image {
                x,
                y,
                w,
                h,
                asset_id,
                stroke,
                transform,
                ..
            } => Frame::new(*x, *y, *w, *h).and_then(|frame| {
                self.draw_image(
                    frame,
                    asset_id.as_deref(),
                    stroke.as_ref(),
                    *transform,
                    parent,
                )
            }),
            Primitive::TextBox {
                x,
                y,
                w,
                h,
                paragraphs,
                lines,
                transform,
                ..
            } => Frame::new(*x, *y, *w, *h).and_then(|frame| {
                self.draw_text_box(
                    frame,
                    paragraphs
                        .iter()
                        .flat_map(|paragraph| &paragraph.runs)
                        .any(|run| !run.text.is_empty()),
                    lines,
                    *transform,
                    parent,
                )
            }),
            Primitive::Placeholder { label, name, .. } => Err(format!(
                "placeholder primitive is not translated: {}",
                label.as_deref().unwrap_or(name)
            )),
            Primitive::Chart { .. } => unreachable!(),
        };
        match result {
            Ok(()) => self.summary.translated(kind),
            Err(reason) => self.skip(primitive, parent, kind, reason),
        }
    }

    fn translate_chart(
        &mut self,
        frame: Result<Frame, String>,
        primitives: &[Primitive],
        transform: Transform,
        parent: Affine,
    ) {
        let kind = "chart";
        self.summary.seen(kind);
        let placeholder = frame.as_ref().ok().map(|frame| frame.rect);
        let prepared = frame.and_then(|frame| {
            let affine = frame.transform(transform, parent)?;
            Ok((frame, affine))
        });
        let (frame, affine) = match prepared {
            Ok(prepared) => prepared,
            Err(reason) => {
                self.summary.skipped(kind);
                self.skipped.record(kind, reason.clone());
                draw_transformed_placeholder(&mut self.scene, placeholder, parent);
                for primitive in primitives {
                    self.skip_subtree(primitive, parent, format!("parent chart skipped: {reason}"));
                }
                return;
            }
        };
        self.summary.translated(kind);
        self.scene
            .push_clip_layer(Fill::NonZero, affine, &frame.rect);
        for primitive in primitives {
            self.translate_primitive(primitive, affine);
        }
        self.scene.pop_layer();
    }

    fn draw_shape(
        &mut self,
        frame: Frame,
        commands: &[GeometryPathCommand],
        fill: Option<&Paint>,
        stroke: Option<&DisplayStroke>,
        transform: Transform,
        parent: Affine,
    ) -> Result<(), String> {
        let path = build_path(commands, frame)?;
        if path.is_empty() && (fill.is_some() || stroke.is_some()) {
            return Err("shape path is empty".to_owned());
        }
        let fill = fill
            .map(|paint| prepare_paint(paint, frame.rect))
            .transpose()?;
        let stroke = stroke.map(prepare_stroke).transpose()?;
        let affine = frame.transform(transform, parent)?;
        if let Some(fill) = fill {
            fill.fill(&mut self.scene, affine, &path);
        }
        if let Some((style, color)) = stroke
            && style.width > 0.0
        {
            self.scene.stroke(&style, affine, color, None, &path);
        }
        Ok(())
    }

    fn draw_image(
        &mut self,
        frame: Frame,
        asset_id: Option<&str>,
        stroke: Option<&DisplayStroke>,
        transform: Transform,
        parent: Affine,
    ) -> Result<(), String> {
        let asset_id = asset_id.ok_or_else(|| "image has no asset id".to_owned())?;
        let image = self.images.get(asset_id)?;
        if image.width == 0 || image.height == 0 {
            return Err(format!("image asset {asset_id} has no pixels"));
        }
        let stroke = stroke.map(prepare_stroke).transpose()?;
        let affine = frame.transform(transform, parent)?;
        let mapping = Affine::translate((frame.rect.x0, frame.rect.y0))
            * Affine::scale_non_uniform(
                frame.rect.width() / f64::from(image.width),
                frame.rect.height() / f64::from(image.height),
            );
        self.scene
            .push_clip_layer(Fill::NonZero, affine, &frame.rect);
        self.scene
            .draw_image(&ImageBrush::new(image.clone()), affine * mapping);
        self.scene.pop_layer();
        if let Some((style, color)) = stroke
            && style.width > 0.0
        {
            self.scene.stroke(&style, affine, color, None, &frame.rect);
        }
        Ok(())
    }

    fn draw_text_box(
        &mut self,
        frame: Frame,
        has_text: bool,
        lines: &[PositionedTextLine],
        transform: Transform,
        parent: Affine,
    ) -> Result<(), String> {
        if has_text && !lines.iter().any(|line| !line.runs.is_empty()) {
            return Err("text box has text but no positioned runs".to_owned());
        }
        let affine = frame.transform(transform, parent)?;
        let mut prepared = Vec::new();
        for line in lines {
            finite(line.baseline, "text baseline")?;
            for run in &line.runs {
                prepared.push(self.prepare_text_run(line, run)?);
            }
        }
        self.scene
            .push_clip_layer(Fill::NonZero, affine, &frame.rect);
        for run in prepared {
            self.scene
                .draw_glyphs(&run.font)
                .font_size(run.font_size_px)
                .brush(run.color)
                .transform(affine)
                .draw(Fill::NonZero, run.glyphs.iter().copied());
            if run.underline && run.width > 0.0 {
                let top = run.baseline + run.font_size_px * 0.08;
                let thickness = (run.font_size_px * 0.05).max(1.0);
                let underline = Rect::new(
                    f64::from(run.x),
                    f64::from(top),
                    f64::from(run.x + run.width),
                    f64::from(top + thickness),
                );
                self.scene
                    .fill(Fill::NonZero, affine, run.color, None, &underline);
            }
        }
        self.scene.pop_layer();
        Ok(())
    }

    fn prepare_text_run(
        &mut self,
        line: &PositionedTextLine,
        run: &PositionedTextRun,
    ) -> Result<PreparedTextRun, String> {
        let x = finite(run.x, "text run x")?;
        let width = nonnegative(run.width, "text run width")?;
        let font_size_px = positive(run.font_size_px, "text font size")?;
        let font = self.fonts.vello_face(run.font_id)?.clone();
        let glyphs = run
            .glyphs
            .iter()
            .map(prepare_positioned_glyph)
            .collect::<Result<Vec<_>, String>>()?;
        self.summary.glyph_audit.glyph_runs += 1;
        self.summary.glyph_audit.glyphs += glyphs.len();
        self.check_width(run);
        self.check_carets(line, run);
        Ok(PreparedTextRun {
            font,
            font_size_px,
            x,
            baseline: line.baseline,
            width,
            glyphs,
            color: color(&run.color, 1.0)?,
            underline: run.underline,
        })
    }

    fn check_width(&mut self, run: &PositionedTextRun) {
        let glyph_width = run
            .glyphs
            .last()
            .map_or(0.0, |glyph| glyph.x + glyph.advance - run.x);
        let drift = (glyph_width - run.width).abs();
        let audit = &mut self.summary.glyph_audit;
        audit.width_checks += 1;
        audit.max_width_drift_px = audit.max_width_drift_px.max(drift);
        if drift > CARET_DRIFT_TOLERANCE_PX {
            audit.drifted_widths += 1;
            push_drift_sample(
                audit,
                format!(
                    "run {}..{} {:?} width {:.4} vs glyph advances {:.4}",
                    run.start,
                    run.end,
                    text_sample(&run.text),
                    run.width,
                    glyph_width
                ),
            );
        }
    }

    fn check_carets(&mut self, line: &PositionedTextLine, run: &PositionedTextRun) {
        let mut starts = run
            .glyphs
            .iter()
            .map(|glyph| glyph.cluster)
            .filter(|start| *start >= run.start && *start < run.end)
            .collect::<Vec<_>>();
        starts.push(run.start);
        starts.push(run.end);
        starts.sort_unstable();
        starts.dedup();
        let mut expected = vec![(run.start, run.x)];
        let mut cursor = run.x;
        for pair in starts.windows(2) {
            if let Some(glyph) = run.glyphs.iter().rfind(|glyph| glyph.cluster == pair[0]) {
                cursor = glyph.x + glyph.advance;
            }
            expected.push((pair[1], cursor));
        }
        let mut largest_drift = None;
        for (position, expected_x) in expected {
            let Some(actual) = line
                .caret_stops
                .iter()
                .filter(|stop| stop.position == position)
                .min_by(|left, right| {
                    (left.x - expected_x)
                        .abs()
                        .total_cmp(&(right.x - expected_x).abs())
                })
            else {
                self.summary.glyph_audit.missing_caret_stops += 1;
                continue;
            };
            let drift = (actual.x - expected_x).abs();
            let audit = &mut self.summary.glyph_audit;
            audit.caret_stops_checked += 1;
            audit.max_caret_drift_px = audit.max_caret_drift_px.max(drift);
            if drift > CARET_DRIFT_TOLERANCE_PX {
                audit.drifted_caret_stops += 1;
                if largest_drift
                    .as_ref()
                    .is_none_or(|(_, _, _, largest)| drift > *largest)
                {
                    largest_drift = Some((position, actual.x, expected_x, drift));
                }
            }
        }
        if let Some((position, actual_x, expected_x, drift)) = largest_drift {
            push_drift_sample(
                &mut self.summary.glyph_audit,
                format!(
                    "run {}..{} {:?} caret {position} at {:.4} vs glyphs {:.4} (drift {:.4})",
                    run.start,
                    run.end,
                    text_sample(&run.text),
                    actual_x,
                    expected_x,
                    drift
                ),
            );
        }
    }

    fn skip(&mut self, primitive: &Primitive, parent: Affine, kind: &str, reason: String) {
        self.summary.skipped(kind);
        self.skipped.record(kind, reason);
        let (bounds, transform) = placeholder_frame(primitive, parent);
        draw_transformed_placeholder(&mut self.scene, bounds, transform);
    }

    fn skip_subtree(&mut self, primitive: &Primitive, parent: Affine, reason: String) {
        let kind = primitive_kind(primitive);
        self.summary.seen(kind);
        self.summary.skipped(kind);
        self.skipped.record(kind, reason.clone());
        let (bounds, transform) = placeholder_frame(primitive, parent);
        draw_transformed_placeholder(&mut self.scene, bounds, transform);
        if let Primitive::Chart { primitives, .. } = primitive {
            for child in primitives {
                self.skip_subtree(child, transform, reason.clone());
            }
        }
    }
}

struct PreparedTextRun {
    font: FontData,
    font_size_px: f32,
    x: f32,
    baseline: f32,
    width: f32,
    glyphs: Vec<Glyph>,
    color: Color,
    underline: bool,
}

enum PreparedPaint {
    Solid(Color),
    Gradient(Gradient),
}

impl PreparedPaint {
    fn fill(&self, scene: &mut Scene, transform: Affine, shape: &impl vello::kurbo::Shape) {
        match self {
            Self::Solid(color) => scene.fill(Fill::NonZero, transform, *color, None, shape),
            Self::Gradient(gradient) => scene.fill(Fill::NonZero, transform, gradient, None, shape),
        }
    }
}

#[derive(Clone, Copy)]
struct Frame {
    rect: Rect,
}

impl Frame {
    fn new(x: f32, y: f32, w: f32, h: f32) -> Result<Self, String> {
        let x = finite(x, "primitive x")?;
        let y = finite(y, "primitive y")?;
        let w = nonnegative(w, "primitive width")?;
        let h = nonnegative(h, "primitive height")?;
        Ok(Self {
            rect: Rect::new(
                f64::from(x),
                f64::from(y),
                f64::from(x + w),
                f64::from(y + h),
            ),
        })
    }

    fn transform(self, transform: Transform, parent: Affine) -> Result<Affine, String> {
        let rotation = finite(transform.rotation_deg, "primitive rotation")?;
        let center = self.rect.center();
        let flip = Affine::translate(center.to_vec2())
            * Affine::scale_non_uniform(
                if transform.flip_h { -1.0 } else { 1.0 },
                if transform.flip_v { -1.0 } else { 1.0 },
            )
            * Affine::translate(-center.to_vec2());
        Ok(parent * Affine::rotate_about(f64::from(rotation).to_radians(), center) * flip)
    }
}

fn prepare_paint(paint: &Paint, bounds: Rect) -> Result<PreparedPaint, String> {
    match paint {
        Paint::Solid { color: value } => Ok(PreparedPaint::Solid(color(value, 1.0)?)),
        Paint::Gradient {
            gradient_type,
            angle_deg,
            stops,
        } => {
            if stops.is_empty() {
                return Err("gradient has no stops".to_owned());
            }
            let stops = stops
                .iter()
                .map(|stop| {
                    if !stop.position.is_finite() {
                        return Err("gradient stop position is not finite".to_owned());
                    }
                    Ok((stop.position.clamp(0.0, 1.0), color(&stop.color, 1.0)?))
                })
                .collect::<Result<Vec<_>, String>>()?;
            let center = bounds.center();
            let radius = bounds.width().hypot(bounds.height()) / 2.0;
            let gradient = match gradient_type {
                GradientType::Linear => {
                    let angle = finite(angle_deg.unwrap_or_default(), "gradient angle")?;
                    let radians = f64::from(angle).to_radians();
                    let dx = radians.cos() * radius;
                    let dy = radians.sin() * radius;
                    Gradient::new_linear(
                        (center.x - dx, center.y - dy),
                        (center.x + dx, center.y + dy),
                    )
                    .with_stops(stops.as_slice())
                }
                GradientType::Radial => {
                    Gradient::new_radial(center, radius as f32).with_stops(stops.as_slice())
                }
                GradientType::Rectangular => {
                    return Err("rectangular gradient is not translated".to_owned());
                }
                GradientType::Path => return Err("path gradient is not translated".to_owned()),
            };
            Ok(PreparedPaint::Gradient(gradient))
        }
    }
}

fn prepare_stroke(stroke: &DisplayStroke) -> Result<(Stroke, Color), String> {
    let width = nonnegative(stroke.width, "stroke width")?;
    let mut style = Stroke::new(f64::from(width));
    if stroke.dashed && width > 0.0 {
        style = with_dashes(
            style,
            [f64::from((width * 2.0).max(3.0)), f64::from(width.max(2.0))],
        );
    }
    Ok((style, color(&stroke.color, 1.0)?))
}

fn build_path(commands: &[GeometryPathCommand], frame: Frame) -> Result<BezPath, String> {
    let x = frame.rect.x0;
    let y = frame.rect.y0;
    let w = frame.rect.width();
    let h = frame.rect.height();
    let point = |px: f64, py: f64| -> Result<(f64, f64), String> {
        if !px.is_finite() || !py.is_finite() {
            return Err("shape path coordinate is not finite".to_owned());
        }
        Ok((x + px * w, y + py * h))
    };
    let mut path = BezPath::new();
    for command in commands {
        match *command {
            GeometryPathCommand::Move { x, y } => path.move_to(point(x, y)?),
            GeometryPathCommand::Line { x, y } => path.line_to(point(x, y)?),
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                path.quad_to(point(cpx, cpy)?, point(x, y)?);
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => path.curve_to(point(cp1x, cp1y)?, point(cp2x, cp2y)?, point(x, y)?),
            GeometryPathCommand::Close => path.close_path(),
        }
    }
    Ok(path)
}

fn placeholder_frame(primitive: &Primitive, parent: Affine) -> (Option<Rect>, Affine) {
    let (x, y, w, h, transform) = primitive_frame(primitive);
    let frame = Frame::new(x, y, w, h);
    match frame {
        Ok(frame) => (
            Some(frame.rect),
            frame.transform(transform, parent).unwrap_or(parent),
        ),
        Err(_) => (None, parent),
    }
}

fn primitive_frame(primitive: &Primitive) -> (f32, f32, f32, f32, Transform) {
    match primitive {
        Primitive::Shape {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::Image {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::TextBox {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::Placeholder {
            x,
            y,
            w,
            h,
            transform,
            ..
        }
        | Primitive::Chart {
            x,
            y,
            w,
            h,
            transform,
            ..
        } => (*x, *y, *w, *h, *transform),
    }
}

fn primitive_kind(primitive: &Primitive) -> &'static str {
    match primitive {
        Primitive::Shape { .. } => "shape",
        Primitive::Image { .. } => "image",
        Primitive::TextBox { .. } => "textBox",
        Primitive::Placeholder { .. } => "placeholder",
        Primitive::Chart { .. } => "chart",
    }
}

fn prepare_positioned_glyph(glyph: &PositionedGlyph) -> Result<Glyph, String> {
    let x = finite(glyph.x, "glyph x")?;
    let x_offset = finite(glyph.x_offset, "glyph x offset")?;
    finite(glyph.advance, "glyph advance")?;
    Ok(Glyph {
        id: glyph.glyph_id,
        x: finite(x + x_offset, "positioned glyph x")?,
        y: finite(glyph.y_offset, "glyph y offset")?,
    })
}

fn push_drift_sample(audit: &mut GlyphAudit, sample: String) {
    if audit.drift_samples.len() < 16 {
        audit.drift_samples.push(sample);
    }
}

fn text_sample(text: &str) -> String {
    text.chars().take(48).collect()
}

fn dimension(value: f32, label: &str, max_texture_dimension_2d: u32) -> Result<f64> {
    let value = positive(value, label).map_err(anyhow::Error::msg)?;
    if value > max_texture_dimension_2d as f32 {
        bail!(
            "requested {label} {value} exceeds GPU texture dimension ceiling {max_texture_dimension_2d}"
        );
    }
    Ok(f64::from(value))
}

fn finite(value: f32, label: &str) -> Result<f32, String> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(format!("{label} is not finite"))
    }
}

fn nonnegative(value: f32, label: &str) -> Result<f32, String> {
    let value = finite(value, label)?;
    if value >= 0.0 {
        Ok(value)
    } else {
        Err(format!("{label} is negative"))
    }
}

fn positive(value: f32, label: &str) -> Result<f32, String> {
    let value = finite(value, label)?;
    if value > 0.0 {
        Ok(value)
    } else {
        Err(format!("{label} is not positive"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use betteroffice_pptx::{CONTRACT_VERSION, GradientStop};
    use vello::kurbo::Shape;

    #[test]
    fn rejects_advanced_gradient_types_with_visible_reasons() {
        let paint = Paint::Gradient {
            gradient_type: GradientType::Rectangular,
            angle_deg: None,
            stops: vec![GradientStop {
                position: 0.0,
                color: "#ffffff".to_owned(),
            }],
        };
        assert_eq!(
            prepare_paint(&paint, Rect::new(0.0, 0.0, 100.0, 100.0))
                .err()
                .as_deref(),
            Some("rectangular gradient is not translated")
        );
    }

    #[test]
    fn counts_unsupported_paint_as_a_shape_skip() {
        let display_list = SurfaceDisplayList {
            contract_version: CONTRACT_VERSION,
            width: 100.0,
            height: 80.0,
            background: None,
            primitives: vec![Primitive::Shape {
                object_id: 1,
                shape_id: None,
                name: "advanced fill".to_owned(),
                x: 10.0,
                y: 10.0,
                w: 40.0,
                h: 30.0,
                geometry: "rect".to_owned(),
                path: vec![
                    GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 1.0, y: 0.0 },
                    GeometryPathCommand::Line { x: 1.0, y: 1.0 },
                    GeometryPathCommand::Close,
                ],
                adjust_values: BTreeMap::new(),
                fill: Some(Paint::Gradient {
                    gradient_type: GradientType::Path,
                    angle_deg: None,
                    stops: vec![GradientStop {
                        position: 0.0,
                        color: "#ffffff".to_owned(),
                    }],
                }),
                stroke: None,
                transform: Transform::default(),
            }],
        };
        let fonts = PptxFonts {
            vello: HashMap::new(),
            labels: Vec::new(),
        };
        let images = PptxImages {
            assets: HashMap::new(),
        };
        let (page, summary) = translate_slide(&display_list, &fonts, &images, 8_192).unwrap();
        let structured = summary.structured(&page.skipped);
        assert_eq!(page.skipped.counts["shape"], 1);
        assert_eq!(structured["primitives"]["shape"]["skipped"], 1);
        assert_eq!(
            structured["skipReasons"]["path gradient is not translated"],
            1
        );
    }

    #[test]
    fn builds_normalized_paths_in_the_primitive_frame() {
        let path = build_path(
            &[
                GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                GeometryPathCommand::Line { x: 1.0, y: 1.0 },
            ],
            Frame::new(10.0, 20.0, 30.0, 40.0).unwrap(),
        )
        .unwrap();
        assert_eq!(path.bounding_box(), Rect::new(10.0, 20.0, 40.0, 60.0));
    }

    #[test]
    fn uses_display_list_glyph_positions() {
        let glyph = prepare_positioned_glyph(&PositionedGlyph {
            glyph_id: 42,
            cluster: 3,
            x: 12.0,
            advance: 8.0,
            x_offset: 1.5,
            y_offset: 20.0,
        })
        .unwrap();
        assert_eq!(glyph.id, 42);
        assert_eq!((glyph.x, glyph.y), (13.5, 20.0));
    }

    #[test]
    fn rejects_slide_dimension_over_device_limit() {
        let error = dimension(8_200.0, "slide height", 8_192).unwrap_err();
        assert!(error.to_string().contains("slide height 8200"));
        assert!(error.to_string().contains("8192"));
    }
}
