use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
use std::sync::{Arc, Mutex};

use ooxml_drawingml::chart::{PlotRect, PlotTextAlign};
use ooxml_drawingml::{
    ColorValue, GeometryPathCommand, GradientFill, LineEnd, ResolvedCellStyle, ShapeEffects,
    ShapeFill, ShapeOutline, ShapeStyle, StyleReference, TableCellBorder,
    TableCellBorders as StyleCellBorders, TableCellPosition, TableCellStyle, TableStyle,
    TableStyleFlags, Theme, ThemeFormatScheme, normalize_table_column_widths,
    preset_geometry_to_path, resolve_color_value_to_hex_with_theme,
    resolve_color_value_to_rgba_hex, resolve_theme_font_ref, style_fill, style_outline,
};
use ooxml_text::{
    CompatFlags, FontId, FontStore, ShapeFeature, WORD_SMALL_CAPS_ADVANCE_SCALE,
    break_opportunities, shape, single_line_box, uppercase_for_language,
};
use pptx_edit::{
    DeckSnapshot, ShapeKind, ShapeSnapshot, SlideScope, SlideSnapshot, StorySnapshot, TextStyle,
};
use pptx_parse::{
    BlipEffect, Bullet, BulletColor, BulletFont, BulletSize, ChartPart, ChartSpace,
    CustomGeometryPath, GraphicFrameData, LineSpacing, MediaPart, ParagraphProperties, Picture,
    PictureCrop, PictureFill, Placeholder, PptxPackage, RunProperties, ShapeNode, ShapeTransform,
    Slide, SlideLayout, SlideMaster, Table, TableCell, TextAutofit, TextBody, TextCaps,
    TextOverflow, builtin_table_style, effective_color_map,
};
use serde::{Deserialize, Serialize};
use thiserror::Error;

use crate::chart::{ChartFrame, ChartText, chart_primitive};
use crate::family_metrics::{FamilyMetrics, family_advance, family_metrics};
use crate::metafile::{MetafileDrawing, decode as decode_metafile, is_metafile};
use crate::{
    CONTRACT_VERSION, CaretStop, GradientStop, GradientType, ImageCrop, ImageEffect, Paint,
    PositionedGlyph, PositionedTextLine, PositionedTextRun, Primitive, Shadow, Stroke, StrokeEnd,
    SurfaceDisplayList, TextAlign, TextAnchor, TextParagraph, TextRun, Transform,
};

const EMU_PER_CSS_PIXEL: f32 = 9_525.0;
const EMU_PER_POINT: f64 = 12_700.0;
const CSS_PIXELS_PER_POINT: f64 = 96.0 / 72.0;
const PICTURE_FILL: &str = "picture";
const LINE_END_MIN_BASE_PX: f32 = 0.7 / 25.4 * 96.0;
const ANGLE_UNITS_PER_DEGREE: f64 = 60_000.0;
const DEFAULT_INSET_HORIZONTAL_EMU: i64 = 91_440;
const DEFAULT_INSET_VERTICAL_EMU: i64 = 45_720;
const DEFAULT_FONT_SIZE_PT: f32 = 18.0;
/// Size a super/subscript run shapes at, relative to its own `sz`.
const SCRIPT_SIZE_RATIO: f32 = 0.58;
/// `p:bgRef/@idx` counts `a:bgFillStyleLst` entries from here.
const BACKGROUND_FILL_BASE: u32 = 1_001;
/// Measured on PowerPoint 16.113 exports: 1.2 em for every face.
const SINGLE_LINE_PITCH_EM: f32 = 1.2;
const MAX_FONT_BYTES: usize = 32 * 1024 * 1024;
const MAX_FONTS: usize = 256;
pub(crate) const MAX_RENDER_SHAPES: usize = 20_000;
const MAX_TEXT_BYTES: usize = 4 * 1024 * 1024;
const MAX_TEXT_LINES: usize = 100_000;
const MAX_TEXT_PARAGRAPHS: usize = 20_000;
const MAX_AUTONUM_VALUE: u32 = 32_767 + MAX_TEXT_PARAGRAPHS as u32;
const MAX_TEXT_RUNS: usize = 100_000;
/// Chart parts one slide may draw, shared across its charts.
pub(crate) const MAX_CHART_PRIMITIVES: usize = 100_000;
/// Metafile fills and strokes one slide may draw, shared across its pictures.
const MAX_METAFILE_PRIMITIVES: usize = 4_000;

#[derive(Debug, Error)]
pub enum RenderError {
    #[error("slide index {0} is outside the deck")]
    SlideNotFound(usize),
    #[error("no font has been registered for slide text")]
    NoFont,
    #[error("font error: {0}")]
    Font(String),
    #[error("render resource limit exceeded: {0}")]
    ResourceLimit(String),
}

#[derive(Clone)]
struct FontFace {
    id: FontId,
    family: String,
    requested_family: String,
    /// The named family's own advance widths, where this face stands in for a
    /// family that does not already run at them.
    widths: Option<&'static FamilyMetrics>,
    /// The same family's `hhea` metrics, which decide where the line box sits.
    line: Option<&'static FamilyMetrics>,
}

pub struct SlideRenderer {
    fonts: FontStore,
    faces: HashMap<(String, bool, bool), FontFace>,
    fallback: Option<FontFace>,
    /// Normalized fallback family.
    fallback_family: Option<String>,
    font_count: usize,
    /// Laid-out text reused across display-list builds and slides.
    text_layouts: Mutex<TextLayoutCache>,
}

impl Default for SlideRenderer {
    fn default() -> Self {
        Self::new()
    }
}

impl SlideRenderer {
    pub fn new() -> Self {
        Self {
            fonts: FontStore::new(),
            faces: HashMap::new(),
            fallback: None,
            fallback_family: None,
            font_count: 0,
            text_layouts: Mutex::new(TextLayoutCache::default()),
        }
    }

    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32, RenderError> {
        if bytes.len() > MAX_FONT_BYTES {
            return Err(RenderError::ResourceLimit(format!(
                "font exceeds {MAX_FONT_BYTES} bytes"
            )));
        }
        if self.font_count >= MAX_FONTS {
            return Err(RenderError::ResourceLimit(format!(
                "more than {MAX_FONTS} font faces"
            )));
        }
        let family = family.trim();
        if family.is_empty() {
            return Err(RenderError::Font("font family is empty".to_owned()));
        }
        let id = self
            .fonts
            .register(bytes.to_vec())
            .map_err(|error| RenderError::Font(error.to_string()))?;
        let requested = normalize_family(family);
        let metrics = family_metrics(&requested, bold, italic);
        let face = FontFace {
            id,
            family: family.to_owned(),
            requested_family: requested.clone(),
            widths: metrics.filter(|metrics| !runs_at_own_widths(&self.fonts, id, metrics)),
            line: metrics.filter(|metrics| !sits_on_own_baseline(&self.fonts, id, metrics)),
        };
        self.faces.insert((requested, bold, italic), face.clone());
        self.fallback.get_or_insert(face);
        self.fallback_family
            .get_or_insert_with(|| normalize_family(family));
        self.font_count += 1;
        if let Ok(cache) = self.text_layouts.get_mut() {
            cache.clear();
        }
        Ok(id.to_u32())
    }

    /// The store holding every registered face, so a raster backend can resolve
    /// the `font_id`s the display list references.
    pub fn fonts(&self) -> &FontStore {
        &self.fonts
    }

    /// First registered face for placeholder labels.
    pub fn fallback_font(&self) -> Option<FontId> {
        self.fallback.as_ref().map(|face| face.id)
    }

    pub fn layout_slide(
        &self,
        package: &PptxPackage,
        deck: &DeckSnapshot,
        slide_index: usize,
    ) -> Result<RenderedSlide, RenderError> {
        let deck_slide = deck
            .slides
            .get(slide_index)
            .ok_or(RenderError::SlideNotFound(slide_index))?;
        self.layout_resolved(
            package,
            deck_slide,
            deck.width_emu,
            deck.height_emu,
            slide_index,
        )
    }

    /// Render one slide from a [`pptx_edit::DeckSession::slide_scope`] snapshot,
    /// which materializes only that slide instead of the whole deck.
    pub fn layout_scoped_slide(
        &self,
        package: &PptxPackage,
        scope: &SlideScope,
    ) -> Result<RenderedSlide, RenderError> {
        self.layout_resolved(
            package,
            &scope.slide,
            scope.width_emu,
            scope.height_emu,
            scope.index,
        )
    }

    fn layout_resolved(
        &self,
        package: &PptxPackage,
        deck_slide: &SlideSnapshot,
        width_emu: i64,
        height_emu: i64,
        slide_index: usize,
    ) -> Result<RenderedSlide, RenderError> {
        let parsed_slide = deck_slide
            .source_part_path
            .as_deref()
            .and_then(|path| package.slides.iter().find(|slide| slide.part_path == path));
        let layout_path = deck_slide
            .layout_part_path
            .as_deref()
            .or_else(|| parsed_slide.and_then(|slide| slide.layout_part_path.as_deref()));
        let layout = layout_path
            .and_then(|path| {
                package
                    .layouts
                    .iter()
                    .find(|layout| layout.part_path == path)
            })
            .or_else(|| package.layouts.first());
        let master = layout
            .and_then(|layout| layout.master_part_path.as_deref())
            .and_then(|path| {
                package
                    .masters
                    .iter()
                    .find(|master| master.part_path == path)
            })
            .or_else(|| {
                layout.and_then(|layout| {
                    package.masters.iter().find(|master| {
                        master
                            .layout_part_paths
                            .iter()
                            .any(|path| path == &layout.part_path)
                    })
                })
            })
            .or_else(|| package.masters.first());
        let theme_part = master
            .and_then(|master| master.theme_part_path.as_deref())
            .and_then(|path| package.themes.iter().find(|theme| theme.part_path == path))
            .or_else(|| package.themes.first());
        let default_theme = Theme::default();
        let base_theme = theme_part.map(|part| &part.theme).unwrap_or(&default_theme);
        // The slot mapping is per slide, not per theme part.
        let color_map = effective_color_map(parsed_slide, layout, master);
        let mapped_theme;
        let theme = if color_map.is_identity() {
            base_theme
        } else {
            mapped_theme = Theme {
                color_map,
                ..base_theme.clone()
            };
            &mapped_theme
        };
        let default_format_scheme = ThemeFormatScheme::default();
        let format_scheme = theme_part
            .map(|part| &part.format_scheme)
            .unwrap_or(&default_format_scheme);
        let background_source = [
            parsed_slide.map(BackgroundSource::from_slide),
            layout.map(BackgroundSource::from_layout),
            master.map(BackgroundSource::from_master),
        ]
        .into_iter()
        .flatten()
        .find(BackgroundSource::is_declared);
        let background_fill =
            background_source
                .as_ref()
                .and_then(|source| match source.reference {
                    Some(reference) => Some(style_fill(format_scheme, reference, theme)),
                    None => source.fill.cloned(),
                });
        let background_picture = background_source.as_ref().and_then(|source| {
            source.picture.or_else(|| {
                let index = source.reference?.index.checked_sub(BACKGROUND_FILL_BASE)?;
                theme_part?
                    .background_pictures
                    .get(index as usize)?
                    .as_ref()
            })
        });
        // A referenced picture style carries no colour of its own, so keep the
        // reference's own colour for the fills we cannot paint, such as a tile.
        let background = background_fill
            .as_ref()
            .and_then(|fill| paint(fill, theme))
            .or_else(|| {
                background_source
                    .as_ref()
                    .and_then(|source| source.fill)
                    .and_then(|fill| paint(fill, theme))
            })
            .filter(|paint| !is_invisible(paint))
            .or_else(|| {
                Some(Paint::Solid {
                    color: "#ffffff".to_owned(),
                })
            });
        let width = slide_extent_px(width_emu);
        let height = slide_extent_px(height_emu);
        let mut builder = LayoutBuilder {
            renderer: self,
            package,
            theme,
            format_scheme,
            theme_part_path: theme_part.map(|part| part.part_path.as_str()),
            master,
            layout,
            parsed_slide,
            primitives: Vec::new(),
            hit_regions: Vec::new(),
            shape_count: 0,
            line_count: 0,
            chart_budget: MAX_CHART_PRIMITIVES,
            metafile_budget: MAX_METAFILE_PRIMITIVES,
            metafile_bytes: 32 * 1024 * 1024,
            metafiles: HashMap::new(),
            media_parts: None,
            chart_parts: None,
            slide_number: i64::from(package.presentation.first_slide_num) + slide_index as i64,
        };
        let root_space = Space::root();
        if let Some(picture) = background_picture
            && let Some(asset_id) = picture.media_part_path.clone()
        {
            builder.primitives.push(Primitive::Image {
                geometry_fallback: false,
                object_id: 0,
                shape_id: None,
                name: "Background".to_owned(),
                x: 0.0,
                y: 0.0,
                w: width,
                h: height,
                asset_id: Some(asset_id),
                effects: Vec::new(),
                crop: picture_fill_crop(picture),
                path: None,
                stroke: None,
                shadow: None,
                transform: Transform::default(),
            });
        }
        // `p:sld/@showMasterSp` hides the layout's own decoration as well as the
        // master's, because the master reaches the slide through the layout.
        let show_layout = parsed_slide.is_none_or(|slide| slide.show_master_shapes);
        let show_master = show_layout && layout.is_none_or(|layout| layout.show_master_shapes);
        if show_master && let Some(master) = master {
            for (index, shape) in master.shapes.iter().enumerate() {
                if node_placeholder(shape).is_none() {
                    builder.render_parsed_shape(
                        shape,
                        &format!("master:{}:{index}", master.part_path),
                        root_space,
                    )?;
                }
            }
        }
        if show_layout && let Some(layout) = layout {
            for (index, shape) in layout.shapes.iter().enumerate() {
                if node_placeholder(shape).is_none() {
                    builder.render_parsed_shape(
                        shape,
                        &format!("layout:{}:{index}", layout.part_path),
                        root_space,
                    )?;
                }
            }
        }
        for shape in &deck_slide.shapes {
            builder.render_snapshot_shape(shape, root_space)?;
        }
        Ok(RenderedSlide {
            display_list: SurfaceDisplayList {
                contract_version: CONTRACT_VERSION,
                width,
                height,
                background,
                primitives: builder.primitives,
            },
            hit_regions: builder.hit_regions,
        })
    }

    fn resolve_face(
        &self,
        family: &str,
        bold: bool,
        italic: bool,
    ) -> Result<FontFace, RenderError> {
        let requested = normalize_family(family);
        let styles = [
            (bold, italic),
            (bold, false),
            (false, italic),
            (false, false),
        ];
        for (face_bold, face_italic) in styles {
            if let Some(face) = self.faces.get(&(requested.clone(), face_bold, face_italic)) {
                return Ok(if (face_bold, face_italic) == (bold, italic) {
                    face.clone()
                } else {
                    self.with_requested_metrics(face, &requested, bold, italic)
                });
            }
        }
        self.faces
            .iter()
            .filter(|((name, _, _), _)| Some(name) == self.fallback_family.as_ref())
            .min_by_key(|((_, face_bold, face_italic), _)| {
                (
                    2 * u8::from(*face_bold != bold) + u8::from(*face_italic != italic),
                    *face_bold,
                    *face_italic,
                )
            })
            .map(|(_, face)| self.with_requested_metrics(face, &requested, bold, italic))
            .ok_or(RenderError::NoFont)
    }

    fn with_requested_metrics(
        &self,
        face: &FontFace,
        requested: &str,
        bold: bool,
        italic: bool,
    ) -> FontFace {
        let metrics = family_metrics(requested, bold, italic);
        FontFace {
            requested_family: requested.to_owned(),
            widths: metrics.filter(|metrics| !runs_at_own_widths(&self.fonts, face.id, metrics)),
            line: metrics.filter(|metrics| !sits_on_own_baseline(&self.fonts, face.id, metrics)),
            ..face.clone()
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum HitTestResult {
    Shape {
        shape_id: String,
    },
    Text {
        shape_id: String,
        story_id: String,
        position: u32,
    },
}

/// The first of slide, layout and master that declares `p:bg`.
struct BackgroundSource<'a> {
    fill: Option<&'a ShapeFill>,
    picture: Option<&'a PictureFill>,
    reference: Option<&'a StyleReference>,
}

impl<'a> BackgroundSource<'a> {
    fn from_slide(slide: &'a Slide) -> Self {
        Self {
            fill: slide.background.as_ref(),
            picture: slide.background_picture.as_deref(),
            reference: slide.background_reference.as_ref(),
        }
    }

    fn from_layout(layout: &'a SlideLayout) -> Self {
        Self {
            fill: layout.background.as_ref(),
            picture: layout.background_picture.as_deref(),
            reference: layout.background_reference.as_ref(),
        }
    }

    fn from_master(master: &'a SlideMaster) -> Self {
        Self {
            fill: master.background.as_ref(),
            picture: master.background_picture.as_deref(),
            reference: master.background_reference.as_ref(),
        }
    }

    fn is_declared(&self) -> bool {
        self.fill.is_some() || self.picture.is_some() || self.reference.is_some()
    }
}

pub struct RenderedSlide {
    pub display_list: SurfaceDisplayList,
    hit_regions: Vec<HitRegion>,
}

impl RenderedSlide {
    pub fn hit_test(&self, x: f32, y: f32) -> Option<HitTestResult> {
        if !x.is_finite() || !y.is_finite() {
            return None;
        }
        for region in self.hit_regions.iter().rev() {
            let (shape_x, shape_y) = region.local_point(x, y);
            let inside_shape = region.rect.contains(shape_x, shape_y);
            if let Some(text) = &region.text {
                let (text_x, text_y) = text.local_point(x, y);
                if !inside_shape && !(text.overflow && region.hit_rect.contains(text_x, text_y)) {
                    continue;
                }
                if let Some(line) = nearest_line(&text.lines, text_y)
                    && let Some(caret) = line.caret_stops.iter().min_by(|left, right| {
                        (left.x - text_x).abs().total_cmp(&(right.x - text_x).abs())
                    })
                {
                    return Some(HitTestResult::Text {
                        shape_id: region.shape_id.clone(),
                        story_id: text.story_id.clone(),
                        position: caret.position,
                    });
                }
            } else if !inside_shape {
                continue;
            }
            return Some(HitTestResult::Shape {
                shape_id: region.shape_id.clone(),
            });
        }
        None
    }
}

struct LayoutBuilder<'a> {
    renderer: &'a SlideRenderer,
    package: &'a PptxPackage,
    theme: &'a Theme,
    format_scheme: &'a ThemeFormatScheme,
    theme_part_path: Option<&'a str>,
    master: Option<&'a SlideMaster>,
    layout: Option<&'a SlideLayout>,
    parsed_slide: Option<&'a Slide>,
    primitives: Vec<Primitive>,
    hit_regions: Vec<HitRegion>,
    shape_count: usize,
    line_count: usize,
    chart_budget: usize,
    metafile_budget: usize,
    metafile_bytes: usize,
    metafiles: HashMap<&'a str, Option<Arc<MetafileDrawing>>>,
    /// Package media keyed by part path, built on first picture lookup so
    /// per-shape work stays independent of the deck's part count.
    media_parts: Option<HashMap<&'a str, &'a MediaPart>>,
    /// Package charts keyed by (part path, theme part path); same rationale.
    chart_parts: Option<HashMap<(&'a str, Option<&'a str>), &'a ChartPart>>,
    /// The number a `slidenum` field resolves to on this slide.
    slide_number: i64,
}

impl<'a> LayoutBuilder<'a> {
    fn resolved_fill(&self, nodes: &[Option<&ShapeNode>]) -> Option<ShapeFill> {
        nodes
            .iter()
            .flatten()
            .find_map(|node| {
                if node_style(node).is_some_and(|style| style.fill_disabled) {
                    Some(ShapeFill::named("none"))
                } else {
                    node_fill(node).cloned()
                }
            })
            .or_else(|| {
                let reference = nodes
                    .iter()
                    .flatten()
                    .find_map(|node| node_style(node)?.fill.as_ref())?;
                Some(style_fill(self.format_scheme, reference, self.theme))
            })
    }

    fn resolved_outline(&self, nodes: &[Option<&ShapeNode>]) -> Option<ShapeOutline> {
        let mut outline = nodes.iter().flatten().find_map(|node| {
            let reference = node_style(node)?.line.as_ref()?;
            Some(style_outline(self.format_scheme, reference, self.theme).unwrap_or_default())
        });
        for node in nodes.iter().rev().flatten() {
            if node_style(node).is_some_and(|style| style.line_disabled) {
                outline = Some(ShapeOutline::default());
            }
            if let Some(direct) = node_outline(node) {
                outline = Some(merge_outline(direct, outline.as_ref()));
            }
        }
        outline
    }

    fn charge_shape(&mut self) -> Result<(), RenderError> {
        self.shape_count += 1;
        if self.shape_count > MAX_RENDER_SHAPES {
            return Err(RenderError::ResourceLimit(format!(
                "more than {MAX_RENDER_SHAPES} shapes"
            )));
        }
        Ok(())
    }

    fn render_snapshot_shape(
        &mut self,
        shape: &ShapeSnapshot,
        space: Space,
    ) -> Result<(), RenderError> {
        self.charge_shape()?;
        if shape.hidden {
            return Ok(());
        }
        let original = (shape.source_id != 0)
            .then(|| {
                self.parsed_slide
                    .and_then(|slide| find_node(&slide.shapes, shape.source_id))
            })
            .flatten();
        let layout_node = shape.placeholder.as_ref().and_then(|placeholder| {
            self.layout
                .and_then(|layout| find_placeholder(&layout.shapes, placeholder))
        });
        let master_node = shape.placeholder.as_ref().and_then(|placeholder| {
            self.master
                .and_then(|master| find_placeholder(&master.shapes, placeholder))
        });
        let resolved = resolved_transform_value(shape, original, layout_node, master_node);
        let rect = space.map_transform(&resolved);
        if shape.kind == ShapeKind::Group {
            let group_transform = original
                .and_then(node_group_transform)
                .or_else(|| layout_node.and_then(node_group_transform))
                .or_else(|| master_node.and_then(node_group_transform));
            let child_space = group_transform
                .map(|transform| Space::for_group(rect, transform))
                .unwrap_or(space);
            for child in &shape.children {
                self.render_snapshot_shape(child, child_space)?;
            }
            return Ok(());
        }
        let stable_id = shape.id.clone();
        let inherited = [original, layout_node, master_node];
        let inherited_fill = self.resolved_fill(&inherited);
        let inherited_outline = self.resolved_outline(&inherited);
        let effective_fill = shape.fill.as_ref().or(inherited_fill.as_ref());
        let picture = effective_fill
            .filter(|fill| fill.fill_type == PICTURE_FILL)
            .and_then(|_| picture_fill(&inherited));
        let fill = effective_fill.and_then(|fill| paint(fill, self.theme));
        let outline = if shape.outline.as_ref() == original.and_then(node_outline) {
            inherited_outline
        } else {
            shape
                .outline
                .as_ref()
                .map(|edited| {
                    if *edited == ShapeOutline::default() {
                        edited.clone()
                    } else {
                        merge_outline(edited, inherited_outline.as_ref())
                    }
                })
                .or(inherited_outline)
        }
        .and_then(|outline| stroke(&outline, self.theme));
        let node_effects = original
            .and_then(node_effects)
            .or_else(|| layout_node.and_then(node_effects))
            .or_else(|| master_node.and_then(node_effects));
        let shadow = node_effects
            .filter(|_| match shape.kind {
                ShapeKind::Shape => fill.is_some() || outline.is_some() || picture.is_some(),
                ShapeKind::Picture => true,
                ShapeKind::GraphicFrame | ShapeKind::Group => false,
            })
            .and_then(|effects| {
                shadow(
                    effects,
                    self.theme,
                    space,
                    rect,
                    shape.rotation_deg as f32,
                    shape.flip_h,
                    shape.flip_v,
                )
            });
        let transform = Transform {
            rotation_deg: shape.rotation_deg as f32,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
        };
        match shape.kind {
            ShapeKind::Shape => {
                let paths = custom_paths(original);
                let (path, mut geometry_fallback) = geometry_path_with_fallback(
                    &shape.geometry,
                    &shape.adjust_values,
                    f64::from(rect.w) / f64::from(rect.h),
                );
                if shape.geometry == "custom" && !paths.is_empty() {
                    geometry_fallback = false;
                }
                self.push_shape(
                    Primitive::Shape {
                        clip: None,
                        even_odd: false,
                        object_id: shape.source_id,
                        shape_id: Some(stable_id.clone()),
                        name: shape.name.clone(),
                        x: rect.x,
                        y: rect.y,
                        w: rect.w,
                        h: rect.h,
                        geometry: shape.geometry.clone(),
                        path,
                        geometry_fallback,
                        adjust_values: shape
                            .adjust_values
                            .iter()
                            .map(|(name, value)| (name.clone(), *value as f32))
                            .collect(),
                        fill,
                        stroke: outline,
                        shadow,
                        transform,
                    },
                    paths,
                    picture,
                )?;
            }
            ShapeKind::Picture => {
                let source = picture_source(original);
                let asset_id = picture_asset_id(shape);
                self.render_picture(
                    shape.source_id,
                    &stable_id,
                    &shape.name,
                    rect,
                    transform,
                    asset_id.as_deref(),
                    &shape.blip_effects,
                    source.map(|picture| &picture.crop),
                    source
                        .map(|picture| picture_mask(picture, rect))
                        .unwrap_or_default(),
                    outline,
                    shadow,
                );
            }
            ShapeKind::GraphicFrame => {
                self.render_graphic_frame(
                    shape.source_id,
                    &stable_id,
                    &shape.name,
                    rect,
                    transform,
                    space,
                    shape.graphic.as_ref(),
                    &shape.text_stories,
                )?;
            }
            ShapeKind::Group => unreachable!(),
        }
        let body_cascade = BodyCascade {
            primary: original.and_then(node_text),
            layout: layout_node.and_then(node_text),
            master: master_node.and_then(node_text),
            master_slide: self.master,
            placeholder: shape.placeholder.as_ref(),
            style_color: shape_style_color(original),
        };
        let text = match shape.kind {
            ShapeKind::GraphicFrame => None,
            ShapeKind::Shape | ShapeKind::Picture | ShapeKind::Group => {
                shape.text_stories.first().map(content_from_story)
            }
        };
        let text_hit = if let Some(content) = text {
            Some(self.render_text_box(
                shape.source_id,
                &stable_id,
                rect,
                transform,
                content,
                body_cascade,
            )?)
        } else {
            None
        };
        self.hit_regions.push(HitRegion {
            shape_id: stable_id,
            rect,
            hit_rect: rect_covering_text(rect, text_hit.as_ref()),
            transform,
            text: text_hit,
        });
        Ok(())
    }

    fn render_parsed_shape(
        &mut self,
        shape: &ShapeNode,
        stable_id: &str,
        space: Space,
    ) -> Result<(), RenderError> {
        self.charge_shape()?;
        if node_base(shape).hidden {
            return Ok(());
        }
        if let ShapeNode::Group(group) = shape {
            let rect = space.map_transform(&group.base.transform);
            let child_space = Space::for_group(rect, &group.base.transform);
            for (index, child) in group.children.iter().enumerate() {
                self.render_parsed_shape(child, &format!("{stable_id}:{index}"), child_space)?;
            }
            return Ok(());
        }
        let base = node_base(shape);
        let rect = space.map_transform(&base.transform);
        let transform = Transform {
            rotation_deg: base.transform.rotation_deg as f32,
            flip_h: base.transform.flip_h,
            flip_v: base.transform.flip_v,
        };
        match shape {
            ShapeNode::Shape(value) => {
                let resolved_fill = self.resolved_fill(&[Some(shape)]);
                let fill = resolved_fill
                    .as_ref()
                    .and_then(|fill| paint(fill, self.theme));
                let outline = self
                    .resolved_outline(&[Some(shape)])
                    .and_then(|outline| stroke(&outline, self.theme));
                let picture = resolved_fill
                    .as_ref()
                    .filter(|fill| fill.fill_type == PICTURE_FILL)
                    .and(value.picture_fill.as_deref());
                let shadow = value
                    .effects
                    .as_ref()
                    .filter(|_| fill.is_some() || outline.is_some() || picture.is_some())
                    .and_then(|effects| {
                        shadow(
                            effects,
                            self.theme,
                            space,
                            rect,
                            transform.rotation_deg,
                            transform.flip_h,
                            transform.flip_v,
                        )
                    });
                let (path, mut geometry_fallback) = geometry_path_with_fallback(
                    &value.geometry,
                    &value.adjust_values,
                    f64::from(rect.w) / f64::from(rect.h),
                );
                if value.geometry == "custom" && !value.paths.is_empty() {
                    geometry_fallback = false;
                }
                self.push_shape(
                    Primitive::Shape {
                        clip: None,
                        even_odd: false,
                        object_id: base.id,
                        shape_id: Some(stable_id.to_owned()),
                        name: base.name.clone(),
                        x: rect.x,
                        y: rect.y,
                        w: rect.w,
                        h: rect.h,
                        geometry: value.geometry.clone(),
                        path,
                        geometry_fallback,
                        adjust_values: value
                            .adjust_values
                            .iter()
                            .map(|(name, value)| (name.clone(), *value as f32))
                            .collect(),
                        fill,
                        stroke: outline,
                        shadow,
                        transform,
                    },
                    &value.paths,
                    picture,
                )?;
            }
            ShapeNode::Picture(value) => {
                let outline = self
                    .resolved_outline(&[Some(shape)])
                    .and_then(|outline| stroke(&outline, self.theme));
                self.render_picture(
                    base.id,
                    stable_id,
                    &base.name,
                    rect,
                    transform,
                    value.media_part_path.as_deref(),
                    &value.effects,
                    Some(&value.crop),
                    picture_mask(value, rect),
                    outline,
                    value.shape_effects.as_ref().and_then(|effects| {
                        shadow(
                            effects,
                            self.theme,
                            space,
                            rect,
                            transform.rotation_deg,
                            transform.flip_h,
                            transform.flip_v,
                        )
                    }),
                );
            }
            ShapeNode::GraphicFrame(value) => {
                self.render_graphic_frame(
                    base.id,
                    stable_id,
                    &base.name,
                    rect,
                    transform,
                    space,
                    Some(&value.data),
                    &[],
                )?;
            }
            ShapeNode::Group(_) => unreachable!(),
        }
        let text_hit = if let Some(body) = node_text(shape) {
            let content = content_from_body(stable_id, body, self.theme, self.slide_number);
            Some(self.render_text_box(
                base.id,
                stable_id,
                rect,
                transform,
                content,
                BodyCascade {
                    primary: Some(body),
                    layout: None,
                    master: None,
                    master_slide: self.master,
                    placeholder: base.placeholder.as_ref(),
                    style_color: shape_style_color(Some(shape)),
                },
            )?)
        } else {
            None
        };
        self.hit_regions.push(HitRegion {
            shape_id: stable_id.to_owned(),
            rect,
            hit_rect: rect_covering_text(rect, text_hit.as_ref()),
            transform,
            text: text_hit,
        });
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn render_picture(
        &mut self,
        object_id: u32,
        shape_id: &str,
        name: &str,
        rect: PxRect,
        transform: Transform,
        media_part_path: Option<&str>,
        effects: &[BlipEffect],
        crop: Option<&PictureCrop>,
        mask: PictureMask,
        outline: Option<Stroke>,
        shadow: Option<Shadow>,
    ) {
        if effects.is_empty()
            && self.push_metafile(
                object_id,
                name,
                rect,
                transform,
                media_part_path,
                crop,
                &mask,
            )
        {
            if let Some(outline) = outline {
                self.primitives.push(Primitive::Shape {
                    clip: None,
                    even_odd: false,
                    object_id,
                    shape_id: Some(shape_id.to_owned()),
                    name: name.to_owned(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    geometry: "rect".to_owned(),
                    path: mask
                        .path
                        .clone()
                        .unwrap_or_else(|| geometry_path("rect", &BTreeMap::new(), 1.0)),
                    geometry_fallback: mask.geometry_fallback,
                    adjust_values: BTreeMap::new(),
                    fill: None,
                    stroke: Some(outline),
                    shadow: None,
                    transform,
                });
            }
            return;
        }
        self.primitives.push(Primitive::Image {
            object_id,
            shape_id: Some(shape_id.to_owned()),
            name: name.to_owned(),
            x: rect.x,
            y: rect.y,
            w: rect.w,
            h: rect.h,
            asset_id: media_part_path.map(str::to_owned),
            effects: image_effects(effects, self.theme),
            crop: crop.map(image_crop).unwrap_or_default(),
            path: mask.path,
            geometry_fallback: mask.geometry_fallback,
            stroke: outline,
            shadow,
            transform,
        });
    }

    #[allow(clippy::too_many_arguments)]
    fn push_metafile(
        &mut self,
        object_id: u32,
        name: &str,
        rect: PxRect,
        transform: Transform,
        media_part_path: Option<&str>,
        crop: Option<&PictureCrop>,
        mask: &PictureMask,
    ) -> bool {
        if self.metafile_budget == 0 {
            return false;
        }
        let Some(drawing) = self.metafile(media_part_path) else {
            return false;
        };
        if drawing.ops.len() > self.metafile_budget {
            return false;
        }
        self.metafile_budget -= drawing.ops.len();
        let Some(crop) = source_rect(crop) else {
            return true;
        };
        let masked = mask.path.is_some();
        let clip = mask
            .path
            .clone()
            .unwrap_or_else(|| geometry_path("rect", &BTreeMap::new(), 1.0));
        for mut op in drawing.ops.iter().cloned() {
            place_in_source_rect(&mut op.path, crop);
            // one clip per primitive, so a picture mask and the metafile's own
            // clip cannot both be carried: draw under the mask only where the
            // metafile clip would remove nothing, and drop the op otherwise
            // rather than paint what either region hides.
            let op_clip = match (op.clip, masked) {
                (Some(rect), false) => match metafile_clip_path(rect, crop) {
                    Some(path) => path,
                    None => continue,
                },
                (Some(rect), true) => match clipped_away(&op.path, rect, crop) {
                    true => continue,
                    false => clip.clone(),
                },
                (None, _) => clip.clone(),
            };
            self.primitives.push(Primitive::Shape {
                clip: Some(op_clip),
                even_odd: op.even_odd,
                object_id,
                shape_id: None,
                name: name.to_owned(),
                x: rect.x,
                y: rect.y,
                w: rect.w,
                h: rect.h,
                geometry: "custom".to_owned(),
                path: op.path,
                geometry_fallback: mask.geometry_fallback,
                adjust_values: BTreeMap::new(),
                fill: op.fill.map(|color| Paint::Solid { color }),
                stroke: op.stroke.map(|stroke| Stroke {
                    join: None,
                    color: stroke.color,
                    paint: None,
                    width: ((stroke.width * crop.2 * f64::from(rect.w)) as f32).max(1.0),
                    dashed: false,
                    head_end: None,
                    tail_end: None,
                }),
                shadow: None,
                transform,
            });
        }
        true
    }

    fn media_part(&mut self, part_path: &str) -> Option<&'a MediaPart> {
        self.media_parts
            .get_or_insert_with(|| {
                let mut index = HashMap::new();
                for part in &self.package.media {
                    index.entry(part.part_path.as_str()).or_insert(part);
                }
                index
            })
            .get(part_path)
            .copied()
    }

    fn metafile(&mut self, media_part_path: Option<&str>) -> Option<Arc<MetafileDrawing>> {
        let part_path = media_part_path?;
        if let Some(drawing) = self.metafiles.get(part_path) {
            return drawing.clone();
        }
        let part = self.media_part(part_path)?;
        if !is_metafile(&part.bytes) || part.bytes.len() > self.metafile_bytes {
            self.metafiles.insert(part.part_path.as_str(), None);
            return None;
        }
        self.metafile_bytes -= part.bytes.len();
        let drawing = decode_metafile(&part.bytes).map(Arc::new);
        self.metafiles
            .insert(part.part_path.as_str(), drawing.clone());
        drawing
    }

    fn push_shape(
        &mut self,
        primitive: Primitive,
        paths: &[CustomGeometryPath],
        picture: Option<&PictureFill>,
    ) -> Result<(), RenderError> {
        if paths.is_empty()
            || !matches!(&primitive, Primitive::Shape { geometry, .. } if geometry == "custom")
        {
            for (index, (primitive, has_fill)) in crate::geometry::preset_primitives(primitive)
                .into_iter()
                .enumerate()
            {
                if index > 0 {
                    self.charge_shape()?;
                }
                let picture = picture.filter(|_| has_fill);
                self.primitives.push(picture_filled(primitive, picture));
            }
            return Ok(());
        }
        for (index, custom) in paths.iter().enumerate() {
            if index > 0 {
                self.charge_shape()?;
            }
            let mut primitive = primitive.clone();
            if let Primitive::Shape {
                path,
                fill,
                stroke,
                shadow,
                ..
            } = &mut primitive
            {
                *path = custom.commands.clone();
                if custom.no_fill {
                    *fill = None;
                }
                if custom.no_stroke {
                    *stroke = None;
                }
                if fill.is_none() && stroke.is_none() {
                    *shadow = None;
                }
            }
            let picture = picture.filter(|_| !custom.no_fill);
            self.primitives.push(picture_filled(primitive, picture));
        }
        Ok(())
    }

    /// Plots a chart or table frame, or keeps the placeholder for graphics
    /// that carry no drawable data.
    #[allow(clippy::too_many_arguments)]
    fn render_graphic_frame(
        &mut self,
        object_id: u32,
        shape_id: &str,
        name: &str,
        rect: PxRect,
        transform: Transform,
        frame_space: Space,
        graphic: Option<&GraphicFrameData>,
        stories: &[StorySnapshot],
    ) -> Result<(), RenderError> {
        if let Some(space) = self.chart_space(graphic) {
            let frame = ChartFrame {
                object_id,
                shape_id: Some(shape_id),
                name,
                rect: PlotRect {
                    x: f64::from(rect.x),
                    y: f64::from(rect.y),
                    w: f64::from(rect.w),
                    h: f64::from(rect.h),
                },
                transform,
            };
            let (renderer, theme) = (self.renderer, self.theme);
            let default_font = resolve_theme_font_ref(Some(theme), "+mn-lt");
            let chart = chart_primitive(
                frame,
                space,
                &default_font,
                self.chart_budget,
                &mut |text| chart_text_primitive(renderer, theme, shape_id, text),
            )?;
            if let Primitive::Chart { primitives, .. } = &chart {
                self.chart_budget -= primitives.len();
            }
            self.primitives.push(chart);
            return Ok(());
        }
        if let Some(GraphicFrameData::Unknown {
            picture: Some(picture),
            ..
        }) = graphic
        {
            let outline = picture
                .outline
                .as_ref()
                .and_then(|outline| stroke(outline, self.theme));
            self.render_picture(
                object_id,
                shape_id,
                name,
                rect,
                transform,
                picture.media_part_path.as_deref(),
                &picture.effects,
                Some(&picture.crop),
                picture_mask(picture, rect),
                outline,
                picture.shape_effects.as_ref().and_then(|effects| {
                    shadow(
                        effects,
                        self.theme,
                        frame_space,
                        rect,
                        transform.rotation_deg,
                        transform.flip_h,
                        transform.flip_v,
                    )
                }),
            );
            return Ok(());
        }
        if let Some(GraphicFrameData::Table(table)) = graphic {
            return self.render_table(object_id, shape_id, name, rect, transform, table, stories);
        }
        self.primitives.push(Primitive::Placeholder {
            object_id,
            shape_id: Some(shape_id.to_owned()),
            name: name.to_owned(),
            x: rect.x,
            y: rect.y,
            w: rect.w,
            h: rect.h,
            label: graphic_label(graphic),
            transform,
        });
        Ok(())
    }

    /// Lays a table out into one container primitive. Every fill paints before
    /// every border, or a shared edge is overpainted.
    #[allow(clippy::too_many_arguments)]
    fn render_table(
        &mut self,
        object_id: u32,
        shape_id: &str,
        name: &str,
        rect: PxRect,
        transform: Transform,
        table: &Table,
        stories: &[StorySnapshot],
    ) -> Result<(), RenderError> {
        let row_count = table.rows.len();
        let column_count = table.grid.len().max(
            table
                .rows
                .iter()
                .map(|row| row.cells.len())
                .max()
                .unwrap_or_default(),
        );
        if row_count == 0 || column_count == 0 {
            return Ok(());
        }
        let columns = column_edges(&table.grid, column_count, rect);
        let flags = TableStyleFlags {
            first_row: table.properties.first_row,
            last_row: table.properties.last_row,
            first_column: table.properties.first_col,
            last_column: table.properties.last_col,
            band_row: table.properties.band_row,
        };
        let package = self.package;
        let table_style = table.properties.style_id.as_deref().and_then(|id| {
            package
                .table_styles
                .style(Some(id))
                .or_else(|| builtin_table_style(id))
        });
        let mut heights: Vec<f32> = table.rows.iter().map(|row| emu_to_px(row.height)).collect();
        let mut spans = Vec::new();
        let mut plans = Vec::new();
        let mut story_index = 0;
        for (row_index, row) in table.rows.iter().enumerate() {
            for (column, cell) in row.cells.iter().enumerate() {
                let story = stories.get(story_index);
                story_index += 1;
                if column >= column_count {
                    continue;
                }
                self.charge_shape()?;
                if cell.merged {
                    continue;
                }
                let span = (cell.grid_span as usize).clamp(1, column_count - column);
                let row_span = (cell.row_span as usize).clamp(1, row_count - row_index);
                let content = match story {
                    Some(story) => content_from_story(story),
                    None => content_from_body(
                        &format!("{shape_id}:table:{row_index}:{column}"),
                        &cell.text,
                        self.theme,
                        self.slide_number,
                    ),
                };
                let position = TableCellPosition {
                    row: row_index,
                    column,
                    row_count,
                    column_count,
                };
                let resolve = |position| {
                    table_style
                        .map(|style: &TableStyle| style.resolve_cell(flags, position))
                        .unwrap_or_default()
                };
                let mut style = resolve(position);
                if span > 1 || row_span > 1 {
                    let far = resolve(TableCellPosition {
                        row: row_index + row_span - 1,
                        column: column + span - 1,
                        ..position
                    });
                    style.right = far.right;
                    style.bottom = far.bottom;
                }
                style.apply_cell_style(&direct_cell_style(cell));
                let inherited = style_text_body(&style);
                let height = self.cell_text_height(
                    &content,
                    cell_cascade(&cell.text, &inherited),
                    columns[column + span] - columns[column],
                )?;
                if row_span == 1 {
                    heights[row_index] = heights[row_index].max(height);
                } else {
                    spans.push((row_index, row_span, height));
                }
                plans.push(CellPlan {
                    text: &cell.text,
                    inherited,
                    content,
                    style,
                    row: row_index,
                    column,
                    span,
                    row_span,
                });
            }
        }
        for (row_index, row_span, height) in spans {
            let covered: f32 = heights[row_index..row_index + row_span].iter().sum();
            if height > covered {
                heights[row_index + row_span - 1] += height - covered;
            }
        }
        let mut rows = Vec::with_capacity(row_count + 1);
        let mut bottom = rect.y;
        rows.push(bottom);
        for height in &heights {
            bottom += height;
            rows.push(bottom);
        }
        let start = self.primitives.len();
        for plan in &plans {
            let cell = plan.rect(&columns, &rows);
            if let Some(paint) = plan
                .style
                .fill
                .as_ref()
                .and_then(|fill| paint(fill, self.theme))
            {
                self.primitives.push(cell_fill(object_id, cell, paint));
            }
        }
        for plan in &plans {
            let cell = plan.rect(&columns, &rows);
            let right = cell.x + cell.w;
            let foot = cell.y + cell.h;
            let edges = [
                (&plan.style.top, (cell.x, cell.y), (right, cell.y)),
                (&plan.style.bottom, (cell.x, foot), (right, foot)),
                (&plan.style.left, (cell.x, cell.y), (cell.x, foot)),
                (&plan.style.right, (right, cell.y), (right, foot)),
            ];
            for (outline, from, to) in edges {
                if let Some(stroke) = outline.as_ref().and_then(|line| stroke(line, self.theme)) {
                    self.primitives
                        .push(cell_border(object_id, from, to, stroke));
                }
            }
        }
        for plan in &plans {
            self.render_text_box(
                object_id,
                shape_id,
                plan.rect(&columns, &rows),
                Transform::default(),
                plan.content.clone(),
                cell_cascade(plan.text, &plan.inherited),
            )?;
        }
        let primitives = self.primitives.split_off(start);
        self.primitives.push(Primitive::Table {
            object_id,
            shape_id: Some(shape_id.to_owned()),
            name: name.to_owned(),
            x: rect.x,
            y: rect.y,
            w: rect.w,
            // The pivot for a rotation or flip is the centre of these bounds, so a
            // turned table keeps the frame's height and clips instead of moving.
            h: if transform.is_identity() {
                (bottom - rect.y).max(rect.h)
            } else {
                rect.h
            },
            label: format!("Table, {row_count} rows, {column_count} columns"),
            primitives,
            transform,
        });
        Ok(())
    }

    /// The height a cell's text needs at `width`, insets included. Vertical
    /// writing runs along the height it would be growing, so it keeps `a:tr/@h`
    /// rather than being measured across the cell.
    fn cell_text_height(
        &self,
        content: &TextContent,
        cascade: BodyCascade<'_>,
        width: f32,
    ) -> Result<f32, RenderError> {
        if TextFlow::from_body_vert(cascade.vertical()) != TextFlow::Horizontal {
            return Ok(0.0);
        }
        let resolved = resolve_content(self.renderer, self.theme, content, cascade)?;
        let left = cascade.inset_left().unwrap_or(DEFAULT_INSET_HORIZONTAL_EMU);
        let right = cascade
            .inset_right()
            .unwrap_or(DEFAULT_INSET_HORIZONTAL_EMU);
        let top = cascade.inset_top().unwrap_or(DEFAULT_INSET_VERTICAL_EMU);
        let bottom = cascade.inset_bottom().unwrap_or(DEFAULT_INSET_VERTICAL_EMU);
        let rect = PxRect {
            x: 0.0,
            y: 0.0,
            w: (width - emu_to_px(left + right)).max(1.0),
            h: 0.0,
        };
        let text = self.layout_text(&resolved, rect, 1.0, false)?;
        Ok(text.total_height + emu_to_px(top + bottom))
    }

    fn chart_space(&mut self, graphic: Option<&GraphicFrameData>) -> Option<&'a ChartSpace> {
        let GraphicFrameData::Chart {
            part_path: Some(part_path),
            ..
        } = graphic?
        else {
            return None;
        };
        let theme_part_path = self.theme_part_path;
        self.chart_parts
            .get_or_insert_with(|| {
                let mut index = HashMap::new();
                for part in &self.package.charts {
                    index
                        .entry((part.part_path.as_str(), part.theme_part_path.as_deref()))
                        .or_insert(part);
                }
                index
            })
            .get(&(part_path.as_str(), theme_part_path))
            .map(|part| &part.chart)
    }

    fn render_text_box(
        &mut self,
        object_id: u32,
        shape_id: &str,
        rect: PxRect,
        transform: Transform,
        content: TextContent,
        cascade: BodyCascade<'_>,
    ) -> Result<TextHit, RenderError> {
        let resolved = resolve_content(self.renderer, self.theme, &content, cascade)?;
        let flow = TextFlow::from_body_vert(cascade.vertical());
        let text_transform = text_transform(transform, flow);
        let text_rect = flow.layout_rect(rect);
        let left = cascade.inset_left().unwrap_or(DEFAULT_INSET_HORIZONTAL_EMU);
        let right = cascade
            .inset_right()
            .unwrap_or(DEFAULT_INSET_HORIZONTAL_EMU);
        let top = cascade.inset_top().unwrap_or(DEFAULT_INSET_VERTICAL_EMU);
        let bottom = cascade.inset_bottom().unwrap_or(DEFAULT_INSET_VERTICAL_EMU);
        let [left, top, right, bottom] = flow.layout_insets([left, top, right, bottom]);
        let content_rect = PxRect {
            x: text_rect.x + emu_to_px(left),
            y: text_rect.y + emu_to_px(top),
            w: (text_rect.w - emu_to_px(left + right)).max(1.0),
            h: (text_rect.h - emu_to_px(top + bottom)).max(1.0),
        };
        let scale = autofit_font_scale(cascade.autofit());
        let stacked = flow == TextFlow::Stacked;
        let mut laid_out = self.layout_text(&resolved, content_rect, scale, stacked)?;
        self.line_count += laid_out.lines.len();
        if self.line_count > MAX_TEXT_LINES {
            return Err(RenderError::ResourceLimit(format!(
                "more than {MAX_TEXT_LINES} text lines"
            )));
        }
        let anchor = match cascade.anchor() {
            Some("ctr") => TextAnchor::Center,
            Some("b") => TextAnchor::Bottom,
            _ => TextAnchor::Top,
        };
        let spare_height = content_rect.h - laid_out.total_height;
        let spare_height = if cascade.clips_overflow() {
            spare_height.max(0.0)
        } else {
            spare_height
        };
        let vertical_shift = match anchor {
            TextAnchor::Top => 0.0,
            TextAnchor::Center => spare_height / 2.0,
            TextAnchor::Bottom => spare_height,
        };
        for line in &mut laid_out.lines {
            shift_line(line, 0.0, vertical_shift);
        }
        if flow == TextFlow::VertLeftToRight {
            reverse_line_order(&mut laid_out.lines);
        }
        let display_paragraphs = resolved
            .paragraphs
            .iter()
            .map(|paragraph| TextParagraph {
                align: Some(paragraph.align),
                level: paragraph.level,
                runs: paragraph
                    .runs
                    .iter()
                    .map(|run| TextRun {
                        text: run.text.clone(),
                        font_family: run.style.family.clone(),
                        font_size_pt: autofit_size_pt(run.style.font_size_pt, scale),
                        bold: run.style.bold,
                        italic: run.style.italic,
                        underline: run.style.underline,
                        color: run.style.color.clone(),
                    })
                    .collect(),
            })
            .collect();
        let overflow = laid_out.total_height > content_rect.h && !cascade.clips_overflow();
        let story_id = content.story_id;
        let lines = laid_out.lines;
        self.primitives.push(Primitive::TextBox {
            object_id,
            shape_id: Some(shape_id.to_owned()),
            story_id: Some(story_id.clone()),
            x: text_rect.x,
            y: text_rect.y,
            w: text_rect.w,
            h: text_rect.h,
            anchor,
            paragraphs: display_paragraphs,
            lines: lines.clone(),
            overflow,
            transform: text_transform,
        });
        Ok(TextHit {
            overflow,
            story_id,
            rect: text_rect,
            transform: text_transform,
            lines,
        })
    }

    /// `layout_content` behind the renderer cache.
    fn layout_text(
        &self,
        content: &ResolvedContent,
        rect: PxRect,
        scale: f32,
        stacked: bool,
    ) -> Result<LayoutText, RenderError> {
        let key = text_layout_key(content, rect, scale, stacked);
        if let Some(text) = self
            .renderer
            .text_layouts
            .lock()
            .ok()
            .and_then(|cache| cache.get(&key))
        {
            return Ok(text);
        }
        let laid_out = layout_content(&self.renderer.fonts, content, rect, scale, stacked)?;
        if let Ok(mut cache) = self.renderer.text_layouts.lock() {
            cache.insert(key, &laid_out);
        }
        Ok(laid_out)
    }
}

/// One cell of a table, measured and ready to paint.
struct CellPlan<'a> {
    text: &'a TextBody,
    inherited: TextBody,
    content: TextContent,
    style: ResolvedCellStyle,
    row: usize,
    column: usize,
    span: usize,
    row_span: usize,
}

impl CellPlan<'_> {
    fn rect(&self, columns: &[f32], rows: &[f32]) -> PxRect {
        let x = columns[self.column];
        let y = rows[self.row];
        PxRect {
            x,
            y,
            w: columns[self.column + self.span] - x,
            h: rows[self.row + self.row_span] - y,
        }
    }
}

/// Column edges across the frame, scaling a declared grid that disagrees with
/// the frame's own width.
fn column_edges(grid: &[i64], column_count: usize, rect: PxRect) -> Vec<f32> {
    let declared: Vec<f64> = grid
        .iter()
        .map(|width| f64::from(emu_to_px(*width)))
        .collect();
    let widths = normalize_table_column_widths(&declared, column_count, f64::from(rect.w));
    let total: f64 = widths.iter().sum();
    let scale = if total > 0.0 {
        f64::from(rect.w) / total
    } else {
        0.0
    };
    let mut edges = Vec::with_capacity(column_count + 1);
    let mut offset = 0.0;
    edges.push(rect.x);
    for width in widths {
        offset += width * scale;
        edges.push(rect.x + safe_geometry(offset as f32));
    }
    edges
}

/// A cell's own `a:tcPr`, as the top of the style cascade.
fn direct_cell_style(cell: &TableCell) -> TableCellStyle {
    let edge = |outline: &Option<ShapeOutline>| {
        outline
            .clone()
            .map(|outline| TableCellBorder::Line(Box::new(outline)))
    };
    TableCellStyle {
        fill: cell.fill.clone(),
        borders: StyleCellBorders {
            left: edge(&cell.borders.left),
            right: edge(&cell.borders.right),
            top: edge(&cell.borders.top),
            bottom: edge(&cell.borders.bottom),
            inside_horizontal: None,
            inside_vertical: None,
        },
    }
}

/// The table style's text formatting, shaped as a body the cascade inherits from.
fn style_text_body(style: &ResolvedCellStyle) -> TextBody {
    TextBody {
        default_list_style: Some(Box::new(ParagraphProperties {
            default_run: Some(RunProperties {
                bold: style.bold,
                italic: style.italic,
                color: style.color.clone(),
                ..RunProperties::default()
            }),
            ..ParagraphProperties::default()
        })),
        ..TextBody::default()
    }
}

fn cell_cascade<'a>(text: &'a TextBody, inherited: &'a TextBody) -> BodyCascade<'a> {
    BodyCascade {
        primary: Some(text),
        layout: None,
        master: Some(inherited),
        master_slide: None,
        placeholder: None,
        style_color: None,
    }
}

fn cell_fill(object_id: u32, rect: PxRect, fill: Paint) -> Primitive {
    Primitive::Shape {
        clip: None,
        even_odd: false,
        object_id,
        shape_id: None,
        name: String::new(),
        x: rect.x,
        y: rect.y,
        w: rect.w,
        h: rect.h,
        geometry: "rect".to_owned(),
        path: geometry_path(
            "rect",
            &BTreeMap::new(),
            f64::from(rect.w) / f64::from(rect.h),
        ),
        geometry_fallback: false,
        adjust_values: BTreeMap::new(),
        fill: Some(fill),
        stroke: None,
        shadow: None,
        transform: Transform::default(),
    }
}

fn cell_border(object_id: u32, from: (f32, f32), to: (f32, f32), stroke: Stroke) -> Primitive {
    Primitive::Shape {
        clip: None,
        even_odd: false,
        object_id,
        shape_id: None,
        name: String::new(),
        x: from.0,
        y: from.1,
        w: to.0 - from.0,
        h: to.1 - from.1,
        geometry: "line".to_owned(),
        path: vec![
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            GeometryPathCommand::Line { x: 1.0, y: 1.0 },
        ],
        geometry_fallback: false,
        adjust_values: BTreeMap::new(),
        fill: None,
        stroke: Some(stroke),
        shadow: None,
        transform: Transform::default(),
    }
}

/// Text flow relative to the shape.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum TextFlow {
    Horizontal,
    Vert,
    Vert270,
    VertLeftToRight,
    Stacked,
}

impl TextFlow {
    fn from_body_vert(vertical: Option<&str>) -> Self {
        match vertical {
            Some("vert" | "eaVert") => Self::Vert,
            Some("vert270") => Self::Vert270,
            Some("mongolianVert") => Self::VertLeftToRight,
            Some("wordArtVert" | "wordArtVertRtl") => Self::Stacked,
            _ => Self::Horizontal,
        }
    }

    fn rotation_deg(self) -> f32 {
        match self {
            Self::Horizontal | Self::Stacked => 0.0,
            Self::Vert | Self::VertLeftToRight => 90.0,
            Self::Vert270 => -90.0,
        }
    }

    fn layout_rect(self, rect: PxRect) -> PxRect {
        match self {
            Self::Horizontal | Self::Stacked => rect,
            Self::Vert | Self::Vert270 | Self::VertLeftToRight => PxRect {
                x: rect.x + (rect.w - rect.h) / 2.0,
                y: rect.y + (rect.h - rect.w) / 2.0,
                w: rect.h,
                h: rect.w,
            },
        }
    }

    fn layout_insets(self, [left, top, right, bottom]: [i64; 4]) -> [i64; 4] {
        match self {
            Self::Horizontal | Self::Stacked => [left, top, right, bottom],
            Self::Vert | Self::VertLeftToRight => [top, right, bottom, left],
            Self::Vert270 => [bottom, left, top, right],
        }
    }
}

/// Mirrors the lines within the block they occupy, keeping the block in place.
fn reverse_line_order(lines: &mut [PositionedTextLine]) {
    let Some(top) = lines.iter().map(|line| line.y).reduce(f32::min) else {
        return;
    };
    let bottom = lines
        .iter()
        .map(|line| line.y + line.height)
        .fold(f32::MIN, f32::max);
    for line in lines {
        shift_line(line, 0.0, top + bottom - line.height - 2.0 * line.y);
    }
}

/// Removes glyph reflections and applies vertical flow.
fn text_transform(shape: Transform, flow: TextFlow) -> Transform {
    if flow == TextFlow::Horizontal && !shape.flip_v {
        return Transform {
            flip_h: false,
            ..shape
        };
    }
    let rotation_deg =
        shape.rotation_deg + if shape.flip_v { 180.0 } else { 0.0 } + flow.rotation_deg();
    Transform {
        rotation_deg: rotation_deg.rem_euclid(360.0),
        flip_h: false,
        flip_v: false,
    }
}

#[derive(Clone, Copy)]
struct BodyCascade<'a> {
    primary: Option<&'a TextBody>,
    layout: Option<&'a TextBody>,
    master: Option<&'a TextBody>,
    master_slide: Option<&'a SlideMaster>,
    placeholder: Option<&'a Placeholder>,
    style_color: Option<&'a ColorValue>,
}

impl BodyCascade<'_> {
    fn clips_overflow(&self) -> bool {
        let vertical = cascade_value(self.primary, self.layout, self.master, |body| {
            body.vertical_overflow
        });
        let horizontal = cascade_value(self.primary, self.layout, self.master, |body| {
            body.horizontal_overflow
        });
        [vertical, horizontal]
            .into_iter()
            .flatten()
            .any(|value| value != TextOverflow::Overflow)
    }

    fn anchor(&self) -> Option<&str> {
        self.primary
            .and_then(|body| body.anchor.as_deref())
            .or_else(|| self.layout.and_then(|body| body.anchor.as_deref()))
            .or_else(|| self.master.and_then(|body| body.anchor.as_deref()))
    }

    fn vertical(&self) -> Option<&str> {
        self.primary
            .and_then(|body| body.vertical.as_deref())
            .or_else(|| self.layout.and_then(|body| body.vertical.as_deref()))
            .or_else(|| self.master.and_then(|body| body.vertical.as_deref()))
    }

    fn autofit(&self) -> Option<&TextAutofit> {
        self.primary
            .and_then(|body| body.autofit.as_ref())
            .or_else(|| self.layout.and_then(|body| body.autofit.as_ref()))
            .or_else(|| self.master.and_then(|body| body.autofit.as_ref()))
    }

    fn inset_left(&self) -> Option<i64> {
        cascade_value(self.primary, self.layout, self.master, |body| {
            body.inset_left
        })
    }

    fn inset_top(&self) -> Option<i64> {
        cascade_value(self.primary, self.layout, self.master, |body| {
            body.inset_top
        })
    }

    fn inset_right(&self) -> Option<i64> {
        cascade_value(self.primary, self.layout, self.master, |body| {
            body.inset_right
        })
    }

    fn inset_bottom(&self) -> Option<i64> {
        cascade_value(self.primary, self.layout, self.master, |body| {
            body.inset_bottom
        })
    }

    /// `a:endParaRPr`: what PowerPoint sizes a paragraph by when it carries no
    /// runs of its own.
    fn end_run_properties(&self, index: usize) -> Option<&RunProperties> {
        [self.primary, self.layout, self.master]
            .into_iter()
            .flatten()
            .find_map(|body| {
                body.paragraphs
                    .get(index)
                    .and_then(|paragraph| paragraph.end_properties.as_ref())
            })
    }

    fn paragraph_properties(&self, index: usize, level: u32) -> ParagraphProperties {
        let mut properties = self
            .master_slide
            .and_then(|master| master_style(master, self.placeholder, level))
            .cloned()
            .unwrap_or_default();
        if let Some(color) = self.style_color {
            properties
                .default_run
                .get_or_insert_with(RunProperties::default)
                .color = Some(color.clone());
        }
        for body in [self.master, self.layout, self.primary]
            .into_iter()
            .flatten()
        {
            if let Some(source) = &body.default_list_style {
                merge_paragraph_properties(&mut properties, source);
            }
            if let Some(source) = body.list_style.get(level as usize) {
                merge_paragraph_properties(&mut properties, source);
            }
            if let Some(source) = body
                .paragraphs
                .get(index)
                .or_else(|| body.paragraphs.get(level as usize))
                .map(|paragraph| &paragraph.properties)
            {
                merge_paragraph_properties(&mut properties, source);
            }
        }
        if let Some(Bullet::AutoNumber { restart, .. }) = &mut properties.bullet {
            *restart = self
                .primary
                .and_then(|body| body.paragraphs.get(index))
                .is_some_and(|paragraph| {
                    matches!(
                        paragraph.properties.bullet,
                        Some(Bullet::AutoNumber { restart: true, .. })
                            | Some(Bullet::AutoNumber { start_at: 2.., .. })
                    )
                });
        }
        properties
    }
}

fn cascade_value<T: Copy>(
    primary: Option<&TextBody>,
    layout: Option<&TextBody>,
    master: Option<&TextBody>,
    get: impl Fn(&TextBody) -> Option<T>,
) -> Option<T> {
    primary
        .and_then(&get)
        .or_else(|| layout.and_then(&get))
        .or_else(|| master.and_then(get))
}

#[derive(Clone)]
struct TextContent {
    story_id: String,
    paragraphs: Vec<ContentParagraph>,
}

#[derive(Clone)]
struct ContentParagraph {
    alignment: Option<String>,
    level: u32,
    bullet: Option<Bullet>,
    runs: Vec<ContentRun>,
}

#[derive(Clone)]
struct ContentRun {
    text: String,
    style: TextStyle,
}

fn content_from_story(story: &StorySnapshot) -> TextContent {
    TextContent {
        story_id: story.id.clone(),
        paragraphs: story
            .paragraphs
            .iter()
            .map(|paragraph| ContentParagraph {
                alignment: paragraph.alignment.clone(),
                level: paragraph.level,
                bullet: paragraph
                    .bullet_json
                    .as_deref()
                    .and_then(|json| serde_json::from_str(json).ok()),
                runs: paragraph
                    .runs
                    .iter()
                    .map(|run| ContentRun {
                        text: run.text.clone(),
                        style: run.style.clone(),
                    })
                    .collect(),
            })
            .collect(),
    }
}

/// Resolves inherited slide-number fields.
fn field_text(run: &pptx_parse::TextRun, slide_number: i64) -> String {
    match run.field_type.as_deref() {
        Some("slidenum") => slide_number.to_string(),
        _ => run.text.clone(),
    }
}

fn content_from_body(
    story_id: &str,
    body: &TextBody,
    theme: &Theme,
    slide_number: i64,
) -> TextContent {
    TextContent {
        story_id: format!("inherited:{story_id}"),
        paragraphs: body
            .paragraphs
            .iter()
            .map(|paragraph| ContentParagraph {
                alignment: paragraph.properties.alignment.clone(),
                level: paragraph.properties.level,
                bullet: paragraph.properties.bullet.clone(),
                runs: paragraph
                    .runs
                    .iter()
                    .map(|run| ContentRun {
                        text: field_text(run, slide_number),
                        style: style_from_properties(&run.properties, theme),
                    })
                    .collect(),
            })
            .collect(),
    }
}

struct ResolvedContent {
    paragraphs: Vec<ResolvedParagraph>,
}

struct ResolvedParagraph {
    align: TextAlign,
    justify: bool,
    level: u32,
    margin_left_px: f32,
    margin_right_px: f32,
    line_spacing: Option<LineSpacing>,
    space_before: Option<LineSpacing>,
    space_after: Option<LineSpacing>,
    line_space_reduction: f32,
    indent_px: f32,
    /// `a:tabLst` stops in pixels from the text area's left edge, ascending.
    tab_stops: Vec<f32>,
    /// Pitch of the implicit stops past the last declared one.
    default_tab_px: f32,
    marker: Option<String>,
    bullet_style: Option<ResolvedStyle>,
    runs: Vec<ResolvedRun>,
}

struct ResolvedRun {
    text: String,
    start: u32,
    style: ResolvedStyle,
}

#[derive(Clone)]
struct ResolvedStyle {
    face: FontFace,
    family: String,
    font_size_pt: f32,
    /// Size the line box and percentage line spacing read. Small caps shape
    /// smaller than the run was authored at without shortening its line.
    line_font_size_pt: f32,
    /// `spc`: tracking added after every cluster, in points.
    spacing_pt: f32,
    baseline_shift_px: f32,
    bold: bool,
    italic: bool,
    underline: bool,
    color: String,
    caps: TextCaps,
}

fn resolve_content(
    renderer: &SlideRenderer,
    theme: &Theme,
    content: &TextContent,
    cascade: BodyCascade<'_>,
) -> Result<ResolvedContent, RenderError> {
    let total_bytes = content
        .paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .map(|run| run.text.len())
        .sum::<usize>();
    let total_runs = content
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.runs.len())
        .sum::<usize>();
    if total_bytes > MAX_TEXT_BYTES {
        return Err(RenderError::ResourceLimit(format!(
            "text exceeds {MAX_TEXT_BYTES} bytes"
        )));
    }
    if content.paragraphs.len() > MAX_TEXT_PARAGRAPHS {
        return Err(RenderError::ResourceLimit(format!(
            "more than {MAX_TEXT_PARAGRAPHS} text paragraphs"
        )));
    }
    if total_runs > MAX_TEXT_RUNS {
        return Err(RenderError::ResourceLimit(format!(
            "more than {MAX_TEXT_RUNS} text runs"
        )));
    }
    let line_space_reduction = autofit_line_space_reduction(cascade.autofit());
    let mut story_offset = 0_u32;
    let mut paragraphs = Vec::with_capacity(content.paragraphs.len());
    let mut numbering = AutoNumbering::default();
    for (index, paragraph) in content.paragraphs.iter().enumerate() {
        let mut properties = cascade.paragraph_properties(index, paragraph.level);
        if matches!(paragraph.bullet, Some(Bullet::AutoNumber { .. })) {
            properties.bullet = paragraph.bullet.clone();
            if let Some(Bullet::AutoNumber {
                restart, start_at, ..
            }) = &mut properties.bullet
            {
                *restart |= *start_at != 1;
            }
        }
        let language = properties
            .default_run
            .as_ref()
            .and_then(|value| value.language.as_deref());
        let mut runs = Vec::with_capacity(paragraph.runs.len().max(1));
        for run in &paragraph.runs {
            let style =
                resolve_style(renderer, theme, &run.style, properties.default_run.as_ref())?;
            let start = story_offset;
            story_offset = story_offset.saturating_add(utf16_len(&run.text));
            push_cased_runs(&mut runs, &run.text, start, language, style);
        }
        if runs.is_empty() {
            let end_style = cascade
                .end_run_properties(index)
                .map(|end| style_from_properties(end, theme))
                .unwrap_or_default();
            runs.push(ResolvedRun {
                text: String::new(),
                start: story_offset,
                style: resolve_style(renderer, theme, &end_style, properties.default_run.as_ref())?,
            });
        }
        let alignment = paragraph
            .alignment
            .as_deref()
            .or(properties.alignment.as_deref());
        // A blank paragraph between list items is spacing, not an item:
        // PowerPoint neither marks it nor counts it towards the next number.
        let marker = paragraph
            .runs
            .iter()
            .any(|run| !run.text.is_empty())
            .then(|| resolve_marker(properties.bullet.as_ref(), paragraph.level, &mut numbering))
            .flatten()
            .map(|marker| symbol_bullet(&marker, properties.bullet_font.as_ref(), theme));
        paragraphs.push(ResolvedParagraph {
            align: parse_align(alignment),
            justify: is_full_justification(alignment),
            level: paragraph.level,
            margin_left_px: emu_to_px(properties.margin_left.unwrap_or_default()),
            margin_right_px: emu_to_px(properties.margin_right.unwrap_or_default()),
            line_spacing: properties.line_spacing,
            space_before: properties.space_before,
            space_after: properties.space_after,
            line_space_reduction,
            indent_px: emu_to_px(properties.indent.unwrap_or_default()),
            tab_stops: resolve_tab_stops(
                properties.tab_stops.as_deref(),
                properties.margin_left.unwrap_or_default(),
                properties.indent.unwrap_or_default(),
            ),
            default_tab_px: resolve_default_tab(properties.default_tab_size),
            bullet_style: marker
                .is_some()
                .then(|| resolve_bullet_style(renderer, theme, &properties, &runs[0].style))
                .transpose()?,
            marker,
            runs,
        });
        story_offset = story_offset.saturating_add(1);
    }
    Ok(ResolvedContent { paragraphs })
}

/// Cases one run for drawing. `a:rPr/@cap` is a display property: the stored
/// run text keeps the author's casing, so only the runs handed to the shaper
/// change and the save projection is untouched.
fn push_cased_runs(
    out: &mut Vec<ResolvedRun>,
    text: &str,
    start: u32,
    language: Option<&str>,
    style: ResolvedStyle,
) {
    if text.is_empty() || style.caps == TextCaps::None {
        out.push(ResolvedRun {
            text: text.to_owned(),
            start,
            style,
        });
        return;
    }
    if style.caps == TextCaps::All {
        out.push(ResolvedRun {
            text: display_uppercase(text, language),
            start,
            style,
        });
        return;
    }
    let mut small = style.clone();
    small.font_size_pt = style.font_size_pt * WORD_SMALL_CAPS_ADVANCE_SCALE;
    let mut offset = start;
    for (lowercase, segment) in case_segments(text) {
        let cased = display_uppercase(segment, language);
        let length = utf16_len(&cased);
        out.push(ResolvedRun {
            text: cased,
            start: offset,
            style: if lowercase {
                small.clone()
            } else {
                style.clone()
            },
        });
        offset = offset.saturating_add(length);
    }
}

/// Uppercases for drawing only, keeping every character's UTF-16 width so the
/// display list still reports the authored story offsets. A character whose
/// uppercase form is wider (ß → SS) stays as authored.
fn display_uppercase(text: &str, language: Option<&str>) -> String {
    text.chars()
        .map(|character| {
            let mut upper = uppercase_for_language(character, language).into_iter();
            match (upper.next(), upper.next()) {
                (Some(single), None) if single.len_utf16() == character.len_utf16() => single,
                _ => character,
            }
        })
        .collect()
}

/// Splits `text` where its characters stop being lowercase, so small caps can
/// draw the lowercase stretches at a reduced size.
fn case_segments(text: &str) -> Vec<(bool, &str)> {
    let mut segments = Vec::new();
    let mut start = 0;
    let mut current = None;
    for (index, character) in text.char_indices() {
        let lowercase = character.is_lowercase();
        if current != Some(lowercase) {
            if let Some(previous) = current {
                segments.push((previous, &text[start..index]));
            }
            start = index;
            current = Some(lowercase);
        }
    }
    if let Some(previous) = current {
        segments.push((previous, &text[start..]));
    }
    segments
}

fn resolve_bullet_style(
    renderer: &SlideRenderer,
    theme: &Theme,
    properties: &ParagraphProperties,
    text: &ResolvedStyle,
) -> Result<ResolvedStyle, RenderError> {
    let mut style = text.clone();
    if let Some(BulletFont::Typeface(family)) = &properties.bullet_font {
        let family = if family.starts_with('+') {
            resolve_theme_font_ref(Some(theme), family)
        } else {
            family.clone()
        };
        style.face = renderer.resolve_face(&family, style.bold, style.italic)?;
        style.family = style.face.family.clone();
    }
    if let Some(BulletColor::Color(color)) = &properties.bullet_color
        && let Some(color) = resolve_color_value_to_hex_with_theme(Some(color), Some(theme))
    {
        style.color = color;
    }
    style.font_size_pt = match properties.bullet_size {
        Some(BulletSize::Percent(size)) => text.font_size_pt * size as f32,
        Some(BulletSize::Points(size)) => size as f32,
        _ => text.font_size_pt,
    }
    .clamp(1.0, 4_096.0);
    Ok(style)
}

fn resolve_style(
    renderer: &SlideRenderer,
    theme: &Theme,
    direct: &TextStyle,
    fallback: Option<&RunProperties>,
) -> Result<ResolvedStyle, RenderError> {
    let bold = direct
        .bold
        .or_else(|| fallback.and_then(|value| value.bold))
        .unwrap_or(false);
    let italic = direct
        .italic
        .or_else(|| fallback.and_then(|value| value.italic))
        .unwrap_or(false);
    let family = direct
        .font_family
        .as_deref()
        .filter(|family| family.len() <= 256)
        .or_else(|| {
            fallback
                .and_then(|value| value.font_family.as_deref())
                .filter(|family| family.len() <= 256)
        })
        .map(|family| {
            if family.starts_with('+') {
                resolve_theme_font_ref(Some(theme), family)
            } else {
                family.to_owned()
            }
        })
        .unwrap_or_else(|| resolve_theme_font_ref(Some(theme), "+mn-lt"));
    let face = renderer.resolve_face(&family, bold, italic)?;
    let color = direct
        .color
        .as_deref()
        .filter(|color| valid_color(color))
        .map(str::to_owned)
        .or_else(|| {
            fallback.and_then(|value| {
                resolve_color_value_to_hex_with_theme(value.color.as_ref(), Some(theme))
            })
        })
        .unwrap_or_else(|| "#000000".to_owned());
    let font_size_pt = direct
        .font_size_pt
        .map(|value| value as f32)
        .or_else(|| fallback.and_then(|value| value.font_size_pt.map(|value| value as f32)))
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(DEFAULT_FONT_SIZE_PT)
        .min(4_096.0);
    let spacing_pt = direct
        .spacing_pt
        .or_else(|| fallback.and_then(|value| value.spacing_pt))
        .map(|value| value as f32)
        .filter(|value| value.is_finite())
        .unwrap_or(0.0)
        .clamp(-4_096.0, 4_096.0);
    let baseline_pct = direct
        .baseline_pct
        .or_else(|| fallback.and_then(|value| value.baseline_pct))
        .map(|value| value as f32)
        .filter(|value| value.is_finite())
        .unwrap_or(0.0);
    let baseline_shift_px = points_to_px(font_size_pt) * baseline_pct / 100.0;
    let font_size_pt = if baseline_pct == 0.0 {
        font_size_pt
    } else {
        font_size_pt * SCRIPT_SIZE_RATIO
    };
    Ok(ResolvedStyle {
        family: face.family.clone(),
        face,
        font_size_pt,
        line_font_size_pt: font_size_pt,
        spacing_pt,
        baseline_shift_px,
        bold,
        italic,
        underline: direct
            .underline
            .as_deref()
            .or_else(|| fallback.and_then(|value| value.underline.as_deref()))
            .is_some_and(|value| value != "none"),
        color,
        caps: direct
            .caps
            .or_else(|| fallback.and_then(|value| value.caps))
            .unwrap_or(TextCaps::None),
    })
}

/// Characters used to check whether a registered face matches a metric table.
const ADVANCE_SAMPLE: &str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 .,";

/// Preserve the face's kerning when its individual advances already match.
fn runs_at_own_widths(fonts: &FontStore, id: FontId, metrics: &FamilyMetrics) -> bool {
    let Ok(own) = fonts.metrics(id) else {
        return true;
    };
    let units = f32::from(own.units_per_em);
    if units <= 0.0 {
        return true;
    }
    for character in ADVANCE_SAMPLE.chars() {
        let Some(wanted) = family_advance(metrics, character) else {
            continue;
        };
        let Ok(Some(advance)) = fonts.advance_width(id, character) else {
            return false;
        };
        if (wanted - advance / units).abs() > 0.002 {
            return false;
        }
    }
    true
}

/// Whether `id` already stacks its lines the way `metrics` does.
fn sits_on_own_baseline(fonts: &FontStore, id: FontId, metrics: &FamilyMetrics) -> bool {
    let Ok(own) = fonts.metrics(id) else {
        return true;
    };
    let units = f32::from(own.units_per_em);
    if units <= 0.0 {
        return true;
    }
    let same =
        |theirs: i32, ours: i16| (theirs as f32 / 1000.0 - f32::from(ours) / units).abs() <= 0.002;
    same(metrics.ascent, own.hhea_ascender)
        && same(metrics.descent, own.hhea_descender)
        && same(metrics.line_gap, own.hhea_line_gap)
}

/// The line box the named family measures at `size_px`, in the shape
/// [`single_line_box`] gives a registered face.
fn family_line_box(metrics: &FamilyMetrics, size_px: f32) -> ooxml_text::LineBox {
    let em = size_px / 1000.0;
    ooxml_text::LineBox {
        ascent: (metrics.ascent + metrics.line_gap).max(0) as f32 * em,
        descent: (-metrics.descent).max(0) as f32 * em,
        leading: 0.0,
    }
}

/// How wide the family a run names draws `text`, when that family is one we
/// have widths but no face for. `None` for anything the table cannot measure,
/// which keeps the substitute's own advance.
fn named_cluster_width(style: &ResolvedStyle, text: &str, size_px: f32) -> Option<f32> {
    let metrics = style.face.widths?;
    let mut total = 0.0;
    for character in text.chars() {
        total += family_advance(metrics, character)?;
    }
    Some(total * size_px)
}

/// Optional ligatures are off once glyphs are tracked apart.
fn tracking_features(tracking: f32) -> &'static [ShapeFeature] {
    const OFF: [ShapeFeature; 2] = [
        ShapeFeature {
            tag: *b"liga",
            value: 0,
        },
        ShapeFeature {
            tag: *b"clig",
            value: 0,
        },
    ];
    if tracking == 0.0 { &[] } else { &OFF }
}

/// One shaped line of chart text, in the family, weight, slant and pixel size
/// the plot geometry asked for.
fn chart_text_primitive(
    renderer: &SlideRenderer,
    theme: &Theme,
    shape_id: &str,
    text: ChartText<'_>,
) -> Result<Primitive, RenderError> {
    let bold = text.font.weight >= 600;
    let italic = text.font.italic;
    let family = if text.font.family.starts_with('+') {
        resolve_theme_font_ref(Some(theme), &text.font.family)
    } else {
        text.font.family.clone()
    };
    let face = renderer.resolve_face(&family, bold, italic)?;
    let size_px = safe_geometry(text.font.size_px as f32).clamp(1.0, 4_096.0);
    let tracking = safe_geometry(text.font.letter_spacing_px as f32);
    let shaped = shape(
        &renderer.fonts,
        face.id,
        text.text,
        size_px,
        tracking_features(tracking),
    )
    .map_err(|error| RenderError::Font(error.to_string()))?;
    let metrics = renderer
        .fonts
        .metrics(face.id)
        .map_err(|error| RenderError::Font(error.to_string()))?;
    let line_box = single_line_box(metrics, size_px, &CompatFlags::default());
    let mut offsets = Vec::with_capacity(shaped.len());
    let mut cursor = 0.0_f32;
    let mut cluster_advance = 0.0_f32;
    let mut cluster = None;
    for glyph in &shaped {
        if cluster.is_some_and(|previous| previous != glyph.cluster) {
            cursor += tracking.max(-cluster_advance);
            cluster_advance = 0.0;
        }
        offsets.push(cursor);
        cursor += glyph.x_advance;
        cluster_advance += glyph.x_advance;
        cluster = Some(glyph.cluster);
    }
    let advance = cursor;
    let box_x = safe_geometry(text.x as f32);
    let box_w = safe_geometry(text.width as f32);
    let align = match text.align {
        PlotTextAlign::Center => TextAlign::Center,
        PlotTextAlign::Start => TextAlign::Left,
    };
    let x = match align {
        TextAlign::Center => box_x + ((box_w - advance) / 2.0).max(0.0),
        _ => box_x,
    };
    let baseline = safe_geometry(text.baseline_y as f32);
    let glyphs = shaped
        .iter()
        .zip(&offsets)
        .map(|(glyph, offset)| PositionedGlyph {
            glyph_id: glyph.glyph_id,
            cluster: glyph.cluster,
            x: x + offset,
            advance: glyph.x_advance,
            x_offset: glyph.x_offset,
            y_offset: baseline + glyph.y_offset,
        })
        .collect();
    let run = PositionedTextRun {
        text: text.text.to_owned(),
        start: 0,
        end: utf16_len(text.text),
        x,
        width: advance.max(0.0),
        font_id: face.id.to_u32(),
        font_family: face.family.clone(),
        font_size_px: size_px,
        bold,
        italic,
        underline: false,
        color: text.color.to_owned(),
        letter_spacing_px: tracking,
        baseline_offset_px: 0.0,
        glyphs,
    };
    let width = run.width;
    Ok(Primitive::TextBox {
        object_id: text.object_id,
        shape_id: Some(shape_id.to_owned()),
        story_id: None,
        x: box_x,
        y: baseline - line_box.ascent,
        w: box_w.max(x - box_x + width),
        h: line_box.height(),
        anchor: TextAnchor::Top,
        paragraphs: vec![TextParagraph {
            align: Some(align),
            level: 0,
            runs: vec![TextRun {
                text: text.text.to_owned(),
                font_family: face.family.clone(),
                font_size_pt: size_px * 72.0 / 96.0,
                bold,
                italic,
                underline: false,
                color: text.color.to_owned(),
            }],
        }],
        lines: vec![PositionedTextLine {
            x,
            y: baseline - line_box.ascent,
            width,
            height: line_box.height(),
            baseline,
            start: 0,
            end: run.end,
            runs: vec![run],
            caret_stops: Vec::new(),
        }],
        overflow: false,
        transform: Transform::default(),
    })
}

struct LayoutText {
    lines: Vec<PositionedTextLine>,
    total_height: f32,
}

/// Byte cap on retained text layouts; oldest entry evicts first.
const MAX_TEXT_LAYOUT_BYTES: usize = 16 * 1024 * 1024;

/// Laid-out text keyed by every `layout_content` input.
#[derive(Default)]
struct TextLayoutCache {
    entries: HashMap<Vec<u8>, CachedTextLayout>,
    order: VecDeque<Vec<u8>>,
    bytes: usize,
}

struct CachedTextLayout {
    lines: Vec<PositionedTextLine>,
    total_height: f32,
    bytes: usize,
}

impl TextLayoutCache {
    fn get(&self, key: &[u8]) -> Option<LayoutText> {
        self.entries.get(key).map(|entry| LayoutText {
            lines: entry.lines.clone(),
            total_height: entry.total_height,
        })
    }

    fn insert(&mut self, key: Vec<u8>, text: &LayoutText) {
        let bytes = key.len() * 2 + text_layout_bytes(&text.lines);
        if bytes > MAX_TEXT_LAYOUT_BYTES || self.entries.contains_key(&key) {
            return;
        }
        self.bytes += bytes;
        self.order.push_back(key.clone());
        self.entries.insert(
            key,
            CachedTextLayout {
                lines: text.lines.clone(),
                total_height: text.total_height,
                bytes,
            },
        );
        while self.bytes > MAX_TEXT_LAYOUT_BYTES {
            let Some(oldest) = self.order.pop_front() else {
                break;
            };
            if let Some(entry) = self.entries.remove(&oldest) {
                self.bytes -= entry.bytes;
            }
        }
    }

    fn clear(&mut self) {
        self.entries.clear();
        self.order.clear();
        self.bytes = 0;
    }
}

fn text_layout_bytes(lines: &[PositionedTextLine]) -> usize {
    lines
        .iter()
        .map(|line| {
            size_of::<PositionedTextLine>()
                + line.caret_stops.capacity() * size_of::<CaretStop>()
                + line
                    .runs
                    .iter()
                    .map(|run| {
                        size_of::<PositionedTextRun>()
                            + run.text.capacity()
                            + run.font_family.capacity()
                            + run.color.capacity()
                            + run.glyphs.capacity() * size_of::<PositionedGlyph>()
                    })
                    .sum::<usize>()
        })
        .sum()
}

/// Encodes every argument `layout_content` reads.
fn text_layout_key(content: &ResolvedContent, rect: PxRect, scale: f32, stacked: bool) -> Vec<u8> {
    let text_len: usize = content
        .paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.runs)
        .map(|run| run.text.len())
        .sum();
    let mut key = Vec::with_capacity(64 + text_len);
    key_f32(&mut key, rect.x);
    key_f32(&mut key, rect.y);
    key_f32(&mut key, rect.w);
    key_f32(&mut key, scale);
    key.push(u8::from(stacked));
    key_u32(&mut key, content.paragraphs.len() as u32);
    for paragraph in &content.paragraphs {
        key.push(match paragraph.align {
            TextAlign::Left => 0,
            TextAlign::Center => 1,
            TextAlign::Right => 2,
            TextAlign::Justify => 3,
        });
        key.push(u8::from(paragraph.justify));
        key_u32(&mut key, paragraph.level);
        key_f32(&mut key, paragraph.margin_left_px);
        key_f32(&mut key, paragraph.margin_right_px);
        key_f32(&mut key, paragraph.indent_px);
        key_f32(&mut key, paragraph.line_space_reduction);
        key_spacing(&mut key, &paragraph.line_spacing);
        key_spacing(&mut key, &paragraph.space_before);
        key_spacing(&mut key, &paragraph.space_after);
        key_opt_str(&mut key, &paragraph.marker);
        match &paragraph.bullet_style {
            Some(style) => {
                key.push(1);
                key_style(&mut key, style);
            }
            None => key.push(0),
        }
        key_u32(&mut key, paragraph.runs.len() as u32);
        for run in &paragraph.runs {
            key_u32(&mut key, run.start);
            key_str(&mut key, &run.text);
            key_style(&mut key, &run.style);
        }
    }
    key
}

fn key_u32(key: &mut Vec<u8>, value: u32) {
    key.extend_from_slice(&value.to_le_bytes());
}

fn key_f32(key: &mut Vec<u8>, value: f32) {
    key_u32(key, value.to_bits());
}

fn key_f64(key: &mut Vec<u8>, value: f64) {
    key.extend_from_slice(&value.to_bits().to_le_bytes());
}

fn key_str(key: &mut Vec<u8>, value: &str) {
    key_u32(key, value.len() as u32);
    key.extend_from_slice(value.as_bytes());
}

fn key_opt_str(key: &mut Vec<u8>, value: &Option<String>) {
    match value {
        Some(value) => {
            key.push(1);
            key_str(key, value);
        }
        None => key.push(0),
    }
}

fn key_spacing(key: &mut Vec<u8>, spacing: &Option<LineSpacing>) {
    match spacing {
        Some(LineSpacing::Percent { value }) => {
            key.push(1);
            key_f64(key, *value);
        }
        Some(LineSpacing::Points { value }) => {
            key.push(2);
            key_f64(key, *value);
        }
        None => key.push(0),
    }
}

fn key_style(key: &mut Vec<u8>, style: &ResolvedStyle) {
    key_u32(key, style.face.id.to_u32());
    key_str(key, &style.face.requested_family);
    key_str(key, &style.family);
    key_f32(key, style.font_size_pt);
    key_f32(key, style.spacing_pt);
    key_f32(key, style.baseline_shift_px);
    key.push(u8::from(style.bold));
    key.push(u8::from(style.italic));
    key.push(u8::from(style.underline));
    key_str(key, &style.color);
}

fn layout_content(
    fonts: &FontStore,
    content: &ResolvedContent,
    rect: PxRect,
    scale: f32,
    stacked: bool,
) -> Result<LayoutText, RenderError> {
    let mut lines = Vec::new();
    let mut y = rect.y;
    let mut previous: Option<&ResolvedParagraph> = None;
    for paragraph in &content.paragraphs {
        if let Some(previous) = previous {
            y += spacing_px(previous.space_after, previous, scale)
                + spacing_px(paragraph.space_before, paragraph, scale);
        }
        previous = Some(paragraph);
        let paragraph_x = rect.x + paragraph.margin_left_px.max(0.0);
        let paragraph_width =
            (rect.w - paragraph.margin_left_px.max(0.0) - paragraph.margin_right_px.max(0.0))
                .max(1.0);
        let mut paragraph_lines = layout_paragraph(
            fonts,
            paragraph,
            paragraph_x,
            y,
            paragraph_width,
            scale,
            stacked,
        )?;
        if let Some(last) = paragraph_lines.last() {
            y = last.y + last.height;
        }
        lines.append(&mut paragraph_lines);
    }
    Ok(LayoutText {
        total_height: (y - rect.y).max(0.0),
        lines,
    })
}

/// Height of a `spcBef` or `spcAft`, whose percentages measure a single line.
fn spacing_px(spacing: Option<LineSpacing>, paragraph: &ResolvedParagraph, scale: f32) -> f32 {
    let height = match spacing {
        Some(LineSpacing::Percent { value }) => {
            let size_pt = paragraph
                .runs
                .iter()
                .map(|run| run.style.font_size_pt)
                .fold(0.0_f32, f32::max);
            value as f32 * SINGLE_LINE_PITCH_EM * points_to_px(autofit_size_pt(size_pt, scale))
        }
        // An autofit font scale shrinks the text, never a spacing written in
        // points: PowerPoint keeps `a:spcBef`/`a:spcAft` at their stated size.
        Some(LineSpacing::Points { value }) => points_to_px(value as f32),
        None => 0.0,
    };
    if height.is_finite() {
        height.max(0.0)
    } else {
        0.0
    }
}

/// `a:pPr/@defTabSz` when the cascade declares none: one inch.
const DEFAULT_TAB_EMU: i64 = 914_400;

/// Pitch of the implicit tab stops, in pixels.
fn resolve_default_tab(size: Option<i64>) -> f32 {
    let pitch = emu_to_px(size.unwrap_or(DEFAULT_TAB_EMU));
    if pitch.is_finite() && pitch >= 1.0 {
        pitch
    } else {
        emu_to_px(DEFAULT_TAB_EMU)
    }
}

/// Declared stops in pixels, ascending, plus the implicit stop a hanging indent
/// puts at the paragraph's left margin — the one a leading tab lands on.
fn resolve_tab_stops(stops: Option<&[i64]>, margin_left: i64, indent: i64) -> Vec<f32> {
    let mut resolved = stops
        .unwrap_or_default()
        .iter()
        .map(|position| emu_to_px(*position))
        .filter(|position| position.is_finite() && *position > 0.0)
        .collect::<Vec<_>>();
    if indent < 0 {
        let hanging = emu_to_px(margin_left);
        if hanging.is_finite() && hanging > 0.0 {
            resolved.push(hanging);
        }
    }
    resolved.sort_by(f32::total_cmp);
    resolved.dedup();
    resolved
}

/// How far left of the margin PowerPoint starts the paragraph's first line. A
/// marker owns that space here, so only an unmarked hanging indent has any, and
/// only the first tab of the first line is measured against it.
fn hanging_space(paragraph: &ResolvedParagraph) -> f32 {
    if paragraph.marker.is_some() || !paragraph.indent_px.is_finite() {
        return 0.0;
    }
    (-paragraph.indent_px).clamp(0.0, paragraph.margin_left_px.max(0.0))
}

/// A stop the pen already sits on does not hold the tab.
const TAB_EPSILON_PX: f32 = 0.01;

/// Width one tab takes. The stop is chosen from `offset` — where PowerPoint's
/// pen would be — and the width is measured from `pen`, where ours is; the two
/// differ only for the tab that stands in a hanging indent. The first declared
/// stop past `offset` wins, otherwise the next multiple of the default pitch,
/// and a tab never moves backwards or reaches past the line.
fn tab_advance(offset: f32, pen: f32, stops: &[f32], default_px: f32, limit: f32) -> f32 {
    let limit = if limit.is_finite() {
        limit.max(0.0)
    } else {
        0.0
    };
    if !offset.is_finite() || !pen.is_finite() {
        return 0.0;
    }
    let next = stops
        .iter()
        .copied()
        .find(|stop| *stop > offset + TAB_EPSILON_PX)
        .unwrap_or_else(|| {
            let steps = (offset / default_px).floor() + 1.0;
            if steps.is_finite() {
                steps * default_px
            } else {
                offset + default_px
            }
        });
    (next - pen).clamp(0.0, limit)
}

fn layout_paragraph(
    fonts: &FontStore,
    paragraph: &ResolvedParagraph,
    x: f32,
    y: f32,
    width: f32,
    scale: f32,
    stacked: bool,
) -> Result<Vec<PositionedTextLine>, RenderError> {
    let mut clusters = shape_paragraph(fonts, paragraph, scale)?;
    if clusters.is_empty() {
        let style = &paragraph.runs[0].style;
        let line_box = spaced_line_box(
            style_line_box(fonts, style, scale)?,
            paragraph,
            points_to_px(autofit_size_pt(style.font_size_pt, scale)),
        );
        return Ok(vec![PositionedTextLine {
            x,
            y,
            width: 0.0,
            height: line_box.height(),
            baseline: y + line_box.ascent,
            start: paragraph.runs[0].start,
            end: paragraph.runs[0].start,
            runs: Vec::new(),
            caret_stops: vec![CaretStop {
                position: paragraph.runs[0].start,
                x,
            }],
        }]);
    }
    let ranges = if stacked {
        // A hard break shapes to no glyph, so stacking it would leave a blank cell.
        (0..clusters.len())
            .filter(|index| !clusters[*index].glyphs.is_empty())
            .map(|index| (index, index + 1))
            .collect()
    } else {
        wrap_clusters(
            &mut clusters,
            width,
            paragraph.margin_left_px.max(0.0),
            hanging_space(paragraph),
            &paragraph.tab_stops,
            paragraph.default_tab_px,
        )
    };
    let line_count = ranges.len();
    let mut output = Vec::with_capacity(line_count);
    let mut line_y = y;
    for (line_index, (start, end)) in ranges.into_iter().enumerate() {
        let slice = &clusters[start..end];
        let natural_width = line_advance(slice);
        let stretchable = paragraph.justify
            && line_index + 1 < line_count
            && !slice.last().is_some_and(|cluster| cluster.mandatory);
        let padding = if stretchable {
            justification_padding(&justify_clusters(slice), width)
        } else {
            vec![0.0; slice.len()]
        };
        let stretched = padding.iter().any(|value| *value > 0.0);
        let trailing = if stretched {
            slice
                .iter()
                .rev()
                .take_while(|cluster| cluster_is_blank(cluster))
                .count()
        } else {
            0
        };
        let visible = slice.len() - trailing;
        let advances = slice
            .iter()
            .enumerate()
            .map(|(index, cluster)| {
                if index < visible {
                    cluster.width + padding[index]
                } else {
                    0.0
                }
            })
            .collect::<Vec<_>>();
        let line_width = advances.iter().sum::<f32>() - trailing_tracking(&slice[..visible]);
        let line_x = match paragraph.align {
            TextAlign::Center => x + ((width - natural_width) / 2.0).max(0.0),
            TextAlign::Right => x + (width - natural_width).max(0.0),
            TextAlign::Left | TextAlign::Justify => x,
        };
        let (natural, extents) = clusters_line_box(fonts, slice, scale)?;
        let line_box = shifted_line_box(
            spaced_line_box(natural, paragraph, line_font_size_px(slice, scale)),
            extents,
        );
        let mut caret_stops = vec![CaretStop {
            position: slice[0].start,
            x: line_x,
        }];
        let mut cursor_x = line_x;
        for (cluster, advance) in slice.iter().zip(&advances) {
            cursor_x += advance;
            caret_stops.push(CaretStop {
                position: cluster.end,
                x: cursor_x,
            });
        }
        caret_stops.dedup_by(|left, right| {
            left.position == right.position && left.x.to_bits() == right.x.to_bits()
        });
        let runs = positioned_runs(slice, &advances, line_x, line_y + line_box.ascent, scale);
        output.push(PositionedTextLine {
            x: line_x,
            y: line_y,
            width: line_width,
            height: line_box.height(),
            baseline: line_y + line_box.ascent,
            start: slice[0].start,
            end: slice
                .last()
                .map(|cluster| cluster.end)
                .unwrap_or(slice[0].start),
            runs,
            caret_stops,
        });
        line_y += line_box.height();
    }
    prepend_bullet(fonts, paragraph, x, &mut output, scale)?;
    Ok(output)
}

/// Prepends a marker outside the story's character space.
fn prepend_bullet(
    fonts: &FontStore,
    paragraph: &ResolvedParagraph,
    x: f32,
    lines: &mut [PositionedTextLine],
    scale: f32,
) -> Result<(), RenderError> {
    let Some(value) = &paragraph.marker else {
        return Ok(());
    };
    let (Some(first), Some(style)) = (lines.first_mut(), paragraph.bullet_style.as_ref()) else {
        return Ok(());
    };
    if value.trim().is_empty() || first.runs.is_empty() {
        return Ok(());
    }
    let bullet_x = (x + paragraph.indent_px).max(x - paragraph.margin_left_px.max(0.0));
    let marker = ResolvedParagraph {
        align: paragraph.align,
        justify: false,
        level: paragraph.level,
        margin_left_px: 0.0,
        margin_right_px: 0.0,
        line_spacing: None,
        space_before: None,
        space_after: None,
        line_space_reduction: 0.0,
        indent_px: 0.0,
        tab_stops: Vec::new(),
        default_tab_px: resolve_default_tab(None),
        marker: None,
        bullet_style: None,
        runs: vec![ResolvedRun {
            text: value.clone(),
            start: paragraph.runs[0].start,
            style: style.clone(),
        }],
    };
    let clusters = shape_paragraph(fonts, &marker, scale)?;
    if clusters.is_empty() {
        return Ok(());
    }
    let advances = clusters
        .iter()
        .map(|cluster| cluster.width)
        .collect::<Vec<_>>();
    let mut runs = positioned_runs(&clusters, &advances, bullet_x, first.baseline, scale);
    for run in &mut runs {
        run.start = paragraph.runs[0].start;
        run.end = run.start;
        for glyph in &mut run.glyphs {
            glyph.cluster = run.start;
        }
    }
    runs.append(&mut first.runs);
    first.runs = runs;
    Ok(())
}

struct ShapedCluster {
    text: String,
    start: u32,
    end: u32,
    /// Glyph advances plus `tracking`.
    width: f32,
    /// Part of `width` that is the gap after the cluster, dropped at a line's end.
    tracking: f32,
    run_index: usize,
    style: ResolvedStyle,
    glyphs: Vec<ClusterGlyph>,
    break_after: bool,
    mandatory: bool,
    /// A `U+0009` cluster, whose width the tab stops decide per line.
    tab: bool,
}

struct ClusterGlyph {
    glyph_id: u32,
    cluster: u32,
    x: f32,
    advance: f32,
    x_offset: f32,
    y_offset: f32,
}

fn shape_paragraph(
    fonts: &FontStore,
    paragraph: &ResolvedParagraph,
    scale: f32,
) -> Result<Vec<ShapedCluster>, RenderError> {
    let full_text = paragraph
        .runs
        .iter()
        .map(|run| run.text.as_str())
        .collect::<String>();
    let breaks = break_opportunities(&full_text)
        .into_iter()
        .map(|value| (value.byte_index, value.mandatory))
        .collect::<HashMap<_, _>>();
    let mut clusters = Vec::new();
    let mut global_byte = 0_usize;
    for (run_index, run) in paragraph.runs.iter().enumerate() {
        let mut segment_start = 0_usize;
        for (byte_index, character) in run.text.char_indices() {
            if character != '\n' && character != '\t' {
                continue;
            }
            add_shaped_segment(
                SegmentShape {
                    fonts,
                    run,
                    run_index,
                    text: &run.text[segment_start..byte_index],
                    run_byte_start: segment_start,
                    global_run_byte: global_byte,
                    scale,
                    breaks: &breaks,
                },
                &mut clusters,
            )?;
            let start = run.start + utf16_len(&run.text[..byte_index]);
            let hard = character == '\n';
            clusters.push(ShapedCluster {
                text: character.to_string(),
                start,
                end: start + 1,
                width: 0.0,
                tracking: 0.0,
                run_index,
                style: run.style.clone(),
                glyphs: Vec::new(),
                break_after: true,
                mandatory: hard,
                tab: !hard,
            });
            segment_start = byte_index + character.len_utf8();
        }
        add_shaped_segment(
            SegmentShape {
                fonts,
                run,
                run_index,
                text: &run.text[segment_start..],
                run_byte_start: segment_start,
                global_run_byte: global_byte,
                scale,
                breaks: &breaks,
            },
            &mut clusters,
        )?;
        global_byte += run.text.len();
    }
    Ok(clusters)
}

struct SegmentShape<'a> {
    fonts: &'a FontStore,
    run: &'a ResolvedRun,
    run_index: usize,
    text: &'a str,
    run_byte_start: usize,
    global_run_byte: usize,
    scale: f32,
    breaks: &'a HashMap<usize, bool>,
}

fn add_shaped_segment(
    request: SegmentShape<'_>,
    output: &mut Vec<ShapedCluster>,
) -> Result<(), RenderError> {
    let SegmentShape {
        fonts,
        run,
        run_index,
        text,
        run_byte_start,
        global_run_byte,
        scale,
        breaks,
    } = request;
    if text.is_empty() {
        return Ok(());
    }
    let size_px = points_to_px(autofit_size_pt(run.style.font_size_pt, scale));
    let tracking = points_to_px(run.style.spacing_pt * scale);
    let shaped = shape(
        fonts,
        run.style.face.id,
        text,
        size_px,
        tracking_features(tracking),
    )
    .map_err(|error| RenderError::Font(error.to_string()))?;
    let mut starts = shaped
        .iter()
        .map(|glyph| glyph.cluster as usize)
        .filter(|start| *start < text.len() && text.is_char_boundary(*start))
        .collect::<Vec<_>>();
    starts.push(0);
    starts.push(text.len());
    starts.sort_unstable();
    starts.dedup();
    for pair in starts.windows(2) {
        let start_byte = pair[0];
        let end_byte = pair[1];
        if start_byte == end_byte {
            continue;
        }
        let source_start = run.start + utf16_len(&run.text[..run_byte_start + start_byte]);
        let source_end = run.start + utf16_len(&run.text[..run_byte_start + end_byte]);
        let mut glyph_x = 0.0;
        let mut glyphs = Vec::new();
        for glyph in shaped
            .iter()
            .filter(|glyph| glyph.cluster as usize == start_byte)
        {
            glyphs.push(ClusterGlyph {
                glyph_id: glyph.glyph_id,
                cluster: source_start,
                x: glyph_x,
                advance: glyph.x_advance,
                x_offset: glyph.x_offset,
                y_offset: glyph.y_offset,
            });
            glyph_x += glyph.x_advance;
        }
        let text_slice = &text[start_byte..end_byte];
        // The glyphs keep the advances of the face that draws them, so a
        // backend laying the run out itself still sees where each one ends and
        // places the next cluster at the x this width put it.
        let cluster_width = named_cluster_width(&run.style, text_slice, size_px).unwrap_or(glyph_x);
        let global_end = global_run_byte + run_byte_start + end_byte;
        let tracking = tracking.max(-cluster_width);
        output.push(ShapedCluster {
            text: text_slice.to_owned(),
            start: source_start,
            end: source_end,
            width: cluster_width + tracking,
            tracking,
            run_index,
            style: run.style.clone(),
            glyphs,
            break_after: breaks.contains_key(&global_end),
            mandatory: breaks.get(&global_end).copied().unwrap_or(false),
            tab: false,
        });
    }
    Ok(())
}

/// Measures a line without its trailing tracking gap.
fn line_advance(clusters: &[ShapedCluster]) -> f32 {
    clusters.iter().map(|cluster| cluster.width).sum::<f32>() - trailing_tracking(clusters)
}

fn trailing_tracking(clusters: &[ShapedCluster]) -> f32 {
    clusters
        .iter()
        .rev()
        .find(|cluster| cluster.text != "\n")
        .map_or(0.0, |cluster| cluster.tracking)
}

/// Greedy line fill. Tab clusters take their width here, from the pen's offset
/// in the text area, so the wrap and the painted line agree on where a tab
/// lands: the greedy walk reaches a cluster in its final line's position last.
fn wrap_clusters(
    clusters: &mut [ShapedCluster],
    width: f32,
    left_offset: f32,
    hanging: f32,
    stops: &[f32],
    default_tab_px: f32,
) -> Vec<(usize, usize)> {
    let mut ranges = Vec::new();
    let mut start = 0;
    let mut line_index = 0;
    while start < clusters.len() {
        let mut cursor = start;
        let mut line_width = 0.0;
        let mut tabs_taken = 0_usize;
        let mut last_break = None;
        let mut end = clusters.len();
        while cursor < clusters.len() {
            if clusters[cursor].tab {
                let pen = left_offset + line_width;
                let measured = if line_index == 0 && tabs_taken == 0 {
                    pen - hanging
                } else {
                    pen
                };
                clusters[cursor].width =
                    tab_advance(measured, pen, stops, default_tab_px, width - line_width);
                tabs_taken += 1;
            }
            let cluster = &clusters[cursor];
            if cluster.text != "\n"
                && line_width + cluster.width - cluster.tracking > width
                && cursor > start
            {
                end = last_break
                    .filter(|candidate| *candidate > start)
                    .unwrap_or(cursor);
                break;
            }
            line_width += cluster.width;
            cursor += 1;
            if cluster.break_after {
                last_break = Some(cursor);
            }
            if cluster.mandatory {
                end = cursor;
                break;
            }
        }
        if cursor == clusters.len() {
            end = clusters.len();
        }
        if end <= start {
            end = start + 1;
        }
        ranges.push((start, end));
        start = end;
        line_index += 1;
    }
    ranges
}

fn positioned_runs(
    clusters: &[ShapedCluster],
    advances: &[f32],
    line_x: f32,
    baseline: f32,
    scale: f32,
) -> Vec<PositionedTextRun> {
    let mut output: Vec<PositionedTextRun> = Vec::new();
    let mut cursor_x = line_x;
    let mut trailing_tracking = 0.0_f32;
    for (index, cluster) in clusters.iter().enumerate() {
        if cluster.text == "\n" {
            continue;
        }
        let baseline_offset_px = cluster.style.baseline_shift_px * scale;
        let append = output.last().is_some_and(|run| {
            run.end == cluster.start
                && run.font_id == cluster.style.face.id.to_u32()
                && run.letter_spacing_px == cluster.tracking
                && run.baseline_offset_px == baseline_offset_px
                && run.font_family == cluster.style.family
                && run.font_size_px
                    == points_to_px(autofit_size_pt(cluster.style.font_size_pt, scale))
                && run.bold == cluster.style.bold
                && run.italic == cluster.style.italic
                && run.underline == cluster.style.underline
                && run.color == cluster.style.color
        });
        if !append {
            if let Some(previous) = output.last_mut() {
                previous.width -= trailing_tracking;
            }
            output.push(PositionedTextRun {
                text: String::new(),
                start: cluster.start,
                end: cluster.start,
                x: cursor_x,
                width: 0.0,
                font_id: cluster.style.face.id.to_u32(),
                font_family: cluster.style.family.clone(),
                font_size_px: points_to_px(autofit_size_pt(cluster.style.font_size_pt, scale)),
                bold: cluster.style.bold,
                italic: cluster.style.italic,
                underline: cluster.style.underline,
                color: cluster.style.color.clone(),
                letter_spacing_px: cluster.tracking,
                baseline_offset_px,
                glyphs: Vec::new(),
            });
        }
        let Some(run) = output.last_mut() else {
            continue;
        };
        run.text.push_str(&cluster.text);
        run.end = cluster.end;
        for glyph in &cluster.glyphs {
            run.glyphs.push(PositionedGlyph {
                glyph_id: glyph.glyph_id,
                cluster: glyph.cluster,
                x: cursor_x + glyph.x,
                advance: glyph.advance,
                x_offset: glyph.x_offset,
                y_offset: baseline - baseline_offset_px + glyph.y_offset,
            });
        }
        let advance = advances.get(index).copied().unwrap_or(cluster.width);
        run.width += advance;
        cursor_x += advance;
        trailing_tracking = cluster.tracking;
    }
    if let Some(last) = output.last_mut() {
        last.width -= trailing_tracking;
    }
    output
}

/// What justification needs to know about one cluster.
struct JustifyCluster {
    width: f32,
    break_after: bool,
    blank: bool,
}

fn justify_clusters(clusters: &[ShapedCluster]) -> Vec<JustifyCluster> {
    let last_visible = clusters
        .iter()
        .rposition(|cluster| !cluster_is_blank(cluster));
    clusters
        .iter()
        .enumerate()
        .map(|(index, cluster)| JustifyCluster {
            width: cluster.width
                - if Some(index) == last_visible {
                    cluster.tracking
                } else {
                    0.0
                },
            break_after: cluster.break_after,
            blank: cluster_is_blank(cluster),
        })
        .collect()
}

fn cluster_is_blank(cluster: &ShapedCluster) -> bool {
    !cluster.text.is_empty() && cluster.text.chars().all(char::is_whitespace)
}

/// Extra width to add after each cluster so the line fills `available`. Only
/// the break opportunities inside the line stretch, so the glyphs spread while
/// the words stay whole; trailing blanks are excluded from both the measurement
/// and the stretch, which lands the last glyph on the far edge.
fn justification_padding(clusters: &[JustifyCluster], available: f32) -> Vec<f32> {
    let mut padding = vec![0.0; clusters.len()];
    let trailing = clusters
        .iter()
        .rev()
        .take_while(|cluster| cluster.blank)
        .count();
    let visible = clusters.len() - trailing;
    let width: f32 = clusters[..visible]
        .iter()
        .map(|cluster| cluster.width)
        .sum();
    let extra = available - width;
    if !extra.is_finite() || extra <= 0.0 || visible == 0 {
        return padding;
    }
    let gaps: Vec<usize> = clusters[..visible]
        .iter()
        .enumerate()
        .take(visible.saturating_sub(1))
        .filter(|(_, cluster)| cluster.break_after)
        .map(|(index, _)| index)
        .collect();
    if gaps.is_empty() {
        return padding;
    }
    let share = extra / gaps.len() as f32;
    for index in gaps {
        padding[index] = share;
    }
    padding
}

fn clusters_line_box(
    fonts: &FontStore,
    clusters: &[ShapedCluster],
    scale: f32,
) -> Result<(ooxml_text::LineBox, ShiftExtents), RenderError> {
    let mut ascent: f32 = 0.0;
    let mut descent: f32 = 0.0;
    let mut leading: f32 = 0.0;
    let mut shifted_ascent: f32 = 0.0;
    let mut shifted_descent: f32 = 0.0;
    let mut seen = HashSet::new();
    for cluster in clusters {
        if !seen.insert(cluster.run_index) {
            continue;
        }
        let line = style_line_box(fonts, &cluster.style, scale)?;
        let shift = cluster.style.baseline_shift_px * scale;
        ascent = ascent.max(line.ascent);
        descent = descent.max(line.descent);
        leading = leading.max(line.leading);
        shifted_ascent = shifted_ascent.max((line.ascent + shift).max(0.0));
        shifted_descent = shifted_descent.max((line.descent - shift).max(0.0));
    }
    Ok((
        ooxml_text::LineBox {
            ascent,
            descent,
            leading,
        },
        ShiftExtents {
            ascent: (shifted_ascent - ascent).max(0.0),
            descent: (shifted_descent - descent).max(0.0),
        },
    ))
}

/// How far super/subscript ink reaches past the unshifted line box.
#[derive(Clone, Copy, Default)]
struct ShiftExtents {
    ascent: f32,
    descent: f32,
}

/// Spacing sets the pitch of the unshifted box; shifted ink then pushes the
/// edges back out so a raised or lowered run is never clipped.
fn shifted_line_box(line: ooxml_text::LineBox, extents: ShiftExtents) -> ooxml_text::LineBox {
    ooxml_text::LineBox {
        ascent: line.ascent + extents.ascent,
        descent: line.descent + extents.descent,
        leading: line.leading,
    }
}

fn style_line_box(
    fonts: &FontStore,
    style: &ResolvedStyle,
    scale: f32,
) -> Result<ooxml_text::LineBox, RenderError> {
    let size_px = points_to_px(autofit_size_pt(style.line_font_size_pt, scale));
    if let Some(named) = style.face.line {
        return Ok(family_line_box(named, size_px));
    }
    let metrics = fonts
        .metrics(style.face.id)
        .map_err(|error| RenderError::Font(error.to_string()))?;
    Ok(single_line_box(metrics, size_px, &CompatFlags::default()))
}

/// Largest font size on this line.
fn line_font_size_px(clusters: &[ShapedCluster], scale: f32) -> f32 {
    clusters
        .iter()
        .map(|cluster| points_to_px(autofit_size_pt(cluster.style.line_font_size_pt, scale)))
        .fold(0.0_f32, f32::max)
}

/// Round normal-autofit sizes to whole points, preserving authored sizes otherwise.
fn autofit_size_pt(size_pt: f32, scale: f32) -> f32 {
    if scale >= 1.0 {
        return size_pt;
    }
    (size_pt * scale).round().max(1.0)
}

/// `a:normAutofit/@fontScale`, applied verbatim: PowerPoint stores the scale it
/// computed when the text last changed and re-fits only on edit, never on render.
fn autofit_font_scale(autofit: Option<&TextAutofit>) -> f32 {
    match autofit {
        Some(TextAutofit::Normal { font_scale, .. }) => {
            font_scale.unwrap_or(1.0).clamp(0.1, 1.0) as f32
        }
        _ => 1.0,
    }
}

/// `a:normAutofit/@lnSpcReduction`, subtracted from percentage line spacing.
fn autofit_line_space_reduction(autofit: Option<&TextAutofit>) -> f32 {
    match autofit {
        Some(TextAutofit::Normal {
            line_space_reduction,
            ..
        }) => line_space_reduction.unwrap_or(0.0).clamp(0.0, 0.9) as f32,
        _ => 0.0,
    }
}

/// Apply paragraph spacing to a measured line box.
fn spaced_line_box(
    content: ooxml_text::LineBox,
    paragraph: &ResolvedParagraph,
    size_px: f32,
) -> ooxml_text::LineBox {
    if content.height() <= 0.0 {
        return content;
    }
    let reduction = paragraph.line_space_reduction;
    let single = SINGLE_LINE_PITCH_EM * size_px;
    let target = match paragraph.line_spacing {
        // An exact line spacing is a measurement, not a font size, so the
        // autofit scale leaves it alone.
        Some(LineSpacing::Points { value }) => points_to_px(value as f32),
        Some(LineSpacing::Percent { value }) => (value as f32 - reduction).max(0.0) * single,
        None => (1.0 - reduction) * single,
    };
    if !target.is_finite() || target < 0.0 {
        return content;
    }
    // PowerPoint splits the room a line box does not fill evenly above and
    // below it, so a paragraph whose spacing asks for more than its face
    // measures starts that much lower.
    if target >= content.height() {
        let slack = (target - content.ascent - content.descent) / 2.0;
        return ooxml_text::LineBox {
            ascent: content.ascent + slack,
            leading: slack,
            ..content
        };
    }
    scale_line_box(content, target / content.height())
}

fn scale_line_box(line: ooxml_text::LineBox, factor: f32) -> ooxml_text::LineBox {
    ooxml_text::LineBox {
        ascent: line.ascent * factor,
        descent: line.descent * factor,
        leading: line.leading * factor,
    }
}

fn shift_line(line: &mut PositionedTextLine, x: f32, y: f32) {
    line.x += x;
    line.y += y;
    line.baseline += y;
    for stop in &mut line.caret_stops {
        stop.x += x;
    }
    for run in &mut line.runs {
        run.x += x;
        for glyph in &mut run.glyphs {
            glyph.x += x;
            glyph.y_offset += y;
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct PxRect {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
}

impl PxRect {
    fn contains(self, x: f32, y: f32) -> bool {
        x >= self.x && x <= self.x + self.w && y >= self.y && y <= self.y + self.h
    }
}

#[derive(Clone, Copy)]
struct Space {
    origin_x: f32,
    origin_y: f32,
    scale_x: f32,
    scale_y: f32,
}

impl Space {
    fn root() -> Self {
        Self {
            origin_x: 0.0,
            origin_y: 0.0,
            scale_x: 1.0 / EMU_PER_CSS_PIXEL,
            scale_y: 1.0 / EMU_PER_CSS_PIXEL,
        }
    }

    fn map_transform(self, transform: &ShapeTransform) -> PxRect {
        PxRect {
            x: safe_geometry(self.origin_x + transform.x as f32 * self.scale_x),
            y: safe_geometry(self.origin_y + transform.y as f32 * self.scale_y),
            w: safe_geometry(transform.width as f32 * self.scale_x).abs(),
            h: safe_geometry(transform.height as f32 * self.scale_y).abs(),
        }
    }

    fn for_group(rect: PxRect, transform: &ShapeTransform) -> Self {
        let child_width = transform.child_width.unwrap_or(transform.width);
        let child_height = transform.child_height.unwrap_or(transform.height);
        if child_width == 0 || child_height == 0 {
            return Self::root();
        }
        let scale_x = safe_geometry(rect.w / child_width as f32);
        let scale_y = safe_geometry(rect.h / child_height as f32);
        let child_x = transform.child_x.unwrap_or_default() as f32;
        let child_y = transform.child_y.unwrap_or_default() as f32;
        Self {
            origin_x: safe_geometry(rect.x - child_x * scale_x),
            origin_y: safe_geometry(rect.y - child_y * scale_y),
            scale_x,
            scale_y,
        }
    }
}

struct HitRegion {
    shape_id: String,
    rect: PxRect,
    hit_rect: PxRect,
    transform: Transform,
    text: Option<TextHit>,
}

impl HitRegion {
    /// Maps slide coordinates into the shape frame.
    fn local_point(&self, x: f32, y: f32) -> (f32, f32) {
        local_point(self.rect, self.transform, x, y)
    }
}

struct TextHit {
    overflow: bool,
    story_id: String,
    rect: PxRect,
    transform: Transform,
    lines: Vec<PositionedTextLine>,
}

impl TextHit {
    /// Maps slide coordinates into the text frame.
    fn local_point(&self, x: f32, y: f32) -> (f32, f32) {
        local_point(self.rect, self.transform, x, y)
    }
}

/// Maps a slide point back into the unrotated frame `rect` was laid out in.
fn local_point(rect: PxRect, transform: Transform, x: f32, y: f32) -> (f32, f32) {
    if transform.is_identity() {
        return (x, y);
    }
    let center_x = rect.x + rect.w / 2.0;
    let center_y = rect.y + rect.h / 2.0;
    let (sin, cos) = transform.rotation_deg.to_radians().sin_cos();
    let dx = x - center_x;
    let dy = y - center_y;
    let mut local_x = dx * cos + dy * sin;
    let mut local_y = dy * cos - dx * sin;
    if transform.flip_h {
        local_x = -local_x;
    }
    if transform.flip_v {
        local_y = -local_y;
    }
    (center_x + local_x, center_y + local_y)
}

fn nearest_line(lines: &[PositionedTextLine], y: f32) -> Option<&PositionedTextLine> {
    lines.iter().min_by(|left, right| {
        distance_to_interval(y, left.y, left.y + left.height).total_cmp(&distance_to_interval(
            y,
            right.y,
            right.y + right.height,
        ))
    })
}

fn distance_to_interval(value: f32, start: f32, end: f32) -> f32 {
    if value < start {
        start - value
    } else if value > end {
        value - end
    } else {
        0.0
    }
}

fn find_node(nodes: &[ShapeNode], id: u32) -> Option<&ShapeNode> {
    for node in nodes {
        if node.id() == id {
            return Some(node);
        }
        if let ShapeNode::Group(group) = node
            && let Some(found) = find_node(&group.children, id)
        {
            return Some(found);
        }
    }
    None
}

fn find_placeholder<'a>(nodes: &'a [ShapeNode], target: &Placeholder) -> Option<&'a ShapeNode> {
    for node in nodes {
        if node_placeholder(node).is_some_and(|value| placeholders_match(value, target)) {
            return Some(node);
        }
        if let ShapeNode::Group(group) = node
            && let Some(found) = find_placeholder(&group.children, target)
        {
            return Some(found);
        }
    }
    None
}

fn placeholders_match(left: &Placeholder, right: &Placeholder) -> bool {
    match (left.index, right.index) {
        (Some(left), Some(right)) => left == right,
        _ => {
            normalize_placeholder_type(left.placeholder_type.as_deref())
                == normalize_placeholder_type(right.placeholder_type.as_deref())
        }
    }
}

fn normalize_placeholder_type(value: Option<&str>) -> &str {
    match value.unwrap_or("body") {
        "ctrTitle" => "title",
        "obj" => "body",
        value => value,
    }
}

fn node_base(node: &ShapeNode) -> &pptx_parse::ShapeBase {
    match node {
        ShapeNode::Shape(shape) => &shape.base,
        ShapeNode::Picture(shape) => &shape.base,
        ShapeNode::GraphicFrame(shape) => &shape.base,
        ShapeNode::Group(shape) => &shape.base,
    }
}

fn node_placeholder(node: &ShapeNode) -> Option<&Placeholder> {
    node_base(node).placeholder.as_ref()
}

fn node_fill(node: &ShapeNode) -> Option<&ShapeFill> {
    match node {
        ShapeNode::Shape(shape) => shape.fill.as_ref(),
        ShapeNode::Picture(shape) => shape.fill.as_ref(),
        ShapeNode::GraphicFrame(_) | ShapeNode::Group(_) => None,
    }
}

fn node_style(node: &ShapeNode) -> Option<&ShapeStyle> {
    match node {
        ShapeNode::Shape(shape) => shape.style.as_deref(),
        ShapeNode::Picture(picture) => picture.style.as_deref(),
        _ => None,
    }
}

fn merge_outline(direct: &ShapeOutline, fallback: Option<&ShapeOutline>) -> ShapeOutline {
    let Some(fallback) = fallback else {
        return direct.clone();
    };
    ShapeOutline {
        width: direct.width.or(fallback.width),
        color: direct.color.clone().or_else(|| {
            direct
                .gradient
                .is_none()
                .then(|| fallback.color.clone())
                .flatten()
        }),
        gradient: direct.gradient.clone().or_else(|| {
            direct
                .color
                .is_none()
                .then(|| fallback.gradient.clone())
                .flatten()
        }),
        style: direct.style.clone().or_else(|| fallback.style.clone()),
        cap: direct.cap.clone().or_else(|| fallback.cap.clone()),
        join: direct.join.clone().or_else(|| fallback.join.clone()),
        head_end: direct
            .head_end
            .clone()
            .or_else(|| fallback.head_end.clone()),
        tail_end: direct
            .tail_end
            .clone()
            .or_else(|| fallback.tail_end.clone()),
    }
}

fn node_outline(node: &ShapeNode) -> Option<&ShapeOutline> {
    match node {
        ShapeNode::Shape(shape) => shape.outline.as_ref(),
        ShapeNode::Picture(shape) => shape.outline.as_ref(),
        ShapeNode::GraphicFrame(_) | ShapeNode::Group(_) => None,
    }
}

fn source_rect(crop: Option<&PictureCrop>) -> Option<(f64, f64, f64, f64)> {
    let Some(crop) = crop else {
        return Some((0.0, 0.0, 1.0, 1.0));
    };
    let fraction = |value: i32| f64::from(value) / 100_000.0;
    let (left, top) = (fraction(crop.left), fraction(crop.top));
    let (width, height) = (
        1.0 - left - fraction(crop.right),
        1.0 - top - fraction(crop.bottom),
    );
    if width <= 0.0 || height <= 0.0 {
        return None;
    }
    Some((-left / width, -top / height, 1.0 / width, 1.0 / height))
}

/// Maps a metafile clip rectangle into the picture box, or `None` when the clip
/// leaves nothing of the box visible.
/// whether a metafile clip would hide any of `path`, in the normalised space
/// `metafile_clip_path` works in.
fn clipped_away(path: &[GeometryPathCommand], clip: [f64; 4], rect: (f64, f64, f64, f64)) -> bool {
    let Some(bounds) = metafile_clip_path(clip, rect) else {
        return true;
    };
    let (mut left, mut top, mut right, mut bottom) = (f64::MAX, f64::MAX, f64::MIN, f64::MIN);
    for command in &bounds {
        if let GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } = command {
            left = left.min(*x);
            top = top.min(*y);
            right = right.max(*x);
            bottom = bottom.max(*y);
        }
    }
    let outside = |x: &f64, y: &f64| *x < left || *x > right || *y < top || *y > bottom;
    path.iter().any(|command| match command {
        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => outside(x, y),
        GeometryPathCommand::Quad { cpx, cpy, x, y } => outside(cpx, cpy) || outside(x, y),
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => outside(cp1x, cp1y) || outside(cp2x, cp2y) || outside(x, y),
        GeometryPathCommand::Close => false,
    })
}

fn metafile_clip_path(
    clip: [f64; 4],
    rect: (f64, f64, f64, f64),
) -> Option<Vec<GeometryPathCommand>> {
    let (x0, y0, w, h) = rect;
    let left = (x0 + clip[0] * w).max(0.0);
    let top = (y0 + clip[1] * h).max(0.0);
    let right = (x0 + clip[2] * w).min(1.0);
    let bottom = (y0 + clip[3] * h).min(1.0);
    if !(left < right && top < bottom) {
        return None;
    }
    Some(vec![
        GeometryPathCommand::Move { x: left, y: top },
        GeometryPathCommand::Line { x: right, y: top },
        GeometryPathCommand::Line {
            x: right,
            y: bottom,
        },
        GeometryPathCommand::Line { x: left, y: bottom },
        GeometryPathCommand::Close,
    ])
}

/// Maps a path measured in the metafile's frame into the picture box.
fn place_in_source_rect(path: &mut [GeometryPathCommand], rect: (f64, f64, f64, f64)) {
    if rect == (0.0, 0.0, 1.0, 1.0) {
        return;
    }
    let (x0, y0, w, h) = rect;
    let place = |x: &mut f64, y: &mut f64| {
        *x = x0 + *x * w;
        *y = y0 + *y * h;
    };
    for command in path {
        match command {
            GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => place(x, y),
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                place(cpx, cpy);
                place(x, y);
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                place(cp1x, cp1y);
                place(cp2x, cp2y);
                place(x, y);
            }
            GeometryPathCommand::Close => {}
        }
    }
}

fn node_effects(node: &ShapeNode) -> Option<&ShapeEffects> {
    match node {
        ShapeNode::Shape(shape) => shape.effects.as_ref(),
        ShapeNode::Picture(shape) => shape.shape_effects.as_ref(),
        ShapeNode::GraphicFrame(_) | ShapeNode::Group(_) => None,
    }
}

fn node_text(node: &ShapeNode) -> Option<&TextBody> {
    match node {
        ShapeNode::Shape(shape) => shape.text.as_ref(),
        _ => None,
    }
}

fn node_group_transform(node: &ShapeNode) -> Option<&ShapeTransform> {
    match node {
        ShapeNode::Group(group) => Some(&group.base.transform),
        _ => None,
    }
}

fn master_style<'a>(
    master: &'a SlideMaster,
    placeholder: Option<&Placeholder>,
    level: u32,
) -> Option<&'a ParagraphProperties> {
    let styles = match placeholder {
        Some(placeholder) => {
            match normalize_placeholder_type(placeholder.placeholder_type.as_deref()) {
                "title" => &master.text_styles.title,
                "body" | "subTitle" => &master.text_styles.body,
                _ => &master.text_styles.other,
            }
        }
        None => &master.text_styles.other,
    };
    styles.get(level as usize).or_else(|| styles.first())
}

fn merge_paragraph_properties(target: &mut ParagraphProperties, source: &ParagraphProperties) {
    if source.alignment.is_some() {
        target.alignment.clone_from(&source.alignment);
    }
    if source.margin_left.is_some() {
        target.margin_left = source.margin_left;
    }
    if source.margin_right.is_some() {
        target.margin_right = source.margin_right;
    }
    if source.indent.is_some() {
        target.indent = source.indent;
    }
    if source.bullet.is_some() {
        target.bullet.clone_from(&source.bullet);
    }
    if source.line_spacing.is_some() {
        target.line_spacing = source.line_spacing;
    }
    if source.space_before.is_some() {
        target.space_before = source.space_before;
    }
    if source.space_after.is_some() {
        target.space_after = source.space_after;
    }
    if source.bullet_font.is_some() {
        target.bullet_font.clone_from(&source.bullet_font);
    }
    if source.bullet_color.is_some() {
        target.bullet_color.clone_from(&source.bullet_color);
    }
    if source.bullet_size.is_some() {
        target.bullet_size.clone_from(&source.bullet_size);
    }
    if source.default_tab_size.is_some() {
        target.default_tab_size = source.default_tab_size;
    }
    if source.tab_stops.is_some() {
        target.tab_stops.clone_from(&source.tab_stops);
    }
    if let Some(source) = &source.default_run {
        let target = target
            .default_run
            .get_or_insert_with(RunProperties::default);
        merge_run_properties(target, source);
    }
}

fn merge_run_properties(target: &mut RunProperties, source: &RunProperties) {
    if source.font_size_pt.is_some() {
        target.font_size_pt = source.font_size_pt;
    }
    if source.bold.is_some() {
        target.bold = source.bold;
    }
    if source.italic.is_some() {
        target.italic = source.italic;
    }
    if source.underline.is_some() {
        target.underline.clone_from(&source.underline);
    }
    if source.font_family.is_some() {
        target.font_family.clone_from(&source.font_family);
    }
    if source.color.is_some() {
        target.color.clone_from(&source.color);
    }
    if source.language.is_some() {
        target.language.clone_from(&source.language);
    }
    if source.spacing_pt.is_some() {
        target.spacing_pt = source.spacing_pt;
    }
    if source.baseline_pct.is_some() {
        target.baseline_pct = source.baseline_pct;
    }
    if source.caps.is_some() {
        target.caps = source.caps;
    }
}

fn style_from_properties(properties: &RunProperties, theme: &Theme) -> TextStyle {
    TextStyle {
        bold: properties.bold,
        italic: properties.italic,
        font_size_pt: properties.font_size_pt,
        color: resolve_color_value_to_hex_with_theme(properties.color.as_ref(), Some(theme)),
        font_family: properties.font_family.clone(),
        underline: properties.underline.clone(),
        spacing_pt: properties.spacing_pt,
        baseline_pct: properties.baseline_pct,
        caps: properties.caps,
    }
}

fn resolved_transform_value(
    shape: &ShapeSnapshot,
    original: Option<&ShapeNode>,
    layout: Option<&ShapeNode>,
    master: Option<&ShapeNode>,
) -> ShapeTransform {
    if shape.width > 0 && shape.height > 0 {
        ShapeTransform {
            x: shape.x,
            y: shape.y,
            width: shape.width,
            height: shape.height,
            rotation_deg: shape.rotation_deg,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
            ..ShapeTransform::default()
        }
    } else {
        [original, layout, master]
            .into_iter()
            .flatten()
            .map(|node| &node_base(node).transform)
            .find(|transform| transform.width > 0 && transform.height > 0)
            .cloned()
            .unwrap_or_else(|| ShapeTransform {
                x: shape.x,
                y: shape.y,
                width: shape.width,
                height: shape.height,
                rotation_deg: shape.rotation_deg,
                flip_h: shape.flip_h,
                flip_v: shape.flip_v,
                ..ShapeTransform::default()
            })
    }
}

/// Resolves bitmap effect colours against the theme.
fn image_effects(effects: &[BlipEffect], theme: &Theme) -> Vec<ImageEffect> {
    let rgba = |color: Option<&ColorValue>| resolve_color_value_to_rgba_hex(color, Some(theme));
    effects
        .iter()
        .filter_map(|effect| match effect {
            BlipEffect::BiLevel { threshold } => Some(ImageEffect::BiLevel {
                threshold: (*threshold as f32).clamp(0.0, 1.0),
            }),
            BlipEffect::Grayscale => Some(ImageEffect::Grayscale),
            BlipEffect::Luminance {
                brightness,
                contrast,
            } => Some(ImageEffect::Luminance {
                brightness: (*brightness as f32).clamp(-1.0, 1.0),
                contrast: (*contrast as f32).clamp(-1.0, 1.0),
            }),
            BlipEffect::Duotone { shadow, highlight } => Some(ImageEffect::Duotone {
                shadow: rgba(shadow.as_ref())?,
                highlight: rgba(highlight.as_ref())?,
            }),
            BlipEffect::ColorChange {
                from,
                to,
                use_alpha,
            } => Some(ImageEffect::ColorChange {
                from: rgba(from.as_ref())?,
                to: rgba(to.as_ref())?,
                use_alpha: *use_alpha,
            }),
        })
        .collect()
}

/// A fill or stroke colour, widened to `#RRGGBBAA` only when it is actually translucent.
fn resolve_paint_color(color: Option<&ColorValue>, theme: &Theme) -> Option<String> {
    let rgba = resolve_color_value_to_rgba_hex(color, Some(theme))?;
    match rgba.strip_suffix("FF") {
        Some(opaque) if rgba.len() == 9 => Some(opaque.to_owned()),
        _ => Some(rgba),
    }
}

fn paint(fill: &ShapeFill, theme: &Theme) -> Option<Paint> {
    if fill.fill_type == "none" {
        return None;
    }
    if let Some(paint) = fill
        .gradient
        .as_ref()
        .and_then(|gradient| gradient_paint(gradient, theme))
    {
        return Some(paint);
    }
    resolve_paint_color(fill.color.as_ref(), theme).map(|color| Paint::Solid { color })
}

fn shape_style_color(shape: Option<&ShapeNode>) -> Option<&ColorValue> {
    match shape? {
        ShapeNode::Shape(shape) => shape.style.as_ref()?.font_color.as_ref(),
        _ => None,
    }
}

fn line_end(end: Option<&LineEnd>, stroke_width: f32) -> Option<StrokeEnd> {
    let end = end.filter(|end| end.end_type != "none")?;
    let base = stroke_width.max(LINE_END_MIN_BASE_PX);
    Some(StrokeEnd {
        kind: end.end_type.clone(),
        width: base * line_end_scale(end.width.as_deref()),
        length: base * line_end_scale(end.length.as_deref()),
    })
}

fn line_end_scale(size: Option<&str>) -> f32 {
    match size {
        Some("sm") => 2.0,
        Some("lg") => 5.0,
        _ => 3.0,
    }
}

/// `a:outerShdw` in surface pixels, or `None` when it would paint nothing.
fn shadow(
    effects: &ShapeEffects,
    theme: &Theme,
    space: Space,
    rect: PxRect,
    rotation_deg: f32,
    flip_h: bool,
    flip_v: bool,
) -> Option<Shadow> {
    let outer = effects.outer_shadow.as_ref()?;
    let color = resolve_paint_color(outer.color.as_ref(), theme)?;
    if color.len() == 9 && color.ends_with("00") {
        return None;
    }
    let direction = (outer.direction as f64 / ANGLE_UNITS_PER_DEGREE).to_radians();
    let distance = outer.distance as f64;
    let mut dx = (distance * direction.cos()) as f32 * space.scale_x;
    let mut dy = (distance * direction.sin()) as f32 * space.scale_y;
    if outer.rotate_with_shape {
        dx *= if flip_h { -1.0 } else { 1.0 };
        dy *= if flip_v { -1.0 } else { 1.0 };
        let (sin, cos) = rotation_deg.to_radians().sin_cos();
        (dx, dy) = (cos * dx - sin * dy, sin * dx + cos * dy);
    }
    // Scaling happens about the surface origin, so the anchor `algn` names travels in the
    // offset: a point scaled about A lands at s * p + A * (1 - s).
    let scale_x = safe_scale(outer.scale_x as f32);
    let scale_y = safe_scale(outer.scale_y as f32);
    if scale_x == 0.0 || scale_y == 0.0 {
        return None;
    }
    let (anchor_x, anchor_y) = shadow_anchor(&outer.alignment, rect);
    Some(Shadow {
        paths: Vec::new(),
        color,
        blur: safe_geometry(outer.blur_radius as f32 * (space.scale_x + space.scale_y) / 2.0),
        dx: safe_geometry(dx + anchor_x * (1.0 - scale_x)),
        dy: safe_geometry(dy + anchor_y * (1.0 - scale_y)),
        scale_x,
        scale_y,
    })
}

fn safe_scale(value: f32) -> f32 {
    if value.is_finite() { value } else { 1.0 }
}

/// The point of `rect` that `a:algn` keeps fixed when the shadow is scaled.
fn shadow_anchor(alignment: &str, rect: PxRect) -> (f32, f32) {
    let x = match alignment {
        "tl" | "l" | "bl" => rect.x,
        "tr" | "r" | "br" => rect.x + rect.w,
        _ => rect.x + rect.w / 2.0,
    };
    let y = match alignment {
        "tl" | "t" | "tr" => rect.y,
        "l" | "ctr" | "r" => rect.y + rect.h / 2.0,
        _ => rect.y + rect.h,
    };
    (x, y)
}

fn gradient_paint(gradient: &GradientFill, theme: &Theme) -> Option<Paint> {
    let gradient_type = match gradient.gradient_type.as_str() {
        "radial" => GradientType::Radial,
        "rectangular" => GradientType::Rectangular,
        "path" => GradientType::Path,
        _ => GradientType::Linear,
    };
    let mut stops = gradient
        .stops
        .iter()
        .filter_map(|stop| {
            Some(GradientStop {
                position: (stop.position as f32 / 100_000.0).clamp(0.0, 1.0),
                color: resolve_paint_color(Some(&stop.color), theme)?,
            })
        })
        .collect::<Vec<_>>();
    stops.sort_by(|left, right| {
        left.position
            .partial_cmp(&right.position)
            .unwrap_or_else(|| left.position.total_cmp(&right.position))
    });
    if stops.is_empty() {
        return None;
    }
    Some(Paint::Gradient {
        gradient_type,
        angle_deg: gradient.angle.map(|value| value as f32),
        stops,
    })
}

fn stroke(outline: &ShapeOutline, theme: &Theme) -> Option<Stroke> {
    let solid = resolve_paint_color(outline.color.as_ref(), theme);
    let paint = match &solid {
        Some(_) => None,
        None => outline
            .gradient
            .as_ref()
            .and_then(|gradient| gradient_paint(gradient, theme)),
    };
    let color = match (&solid, &paint) {
        (Some(color), _) => color.clone(),
        (None, Some(Paint::Gradient { stops, .. })) => stops.first()?.color.clone(),
        (None, Some(Paint::Solid { color })) => color.clone(),
        (None, None) => return None,
    };
    let width = outline
        .width
        .filter(|width| width.is_finite() && *width >= 0.0)
        .map(|width| width as f32 / EMU_PER_CSS_PIXEL)
        .unwrap_or(1.0);
    Some(Stroke {
        join: outline.join.clone(),
        color,
        width,
        dashed: outline
            .style
            .as_deref()
            .is_some_and(|style| style != "solid"),
        paint,
        head_end: line_end(outline.head_end.as_ref(), width),
        tail_end: line_end(outline.tail_end.as_ref(), width),
    })
}

/// Looks up a snapshot picture's parsed source.
fn picture_source(shape: Option<&ShapeNode>) -> Option<&Picture> {
    match shape? {
        ShapeNode::Picture(picture) => Some(picture),
        _ => None,
    }
}

/// A picture's part path, or a marker resolving to its unsaved pending bytes.
fn picture_asset_id(shape: &ShapeSnapshot) -> Option<String> {
    shape.media_part_path.clone().or_else(|| {
        shape
            .pending_media
            .as_ref()
            .map(|_| format!("pending-media:{}", shape.id))
    })
}

/// Clamps outsets and discards empty crops.
fn image_crop(crop: &PictureCrop) -> ImageCrop {
    let fraction = |value: i32| (value as f32 / 100_000.0).clamp(0.0, 1.0);
    let cropped = ImageCrop {
        left: fraction(crop.left),
        top: fraction(crop.top),
        right: fraction(crop.right),
        bottom: fraction(crop.bottom),
    };
    let (kept_x, kept_y) = cropped.kept();
    if kept_x <= 0.0 || kept_y <= 0.0 {
        ImageCrop::default()
    } else {
        cropped
    }
}

#[derive(Default)]
struct PictureMask {
    path: Option<Vec<GeometryPathCommand>>,
    geometry_fallback: bool,
}

fn picture_mask(picture: &Picture, rect: PxRect) -> PictureMask {
    if picture.geometry.is_empty() || picture.geometry == "rect" || rect.h <= 0.0 {
        return PictureMask::default();
    }
    let (path, geometry_fallback) = geometry_path_with_fallback(
        &picture.geometry,
        &picture.adjust_values,
        f64::from(rect.w) / f64::from(rect.h),
    );
    PictureMask {
        path: (!path.is_empty()).then_some(path),
        geometry_fallback,
    }
}

fn custom_paths(shape: Option<&ShapeNode>) -> &[CustomGeometryPath] {
    match shape {
        Some(ShapeNode::Shape(shape)) => &shape.paths,
        _ => &[],
    }
}

/// Looks up the blip behind a snapshot shape's picture fill.
fn picture_fill<'a>(nodes: &[Option<&'a ShapeNode>]) -> Option<&'a PictureFill> {
    match nodes
        .iter()
        .flatten()
        .find(|node| node_fill(node).is_some())?
    {
        ShapeNode::Shape(shape) => shape.picture_fill.as_deref(),
        _ => None,
    }
}

/// Redraws a picture-filled shape as an image masked by the shape's own outline.
fn picture_filled(primitive: Primitive, picture: Option<&PictureFill>) -> Primitive {
    let Some(picture) = picture else {
        return primitive;
    };
    match primitive {
        Primitive::Shape {
            object_id,
            shape_id,
            name,
            x,
            y,
            w,
            h,
            geometry,
            path,
            geometry_fallback,
            stroke,
            shadow,
            transform,
            ..
        } => Primitive::Image {
            object_id,
            shape_id,
            name,
            x,
            y,
            w,
            h,
            asset_id: picture.media_part_path.clone(),
            effects: Vec::new(),
            crop: picture_fill_crop(picture),
            path: (geometry != "rect").then_some(path),
            geometry_fallback,
            stroke,
            shadow,
            transform,
        },
        other => other,
    }
}

/// Folds `a:srcRect` and the `a:stretch/a:fillRect` band into source fractions.
/// Insets that letterbox the image rather than crop it stretch to the box instead.
fn picture_fill_crop(picture: &PictureFill) -> ImageCrop {
    let source = image_crop(&picture.crop);
    let (kept_x, kept_y) = source.kept();
    let band = |near: i32, far: i32| {
        let (near, far) = (near as f32 / 100_000.0, far as f32 / 100_000.0);
        let span = 1.0 - near - far;
        if span <= 0.0 {
            return (0.0, 0.0);
        }
        ((-near).max(0.0) / span, (-far).max(0.0) / span)
    };
    let (left, right) = band(picture.fill_rect.left, picture.fill_rect.right);
    let (top, bottom) = band(picture.fill_rect.top, picture.fill_rect.bottom);
    let cropped = ImageCrop {
        left: source.left + left * kept_x,
        top: source.top + top * kept_y,
        right: source.right + right * kept_x,
        bottom: source.bottom + bottom * kept_y,
    };
    let (kept_x, kept_y) = cropped.kept();
    if kept_x <= 0.0 || kept_y <= 0.0 {
        source
    } else {
        cropped
    }
}

fn geometry_path(
    geometry: &str,
    adjustments: &BTreeMap<String, f64>,
    aspect_ratio: f64,
) -> Vec<ooxml_drawingml::GeometryPathCommand> {
    geometry_path_with_fallback(geometry, adjustments, aspect_ratio).0
}

fn geometry_path_with_fallback(
    geometry: &str,
    adjustments: &BTreeMap<String, f64>,
    aspect_ratio: f64,
) -> (Vec<ooxml_drawingml::GeometryPathCommand>, bool) {
    let adjustments = adjustments
        .iter()
        .map(|(name, value)| (name.clone(), *value))
        .collect();
    match preset_geometry_to_path(geometry, &adjustments, aspect_ratio) {
        Some(path) => (path, false),
        None => (
            preset_geometry_to_path("rect", &HashMap::new(), aspect_ratio).unwrap_or_default(),
            true,
        ),
    }
}

fn graphic_label(graphic: Option<&GraphicFrameData>) -> Option<String> {
    match graphic {
        Some(GraphicFrameData::Table { .. }) => Some("Table".to_owned()),
        Some(GraphicFrameData::Chart { .. }) => Some("Chart".to_owned()),
        Some(GraphicFrameData::Diagram { .. }) => Some("Diagram".to_owned()),
        Some(GraphicFrameData::Unknown { .. }) | None => None,
    }
}

fn parse_align(value: Option<&str>) -> TextAlign {
    match value {
        Some("ctr") => TextAlign::Center,
        Some("r") => TextAlign::Right,
        Some("just") | Some("justLow") | Some("dist") | Some("thaiDist") => TextAlign::Justify,
        _ => TextAlign::Left,
    }
}

fn is_full_justification(value: Option<&str>) -> bool {
    value == Some("just")
}

fn rect_covering_text(rect: PxRect, text: Option<&TextHit>) -> PxRect {
    let Some(text) = text.filter(|text| text.overflow) else {
        return rect;
    };
    let rect = text.rect;
    let (mut left, mut right) = (rect.x, rect.x + rect.w);
    let (mut top, mut bottom) = (rect.y, rect.y + rect.h);
    for line in &text.lines {
        left = left.min(line.x);
        right = right.max(line.x + line.width);
        top = top.min(line.y);
        bottom = bottom.max(line.y + line.height);
        for run in &line.runs {
            left = left.min(run.x);
            right = right.max(run.x + run.width);
        }
    }
    PxRect {
        x: left,
        y: top,
        w: (right - left).max(rect.w),
        h: (bottom - top).max(rect.h),
    }
}

/// A `buFont` symbol face reaches its glyphs by font position, so `buChar`
/// names a slot rather than the character to draw. Checked against
/// PowerPoint's own render of the corpus: Wingdings `§` draws a small filled
/// square, Wingdings 3's `U+F075` a solid right-pointing triangle.
fn symbol_bullet(marker: &str, font: Option<&BulletFont>, theme: &Theme) -> String {
    let Some(BulletFont::Typeface(typeface)) = font else {
        return marker.to_owned();
    };
    let typeface = if typeface.starts_with('+') {
        resolve_theme_font_ref(Some(theme), typeface)
    } else {
        typeface.clone()
    };
    let Some(font) = ooxml_text::SymbolFont::named(&typeface) else {
        return marker.to_owned();
    };
    marker
        .chars()
        .map(|character| font.substitute(character).unwrap_or(character))
        .collect()
}

/// Per-level `a:buAutoNum` state: the number last drawn and the `startAt`
/// the run was seeded from.
#[derive(Default)]
struct AutoNumbering {
    numbers: [u32; 9],
    starts: [u32; 9],
}

/// Resolves a marker once per paragraph.
fn resolve_marker(
    bullet: Option<&Bullet>,
    level: u32,
    numbering: &mut AutoNumbering,
) -> Option<String> {
    let level = (level as usize).min(numbering.numbers.len() - 1);
    numbering.numbers[level + 1..].fill(0);
    numbering.starts[level + 1..].fill(0);
    match bullet {
        Some(Bullet::AutoNumber {
            scheme,
            start_at,
            restart,
        }) => {
            let start = (*start_at).clamp(1, 32_767);
            // PowerPoint writes the list's `startAt` on every one of its
            // paragraphs, so repeating the seed continues the run; only a
            // different declared start opens a new list.
            numbering.numbers[level] = match numbering.numbers[level] {
                0 => start,
                _ if *restart && numbering.starts[level] != start => start,
                current => current.saturating_add(1),
            };
            if numbering.numbers[level] == start {
                numbering.starts[level] = start;
            }
            Some(format_autonum(numbering.numbers[level], scheme))
        }
        _ => {
            numbering.numbers[level] = 0;
            numbering.starts[level] = 0;
            match bullet {
                Some(Bullet::Character { value }) if !value.trim().is_empty() => {
                    Some(value.clone())
                }
                _ => None,
            }
        }
    }
}

/// Formats Latin, Roman and decimal markers.
fn format_autonum(value: u32, scheme: &str) -> String {
    let value = value.clamp(1, MAX_AUTONUM_VALUE);
    let (numeral, suffix) = ["ParenBoth", "ParenR", "Period", "Plain"]
        .into_iter()
        .find_map(|suffix| Some((scheme.strip_suffix(suffix)?, suffix)))
        .unwrap_or((scheme, "Period"));
    let body = match numeral {
        "alphaLc" => format_alpha(value, false),
        "alphaUc" => format_alpha(value, true),
        "romanLc" => format_roman(value, false),
        "romanUc" => format_roman(value, true),
        "arabic" => value.to_string(),
        _ => return format!("{value}."),
    };
    match suffix {
        "ParenBoth" => format!("({body})"),
        "ParenR" => format!("{body})"),
        "Plain" => body,
        _ => format!("{body}."),
    }
}

fn format_alpha(value: u32, upper: bool) -> String {
    let base = if upper { b'A' } else { b'a' };
    let mut value = value.max(1);
    let mut out = Vec::new();
    while value > 0 {
        let index = (value - 1) % 26;
        out.push(base + index as u8);
        value = (value - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap_or_default()
}

fn format_roman(value: u32, upper: bool) -> String {
    const NUMERALS: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut value = value.max(1);
    let mut out = String::new();
    for (amount, numeral) in NUMERALS {
        while value >= amount {
            out.push_str(numeral);
            value -= amount;
        }
    }
    if upper { out.to_uppercase() } else { out }
}

fn normalize_family(value: &str) -> String {
    value.trim().to_lowercase()
}

fn valid_color(value: &str) -> bool {
    let value = value.strip_prefix('#').unwrap_or(value);
    value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit())
}

fn points_to_px(value: f32) -> f32 {
    value * 96.0 / 72.0
}

fn emu_to_px(value: i64) -> f32 {
    safe_geometry(value as f32 / EMU_PER_CSS_PIXEL)
}

/// A slide's page box is a whole number of points, the unit PowerPoint
/// exports and prints it in, so the extent snaps there before the px scale.
/// a slide is white paper, so a background the deck made fully transparent is
/// not a hole onto whatever is behind it.
fn is_invisible(paint: &Paint) -> bool {
    fn clear(color: &str) -> bool {
        color.len() == 9 && color.as_bytes()[7..] == *b"00"
    }
    match paint {
        Paint::Solid { color } => clear(color),
        Paint::Gradient { stops, .. } => !stops.is_empty() && stops.iter().all(|s| clear(&s.color)),
    }
}

fn slide_extent_px(value: i64) -> f32 {
    safe_geometry(((value as f64 / EMU_PER_POINT).round() * CSS_PIXELS_PER_POINT) as f32)
}

fn safe_geometry(value: f32) -> f32 {
    if value.is_finite() {
        value.clamp(-1.0e12, 1.0e12)
    } else {
        0.0
    }
}

fn utf16_len(value: &str) -> u32 {
    value.encode_utf16().count() as u32
}

#[cfg(test)]
mod tests {
    use ooxml_drawingml::ThemeColorScheme;
    use std::collections::BTreeSet;

    use pptx_edit::{DeckSession, EditCtx, TextStylePatch};
    use pptx_parse::ShapeBase;

    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    const CHART_FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/chart-deck.pptx");
    const GRADIENT_OUTLINE_FIXTURE: &[u8] =
        include_bytes!("../../pptx-parse/tests/fixtures/gradient-outline.pptx");
    const NUMBERED_FIXTURE: &[u8] =
        include_bytes!("../../pptx-parse/tests/fixtures/slide-number-fields.pptx");
    const STYLE_FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/shape-style.pptx");
    const HIDDEN_FIXTURE: &[u8] =
        include_bytes!("../../pptx-edit/tests/fixtures/hidden-shapes.pptx");
    const V2_UPDATE: &[u8] =
        include_bytes!("../../pptx-edit/tests/fixtures/deck-schema-v2-hidden.update.bin");
    const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
    const BOLD_FONT: &[u8] =
        include_bytes!("../../../packages/fonts/assets/LiberationSans-Bold.ttf");
    const ITALIC_FONT: &[u8] =
        include_bytes!("../../../packages/fonts/assets/LiberationSans-Italic.ttf");
    const BOLD_ITALIC_FONT: &[u8] =
        include_bytes!("../../../packages/fonts/assets/LiberationSans-BoldItalic.ttf");

    #[test]
    fn a_slide_extent_matches_the_page_powerpoint_exports() {
        let page_px = |emu| (f64::from(slide_extent_px(emu)) * 150.0 / 96.0).ceil() as u32;
        for (emu, px, dots) in [
            (12_192_000, 1280.0, 2000),
            (6_858_000, 720.0, 1125),
            (10_691_813, 1122.6666, 1755),
            (7_559_675, 793.3333, 1240),
            (7_556_500, 793.3333, 1240),
            (10_693_400, 1122.6666, 1755),
        ] {
            assert!((slide_extent_px(emu) - px).abs() < 0.001, "{emu}");
            assert_eq!(page_px(emu), dots, "{emu}");
        }
    }

    #[test]
    fn a_source_crop_converts_to_fractions_and_refuses_what_cannot_be_drawn() {
        assert_eq!(
            image_crop(&PictureCrop {
                left: 0,
                top: 251,
                right: 0,
                bottom: 16_720,
            }),
            ImageCrop {
                left: 0.0,
                top: 0.002_51,
                right: 0.0,
                bottom: 0.167_2,
            }
        );
        assert_eq!(
            image_crop(&PictureCrop {
                left: -5_000,
                right: 20_000,
                ..PictureCrop::default()
            }),
            ImageCrop {
                right: 0.2,
                ..ImageCrop::default()
            }
        );
        assert_eq!(
            image_crop(&PictureCrop {
                left: 60_000,
                right: 60_000,
                ..PictureCrop::default()
            }),
            ImageCrop::default()
        );
    }

    #[test]
    fn only_a_picture_with_its_own_geometry_gets_a_mask() {
        let mut picture = Picture {
            base: ShapeBase {
                id: 1,
                name: "Photo".to_owned(),
                description: None,
                hidden: false,
                placeholder: None,
                transform: ShapeTransform::default(),
            },
            relationship_id: None,
            media_part_path: None,
            effects: Vec::new(),
            crop: PictureCrop::default(),
            geometry: "rect".to_owned(),
            adjust_values: BTreeMap::new(),
            fill: None,
            outline: None,
            shape_effects: None,
            style: None,
        };
        let rect = PxRect {
            x: 0.0,
            y: 0.0,
            w: 100.0,
            h: 50.0,
        };
        assert!(picture_mask(&picture, rect).path.is_none());
        picture.geometry = "ellipse".to_owned();
        let path = picture_mask(&picture, rect).path.unwrap();
        assert_eq!(path.len(), 6);
        assert_eq!(
            path[0],
            ooxml_drawingml::GeometryPathCommand::Move { x: 1.0, y: 0.5 }
        );
        assert!(
            path[1..5].iter().all(|command| matches!(
                command,
                ooxml_drawingml::GeometryPathCommand::Cubic { .. }
            ))
        );
        assert_eq!(path[5], ooxml_drawingml::GeometryPathCommand::Close);
    }

    #[test]
    fn a_fixture_picture_keeps_its_crop_mask_and_outline_through_layout() {
        let session = DeckSession::open(
            include_bytes!("../tests/fixtures/picture-crop-mask.pptx"),
            288,
        )
        .unwrap();
        let rendered = renderer()
            .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
            .unwrap();
        let image = rendered
            .display_list
            .primitives
            .iter()
            .find(|primitive| matches!(primitive, Primitive::Image { object_id: 90, .. }))
            .unwrap();
        let Primitive::Image {
            x,
            y,
            w,
            h,
            crop,
            path,
            stroke,
            ..
        } = image
        else {
            unreachable!()
        };
        assert_eq!((*x, *y, *w, *h), (100.0, 50.0, 200.0, 100.0));
        assert_eq!(
            *crop,
            ImageCrop {
                left: 0.1,
                top: 0.2,
                right: 0.3,
                bottom: 0.1
            }
        );
        let path = path.as_ref().unwrap();
        assert_eq!(path.len(), 6);
        assert_eq!(
            path[0],
            ooxml_drawingml::GeometryPathCommand::Move { x: 1.0, y: 0.5 }
        );
        assert!(
            path[1..5].iter().all(|command| matches!(
                command,
                ooxml_drawingml::GeometryPathCommand::Cubic { .. }
            ))
        );
        let stroke = stroke.as_ref().unwrap();
        assert_eq!(stroke.width, 2.0);
        assert_eq!(stroke.color, "#FF00FF");
    }

    #[test]
    fn overflowing_text_grows_the_hit_region_so_a_visible_glyph_stays_clickable() {
        let rect = PxRect {
            x: 10.0,
            y: 100.0,
            w: 200.0,
            h: 20.0,
        };
        let line = |y: f32, height: f32| PositionedTextLine {
            x: 10.0,
            y,
            width: 200.0,
            height,
            baseline: y + height,
            start: 0,
            end: 0,
            runs: Vec::new(),
            caret_stops: Vec::new(),
        };
        let text = TextHit {
            rect,
            transform: Transform::default(),
            overflow: true,
            story_id: "story".to_owned(),
            lines: vec![line(80.0, 40.0), line(120.0, 40.0)],
        };
        let grown = rect_covering_text(rect, Some(&text));
        assert_eq!(grown.y, 80.0);
        assert_eq!(grown.y + grown.h, 160.0);
        assert_eq!(grown.x, rect.x);
        assert_eq!(grown.w, rect.w);
        let fitting = TextHit {
            rect,
            transform: Transform::default(),
            overflow: false,
            story_id: "story".to_owned(),
            lines: vec![line(104.0, 12.0)],
        };
        let same = rect_covering_text(rect, Some(&fitting));
        assert_eq!((same.y, same.h), (rect.y, rect.h));
        assert_eq!(
            (
                rect_covering_text(rect, None).y,
                rect_covering_text(rect, None).h
            ),
            (rect.y, rect.h)
        );
    }

    #[test]
    fn autonumbering_counts_per_level_and_resumes_across_other_levels() {
        let mut counters = AutoNumbering::default();
        let number = |level, counters: &mut AutoNumbering| {
            resolve_marker(
                Some(&Bullet::AutoNumber {
                    scheme: "arabicPeriod".to_owned(),
                    start_at: 1,
                    restart: false,
                }),
                level,
                counters,
            )
        };
        let dash = |level, counters: &mut AutoNumbering| {
            resolve_marker(
                Some(&Bullet::Character {
                    value: "-".to_owned(),
                }),
                level,
                counters,
            )
        };
        assert_eq!(number(0, &mut counters).as_deref(), Some("1."));
        assert_eq!(dash(1, &mut counters).as_deref(), Some("-"));
        assert_eq!(dash(1, &mut counters).as_deref(), Some("-"));
        assert_eq!(number(0, &mut counters).as_deref(), Some("2."));
        assert_eq!(number(0, &mut counters).as_deref(), Some("3."));
        assert_eq!(number(1, &mut counters).as_deref(), Some("1."));
        assert_eq!(number(1, &mut counters).as_deref(), Some("2."));
        assert_eq!(number(0, &mut counters).as_deref(), Some("4."));
        assert_eq!(number(1, &mut counters).as_deref(), Some("1."));
        assert_eq!(resolve_marker(Some(&Bullet::None), 0, &mut counters), None);
    }

    #[test]
    fn inherited_autonumber_start_at_applies_only_to_the_first_item() {
        let mut counters = AutoNumbering::default();
        let number = |counters: &mut AutoNumbering| {
            resolve_marker(
                Some(&Bullet::AutoNumber {
                    scheme: "arabicPeriod".to_owned(),
                    start_at: 7,
                    restart: false,
                }),
                0,
                counters,
            )
        };
        assert_eq!(number(&mut counters).as_deref(), Some("7."));
        assert_eq!(number(&mut counters).as_deref(), Some("8."));
        let mut counters = AutoNumbering::default();
        let last_start = Bullet::AutoNumber {
            scheme: "arabicPeriod".to_owned(),
            start_at: 32_767,
            restart: false,
        };
        assert_eq!(
            resolve_marker(Some(&last_start), 0, &mut counters).as_deref(),
            Some("32767.")
        );
        assert_eq!(
            resolve_marker(Some(&last_start), 0, &mut counters).as_deref(),
            Some("32768.")
        );
    }

    #[test]
    fn autonumber_schemes_format_their_numeral_and_suffix() {
        assert_eq!(format_autonum(4, "arabicPeriod"), "4.");
        assert_eq!(format_autonum(4, "arabicParenR"), "4)");
        assert_eq!(format_autonum(4, "arabicParenBoth"), "(4)");
        assert_eq!(format_autonum(4, "arabicPlain"), "4");
        assert_eq!(format_autonum(1, "alphaLcParenR"), "a)");
        assert_eq!(format_autonum(27, "alphaUcPeriod"), "AA.");
        assert_eq!(format_autonum(9, "romanLcPeriod"), "ix.");
        assert_eq!(format_autonum(2024, "romanUcPeriod"), "MMXXIV.");
        assert_eq!(format_autonum(3, "somethingElse"), "3.");
        assert_eq!(format_autonum(3, "somethingElsePlain"), "3.");
        assert_eq!(format_autonum(3, "somethingElseParenBoth"), "3.");
    }

    #[test]
    fn autonumber_sequences_restart_after_plain_paragraphs_and_explicit_starts() {
        let mut counters = AutoNumbering::default();
        let number = Bullet::AutoNumber {
            scheme: "arabicPeriod".to_owned(),
            start_at: 1,
            restart: false,
        };
        let restart = Bullet::AutoNumber {
            scheme: "arabicPeriod".to_owned(),
            start_at: 7,
            restart: true,
        };
        assert_eq!(
            resolve_marker(Some(&number), 0, &mut counters).as_deref(),
            Some("1.")
        );
        assert_eq!(
            resolve_marker(Some(&restart), 0, &mut counters).as_deref(),
            Some("7.")
        );
        assert_eq!(
            resolve_marker(Some(&number), 0, &mut counters).as_deref(),
            Some("8.")
        );
        assert_eq!(resolve_marker(None, 0, &mut counters), None);
        assert_eq!(
            resolve_marker(Some(&number), 0, &mut counters).as_deref(),
            Some("1.")
        );
        assert_eq!(
            resolve_marker(Some(&number), 1, &mut counters).as_deref(),
            Some("1.")
        );
        assert_eq!(
            resolve_marker(
                Some(&Bullet::Character {
                    value: "•".to_owned()
                }),
                0,
                &mut counters
            )
            .as_deref(),
            Some("•")
        );
        assert_eq!(
            resolve_marker(Some(&number), 1, &mut counters).as_deref(),
            Some("1.")
        );
        assert_eq!(
            resolve_marker(Some(&number), 0, &mut counters).as_deref(),
            Some("1.")
        );
    }

    /// `pptarena-018-original` slide 11 declares `startAt="4"` on all four of
    /// its paragraphs and PowerPoint renders 4, 5, 6, 7;
    /// `pptarena-034-original` slide 11 declares `startAt="1"` on all five and
    /// PowerPoint renders a) through e).
    #[test]
    fn a_repeated_declared_start_continues_the_list() {
        let mut counters = AutoNumbering::default();
        let arabic = Bullet::AutoNumber {
            scheme: "arabicPeriod".to_owned(),
            start_at: 4,
            restart: true,
        };
        let alpha = Bullet::AutoNumber {
            scheme: "alphaLcParenR".to_owned(),
            start_at: 1,
            restart: true,
        };
        assert_eq!(
            (0..4)
                .filter_map(|_| resolve_marker(Some(&arabic), 0, &mut counters))
                .collect::<Vec<_>>(),
            ["4.", "5.", "6.", "7."]
        );
        let mut counters = AutoNumbering::default();
        assert_eq!(
            (0..5)
                .filter_map(|_| resolve_marker(Some(&alpha), 0, &mut counters))
                .collect::<Vec<_>>(),
            ["a)", "b)", "c)", "d)", "e)"]
        );
    }

    #[test]
    fn autonumber_roman_markers_bound_untrusted_start_values() {
        assert_eq!(format_autonum(0, "arabicPeriod"), "1.");
        assert_eq!(format_autonum(u32::MAX, "romanUcPeriod").len(), 61);
        assert_eq!(
            format_autonum(u32::MAX, "romanUcPeriod"),
            format!("{}DCCLXVII.", "M".repeat(52))
        );
    }

    fn renderer() -> SlideRenderer {
        let mut renderer = SlideRenderer::new();
        for bold in [false, true] {
            renderer.register_font("Arial", bold, false, FONT).unwrap();
        }
        renderer
    }

    fn tabbed(
        renderer: &SlideRenderer,
        text: &str,
        stops: Vec<f32>,
        default_px: f32,
    ) -> ResolvedParagraph {
        let mut paragraph = paragraph(renderer, "l", text);
        paragraph.tab_stops = stops;
        paragraph.default_tab_px = default_px;
        paragraph
    }

    fn glyph_positions(lines: &[PositionedTextLine]) -> Vec<f32> {
        lines
            .iter()
            .flat_map(|line| line.runs.iter())
            .flat_map(|run| run.glyphs.iter())
            .map(|glyph| glyph.x)
            .collect()
    }

    #[test]
    fn a_tab_advances_to_the_next_default_stop_without_painting_a_glyph() {
        let renderer = renderer();
        let paragraph = tabbed(&renderer, "A\tB", Vec::new(), 96.0);
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 0.0, 0.0, 1_000.0, 1.0, false).unwrap();
        let positions = glyph_positions(&lines);
        assert_eq!(positions.len(), 2, "the tab paints nothing");
        assert!(positions[0].abs() < 0.01, "{positions:?}");
        assert!((positions[1] - 96.0).abs() < 0.01, "{positions:?}");
    }

    #[test]
    fn a_declared_stop_wins_over_the_default_pitch() {
        let renderer = renderer();
        let paragraph = tabbed(&renderer, "A\tB", vec![40.0, 300.0], 96.0);
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 0.0, 0.0, 1_000.0, 1.0, false).unwrap();
        let positions = glyph_positions(&lines);
        assert!((positions[1] - 40.0).abs() < 0.01, "{positions:?}");
    }

    #[test]
    fn the_left_margin_offsets_the_stop_a_tab_reaches() {
        let renderer = renderer();
        let mut paragraph = tabbed(&renderer, "A\tB", vec![40.0, 200.0], 96.0);
        paragraph.margin_left_px = 50.0;
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 50.0, 0.0, 950.0, 1.0, false).unwrap();
        let positions = glyph_positions(&lines);
        assert!((positions[1] - 200.0).abs() < 0.01, "{positions:?}");
    }

    #[test]
    fn a_tab_never_reaches_past_the_line() {
        let renderer = renderer();
        let paragraph = tabbed(&renderer, "A\tB", Vec::new(), 96.0);
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 0.0, 0.0, 40.0, 1.0, false).unwrap();
        for line in &lines {
            assert!(line.width <= 40.01, "{}", line.width);
        }
    }

    #[test]
    fn a_leading_tab_in_a_hanging_indent_lands_on_the_margin_the_line_already_starts_at() {
        let renderer = renderer();
        let mut paragraph = tabbed(&renderer, "\tItem", Vec::new(), 96.0);
        paragraph.margin_left_px = 30.0;
        paragraph.indent_px = -30.0;
        paragraph.tab_stops = resolve_tab_stops(None, 285_750, -285_750);
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 30.0, 0.0, 300.0, 1.0, false).unwrap();
        let positions = glyph_positions(&lines);
        assert!((positions[0] - 30.0).abs() < 0.01, "{positions:?}");
    }

    #[test]
    fn only_the_first_tab_of_a_hanging_paragraph_measures_from_the_hanging_position() {
        let renderer = renderer();
        let mut paragraph = tabbed(&renderer, "\tItem\tNote", Vec::new(), 96.0);
        paragraph.margin_left_px = 30.0;
        paragraph.indent_px = -30.0;
        paragraph.tab_stops = resolve_tab_stops(None, 285_750, -285_750);
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 30.0, 0.0, 400.0, 1.0, false).unwrap();
        let positions = glyph_positions(&lines);
        assert!((positions[0] - 30.0).abs() < 0.01, "{positions:?}");
        assert!((positions[4] - 96.0).abs() < 0.01, "{positions:?}");
    }

    #[test]
    fn a_marker_leaves_no_hanging_space_for_a_tab_to_take() {
        let renderer = renderer();
        let mut paragraph = tabbed(&renderer, "\tItem", Vec::new(), 96.0);
        paragraph.margin_left_px = 30.0;
        paragraph.indent_px = -30.0;
        paragraph.marker = Some("\u{2022}".to_owned());
        paragraph.tab_stops = resolve_tab_stops(None, 285_750, -285_750);
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 30.0, 0.0, 300.0, 1.0, false).unwrap();
        let text = lines[0]
            .runs
            .iter()
            .find(|run| run.text.contains("Item"))
            .unwrap();
        assert!(
            (text.glyphs[0].x - 96.0).abs() < 0.01,
            "{}",
            text.glyphs[0].x
        );
    }

    #[test]
    fn tab_advance_bounds_a_hostile_pitch_and_offset() {
        assert_eq!(tab_advance(0.0, 0.0, &[], 96.0, 1_000.0), 96.0);
        assert_eq!(tab_advance(96.0, 96.0, &[], 96.0, 1_000.0), 96.0);
        assert_eq!(tab_advance(1.0, 1.0, &[40.0], 96.0, 1_000.0), 39.0);
        assert_eq!(tab_advance(10.0, 40.0, &[30.0], 96.0, 1_000.0), 0.0);
        assert_eq!(tab_advance(f32::NAN, 0.0, &[], 96.0, 1_000.0), 0.0);
        assert_eq!(tab_advance(0.0, f32::NAN, &[], 96.0, 1_000.0), 0.0);
        assert_eq!(tab_advance(0.0, 0.0, &[], 96.0, f32::NAN), 0.0);
        assert!((0.0..=10.0).contains(&tab_advance(f32::MAX, f32::MAX, &[], 96.0, 10.0)));
        assert_eq!(resolve_default_tab(Some(0)), 96.0);
        assert_eq!(resolve_default_tab(Some(-914_400)), 96.0);
    }

    fn justify(widths: &[(f32, bool, bool)]) -> Vec<JustifyCluster> {
        widths
            .iter()
            .map(|(width, break_after, blank)| JustifyCluster {
                width: *width,
                break_after: *break_after,
                blank: *blank,
            })
            .collect()
    }

    fn paragraph(renderer: &SlideRenderer, alignment: &str, text: &str) -> ResolvedParagraph {
        let face = renderer.resolve_face("Arial", false, false).unwrap();
        ResolvedParagraph {
            align: parse_align(Some(alignment)),
            justify: is_full_justification(Some(alignment)),
            level: 0,
            margin_left_px: 0.0,
            margin_right_px: 0.0,
            line_spacing: None,
            space_before: None,
            space_after: None,
            line_space_reduction: 0.0,
            indent_px: 0.0,
            tab_stops: Vec::new(),
            default_tab_px: resolve_default_tab(None),
            marker: None,
            bullet_style: None,
            runs: vec![ResolvedRun {
                text: text.to_owned(),
                start: 0,
                style: ResolvedStyle {
                    face,
                    family: "Arial".to_owned(),
                    font_size_pt: 18.0,
                    line_font_size_pt: 18.0,
                    spacing_pt: 0.0,
                    baseline_shift_px: 0.0,
                    bold: false,
                    italic: false,
                    underline: false,
                    color: "#000000".to_owned(),
                    caps: TextCaps::None,
                },
            }],
        }
    }

    #[test]
    fn justification_spreads_the_slack_across_the_gaps() {
        // "aa bb cc" laid out in 100px: 80px of clusters, two gaps to share 20px
        let line = justify(&[
            (20.0, false, false),
            (10.0, true, true),
            (20.0, false, false),
            (10.0, true, true),
            (20.0, false, false),
        ]);
        let padding = justification_padding(&line, 100.0);
        assert_eq!(padding, vec![0.0, 10.0, 0.0, 10.0, 0.0]);
        let width: f32 = line
            .iter()
            .zip(&padding)
            .map(|(cluster, pad)| cluster.width + pad)
            .sum();
        assert_eq!(width, 100.0);
    }

    #[test]
    fn justification_ignores_a_trailing_blank() {
        // the wrap keeps the space that ended the line; stretching to it would
        // leave the last glyph short of the edge
        let line = justify(&[
            (20.0, false, false),
            (10.0, true, true),
            (20.0, false, false),
            (10.0, true, true),
        ]);
        let padding = justification_padding(&line, 60.0);
        assert_eq!(padding, vec![0.0, 10.0, 0.0, 0.0]);
        let width: f32 = line[..3]
            .iter()
            .zip(&padding)
            .map(|(cluster, pad)| cluster.width + pad)
            .sum();
        assert_eq!(width, 60.0);
    }

    #[test]
    fn justification_leaves_a_line_it_cannot_stretch() {
        // one unbroken word has no gap to widen, and an overfull line no slack
        assert_eq!(
            justification_padding(&justify(&[(80.0, false, false)]), 100.0),
            vec![0.0]
        );
        assert_eq!(
            justification_padding(
                &justify(&[
                    (80.0, false, false),
                    (10.0, true, true),
                    (80.0, false, false)
                ]),
                100.0
            ),
            vec![0.0, 0.0, 0.0]
        );
    }

    #[test]
    fn justified_layout_stretches_non_last_lines_only() {
        let renderer = renderer();
        let text = "alpha beta    gamma delta";
        let justified = paragraph(&renderer, "just", text);
        let clusters = shape_paragraph(&renderer.fonts, &justified, 1.0).unwrap();
        let second_break = clusters
            .iter()
            .enumerate()
            .find(|(_, cluster)| cluster.end == utf16_len("alpha beta    "))
            .map(|(index, _)| index + 1)
            .unwrap();
        let prefix_width = clusters[..second_break]
            .iter()
            .map(|cluster| cluster.width)
            .sum::<f32>();
        let trailing_count = clusters[..second_break]
            .iter()
            .rev()
            .take_while(|cluster| cluster_is_blank(cluster))
            .count();
        let width = prefix_width + clusters[second_break].width / 2.0;
        let lines =
            layout_paragraph(&renderer.fonts, &justified, 20.0, 30.0, width, 1.0, false).unwrap();
        let natural = layout_paragraph(
            &renderer.fonts,
            &paragraph(&renderer, "justLow", text),
            20.0,
            30.0,
            width,
            1.0,
            false,
        )
        .unwrap();

        assert!(lines.len() > 1);
        assert_eq!(lines.len(), natural.len());
        let beta = utf16_len("alpha ");
        let stretched_beta = lines[0]
            .caret_stops
            .iter()
            .find(|stop| stop.position == beta)
            .unwrap();
        let natural_beta = natural[0]
            .caret_stops
            .iter()
            .find(|stop| stop.position == beta)
            .unwrap();
        assert!(stretched_beta.x > natural_beta.x);
        assert!((lines[0].width - width).abs() < 0.001);
        let trailing = &lines[0].caret_stops;
        let edge = trailing.last().unwrap().x;
        assert!(trailing_count >= 2);
        assert!(
            trailing
                .iter()
                .rev()
                .take(trailing_count + 1)
                .all(|stop| stop.x == edge)
        );
        assert_eq!(lines.last(), natural.last());
    }

    #[test]
    fn only_full_justification_enables_stretching() {
        assert!(is_full_justification(Some("just")));
        for alignment in ["justLow", "dist", "thaiDist"] {
            assert_eq!(parse_align(Some(alignment)), TextAlign::Justify);
            assert!(!is_full_justification(Some(alignment)));
        }
    }

    #[test]
    fn duotone_colours_resolve_against_the_theme() {
        let theme = Theme {
            color_scheme: ThemeColorScheme {
                lt2: "FFFFFF".to_owned(),
                ..ThemeColorScheme::default()
            },
            ..Theme::default()
        };
        let effects = image_effects(
            &[
                BlipEffect::Duotone {
                    shadow: Some(ColorValue {
                        theme_color: Some("background2".to_owned()),
                        theme_shade: Some("73".to_owned()),
                        ..ColorValue::default()
                    }),
                    highlight: Some(ColorValue {
                        rgb: Some("FFFFFF".to_owned()),
                        ..ColorValue::default()
                    }),
                },
                BlipEffect::BiLevel { threshold: 0.25 },
            ],
            &theme,
        );

        assert_eq!(
            effects,
            vec![
                ImageEffect::Duotone {
                    shadow: "#737373FF".to_owned(),
                    highlight: "#FFFFFFFF".to_owned(),
                },
                ImageEffect::BiLevel { threshold: 0.25 },
            ]
        );
    }

    #[test]
    fn an_unresolvable_duotone_is_dropped() {
        let effects = image_effects(
            &[BlipEffect::Duotone {
                shadow: Some(ColorValue::default()),
                highlight: Some(ColorValue {
                    rgb: Some("FFFFFF".to_owned()),
                    ..ColorValue::default()
                }),
            }],
            &Theme::default(),
        );

        assert!(effects.is_empty());
    }

    #[test]
    fn luminance_outside_the_legal_range_clamps() {
        let effects = image_effects(
            &[BlipEffect::Luminance {
                brightness: 4.0,
                contrast: -3.0,
            }],
            &Theme::default(),
        );

        assert_eq!(
            effects,
            vec![ImageEffect::Luminance {
                brightness: 1.0,
                contrast: -1.0,
            }]
        );
    }

    #[test]
    fn adjacent_runs_keep_their_own_paint_attributes() {
        let renderer = renderer();
        let style = ResolvedStyle {
            face: renderer.resolve_face("Arial", false, false).unwrap(),
            family: "Arial".to_owned(),
            font_size_pt: 14.0,
            line_font_size_pt: 14.0,
            spacing_pt: 0.0,
            baseline_shift_px: 0.0,
            bold: false,
            italic: false,
            underline: false,
            color: "#000000".to_owned(),
            caps: TextCaps::None,
        };
        let mut variants = vec![style.clone(); 7];
        variants[0].color = "#A99A72".to_owned();
        variants[1].bold = true;
        variants[2].italic = true;
        variants[3].underline = true;
        variants[4].font_size_pt = 28.0;
        variants[5].family = "Fallback".to_owned();
        variants[6].face = renderer.resolve_face("Arial", true, false).unwrap();
        for changed in variants {
            let paragraph = ResolvedParagraph {
                align: TextAlign::Left,
                justify: false,
                level: 0,
                margin_left_px: 0.0,
                margin_right_px: 0.0,
                line_spacing: None,
                space_before: None,
                space_after: None,
                line_space_reduction: 0.0,
                indent_px: 0.0,
                tab_stops: Vec::new(),
                default_tab_px: resolve_default_tab(None),
                marker: None,
                bullet_style: None,
                runs: [style.clone(), changed.clone(), style.clone()]
                    .into_iter()
                    .enumerate()
                    .map(|(index, style)| ResolvedRun {
                        text: "word ".to_owned(),
                        start: index as u32 * 5,
                        style,
                    })
                    .collect(),
            };
            for scale in [1.0, 0.5] {
                let lines = layout_paragraph(
                    &renderer.fonts,
                    &paragraph,
                    10.0,
                    20.0,
                    10_000.0,
                    scale,
                    false,
                )
                .unwrap();
                assert_eq!(lines.len(), 1);
                let runs = &lines[0].runs;
                assert_eq!(runs.len(), 3);
                for (index, (actual, expected)) in runs.iter().zip(&paragraph.runs).enumerate() {
                    assert_eq!(actual.text, expected.text);
                    assert_eq!(
                        (actual.start, actual.end),
                        (index as u32 * 5, index as u32 * 5 + 5)
                    );
                    assert_eq!(actual.color, expected.style.color);
                    assert_eq!(actual.bold, expected.style.bold);
                    assert_eq!(actual.italic, expected.style.italic);
                    assert_eq!(actual.underline, expected.style.underline);
                    assert_eq!(actual.font_id, expected.style.face.id.to_u32());
                    assert_eq!(actual.font_family, expected.style.family);
                    assert_eq!(
                        actual.font_size_px,
                        points_to_px(autofit_size_pt(expected.style.font_size_pt, scale))
                    );
                }
            }
        }
    }

    #[test]
    fn identical_adjacent_runs_keep_the_same_display_list() {
        let renderer = renderer();
        let style = ResolvedStyle {
            face: renderer.resolve_face("Arial", false, false).unwrap(),
            family: "Arial".to_owned(),
            font_size_pt: 14.0,
            line_font_size_pt: 14.0,
            spacing_pt: 0.0,
            baseline_shift_px: 0.0,
            bold: false,
            italic: false,
            underline: true,
            color: "#A99A72".to_owned(),
            caps: TextCaps::None,
        };
        let paragraph = |parts: &[&str]| {
            let mut start = 0;
            ResolvedParagraph {
                align: TextAlign::Justify,
                justify: true,
                level: 0,
                margin_left_px: 0.0,
                margin_right_px: 0.0,
                line_spacing: None,
                space_before: None,
                space_after: None,
                line_space_reduction: 0.0,
                indent_px: 0.0,
                tab_stops: Vec::new(),
                default_tab_px: resolve_default_tab(None),
                marker: None,
                bullet_style: None,
                runs: parts
                    .iter()
                    .map(|text| {
                        let run = ResolvedRun {
                            text: (*text).to_owned(),
                            start,
                            style: style.clone(),
                        };
                        start += utf16_len(text);
                        run
                    })
                    .collect(),
            }
        };
        let split = paragraph(&["alpha ", "", "beta ", "gamma ", "delta"]);
        let joined = paragraph(&["alpha beta gamma delta"]);
        for width in [100.0, 10_000.0] {
            let render = |paragraph| {
                layout_paragraph(&renderer.fonts, paragraph, 10.0, 20.0, width, 1.0, false).unwrap()
            };
            assert_eq!(render(&split), render(&joined));
        }
    }

    #[test]
    fn a_stack_puts_one_glyph_on_each_line_and_no_line_on_a_break() {
        let mut renderer = SlideRenderer::new();
        renderer
            .register_font(
                "Arial",
                false,
                false,
                include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"),
            )
            .unwrap();
        let style = ResolvedStyle {
            face: renderer.resolve_face("Arial", false, false).unwrap(),
            family: "Arial".to_owned(),
            font_size_pt: 24.0,
            line_font_size_pt: 24.0,
            spacing_pt: 0.0,
            baseline_shift_px: 0.0,
            bold: false,
            italic: false,
            underline: false,
            color: "#000000".to_owned(),
            caps: TextCaps::None,
        };
        let stack = |text: &str| {
            let paragraph = ResolvedParagraph {
                align: TextAlign::Left,
                justify: false,
                level: 0,
                margin_left_px: 0.0,
                margin_right_px: 0.0,
                space_before: None,
                space_after: None,
                line_spacing: None,
                line_space_reduction: 0.0,
                indent_px: 0.0,
                tab_stops: Vec::new(),
                default_tab_px: resolve_default_tab(None),
                marker: None,
                bullet_style: None,
                runs: vec![ResolvedRun {
                    text: text.to_owned(),
                    start: 0,
                    style: style.clone(),
                }],
            };
            layout_paragraph(&renderer.fonts, &paragraph, 0.0, 0.0, 1000.0, 1.0, true)
                .unwrap()
                .iter()
                .map(|line| {
                    line.runs
                        .iter()
                        .map(|run| run.text.clone())
                        .collect::<String>()
                })
                .collect::<Vec<_>>()
        };

        assert_eq!(stack("A\nB"), ["A", "B"]);
        assert_eq!(stack("A B"), ["A", " ", "B"]);
        assert_eq!(stack("ffi"), ["f", "f", "i"]);
    }

    #[test]
    fn gradient_stops_reach_the_display_list_in_position_order() {
        use ooxml_drawingml::{ColorValue, GradientStop as ModelStop};

        let stop = |position, rgb: &str, alpha| ModelStop {
            position,
            color: ColorValue {
                rgb: Some(rgb.to_owned()),
                alpha,
                ..ColorValue::default()
            },
        };
        let ordered = [
            stop(0.0, "404040", None),
            stop(50_000.0, "FF0000", Some(0.5)),
            stop(50_000.0, "00FF00", None),
            stop(100_000.0, "262626", None),
        ];
        for (kind, gradient_type) in [
            ("linear", GradientType::Linear),
            ("radial", GradientType::Radial),
            ("rectangular", GradientType::Rectangular),
            ("path", GradientType::Path),
        ] {
            for indices in [[0, 1, 2, 3], [3, 1, 0, 2]] {
                let fill = ShapeFill {
                    fill_type: "gradient".to_owned(),
                    color: None,
                    gradient: Some(ooxml_drawingml::GradientFill {
                        gradient_type: kind.to_owned(),
                        angle: Some(45.0),
                        stops: indices.map(|index| ordered[index].clone()).to_vec(),
                    }),
                };
                let source = serde_json::to_vec(&fill).unwrap();
                assert_eq!(
                    paint(&fill, &Theme::default()),
                    Some(Paint::Gradient {
                        gradient_type,
                        angle_deg: Some(45.0),
                        stops: [
                            (0.0, "#404040"),
                            (0.5, "#FF000080"),
                            (0.5, "#00FF00"),
                            (1.0, "#262626"),
                        ]
                        .map(|(position, color)| GradientStop {
                            position,
                            color: color.to_owned(),
                        })
                        .to_vec(),
                    }),
                    "{kind}: {indices:?}",
                );
                assert_eq!(serde_json::to_vec(&fill).unwrap(), source);
            }
        }
        let session = DeckSession::open(
            include_bytes!("../tests/fixtures/gradient-stop-order.pptx"),
            306,
        )
        .unwrap();
        let rendered = renderer()
            .layout_slide(session.package(), &session.snapshot().unwrap(), 2)
            .unwrap();
        let Some(Paint::Gradient { stops, .. }) = rendered.display_list.background else {
            panic!("expected a gradient");
        };
        assert_eq!(
            stops
                .iter()
                .map(|stop| stop.color.as_str())
                .collect::<Vec<_>>(),
            ["#FF0000", "#FF00FF", "#00FF00", "#FFFF00", "#0000FF"]
        );
    }

    #[test]
    fn an_unregistered_family_keeps_its_weight_through_the_fallback() {
        let mut renderer = SlideRenderer::new();
        renderer.register_font("Arial", false, false, FONT).unwrap();
        let bold = renderer
            .register_font("Arial", true, false, BOLD_FONT)
            .unwrap();

        let resolved = renderer.resolve_face("Segoe UI", true, false).unwrap();
        assert_eq!(resolved.id.to_u32(), bold);
        assert_eq!(renderer.fonts.font_bytes(resolved.id).unwrap(), BOLD_FONT);
    }

    #[test]
    fn an_unregistered_family_keeps_its_slant_through_the_fallback() {
        let mut renderer = SlideRenderer::new();
        renderer.register_font("Arial", false, false, FONT).unwrap();
        let italic = renderer
            .register_font("Arial", false, true, ITALIC_FONT)
            .unwrap();

        let resolved = renderer.resolve_face("Segoe UI", false, true).unwrap();
        assert_eq!(resolved.id.to_u32(), italic);
        assert_eq!(renderer.fonts.font_bytes(resolved.id).unwrap(), ITALIC_FONT);
    }

    #[test]
    fn fallback_keeps_combined_style_and_prefers_a_registered_family() {
        let mut renderer = SlideRenderer::new();
        let regular = renderer.register_font("Arial", false, false, FONT).unwrap();
        renderer
            .register_font("ARIAL", true, false, BOLD_FONT)
            .unwrap();
        renderer
            .register_font("Arial", false, true, ITALIC_FONT)
            .unwrap();
        let bold_italic = renderer
            .register_font(" arial ", true, true, BOLD_ITALIC_FONT)
            .unwrap();
        let georgia = renderer
            .register_font("Georgia", false, false, FONT)
            .unwrap();

        let resolved = renderer.resolve_face(" geORGia ", true, true).unwrap();
        assert_eq!(resolved.id.to_u32(), georgia);
        let resolved = renderer.resolve_face("Segoe UI", true, true).unwrap();
        assert_eq!(resolved.id.to_u32(), bold_italic);
        assert_eq!(renderer.fallback_font().unwrap().to_u32(), regular);
    }

    #[test]
    fn fallback_with_only_bold_and_italic_is_independent_of_registration_order() {
        fn substitute_family(shapes: &mut [ShapeSnapshot]) {
            for shape in shapes {
                for story in &mut shape.text_stories {
                    for paragraph in &mut story.paragraphs {
                        for run in &mut paragraph.runs {
                            run.style.font_family = Some("Segoe UI".to_owned());
                        }
                    }
                }
                substitute_family(&mut shape.children);
            }
        }

        fn normalize_font_ids(primitives: &mut [Primitive], bold_id: u32) {
            for primitive in primitives {
                match primitive {
                    Primitive::TextBox { lines, .. } => {
                        for run in lines.iter_mut().flat_map(|line| &mut line.runs) {
                            run.font_id = u32::from(run.font_id != bold_id);
                        }
                    }
                    Primitive::Chart { primitives, .. } | Primitive::Table { primitives, .. } => {
                        normalize_font_ids(primitives, bold_id)
                    }
                    _ => {}
                }
            }
        }

        assert!(matches!(
            SlideRenderer::new().resolve_face("Segoe UI", false, false),
            Err(RenderError::NoFont)
        ));
        let package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_010).unwrap();
        let mut snapshot = session.snapshot().unwrap();
        for slide in &mut snapshot.slides {
            substitute_family(&mut slide.shapes);
        }
        let mut outputs = Vec::new();
        let mut renderers = Vec::new();
        for reverse in [false, true] {
            let mut renderer = SlideRenderer::new();
            let mut faces = [(true, false, BOLD_FONT), (false, true, ITALIC_FONT)];
            if reverse {
                faces.reverse();
            }
            for (bold, italic, bytes) in faces {
                renderer
                    .register_font("Arial", bold, italic, bytes)
                    .unwrap();
            }
            let bold_id = renderer.resolve_face("Arial", true, false).unwrap().id;
            let mut slides = Vec::new();
            for index in 0..snapshot.slides.len() {
                let mut rendered = renderer.layout_slide(&package, &snapshot, index).unwrap();
                normalize_font_ids(&mut rendered.display_list.primitives, bold_id.to_u32());
                slides.push(rendered.display_list);
            }
            outputs.push(serde_json::to_vec(&slides).unwrap());
            renderers.push(renderer);
        }
        assert!(outputs[0] == outputs[1], "fallback display lists differ");
        for renderer in renderers {
            for (bold, italic, expected) in [
                (false, false, ITALIC_FONT),
                (true, false, BOLD_FONT),
                (false, true, ITALIC_FONT),
                (true, true, BOLD_FONT),
            ] {
                let resolved = renderer.resolve_face("Segoe UI", bold, italic).unwrap();
                assert!(
                    renderer.fonts.font_bytes(resolved.id).unwrap() == expected,
                    "wrong fallback for bold={bold}, italic={italic}"
                );
            }
        }
    }

    #[test]
    fn a_translucent_fill_carries_its_alpha_into_the_display_list() {
        let theme = Theme::default();
        let translucent = ShapeFill {
            fill_type: "solid".to_owned(),
            color: Some(ColorValue {
                rgb: Some("112233".to_owned()),
                alpha: Some(0.5),
                ..ColorValue::default()
            }),
            gradient: None,
        };
        assert_eq!(
            paint(&translucent, &theme),
            Some(Paint::Solid {
                color: "#11223380".to_owned()
            })
        );

        let opaque = ShapeFill {
            color: Some(ColorValue {
                rgb: Some("112233".to_owned()),
                ..ColorValue::default()
            }),
            ..translucent
        };
        assert_eq!(
            paint(&opaque, &theme),
            Some(Paint::Solid {
                color: "#112233".to_owned()
            })
        );
    }

    #[test]
    fn a_translucent_gradient_stop_and_outline_carry_their_alpha() {
        let theme = Theme::default();
        let color = |rgb: &str, alpha: Option<f64>| ColorValue {
            rgb: Some(rgb.to_owned()),
            alpha,
            ..ColorValue::default()
        };

        let gradient = ShapeFill {
            fill_type: "gradient".to_owned(),
            color: None,
            gradient: Some(ooxml_drawingml::GradientFill {
                gradient_type: "linear".to_owned(),
                angle: None,
                stops: vec![
                    ooxml_drawingml::GradientStop {
                        position: 0.0,
                        color: color("112233", Some(0.5)),
                    },
                    ooxml_drawingml::GradientStop {
                        position: 100_000.0,
                        color: color("445566", None),
                    },
                ],
            }),
        };
        let Some(Paint::Gradient { stops, .. }) = paint(&gradient, &theme) else {
            panic!("expected a gradient paint");
        };
        assert_eq!(
            stops
                .iter()
                .map(|stop| stop.color.as_str())
                .collect::<Vec<_>>(),
            ["#11223380", "#445566"]
        );

        let outline = |alpha| ShapeOutline {
            color: Some(color("112233", alpha)),
            ..ShapeOutline::default()
        };
        assert_eq!(
            stroke(&outline(Some(0.5)), &theme).map(|stroke| stroke.color),
            Some("#11223380".to_owned())
        );
        assert_eq!(
            stroke(&outline(None), &theme).map(|stroke| stroke.color),
            Some("#112233".to_owned())
        );
    }

    #[test]
    fn an_outer_shadow_resolves_to_translucent_pixels_offset_along_its_direction() {
        let theme = Theme::default();
        let effects = |alpha| ShapeEffects {
            outer_shadow: Some(ooxml_drawingml::OuterShadow {
                color: Some(ColorValue {
                    rgb: Some("000000".to_owned()),
                    alpha: Some(alpha),
                    ..ColorValue::default()
                }),
                blur_radius: 76_200,
                distance: 38_100,
                direction: 2_700_000,
                ..Default::default()
            }),
        };
        let box_ = PxRect {
            x: 10.0,
            y: 20.0,
            w: 100.0,
            h: 50.0,
        };
        let resolved = shadow(
            &effects(0.4),
            &theme,
            Space::root(),
            box_,
            0.0,
            false,
            false,
        )
        .expect("a painted shadow");
        assert_eq!(resolved.color, "#00000066");
        assert!((resolved.blur - 8.0).abs() < 0.01);
        assert!((resolved.dx - 2.828).abs() < 0.01);
        assert!((resolved.dy - 2.828).abs() < 0.01);

        let rotated = shadow(
            &effects(0.4),
            &theme,
            Space::root(),
            box_,
            90.0,
            true,
            false,
        )
        .unwrap();
        assert!((rotated.dx + 2.828).abs() < 0.01);
        assert!((rotated.dy + 2.828).abs() < 0.01);
        let mut fixed = effects(0.4);
        fixed.outer_shadow.as_mut().unwrap().rotate_with_shape = false;
        assert_eq!(
            shadow(&fixed, &theme, Space::root(), box_, 90.0, true, false).unwrap(),
            resolved
        );

        assert_eq!(
            shadow(
                &effects(0.0),
                &theme,
                Space::root(),
                box_,
                0.0,
                false,
                false
            ),
            None
        );
        assert_eq!(
            shadow(
                &ShapeEffects::default(),
                &theme,
                Space::root(),
                box_,
                0.0,
                false,
                false
            ),
            None
        );
    }

    #[test]
    fn a_scaled_shadow_keeps_the_point_its_alignment_names() {
        let theme = Theme::default();
        let scaled = |alignment: &str, sx: f64| ShapeEffects {
            outer_shadow: Some(ooxml_drawingml::OuterShadow {
                color: Some(ColorValue {
                    rgb: Some("000000".to_owned()),
                    alpha: Some(0.4),
                    ..ColorValue::default()
                }),
                scale_x: sx,
                scale_y: sx,
                alignment: alignment.to_owned(),
                ..Default::default()
            }),
        };
        let box_ = PxRect {
            x: 100.0,
            y: 200.0,
            w: 40.0,
            h: 80.0,
        };
        let centre = shadow(
            &scaled("ctr", 1.02),
            &theme,
            Space::root(),
            box_,
            0.0,
            false,
            false,
        )
        .expect("a painted shadow");
        assert!((centre.scale_x - 1.02).abs() < 1e-6);
        // 102% about the box centre (120, 240) leaves it where it was.
        assert!((120.0 * centre.scale_x + centre.dx - 120.0).abs() < 0.01);
        assert!((240.0 * centre.scale_y + centre.dy - 240.0).abs() < 0.01);

        // The default anchor is the bottom edge, so only that edge stays put.
        let bottom = shadow(
            &scaled("b", 1.02),
            &theme,
            Space::root(),
            box_,
            0.0,
            false,
            false,
        )
        .expect("a painted shadow");
        assert!((280.0 * bottom.scale_y + bottom.dy - 280.0).abs() < 0.01);
        assert!(200.0 * bottom.scale_y + bottom.dy < 200.0);

        let plain = shadow(
            &scaled("ctr", 1.0),
            &theme,
            Space::root(),
            box_,
            0.0,
            false,
            false,
        )
        .expect("a painted shadow");
        assert_eq!((plain.scale_x, plain.dx, plain.dy), (1.0, 0.0, 0.0));

        for (alignment, x, y) in [
            ("tl", 100.0, 200.0),
            ("t", 120.0, 200.0),
            ("tr", 140.0, 200.0),
            ("l", 100.0, 240.0),
            ("ctr", 120.0, 240.0),
            ("r", 140.0, 240.0),
            ("bl", 100.0, 280.0),
            ("b", 120.0, 280.0),
            ("br", 140.0, 280.0),
        ] {
            let mut effects = scaled(alignment, -1.5);
            effects.outer_shadow.as_mut().unwrap().scale_y = 0.5;
            let result = shadow(&effects, &theme, Space::root(), box_, 0.0, false, false).unwrap();
            assert_eq!((result.scale_x, result.scale_y), (-1.5, 0.5));
            assert_eq!(result.dx, x * 2.5, "{alignment}");
            assert_eq!(result.dy, y * 0.5, "{alignment}");
        }
        assert!(
            shadow(
                &scaled("ctr", 0.0),
                &theme,
                Space::root(),
                box_,
                0.0,
                false,
                false
            )
            .is_none()
        );
        let invalid = shadow(
            &scaled("ctr", f64::INFINITY),
            &theme,
            Space::root(),
            box_,
            0.0,
            false,
            false,
        )
        .unwrap();
        assert_eq!(
            (invalid.scale_x, invalid.scale_y, invalid.dx, invalid.dy),
            (1.0, 1.0, 0.0, 0.0)
        );
    }

    #[test]
    fn a_slide_number_field_resolves_to_the_slide_it_is_drawn_on() {
        let run = |field_type: Option<&str>, text: &str| pptx_parse::TextRun {
            text: text.to_owned(),
            properties: RunProperties::default(),
            field_id: field_type.map(|_| "{GUID}".to_owned()),
            field_type: field_type.map(str::to_owned),
            line_break: false,
        };

        assert_eq!(
            field_text(&run(Some("slidenum"), "\u{2039}#\u{203a}"), 7),
            "7"
        );
        assert_eq!(
            field_text(&run(Some("datetime"), "16/08/2026"), 7),
            "16/08/2026"
        );
        assert_eq!(field_text(&run(None, "Chapter 3"), 7), "Chapter 3");
    }

    #[test]
    fn slide_number_fields_count_from_first_slide_num_through_the_edit_snapshot() {
        for first in [10, -3, i32::MAX] {
            assert_slide_number_fields(first);
        }
    }

    fn assert_slide_number_fields(first: i32) {
        let mut source = pptx_parse::parse_pptx(NUMBERED_FIXTURE).unwrap();
        assert_eq!(source.presentation.first_slide_num, 10);
        assert!(
            std::str::from_utf8(source.part_bytes("ppt/slides/slide1.xml").unwrap())
                .unwrap()
                .contains("show=\"0\"")
        );
        let presentation = std::str::from_utf8(source.part_bytes("ppt/presentation.xml").unwrap())
            .unwrap()
            .replace(
                "firstSlideNum=\"10\"",
                &format!("firstSlideNum=\"{first}\""),
            );
        assert!(source.replace_part("ppt/presentation.xml", presentation.into_bytes()));
        let bytes = pptx_parse::write_pptx(&source).unwrap();
        let parsed = pptx_parse::parse_pptx(&bytes).unwrap();
        let opened = DeckSession::from_package_with_source(parsed, &bytes, 8_006).unwrap();
        let session =
            DeckSession::open_from_update(&opened.encode_state_as_update_v1(), 8_007).unwrap();
        let package = session.package();
        assert_eq!(package.presentation.first_slide_num, first);
        let deck = session.snapshot().unwrap();
        assert_eq!(deck.slides.len(), 3);
        let master = &package.masters[0];
        let layout = &package.layouts[0];
        let master_probe = format!("master:{}:{}", master.part_path, master.shapes.len() - 1);
        let layout_probe = format!("layout:{}:{}", layout.part_path, layout.shapes.len() - 1);
        let slide_probe = deck.slides[1].shapes.last().unwrap();
        assert_eq!(slide_probe.text_stories[0].plain_text(), "77");

        let run = |text: &str, size: f32, emphasis: bool, color: &str| TextRun {
            text: text.to_owned(),
            font_family: "Arial".to_owned(),
            font_size_pt: size,
            bold: emphasis,
            italic: emphasis,
            underline: emphasis,
            color: color.to_owned(),
        };
        let renderer = renderer();
        for index in 0..3 {
            let number = (i64::from(first) + index as i64).to_string();
            let number = number.as_str();
            let rendered = renderer.layout_slide(package, &deck, index).unwrap();
            let (runs, line) = drawn_text(&rendered, &master_probe);
            assert_eq!(
                runs.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(),
                [number, "|", "CACHED-DATE"]
            );
            assert_eq!(runs[0], run(number, 23.0, true, "#FF0066"));
            assert_eq!(runs[2], run("CACHED-DATE", 13.0, false, "#00AA55"));
            assert_eq!(line, format!("{number}|CACHED-DATE"));
            let (runs, line) = drawn_text(&rendered, &layout_probe);
            assert_eq!(runs, [run(number, 17.0, false, "#1122CC")]);
            assert_eq!(line, number);
            if index == 1 {
                let (runs, line) = drawn_text(&rendered, &slide_probe.id);
                assert_eq!(runs, [run("77", 19.0, false, "#AA5500")]);
                assert_eq!(line, "77");
            }
        }
    }

    fn drawn_text(rendered: &RenderedSlide, shape_id: &str) -> (Vec<TextRun>, String) {
        rendered
            .display_list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::TextBox {
                    shape_id: Some(id),
                    paragraphs,
                    lines,
                    ..
                } if id == shape_id => Some((
                    paragraphs
                        .iter()
                        .flat_map(|paragraph| paragraph.runs.iter().cloned())
                        .collect(),
                    lines
                        .iter()
                        .flat_map(|line| line.runs.iter().map(|run| run.text.as_str()))
                        .collect(),
                )),
                _ => None,
            })
            .unwrap_or_else(|| panic!("{shape_id} was not drawn"))
    }

    #[test]
    fn hiding_master_shapes_drops_the_layout_decoration_too() {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_301).unwrap();
        let snapshot = session.snapshot().unwrap();
        package.layouts[0]
            .shapes
            .push(package.slides[0].shapes[0].clone());
        let shown = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        assert!(
            painted_shape_ids(&shown)
                .iter()
                .any(|id| id.starts_with("layout:"))
        );
        package.slides[0].show_master_shapes = false;
        let hidden = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        assert!(
            painted_shape_ids(&hidden)
                .iter()
                .all(|id| !id.starts_with("layout:") && !id.starts_with("master:"))
        );
        assert_eq!(
            hidden.display_list.background,
            shown.display_list.background
        );
    }

    #[test]
    fn a_background_picture_paints_first_from_the_slide_or_the_theme() {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_302).unwrap();
        let snapshot = session.snapshot().unwrap();
        let picture = PictureFill {
            relationship_id: None,
            media_part_path: Some("ppt/media/image1.png".to_owned()),
            crop: PictureCrop::default(),
            fill_rect: PictureCrop {
                top: -6_000,
                bottom: -6_000,
                ..PictureCrop::default()
            },
        };
        package.slides[0].background_picture = Some(Box::new(picture.clone()));
        let rendered = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        let Primitive::Image {
            object_id,
            asset_id,
            x,
            y,
            w,
            h,
            crop,
            ..
        } = &rendered.display_list.primitives[0]
        else {
            panic!("the background paints before every shape")
        };
        assert_eq!(*object_id, 0);
        assert_eq!(asset_id.as_deref(), Some("ppt/media/image1.png"));
        assert_eq!(
            (*x, *y, *w, *h),
            (
                0.0,
                0.0,
                rendered.display_list.width,
                rendered.display_list.height
            )
        );
        assert!(crop.top > 0.0 && crop.bottom > 0.0);

        package.slides[0].background_picture = None;
        package.slides[0].background = None;
        package.slides[0].background_reference = Some(StyleReference {
            index: 1_003,
            color: None,
        });
        package.themes[0].background_pictures = vec![None, None, Some(picture)];
        let referenced = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        assert_eq!(
            referenced.display_list.primitives[0],
            rendered.display_list.primitives[0]
        );
    }

    /// a deck whose master fills the slide at alpha 0 still renders on paper.
    #[test]
    fn a_fully_transparent_background_renders_as_white_paper() {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_311).unwrap();
        let snapshot = session.snapshot().unwrap();
        package.slides[0].background_reference = None;
        package.slides[0].background = Some(ShapeFill {
            fill_type: "solid".to_owned(),
            color: Some(ColorValue {
                rgb: Some("123456".to_owned()),
                alpha: Some(0.0),
                ..ColorValue::default()
            }),
            gradient: None,
        });
        let rendered = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        assert_eq!(
            rendered.display_list.background,
            Some(Paint::Solid {
                color: "#ffffff".to_owned()
            })
        );
    }

    /// a gradient whose every stop is transparent is paper too, while one
    /// visible stop keeps the gradient.
    #[test]
    fn a_gradient_paints_unless_every_stop_is_transparent() {
        let clear = |alpha: f64| ColorValue {
            rgb: Some("123456".to_owned()),
            alpha: Some(alpha),
            ..ColorValue::default()
        };
        let gradient = |stops: Vec<ColorValue>| ShapeFill {
            fill_type: "gradient".to_owned(),
            color: None,
            gradient: Some(ooxml_drawingml::GradientFill {
                gradient_type: "linear".to_owned(),
                angle: None,
                stops: stops
                    .into_iter()
                    .enumerate()
                    .map(|(index, color)| ooxml_drawingml::GradientStop {
                        position: index as f64,
                        color,
                    })
                    .collect(),
            }),
        };
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_317).unwrap();
        let snapshot = session.snapshot().unwrap();
        package.slides[0].background_reference = None;
        package.slides[0].background = Some(gradient(vec![clear(0.0), clear(0.0)]));
        assert_eq!(
            renderer()
                .layout_slide(&package, &snapshot, 0)
                .unwrap()
                .display_list
                .background,
            Some(Paint::Solid {
                color: "#ffffff".to_owned()
            })
        );
        package.slides[0].background = Some(gradient(vec![clear(0.0), clear(1.0)]));
        assert!(matches!(
            renderer()
                .layout_slide(&package, &snapshot, 0)
                .unwrap()
                .display_list
                .background,
            Some(Paint::Gradient { .. })
        ));
    }

    #[test]
    fn an_unpaintable_background_style_keeps_the_reference_colour() {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_303).unwrap();
        let snapshot = session.snapshot().unwrap();
        let color = ColorValue {
            rgb: Some("123456".to_owned()),
            ..ColorValue::default()
        };
        package.slides[0].background = Some(ShapeFill {
            fill_type: "theme".to_owned(),
            color: Some(color),
            gradient: None,
        });
        package.slides[0].background_reference = Some(StyleReference {
            index: 1_003,
            color: None,
        });
        package.themes[0].format_scheme.background_fills =
            vec![None, None, Some(ShapeFill::named("picture"))];
        package.themes[0].background_pictures = Vec::new();
        let rendered = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        assert_eq!(
            rendered.display_list.background,
            Some(Paint::Solid {
                color: "#123456".to_owned()
            })
        );
    }

    #[test]
    fn seeded_hidden_shapes_are_neither_painted_nor_hit_testable() {
        let package = pptx_parse::parse_pptx(HIDDEN_FIXTURE).unwrap();
        let session = DeckSession::open(HIDDEN_FIXTURE, 8_103).unwrap();
        let snapshot = session.snapshot().unwrap();
        let shapes = &snapshot.slides[0].shapes;
        assert!(shapes[0].hidden);
        assert!(shapes[8].hidden);
        assert_eq!(shapes[8].children.len(), 14);
        assert!(shapes[8].children[..13].iter().all(|child| !child.hidden));
        assert!(shapes[8].children[13].hidden);
        assert!(
            shapes
                .iter()
                .enumerate()
                .all(|(index, shape)| shape.hidden == (index == 0 || index == 8))
        );

        let mut visible = snapshot.clone();
        clear_hidden(&mut visible.slides[0].shapes);
        let before = renderer().layout_slide(&package, &visible, 0).unwrap();
        let after = renderer().layout_slide(&package, &snapshot, 0).unwrap();

        let hidden_ids: BTreeSet<String> = std::iter::once("slide:0:256:shape:0".to_owned())
            .chain((0..14).map(|child| format!("slide:0:256:shape:8.{child}")))
            .collect();
        let removed: BTreeSet<String> = painted_shape_ids(&before)
            .difference(&painted_shape_ids(&after))
            .cloned()
            .collect();
        assert_eq!(removed, hidden_ids);
        assert_eq!(before.display_list.primitives.len(), 41);
        assert_eq!(after.display_list.primitives.len(), 20);
        let expected: Vec<Primitive> = before
            .display_list
            .primitives
            .iter()
            .filter(|primitive| {
                !primitive_shape_id(primitive).is_some_and(|id| hidden_ids.contains(id))
            })
            .cloned()
            .collect();
        assert_eq!(after.display_list.primitives, expected);

        assert_eq!(
            before.hit_test(12.0, 200.0),
            Some(HitTestResult::Shape {
                shape_id: "slide:0:256:shape:0".to_owned()
            })
        );
        assert_eq!(after.hit_test(12.0, 200.0), None);
        let child = match before.hit_test(760.0, 100.0) {
            Some(HitTestResult::Shape { shape_id } | HitTestResult::Text { shape_id, .. }) => {
                shape_id
            }
            None => panic!("a child of the visible group is painted"),
        };
        assert!(child.starts_with("slide:0:256:shape:8."));
        assert_eq!(after.hit_test(760.0, 100.0), None);
    }

    #[test]
    fn a_v2_deck_renders_its_hidden_shapes_hidden_after_migration() {
        let fresh = DeckSession::open(HIDDEN_FIXTURE, 8_104).unwrap();
        let mut visible = fresh.snapshot().unwrap();
        let expected = renderer()
            .layout_slide(fresh.package(), &visible, 0)
            .unwrap();
        clear_hidden(&mut visible.slides[1].shapes);
        let painted_when_visible = painted_shape_ids(
            &renderer()
                .layout_slide(fresh.package(), &visible, 1)
                .unwrap(),
        );
        assert!(painted_when_visible.contains("slide:1:257:shape:16"));

        let session = DeckSession::open_from_update(V2_UPDATE, 8_105).unwrap();
        let snapshot = session.snapshot().unwrap();
        let slide = |id: &str| {
            snapshot
                .slides
                .iter()
                .position(|slide| slide.id == id)
                .unwrap()
        };
        let migrated = renderer()
            .layout_slide(session.package(), &snapshot, slide("slide:0:256"))
            .unwrap();
        assert_eq!(migrated.display_list, expected.display_list);
        assert_eq!(migrated.hit_test(12.0, 200.0), None);
        assert_eq!(migrated.hit_test(760.0, 100.0), None);

        let painted = painted_shape_ids(
            &renderer()
                .layout_slide(session.package(), &snapshot, slide("slide:1:257"))
                .unwrap(),
        );
        assert!(!painted.contains("slide:1:257:shape:16"));
        assert!(painted.contains("shape:4343:0"));
    }

    fn clear_hidden(shapes: &mut [ShapeSnapshot]) {
        for shape in shapes {
            shape.hidden = false;
            clear_hidden(&mut shape.children);
        }
    }

    fn primitive_shape_id(primitive: &Primitive) -> Option<&str> {
        match primitive {
            Primitive::Shape { shape_id, .. }
            | Primitive::Image { shape_id, .. }
            | Primitive::TextBox { shape_id, .. }
            | Primitive::Placeholder { shape_id, .. }
            | Primitive::Chart { shape_id, .. }
            | Primitive::Table { shape_id, .. } => shape_id.as_deref(),
        }
    }

    fn painted_shape_ids(slide: &RenderedSlide) -> BTreeSet<String> {
        slide
            .display_list
            .primitives
            .iter()
            .filter_map(primitive_shape_id)
            .map(str::to_owned)
            .collect()
    }

    #[test]
    fn hit_testing_a_rotated_shape_follows_its_painted_frame() {
        let slide = |transform| RenderedSlide {
            display_list: SurfaceDisplayList {
                contract_version: CONTRACT_VERSION,
                width: 400.0,
                height: 400.0,
                background: None,
                primitives: Vec::new(),
            },
            hit_regions: vec![HitRegion {
                shape_id: "rotated".to_owned(),
                rect: PxRect {
                    x: 100.0,
                    y: 100.0,
                    w: 100.0,
                    h: 100.0,
                },
                hit_rect: PxRect {
                    x: 100.0,
                    y: 100.0,
                    w: 100.0,
                    h: 100.0,
                },
                transform,
                text: None,
            }],
        };
        let turned = |rotation_deg| Transform {
            rotation_deg,
            ..Transform::default()
        };

        // turned 45°, the square's corners point along the axes: its bounding box
        // corners fall outside it and the axis midpoints fall inside
        assert!(slide(turned(0.0)).hit_test(105.0, 105.0).is_some());
        assert!(slide(turned(45.0)).hit_test(105.0, 105.0).is_none());
        assert!(slide(turned(0.0)).hit_test(150.0, 90.0).is_none());
        assert!(slide(turned(45.0)).hit_test(150.0, 90.0).is_some());

        // a flip maps the square onto itself, so membership is unchanged
        let flipped = Transform {
            rotation_deg: 45.0,
            flip_h: true,
            flip_v: false,
        };
        assert!(slide(flipped).hit_test(150.0, 90.0).is_some());
        assert!(slide(flipped).hit_test(105.0, 105.0).is_none());
    }

    #[test]
    fn hit_testing_flipped_text_keeps_the_upright_caret() {
        let slide = |flip_h| RenderedSlide {
            display_list: SurfaceDisplayList {
                contract_version: CONTRACT_VERSION,
                width: 400.0,
                height: 400.0,
                background: None,
                primitives: Vec::new(),
            },
            hit_regions: vec![HitRegion {
                shape_id: "mirrored".to_owned(),
                rect: PxRect {
                    x: 100.0,
                    y: 100.0,
                    w: 120.0,
                    h: 40.0,
                },
                hit_rect: PxRect {
                    x: 100.0,
                    y: 100.0,
                    w: 120.0,
                    h: 40.0,
                },
                transform: Transform {
                    rotation_deg: 0.0,
                    flip_h,
                    flip_v: false,
                },
                text: Some(TextHit {
                    overflow: false,
                    story_id: "story".to_owned(),
                    rect: PxRect {
                        x: 100.0,
                        y: 100.0,
                        w: 120.0,
                        h: 40.0,
                    },
                    transform: text_transform(
                        Transform {
                            rotation_deg: 0.0,
                            flip_h,
                            flip_v: false,
                        },
                        TextFlow::Horizontal,
                    ),
                    lines: vec![PositionedTextLine {
                        x: 110.0,
                        y: 110.0,
                        width: 100.0,
                        height: 20.0,
                        baseline: 125.0,
                        start: 0,
                        end: 5,
                        runs: Vec::new(),
                        caret_stops: vec![
                            CaretStop {
                                position: 0,
                                x: 110.0,
                            },
                            CaretStop {
                                position: 5,
                                x: 210.0,
                            },
                        ],
                    }],
                }),
            }],
        };
        let position = |flip_h| match slide(flip_h).hit_test(120.0, 120.0) {
            Some(HitTestResult::Text { position, .. }) => position,
            other => panic!("expected a text hit, got {other:?}"),
        };

        assert_eq!(position(false), 0);
        assert_eq!(position(true), 0);
    }

    #[test]
    fn hit_testing_vertical_text_reads_the_caret_in_the_turned_frame() {
        let shape = Transform {
            rotation_deg: 0.0,
            flip_h: false,
            flip_v: false,
        };
        let rect = PxRect {
            x: 100.0,
            y: 100.0,
            w: 40.0,
            h: 120.0,
        };
        let slide = RenderedSlide {
            display_list: SurfaceDisplayList {
                contract_version: CONTRACT_VERSION,
                width: 400.0,
                height: 400.0,
                background: None,
                primitives: Vec::new(),
            },
            hit_regions: vec![HitRegion {
                shape_id: "sideways".to_owned(),
                hit_rect: rect,
                rect,
                transform: shape,
                text: Some(TextHit {
                    overflow: false,
                    story_id: "story".to_owned(),
                    rect: TextFlow::Vert270.layout_rect(rect),
                    transform: text_transform(shape, TextFlow::Vert270),
                    lines: vec![PositionedTextLine {
                        x: 60.0,
                        y: 150.0,
                        width: 100.0,
                        height: 20.0,
                        baseline: 165.0,
                        start: 0,
                        end: 5,
                        runs: Vec::new(),
                        caret_stops: vec![
                            CaretStop {
                                position: 0,
                                x: 60.0,
                            },
                            CaretStop {
                                position: 5,
                                x: 160.0,
                            },
                        ],
                    }],
                }),
            }],
        };
        let position = |y| match slide.hit_test(120.0, y) {
            Some(HitTestResult::Text { position, .. }) => position,
            other => panic!("expected a text hit, got {other:?}"),
        };

        assert_eq!(position(205.0), 0);
        assert_eq!(position(115.0), 5);
    }

    #[test]
    fn text_follows_the_shape_rotation_but_is_never_mirrored() {
        let turned = |rotation_deg, flip_h, flip_v| {
            text_transform(
                Transform {
                    rotation_deg,
                    flip_h,
                    flip_v,
                },
                TextFlow::Horizontal,
            )
        };
        let upright = |rotation_deg: f32| Transform {
            rotation_deg,
            flip_h: false,
            flip_v: false,
        };

        assert_eq!(turned(180.0, false, true), upright(0.0));
        assert_eq!(turned(0.0, true, false), upright(0.0));
        assert_eq!(turned(90.0, true, false), upright(90.0));
        assert_eq!(turned(0.0, true, true), upright(180.0));
        assert_eq!(turned(45.0, false, false), upright(45.0));
    }

    #[test]
    fn body_vert_turns_the_text_and_lays_it_out_in_the_swapped_box() {
        assert_eq!(TextFlow::from_body_vert(Some("vert")), TextFlow::Vert);
        assert_eq!(TextFlow::from_body_vert(Some("vert270")), TextFlow::Vert270);
        assert_eq!(TextFlow::from_body_vert(Some("horz")), TextFlow::Horizontal);
        assert_eq!(TextFlow::from_body_vert(None), TextFlow::Horizontal);
        assert_eq!(TextFlow::from_body_vert(Some("eaVert")), TextFlow::Vert);
        assert_eq!(
            TextFlow::from_body_vert(Some("mongolianVert")),
            TextFlow::VertLeftToRight
        );
        for mode in ["wordArtVert", "wordArtVertRtl"] {
            assert_eq!(TextFlow::from_body_vert(Some(mode)), TextFlow::Stacked);
        }

        let shape = |rotation_deg| Transform {
            rotation_deg,
            flip_h: false,
            flip_v: false,
        };
        assert_eq!(
            text_transform(shape(90.0), TextFlow::Vert270).rotation_deg,
            0.0
        );
        assert_eq!(
            text_transform(shape(270.0), TextFlow::Vert).rotation_deg,
            0.0
        );

        let rect = PxRect {
            x: 100.0,
            y: 0.0,
            w: 40.0,
            h: 240.0,
        };
        assert_eq!(
            TextFlow::Vert.layout_rect(rect),
            PxRect {
                x: 0.0,
                y: 100.0,
                w: 240.0,
                h: 40.0,
            }
        );
        assert_eq!(TextFlow::Horizontal.layout_rect(rect), rect);
        assert_eq!(
            TextFlow::VertLeftToRight.layout_rect(rect),
            TextFlow::Vert.layout_rect(rect)
        );
        assert_eq!(TextFlow::Stacked.layout_rect(rect), rect);
        assert_eq!(TextFlow::Stacked.rotation_deg(), 0.0);
    }

    /// Renders a synthetic master text box.
    fn text_box_on_master(
        transform: ShapeTransform,
        vertical: Option<&str>,
        text: &str,
    ) -> Primitive {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_010).unwrap();
        package.masters[0]
            .shapes
            .push(ShapeNode::Shape(pptx_parse::Shape {
                has_preset_geometry: true,
                effects: None,
                base: pptx_parse::ShapeBase {
                    id: 9_001,
                    name: "turned".to_owned(),
                    description: None,
                    hidden: false,
                    placeholder: None,
                    transform,
                },
                geometry: "rect".to_owned(),
                adjust_values: BTreeMap::new(),
                paths: Vec::new(),
                style: None,
                fill: None,
                picture_fill: None,
                outline: None,
                text: Some(TextBody {
                    vertical_overflow: None,
                    horizontal_overflow: None,
                    anchor: Some("ctr".to_owned()),
                    vertical: vertical.map(str::to_owned),
                    compat_line_spacing: None,
                    autofit: None,
                    inset_left: None,
                    inset_top: None,
                    inset_right: None,
                    inset_bottom: None,
                    default_list_style: None,
                    list_style: Vec::new(),
                    paragraphs: vec![pptx_parse::TextParagraph {
                        properties: ParagraphProperties::default(),
                        runs: vec![pptx_parse::TextRun {
                            text: text.to_owned(),
                            properties: RunProperties {
                                font_size_pt: Some(12.0),
                                font_family: Some("Arial".to_owned()),
                                ..RunProperties::default()
                            },
                            field_id: None,
                            field_type: None,
                            line_break: false,
                        }],
                        end_properties: None,
                    }],
                }),
            }));
        renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap()
            .display_list
            .primitives
            .into_iter()
            .find(|primitive| {
                matches!(
                    primitive,
                    Primitive::TextBox {
                        object_id: 9_001,
                        ..
                    }
                )
            })
            .expect("the master text box was laid out")
    }

    #[test]
    fn a_flipped_shape_lays_its_text_out_unmirrored() {
        let primitive = text_box_on_master(
            ShapeTransform {
                x: 1_000_000,
                y: 1_000_000,
                width: 4_000_000,
                height: 1_000_000,
                rotation_deg: 180.0,
                flip_v: true,
                ..ShapeTransform::default()
            },
            None,
            "Audit",
        );
        let Primitive::TextBox { transform, .. } = &primitive else {
            unreachable!()
        };
        assert_eq!(
            *transform,
            Transform {
                rotation_deg: 0.0,
                flip_h: false,
                flip_v: false,
            }
        );
    }

    #[test]
    fn a_vert270_label_bar_lays_its_sentence_out_on_one_line() {
        let turned = ShapeTransform {
            x: 4_601_452,
            y: -2_333_065,
            width: 639_990,
            height: 7_718_032,
            rotation_deg: 90.0,
            ..ShapeTransform::default()
        };
        let sentence = "Take the Shadow IT and Data Risk Assessments";
        let primitive = text_box_on_master(turned.clone(), Some("vert270"), sentence);
        let Primitive::TextBox {
            w,
            h,
            lines,
            transform,
            ..
        } = &primitive
        else {
            unreachable!()
        };
        assert!(w > h, "the layout box is the shape box on its side");
        assert_eq!(transform.rotation_deg, 0.0);
        assert_eq!(lines.len(), 1);

        let upright = text_box_on_master(turned, None, sentence);
        let Primitive::TextBox { lines, .. } = &upright else {
            unreachable!()
        };
        assert!(lines.len() > 1);
    }

    #[test]
    fn lays_out_demo_with_master_shapes_geometry_and_glyphs() {
        let package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_001).unwrap();
        let rendered = renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap();
        assert_eq!(
            (rendered.display_list.width, rendered.display_list.height),
            (1280.0, 720.0)
        );
        let shape_index = rendered
            .display_list
            .primitives
            .iter()
            .position(|primitive| matches!(primitive, Primitive::Shape { .. }))
            .unwrap();
        let text_index = rendered
            .display_list
            .primitives
            .iter()
            .position(|primitive| matches!(primitive, Primitive::TextBox { lines, .. } if !lines.is_empty()))
            .unwrap();
        assert!(shape_index < text_index);
        assert!(rendered.display_list.primitives.iter().any(|primitive| {
            matches!(primitive, Primitive::Shape { path, .. } if !path.is_empty())
        }));
        assert!(rendered.display_list.primitives.iter().any(|primitive| {
            matches!(primitive, Primitive::TextBox { lines, .. } if lines.iter().flat_map(|line| &line.runs).any(|run| !run.glyphs.is_empty()))
        }));
        assert!(rendered.display_list.primitives.iter().any(|primitive| {
            matches!(primitive, Primitive::TextBox { shape_id: Some(id), .. } if id.starts_with("master:"))
        }));
        let master_index = rendered
            .display_list
            .primitives
            .iter()
            .position(|primitive| {
                matches!(primitive, Primitive::Shape { shape_id: Some(id), .. } if id.starts_with("master:"))
            })
            .unwrap();
        let slide_index = rendered
            .display_list
            .primitives
            .iter()
            .position(|primitive| {
                matches!(primitive, Primitive::Shape { shape_id: Some(id), .. } if id.starts_with("slide:"))
            })
            .unwrap();
        assert!(master_index < slide_index);
    }

    #[test]
    fn a_baseline_shifted_run_is_raised_lowered_and_shrunk() {
        let session = DeckSession::open(
            include_bytes!("../tests/fixtures/text-baseline-script.pptx"),
            289,
        )
        .unwrap();
        let rendered = renderer()
            .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
            .unwrap();
        let Some(Primitive::TextBox { lines, .. }) =
            rendered.display_list.primitives.iter().find(|primitive| {
                matches!(primitive, Primitive::TextBox { lines, .. }
                    if lines.iter().flat_map(|line| &line.runs).any(|run| run.text.starts_with("E = mc")))
            })
        else {
            panic!("expected the script text box");
        };
        let line = lines
            .iter()
            .find(|line| line.runs.iter().any(|run| run.text == "E = mc"))
            .unwrap();
        let base = &line.runs[0];
        let raised = &line.runs[1];
        let plain = &line.runs[2];
        let lowered = &line.runs[3];
        assert_eq!(
            [
                base.text.as_str(),
                raised.text.as_str(),
                plain.text.as_str(),
                lowered.text.as_str()
            ],
            ["E = mc", "2", " and H", "2"]
        );
        assert_eq!(base.baseline_offset_px, 0.0);
        assert_eq!(plain.baseline_offset_px, 0.0);
        assert_eq!(base.font_size_px, plain.font_size_px);

        let em = points_to_px(17.0);
        assert!((raised.baseline_offset_px - em * 0.30).abs() < 0.01);
        assert!((lowered.baseline_offset_px + em * 0.25).abs() < 0.01);
        assert!((raised.font_size_px - 13.146666).abs() < 0.01);
        assert!((lowered.font_size_px - 13.146666).abs() < 0.01);
        assert!(
            (raised.glyphs[0].y_offset - (base.glyphs[0].y_offset - raised.baseline_offset_px))
                .abs()
                < 0.01
        );
        assert!(lowered.glyphs[0].y_offset > base.glyphs[0].y_offset);
        assert!(line.baseline - raised.baseline_offset_px - raised.font_size_px >= line.y);
    }

    #[test]
    fn baseline_cascade_keeps_zero_overrides_and_expands_line_metrics() {
        let mut package = pptx_parse::parse_pptx(include_bytes!(
            "../tests/fixtures/text-baseline-script.pptx"
        ))
        .unwrap();
        let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[3] else {
            panic!("shape")
        };
        let body = shape.text.as_mut().unwrap();
        body.list_style = vec![pptx_parse::ParagraphProperties {
            default_run: Some(RunProperties {
                baseline_pct: Some(150.0),
                ..Default::default()
            }),
            ..Default::default()
        }];
        body.paragraphs[0]
            .properties
            .default_run
            .as_mut()
            .unwrap()
            .baseline_pct = None;
        body.paragraphs[0].runs = [None, Some(-150.0), Some(0.0)]
            .into_iter()
            .enumerate()
            .map(|(i, baseline)| pptx_parse::TextRun {
                text: ["A", "B", "C"][i].into(),
                properties: RunProperties {
                    font_size_pt: Some(17.0),
                    baseline_pct: baseline,
                    ..Default::default()
                },
                field_id: None,
                field_type: None,
                line_break: false,
            })
            .collect();
        let session = DeckSession::from_package(package, 331).unwrap();
        let rendered = renderer()
            .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
            .unwrap();
        let Primitive::TextBox { lines, .. } = rendered
            .display_list
            .primitives
            .iter()
            .find(|p| matches!(p, Primitive::TextBox { object_id: 5, .. }))
            .unwrap()
        else {
            panic!("text")
        };
        let line = &lines[0];
        assert_eq!(line.runs.len(), 3);
        for (run, shift, size) in [
            (&line.runs[0], 34.0, 13.146666),
            (&line.runs[1], -34.0, 13.146666),
            (&line.runs[2], 0.0, 22.666666),
        ] {
            assert!((run.baseline_offset_px - shift).abs() < 0.001);
            assert!((run.font_size_px - size).abs() < 0.001);
            assert!((run.glyphs[0].y_offset - (line.baseline - shift)).abs() < 0.001);
        }
        assert!(line.baseline - line.y > 34.0);
        assert!(line.y + line.height - line.baseline > 34.0);
    }

    #[test]
    fn edited_text_reflows_and_hit_testing_returns_a_story_position() {
        let package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_002).unwrap();
        let initial = session.snapshot().unwrap();
        let slide_id = initial.slides[0].id.clone();
        let shape = initial.slides[0]
            .shapes
            .iter()
            .find(|shape| !shape.text_stories.is_empty())
            .unwrap();
        let shape_id = shape.id.clone();
        let story_id = shape.text_stories[0].id.clone();
        let index = shape.text_stories[0].length - 1;
        session
            .resize_shape(
                &EditCtx::local("test"),
                &slide_id,
                &shape_id,
                1_200_000,
                1_524_000,
            )
            .unwrap();
        session
            .insert_text(
                &EditCtx::local("test"),
                &story_id,
                index,
                " with enough collaborative text to wrap across several shaped lines",
                &TextStyle::default(),
            )
            .unwrap();
        let rendered = renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap();
        let (line_count, first_line) = rendered
            .display_list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::TextBox {
                    shape_id: Some(id),
                    lines,
                    ..
                } if id == &shape_id => Some((lines.len(), lines.first().unwrap())),
                _ => None,
            })
            .unwrap();
        assert!(line_count > 1);
        assert!(matches!(
            rendered.hit_test(first_line.x + 1.0, first_line.y + 1.0),
            Some(HitTestResult::Text {
                story_id: hit_story,
                ..
            }) if hit_story == story_id
        ));
    }

    /// Wrap the demo text with inherited spacing.
    fn wrapped_text_box(
        seed: u64,
        spacing: Option<LineSpacing>,
        paragraph_spacing: Option<LineSpacing>,
        height_emu: i64,
        compat: bool,
    ) -> (Vec<PositionedTextLine>, f32) {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, seed).unwrap();
        let initial = session.snapshot().unwrap();
        let slide_id = initial.slides[0].id.clone();
        let shape = initial.slides[0]
            .shapes
            .iter()
            .find(|shape| !shape.text_stories.is_empty())
            .unwrap();
        let shape_id = shape.id.clone();
        let source_id = shape.source_id;
        let story_id = shape.text_stories[0].id.clone();
        let index = shape.text_stories[0].length - 1;
        for master in &mut package.masters {
            for level in [
                &mut master.text_styles.title,
                &mut master.text_styles.body,
                &mut master.text_styles.other,
            ] {
                *level = vec![ParagraphProperties {
                    line_spacing: spacing,
                    ..ParagraphProperties::default()
                }];
            }
        }
        let parsed = package.slides[0]
            .shapes
            .iter_mut()
            .find(|shape| shape.id() == source_id)
            .unwrap();
        let ShapeNode::Shape(parsed) = parsed else {
            panic!("expected text shape");
        };
        let body = parsed.text.as_mut().unwrap();
        body.compat_line_spacing = compat.then_some(true);
        body.anchor = Some("t".to_owned());
        for paragraph in &mut body.paragraphs {
            paragraph.properties.line_spacing = paragraph_spacing;
        }
        session
            .resize_shape(
                &EditCtx::local("test"),
                &slide_id,
                &shape_id,
                1_200_000,
                height_emu,
            )
            .unwrap();
        session
            .insert_text(
                &EditCtx::local("test"),
                &story_id,
                index,
                " with enough text to wrap across several shaped lines",
                &TextStyle::default(),
            )
            .unwrap();
        renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap()
            .display_list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::TextBox {
                    shape_id: Some(id),
                    lines,
                    paragraphs,
                    ..
                } if id == &shape_id => Some((lines.clone(), paragraphs[0].runs[0].font_size_pt)),
                _ => None,
            })
            .unwrap()
    }

    #[test]
    fn master_line_spacing_sets_the_line_pitch() {
        let percent = |value| Some(LineSpacing::Percent { value });
        let (single, size_pt) = wrapped_text_box(8_010, None, None, 1_524_000, true);
        let (tight, _) = wrapped_text_box(8_011, percent(0.8), None, 1_524_000, true);
        let (loose, _) = wrapped_text_box(
            8_012,
            Some(LineSpacing::Points { value: 40.0 }),
            None,
            1_524_000,
            true,
        );
        let (no_compat, _) = wrapped_text_box(8_017, percent(0.8), None, 1_524_000, false);
        let (tighter, _) = wrapped_text_box(8_018, percent(0.6), None, 1_524_000, true);
        let pitch = |lines: &[PositionedTextLine]| {
            assert!(lines.len() > 1, "expected the text to wrap");
            lines[1].y - lines[0].y
        };
        let size_px = points_to_px(size_pt);

        assert!((pitch(&single) - 1.2 * size_px).abs() < 0.05);
        assert!((pitch(&tight) - 0.8 * 1.2 * size_px).abs() < 0.05);
        assert!((pitch(&no_compat) - pitch(&tight)).abs() < 0.05);
        assert!(pitch(&tight) < pitch(&single));
        assert!((pitch(&loose) - points_to_px(40.0)).abs() < 0.05);

        let share =
            |lines: &[PositionedTextLine]| (lines[0].baseline - lines[0].y) / lines[0].height;
        assert!((share(&tight) - share(&tighter)).abs() < 0.01);
    }

    #[test]
    fn a_slide_paragraph_overrides_the_master_line_spacing() {
        let (lines, size_pt) = wrapped_text_box(
            8_015,
            Some(LineSpacing::Percent { value: 0.8 }),
            Some(LineSpacing::Percent { value: 1.5 }),
            1_524_000,
            true,
        );
        assert!(lines.len() > 1, "expected the text to wrap");
        let pitch = lines[1].y - lines[0].y;
        assert!((pitch - 1.5 * 1.2 * points_to_px(size_pt)).abs() < 0.05);

        // The room the spacing adds is split above and below the line, so a
        // looser paragraph starts lower by half of what it added.
        let (single, _) = wrapped_text_box(8_016, None, None, 1_524_000, true);
        let added = (lines[0].height - single[0].height) / 2.0;
        assert!((lines[0].baseline - single[0].baseline - added).abs() < 0.001);
    }

    #[test]
    fn adjacent_runs_keep_their_own_tracking() {
        let renderer = renderer();
        let mut paragraph = paragraph(&renderer, "l", "AA");
        let mut second = ResolvedRun {
            text: "BB".into(),
            start: 2,
            style: paragraph.runs[0].style.clone(),
        };
        second.style.spacing_pt = 6.0;
        paragraph.runs.push(second);
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 0.0, 0.0, 1000.0, 1.0, false).unwrap();
        assert_eq!(lines[0].runs.len(), 2);
        assert_eq!(lines[0].runs[0].letter_spacing_px, 0.0);
        assert_eq!(lines[0].runs[1].letter_spacing_px, 8.0);
    }

    #[test]
    fn tracking_separates_optional_ligatures_and_preserves_combining_clusters() {
        let mut renderer = SlideRenderer::new();
        renderer
            .register_font(
                "Arial",
                false,
                false,
                include_bytes!("../../../packages/fonts/assets/Carlito-Regular.ttf"),
            )
            .unwrap();
        let mut paragraph = paragraph(&renderer, "l", "fi a\u{301}");
        let plain = shape_paragraph(&renderer.fonts, &paragraph, 1.0).unwrap();
        assert!(plain.iter().any(|cluster| cluster.text == "fi"));
        paragraph.runs[0].style.spacing_pt = 3.0;
        let tracked = shape_paragraph(&renderer.fonts, &paragraph, 1.0).unwrap();
        assert_eq!(
            tracked
                .iter()
                .map(|cluster| cluster.text.as_str())
                .collect::<Vec<_>>(),
            ["f", "i", " ", "a\u{301}"]
        );
    }

    #[test]
    fn tracking_justification_reaches_the_right_edge() {
        let renderer = renderer();
        let mut paragraph = paragraph(&renderer, "just", "AA BB CC AA BB CC");
        paragraph.runs[0].style.spacing_pt = 6.0;
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 0.0, 0.0, 160.0, 1.0, false).unwrap();
        assert!(lines.len() > 1);
        assert!((lines[0].width - 160.0).abs() < 0.001, "{}", lines[0].width);
    }

    #[test]
    fn tracking_before_a_hard_break_has_no_trailing_gap() {
        let renderer = renderer();
        let mut paragraph = paragraph(&renderer, "ctr", "AA\nAA");
        paragraph.runs[0].style.spacing_pt = 6.0;
        let lines =
            layout_paragraph(&renderer.fonts, &paragraph, 0.0, 0.0, 1000.0, 1.0, false).unwrap();
        assert_eq!(lines.len(), 2);
        assert!((lines[0].width - lines[1].width).abs() < 0.001);
        assert!((lines[0].x - lines[1].x).abs() < 0.001);
        let tight = layout_paragraph(
            &renderer.fonts,
            &paragraph,
            0.0,
            0.0,
            lines[1].width + 0.001,
            1.0,
            false,
        )
        .unwrap();
        assert_eq!(tight.len(), 2);
    }

    #[test]
    fn shape_autofit_keeps_tracked_font_size() {
        let session =
            DeckSession::open(include_bytes!("../tests/fixtures/run-spacing.pptx"), 32510).unwrap();
        let snapshot = session.snapshot().unwrap();
        let id = &snapshot.slides[0].shapes[2].id;
        let rendered = renderer()
            .layout_slide(session.package(), &snapshot, 0)
            .unwrap();
        let lines = rendered
            .display_list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::TextBox {
                    shape_id: Some(shape_id),
                    lines,
                    ..
                } if shape_id == id => Some(lines),
                _ => None,
            })
            .unwrap();
        assert!(lines.len() > 1);
        assert!(
            lines
                .iter()
                .flat_map(|line| &line.runs)
                .all(|run| (run.font_size_px - points_to_px(32.0)).abs() < 0.01)
        );
    }

    /// Formats the demo title for tracking assertions.
    fn tracked_text_box(
        seed: u64,
        spacing_pt: Option<f64>,
        align: Option<&str>,
        width_emu: i64,
    ) -> Vec<PositionedTextLine> {
        let package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, seed).unwrap();
        let initial = session.snapshot().unwrap();
        let slide_id = initial.slides[0].id.clone();
        let shape = initial.slides[0]
            .shapes
            .iter()
            .find(|shape| !shape.text_stories.is_empty())
            .unwrap();
        let shape_id = shape.id.clone();
        let story_id = shape.text_stories[0].id.clone();
        let context = EditCtx::local("test");
        session
            .resize_shape(&context, &slide_id, &shape_id, width_emu, 3_000_000)
            .unwrap();
        let end = session.story(&story_id).unwrap().length.saturating_sub(1);
        session
            .format_text(
                &context,
                &story_id,
                0,
                end,
                &TextStylePatch {
                    spacing_pt,
                    ..TextStylePatch::default()
                },
            )
            .unwrap();
        session
            .set_paragraph_alignment(&context, &story_id, 0, end, align)
            .unwrap();
        renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap()
            .display_list
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::TextBox {
                    shape_id: Some(id),
                    lines,
                    ..
                } if id == &shape_id => Some(lines.clone()),
                _ => None,
            })
            .unwrap()
    }

    #[test]
    fn run_spacing_tracks_the_line_and_keeps_it_centred() {
        let wide = 5_000_000;
        let plain = tracked_text_box(9_001, None, Some("ctr"), wide);
        let tracked = tracked_text_box(9_002, Some(6.0), Some("ctr"), wide);
        assert_eq!(plain.len(), 1);
        assert_eq!(tracked.len(), 1);

        let gaps = (plain[0].end - plain[0].start - 1) as f32;
        let gap_px = points_to_px(6.0);
        assert!(
            (tracked[0].width - plain[0].width - gaps * gap_px).abs() < 0.5,
            "{} vs {}",
            tracked[0].width,
            plain[0].width
        );
        let centre = |line: &PositionedTextLine| line.x + line.width / 2.0;
        assert!((centre(&tracked[0]) - centre(&plain[0])).abs() < 0.5);

        let caret = tracked[0].caret_stops.last().unwrap().x;
        assert!((caret - (tracked[0].x + tracked[0].width + gap_px)).abs() < 0.5);
    }

    #[test]
    fn run_spacing_moves_the_wrap_point() {
        let plain = tracked_text_box(9_003, None, None, 1_500_000);
        let tracked = tracked_text_box(9_004, Some(6.0), None, 1_500_000);

        assert!(
            tracked.len() > plain.len(),
            "{} vs {}",
            tracked.len(),
            plain.len()
        );
        assert!(tracked[0].end < plain[0].end);
    }

    #[test]
    fn tightening_past_a_glyph_advance_never_widens_the_line() {
        let wide = 5_000_000;
        let plain = tracked_text_box(9_005, None, None, wide);
        let tight = tracked_text_box(9_006, Some(-4_000.0), None, wide);

        assert_eq!(tight.len(), plain.len());
        assert!(
            tight[0].width <= plain[0].width,
            "{} vs {}",
            tight[0].width,
            plain[0].width
        );
        assert!(tight[0].width >= 0.0, "{}", tight[0].width);
    }

    #[test]
    fn normal_autofit_steps_the_type_down_by_its_stored_scale() {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_003).unwrap();
        let initial = session.snapshot().unwrap();
        let slide_id = initial.slides[0].id.clone();
        let shape = initial.slides[0]
            .shapes
            .iter()
            .find(|shape| !shape.text_stories.is_empty())
            .unwrap();
        let shape_id = shape.id.clone();
        let source_id = shape.source_id;
        let story_id = shape.text_stories[0].id.clone();
        let index = shape.text_stories[0].length - 1;
        let parsed = package.slides[0]
            .shapes
            .iter_mut()
            .find(|shape| shape.id() == source_id)
            .unwrap();
        let ShapeNode::Shape(parsed) = parsed else {
            panic!("expected text shape");
        };
        parsed.text.as_mut().unwrap().autofit = Some(TextAutofit::Normal {
            font_scale: None,
            line_space_reduction: None,
        });
        session
            .resize_shape(
                &EditCtx::local("test"),
                &slide_id,
                &shape_id,
                2_000_000,
                500_000,
            )
            .unwrap();
        session
            .insert_text(
                &EditCtx::local("test"),
                &story_id,
                index,
                " text that must shrink",
                &TextStyle::default(),
            )
            .unwrap();
        let snapshot = session.snapshot().unwrap();
        let font_size = |package: &PptxPackage| {
            renderer()
                .layout_slide(package, &snapshot, 0)
                .unwrap()
                .display_list
                .primitives
                .iter()
                .find_map(|primitive| match primitive {
                    Primitive::TextBox {
                        shape_id: Some(id),
                        paragraphs,
                        ..
                    } if id == &shape_id => Some(paragraphs[0].runs[0].font_size_pt),
                    _ => None,
                })
                .unwrap()
        };
        let scaled = font_size(&package);
        let ShapeNode::Shape(parsed) = package.slides[0]
            .shapes
            .iter_mut()
            .find(|shape| shape.id() == source_id)
            .unwrap()
        else {
            panic!("expected text shape");
        };
        parsed.text.as_mut().unwrap().autofit = Some(TextAutofit::None);
        let natural = font_size(&package);
        assert!((scaled - natural).abs() < 0.001);
        let ShapeNode::Shape(parsed) = package.slides[0]
            .shapes
            .iter_mut()
            .find(|shape| shape.id() == source_id)
            .unwrap()
        else {
            panic!("expected text shape");
        };
        parsed.text.as_mut().unwrap().autofit = Some(TextAutofit::Normal {
            font_scale: Some(0.5),
            line_space_reduction: None,
        });
        // The scale lands the type on a whole point, as PowerPoint's own
        // shrink-to-fit does.
        assert!((font_size(&package) - (natural * 0.5).round()).abs() < 0.001);
    }

    #[test]
    fn a_gradient_outline_strokes_with_its_gradient_and_keeps_a_flat_fallback() {
        let package = pptx_parse::parse_pptx(GRADIENT_OUTLINE_FIXTURE).unwrap();
        let session = DeckSession::open(GRADIENT_OUTLINE_FIXTURE, 8_010).unwrap();
        let primitives = renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap()
            .display_list
            .primitives;
        let strokes = primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape { name, stroke, .. } => Some((name.clone(), stroke.clone()?)),
                _ => None,
            })
            .collect::<Vec<_>>();

        let (_, gradient) = strokes
            .iter()
            .find(|(name, _)| name == "Gradient outline rectangle")
            .expect("the gradient-stroked rectangle must reach the display list");
        assert_eq!(
            gradient.paint,
            Some(Paint::Gradient {
                gradient_type: GradientType::Linear,
                angle_deg: Some(0.0),
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: "#C00000".to_owned(),
                    },
                    GradientStop {
                        position: 0.5,
                        color: "#FFC000".to_owned(),
                    },
                    GradientStop {
                        position: 1.0,
                        color: "#1F7A3D".to_owned(),
                    },
                ],
            })
        );
        assert_eq!(gradient.color, "#C00000");
        assert_eq!(gradient.width, 8.0);

        assert!(
            strokes
                .iter()
                .any(|(name, stroke)| name == "Gradient outline line" && stroke.paint.is_some())
        );

        let (_, solid) = strokes
            .iter()
            .find(|(name, _)| name == "Solid outline rectangle")
            .expect("the solid-stroked rectangle must reach the display list");
        assert_eq!(solid.paint, None);
        assert_eq!(solid.color, "#C00000");
    }

    #[test]
    fn gradient_outline_keeps_sorted_alpha_paint_through_theme_fallback() {
        let gradient: ShapeOutline = serde_json::from_value(serde_json::json!({
            "gradient": {
                "type": "linear", "angle": 0.0,
                "stops": [
                    {"position": 100000.0, "color": {"rgb": "0000FF"}},
                    {"position": 50000.0, "color": {"rgb": "00FF00"}},
                    {"position": 0.0, "color": {"rgb": "FF0000", "alpha": 0.5}},
                    {"position": 50000.0, "color": {"rgb": "FFFF00"}}
                ]
            }
        }))
        .unwrap();
        let solid = ShapeOutline {
            color: Some(ColorValue {
                rgb: Some("123456".into()),
                ..Default::default()
            }),
            width: Some(76_200.0),
            head_end: Some(LineEnd {
                end_type: "triangle".into(),
                width: None,
                length: None,
            }),
            ..Default::default()
        };
        let merged = merge_outline(&gradient, Some(&solid));
        let stroke = stroke(&merged, &Theme::default()).unwrap();
        assert_eq!(stroke.width, 8.0);
        assert_eq!(stroke.color, "#FF000080");
        assert!(stroke.head_end.is_some());
        let Some(Paint::Gradient { stops, .. }) = stroke.paint else {
            panic!("gradient")
        };
        assert_eq!(
            stops
                .iter()
                .map(|stop| (stop.position, stop.color.as_str()))
                .collect::<Vec<_>>(),
            [
                (0.0, "#FF000080"),
                (0.5, "#00FF00"),
                (0.5, "#FFFF00"),
                (1.0, "#0000FF")
            ]
        );
        assert!(merge_outline(&solid, Some(&gradient)).gradient.is_none());
        assert_eq!(
            merge_outline(
                &ShapeOutline {
                    width: Some(12_700.0),
                    ..Default::default()
                },
                Some(&gradient)
            )
            .gradient,
            gradient.gradient
        );
    }

    fn chart_slide(index: usize) -> Vec<Primitive> {
        let package = pptx_parse::parse_pptx(CHART_FIXTURE).unwrap();
        let session = DeckSession::open(CHART_FIXTURE, 8_004 + index as u64).unwrap();
        renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), index)
            .unwrap()
            .display_list
            .primitives
    }

    #[test]
    fn a_chart_frame_plots_with_the_deck_theme_and_an_aria_label() {
        let primitives = chart_slide(0);
        let (label, parts, rect) = primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Chart {
                    label,
                    primitives,
                    x,
                    y,
                    w,
                    h,
                    name,
                    shape_id: Some(_),
                    ..
                } if name == "Revenue chart" => Some((label, primitives, (*x, *y, *w, *h))),
                _ => None,
            })
            .expect("the chart frame plots");
        assert_eq!(label, "Revenue, column chart, 2 series, 3 categories");
        assert_eq!(rect, (96.0, 96.0, 576.0, 336.0));
        let fills = parts
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape {
                    fill: Some(Paint::Solid { color }),
                    ..
                } => Some(color.as_str()),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(fills.contains(&"#6254E7"), "theme accent1 bars are missing");
        assert!(fills.contains(&"#1FA97A"), "theme accent2 bars are missing");
        let text = parts
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::TextBox { lines, .. } => Some(
                    lines
                        .iter()
                        .flat_map(|line| line.runs.iter())
                        .map(|run| run.text.clone())
                        .collect::<String>(),
                ),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(text.contains(&"Revenue".to_owned()));
        assert!(text.contains(&"North".to_owned()));
        assert!(text.contains(&"Q1".to_owned()));
        assert!(text.contains(&"12".to_owned()), "data labels are missing");
        assert!(
            text.contains(&"Quarter".to_owned()),
            "no category axis title"
        );
        assert!(text.contains(&"Millions".to_owned()), "no value axis title");
        assert!(
            parts.iter().any(|primitive| matches!(
                primitive,
                Primitive::TextBox { lines, .. }
                    if lines.iter().flat_map(|line| &line.runs).any(|run| !run.glyphs.is_empty())
            )),
            "chart text is not shaped"
        );
    }

    /// A one-record EMF: a black triangle filling the left half of its frame.
    fn triangle_emf() -> Vec<u8> {
        let mut bytes = vec![0u8; 88];
        let mut put = |offset: usize, value: i32| {
            bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
        };
        put(0, 1);
        put(4, 88);
        put(16, 99);
        put(20, 99);
        put(32, 10_000);
        put(36, 10_000);
        put(40, 0x464D_4520);
        put(56, 4);
        put(72, 100);
        put(76, 100);
        put(80, 100);
        put(84, 100);
        let record = |kind: u32, body: Vec<u8>| {
            let mut out = kind.to_le_bytes().to_vec();
            out.extend_from_slice(&((8 + body.len()) as u32).to_le_bytes());
            out.extend(body);
            out
        };
        let i32s =
            |values: &[i32]| -> Vec<u8> { values.iter().flat_map(|v| v.to_le_bytes()).collect() };
        bytes.extend(record(39, i32s(&[1, 0, 0, 0])));
        bytes.extend(record(37, i32s(&[1])));
        let mut polygon = i32s(&[0, 0, 0, 0, 3]);
        for (x, y) in [(0i16, 0i16), (50, 0), (0, 100)] {
            polygon.extend_from_slice(&x.to_le_bytes());
            polygon.extend_from_slice(&y.to_le_bytes());
        }
        bytes.extend(record(86, polygon));
        bytes.extend(record(14, i32s(&[0, 0, 0])));
        bytes
    }

    /// `triangle_emf` with the left-top quarter of the frame selected as the
    /// clip path before the fill.
    fn clipped_triangle_emf() -> Vec<u8> {
        let mut bytes = triangle_emf();
        let record = |kind: u32, body: Vec<u8>| {
            let mut out = kind.to_le_bytes().to_vec();
            out.extend_from_slice(&((8 + body.len()) as u32).to_le_bytes());
            out.extend(body);
            out
        };
        let i32s =
            |values: &[i32]| -> Vec<u8> { values.iter().flat_map(|v| v.to_le_bytes()).collect() };
        let mut polyline = i32s(&[0, 0, 0, 0, 4]);
        for (x, y) in [(0i16, 0i16), (0, 50), (50, 50), (50, 0)] {
            polyline.extend_from_slice(&x.to_le_bytes());
            polyline.extend_from_slice(&y.to_le_bytes());
        }
        let mut clip = record(59, Vec::new());
        clip.extend(record(27, i32s(&[0, 0])));
        clip.extend(record(89, polyline));
        clip.extend(record(60, Vec::new()));
        clip.extend(record(67, i32s(&[5])));
        bytes.splice(88..88, clip);
        bytes
    }

    /// The demo deck with one extra master picture pointing at `media`.
    fn deck_with_master_picture(media: Vec<u8>) -> (PptxPackage, DeckSnapshot) {
        deck_with_cropped_master_picture(media, PictureCrop::default())
    }

    fn deck_with_cropped_master_picture(
        media: Vec<u8>,
        crop: PictureCrop,
    ) -> (PptxPackage, DeckSnapshot) {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_100).unwrap();
        let snapshot = session.snapshot().unwrap();
        package.media.push(pptx_parse::MediaPart {
            part_path: "ppt/media/logo.emf".to_owned(),
            content_type: "image/x-emf".to_owned(),
            bytes: media,
        });
        let picture = pptx_parse::Picture {
            effects: Vec::new(),
            shape_effects: None,
            base: pptx_parse::ShapeBase {
                id: 4_242,
                name: "Logo".to_owned(),
                description: None,
                hidden: false,
                placeholder: None,
                transform: ShapeTransform {
                    width: 914_400,
                    height: 914_400,
                    ..ShapeTransform::default()
                },
            },
            relationship_id: None,
            media_part_path: Some("ppt/media/logo.emf".to_owned()),
            crop,
            geometry: "rect".to_owned(),
            adjust_values: BTreeMap::new(),
            style: None,
            fill: None,
            outline: None,
        };
        package.masters[0].shapes.push(ShapeNode::Picture(picture));
        (package, snapshot)
    }

    #[test]
    fn metafile_masks_borders_and_empty_crops_preserve_picture_geometry() {
        for empty in [false, true] {
            let (mut package, snapshot) = deck_with_cropped_master_picture(
                triangle_emf(),
                PictureCrop {
                    left: if empty { 100_000 } else { 25_000 },
                    ..Default::default()
                },
            );
            let ShapeNode::Picture(picture) = package.masters[0].shapes.last_mut().unwrap() else {
                panic!()
            };
            picture.geometry = "ellipse".into();
            picture.outline = Some(ShapeOutline {
                width: Some(9525.0),
                color: Some(ColorValue {
                    rgb: Some("0000FF".into()),
                    ..Default::default()
                }),
                ..Default::default()
            });
            let mask = picture_mask(
                picture,
                PxRect {
                    x: 0.0,
                    y: 0.0,
                    w: 96.0,
                    h: 96.0,
                },
            )
            .path
            .unwrap();
            let rendered = renderer().layout_slide(&package, &snapshot, 0).unwrap();
            let shapes = rendered
                .display_list
                .primitives
                .iter()
                .filter(|p| matches!(p, Primitive::Shape { name, .. } if name == "Logo"))
                .collect::<Vec<_>>();
            assert_eq!(shapes.len(), if empty { 1 } else { 2 });
            if !empty {
                let Primitive::Shape { path, clip, .. } = shapes[0] else {
                    panic!()
                };
                assert_eq!(clip.as_ref(), Some(&mask));
                assert!(
                    matches!(path[0], GeometryPathCommand::Move { x, y: 0.0 } if (x + 1.0 / 3.0).abs() < 1e-6)
                );
            }
            let Primitive::Shape {
                path,
                fill,
                stroke: Some(stroke),
                ..
            } = shapes.last().unwrap()
            else {
                panic!()
            };
            assert_eq!(*path, mask);
            assert!(fill.is_none());
            assert_eq!(stroke.color, "#0000FF");
        }
    }

    #[test]
    fn unsupported_picture_masks_report_fallbacks_for_bitmaps_and_metafiles() {
        for media in [b"not a metafile".to_vec(), triangle_emf()] {
            let (mut package, snapshot) = deck_with_master_picture(media);
            let ShapeNode::Picture(picture) = package.masters[0].shapes.last_mut().unwrap() else {
                panic!()
            };
            picture.geometry = "unknownPreset".into();
            let rendered = renderer().layout_slide(&package, &snapshot, 0).unwrap();
            let primitives: Vec<_> = rendered
                .display_list
                .primitives
                .iter()
                .filter(|p| match p {
                    Primitive::Shape { name, .. } | Primitive::Image { name, .. } => name == "Logo",
                    _ => false,
                })
                .collect();
            assert!(!primitives.is_empty());
            for primitive in primitives {
                assert_eq!(
                    serde_json::to_value(primitive).unwrap()["geometryFallback"],
                    true
                );
            }
        }
    }

    #[test]
    fn metafile_picture_operations_share_the_slide_budget() {
        let mut bytes = triangle_emf();
        let polygon = bytes[124..164].to_vec();
        let eof = bytes.split_off(bytes.len() - 20);
        for _ in 1..2001 {
            bytes.extend_from_slice(&polygon);
        }
        bytes.extend(eof);
        let (mut package, snapshot) = deck_with_master_picture(bytes);
        let mut duplicate = package.masters[0].shapes.last().unwrap().clone();
        if let ShapeNode::Picture(picture) = &mut duplicate {
            picture.base.id += 1;
        }
        package.masters[0].shapes.push(duplicate);
        let rendered = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        let primitives = rendered.display_list.primitives;
        assert_eq!(
            primitives
                .iter()
                .filter(|p| matches!(p, Primitive::Shape { name, .. } if name == "Logo"))
                .count(),
            2001
        );
        assert_eq!(
            primitives
                .iter()
                .filter(|p| matches!(p, Primitive::Image { name, .. } if name == "Logo"))
                .count(),
            1
        );
    }

    #[test]
    fn a_metafile_picture_paints_shapes_instead_of_an_undecodable_image() {
        let (package, snapshot) = deck_with_master_picture(triangle_emf());
        let primitives = renderer()
            .layout_slide(&package, &snapshot, 0)
            .unwrap()
            .display_list
            .primitives;

        assert!(
            !primitives.iter().any(|primitive| matches!(
                primitive,
                Primitive::Image { asset_id: Some(asset), .. } if asset == "ppt/media/logo.emf"
            )),
            "the metafile is still handed to the image decoder"
        );
        let shapes: Vec<_> = primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape {
                    name,
                    path,
                    fill,
                    w,
                    h,
                    ..
                } if name == "Logo" => Some((path, fill, *w, *h)),
                _ => None,
            })
            .collect();
        assert_eq!(shapes.len(), 1, "one fill becomes one shape");
        let (path, fill, w, h) = shapes[0];
        assert_eq!(
            *fill,
            Some(Paint::Solid {
                color: "#000000".to_owned()
            })
        );
        assert_eq!((w, h), (96.0, 96.0));
        assert_eq!(
            path[1],
            ooxml_drawingml::GeometryPathCommand::Line { x: 0.5, y: 0.0 }
        );
    }

    #[test]
    fn a_metafile_clip_narrows_the_shape_clip_to_the_clipped_region() {
        let (package, snapshot) = deck_with_master_picture(clipped_triangle_emf());
        let primitives = renderer()
            .layout_slide(&package, &snapshot, 0)
            .unwrap()
            .display_list
            .primitives;

        let clip = primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Shape { name, clip, .. } if name == "Logo" => clip.clone(),
                _ => None,
            })
            .expect("the clipped metafile paints");
        assert_eq!(
            clip,
            vec![
                ooxml_drawingml::GeometryPathCommand::Move { x: 0.0, y: 0.0 },
                ooxml_drawingml::GeometryPathCommand::Line { x: 0.5, y: 0.0 },
                ooxml_drawingml::GeometryPathCommand::Line { x: 0.5, y: 0.5 },
                ooxml_drawingml::GeometryPathCommand::Line { x: 0.0, y: 0.5 },
                ooxml_drawingml::GeometryPathCommand::Close,
            ]
        );
    }

    #[test]
    fn a_negative_source_rectangle_insets_the_metafile_inside_its_box() {
        let (package, snapshot) = deck_with_cropped_master_picture(
            triangle_emf(),
            PictureCrop {
                left: -25_000,
                right: -25_000,
                ..PictureCrop::default()
            },
        );
        let primitives = renderer()
            .layout_slide(&package, &snapshot, 0)
            .unwrap()
            .display_list
            .primitives;

        let path = primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Shape { name, path, .. } if name == "Logo" => Some(path),
                _ => None,
            })
            .expect("the metafile paints");
        let GeometryPathCommand::Move { x, .. } = path[0] else {
            panic!("expected a move");
        };
        assert!((x - 1.0 / 6.0).abs() < 1e-6, "left edge at {x}");
        let GeometryPathCommand::Line { x, .. } = path[1] else {
            panic!("expected a line");
        };
        assert!((x - 0.5).abs() < 1e-6, "apex at {x}");
    }

    #[test]
    fn media_that_is_not_a_metafile_still_becomes_an_image() {
        let (package, snapshot) = deck_with_master_picture(b"\x89PNG\r\n\x1a\n".to_vec());
        let primitives = renderer()
            .layout_slide(&package, &snapshot, 0)
            .unwrap()
            .display_list
            .primitives;

        assert!(primitives.iter().any(|primitive| matches!(
            primitive,
            Primitive::Image { asset_id: Some(asset), .. } if asset == "ppt/media/logo.emf"
        )));
    }

    #[test]
    fn a_chart_title_is_shaped_into_the_centre_of_its_box() {
        let parts = chart_slide(0)
            .into_iter()
            .find_map(|primitive| match primitive {
                Primitive::Chart {
                    primitives, name, ..
                } if name == "Revenue chart" => Some(primitives),
                _ => None,
            })
            .expect("the chart frame plots");
        let (align, x, width) = parts
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::TextBox {
                    lines, paragraphs, ..
                } if lines
                    .iter()
                    .flat_map(|line| &line.runs)
                    .any(|run| run.text == "Revenue") =>
                {
                    let run = &lines[0].runs[0];
                    Some((paragraphs[0].align, run.x, run.width))
                }
                _ => None,
            })
            .expect("the chart title is laid out");
        assert_eq!(align, Some(TextAlign::Center));
        let centre = x + width / 2.0;
        assert!(
            (centre - (96.0 + 576.0 / 2.0)).abs() < 2.0,
            "the title is not centred on the frame: {centre}"
        );
    }

    #[test]
    fn a_graphic_frame_without_a_chart_part_keeps_its_placeholder() {
        assert!(chart_slide(0).iter().any(|primitive| matches!(
            primitive,
            Primitive::Placeholder { name, label, .. }
                if name == "Broken chart" && label.as_deref() == Some("Chart")
        )));
    }

    #[test]
    fn a_chart_inside_a_group_plots_in_the_group_space() {
        let primitives = chart_slide(1);
        let (label, parts) = primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::Chart {
                    label, primitives, ..
                } => Some((label, primitives)),
                _ => None,
            })
            .expect("the grouped chart plots");
        assert_eq!(label, "Untitled chart, pie chart, 1 series, 2 categories");
        assert_eq!(
            parts
                .iter()
                .filter(|primitive| matches!(primitive, Primitive::Shape { geometry, .. } if geometry == "custom"))
                .count(),
            2
        );
        assert!(parts.iter().any(|primitive| matches!(
            primitive,
            Primitive::Shape { fill: Some(Paint::Solid { color }), .. } if color == "#E7A954"
        )));
    }

    #[test]
    fn a_shared_chart_uses_the_rendering_slides_theme_key() {
        let mut package = pptx_parse::parse_pptx(CHART_FIXTURE).unwrap();
        let session = DeckSession::open(CHART_FIXTURE, 8_006).unwrap();
        let mut snapshot = session.snapshot().unwrap();
        let mut theme = package.themes[0].clone();
        theme.part_path = "ppt/theme/theme2.xml".to_owned();
        let mut master = package.masters[0].clone();
        master.part_path = "ppt/slideMasters/slideMaster2.xml".to_owned();
        master.theme_part_path = Some(theme.part_path.clone());
        master.layout_part_paths = vec!["ppt/slideLayouts/slideLayout2.xml".to_owned()];
        let mut layout = package.layouts[0].clone();
        layout.part_path = "ppt/slideLayouts/slideLayout2.xml".to_owned();
        layout.master_part_path = Some(master.part_path.clone());
        snapshot.slides[1].layout_part_path = Some(layout.part_path.clone());
        package.themes.push(theme);
        package.masters.push(master);
        package.layouts.push(layout);
        let mut chart = package
            .charts
            .iter()
            .find(|part| part.part_path == "ppt/charts/chart2.xml")
            .unwrap()
            .clone();
        chart.theme_part_path = Some("ppt/theme/theme2.xml".to_owned());
        for series in &mut chart.chart.series {
            series.color = "#ABCDEF".to_owned();
        }
        for series in chart
            .chart
            .plot_groups
            .iter_mut()
            .flat_map(|group| &mut group.series)
        {
            series.color = "#ABCDEF".to_owned();
            for point in series.points.iter_mut().flatten() {
                point.color = "#ABCDEF".to_owned();
            }
        }
        package.charts.push(chart);

        let rendered = renderer()
            .layout_slide(&package, &snapshot, 1)
            .unwrap()
            .display_list
            .primitives;

        assert!(rendered.iter().any(|primitive| matches!(
            primitive,
            Primitive::Chart { primitives, .. }
                if primitives.iter().any(|primitive| matches!(
                    primitive,
                    Primitive::Shape {
                        fill: Some(Paint::Solid { color }),
                        ..
                    } if color == "#ABCDEF"
                ))
        )));
    }

    #[test]
    fn a_deck_without_charts_plots_no_chart_primitives() {
        let package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let session = DeckSession::open(FIXTURE, 8_009).unwrap();
        let rendered = renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap();
        assert!(
            !rendered
                .display_list
                .primitives
                .iter()
                .any(|primitive| matches!(primitive, Primitive::Chart { .. }))
        );
    }

    #[test]
    fn placeholder_matching_prefers_indices_and_normalizes_common_types() {
        let indexed = Placeholder {
            placeholder_type: Some("body".to_owned()),
            index: Some(4),
            orientation: None,
            size: None,
        };
        let same_index = Placeholder {
            placeholder_type: Some("title".to_owned()),
            index: Some(4),
            orientation: None,
            size: None,
        };
        let centered_title = Placeholder {
            placeholder_type: Some("ctrTitle".to_owned()),
            index: None,
            orientation: None,
            size: None,
        };
        let title = Placeholder {
            placeholder_type: Some("title".to_owned()),
            index: None,
            orientation: None,
            size: None,
        };
        assert!(placeholders_match(&indexed, &same_index));
        assert!(placeholders_match(&centered_title, &title));

        let snapshot = ShapeSnapshot {
            id: "placeholder".to_owned(),
            source_id: 1,
            kind: ShapeKind::Shape,
            name: "Title".to_owned(),
            x: 0,
            y: 0,
            width: 0,
            height: 0,
            rotation_deg: 0.0,
            flip_h: false,
            flip_v: false,
            hidden: false,
            geometry: "rect".to_owned(),
            adjust_values: BTreeMap::new(),
            placeholder: Some(title.clone()),
            fill: None,
            resolved_fill_color: None,
            outline: None,
            resolved_outline_color: None,
            media_part_path: None,
            pending_media: None,
            blip_effects: Vec::new(),
            graphic: None,
            text_stories: Vec::new(),
            children: Vec::new(),
        };
        let layout_shape = ShapeNode::Shape(pptx_parse::Shape {
            has_preset_geometry: true,
            style: None,
            paths: Vec::new(),
            base: pptx_parse::ShapeBase {
                id: 2,
                name: "Layout title".to_owned(),
                description: None,
                hidden: false,
                placeholder: Some(title),
                transform: ShapeTransform {
                    x: 100,
                    y: 200,
                    width: 300,
                    height: 400,
                    ..ShapeTransform::default()
                },
            },
            geometry: "rect".to_owned(),
            adjust_values: BTreeMap::new(),
            fill: None,
            picture_fill: None,
            outline: None,
            effects: None,
            text: None,
        });
        let resolved = resolved_transform_value(&snapshot, None, Some(&layout_shape), None);
        assert_eq!(
            (resolved.x, resolved.y, resolved.width, resolved.height),
            (100, 200, 300, 400)
        );
    }

    #[test]
    fn a_shape_style_colour_sits_between_inherited_bodies_and_master_defaults() {
        let package = pptx_parse::parse_pptx(STYLE_FIXTURE).unwrap();
        let session = DeckSession::open(STYLE_FIXTURE, 8_010).unwrap();
        let snapshot = session.snapshot().unwrap();
        let renderer = renderer();
        let expected = [
            (
                "#EEEEEE",
                "fontRef schemeClr resolves through the deck theme",
            ),
            ("#FF0000", "fontRef srgbClr"),
            ("#00B050", "run colour beats fontRef"),
            ("#0070C0", "paragraph defRPr beats fontRef"),
            ("#0070C0", "layout placeholder colour beats fontRef"),
            ("#595959", "fontRef without a colour falls to otherStyle"),
            ("#595959", "no p:style falls to otherStyle"),
            (
                "#EEEEEE",
                "fontRef beats bodyStyle when no placeholder sets a colour",
            ),
        ];
        let mut failures = Vec::new();
        for (index, (color, case)) in expected.iter().enumerate() {
            let rendered = renderer.layout_slide(&package, &snapshot, index).unwrap();
            let colors: Vec<&str> = rendered
                .display_list
                .primitives
                .iter()
                .filter_map(|primitive| match primitive {
                    Primitive::TextBox {
                        shape_id: Some(id),
                        paragraphs,
                        ..
                    } if id.starts_with("slide:") => Some(paragraphs),
                    _ => None,
                })
                .flat_map(|paragraphs| paragraphs.iter().flat_map(|paragraph| &paragraph.runs))
                .map(|run| run.color.as_str())
                .collect();
            if colors.is_empty() || colors.iter().any(|value| value != color) {
                failures.push(format!("slide {}: {case}: {colors:?}", index + 1));
            }
        }
        assert!(failures.is_empty(), "{}", failures.join("\n"));
    }

    fn assert_style_colours(rendered: &RenderedSlide, prefix: &str, expected: &[(&str, &str)]) {
        let mut paragraphs = Vec::new();
        let mut positioned = Vec::new();
        for primitive in &rendered.display_list.primitives {
            if let Primitive::TextBox {
                shape_id: Some(id),
                paragraphs: text,
                lines,
                ..
            } = primitive
                && id.starts_with(prefix)
            {
                paragraphs.extend(text.iter().flat_map(|paragraph| {
                    paragraph
                        .runs
                        .iter()
                        .map(|run| (run.text.as_str(), run.color.as_str()))
                }));
                positioned.extend(lines.iter().flat_map(|line| {
                    line.runs
                        .iter()
                        .filter(|run| run.start != run.end)
                        .map(|run| (run.text.as_str(), run.color.as_str()))
                }));
            }
        }
        assert_eq!(paragraphs, expected);
        assert_eq!(positioned, expected);
    }

    #[test]
    fn placeholder_colours_outrank_font_refs_per_paragraph() {
        let session = DeckSession::open(STYLE_FIXTURE, 8_011).unwrap();
        let snapshot = session.snapshot().unwrap();
        let renderer = renderer();
        let layout_placeholder = renderer
            .layout_slide(session.package(), &snapshot, 4)
            .unwrap();
        assert_style_colours(
            &layout_placeholder,
            "slide:",
            &[("layout title 0070C0", "#0070C0")],
        );
        let master_placeholder = renderer
            .layout_slide(session.package(), &snapshot, 9)
            .unwrap();
        assert_style_colours(
            &master_placeholder,
            "slide:",
            &[
                ("master placeholder", "#7030A0"),
                ("paragraph default", "#0070C0"),
                ("explicit run", "#00B050"),
            ],
        );
    }

    #[test]
    fn font_ref_scheme_colours_follow_the_slides_layout_master_and_theme() {
        let session = DeckSession::open(STYLE_FIXTURE, 8_012).unwrap();
        let snapshot = session.snapshot().unwrap();
        let renderer = renderer();
        let first_theme = renderer
            .layout_slide(session.package(), &snapshot, 0)
            .unwrap();
        assert_style_colours(&first_theme, "slide:", &[("fontRef lt1", "#EEEEEE")]);
        let second_theme = renderer
            .layout_slide(session.package(), &snapshot, 8)
            .unwrap();
        for (prefix, expected) in [
            ("slide:", ("second theme", "#123456")),
            ("layout:", ("layout theme", "#234567")),
            ("master:", ("master theme", "#345678")),
        ] {
            assert_style_colours(&second_theme, prefix, &[expected]);
        }
    }

    fn end(kind: &str, width: Option<&str>, length: Option<&str>) -> LineEnd {
        LineEnd {
            end_type: kind.to_owned(),
            width: width.map(str::to_owned),
            length: length.map(str::to_owned),
        }
    }

    fn red_outline(width_emu: f64) -> ShapeOutline {
        ShapeOutline {
            width: Some(width_emu),
            color: Some(ooxml_drawingml::ColorValue {
                rgb: Some("FF0000".to_owned()),
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    #[test]
    fn line_ends_carry_their_kind_and_sizes_from_ooxml() {
        let bytes = include_bytes!("../tests/fixtures/line-ends.pptx");
        let package = pptx_parse::parse_pptx(bytes).unwrap();
        let session = DeckSession::open(bytes, 8_005).unwrap();
        let snapshot = session.snapshot().unwrap();
        let rendered = renderer().layout_slide(&package, &snapshot, 0).unwrap();
        let strokes: BTreeMap<_, _> = rendered
            .display_list
            .primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::Shape {
                    name,
                    stroke: Some(stroke),
                    ..
                } => Some((name.as_str(), stroke)),
                _ => None,
            })
            .collect();
        for kind in ["triangle", "arrow", "stealth", "diamond", "oval"] {
            for (size, width, length) in [
                ("sm-lg", 8.0, 20.0),
                ("med-sm", 12.0, 8.0),
                ("lg-med", 20.0, 12.0),
            ] {
                let name = format!("{kind}-{size}");
                let stroke = strokes[name.as_str()];
                assert_eq!(stroke.color, "#315EFB");
                assert_eq!(stroke.width, 4.0);
                assert_eq!(
                    stroke.head_end,
                    Some(StrokeEnd {
                        kind: kind.to_owned(),
                        width,
                        length,
                    }),
                    "{name} head"
                );
                assert_eq!(
                    stroke.tail_end,
                    Some(StrokeEnd {
                        kind: kind.to_owned(),
                        width: length,
                        length: width,
                    }),
                    "{name} tail"
                );
            }
        }
        let defaults = serde_json::to_string(strokes["default-medium"]).unwrap();
        assert_eq!(
            defaults,
            r##"{"color":"#315EFB","width":4.0,"headEnd":{"kind":"triangle","width":12.0,"length":12.0},"tailEnd":{"kind":"stealth","width":12.0,"length":12.0}}"##
        );
        assert_eq!(
            strokes
                .values()
                .filter(|stroke| stroke.head_end.is_some())
                .count(),
            20
        );
    }

    #[test]
    fn strokes_without_ends_serialise_as_before() {
        let theme = Theme::default();
        let plain = stroke(&red_outline(9_525.0), &theme).unwrap();
        let none = stroke(
            &ShapeOutline {
                head_end: Some(end("none", Some("lg"), None)),
                tail_end: Some(end("none", None, None)),
                ..red_outline(9_525.0)
            },
            &theme,
        )
        .unwrap();
        let expected = r##"{"color":"#FF0000","width":1.0}"##;
        assert_eq!(serde_json::from_str::<Stroke>(expected).unwrap(), plain);
        assert_eq!(serde_json::to_string(&plain).unwrap(), expected);
        assert_eq!(serde_json::to_string(&none).unwrap(), expected);
    }

    #[test]
    fn thin_lines_keep_a_visible_end() {
        let end = line_end(Some(&end("oval", Some("sm"), Some("lg"))), 1.0).unwrap();
        assert!((end.width - 5.291_339).abs() < 1e-6);
        assert!((end.length - 13.228_347).abs() < 1e-6);
    }

    #[test]
    fn a_family_we_cannot_ship_lends_its_own_metrics_to_the_substitute() {
        let mut renderer = SlideRenderer::new();
        renderer
            .register_font("Trebuchet MS", false, false, FONT)
            .unwrap();
        renderer.register_font("Arial", false, false, FONT).unwrap();
        let substituted = renderer.resolve_face("Trebuchet MS", false, false).unwrap();
        let plain = renderer.resolve_face("Arial", false, false).unwrap();
        assert!(plain.widths.is_none() && plain.line.is_none());

        let metrics = substituted.widths.expect("trebuchet ms widths");
        assert!((family_advance(metrics, 'M').unwrap() - 0.709).abs() < 1e-6);
        assert!((family_advance(metrics, ' ').unwrap() - 0.301).abs() < 1e-6);
        assert!((family_advance(metrics, '\u{2019}').unwrap() - 0.367).abs() < 1e-6);
        assert_eq!(family_advance(metrics, '\u{4e00}'), None);
        assert!(family_metrics("trebuchet ms", true, false).is_some());
        assert!(family_metrics("Trebuchet MS", false, false).is_none());
        assert!(family_metrics("calibri", false, false).is_none());

        let line = family_line_box(substituted.line.expect("trebuchet ms lines"), 1000.0);
        assert!((line.ascent - 939.0).abs() < 1e-3);
        assert!((line.descent - 222.0).abs() < 1e-3);
    }

    #[test]
    fn missing_faces_keep_the_requested_family_and_style_metrics() {
        for registered in ["Arial", "Trebuchet MS"] {
            let mut renderer = SlideRenderer::new();
            renderer
                .register_font(registered, false, false, FONT)
                .unwrap();
            for (bold, italic) in [(false, false), (true, false), (false, true), (true, true)] {
                let face = renderer.resolve_face("Trebuchet MS", bold, italic).unwrap();
                let expected = family_metrics("trebuchet ms", bold, italic).unwrap();
                assert_eq!(renderer.fonts.font_bytes(face.id).unwrap(), FONT);
                assert!(std::ptr::eq(face.widths.unwrap(), expected));
                assert!(std::ptr::eq(face.line.unwrap(), expected));
            }
            let unknown = renderer
                .resolve_face("Unknown family", false, false)
                .unwrap();
            assert!(unknown.widths.is_none() && unknown.line.is_none());
        }
    }

    #[test]
    fn different_requested_metrics_do_not_share_cached_fallback_layouts() {
        let session = DeckSession::open(
            include_bytes!("../tests/fixtures/paragraph-spacing.pptx"),
            8_020,
        )
        .unwrap();
        let new_renderer = || {
            let mut renderer = SlideRenderer::new();
            renderer.register_font("Arial", false, false, FONT).unwrap();
            renderer
        };
        let render = |renderer: &SlideRenderer, family: &str| {
            let mut snapshot = session.snapshot().unwrap();
            let story = &mut snapshot.slides[0]
                .shapes
                .iter_mut()
                .find(|shape| shape.source_id == 2)
                .unwrap()
                .text_stories[0];
            for paragraph in &mut story.paragraphs {
                for run in &mut paragraph.runs {
                    run.style.font_family = Some(family.to_owned());
                    run.style.font_size_pt = Some(24.0);
                    run.text = "MMMMMMMM".to_owned();
                }
            }
            renderer
                .layout_slide(session.package(), &snapshot, 0)
                .unwrap()
                .display_list
                .primitives
                .into_iter()
                .find_map(|primitive| match primitive {
                    Primitive::TextBox {
                        object_id: 2,
                        lines,
                        ..
                    } => Some(lines),
                    _ => None,
                })
                .unwrap()
        };
        let renderer = new_renderer();
        let trebuchet = render(&renderer, "Trebuchet MS");
        let consolas = render(&renderer, "Consolas");
        assert!((trebuchet[0].width - consolas[0].width).abs() > 1.0);
        assert_eq!(consolas, render(&new_renderer(), "Consolas"));
        assert_eq!(trebuchet, render(&renderer, "Trebuchet MS"));
    }
}
