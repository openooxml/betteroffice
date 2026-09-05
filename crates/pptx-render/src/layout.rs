use std::collections::{BTreeMap, HashMap, HashSet};

use ooxml_drawingml::chart::PlotRect;
use ooxml_drawingml::{
    ColorValue, LineEnd, ShapeFill, ShapeOutline, ShapeStyle, Theme, ThemeFormatScheme,
    preset_geometry_to_path, resolve_color_value_to_hex_with_theme,
    resolve_color_value_to_rgba_hex, resolve_theme_font_ref, style_fill, style_outline,
};
use ooxml_text::{CompatFlags, FontId, FontStore, break_opportunities, shape, single_line_box};
use pptx_edit::{DeckSnapshot, ShapeKind, ShapeSnapshot, StorySnapshot, TextStyle};
use pptx_parse::{
    ChartSpace, CustomGeometryPath, GraphicFrameData, ParagraphProperties, Picture, PictureCrop,
    Placeholder, PptxPackage, RunProperties, ShapeNode, ShapeTransform, Slide, SlideLayout,
    SlideMaster, TextAutofit, TextBody,
};
use serde::{Deserialize, Serialize};
use thiserror::Error;

use crate::chart::{ChartFrame, ChartText, chart_primitive};
use crate::{
    CONTRACT_VERSION, CaretStop, GradientStop, GradientType, ImageCrop, Paint, PositionedGlyph,
    PositionedTextLine, PositionedTextRun, Primitive, Stroke, StrokeEnd, SurfaceDisplayList,
    TextAlign, TextAnchor, TextParagraph, TextRun, Transform,
};

const EMU_PER_CSS_PIXEL: f32 = 9_525.0;
const LINE_END_MIN_BASE_PX: f32 = 0.7 / 25.4 * 96.0;
const DEFAULT_INSET_HORIZONTAL_EMU: i64 = 91_440;
const DEFAULT_INSET_VERTICAL_EMU: i64 = 45_720;
const DEFAULT_FONT_SIZE_PT: f32 = 18.0;
const MIN_AUTOFIT_SCALE: f32 = 0.5;
const MAX_FONT_BYTES: usize = 32 * 1024 * 1024;
const MAX_FONTS: usize = 256;
const MAX_RENDER_SHAPES: usize = 20_000;
const MAX_TEXT_BYTES: usize = 4 * 1024 * 1024;
const MAX_TEXT_LINES: usize = 100_000;
const MAX_TEXT_PARAGRAPHS: usize = 20_000;
const MAX_TEXT_RUNS: usize = 100_000;
/// Chart parts one slide may draw, shared across its charts.
pub(crate) const MAX_CHART_PRIMITIVES: usize = 100_000;

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
}

pub struct SlideRenderer {
    fonts: FontStore,
    faces: HashMap<(String, bool, bool), FontFace>,
    fallback: Option<FontFace>,
    /// Normalized fallback family.
    fallback_family: Option<String>,
    font_count: usize,
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
        let face = FontFace {
            id,
            family: family.to_owned(),
        };
        self.faces
            .insert((normalize_family(family), bold, italic), face.clone());
        self.fallback.get_or_insert(face);
        self.fallback_family
            .get_or_insert_with(|| normalize_family(family));
        self.font_count += 1;
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
        let theme = theme_part.map(|part| &part.theme).unwrap_or(&default_theme);
        let default_format_scheme = ThemeFormatScheme::default();
        let format_scheme = theme_part
            .map(|part| &part.format_scheme)
            .unwrap_or(&default_format_scheme);
        let background = parsed_slide
            .and_then(|slide| slide.background.as_ref())
            .or_else(|| layout.and_then(|layout| layout.background.as_ref()))
            .or_else(|| master.and_then(|master| master.background.as_ref()))
            .and_then(|fill| paint(fill, theme))
            .or_else(|| {
                Some(Paint::Solid {
                    color: "#ffffff".to_owned(),
                })
            });
        let width = emu_to_px(deck.width_emu);
        let height = emu_to_px(deck.height_emu);
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
            slide_number: i64::from(package.presentation.first_slide_num) + slide_index as i64,
        };
        let root_space = Space::root();
        let show_master = parsed_slide.is_none_or(|slide| slide.show_master_shapes)
            && layout.is_none_or(|layout| layout.show_master_shapes);
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
        if let Some(layout) = layout {
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
        for (bold, italic) in styles {
            if let Some(face) = self.faces.get(&(requested.clone(), bold, italic)) {
                return Ok(face.clone());
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
            .map(|(_, face)| face.clone())
            .ok_or(RenderError::NoFont)
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
            let (x, y) = region.local_point(x, y);
            if !region.rect.contains(x, y) {
                continue;
            }
            if let Some(text) = &region.text
                && let Some(line) = nearest_line(&text.lines, y)
                && let Some(caret) = line
                    .caret_stops
                    .iter()
                    .min_by(|left, right| (left.x - x).abs().total_cmp(&(right.x - x).abs()))
            {
                return Some(HitTestResult::Text {
                    shape_id: region.shape_id.clone(),
                    story_id: text.story_id.clone(),
                    position: caret.position,
                });
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
        let fill = shape
            .fill
            .as_ref()
            .or(inherited_fill.as_ref())
            .and_then(|fill| paint(fill, self.theme));
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
        let transform = Transform {
            rotation_deg: shape.rotation_deg as f32,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
        };
        match shape.kind {
            ShapeKind::Shape => {
                self.push_shape(
                    Primitive::Shape {
                        object_id: shape.source_id,
                        shape_id: Some(stable_id.clone()),
                        name: shape.name.clone(),
                        x: rect.x,
                        y: rect.y,
                        w: rect.w,
                        h: rect.h,
                        geometry: shape.geometry.clone(),
                        path: geometry_path(
                            &shape.geometry,
                            &shape.adjust_values,
                            f64::from(rect.w) / f64::from(rect.h),
                        ),
                        adjust_values: shape
                            .adjust_values
                            .iter()
                            .map(|(name, value)| (name.clone(), *value as f32))
                            .collect(),
                        fill,
                        stroke: outline,
                        transform,
                    },
                    custom_paths(original),
                )?;
            }
            ShapeKind::Picture => {
                let source = picture_source(original);
                self.primitives.push(Primitive::Image {
                    object_id: shape.source_id,
                    shape_id: Some(stable_id.clone()),
                    name: shape.name.clone(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    asset_id: shape.media_part_path.clone(),
                    crop: source
                        .map(|value| image_crop(&value.crop))
                        .unwrap_or_default(),
                    path: source.and_then(|value| picture_mask(value, rect)),
                    stroke: outline,
                    transform,
                });
            }
            ShapeKind::GraphicFrame => {
                self.render_graphic_frame(
                    shape.source_id,
                    &stable_id,
                    &shape.name,
                    rect,
                    transform,
                    shape.graphic.as_ref(),
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
        let text = shape.text_stories.first().map(content_from_story);
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
                self.push_shape(
                    Primitive::Shape {
                        object_id: base.id,
                        shape_id: Some(stable_id.to_owned()),
                        name: base.name.clone(),
                        x: rect.x,
                        y: rect.y,
                        w: rect.w,
                        h: rect.h,
                        geometry: value.geometry.clone(),
                        path: geometry_path(
                            &value.geometry,
                            &value.adjust_values,
                            f64::from(rect.w) / f64::from(rect.h),
                        ),
                        adjust_values: value
                            .adjust_values
                            .iter()
                            .map(|(name, value)| (name.clone(), *value as f32))
                            .collect(),
                        fill: self
                            .resolved_fill(&[Some(shape)])
                            .and_then(|fill| paint(&fill, self.theme)),
                        stroke: self
                            .resolved_outline(&[Some(shape)])
                            .and_then(|outline| stroke(&outline, self.theme)),
                        transform,
                    },
                    &value.paths,
                )?;
            }
            ShapeNode::Picture(value) => {
                self.primitives.push(Primitive::Image {
                    object_id: base.id,
                    shape_id: Some(stable_id.to_owned()),
                    name: base.name.clone(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    asset_id: value.media_part_path.clone(),
                    crop: image_crop(&value.crop),
                    path: picture_mask(value, rect),
                    stroke: self
                        .resolved_outline(&[Some(shape)])
                        .and_then(|outline| stroke(&outline, self.theme)),
                    transform,
                });
            }
            ShapeNode::GraphicFrame(value) => {
                self.render_graphic_frame(
                    base.id,
                    stable_id,
                    &base.name,
                    rect,
                    transform,
                    Some(&value.data),
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
            transform,
            text: text_hit,
        });
        Ok(())
    }

    fn push_shape(
        &mut self,
        primitive: Primitive,
        paths: &[CustomGeometryPath],
    ) -> Result<(), RenderError> {
        if paths.is_empty()
            || !matches!(&primitive, Primitive::Shape { geometry, .. } if geometry == "custom")
        {
            self.primitives.push(primitive);
            return Ok(());
        }
        for (index, custom) in paths.iter().enumerate() {
            if index > 0 {
                self.charge_shape()?;
            }
            let mut primitive = primitive.clone();
            if let Primitive::Shape {
                path, fill, stroke, ..
            } = &mut primitive
            {
                *path = custom.commands.clone();
                if custom.no_fill {
                    *fill = None;
                }
                if custom.no_stroke {
                    *stroke = None;
                }
            }
            self.primitives.push(primitive);
        }
        Ok(())
    }

    /// Plots a chart frame, or keeps the placeholder for graphics that carry
    /// no drawable data.
    fn render_graphic_frame(
        &mut self,
        object_id: u32,
        shape_id: &str,
        name: &str,
        rect: PxRect,
        transform: Transform,
        graphic: Option<&GraphicFrameData>,
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
            let chart = chart_primitive(frame, space, self.chart_budget, &mut |text| {
                chart_text_primitive(renderer, theme, shape_id, text)
            })?;
            if let Primitive::Chart { primitives, .. } = &chart {
                self.chart_budget -= primitives.len();
            }
            self.primitives.push(chart);
            return Ok(());
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

    fn chart_space(&self, graphic: Option<&GraphicFrameData>) -> Option<&'a ChartSpace> {
        let GraphicFrameData::Chart {
            part_path: Some(part_path),
            ..
        } = graphic?
        else {
            return None;
        };
        self.package
            .charts
            .iter()
            .find(|part| {
                &part.part_path == part_path
                    && part.theme_part_path.as_deref() == self.theme_part_path
            })
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
        let left = cascade.inset_left().unwrap_or(DEFAULT_INSET_HORIZONTAL_EMU);
        let right = cascade
            .inset_right()
            .unwrap_or(DEFAULT_INSET_HORIZONTAL_EMU);
        let top = cascade.inset_top().unwrap_or(DEFAULT_INSET_VERTICAL_EMU);
        let bottom = cascade.inset_bottom().unwrap_or(DEFAULT_INSET_VERTICAL_EMU);
        let content_rect = PxRect {
            x: rect.x + emu_to_px(left),
            y: rect.y + emu_to_px(top),
            w: (rect.w - emu_to_px(left + right)).max(1.0),
            h: (rect.h - emu_to_px(top + bottom)).max(1.0),
        };
        let autofit = cascade.autofit();
        let mut scale = match autofit {
            Some(TextAutofit::Normal { font_scale, .. }) => {
                font_scale.unwrap_or(1.0).clamp(0.1, 1.0) as f32
            }
            _ => 1.0,
        };
        let mut laid_out = layout_content(&self.renderer.fonts, &resolved, content_rect, scale)?;
        if matches!(
            autofit,
            Some(TextAutofit::Normal { .. } | TextAutofit::Shape)
        ) {
            while laid_out.total_height > content_rect.h && scale > MIN_AUTOFIT_SCALE {
                scale = (scale * 0.9).max(MIN_AUTOFIT_SCALE);
                laid_out = layout_content(&self.renderer.fonts, &resolved, content_rect, scale)?;
                if scale == MIN_AUTOFIT_SCALE {
                    break;
                }
            }
        }
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
        let vertical_shift = match anchor {
            TextAnchor::Top => 0.0,
            TextAnchor::Center => ((content_rect.h - laid_out.total_height) / 2.0).max(0.0),
            TextAnchor::Bottom => (content_rect.h - laid_out.total_height).max(0.0),
        };
        for line in &mut laid_out.lines {
            shift_line(line, 0.0, vertical_shift);
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
                        font_size_pt: run.style.font_size_pt * scale,
                        bold: run.style.bold,
                        italic: run.style.italic,
                        underline: run.style.underline,
                        color: run.style.color.clone(),
                    })
                    .collect(),
            })
            .collect();
        let overflow = laid_out.total_height > content_rect.h;
        let story_id = content.story_id;
        let lines = laid_out.lines;
        self.primitives.push(Primitive::TextBox {
            object_id,
            shape_id: Some(shape_id.to_owned()),
            story_id: Some(story_id.clone()),
            x: rect.x,
            y: rect.y,
            w: rect.w,
            h: rect.h,
            anchor,
            paragraphs: display_paragraphs,
            lines: lines.clone(),
            overflow,
            transform,
        });
        Ok(TextHit { story_id, lines })
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
    fn anchor(&self) -> Option<&str> {
        self.primary
            .and_then(|body| body.anchor.as_deref())
            .or_else(|| self.layout.and_then(|body| body.anchor.as_deref()))
            .or_else(|| self.master.and_then(|body| body.anchor.as_deref()))
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
            if let Some(source) = body
                .paragraphs
                .get(index)
                .or_else(|| body.paragraphs.get(level as usize))
                .map(|paragraph| &paragraph.properties)
            {
                merge_paragraph_properties(&mut properties, source);
            }
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
    bold: bool,
    italic: bool,
    underline: bool,
    color: String,
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
    let mut story_offset = 0_u32;
    let mut paragraphs = Vec::with_capacity(content.paragraphs.len());
    for (index, paragraph) in content.paragraphs.iter().enumerate() {
        let properties = cascade.paragraph_properties(index, paragraph.level);
        let mut runs = Vec::with_capacity(paragraph.runs.len().max(1));
        for run in &paragraph.runs {
            let style =
                resolve_style(renderer, theme, &run.style, properties.default_run.as_ref())?;
            let start = story_offset;
            story_offset = story_offset.saturating_add(utf16_len(&run.text));
            runs.push(ResolvedRun {
                text: run.text.clone(),
                start,
                style,
            });
        }
        if runs.is_empty() {
            runs.push(ResolvedRun {
                text: String::new(),
                start: story_offset,
                style: resolve_style(
                    renderer,
                    theme,
                    &TextStyle::default(),
                    properties.default_run.as_ref(),
                )?,
            });
        }
        let alignment = paragraph
            .alignment
            .as_deref()
            .or(properties.alignment.as_deref());
        paragraphs.push(ResolvedParagraph {
            align: parse_align(alignment),
            justify: is_full_justification(alignment),
            level: paragraph.level,
            margin_left_px: emu_to_px(properties.margin_left.unwrap_or_default()),
            runs,
        });
        story_offset = story_offset.saturating_add(1);
    }
    Ok(ResolvedContent { paragraphs })
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
    Ok(ResolvedStyle {
        family: face.family.clone(),
        face,
        font_size_pt,
        bold,
        italic,
        underline: direct
            .underline
            .as_deref()
            .or_else(|| fallback.and_then(|value| value.underline.as_deref()))
            .is_some_and(|value| value != "none"),
        color,
    })
}

/// One shaped line of chart text, in the deck's minor font at the weight and
/// pixel size the plot geometry asked for.
fn chart_text_primitive(
    renderer: &SlideRenderer,
    theme: &Theme,
    shape_id: &str,
    text: ChartText<'_>,
) -> Result<Primitive, RenderError> {
    let bold = text.font.weight >= 600;
    let family = resolve_theme_font_ref(Some(theme), "+mn-lt");
    let face = renderer.resolve_face(&family, bold, false)?;
    let size_px = safe_geometry(text.font.size_px as f32).clamp(1.0, 4_096.0);
    let shaped = shape(&renderer.fonts, face.id, text.text, size_px, &[])
        .map_err(|error| RenderError::Font(error.to_string()))?;
    let metrics = renderer
        .fonts
        .metrics(face.id)
        .map_err(|error| RenderError::Font(error.to_string()))?;
    let line_box = single_line_box(metrics, size_px, &CompatFlags::default());
    let x = safe_geometry(text.x as f32);
    let baseline = safe_geometry(text.baseline_y as f32);
    let mut glyphs = Vec::with_capacity(shaped.len());
    let mut cursor = 0.0_f32;
    for glyph in &shaped {
        glyphs.push(PositionedGlyph {
            glyph_id: glyph.glyph_id,
            cluster: glyph.cluster,
            x: x + cursor,
            advance: glyph.x_advance,
            x_offset: glyph.x_offset,
            y_offset: baseline + glyph.y_offset,
        });
        cursor += glyph.x_advance;
    }
    let run = PositionedTextRun {
        text: text.text.to_owned(),
        start: 0,
        end: utf16_len(text.text),
        x,
        width: cursor.max(0.0),
        font_id: face.id.to_u32(),
        font_family: face.family.clone(),
        font_size_px: size_px,
        bold,
        italic: false,
        underline: false,
        color: text.color.to_owned(),
        glyphs,
    };
    let width = run.width;
    Ok(Primitive::TextBox {
        object_id: text.object_id,
        shape_id: Some(shape_id.to_owned()),
        story_id: None,
        x,
        y: baseline - line_box.ascent,
        w: safe_geometry(text.width as f32).max(width),
        h: line_box.height(),
        anchor: TextAnchor::Top,
        paragraphs: vec![TextParagraph {
            align: Some(TextAlign::Left),
            level: 0,
            runs: vec![TextRun {
                text: text.text.to_owned(),
                font_family: face.family.clone(),
                font_size_pt: size_px * 72.0 / 96.0,
                bold,
                italic: false,
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

fn layout_content(
    fonts: &FontStore,
    content: &ResolvedContent,
    rect: PxRect,
    scale: f32,
) -> Result<LayoutText, RenderError> {
    let mut lines = Vec::new();
    let mut y = rect.y;
    for paragraph in &content.paragraphs {
        let paragraph_x = rect.x + paragraph.margin_left_px.max(0.0);
        let paragraph_width = (rect.w - paragraph.margin_left_px.max(0.0)).max(1.0);
        let mut paragraph_lines =
            layout_paragraph(fonts, paragraph, paragraph_x, y, paragraph_width, scale)?;
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

fn layout_paragraph(
    fonts: &FontStore,
    paragraph: &ResolvedParagraph,
    x: f32,
    y: f32,
    width: f32,
    scale: f32,
) -> Result<Vec<PositionedTextLine>, RenderError> {
    let clusters = shape_paragraph(fonts, paragraph, scale)?;
    if clusters.is_empty() {
        let style = &paragraph.runs[0].style;
        let line_box = style_line_box(fonts, style, scale)?;
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
    let ranges = wrap_clusters(&clusters, width);
    let line_count = ranges.len();
    let mut output = Vec::with_capacity(line_count);
    let mut line_y = y;
    for (line_index, (start, end)) in ranges.into_iter().enumerate() {
        let slice = &clusters[start..end];
        let natural_width = slice.iter().map(|cluster| cluster.width).sum::<f32>();
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
        let line_width = advances.iter().sum::<f32>();
        let line_x = match paragraph.align {
            TextAlign::Center => x + ((width - natural_width) / 2.0).max(0.0),
            TextAlign::Right => x + (width - natural_width).max(0.0),
            TextAlign::Left | TextAlign::Justify => x,
        };
        let line_box = clusters_line_box(fonts, slice, scale)?;
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
    Ok(output)
}

struct ShapedCluster {
    text: String,
    start: u32,
    end: u32,
    width: f32,
    run_index: usize,
    style: ResolvedStyle,
    glyphs: Vec<ClusterGlyph>,
    break_after: bool,
    mandatory: bool,
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
            if character != '\n' {
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
            clusters.push(ShapedCluster {
                text: "\n".to_owned(),
                start,
                end: start + 1,
                width: 0.0,
                run_index,
                style: run.style.clone(),
                glyphs: Vec::new(),
                break_after: true,
                mandatory: true,
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
    let size_px = points_to_px(run.style.font_size_pt * scale);
    let shaped = shape(fonts, run.style.face.id, text, size_px, &[])
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
        let global_end = global_run_byte + run_byte_start + end_byte;
        output.push(ShapedCluster {
            text: text[start_byte..end_byte].to_owned(),
            start: source_start,
            end: source_end,
            width: glyph_x.max(0.0),
            run_index,
            style: run.style.clone(),
            glyphs,
            break_after: breaks.contains_key(&global_end),
            mandatory: breaks.get(&global_end).copied().unwrap_or(false),
        });
    }
    Ok(())
}

fn wrap_clusters(clusters: &[ShapedCluster], width: f32) -> Vec<(usize, usize)> {
    let mut ranges = Vec::new();
    let mut start = 0;
    while start < clusters.len() {
        let mut cursor = start;
        let mut line_width = 0.0;
        let mut last_break = None;
        let mut end = clusters.len();
        while cursor < clusters.len() {
            let cluster = &clusters[cursor];
            if line_width + cluster.width > width && cursor > start {
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
    for (index, cluster) in clusters.iter().enumerate() {
        if cluster.text == "\n" {
            continue;
        }
        let append = output.last().is_some_and(|run| {
            run.end == cluster.start && run.font_id == cluster.style.face.id.to_u32()
        });
        if !append {
            output.push(PositionedTextRun {
                text: String::new(),
                start: cluster.start,
                end: cluster.start,
                x: cursor_x,
                width: 0.0,
                font_id: cluster.style.face.id.to_u32(),
                font_family: cluster.style.family.clone(),
                font_size_px: points_to_px(cluster.style.font_size_pt * scale),
                bold: cluster.style.bold,
                italic: cluster.style.italic,
                underline: cluster.style.underline,
                color: cluster.style.color.clone(),
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
                y_offset: baseline + glyph.y_offset,
            });
        }
        let advance = advances.get(index).copied().unwrap_or(cluster.width);
        run.width += advance;
        cursor_x += advance;
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
    clusters
        .iter()
        .map(|cluster| JustifyCluster {
            width: cluster.width,
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
) -> Result<ooxml_text::LineBox, RenderError> {
    let mut ascent: f32 = 0.0;
    let mut descent: f32 = 0.0;
    let mut leading: f32 = 0.0;
    let mut seen = HashSet::new();
    for cluster in clusters {
        if !seen.insert(cluster.run_index) {
            continue;
        }
        let line = style_line_box(fonts, &cluster.style, scale)?;
        ascent = ascent.max(line.ascent);
        descent = descent.max(line.descent);
        leading = leading.max(line.leading);
    }
    Ok(ooxml_text::LineBox {
        ascent,
        descent,
        leading,
    })
}

fn style_line_box(
    fonts: &FontStore,
    style: &ResolvedStyle,
    scale: f32,
) -> Result<ooxml_text::LineBox, RenderError> {
    let metrics = fonts
        .metrics(style.face.id)
        .map_err(|error| RenderError::Font(error.to_string()))?;
    Ok(single_line_box(
        metrics,
        points_to_px(style.font_size_pt * scale),
        &CompatFlags::default(),
    ))
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

#[derive(Clone, Copy)]
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
    transform: Transform,
    text: Option<TextHit>,
}

impl HitRegion {
    /// Undoes the rotate-then-flip the shape paints with, so `rect` and the text
    /// lines can both be read in their unrotated frame.
    fn local_point(&self, x: f32, y: f32) -> (f32, f32) {
        if self.transform.is_identity() {
            return (x, y);
        }
        let center_x = self.rect.x + self.rect.w / 2.0;
        let center_y = self.rect.y + self.rect.h / 2.0;
        let (sin, cos) = self.transform.rotation_deg.to_radians().sin_cos();
        let dx = x - center_x;
        let dy = y - center_y;
        let mut local_x = dx * cos + dy * sin;
        let mut local_y = dy * cos - dx * sin;
        if self.transform.flip_h {
            local_x = -local_x;
        }
        if self.transform.flip_v {
            local_y = -local_y;
        }
        (center_x + local_x, center_y + local_y)
    }
}

struct TextHit {
    story_id: String,
    lines: Vec<PositionedTextLine>,
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
        color: direct.color.clone().or_else(|| fallback.color.clone()),
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
    if source.indent.is_some() {
        target.indent = source.indent;
    }
    if source.bullet.is_some() {
        target.bullet.clone_from(&source.bullet);
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
}

fn style_from_properties(properties: &RunProperties, theme: &Theme) -> TextStyle {
    TextStyle {
        bold: properties.bold,
        italic: properties.italic,
        font_size_pt: properties.font_size_pt,
        color: resolve_color_value_to_hex_with_theme(properties.color.as_ref(), Some(theme)),
        font_family: properties.font_family.clone(),
        underline: properties.underline.clone(),
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
    if let Some(gradient) = &fill.gradient {
        let gradient_type = match gradient.gradient_type.as_str() {
            "radial" => GradientType::Radial,
            "rectangular" => GradientType::Rectangular,
            "path" => GradientType::Path,
            _ => GradientType::Linear,
        };
        let stops = gradient
            .stops
            .iter()
            .filter_map(|stop| {
                Some(GradientStop {
                    position: (stop.position as f32 / 100_000.0).clamp(0.0, 1.0),
                    color: resolve_paint_color(Some(&stop.color), theme)?,
                })
            })
            .collect::<Vec<_>>();
        if !stops.is_empty() {
            return Some(Paint::Gradient {
                gradient_type,
                angle_deg: gradient.angle.map(|value| value as f32),
                stops,
            });
        }
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

fn stroke(outline: &ShapeOutline, theme: &Theme) -> Option<Stroke> {
    let color = resolve_paint_color(outline.color.as_ref(), theme)?;
    let width = outline
        .width
        .filter(|width| width.is_finite() && *width >= 0.0)
        .map(|width| width as f32 / EMU_PER_CSS_PIXEL)
        .unwrap_or(1.0);
    Some(Stroke {
        color,
        width,
        dashed: outline
            .style
            .as_deref()
            .is_some_and(|style| style != "solid"),
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

/// Resolves a picture's nonrectangular preset mask.
fn picture_mask(
    picture: &Picture,
    rect: PxRect,
) -> Option<Vec<ooxml_drawingml::GeometryPathCommand>> {
    if picture.geometry.is_empty() || picture.geometry == "rect" || rect.h <= 0.0 {
        return None;
    }
    let path = geometry_path(
        &picture.geometry,
        &picture.adjust_values,
        f64::from(rect.w) / f64::from(rect.h),
    );
    (!path.is_empty()).then_some(path)
}

fn custom_paths(shape: Option<&ShapeNode>) -> &[CustomGeometryPath] {
    match shape {
        Some(ShapeNode::Shape(shape)) => &shape.paths,
        _ => &[],
    }
}

fn geometry_path(
    geometry: &str,
    adjustments: &BTreeMap<String, f64>,
    aspect_ratio: f64,
) -> Vec<ooxml_drawingml::GeometryPathCommand> {
    let adjustments = adjustments
        .iter()
        .map(|(name, value)| (name.clone(), *value))
        .collect();
    preset_geometry_to_path(geometry, &adjustments, aspect_ratio)
        .or_else(|| preset_geometry_to_path("rect", &HashMap::new(), aspect_ratio))
        .unwrap_or_default()
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
    use std::collections::BTreeSet;

    use pptx_edit::{DeckSession, EditCtx};
    use pptx_parse::ShapeBase;

    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    const CHART_FIXTURE: &[u8] = include_bytes!("../../pptx-parse/tests/fixtures/chart-deck.pptx");
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
            crop: PictureCrop::default(),
            geometry: "rect".to_owned(),
            adjust_values: BTreeMap::new(),
            fill: None,
            outline: None,
            style: None,
        };
        let rect = PxRect {
            x: 0.0,
            y: 0.0,
            w: 100.0,
            h: 50.0,
        };
        assert!(picture_mask(&picture, rect).is_none());
        picture.geometry = "ellipse".to_owned();
        let path = picture_mask(&picture, rect).unwrap();
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

    fn renderer() -> SlideRenderer {
        let mut renderer = SlideRenderer::new();
        for bold in [false, true] {
            renderer.register_font("Arial", bold, false, FONT).unwrap();
        }
        renderer
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
            runs: vec![ResolvedRun {
                text: text.to_owned(),
                start: 0,
                style: ResolvedStyle {
                    face,
                    family: "Arial".to_owned(),
                    font_size_pt: 18.0,
                    bold: false,
                    italic: false,
                    underline: false,
                    color: "#000000".to_owned(),
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
        let text = "alpha beta \t  gamma delta";
        let justified = paragraph(&renderer, "just", text);
        let clusters = shape_paragraph(&renderer.fonts, &justified, 1.0).unwrap();
        let second_break = clusters
            .iter()
            .enumerate()
            .find(|(_, cluster)| cluster.end == utf16_len("alpha beta \t  "))
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
        let lines = layout_paragraph(&renderer.fonts, &justified, 20.0, 30.0, width, 1.0).unwrap();
        let natural = layout_paragraph(
            &renderer.fonts,
            &paragraph(&renderer, "justLow", text),
            20.0,
            30.0,
            width,
            1.0,
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
                    Primitive::Chart { primitives, .. } => normalize_font_ids(primitives, bold_id),
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
            | Primitive::Chart { shape_id, .. } => shape_id.as_deref(),
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
    fn hit_testing_flipped_text_reads_the_mirrored_caret() {
        // membership cannot catch a flip — the rect maps onto itself — so pin it
        // on the caret, where the sign of the local point decides the answer
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
                transform: Transform {
                    rotation_deg: 0.0,
                    flip_h,
                    flip_v: false,
                },
                text: Some(TextHit {
                    story_id: "story".to_owned(),
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
        assert_eq!(position(true), 5);
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

    #[test]
    fn normal_autofit_scales_text_until_the_shape_height_is_respected() {
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
        let rendered = renderer()
            .layout_slide(&package, &session.snapshot().unwrap(), 0)
            .unwrap();
        let font_size = rendered
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
            .unwrap();
        assert!(font_size < 40.0);
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
            graphic: None,
            text_stories: Vec::new(),
            children: Vec::new(),
        };
        let layout_shape = ShapeNode::Shape(pptx_parse::Shape {
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
            outline: None,
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
}
