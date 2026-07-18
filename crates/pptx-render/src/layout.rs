use std::collections::{BTreeMap, HashMap, HashSet};

use ooxml_drawingml::{
    ShapeFill, ShapeOutline, Theme, preset_geometry_to_path, resolve_color_value_to_hex_with_theme,
    resolve_theme_font_ref,
};
use ooxml_text::{CompatFlags, FontId, FontStore, break_opportunities, shape, single_line_box};
use pptx_edit::{DeckSnapshot, ShapeKind, ShapeSnapshot, StorySnapshot, TextStyle};
use pptx_parse::{
    GraphicFrameData, ParagraphProperties, Placeholder, PptxPackage, RunProperties, ShapeNode,
    ShapeTransform, Slide, SlideLayout, SlideMaster, TextAutofit, TextBody,
};
use serde::{Deserialize, Serialize};
use thiserror::Error;

use crate::{
    CONTRACT_VERSION, CaretStop, GradientStop, GradientType, Paint, PositionedGlyph,
    PositionedTextLine, PositionedTextRun, Primitive, Stroke, SurfaceDisplayList, TextAlign,
    TextAnchor, TextParagraph, TextRun, Transform,
};

const EMU_PER_CSS_PIXEL: f32 = 9_525.0;
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
        self.font_count += 1;
        Ok(id.to_u32())
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
        let theme = master
            .and_then(|master| master.theme_part_path.as_deref())
            .and_then(|path| package.themes.iter().find(|theme| theme.part_path == path))
            .map(|part| &part.theme)
            .or_else(|| package.themes.first().map(|part| &part.theme));
        let default_theme = Theme::default();
        let theme = theme.unwrap_or(&default_theme);
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
            theme,
            master,
            layout,
            parsed_slide,
            primitives: Vec::new(),
            hit_regions: Vec::new(),
            shape_count: 0,
            line_count: 0,
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
        let normalized = normalize_family(family);
        self.faces
            .get(&(normalized.clone(), bold, italic))
            .or_else(|| self.faces.get(&(normalized, false, false)))
            .or(self.fallback.as_ref())
            .cloned()
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
    theme: &'a Theme,
    master: Option<&'a SlideMaster>,
    layout: Option<&'a SlideLayout>,
    parsed_slide: Option<&'a Slide>,
    primitives: Vec<Primitive>,
    hit_regions: Vec<HitRegion>,
    shape_count: usize,
    line_count: usize,
}

impl LayoutBuilder<'_> {
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
        let node_fill = original
            .and_then(node_fill)
            .or_else(|| layout_node.and_then(node_fill))
            .or_else(|| master_node.and_then(node_fill));
        let node_outline = original
            .and_then(node_outline)
            .or_else(|| layout_node.and_then(node_outline))
            .or_else(|| master_node.and_then(node_outline));
        let fill = shape
            .fill
            .as_ref()
            .or(node_fill)
            .and_then(|fill| paint(fill, self.theme));
        let outline = shape
            .outline
            .as_ref()
            .or(node_outline)
            .and_then(|outline| stroke(outline, self.theme));
        let transform = Transform {
            rotation_deg: shape.rotation_deg as f32,
            flip_h: shape.flip_h,
            flip_v: shape.flip_v,
        };
        match shape.kind {
            ShapeKind::Shape => {
                self.primitives.push(Primitive::Shape {
                    object_id: shape.source_id,
                    shape_id: Some(stable_id.clone()),
                    name: shape.name.clone(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    geometry: shape.geometry.clone(),
                    path: geometry_path(&shape.geometry),
                    adjust_values: BTreeMap::new(),
                    fill,
                    stroke: outline,
                    transform,
                });
            }
            ShapeKind::Picture => {
                self.primitives.push(Primitive::Image {
                    object_id: shape.source_id,
                    shape_id: Some(stable_id.clone()),
                    name: shape.name.clone(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    asset_id: shape.media_part_path.clone(),
                    stroke: outline,
                    transform,
                });
            }
            ShapeKind::GraphicFrame => {
                self.primitives.push(Primitive::Placeholder {
                    object_id: shape.source_id,
                    shape_id: Some(stable_id.clone()),
                    name: shape.name.clone(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    label: graphic_label(shape.graphic.as_ref()),
                    transform,
                });
            }
            ShapeKind::Group => unreachable!(),
        }
        let body_cascade = BodyCascade {
            primary: original.and_then(node_text),
            layout: layout_node.and_then(node_text),
            master: master_node.and_then(node_text),
            master_slide: self.master,
            placeholder: shape.placeholder.as_ref(),
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
                self.primitives.push(Primitive::Shape {
                    object_id: base.id,
                    shape_id: Some(stable_id.to_owned()),
                    name: base.name.clone(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    geometry: value.geometry.clone(),
                    path: geometry_path(&value.geometry),
                    adjust_values: BTreeMap::new(),
                    fill: value.fill.as_ref().and_then(|fill| paint(fill, self.theme)),
                    stroke: value
                        .outline
                        .as_ref()
                        .and_then(|outline| stroke(outline, self.theme)),
                    transform,
                });
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
                    stroke: value
                        .outline
                        .as_ref()
                        .and_then(|outline| stroke(outline, self.theme)),
                    transform,
                });
            }
            ShapeNode::GraphicFrame(value) => {
                self.primitives.push(Primitive::Placeholder {
                    object_id: base.id,
                    shape_id: Some(stable_id.to_owned()),
                    name: base.name.clone(),
                    x: rect.x,
                    y: rect.y,
                    w: rect.w,
                    h: rect.h,
                    label: graphic_label(Some(&value.data)),
                    transform,
                });
            }
            ShapeNode::Group(_) => unreachable!(),
        }
        let text_hit = if let Some(body) = node_text(shape) {
            let content = content_from_body(stable_id, body, self.theme);
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
                },
            )?)
        } else {
            None
        };
        self.hit_regions.push(HitRegion {
            shape_id: stable_id.to_owned(),
            rect,
            text: text_hit,
        });
        Ok(())
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

fn content_from_body(story_id: &str, body: &TextBody, theme: &Theme) -> TextContent {
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
                        text: run.text.clone(),
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
        paragraphs.push(ResolvedParagraph {
            align: parse_align(
                paragraph
                    .alignment
                    .as_deref()
                    .or(properties.alignment.as_deref()),
            ),
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
    let mut output = Vec::with_capacity(ranges.len());
    let mut line_y = y;
    for (start, end) in ranges {
        let slice = &clusters[start..end];
        let line_width = slice.iter().map(|cluster| cluster.width).sum::<f32>();
        let line_x = match paragraph.align {
            TextAlign::Center => x + ((width - line_width) / 2.0).max(0.0),
            TextAlign::Right => x + (width - line_width).max(0.0),
            TextAlign::Left | TextAlign::Justify => x,
        };
        let line_box = clusters_line_box(fonts, slice, scale)?;
        let mut caret_stops = vec![CaretStop {
            position: slice[0].start,
            x: line_x,
        }];
        let mut cursor_x = line_x;
        for cluster in slice {
            cursor_x += cluster.width;
            caret_stops.push(CaretStop {
                position: cluster.end,
                x: cursor_x,
            });
        }
        caret_stops.dedup_by(|left, right| {
            left.position == right.position && left.x.to_bits() == right.x.to_bits()
        });
        let runs = positioned_runs(slice, line_x, line_y + line_box.ascent, scale);
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
    line_x: f32,
    baseline: f32,
    scale: f32,
) -> Vec<PositionedTextRun> {
    let mut output: Vec<PositionedTextRun> = Vec::new();
    let mut cursor_x = line_x;
    for cluster in clusters {
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
        run.width += cluster.width;
        cursor_x += cluster.width;
    }
    output
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
    text: Option<TextHit>,
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
                    color: resolve_color_value_to_hex_with_theme(Some(&stop.color), Some(theme))?,
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
    resolve_color_value_to_hex_with_theme(fill.color.as_ref(), Some(theme))
        .map(|color| Paint::Solid { color })
}

fn stroke(outline: &ShapeOutline, theme: &Theme) -> Option<Stroke> {
    let color = resolve_color_value_to_hex_with_theme(outline.color.as_ref(), Some(theme))?;
    Some(Stroke {
        color,
        width: outline
            .width
            .filter(|width| width.is_finite() && *width >= 0.0)
            .map(|width| width as f32 / EMU_PER_CSS_PIXEL)
            .unwrap_or(1.0),
        dashed: outline
            .style
            .as_deref()
            .is_some_and(|style| style != "solid"),
    })
}

fn geometry_path(geometry: &str) -> Vec<ooxml_drawingml::GeometryPathCommand> {
    preset_geometry_to_path(geometry, &HashMap::new())
        .or_else(|| preset_geometry_to_path("rect", &HashMap::new()))
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
    use pptx_edit::{DeckSession, EditCtx};

    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    fn renderer() -> SlideRenderer {
        let mut renderer = SlideRenderer::new();
        for bold in [false, true] {
            renderer.register_font("Inter", bold, false, FONT).unwrap();
        }
        renderer
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
            geometry: "rect".to_owned(),
            placeholder: Some(title.clone()),
            fill: None,
            outline: None,
            media_part_path: None,
            graphic: None,
            text_stories: Vec::new(),
            children: Vec::new(),
        };
        let layout_shape = ShapeNode::Shape(pptx_parse::Shape {
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
}
