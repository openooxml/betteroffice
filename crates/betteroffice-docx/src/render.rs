//! Server-side rasterization of a display list.

#[cfg(feature = "render-path-counter")]
use std::cell::Cell;
use std::collections::BTreeMap;
use std::sync::{Mutex, PoisonError};

use docx_layout::display_list::DisplayList;
use docx_parse::notes::Note;
use docx_parse::s9::S9SectionWire;
use docx_raster::{
    ImageCache, ImageMap, ImageScope, RenderResources, RenderedPage, scoped_image_key,
};
use serde_json::{Number, Value, json};

use crate::{Document, DocumentModel, Error, Result};

#[cfg(feature = "render-path-counter")]
std::thread_local! {
    static MODEL_RENDERS: Cell<usize> = const { Cell::new(0) };
}

const MAX_IMAGE_BYTES: usize = 32 * 1024 * 1024;
const MAX_IMAGES: usize = 256;
pub const MAX_PIXMAP_DIM: u32 = docx_raster::MAX_PAGE_DIM;
pub const MAX_PIXMAP_PIXELS: u64 = docx_raster::MAX_PAGE_PIXELS;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PngTarget {
    pub page: usize,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RenderOptions {
    pub scale: f32,
}

impl Default for RenderOptions {
    fn default() -> Self {
        Self { scale: 1.0 }
    }
}

/// Caller-supplied bytes for relationship ids the display list did not already
/// carry as `data:` URLs, keyed by owning part, plus the decoded images a page
/// leaves for the next — behind a `Mutex` so the document stays `Send + Sync`.
#[derive(Default)]
pub(crate) struct ImageRegistry {
    entries: ImageMap,
    cache: Mutex<ImageCache>,
}

impl ImageRegistry {
    fn register(&mut self, scope: ImageScope<'_>, rel_id: &str, bytes: &[u8]) -> Result<()> {
        if bytes.len() > MAX_IMAGE_BYTES {
            return Err(Error::ResourceLimit(format!(
                "image exceeds {MAX_IMAGE_BYTES} bytes"
            )));
        }
        if rel_id.is_empty() {
            return Err(Error::Image("image relationship id is empty".to_owned()));
        }
        let key = scoped_image_key(scope, rel_id);
        if self.entries.len() >= MAX_IMAGES && !self.entries.contains_key(&key) {
            return Err(Error::ResourceLimit(format!(
                "more than {MAX_IMAGES} images"
            )));
        }
        self.entries.insert(key, bytes.to_vec());
        Ok(())
    }
}

fn validate_render_size(width: u32, height: u32) -> Result<()> {
    if width > MAX_PIXMAP_DIM || height > MAX_PIXMAP_DIM {
        return Err(Error::RenderTooLarge {
            width,
            height,
            max: MAX_PIXMAP_DIM,
        });
    }
    if u64::from(width) * u64::from(height) > MAX_PIXMAP_PIXELS {
        return Err(Error::RenderAreaTooLarge {
            width,
            height,
            max_pixels: MAX_PIXMAP_PIXELS,
        });
    }
    Ok(())
}

/// The pixel extent the raster backend would allocate, or `None` when the
/// backend rejects the number before allocating anything.
fn page_pixels(number: &Number, scale: f32) -> Option<u32> {
    let value = number.as_f64()? * f64::from(scale);
    if !value.is_finite() || value <= 0.0 {
        return None;
    }
    Some(value.ceil().min(f64::from(u32::MAX)) as u32)
}

fn validate_scale(options: &RenderOptions) -> Result<()> {
    if !options.scale.is_finite() || options.scale <= 0.0 {
        return Err(Error::Render("scale must be finite and positive".into()));
    }
    Ok(())
}

fn json_error(error: serde_json::Error) -> Error {
    Error::Render(error.to_string())
}

/// The package fields a region layout request reads, serialized as in the parsed package.
fn layout_package(model: &DocumentModel) -> Result<Value> {
    let notes = |notes: &[Note]| -> Vec<Value> {
        notes
            .iter()
            .map(|note| json!({"id": note.id, "noteType": note.note_type}))
            .collect()
    };
    let mut document = serde_json::Map::new();
    if let Some(sections) = &model.body.sections {
        let sections: Vec<S9SectionWire> = sections
            .iter()
            .map(|section| S9SectionWire {
                id: section.id.clone(),
                properties: section.properties.clone(),
                content_start: 0,
                content_end: 0,
            })
            .collect();
        let sections = serde_json::to_value(sections).map_err(json_error)?;
        document.insert("sections".to_owned(), sections);
    }
    if let Some(properties) = &model.body.final_section_properties {
        let properties = serde_json::to_value(properties).map_err(json_error)?;
        document.insert("finalSectionProperties".to_owned(), properties);
    }
    let theme = serde_json::to_value(&model.theme).map_err(json_error)?;
    let settings = serde_json::to_value(&model.settings).map_err(json_error)?;
    let mut package = json!({
        "document": document,
        "theme": theme,
        "settings": settings,
        "footnotes": notes(&model.footnotes),
        "endnotes": notes(&model.endnotes),
    });
    if let Some(styles) = &model.styles {
        package["styles"] = serde_json::to_value(styles).map_err(json_error)?;
    }
    Ok(package)
}

impl Document {
    pub fn render_png(&self, target: PngTarget, options: &RenderOptions) -> Result<RenderedPage> {
        if !self.renders_from_model() {
            return self.render_png_saved(target, options);
        }
        let Some(rendered) = self.render_png_from_model(target, options).transpose() else {
            return self.render_png_saved(target, options);
        };
        #[cfg(feature = "render-path-counter")]
        MODEL_RENDERS.with(|renders| renders.set(renders.get() + 1));
        rendered
    }

    /// The page laid out from the model, or `None` when its conversion left content unpainted.
    fn render_png_from_model(
        &self,
        target: PngTarget,
        options: &RenderOptions,
    ) -> Result<Option<RenderedPage>> {
        validate_scale(options)?;
        let package = layout_package(self.model())?;
        let engine = docx_edit::EngineSession::new(1);
        docx_edit::seed_from_parsed_docx(
            engine.doc(),
            self.envelope(),
            self.original_parts.clone(),
            self.source_bytes(),
        )
        .map_err(Error::Render)?;
        let (display_list, unreadable) =
            self.seeded_display_list(&engine, &package, |request, chains| {
                engine
                    .layout_document_with_regions_retained(request)
                    .map_err(Error::Render)?;
                engine
                    .retained_page_display_list(chains, target.page)
                    .map_err(Error::Render)
            })?;
        if !unreadable.is_empty() {
            return Ok(None);
        }
        self.render_display_list_scaled(&display_list, target.page, options.scale).map(Some)
    }

    /// [`Document::render_png`] through a save and a parse of the saved bytes, in every state.
    // Test hook; not part of the stable API.
    #[doc(hidden)]
    pub fn render_png_saved(
        &self,
        target: PngTarget,
        options: &RenderOptions,
    ) -> Result<RenderedPage> {
        validate_scale(options)?;
        let bytes = self.save()?;
        let package = docx_parse::parse_docx_s9_wire(
            &bytes,
            docx_parse::S9ParseOptions {
                detect_variables: false,
                ..Default::default()
            },
        )?
        .document
        .package;
        let engine = docx_edit::EngineSession::new(1);
        docx_edit::seed_from_docx(engine.doc(), &bytes).map_err(Error::Render)?;
        let package = serde_json::to_value(package).map_err(json_error)?;
        let display_list = self.seeded_display_list(&engine, &package, |request, chains| {
            let output = engine
                .layout_document_with_regions_json(request)
                .map_err(Error::Render)?;
            let mut output: Value = serde_json::from_str(&output).map_err(json_error)?;
            output["fontChains"] = json!(chains);
            docx_layout::build_display_list_value(&output.to_string()).map_err(Error::Render)
        })?;
        self.render_display_list_scaled(&display_list, target.page, options.scale)
    }

    /// How many [`Document::render_png`] calls on this thread laid out from the model.
    // Test hook; not part of the stable API.
    #[cfg(feature = "render-path-counter")]
    #[doc(hidden)]
    pub fn model_renders() -> usize {
        MODEL_RENDERS.with(Cell::get)
    }

    fn seeded_display_list<T>(
        &self,
        engine: &docx_edit::EngineSession,
        package: &Value,
        layout: impl FnOnce(&str, &BTreeMap<String, Vec<u32>>) -> Result<T>,
    ) -> Result<T> {
        let mut request = docx_edit::layout_request::region_layout_request(package, 24.0);
        request["measurement"] = docx_edit::layout_request::measurement(package);
        let mut fonts = self.fonts.lock().unwrap_or_else(PoisonError::into_inner);
        #[cfg(feature = "fonts")]
        {
            let requirements = engine
                .layout_font_requirements_json(&request.to_string())
                .map_err(Error::Render)?;
            let requirements: Vec<serde_json::Value> = serde_json::from_str(&requirements)
                .map_err(|error| Error::Render(error.to_string()))?;
            let requirements: Vec<_> = requirements
                .iter()
                .map(|value| docx_layout::measure_blocks::FontRequirement {
                    key: value["key"].as_str().unwrap_or_default().to_owned(),
                    family: value["family"].as_str().unwrap_or("Calibri").to_owned(),
                    bold: value["bold"].as_bool().unwrap_or(false),
                    italic: value["italic"].as_bool().unwrap_or(false),
                    scripts: value["scripts"]
                        .as_array()
                        .into_iter()
                        .flatten()
                        .filter_map(|script| script.as_str().map(str::to_owned))
                        .collect(),
                    characters: Default::default(),
                })
                .collect();
            fonts.resolve(&self.font_manager, &requirements)?;
        }
        let chains: BTreeMap<String, Vec<u32>> = fonts
            .chains
            .iter()
            .map(|(key, ids)| (key.clone(), ids.iter().map(|id| id.to_u32()).collect()))
            .collect();
        request["measurement"]["fontChains"] = json!(chains);
        let request = request.to_string();
        let display_list =
            docx_layout::with_measure_font_store(&mut fonts.store, || layout(request.as_str(), &chains));
        drop(fonts);
        display_list
    }

    /// Supplies bytes for one relationship id the display list left
    /// unresolved, scoped to the part that owns the id. Images are capped at
    /// 256 and each at 32 MiB.
    pub fn register_image(
        &mut self,
        scope: ImageScope<'_>,
        rel_id: &str,
        bytes: &[u8],
    ) -> Result<()> {
        self.images.register(scope, rel_id, bytes)
    }

    /// Rasterizes one display-list page to deterministic PNG bytes. A page past
    /// [`MAX_PIXMAP_DIM`] or [`MAX_PIXMAP_PIXELS`] is refused before any surface
    /// is allocated; an image the backend cannot draw is drawn as a placeholder
    /// in its frame and counted.
    /// Glyph outlines and decoded images are cached on the document, so a page
    /// reuses what earlier pages extracted and decoded; an export past either
    /// cache's cap re-does what it evicted.
    pub fn render_display_list_png(
        &self,
        display_list: &DisplayList,
        page_ordinal: usize,
    ) -> Result<RenderedPage> {
        self.render_display_list_scaled(display_list, page_ordinal, 1.0)
    }

    fn render_display_list_scaled(
        &self,
        display_list: &DisplayList,
        page_ordinal: usize,
        scale: f32,
    ) -> Result<RenderedPage> {
        if let Some(page) = display_list.pages.get(page_ordinal)
            && let (Some(width), Some(height)) = (
                page_pixels(&page.width, scale),
                page_pixels(&page.height, scale),
            )
        {
            validate_render_size(width, height)?;
        }
        #[cfg(feature = "fonts")]
        self.prepare_display_fonts(display_list, page_ordinal)?;
        let mut fonts = self.fonts.lock().unwrap_or_else(PoisonError::into_inner);
        let mut images = self
            .images
            .cache
            .lock()
            .unwrap_or_else(PoisonError::into_inner);
        let crate::font::FontRegistry {
            store,
            chains,
            glyphs,
            ..
        } = &mut *fonts;
        let resources = RenderResources::new(store, chains, &self.images.entries);
        docx_raster::render_page_cached_scaled(
            display_list,
            page_ordinal,
            &resources,
            glyphs,
            &mut images,
            scale,
        )
        .map_err(Error::Render)
    }
}
