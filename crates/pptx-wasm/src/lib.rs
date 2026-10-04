//! PPTX display-list wasm boundary.

use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, HashMap};
use std::io::{self, Write};
use wasm_bindgen::prelude::*;

use pptx_edit::structured;
pub use pptx_edit::wasm::PptxDocument;

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(js_namespace = performance, js_name = now)]
    fn performance_now() -> f64;
}

/// stage latencies of one slide layout, in milliseconds. the clock comes from
/// the caller, so the ordinary layout path pays no timer calls.
struct LayoutProfile {
    scope_ms: f64,
    layout_ms: f64,
    serialize_ms: f64,
}

#[wasm_bindgen]
pub struct PptxRenderer {
    renderer: pptx_render::SlideRenderer,
    rendered: HashMap<String, CachedSlide>,
    last_slide: Option<String>,
    active_slide: Option<String>,
    clock: u64,
    font_epoch: u64,
    resources: Option<(yrs::Doc, usize, String)>,
    #[cfg(test)]
    layout_count: usize,
}

const LAYOUT_CACHE_CAPACITY: usize = 8;

struct CachedSlide {
    key: String,
    rendered: pptx_render::RenderedSlide,
    used: u64,
    validated_at: (String, u64),
}

#[wasm_bindgen]
impl PptxRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new() -> PptxRenderer {
        Self {
            renderer: pptx_render::SlideRenderer::new(),
            rendered: HashMap::new(),
            last_slide: None,
            active_slide: None,
            clock: 0,
            font_epoch: 0,
            resources: None,
            #[cfg(test)]
            layout_count: 0,
        }
    }

    #[wasm_bindgen(js_name = registerFont)]
    pub fn register_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32, JsValue> {
        let id = self
            .renderer
            .register_font(family, bold, italic, bytes)
            .map_err(js_error)?;
        self.font_epoch += 1;
        Ok(id)
    }

    #[wasm_bindgen(js_name = registerFallbackFont)]
    pub fn register_fallback_font(
        &mut self,
        family: &str,
        bold: bool,
        italic: bool,
        bytes: &[u8],
    ) -> Result<u32, JsValue> {
        let id = self
            .renderer
            .register_fallback_font(family, bold, italic, bytes)
            .map_err(js_error)?;
        self.font_epoch += 1;
        Ok(id)
    }

    #[wasm_bindgen(js_name = layoutSlideJson)]
    pub fn layout_slide_json(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        Ok(self
            .layout_slide_timed(document, slide_index, &mut || 0.0)?
            .0)
    }

    /// `layoutSlideJson` with its stages timed, returned as
    /// `{"layout": ..., "profile": {"scopeMs", "layoutMs", "serializeMs"}}`.
    #[wasm_bindgen(js_name = layoutSlideProfiledJson)]
    pub fn layout_slide_profiled_json(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let (json, profile) =
            self.layout_slide_timed(document, slide_index, &mut performance_now)?;
        let profile = serde_json::json!({
            "scopeMs": profile.scope_ms,
            "layoutMs": profile.layout_ms,
            "serializeMs": profile.serialize_ms,
        });
        Ok(format!("{{\"layout\":{json},\"profile\":{profile}}}"))
    }

    #[wasm_bindgen(js_name = hitTestJson)]
    pub fn hit_test_json(&self, x: f32, y: f32) -> Result<String, JsValue> {
        let result = self
            .last_slide
            .as_ref()
            .and_then(|id| self.rendered.get(id))
            .and_then(|cached| cached.rendered.hit_test(x, y));
        serde_json::to_string(&result).map_err(js_error)
    }

    #[wasm_bindgen(js_name = setActiveSlide)]
    pub fn set_active_slide(&mut self, id: &str, key: &str) -> bool {
        self.active_slide = Some(id.to_owned());
        let current = self
            .rendered
            .get(id)
            .is_some_and(|cached| cached.key == key);
        if current {
            self.last_slide = Some(id.to_owned());
        }
        current
    }

    #[wasm_bindgen(js_name = snapshotWithLayoutKeysJson)]
    pub fn snapshot_with_layout_keys_json(
        &mut self,
        document: &PptxDocument,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let snapshot = session.snapshot().map_err(js_error)?;
        let (version, resources) = self.resource_key(session)?;
        let mut keys = BTreeMap::new();
        for (index, slide) in snapshot.slides.iter().enumerate() {
            let key = layout_key(
                session.package(),
                slide,
                index,
                snapshot.width_emu,
                snapshot.height_emu,
                self.font_epoch,
                &resources,
            )
            .map_err(js_error)?;
            if let Some(cached) = self.rendered.get_mut(&slide.id)
                && cached.key == key
            {
                cached.validated_at = (version.clone(), self.font_epoch);
            }
            keys.insert(slide.id.clone(), key);
        }
        self.rendered.retain(|id, _| keys.contains_key(id));
        #[derive(serde::Serialize)]
        struct LayoutSnapshot<'a> {
            snapshot: &'a pptx_edit::DeckSnapshot,
            keys: BTreeMap<String, String>,
        }
        serde_json::to_string(&LayoutSnapshot {
            snapshot: &snapshot,
            keys,
        })
        .map_err(js_error)
    }

    #[wasm_bindgen(js_name = hitTestSlideJson)]
    pub fn hit_test_slide_json(
        &mut self,
        document: &PptxDocument,
        id: &str,
        x: f32,
        y: f32,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let version = session.version().to_string();
        let current = self
            .rendered
            .get(id)
            .is_some_and(|cached| cached.validated_at == (version, self.font_epoch));
        if !current {
            let index = session
                .slide_ids()
                .map_err(js_error)?
                .iter()
                .position(|slide| slide == id);
            let Some(index) = index else {
                return Ok("null".to_owned());
            };
            self.layout_slide_timed(document, index as u32, &mut || 0.0)?;
        }
        self.clock += 1;
        self.last_slide = Some(id.to_owned());
        let cached = self.rendered.get_mut(id).expect("layout cached the slide");
        cached.used = self.clock;
        serde_json::to_string(&cached.rendered.hit_test(x, y)).map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutProposalSlideJson)]
    pub fn layout_proposal_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let preview = document
            .session()
            .proposal_preview_session(id)
            .map_err(js_error)?;
        let scope = slide_scope(&preview, slide_index)?;
        let rendered = self
            .renderer
            .layout_scoped_slide(preview.package(), &scope)
            .map_err(js_error)?;
        serde_json::to_string(&rendered.display_list).map_err(js_error)
    }

    #[wasm_bindgen(js_name = layoutProposalDiffSlideJson)]
    pub fn layout_proposal_diff_slide_json(
        &self,
        document: &PptxDocument,
        id: &str,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let preview = session
            .preview_proposal_diff_slide(id, slide_index as usize)
            .map_err(|error| match error {
                pptx_edit::ProposalError::Edit(pptx_edit::EditError::OutOfBounds { .. }) => {
                    js_error(pptx_render::RenderError::SlideNotFound(
                        slide_index as usize,
                    ))
                }
                error => js_error(error),
            })?;
        let rendered = self
            .renderer
            .layout_scoped_slide(session.package(), &preview.scope)
            .map_err(js_error)?;
        serde_json::to_string(&serde_json::json!({
            "proposal": preview.proposal,
            "snapshot": preview.snapshot,
            "textChanges": preview.text_changes,
            "frame": rendered.display_list,
        }))
        .map_err(js_error)
    }
}

impl PptxRenderer {
    fn resource_key(
        &mut self,
        session: &pptx_edit::DeckSession,
    ) -> Result<(String, String), JsValue> {
        let version = session.version().to_string();
        let package = session.package() as *const pptx_parse::PptxPackage as usize;
        // A session's parsed package is immutable; retain the doc to prevent identity reuse.
        if self.resources.as_ref().is_none_or(|(doc, address, _)| {
            !yrs::Doc::ptr_eq(doc, session.yrs_doc()) || *address != package
        }) {
            self.resources = Some((
                session.yrs_doc().clone(),
                package,
                resource_key(session.package()).map_err(js_error)?,
            ));
        }
        Ok((version, self.resources.as_ref().unwrap().2.clone()))
    }

    fn cache_slide(
        &mut self,
        session: &pptx_edit::DeckSession,
        scope: &pptx_edit::SlideScope,
    ) -> Result<(), JsValue> {
        let (version, resources) = self.resource_key(session)?;
        let key = layout_key(
            session.package(),
            &scope.slide,
            scope.index,
            scope.width_emu,
            scope.height_emu,
            self.font_epoch,
            &resources,
        )
        .map_err(js_error)?;
        let id = &scope.slide.id;
        self.clock += 1;
        if self.rendered.get(id).is_none_or(|cached| cached.key != key) {
            let rendered = self
                .renderer
                .layout_scoped_slide(session.package(), scope)
                .map_err(js_error)?;
            #[cfg(test)]
            {
                self.layout_count += 1;
            }
            self.rendered.insert(
                id.clone(),
                CachedSlide {
                    key,
                    rendered,
                    used: self.clock,
                    validated_at: (version.clone(), self.font_epoch),
                },
            );
        }
        let cached = self.rendered.get_mut(id).unwrap();
        cached.used = self.clock;
        cached.validated_at = (version, self.font_epoch);
        self.last_slide = Some(id.clone());
        while self.rendered.len() > LAYOUT_CACHE_CAPACITY {
            let oldest = self
                .rendered
                .iter()
                .filter(|(candidate, _)| {
                    Some(*candidate) != self.active_slide.as_ref() && *candidate != id
                })
                .min_by_key(|(_, cached)| cached.used)
                .map(|(id, _)| id.clone())
                .unwrap();
            self.rendered.remove(&oldest);
        }
        Ok(())
    }

    fn layout_slide_timed(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
        now: &mut impl FnMut() -> f64,
    ) -> Result<(String, LayoutProfile), JsValue> {
        let started = now();
        let session = document.session();
        let scope = slide_scope(session, slide_index)?;
        let scoped = now();
        self.cache_slide(session, &scope)?;
        let laid_out = now();
        let json = serde_json::to_string(&self.rendered[&scope.slide.id].rendered.display_list)
            .map_err(js_error)?;
        let serialized = now();
        Ok((
            json,
            LayoutProfile {
                scope_ms: scoped - started,
                layout_ms: laid_out - scoped,
                serialize_ms: serialized - laid_out,
            },
        ))
    }
}

struct HashWriter(Sha256);

impl Write for HashWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.0.update(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn json_key(value: &impl serde::Serialize) -> Result<String, serde_json::Error> {
    let mut writer = HashWriter(Sha256::new());
    serde_json::to_writer(&mut writer, value)?;
    Ok(format!("{:x}", writer.0.finalize()))
}

fn resource_key(package: &pptx_parse::PptxPackage) -> Result<String, serde_json::Error> {
    json_key(&(
        &package.charts,
        &package.media,
        &package.diagram_drawings,
        &package.table_styles,
        package.presentation.first_slide_num,
        &package.presentation.default_text_style,
        &package.presentation.default_text_paragraph,
    ))
}

fn layout_key(
    package: &pptx_parse::PptxPackage,
    slide: &pptx_edit::SlideSnapshot,
    index: usize,
    width: i64,
    height: i64,
    font_epoch: u64,
    resources: &str,
) -> Result<String, serde_json::Error> {
    let parents = pptx_edit::paragraph::SlideParents::resolve(
        package,
        slide.source_part_path.as_deref(),
        slide.layout_part_path.as_deref(),
    );
    let theme = parents
        .master
        .and_then(|master| master.theme_part_path.as_deref())
        .and_then(|path| package.themes.iter().find(|theme| theme.part_path == path))
        .or_else(|| package.themes.first());
    json_key(&(
        slide,
        index,
        width,
        height,
        font_epoch,
        resources,
        parents.slide,
        parents.layout,
        parents.master,
        theme,
    ))
}

impl Default for PptxRenderer {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen(js_name = parsePptxJson)]
pub fn parse_pptx_json(data: &[u8]) -> Result<String, JsValue> {
    let package = pptx_parse::parse_pptx(data).map_err(js_error)?;
    serde_json::to_string(&package).map_err(js_error)
}

/// Structured export of PPTX bytes as a snapshot, with the options of
/// `PptxDocument.exportStructuredJson`: `{"ok":true,"content"}` or `{"ok":false,"failure"}`.
/// Bytes that are not a readable PPTX throw.
#[wasm_bindgen(js_name = exportPptxStructuredJson)]
pub fn export_pptx_structured_json(data: &[u8], options: &str) -> Result<String, JsValue> {
    let options = serde_json::from_str(options).map_err(js_error)?;
    structured::snapshot_outcome_json(structured::export_pptx_structured(data, &options))
        .map_err(|message| JsValue::from_str(&message))
}

/// `exportPptxStructuredJson` rendered as Markdown.
#[wasm_bindgen(js_name = exportPptxMarkdownJson)]
pub fn export_pptx_markdown_json(data: &[u8], options: &str) -> Result<String, JsValue> {
    let options = serde_json::from_str(options).map_err(js_error)?;
    structured::snapshot_outcome_json(structured::export_pptx_markdown(data, &options))
        .map_err(|message| JsValue::from_str(&message))
}

/// Renders schema-version-1 structured content as Markdown; `options` is `{"maxBytes"?}`.
#[wasm_bindgen(js_name = renderPptxMarkdownJson)]
pub fn render_pptx_markdown_json(content: &str, options: &str) -> Result<String, JsValue> {
    let content: structured::PptxStructuredContent =
        serde_json::from_str(content).map_err(js_error)?;
    let options: structured::PptxMarkdownOptions =
        serde_json::from_str(options).map_err(js_error)?;
    structured::snapshot_outcome_json(
        structured::render_pptx_markdown(&content, &options)
            .map_err(structured::ExportError::Refused),
    )
    .map_err(|message| JsValue::from_str(&message))
}

#[wasm_bindgen(js_name = compileSlideJson)]
pub fn compile_slide_json(slide_json: &str) -> Result<String, JsValue> {
    pptx_render::compile_json(slide_json).map_err(|error| JsValue::from_str(&error))
}

#[wasm_bindgen(js_name = rendererVersion)]
pub fn renderer_version() -> String {
    env!("CARGO_PKG_VERSION").to_owned()
}

#[wasm_bindgen(js_name = decodeTiffPng)]
pub fn decode_tiff_png(data: &[u8]) -> Result<Vec<u8>, JsValue> {
    ooxml_drawingml::media::decode_tiff_png(data).map_err(js_error)
}

fn slide_scope(
    session: &pptx_edit::DeckSession,
    slide_index: u32,
) -> Result<pptx_edit::SlideScope, JsValue> {
    session
        .slide_scope(slide_index as usize)
        .map_err(|error| match error {
            pptx_edit::EditError::OutOfBounds { .. } => js_error(
                pptx_render::RenderError::SlideNotFound(slide_index as usize),
            ),
            error => js_error(error),
        })
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pptx_edit::{EditCtx, SlideScope, TextStyle};

    const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    fn key(package: &pptx_parse::PptxPackage, scope: &SlideScope, epoch: u64) -> String {
        layout_key(
            package,
            &scope.slide,
            scope.index,
            scope.width_emu,
            scope.height_emu,
            epoch,
            &resource_key(package).unwrap(),
        )
        .unwrap()
    }

    #[test]
    fn keys_track_local_text_and_every_inherited_input() {
        let document = PptxDocument::open_collaborative(DECK, 910.0).unwrap();
        let session = document.session();
        let scope = session.slide_scope(0).unwrap();
        let package = session.package();
        let original = key(package, &scope, 0);
        let other = session.slide_scope(1).unwrap();
        let other_key = key(package, &other, 0);
        let story = &scope
            .slide
            .shapes
            .iter()
            .flat_map(|shape| &shape.text_stories)
            .next()
            .unwrap()
            .id;
        session
            .insert_text(
                &EditCtx::local("test"),
                story,
                0,
                "Edited ",
                &TextStyle::default(),
            )
            .unwrap();
        assert_ne!(original, key(package, &session.slide_scope(0).unwrap(), 0));
        assert_eq!(other_key, key(package, &session.slide_scope(1).unwrap(), 0));

        let parents = pptx_edit::paragraph::SlideParents::resolve(
            package,
            scope.slide.source_part_path.as_deref(),
            scope.slide.layout_part_path.as_deref(),
        );
        let mut changed = package.clone();
        let layout = parents.layout.unwrap();
        changed
            .layouts
            .iter_mut()
            .find(|item| item.part_path == layout.part_path)
            .unwrap()
            .show_master_shapes ^= true;
        assert_ne!(original, key(&changed, &scope, 0));
        let mut changed = package.clone();
        let master = parents.master.unwrap();
        changed
            .masters
            .iter_mut()
            .find(|item| item.part_path == master.part_path)
            .unwrap()
            .color_map
            .set("bg1", "accent1");
        assert_ne!(original, key(&changed, &scope, 0));
        let mut changed = package.clone();
        let theme = master.theme_part_path.as_deref().unwrap();
        changed
            .themes
            .iter_mut()
            .find(|item| item.part_path == theme)
            .unwrap()
            .theme
            .color_scheme
            .accent1 = "123456".to_owned();
        assert_ne!(original, key(&changed, &scope, 0));
        let mut resized = scope.clone();
        resized.width_emu += 1;
        assert_ne!(original, key(package, &resized, 0));
        resized = scope.clone();
        resized.height_emu += 1;
        assert_ne!(original, key(package, &resized, 0));
        resized = scope.clone();
        resized.index += 1;
        assert_ne!(original, key(package, &resized, 0));
        let mut changed = package.clone();
        changed.presentation.first_slide_num += 1;
        assert_ne!(original, key(&changed, &scope, 0));
        let mut changed = package.clone();
        changed.media[0].bytes.push(0);
        assert_ne!(original, key(&changed, &scope, 0));

        let mut renderer = PptxRenderer::new();
        renderer
            .register_font("Liberation Sans", false, false, FONT)
            .unwrap();
        assert_ne!(original, key(package, &scope, renderer.font_epoch));
        let regular = key(package, &scope, renderer.font_epoch);
        renderer
            .register_fallback_font("Fallback", false, false, FONT)
            .unwrap();
        assert_ne!(regular, key(package, &scope, renderer.font_epoch));
    }

    #[test]
    fn cached_hits_and_display_lists_match_fresh_layouts_and_remain_bounded() {
        let document = PptxDocument::open_collaborative(DECK, 911.0).unwrap();
        let session = document.session();
        let mut renderer = PptxRenderer::new();
        renderer
            .register_font("Liberation Sans", false, false, FONT)
            .unwrap();
        let scope = session.slide_scope(0).unwrap();
        renderer.set_active_slide(&scope.slide.id, "");
        let json = renderer.layout_slide_json(&document, 0).unwrap();
        let fresh = renderer
            .renderer
            .layout_scoped_slide(session.package(), &scope)
            .unwrap();
        assert_eq!(json, serde_json::to_string(&fresh.display_list).unwrap());
        assert_eq!(renderer.layout_slide_json(&document, 0).unwrap(), json);
        assert_eq!(renderer.layout_count, 1);
        for _ in 0..16 {
            let index = session.slide_ids().unwrap().len() as u32;
            session
                .insert_slide(&EditCtx::local("test"), index, None)
                .unwrap();
            renderer.layout_slide_json(&document, index).unwrap();
            assert!(renderer.rendered.len() <= LAYOUT_CACHE_CAPACITY);
        }
        let layouts = renderer.layout_count;
        let mut hits = 0;
        for y in (0..fresh.display_list.height as usize).step_by(32) {
            for x in (0..fresh.display_list.width as usize).step_by(32) {
                let expected = fresh.hit_test(x as f32, y as f32);
                hits += usize::from(expected.is_some());
                assert_eq!(
                    renderer
                        .hit_test_slide_json(&document, &scope.slide.id, x as f32, y as f32)
                        .unwrap(),
                    serde_json::to_string(&expected).unwrap(),
                );
            }
        }
        assert!(hits > 0);
        assert_eq!(renderer.layout_count, layouts);
        let story = &scope
            .slide
            .shapes
            .iter()
            .flat_map(|shape| &shape.text_stories)
            .next()
            .unwrap()
            .id;
        session
            .insert_text(
                &EditCtx::local("test"),
                story,
                0,
                "New text ",
                &TextStyle::default(),
            )
            .unwrap();
        renderer
            .hit_test_slide_json(&document, &scope.slide.id, 10.0, 10.0)
            .unwrap();
        assert_eq!(renderer.layout_count, layouts + 1);
        let updated = session.slide_scope(0).unwrap();
        let fresh = renderer
            .renderer
            .layout_scoped_slide(session.package(), &updated)
            .unwrap();
        assert_eq!(
            renderer.layout_slide_json(&document, 0).unwrap(),
            serde_json::to_string(&fresh.display_list).unwrap()
        );
        assert_eq!(renderer.layout_count, layouts + 1);
        renderer
            .register_fallback_font("Fallback", false, false, FONT)
            .unwrap();
        renderer
            .hit_test_slide_json(&document, &scope.slide.id, 10.0, 10.0)
            .unwrap();
        assert_eq!(renderer.layout_count, layouts + 2);
        let fresh = renderer
            .renderer
            .layout_scoped_slide(session.package(), &updated)
            .unwrap();
        assert_eq!(
            renderer.hit_test_json(10.0, 10.0).unwrap(),
            serde_json::to_string(&fresh.hit_test(10.0, 10.0)).unwrap()
        );
    }
}
