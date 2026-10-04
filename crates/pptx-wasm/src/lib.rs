//! PPTX display-list wasm boundary.

use serde::{Serialize, ser};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::rc::Rc;
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
    package_identity: Option<(yrs::Doc, usize)>,
    generation: u64,
    next_token: u64,
    layout_keys: HashMap<String, LayoutKeyRecord>,
    #[cfg(test)]
    layout_count: usize,
    #[cfg(test)]
    key_count: usize,
}

const LAYOUT_CACHE_CAPACITY: usize = 8;

struct CachedSlide {
    key: String,
    rendered: pptx_render::RenderedSlide,
    used: u64,
    validated_at: (Rc<pptx_edit::DocumentVersion>, u64),
}

struct LayoutKeyRecord {
    scope: pptx_edit::SlideScope,
    font_epoch: u64,
    generation: u64,
    validated_version: Rc<pptx_edit::DocumentVersion>,
    token: u64,
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
            package_identity: None,
            generation: 0,
            next_token: 0,
            layout_keys: HashMap::new(),
            #[cfg(test)]
            layout_count: 0,
            #[cfg(test)]
            key_count: 0,
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
        self.layout_keys.clear();
        self.rendered.clear();
        self.last_slide = None;
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
        self.layout_keys.clear();
        self.rendered.clear();
        self.last_slide = None;
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
    pub fn set_active_slide(&mut self, document: &PptxDocument, id: &str, key: &str) -> bool {
        let session = document.session();
        self.sync_package(session);
        let version = session.version();
        self.active_slide = Some(id.to_owned());
        let current = self.layout_keys.get(id).is_some_and(|record| {
            record.generation == self.generation
                && record.font_epoch == self.font_epoch
                && record.validated_version.as_ref() == &version
                && record.token.to_string() == key
        }) && self
            .rendered
            .get(id)
            .is_some_and(|cached| cached.key == key);
        self.last_slide = current.then(|| id.to_owned());
        current
    }

    #[wasm_bindgen(js_name = snapshotWithLayoutKeysJson)]
    pub fn snapshot_with_layout_keys_json(
        &mut self,
        document: &PptxDocument,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let snapshot = session.snapshot().map_err(js_error)?;
        let version = Rc::new(session.version());
        self.sync_package(session);
        let mut keys = BTreeMap::new();
        for (index, slide) in snapshot.slides.iter().enumerate() {
            let key = self.slide_key(
                &version,
                slide,
                index,
                snapshot.width_emu,
                snapshot.height_emu,
            );
            keys.insert(slide.id.clone(), key);
        }
        self.rendered.retain(|id, _| keys.contains_key(id));
        self.layout_keys.retain(|id, _| keys.contains_key(id));
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

    #[wasm_bindgen(js_name = slideLayoutKey)]
    pub fn slide_layout_key(
        &mut self,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String, JsValue> {
        let session = document.session();
        let scope = slide_scope(session, slide_index)?;
        let version = Rc::new(session.version());
        self.sync_package(session);
        let key = self.slide_key(
            &version,
            &scope.slide,
            scope.index,
            scope.width_emu,
            scope.height_emu,
        );
        self.prune_layout_keys(session)?;
        Ok(key)
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
        let version = Rc::new(session.version());
        self.sync_package(session);
        let current = self.rendered.get(id).is_some_and(|cached| {
            cached.validated_at.0 == version && cached.validated_at.1 == self.font_epoch
        });
        if !current {
            let index = session
                .slide_ids()
                .map_err(js_error)?
                .iter()
                .position(|slide| slide == id);
            let Some(index) = index else {
                return Ok("null".to_owned());
            };
            let scope = slide_scope(session, index as u32)?;
            self.cache_slide(session, &scope, &version)?;
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
    fn prune_layout_keys(&mut self, session: &pptx_edit::DeckSession) -> Result<(), JsValue> {
        if self.layout_keys.len() > 256 {
            let ids = session.slide_ids().map_err(js_error)?;
            if self.layout_keys.len() > 256.max(ids.len().saturating_mul(2)) {
                let live: HashSet<_> = ids.into_iter().collect();
                self.layout_keys.retain(|id, _| live.contains(id));
            }
        }
        Ok(())
    }

    fn sync_package(&mut self, session: &pptx_edit::DeckSession) {
        let package = session.package() as *const pptx_parse::PptxPackage as usize;
        if self.package_identity.as_ref().is_none_or(|(doc, address)| {
            !yrs::Doc::ptr_eq(doc, session.yrs_doc()) || *address != package
        }) {
            self.package_identity = Some((session.yrs_doc().clone(), package));
            self.generation = self
                .generation
                .checked_add(1)
                .expect("generation exhausted");
            self.layout_keys.clear();
            self.rendered.clear();
            self.last_slide = None;
        }
    }

    fn slide_key(
        &mut self,
        version: &Rc<pptx_edit::DocumentVersion>,
        slide: &pptx_edit::SlideSnapshot,
        index: usize,
        width: i64,
        height: i64,
    ) -> String {
        if let Some(record) = self.layout_keys.get_mut(&slide.id)
            && record.generation == self.generation
            && record.font_epoch == self.font_epoch
            && (record.validated_version == *version
                || (record.scope.slide == *slide
                    && record.scope.index == index
                    && record.scope.width_emu == width
                    && record.scope.height_emu == height
                    && float_bits(&record.scope.slide)
                        .is_some_and(|bits| Some(bits) == float_bits(slide))))
        {
            if record.validated_version != *version {
                record.validated_version = Rc::clone(version);
            }
            let key = record.token.to_string();
            if let Some(cached) = self.rendered.get_mut(&slide.id)
                && cached.key == key
            {
                cached.validated_at = (Rc::clone(version), self.font_epoch);
            }
            return key;
        }
        self.next_token = self
            .next_token
            .checked_add(1)
            .expect("layout token exhausted");
        #[cfg(test)]
        {
            self.key_count += 1;
        }
        self.layout_keys.insert(
            slide.id.clone(),
            LayoutKeyRecord {
                scope: pptx_edit::SlideScope {
                    slide: slide.clone(),
                    index,
                    width_emu: width,
                    height_emu: height,
                },
                font_epoch: self.font_epoch,
                generation: self.generation,
                validated_version: Rc::clone(version),
                token: self.next_token,
            },
        );
        self.next_token.to_string()
    }

    fn cache_slide(
        &mut self,
        session: &pptx_edit::DeckSession,
        scope: &pptx_edit::SlideScope,
        version: &Rc<pptx_edit::DocumentVersion>,
    ) -> Result<(), JsValue> {
        let key = self.slide_key(
            version,
            &scope.slide,
            scope.index,
            scope.width_emu,
            scope.height_emu,
        );
        self.prune_layout_keys(session)?;
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
                    validated_at: (Rc::clone(version), self.font_epoch),
                },
            );
        }
        let cached = self.rendered.get_mut(id).unwrap();
        cached.used = self.clock;
        cached.validated_at = (Rc::clone(version), self.font_epoch);
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
        let version = Rc::new(session.version());
        self.sync_package(session);
        self.cache_slide(session, &scope, &version)?;
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

#[derive(Default)]
struct FloatBits(Vec<u64>);

fn float_bits(value: &impl Serialize) -> Option<Vec<u64>> {
    let mut bits = FloatBits::default();
    value.serialize(&mut bits).ok()?;
    Some(bits.0)
}

macro_rules! ignore_primitives {
    ($($method:ident($ty:ty)),* $(,)?) => {
        $(fn $method(self, _: $ty) -> Result<(), Self::Error> { Ok(()) })*
    };
}

impl ser::Serializer for &mut FloatBits {
    type Ok = ();
    type Error = serde::de::value::Error;
    type SerializeSeq = Self;
    type SerializeTuple = Self;
    type SerializeTupleStruct = Self;
    type SerializeTupleVariant = Self;
    type SerializeMap = Self;
    type SerializeStruct = Self;
    type SerializeStructVariant = Self;

    ignore_primitives! {
        serialize_bool(bool),
        serialize_i8(i8), serialize_i16(i16), serialize_i32(i32),
        serialize_i64(i64), serialize_i128(i128),
        serialize_u8(u8), serialize_u16(u16), serialize_u32(u32),
        serialize_u64(u64), serialize_u128(u128),
        serialize_char(char), serialize_str(&str), serialize_bytes(&[u8]),
    }

    fn serialize_f32(self, value: f32) -> Result<(), Self::Error> {
        self.0.push(u64::from(value.to_bits()));
        Ok(())
    }

    fn serialize_f64(self, value: f64) -> Result<(), Self::Error> {
        self.0.push(value.to_bits());
        Ok(())
    }

    fn serialize_none(self) -> Result<(), Self::Error> {
        Ok(())
    }

    fn serialize_some<T: ?Sized + Serialize>(self, value: &T) -> Result<(), Self::Error> {
        value.serialize(self)
    }

    fn serialize_unit(self) -> Result<(), Self::Error> {
        Ok(())
    }

    fn serialize_unit_struct(self, _: &'static str) -> Result<(), Self::Error> {
        Ok(())
    }

    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
    ) -> Result<(), Self::Error> {
        Ok(())
    }

    fn serialize_newtype_struct<T: ?Sized + Serialize>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<(), Self::Error> {
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: ?Sized + Serialize>(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        value: &T,
    ) -> Result<(), Self::Error> {
        value.serialize(self)
    }

    fn serialize_seq(self, _: Option<usize>) -> Result<Self, Self::Error> {
        Ok(self)
    }

    fn serialize_tuple(self, _: usize) -> Result<Self, Self::Error> {
        Ok(self)
    }

    fn serialize_tuple_struct(self, _: &'static str, _: usize) -> Result<Self, Self::Error> {
        Ok(self)
    }

    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self, Self::Error> {
        Ok(self)
    }

    fn serialize_map(self, _: Option<usize>) -> Result<Self, Self::Error> {
        Ok(self)
    }

    fn serialize_struct(self, _: &'static str, _: usize) -> Result<Self, Self::Error> {
        Ok(self)
    }

    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self, Self::Error> {
        Ok(self)
    }

    fn collect_str<T: ?Sized + std::fmt::Display>(self, _: &T) -> Result<(), Self::Error> {
        Ok(())
    }
}

macro_rules! float_sequence {
    ($trait:ident, $method:ident) => {
        impl ser::$trait for &mut FloatBits {
            type Ok = ();
            type Error = serde::de::value::Error;

            fn $method<T: ?Sized + Serialize>(&mut self, value: &T) -> Result<(), Self::Error> {
                value.serialize(&mut **self)
            }

            fn end(self) -> Result<(), Self::Error> {
                Ok(())
            }
        }
    };
}

float_sequence!(SerializeSeq, serialize_element);
float_sequence!(SerializeTuple, serialize_element);
float_sequence!(SerializeTupleStruct, serialize_field);
float_sequence!(SerializeTupleVariant, serialize_field);

impl ser::SerializeMap for &mut FloatBits {
    type Ok = ();
    type Error = serde::de::value::Error;

    fn serialize_key<T: ?Sized + Serialize>(&mut self, key: &T) -> Result<(), Self::Error> {
        key.serialize(&mut **self)
    }

    fn serialize_value<T: ?Sized + Serialize>(&mut self, value: &T) -> Result<(), Self::Error> {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Self::Error> {
        Ok(())
    }
}

macro_rules! float_struct {
    ($trait:ident) => {
        impl ser::$trait for &mut FloatBits {
            type Ok = ();
            type Error = serde::de::value::Error;

            fn serialize_field<T: ?Sized + Serialize>(
                &mut self,
                _: &'static str,
                value: &T,
            ) -> Result<(), Self::Error> {
                value.serialize(&mut **self)
            }

            fn end(self) -> Result<(), Self::Error> {
                Ok(())
            }
        }
    };
}

float_struct!(SerializeStruct);
float_struct!(SerializeStructVariant);

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
    use yrs::{Map, ReadTxn, Transact};

    const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

    fn key(
        renderer: &mut PptxRenderer,
        session: &pptx_edit::DeckSession,
        scope: &SlideScope,
    ) -> String {
        renderer.sync_package(session);
        renderer.slide_key(
            &Rc::new(session.version()),
            &scope.slide,
            scope.index,
            scope.width_emu,
            scope.height_emu,
        )
    }

    fn layout_input_oracle(scope: &SlideScope, generation: u64, font_epoch: u64) -> String {
        serde_json::to_string(&(scope, generation, font_epoch)).unwrap()
    }

    fn assert_key_oracles(
        renderer: &mut PptxRenderer,
        document: &PptxDocument,
        seen: &mut HashMap<(u64, String), String>,
        previous: &mut HashMap<String, (String, String)>,
    ) {
        let session = document.session();
        for index in 0..session.slide_ids().unwrap().len() {
            let scope = session.slide_scope(index).unwrap();
            let token = renderer.slide_layout_key(document, index as u32).unwrap();
            let oracle = layout_input_oracle(&scope, renderer.generation, renderer.font_epoch);
            if let Some(known) = seen.insert((renderer.generation, token.clone()), oracle.clone()) {
                assert_eq!(known, oracle, "{}", scope.slide.id);
            }
            if let Some((previous_token, previous_oracle)) =
                previous.insert(scope.slide.id.clone(), (token.clone(), oracle.clone()))
            {
                assert_eq!(
                    previous_token == token,
                    previous_oracle == oracle,
                    "{}",
                    scope.slide.id
                );
            }
        }
    }

    #[test]
    fn public_layout_prunes_keys_during_slide_churn() {
        let document = PptxDocument::open_collaborative(DECK, 913.0).unwrap();
        let session = document.session();
        let context = EditCtx::local("test");
        let mut renderer = PptxRenderer::new();
        renderer
            .register_fallback_font("Fallback", false, false, FONT)
            .unwrap();
        renderer.register_font("Test", false, false, FONT).unwrap();
        let live = session.slide_ids().unwrap().len();
        let bound = 256.max(2 * (live + 1));
        renderer.layout_slide_json(&document, 0).unwrap();
        let original = renderer.slide_layout_key(&document, 0).unwrap();
        for n in 0..1_000 {
            let receipt = session.insert_slide(&context, live as u32, None).unwrap();
            renderer.layout_slide_json(&document, live as u32).unwrap();
            assert!(renderer.layout_keys.len() <= bound);
            assert!(renderer.rendered.len() <= LAYOUT_CACHE_CAPACITY);
            let record = &renderer.layout_keys[&receipt.slide_id];
            assert_eq!(record.validated_version.as_ref(), &session.version());
            assert_eq!(record.font_epoch, renderer.font_epoch);
            assert_eq!(record.generation, renderer.generation);
            session.delete_slide(&context, &receipt.slide_id).unwrap();
            renderer.layout_slide_json(&document, 0).unwrap();
            assert!(renderer.layout_keys.len() <= bound);
            assert!(renderer.rendered.len() <= LAYOUT_CACHE_CAPACITY);
            assert_eq!(renderer.slide_layout_key(&document, 0).unwrap(), original);
            if n % 300 == 0 {
                let json = renderer.snapshot_with_layout_keys_json(&document).unwrap();
                let snapshot: serde_json::Value = serde_json::from_str(&json).unwrap();
                assert_eq!(
                    snapshot["snapshot"]["slides"].as_array().unwrap().len(),
                    live
                );
                assert_eq!(snapshot["keys"].as_object().unwrap().len(), live);
                assert!(!renderer.layout_keys.contains_key(&receipt.slide_id));
                assert!(!renderer.rendered.contains_key(&receipt.slide_id));
                assert_eq!(renderer.layout_keys.len(), live);
                assert!(renderer.layout_keys.values().all(|record| {
                    record.validated_version.as_ref() == &session.version()
                        && record.font_epoch == renderer.font_epoch
                        && record.generation == renderer.generation
                }));
            }
        }
        renderer.snapshot_with_layout_keys_json(&document).unwrap();
        assert_eq!(renderer.layout_keys.len(), live);
        let ids = session.slide_ids().unwrap();
        assert_eq!(ids.len(), live);
        assert!(renderer.layout_keys.keys().all(|id| ids.contains(id)));
        assert!(renderer.rendered.keys().all(|id| ids.contains(id)));
    }

    #[test]
    fn public_key_lookup_caps_entries_and_prunes_old_font_epochs() {
        let document = PptxDocument::open_collaborative(DECK, 914.0).unwrap();
        let session = document.session();
        let context = EditCtx::local("test");
        let mut renderer = PptxRenderer::new();
        let live = session.slide_ids().unwrap().len();
        let count = 258;
        for index in live..count {
            session.insert_slide(&context, index as u32, None).unwrap();
        }
        let mut keys = Vec::new();
        for index in 0..count {
            keys.push(renderer.slide_layout_key(&document, index as u32).unwrap());
            assert!(renderer.layout_keys.len() <= 2 * count);
        }
        assert_eq!(renderer.layout_keys.len(), count);
        for _ in 0..600 {
            let receipt = session.insert_slide(&context, count as u32, None).unwrap();
            renderer.slide_layout_key(&document, count as u32).unwrap();
            assert!(renderer.layout_keys.len() <= 2 * (count + 1));
            session.delete_slide(&context, &receipt.slide_id).unwrap();
            assert_eq!(renderer.slide_layout_key(&document, 0).unwrap(), keys[0]);
            assert!(renderer.layout_keys.len() <= 2 * count);
        }
        for (index, expected) in keys.iter().enumerate() {
            assert_eq!(
                &renderer.slide_layout_key(&document, index as u32).unwrap(),
                expected
            );
        }
        let index = (count - 1) as u32;
        let previous = renderer.slide_layout_key(&document, index).unwrap();
        let computed = renderer.key_count;
        assert_eq!(
            renderer.slide_layout_key(&document, index).unwrap(),
            previous
        );
        assert_eq!(renderer.key_count, computed);
        renderer
            .register_fallback_font("Fallback", false, false, FONT)
            .unwrap();
        assert_ne!(
            renderer.slide_layout_key(&document, index).unwrap(),
            previous
        );
        assert_eq!(renderer.layout_keys.len(), 1);
        assert_eq!(renderer.key_count, computed + 1);
        assert!(renderer.layout_keys.values().all(|record| {
            record.validated_version.as_ref() == &session.version()
                && record.font_epoch == renderer.font_epoch
                && record.generation == renderer.generation
        }));
    }

    #[test]
    fn large_deck_snapshot_keys_survive_every_layout() {
        let document = PptxDocument::open_collaborative(DECK, 919.0).unwrap();
        let session = document.session();
        let context = EditCtx::local("test");
        for index in session.slide_ids().unwrap().len()..260 {
            session.insert_slide(&context, index as u32, None).unwrap();
        }
        let mut renderer = PptxRenderer::new();
        renderer.register_font("Test", false, false, FONT).unwrap();
        let first: serde_json::Value =
            serde_json::from_str(&renderer.snapshot_with_layout_keys_json(&document).unwrap())
                .unwrap();
        let second: serde_json::Value =
            serde_json::from_str(&renderer.snapshot_with_layout_keys_json(&document).unwrap())
                .unwrap();
        assert_eq!(first["keys"], second["keys"]);
        assert_eq!(renderer.layout_keys.len(), 260);
        let computed = renderer.key_count;
        for (index, id) in session.slide_ids().unwrap().iter().enumerate() {
            renderer.layout_slide_json(&document, index as u32).unwrap();
            let expected = first["keys"][id].as_str().unwrap();
            assert_eq!(renderer.rendered[id].key, expected);
            assert_eq!(
                renderer.slide_layout_key(&document, index as u32).unwrap(),
                expected
            );
        }
        assert_eq!(renderer.key_count, computed);
        assert_eq!(renderer.layout_keys.len(), 260);
    }

    #[test]
    fn new_document_resets_pruning_after_large_deck_snapshot() {
        let large = PptxDocument::open_collaborative(DECK, 922.0).unwrap();
        let context = EditCtx::local("test");
        for index in large.session().slide_ids().unwrap().len()..260 {
            large
                .session()
                .insert_slide(&context, index as u32, None)
                .unwrap();
        }
        let mut renderer = PptxRenderer::new();
        renderer.snapshot_with_layout_keys_json(&large).unwrap();
        assert_eq!(renderer.layout_keys.len(), 260);

        let document = PptxDocument::open_collaborative(DECK, 923.0).unwrap();
        let session = document.session();
        let live = session.slide_ids().unwrap().len();
        let original = renderer.slide_layout_key(&document, 0).unwrap();
        assert_eq!(renderer.layout_keys.len(), 1);
        for _ in 0..520 {
            let receipt = session.insert_slide(&context, live as u32, None).unwrap();
            renderer.slide_layout_key(&document, live as u32).unwrap();
            assert!(renderer.layout_keys.len() <= 256.max(2 * (live + 1)));
            session.delete_slide(&context, &receipt.slide_id).unwrap();
            assert_eq!(renderer.slide_layout_key(&document, 0).unwrap(), original);
            assert!(renderer.layout_keys.len() <= 256.max(2 * live));
        }
    }

    #[test]
    fn signed_zero_changes_tokens_and_cached_layouts() {
        let document = PptxDocument::open_collaborative(DECK, 920.0).unwrap();
        let session = document.session();
        let context = EditCtx::local("test");
        let slide_id = session.slide_ids().unwrap()[0].clone();
        let shape = session
            .add_shape(
                &context,
                &slide_id,
                &pptx_edit::PresetShapeDraft {
                    name: "Callout".to_owned(),
                    geometry: "wedgeEllipseCallout".to_owned(),
                    rect: pptx_edit::ShapeRect {
                        x: 0,
                        y: 0,
                        width: 1_000_000,
                        height: 1_000_000,
                    },
                    fill: Some("#123456".to_owned()),
                },
            )
            .unwrap();
        let mut adjustments = BTreeMap::from([("adj1".to_owned(), 0.0), ("adj2".to_owned(), 0.0)]);
        session
            .set_shape_adjust(&context, &slide_id, &shape.shape_id, &adjustments)
            .unwrap();
        let positive = session.slide_scope(0).unwrap();
        let mut renderer = PptxRenderer::new();
        renderer.register_font("Test", false, false, FONT).unwrap();
        let original_frame = renderer.layout_slide_json(&document, 0).unwrap();
        let original = renderer.slide_layout_key(&document, 0).unwrap();
        adjustments.insert("adj1".to_owned(), -0.0);
        session
            .set_shape_adjust(&context, &slide_id, &shape.shape_id, &adjustments)
            .unwrap();
        let negative = session.slide_scope(0).unwrap();
        assert_eq!(positive, negative);
        assert_ne!(
            layout_input_oracle(&positive, renderer.generation, renderer.font_epoch),
            layout_input_oracle(&negative, renderer.generation, renderer.font_epoch)
        );
        let adjustment = negative.slide.shapes.last().unwrap().adjust_values["adj1"];
        assert_eq!(adjustment.to_bits(), (-0.0_f64).to_bits());
        assert_ne!(renderer.slide_layout_key(&document, 0).unwrap(), original);
        let frame = renderer.layout_slide_json(&document, 0).unwrap();
        assert_ne!(frame, original_frame);
        assert_eq!(renderer.layout_count, 2);
        let fresh = renderer
            .renderer
            .layout_scoped_slide(session.package(), &negative)
            .unwrap();
        assert_eq!(frame, serde_json::to_string(&fresh.display_list).unwrap());
    }

    #[test]
    fn invalidated_frames_cannot_be_activated() {
        let document = PptxDocument::open_collaborative(DECK, 921.0).unwrap();
        let session = document.session();
        let scope = session.slide_scope(0).unwrap();
        let mut renderer = PptxRenderer::new();
        renderer.register_font("Test", false, false, FONT).unwrap();
        for fallback in [false, true] {
            renderer.layout_slide_json(&document, 0).unwrap();
            let old = renderer.slide_layout_key(&document, 0).unwrap();
            assert!(renderer.set_active_slide(&document, &scope.slide.id, &old));
            let frame = &renderer.rendered[&scope.slide.id].rendered;
            let point = (0..frame.display_list.height as usize)
                .step_by(32)
                .flat_map(|y| {
                    (0..frame.display_list.width as usize)
                        .step_by(32)
                        .map(move |x| (x as f32, y as f32))
                })
                .find(|(x, y)| frame.hit_test(*x, *y).is_some())
                .unwrap();
            assert_ne!(renderer.hit_test_json(point.0, point.1).unwrap(), "null");
            if fallback {
                renderer
                    .register_fallback_font("Fallback", false, false, FONT)
                    .unwrap();
            } else {
                renderer.register_font("Test", false, false, FONT).unwrap();
            }
            assert!(renderer.rendered.is_empty());
            assert!(renderer.last_slide.is_none());
            assert!(!renderer.set_active_slide(&document, &scope.slide.id, &old));
            assert_eq!(renderer.hit_test_json(point.0, point.1).unwrap(), "null");
            let layouts = renderer.layout_count;
            renderer
                .hit_test_slide_json(&document, &scope.slide.id, point.0, point.1)
                .unwrap();
            assert_eq!(renderer.layout_count, layouts + 1);
        }
        let old = renderer.slide_layout_key(&document, 0).unwrap();
        assert!(renderer.set_active_slide(&document, &scope.slide.id, &old));
        let frame = &renderer.rendered[&scope.slide.id].rendered;
        let point = (0..frame.display_list.height as usize)
            .step_by(32)
            .flat_map(|y| {
                (0..frame.display_list.width as usize)
                    .step_by(32)
                    .map(move |x| (x as f32, y as f32))
            })
            .find(|(x, y)| frame.hit_test(*x, *y).is_some())
            .unwrap();
        assert_ne!(renderer.hit_test_json(point.0, point.1).unwrap(), "null");
        let story = scope
            .slide
            .shapes
            .iter()
            .flat_map(|shape| &shape.text_stories)
            .next()
            .unwrap();
        session
            .insert_text(
                &EditCtx::local("test"),
                &story.id,
                0,
                "X",
                &TextStyle::default(),
            )
            .unwrap();
        assert!(!renderer.set_active_slide(&document, &scope.slide.id, &old));
        assert_eq!(renderer.hit_test_json(point.0, point.1).unwrap(), "null");
        let current = renderer.slide_layout_key(&document, 0).unwrap();
        assert_ne!(old, current);
        assert!(!renderer.set_active_slide(&document, &scope.slide.id, &old));
        assert!(!renderer.set_active_slide(&document, &scope.slide.id, &current));
        renderer.layout_slide_json(&document, 0).unwrap();
        assert!(renderer.set_active_slide(&document, &scope.slide.id, &current));
        let layouts = renderer.layout_count;
        let index = session.slide_ids().unwrap().len() as u32;
        session
            .insert_slide(&EditCtx::local("test"), index, None)
            .unwrap();
        assert_eq!(renderer.slide_layout_key(&document, 0).unwrap(), current);
        assert!(renderer.set_active_slide(&document, &scope.slide.id, &current));
        assert_eq!(renderer.layout_count, layouts);
    }

    #[test]
    fn tokens_match_serialized_inputs_across_edit_history() {
        let document = PptxDocument::open_collaborative(DECK, 924.0).unwrap();
        let session = document.session();
        let context = EditCtx::local("test");
        let mut renderer = PptxRenderer::new();
        renderer.register_font("Test", false, false, FONT).unwrap();
        let mut seen = HashMap::new();
        let mut previous = HashMap::new();
        assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);
        let stories: Vec<_> = (0..2)
            .map(|index| {
                session
                    .slide_scope(index)
                    .unwrap()
                    .slide
                    .shapes
                    .iter()
                    .flat_map(|shape| &shape.text_stories)
                    .next()
                    .unwrap()
                    .id
                    .clone()
            })
            .collect();
        session
            .insert_text(&context, &stories[0], 0, "Edited ", &TextStyle::default())
            .unwrap();
        assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);

        for index in [1, 0] {
            let scope = session.slide_scope(index).unwrap();
            renderer.layout_slide_json(&document, index as u32).unwrap();
            let token = renderer.slide_layout_key(&document, index as u32).unwrap();
            assert!(renderer.set_active_slide(&document, &scope.slide.id, &token));
            assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);
        }
        assert!(session.undo());
        assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);
        assert!(session.redo());
        assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);
        session
            .insert_text(&context, &stories[1], 0, "Other ", &TextStyle::default())
            .unwrap();
        assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);
        renderer.register_font("Test", false, false, FONT).unwrap();
        assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);
        renderer
            .register_fallback_font("Fallback", false, false, FONT)
            .unwrap();
        assert_key_oracles(&mut renderer, &document, &mut seen, &mut previous);
    }

    #[test]
    fn memoized_keys_match_unmemoized_keys_before_and_after_edit() {
        let document = PptxDocument::open_collaborative(DECK, 912.0).unwrap();
        let session = document.session();
        let mut renderer = PptxRenderer::new();
        let mut fresh = PptxRenderer::new();
        for item in [&mut renderer, &mut fresh] {
            item.register_font("Liberation Sans", false, false, FONT)
                .unwrap();
        }
        let original = session.package().clone();
        let mut before = Vec::new();
        for index in 0..session.slide_ids().unwrap().len() {
            let scope = session.slide_scope(index).unwrap();
            let memoized = renderer.slide_layout_key(&document, index as u32).unwrap();
            let uncached = key(&mut fresh, session, &scope);
            assert_eq!(
                renderer.slide_layout_key(&document, index as u32).unwrap(),
                memoized
            );
            assert_eq!(key(&mut fresh, session, &scope), uncached);
            before.push((memoized, uncached));
        }
        let scope = session.slide_scope(0).unwrap();
        renderer.set_active_slide(&document, &scope.slide.id, &before[0].0);
        let story = scope
            .slide
            .shapes
            .iter()
            .flat_map(|shape| &shape.text_stories)
            .next()
            .unwrap();
        session
            .insert_text(
                &EditCtx::local("test"),
                &story.id,
                0,
                "X",
                &TextStyle::default(),
            )
            .unwrap();
        let keys = renderer.key_count;
        let changed = renderer.slide_layout_key(&document, 0).unwrap();
        renderer.layout_slide_json(&document, 0).unwrap();
        assert_eq!(renderer.key_count, keys + 1);
        assert_eq!(renderer.layout_count, 1);
        assert_ne!(before[0].0, changed);
        for (index, (previous, fresh_previous)) in before.iter().enumerate() {
            let scope = session.slide_scope(index).unwrap();
            let memoized = renderer.slide_layout_key(&document, index as u32).unwrap();
            let uncached = key(&mut fresh, session, &scope);
            assert_eq!(key(&mut fresh, session, &scope), uncached);
            if index == 0 {
                assert_ne!(&memoized, previous);
                assert_ne!(&uncached, fresh_previous);
            } else {
                assert_eq!(&memoized, previous);
                assert_eq!(&uncached, fresh_previous);
            }
            let expected = fresh
                .renderer
                .layout_scoped_slide(session.package(), &scope)
                .unwrap();
            let json = renderer.layout_slide_json(&document, index as u32).unwrap();
            assert_eq!(json, serde_json::to_string(&expected.display_list).unwrap());
            let layouts = renderer.layout_count;
            assert_eq!(
                renderer.layout_slide_json(&document, index as u32).unwrap(),
                json
            );
            assert_eq!(renderer.layout_count, layouts);
        }
        assert_eq!(renderer.key_count, keys + 1);
        assert_eq!(session.package(), &original);
    }

    #[test]
    fn keys_track_local_text_and_every_inherited_input() {
        let document = PptxDocument::open_collaborative(DECK, 910.0).unwrap();
        let session = document.session();
        let scope = session.slide_scope(0).unwrap();
        let package = session.package();
        let mut renderer = PptxRenderer::new();
        let original = key(&mut renderer, session, &scope);
        assert_eq!(key(&mut renderer, session, &scope), original);
        let other = session.slide_scope(1).unwrap();
        let other_key = key(&mut renderer, session, &other);
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
        let edited = renderer.slide_layout_key(&document, 0).unwrap();
        assert_ne!(original, edited);
        assert_eq!(renderer.slide_layout_key(&document, 1).unwrap(), other_key);
        session
            .move_slide(&EditCtx::local("test"), &scope.slide.id, 1)
            .unwrap();
        let moved = renderer.slide_layout_key(&document, 1).unwrap();
        assert_ne!(edited, moved);
        let mut previous = moved;
        for field in ["widthEmu", "heightEmu"] {
            let snapshot = session.snapshot().unwrap();
            let value = if field == "widthEmu" {
                snapshot.width_emu
            } else {
                snapshot.height_emu
            };
            {
                let mut txn = session.yrs_doc().transact_mut();
                txn.get_map("pptx:meta")
                    .unwrap()
                    .insert(&mut txn, field, (value + 1) as f64);
            }
            let resized = renderer.slide_layout_key(&document, 1).unwrap();
            assert_ne!(previous, resized);
            previous = resized;
        }
        renderer
            .register_font("Liberation Sans", false, false, FONT)
            .unwrap();
        let regular = renderer.slide_layout_key(&document, 1).unwrap();
        assert_ne!(previous, regular);
        renderer
            .register_fallback_font("Fallback", false, false, FONT)
            .unwrap();
        assert_ne!(regular, renderer.slide_layout_key(&document, 1).unwrap());

        let parents = pptx_edit::paragraph::SlideParents::resolve(
            package,
            scope.slide.source_part_path.as_deref(),
            scope.slide.layout_part_path.as_deref(),
        );
        let mut packages = Vec::new();
        let mut changed = package.clone();
        let source = parents.slide.unwrap();
        changed
            .slides
            .iter_mut()
            .find(|item| item.part_path == source.part_path)
            .unwrap()
            .show_master_shapes ^= true;
        packages.push(changed);
        let mut changed = package.clone();
        let layout = parents.layout.unwrap();
        changed
            .layouts
            .iter_mut()
            .find(|item| item.part_path == layout.part_path)
            .unwrap()
            .show_master_shapes ^= true;
        packages.push(changed);
        let mut changed = package.clone();
        let master = parents.master.unwrap();
        changed
            .masters
            .iter_mut()
            .find(|item| item.part_path == master.part_path)
            .unwrap()
            .color_map
            .set("bg1", "accent1");
        packages.push(changed);
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
        packages.push(changed);
        let mut changed = package.clone();
        changed.presentation.first_slide_num += 1;
        packages.push(changed);
        let mut changed = package.clone();
        changed
            .presentation
            .default_text_style
            .push(Default::default());
        packages.push(changed);
        let mut changed = package.clone();
        let paragraph = changed
            .presentation
            .default_text_paragraph
            .get_or_insert_with(Default::default);
        paragraph.margin_left = Some(paragraph.margin_left.unwrap_or(0) + 1);
        packages.push(changed);
        let mut changed = package.clone();
        changed.media[0].bytes.push(0);
        packages.push(changed);
        let mut changed = package.clone();
        changed.charts.push(pptx_parse::ChartPart {
            part_path: "ppt/charts/extra.xml".to_owned(),
            theme_part_path: None,
            chart: Default::default(),
        });
        packages.push(changed);
        let mut changed = package.clone();
        changed.diagram_drawings.push(pptx_parse::DiagramDrawing {
            part_path: "ppt/diagrams/extra.xml".to_owned(),
            shapes: Vec::new(),
        });
        packages.push(changed);
        let mut changed = package.clone();
        changed.table_styles.default_style_id = Some("test-style".to_owned());
        packages.push(changed);
        let baseline_session =
            pptx_edit::DeckSession::from_package_with_source(package.clone(), DECK, 918).unwrap();
        let baseline_scope = baseline_session.slide_scope(0).unwrap();
        for changed in packages {
            assert_ne!(&changed, package);
            let baseline = key(&mut renderer, &baseline_session, &baseline_scope);
            let generation = renderer.generation;
            let fresh =
                pptx_edit::DeckSession::from_package_with_source(changed, DECK, 916).unwrap();
            let input = fresh.slide_scope(0).unwrap();
            let inherited = key(&mut renderer, &fresh, &input);
            assert_ne!(baseline, inherited);
            assert!(inherited.parse::<u64>().unwrap() > baseline.parse::<u64>().unwrap());
            assert_eq!(renderer.generation, generation + 1);
            assert_eq!(renderer.layout_keys.len(), 1);
            assert_eq!(key(&mut renderer, &fresh, &input), inherited);
        }
    }

    #[test]
    fn tokens_are_not_reused_after_edit_and_undo() {
        let document = PptxDocument::open_collaborative(DECK, 917.0).unwrap();
        let session = document.session();
        let mut renderer = PptxRenderer::new();
        renderer
            .register_font("Liberation Sans", false, false, FONT)
            .unwrap();
        let scope = session.slide_scope(0).unwrap();
        let original = renderer.slide_layout_key(&document, 0).unwrap();
        let original_frame = renderer.layout_slide_json(&document, 0).unwrap();
        assert!(renderer.set_active_slide(&document, &scope.slide.id, &original));
        let story = scope
            .slide
            .shapes
            .iter()
            .flat_map(|shape| &shape.text_stories)
            .next()
            .unwrap();
        session
            .insert_text(
                &EditCtx::local("test"),
                &story.id,
                0,
                "Edited ",
                &TextStyle::default(),
            )
            .unwrap();
        let edited = renderer.slide_layout_key(&document, 0).unwrap();
        assert_ne!(original, edited);
        let edited_frame = renderer.layout_slide_json(&document, 0).unwrap();
        assert_ne!(original_frame, edited_frame);
        assert!(!renderer.set_active_slide(&document, &scope.slide.id, &original));
        assert!(session.undo());
        assert_eq!(session.slide_scope(0).unwrap(), scope);
        let restored = renderer.slide_layout_key(&document, 0).unwrap();
        assert_ne!(restored, edited);
        assert!(restored.parse::<u64>().unwrap() > edited.parse::<u64>().unwrap());
        assert!(!renderer.set_active_slide(&document, &scope.slide.id, &edited));
        let expected = renderer
            .renderer
            .layout_scoped_slide(session.package(), &scope)
            .unwrap();
        let restored_frame = renderer.layout_slide_json(&document, 0).unwrap();
        assert_eq!(restored_frame, original_frame);
        assert_eq!(
            restored_frame,
            serde_json::to_string(&expected.display_list).unwrap()
        );
        assert!(renderer.set_active_slide(&document, &scope.slide.id, &restored));
        assert_eq!(renderer.key_count, 3);
        assert_eq!(renderer.layout_count, 3);
        assert!(session.redo());
        let redone = renderer.slide_layout_key(&document, 0).unwrap();
        assert!(redone.parse::<u64>().unwrap() > restored.parse::<u64>().unwrap());
        let updated = session.slide_scope(0).unwrap();
        let expected = renderer
            .renderer
            .layout_scoped_slide(session.package(), &updated)
            .unwrap();
        assert_eq!(
            renderer.layout_slide_json(&document, 0).unwrap(),
            serde_json::to_string(&expected.display_list).unwrap()
        );
        assert_eq!(renderer.key_count, 4);
        assert_eq!(renderer.layout_count, 4);
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
        renderer.set_active_slide(&document, &scope.slide.id, "");
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
