//! Interned media sources.
//!
//! A resolved image `src` is a `data:`/`blob:` URL the canvas image resolver
//! decodes, so until now every stage between the parser and the wire stored a
//! copy of it: the yrs embed payload, the layout input, the resident display
//! list, and every snapshot derived from those. On media-heavy documents that
//! is most of the session's retained heap.
//!
//! Seeding instead writes a `media:{n}` token anywhere a parser-minted URL
//! would appear and keeps one `MediaSrcTable` of interned URLs on the
//! [`EditingDoc`]. The frame encoder swaps the token back for the URL as it
//! serializes a page, so the wire carries exactly what it always did — the
//! strings only exist in the frame bytes and in one table.

use std::borrow::Cow;
use std::collections::HashMap;
use std::sync::Arc;

use docx_layout::display_list::{DisplayPage, MEDIA_SRC_PREFIX, Primitive};
use yrs::Any;

use crate::raw::RawOp;

const TOKEN_PREFIX: &str = MEDIA_SRC_PREFIX;

/// Payload keys a resolved image source travels under. `src` covers image
/// embeds and drawing-scene image nodes; `dataUrl`/`data_url` and `pictureSrc`
/// cover watermark and shape picture-fill sources.
const SRC_KEYS: &[&str] = &["src", "dataUrl", "data_url", "pictureSrc"];

/// One interned URL per distinct parser-minted media source in a seeded
/// document. `media:{n}` tokens index back into `srcs`.
#[derive(Default)]
pub(crate) struct MediaSrcTable {
    srcs: Vec<Arc<str>>,
    by_src: HashMap<String, u32>,
}

impl MediaSrcTable {
    /// The token for a parser-minted URL, or `None` for anything else —
    /// external and empty sources pass through untouched.
    fn intern(&mut self, src: &str) -> Option<Arc<str>> {
        if !src.starts_with("data:") && !src.starts_with("blob:") {
            return None;
        }
        let index = match self.by_src.get(src) {
            Some(&index) => index,
            None => {
                let index = u32::try_from(self.srcs.len()).ok()?;
                self.srcs.push(Arc::from(src));
                self.by_src.insert(src.to_owned(), index);
                index
            }
        };
        Some(format!("{TOKEN_PREFIX}{index}").into())
    }

    /// A table continuing an existing `srcs` list, for seeds that add to a
    /// document rather than open it — the tokens they emit must not renumber
    /// what earlier seeds interned.
    pub(crate) fn from_srcs(srcs: Arc<[Arc<str>]>) -> Self {
        let by_src = srcs
            .iter()
            .enumerate()
            .map(|(index, src)| (src.to_string(), index as u32))
            .collect();
        Self {
            srcs: srcs.to_vec(),
            by_src,
        }
    }

    pub(crate) fn into_srcs(self) -> Arc<[Arc<str>]> {
        self.srcs.into()
    }
}

/// `token` -> `Some(index)` when it names an interned media source.
fn token_index(value: &str) -> Option<usize> {
    value
        .strip_prefix(TOKEN_PREFIX)
        .and_then(|rest| rest.parse::<usize>().ok())
}

/// `media:{n}` token -> its interned URL, or `value` unchanged — the single
/// resolve the frame serializers apply to strings under media-source keys.
pub(crate) fn resolve_token<'a>(value: &'a str, srcs: &'a [Arc<str>]) -> &'a str {
    token_index(value)
        .and_then(|index| srcs.get(index))
        .map_or(value, |src| src.as_ref())
}

/// Interns every media URL inside `ops` payloads in place, so the yrs store
/// only ever sees tokens.
pub(crate) fn intern_ops_media(ops: &mut [RawOp], table: &mut MediaSrcTable) {
    for op in ops {
        match op {
            RawOp::InsertEmbed { payload, .. } => {
                for (key, value) in payload.iter_mut() {
                    if let Some(interned) = intern_any(value, key, table) {
                        *value = interned;
                    }
                }
            }
            RawOp::SetEmbedAttr { key, value, .. } => {
                if let Some(interned) = intern_any(value, key, table) {
                    *value = interned;
                }
            }
            RawOp::SetComment { body, .. } => {
                if let Some(interned) = intern_any(body, "", table) {
                    *body = interned;
                }
            }
            RawOp::Insert { .. }
            | RawOp::Delete { .. }
            | RawOp::Format { .. }
            | RawOp::RemoveComment { .. } => {}
        }
    }
}

/// Replaces an `Any` subtree's media URLs with tokens, returning `None` when
/// nothing under it changed.
fn intern_any(value: &Any, key: &str, table: &mut MediaSrcTable) -> Option<Any> {
    match value {
        Any::String(text) if SRC_KEYS.contains(&key) => table.intern(text).map(Any::String),
        Any::String(text) if embedded_media_json(text) => {
            patch_embedded_json(text, table).map(|patched| Any::String(patched.into()))
        }
        Any::Array(items) => {
            let mut changed = false;
            let next: Vec<Any> = items
                .iter()
                .map(|item| match intern_any(item, "", table) {
                    Some(interned) => {
                        changed = true;
                        interned
                    }
                    None => item.clone(),
                })
                .collect();
            changed.then(|| Any::Array(next.into()))
        }
        Any::Map(map) => {
            let mut changed = false;
            let next: HashMap<String, Any> = map
                .iter()
                .map(|(key, value)| {
                    let value = match intern_any(value, key, table) {
                        Some(interned) => {
                            changed = true;
                            interned
                        }
                        None => value.clone(),
                    };
                    (key.clone(), value)
                })
                .collect();
            changed.then(|| Any::Map(Arc::new(next)))
        }
        _ => None,
    }
}

/// Whether a string might be a serialized JSON tree holding media URLs.
fn embedded_media_json(text: &str) -> bool {
    matches!(text.as_bytes().first(), Some(b'{') | Some(b'['))
        && text.contains("\"data:")
        && SRC_KEYS
            .iter()
            .any(|key| text.contains(&format!("\"{key}\"")))
}

/// Reparses a serialized-JSON payload string and interns every media URL in
/// it, mirroring what [`intern_any`] does at the `Any` level.
fn patch_embedded_json(text: &str, table: &mut MediaSrcTable) -> Option<String> {
    let mut value: serde_json::Value = serde_json::from_str(text).ok()?;
    if !patch_json_media(&mut value, table) {
        return None;
    }
    serde_json::to_string(&value).ok()
}

/// Returns whether any `src`-keyed URL under `value` was interned.
fn patch_json_media(value: &mut serde_json::Value, table: &mut MediaSrcTable) -> bool {
    match value {
        serde_json::Value::Object(map) => {
            let mut changed = false;
            for (key, entry) in map.iter_mut() {
                if let serde_json::Value::String(text) = entry
                    && SRC_KEYS.contains(&key.as_str())
                    && let Some(token) = table.intern(text)
                {
                    *text = token.to_string();
                    changed = true;
                    continue;
                }
                changed |= patch_json_media(entry, table);
            }
            changed
        }
        serde_json::Value::Array(items) => {
            let mut changed = false;
            for item in items {
                changed |= patch_json_media(item, table);
            }
            changed
        }
        _ => false,
    }
}

/// Whether `page` carries a `media:` token — the gate that keeps non-media
/// pages borrowing instead of cloning. Payload Values flag themselves at
/// build time via [`DocAttrs::media_token`]; only the image `rel_id` needs a
/// string check here.
fn page_has_media_token(page: &DisplayPage) -> bool {
    fn prims_have_token(primitives: &[Primitive]) -> bool {
        primitives.iter().any(|primitive| {
            primitive_attrs(primitive).media_token
                || matches!(primitive, Primitive::Image(image)
                    if image.rel_id.starts_with(TOKEN_PREFIX))
        })
    }
    prims_have_token(&page.primitives)
        || page
            .header
            .iter()
            .chain(&page.footer)
            .any(|region| prims_have_token(&region.primitives))
        || page.note_areas.iter().any(|area| {
            prims_have_token(&area.primitives) || prims_have_token(&area.separator_primitives)
        })
}

/// The display page a frame ships, with `media:` tokens materialized back into
/// their URLs. Borrows the page unchanged when no primitive carries a token.
pub(crate) fn materialize_page_media<'a>(
    page: &'a DisplayPage,
    srcs: &[Arc<str>],
) -> Cow<'a, DisplayPage> {
    if srcs.is_empty() || !page_has_media_token(page) {
        return Cow::Borrowed(page);
    }
    let mut page = page.clone();
    let patch_primitives = |primitives: &mut Vec<Primitive>| {
        for primitive in primitives {
            if let Primitive::Image(image) = primitive {
                materialize_token(&mut image.rel_id, srcs);
            }
            let attrs = primitive_attrs_mut(primitive);
            if let Some(fill_paint) = &mut attrs.fill_paint {
                patch_value_tokens(fill_paint, srcs);
            }
            if let Some(drawing_scene) = &mut attrs.drawing_scene {
                patch_value_tokens(drawing_scene, srcs);
            }
        }
    };
    patch_primitives(&mut page.primitives);
    for region in page.header.iter_mut().chain(&mut page.footer) {
        patch_primitives(&mut region.primitives);
    }
    for area in &mut page.note_areas {
        patch_primitives(&mut area.primitives);
        patch_primitives(&mut area.separator_primitives);
    }
    Cow::Owned(page)
}

fn primitive_attrs(primitive: &Primitive) -> &docx_layout::display_list::DocAttrs {
    match primitive {
        Primitive::Text(value) => &value.attrs,
        Primitive::GlyphRun(value) => &value.attrs,
        Primitive::Rect(value) => &value.attrs,
        Primitive::Line(value) => &value.attrs,
        Primitive::Image(value) => &value.attrs,
        Primitive::Shape(value) => &value.attrs,
        Primitive::Decoration(value) => &value.attrs,
    }
}

fn primitive_attrs_mut(primitive: &mut Primitive) -> &mut docx_layout::display_list::DocAttrs {
    match primitive {
        Primitive::Text(value) => &mut value.attrs,
        Primitive::GlyphRun(value) => &mut value.attrs,
        Primitive::Rect(value) => &mut value.attrs,
        Primitive::Line(value) => &mut value.attrs,
        Primitive::Image(value) => &mut value.attrs,
        Primitive::Shape(value) => &mut value.attrs,
        Primitive::Decoration(value) => &mut value.attrs,
    }
}

/// Replaces a `media:` token in place with the URL it names; plain strings
/// (including URLs inserted post-seed, which were never interned) pass through.
fn materialize_token(value: &mut String, srcs: &[Arc<str>]) {
    if let Some(index) = token_index(value)
        && let Some(src) = srcs.get(index)
    {
        *value = src.to_string();
    }
}

/// Token pass for the untyped `Value` payloads primitives carry — fill paints
/// and drawing scenes address their images through the same `src`-keyed
/// fields as embeds.
fn patch_value_tokens(value: &mut serde_json::Value, srcs: &[Arc<str>]) {
    match value {
        serde_json::Value::Object(map) => {
            for (key, entry) in map.iter_mut() {
                if let serde_json::Value::String(text) = entry
                    && SRC_KEYS.contains(&key.as_str())
                {
                    materialize_token(text, srcs);
                    continue;
                }
                patch_value_tokens(entry, srcs);
            }
        }
        serde_json::Value::Array(items) => {
            for item in items {
                patch_value_tokens(item, srcs);
            }
        }
        _ => {}
    }
}
