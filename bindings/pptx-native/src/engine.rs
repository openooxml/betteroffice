use std::cell::{Cell, RefCell};
use std::sync::OnceLock;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use napi::bindgen_prelude::{
    ArrayBuffer, Either, JsObjectValue, Uint8Array, Uint8ArraySlice, Undefined,
};
use napi::{Env, Error, Result, Status};
use napi_derive::napi;
use pptx_engine::{Document, EngineError, ErrorKind, Host, Renderer};

#[napi(object)]
pub struct DecodedFontFace {
    pub font: Uint8Array,
    pub payload: Option<Uint8Array>,
}

/// Decodes one font pack face using this engine's linked decoder.
#[napi(
    catch_unwind,
    js_name = "decodeFontFace",
    ts_args_type = "bytes: Uint8Array, codec: string, dict?: Uint8Array | null, maxOutput?: number | null",
    ts_return_type = "DecodedFontFace"
)]
pub fn decode_font_face(
    env: Env,
    bytes: Uint8Array,
    codec: String,
    dict: Option<Uint8Array>,
    max_output: Option<f64>,
) -> Result<DecodedFontFace> {
    let decoded =
        betteroffice_fonts::decode_font_face_js(&bytes, &codec, dict.as_deref(), max_output)
            .map_err(|error| throw_error(&env, &error.to_string()))?;
    Ok(DecodedFontFace {
        font: Uint8Array::new(decoded.font),
        payload: decoded.payload.map(Uint8Array::new),
    })
}

thread_local! {
    static CURRENT_OBJECT: Cell<*const Failure> = const { Cell::new(std::ptr::null()) };
}

const FAILED_OBJECT: &str = "This engine object failed in an earlier call and can only be freed";

type Failure = RefCell<Option<String>>;

fn enter_object(failure: &Failure, env: Env) -> Result<()> {
    if let Some(message) = failure.borrow().as_deref() {
        return Err(throw_error(&env, message));
    }
    CURRENT_OBJECT.set(failure);
    Ok(())
}

struct EngineCall {
    env: Env,
    previous_object: *const Failure,
}

impl EngineCall {
    fn new(env: Env) -> Self {
        Self {
            env,
            previous_object: CURRENT_OBJECT.replace(std::ptr::null()),
        }
    }
}

impl Drop for EngineCall {
    fn drop(&mut self) {
        let object = CURRENT_OBJECT.replace(self.previous_object);
        if std::thread::panicking() {
            // SAFETY: the object registered itself from a method that is still on the stack.
            if !object.is_null()
                && let Ok(mut failure) = unsafe { &*object }.try_borrow_mut()
            {
                failure.get_or_insert_with(|| FAILED_OBJECT.to_owned());
            }
            let mut exception = std::ptr::null_mut();
            // SAFETY: the env belongs to the call that is ending on this thread.
            unsafe { napi::sys::napi_get_and_clear_last_exception(self.env.raw(), &mut exception) };
        }
    }
}

fn byte_result(env: Env, bytes: Vec<u8>) -> Result<Uint8Array> {
    let mut buffer = ArrayBuffer::copy_from(&env, &bytes)?;
    unsafe { buffer.as_mut() }.copy_from_slice(&bytes);
    Uint8ArraySlice::from_arraybuffer(&buffer, 0, bytes.len())?.into_typed_array(&env)
}

fn native_entropy() -> u64 {
    let mut bytes = [0; 8];
    if getrandom::fill(&mut bytes).is_ok() {
        return u64::from_le_bytes(bytes).max(1);
    }
    static NEXT: AtomicU64 = AtomicU64::new(1);
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    let seed = (nanos as u64)
        ^ ((nanos >> 64) as u64).rotate_left(17)
        ^ u64::from(std::process::id()).rotate_left(32)
        ^ NEXT.fetch_add(1, Ordering::Relaxed);
    let seed = (seed ^ (seed >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    let seed = (seed ^ (seed >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    (seed ^ (seed >> 31)).max(1)
}

fn native_timer_ms() -> f64 {
    static START: OnceLock<Instant> = OnceLock::new();
    START.get_or_init(Instant::now).elapsed().as_secs_f64() * 1_000.0
}

fn native_host() -> Host {
    Host {
        entropy: native_entropy,
        timer_ms: native_timer_ms,
    }
}

fn throw_error(env: &Env, message: &str) -> Error {
    let result = (|| -> Result<()> {
        let mut exception = env.create_error(Error::from_reason(message))?;
        exception.delete_named_property("code")?;
        env.throw(exception)
    })();
    match result {
        Ok(()) => Error::new(Status::PendingException, message.to_owned()),
        Err(error) => error,
    }
}

/// `{ slide }` and `{ scale, background, maxShadowPixels }` as the JS handle sends them.
fn png_request(
    target_json: &str,
    options_json: &str,
) -> std::result::Result<(u32, pptx_engine::PngOptions), String> {
    #[derive(serde::Deserialize)]
    #[serde(deny_unknown_fields)]
    struct Target {
        slide: u32,
    }
    #[derive(serde::Deserialize)]
    #[serde(deny_unknown_fields, rename_all = "camelCase")]
    struct Options {
        scale: Option<f64>,
        background: Option<String>,
        max_shadow_pixels: Option<u64>,
    }
    let target: Target = serde_json::from_str(target_json)
        .map_err(|_| "png-invalid: target must be { slide } with a slide index".to_owned())?;
    let options: Options = serde_json::from_str(options_json)
        .map_err(|error| format!("png-invalid: bad options: {error}"))?;
    let mut png = pptx_engine::PngOptions::default();
    if let Some(scale) = options.scale {
        if !scale.is_finite() || scale <= 0.0 {
            return Err("png-invalid: scale must be finite and positive".to_owned());
        }
        png.scale = scale as f32;
        if !png.scale.is_finite() {
            return Err(format!(
                "png-too-large: scale {scale} is past any slide size"
            ));
        }
    }
    if let Some(budget) = options.max_shadow_pixels {
        png.max_shadow_pixels = budget;
    }
    png.background = match options.background.as_deref() {
        None | Some("slide") => pptx_engine::PngBackground::Slide,
        Some("transparent") => pptx_engine::PngBackground::Transparent,
        Some(color) if color.starts_with('#') => {
            pptx_engine::PngBackground::Color(color.to_owned())
        }
        Some(other) => {
            return Err(format!(
                "png-invalid: background must be \"slide\", \"transparent\", or a #rrggbb color, not {other:?}"
            ));
        }
    };
    Ok((target.slide, png))
}

/// Prefixes a render failure with what kind it is, as the other native PNG methods do.
fn png_error(message: &str) -> String {
    let kind = if message.contains("no font has been registered") {
        "png-font-unavailable"
    } else if message.contains("px limit") {
        "png-too-large"
    } else {
        "png-invalid"
    };
    format!("{kind}: {message}")
}

fn map_error(env: Env, error: EngineError) -> Error {
    let result = (|| -> Result<()> {
        let (name, code) = match error.kind {
            ErrorKind::Message => return env.throw(&error.message),
            ErrorKind::Peer { code } => ("PresentationPeerError", Some(code)),
            ErrorKind::Decode => ("PptxEditDecodeError", None),
            ErrorKind::ResourceLimit => ("PptxResourceLimitError", None),
        };
        let mut exception = env.create_error(Error::from_reason(error.message.clone()))?;
        exception.set_named_property("name", name)?;
        if let Some(code) = code {
            exception.set_named_property("code", code)?;
        } else {
            exception.delete_named_property("code")?;
        }
        env.throw(exception)
    })();
    match result {
        Ok(()) => Error::new(Status::PendingException, error.message),
        Err(error) => error,
    }
}

#[napi]
pub struct PptxDocument {
    inner: Option<Document>,
    failure: Failure,
}

#[napi]
pub struct PptxRenderer {
    inner: Option<Renderer>,
    failure: Failure,
    #[cfg(feature = "test-hooks")]
    panic_next_png: Cell<bool>,
}

impl PptxDocument {
    fn document(&self, env: Env) -> Result<&Document> {
        enter_object(&self.failure, env)?;
        self.inner
            .as_ref()
            .ok_or_else(|| throw_error(&env, "PptxDocument was freed"))
    }

    fn document_mut(&mut self, env: Env) -> Result<&mut Document> {
        enter_object(&self.failure, env)?;
        self.inner
            .as_mut()
            .ok_or_else(|| throw_error(&env, "PptxDocument was freed"))
    }
}

impl PptxRenderer {
    fn renderer(&self, env: Env) -> Result<&Renderer> {
        enter_object(&self.failure, env)?;
        self.inner
            .as_ref()
            .ok_or_else(|| throw_error(&env, "PptxRenderer was freed"))
    }

    fn renderer_mut(&mut self, env: Env) -> Result<&mut Renderer> {
        enter_object(&self.failure, env)?;
        self.inner
            .as_mut()
            .ok_or_else(|| throw_error(&env, "PptxRenderer was freed"))
    }
}

impl Default for PptxRenderer {
    fn default() -> Self {
        Self {
            inner: Some(Renderer::new()),
            failure: RefCell::new(None),
            #[cfg(feature = "test-hooks")]
            panic_next_png: Cell::new(false),
        }
    }
}

#[cfg(feature = "test-hooks")]
#[napi]
impl PptxRenderer {
    #[napi(catch_unwind, js_name = "panicNextPngForTest")]
    pub fn panic_next_png_for_test(&self) {
        self.panic_next_png.set(true);
    }
}

#[cfg(feature = "test-hooks")]
#[napi]
impl PptxDocument {
    #[napi(catch_unwind, js_name = "panicForTest")]
    pub fn panic_for_test(&self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        let _document = self.document(env)?;
        panic!("panicForTest");
    }
}

#[napi]
impl PptxDocument {
    #[napi(catch_unwind)]
    pub fn free(&mut self, env: Env) {
        let _call = EngineCall::new(env);
        self.inner = None;
    }

    #[napi(catch_unwind, factory, js_name = "openReplayBaseline")]
    pub fn open_replay_baseline(
        env: Env,
        source: Uint8Array,
        client_id: Option<f64>,
        initial_update: Option<Uint8Array>,
    ) -> Result<Self> {
        let _call = EngineCall::new(env);
        let source = source.to_vec();
        let initial_update = initial_update.map(|bytes| bytes.to_vec());
        Document::open_replay_baseline(&source, client_id, initial_update, native_host())
            .map(|inner| Self {
                inner: Some(inner),
                failure: RefCell::new(None),
            })
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind, factory, js_name = "openPeerDeckJson")]
    pub fn open_peer_deck_json(
        env: Env,
        source: Uint8Array,
        identity_json: String,
        initial_update: Option<Uint8Array>,
    ) -> Result<Self> {
        let _call = EngineCall::new(env);
        let source = source.to_vec();
        let initial_update = initial_update.map(|bytes| bytes.to_vec());
        Document::open_peer_deck_json(&source, &identity_json, initial_update, native_host())
            .map(|inner| Self {
                inner: Some(inner),
                failure: RefCell::new(None),
            })
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind, factory, js_name = "openCollaborative")]
    pub fn open_collaborative(env: Env, bytes: Uint8Array, client_id: f64) -> Result<Self> {
        let _call = EngineCall::new(env);
        let bytes = bytes.to_vec();
        Document::open_collaborative(&bytes, client_id, native_host())
            .map(|inner| Self {
                inner: Some(inner),
                failure: RefCell::new(None),
            })
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind, factory, js_name = "openCollaborativeFromUpdate")]
    pub fn open_collaborative_from_update(
        env: Env,
        update: Uint8Array,
        client_id: f64,
        source: Option<Uint8Array>,
    ) -> Result<Self> {
        let _call = EngineCall::new(env);
        let update = update.to_vec();
        let source = source.map(|bytes| bytes.to_vec());
        Document::open_collaborative_from_update(&update, client_id, source, native_host())
            .map(|inner| Self {
                inner: Some(inner),
                failure: RefCell::new(None),
            })
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn peer_hydration_json(&self, env: Env, fonts_json: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .peer_hydration_json(&fonts_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn register_peer_fonts_json(&self, env: Env, fonts_json: String) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .register_peer_fonts_json(&fonts_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn adopt_peer_identity(&self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .adopt_peer_identity()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn replay_json(&self, env: Env, envelope: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .replay_json(&envelope)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn propose_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .propose_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn validate_proposal_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .validate_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn list_proposals_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .list_proposals_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn proposal_slide_strip_json(&self, env: Env, id: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .proposal_slide_strip_json(&id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn preview_proposal_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .preview_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn accept_proposal_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .accept_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn reject_proposal_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .reject_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind, getter, js_name = "clientId")]
    pub fn client_id(&self, env: Env) -> Result<f64> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.client_id())
    }

    #[napi(catch_unwind)]
    pub fn snapshot_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .snapshot_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn session_metadata_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .session_metadata_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn document_version(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.document_version())
    }

    #[napi(catch_unwind)]
    pub fn read_content_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .read_content_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn read_command_content_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .read_command_content_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn find_text_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .find_text_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn validate_edits_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .validate_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn apply_edits_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .apply_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn export_structured_json(&self, env: Env, options: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .export_structured_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn export_markdown_json(&self, env: Env, options: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .export_markdown_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn search_text_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .search_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn story_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .story_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn story_formatting_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .story_formatting_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn anchor_caret_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .anchor_caret_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn resolve_caret_anchor_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .resolve_caret_anchor_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn media_bytes(&self, env: Env, part_path: String) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .media_bytes(&part_path)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(catch_unwind)]
    pub fn metafile_svg(
        &self,
        env: Env,
        part_path: String,
    ) -> Result<Either<Uint8Array, Undefined>> {
        let _call = EngineCall::new(env);
        match self.document(env)?.session().metafile_svg(&part_path) {
            Some(svg) => byte_result(env, svg.to_vec()).map(Either::A),
            None => Ok(Either::B(())),
        }
    }

    #[napi(catch_unwind)]
    pub fn save_bytes(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .save_bytes()
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(catch_unwind)]
    pub fn encode_state_vector(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        byte_result(env, self.document(env)?.encode_state_vector())
    }

    #[napi(catch_unwind)]
    pub fn encode_state_as_update(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        byte_result(env, self.document(env)?.encode_state_as_update())
    }

    #[napi(catch_unwind)]
    pub fn encode_diff(&self, env: Env, remote_state_vector: Uint8Array) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        let remote_state_vector = remote_state_vector.to_vec();
        self.document(env)?
            .encode_diff(&remote_state_vector)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(catch_unwind)]
    pub fn set_collaborative(&self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_collaborative()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn apply_update_json(&self, env: Env, update: Uint8Array) -> Result<String> {
        let _call = EngineCall::new(env);
        let update = update.to_vec();
        self.document(env)?
            .apply_update_json(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn start_update_observation(&mut self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .start_update_observation()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn clear_update_observation(&mut self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?.clear_update_observation();
        Ok(())
    }

    #[napi(catch_unwind)]
    pub fn drain_update_event(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        byte_result(env, self.document(env)?.drain_update_event())
    }

    #[napi(catch_unwind)]
    pub fn insert_text_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .insert_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn insert_text_profiled_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .insert_text_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn delete_text_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .delete_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn delete_text_profiled_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .delete_text_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn format_text_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .format_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_paragraph_alignment_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_paragraph_alignment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn add_comment_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .add_comment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn reply_to_comment_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .reply_to_comment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_comment_position_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_comment_position_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_comment_status_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_comment_status_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn remove_comment_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .remove_comment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_comment_flavor_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_comment_flavor_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn comments_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .comments_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn insert_paragraph_break_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .insert_paragraph_break_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_story_paragraphs_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_story_paragraphs_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn insert_table_row_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .insert_table_row_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn check_table_edit_json(&self, env: Env, args: String) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .check_table_edit_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn delete_table_row_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .delete_table_row_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn insert_table_column_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .insert_table_column_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn delete_table_column_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .delete_table_column_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn insert_slide_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .insert_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn duplicate_slide_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .duplicate_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn insert_slide_profiled_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .insert_slide_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn delete_slide_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .delete_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_slide_layout_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_slide_layout_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_slide_notes_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_slide_notes_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn move_slide_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .move_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn add_text_box_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .add_text_box_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn add_text_box_profiled_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .add_text_box_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn add_shape_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .add_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn add_picture_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .add_picture_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn remove_shape_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .remove_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn bring_shape_to_front_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .bring_shape_to_front_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn send_shape_to_back_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .send_shape_to_back_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn bring_shape_forward_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .bring_shape_forward_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn send_shape_backward_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .send_shape_backward_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn move_shape_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .move_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn move_shape_profiled_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .move_shape_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn resize_shape_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .resize_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_shape_rect_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_shape_rect_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_shape_fill_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_shape_fill_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_shape_stroke_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_shape_stroke_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_shape_adjust_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_shape_adjust_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn undo_capture_mode(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.undo_capture_mode())
    }

    #[napi(catch_unwind)]
    pub fn set_undo_capture_mode(&self, env: Env, mode: String) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .set_undo_capture_mode(&mode)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn add_undo_boundary(&self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document(env)?.add_undo_boundary();
        Ok(())
    }

    #[napi(catch_unwind)]
    pub fn undo_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .undo_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn undo_receipt_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .undo_receipt_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn redo_receipt_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .redo_receipt_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn undo_profiled_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .undo_profiled_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn redo_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .redo_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn can_undo(&self, env: Env) -> Result<bool> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.can_undo())
    }

    #[napi(catch_unwind)]
    pub fn can_redo(&self, env: Env) -> Result<bool> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.can_redo())
    }

    #[napi(catch_unwind)]
    pub fn version(env: Env) -> String {
        let _call = EngineCall::new(env);
        pptx_engine::version().to_owned()
    }
}

#[napi]
impl PptxRenderer {
    #[napi(catch_unwind, constructor)]
    pub fn new(env: Env) -> Self {
        let _call = EngineCall::new(env);
        Self::default()
    }

    #[napi(catch_unwind)]
    pub fn free(&mut self, env: Env) {
        let _call = EngineCall::new(env);
        self.inner = None;
    }

    #[napi(catch_unwind)]
    pub fn register_font(
        &mut self,
        env: Env,
        family: String,
        bold: bool,
        italic: bool,
        bytes: Uint8Array,
    ) -> Result<u32> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let bytes = bytes.to_vec();
        renderer
            .register_font(&family, bold, italic, bytes)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn register_fallback_font(
        &mut self,
        env: Env,
        family: String,
        bold: bool,
        italic: bool,
        bytes: Uint8Array,
    ) -> Result<u32> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let bytes = bytes.to_vec();
        renderer
            .register_fallback_font(&family, bold, italic, bytes)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn save_bytes(&self, env: Env, document: &PptxDocument) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .save_bytes(document.session())
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(catch_unwind)]
    pub fn layout_slide_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .layout_slide_json(document.session(), slide_index)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn export_slide_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .export_slide_json(document.session(), slide_index)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn render_png_target(
        &mut self,
        env: Env,
        document: &PptxDocument,
        target_json: String,
        options_json: String,
    ) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        let (slide, options) = png_request(&target_json, &options_json)
            .map_err(|message| throw_error(&env, &message))?;
        #[cfg(feature = "test-hooks")]
        let panic_next = self.panic_next_png.replace(false);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        #[cfg(feature = "test-hooks")]
        if panic_next {
            panic!("panicNextPngForTest");
        }
        renderer
            .render_png(document.session(), slide, &options)
            .map_err(|error| match error.kind {
                ErrorKind::Message => throw_error(&env, &png_error(&error.message)),
                _ => map_error(env, error),
            })
            .and_then(|rendered| byte_result(env, rendered.bytes))
    }

    #[napi(catch_unwind)]
    pub fn layout_slide_profiled_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .layout_slide_profiled_json(document.session(), slide_index, &mut native_timer_ms)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn hit_test_json(&self, env: Env, x: f64, y: f64) -> Result<String> {
        let _call = EngineCall::new(env);
        self.renderer(env)?
            .hit_test_json(x as f32, y as f32)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn set_active_slide(
        &mut self,
        env: Env,
        document: &PptxDocument,
        id: String,
        key: String,
    ) -> Result<bool> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        Ok(renderer.set_active_slide(document.session(), &id, &key))
    }

    #[napi(catch_unwind)]
    pub fn snapshot_with_layout_keys_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .snapshot_with_layout_keys_json(document.session())
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn slide_layout_key(
        &mut self,
        env: Env,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .slide_layout_key(document.session(), slide_index)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn hit_test_slide_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
        id: String,
        x: f64,
        y: f64,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer_mut(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .hit_test_slide_json(document.session(), &id, x as f32, y as f32)
            .map_err(|error| map_error(env, error))
    }

    /// Enables browser metrics for uncovered characters. Off by default.
    #[napi(catch_unwind, js_name = "setBrowserCoverageFallback")]
    pub fn set_browser_coverage_fallback(&mut self, env: Env, enabled: bool) -> Result<()> {
        let _call = EngineCall::new(env);
        self.renderer_mut(env)?
            .set_browser_coverage_fallback(enabled);
        Ok(())
    }

    /// Requested faces with optional script coverage buckets, before layout.
    #[napi(catch_unwind, js_name = "requestedFontFacesJson")]
    pub fn requested_font_faces_json(
        &self,
        env: Env,
        document: &PptxDocument,
        slide_id: Option<String>,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .requested_font_faces_json(document.session(), slide_id.as_deref())
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn layout_proposal_slide_json(
        &self,
        env: Env,
        document: &PptxDocument,
        id: String,
        slide_index: u32,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .layout_proposal_slide_json(document.session(), &id, slide_index)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn layout_proposal_diff_slide_by_id_json(
        &self,
        env: Env,
        document: &PptxDocument,
        id: String,
        slide_id: String,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .layout_proposal_diff_slide_by_id_json(document.session(), &id, &slide_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(catch_unwind)]
    pub fn layout_proposal_diff_slide_json(
        &self,
        env: Env,
        document: &PptxDocument,
        id: String,
        slide_index: u32,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        let renderer = self.renderer(env)?;
        let _document_call = EngineCall::new(env);
        let document = document.document(env)?;
        renderer
            .layout_proposal_diff_slide_json(document.session(), &id, slide_index)
            .map_err(|error| map_error(env, error))
    }
}

#[napi(catch_unwind, js_name = "parsePptxJson")]
pub fn parse_pptx_json(env: Env, data: Uint8Array) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::parse_pptx_json(&data).map_err(|error| map_error(env, error))
}

#[napi(catch_unwind, js_name = "parsePptxJsonWithoutMedia")]
pub fn parse_pptx_json_without_media(env: Env, data: Uint8Array) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::parse_pptx_json_without_media(&data).map_err(|error| map_error(env, error))
}

#[napi(catch_unwind, js_name = "exportPptxStructuredJson")]
pub fn export_pptx_structured_json(env: Env, data: Uint8Array, options: String) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::export_pptx_structured_json(&data, &options).map_err(|error| map_error(env, error))
}

#[napi(catch_unwind, js_name = "exportPptxMarkdownJson")]
pub fn export_pptx_markdown_json(env: Env, data: Uint8Array, options: String) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::export_pptx_markdown_json(&data, &options).map_err(|error| map_error(env, error))
}

#[napi(catch_unwind, js_name = "renderPptxMarkdownJson")]
pub fn render_pptx_markdown_json(env: Env, content: String, options: String) -> Result<String> {
    pptx_engine::render_pptx_markdown_json(&content, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(catch_unwind, js_name = "compileSlideJson")]
pub fn compile_slide_json(env: Env, slide_json: String) -> Result<String> {
    pptx_engine::compile_slide_json(&slide_json).map_err(|error| map_error(env, error))
}

#[napi(catch_unwind, js_name = "rendererVersion")]
pub fn renderer_version() -> String {
    pptx_engine::renderer_version().to_owned()
}

#[napi(catch_unwind, js_name = "decodeTiffPng")]
pub fn decode_tiff_png(env: Env, data: Uint8Array) -> Result<Uint8Array> {
    let data = data.to_vec();
    pptx_engine::decode_tiff_png(&data)
        .map_err(|error| map_error(env, error))
        .and_then(|bytes| byte_result(env, bytes))
}

#[napi(catch_unwind, js_name = "engineVersion")]
pub fn engine_version() -> String {
    env!("CARGO_PKG_VERSION").to_owned()
}
