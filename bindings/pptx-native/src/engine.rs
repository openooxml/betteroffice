use std::sync::OnceLock;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use napi::bindgen_prelude::{ArrayBuffer, JsObjectValue, Uint8Array, Uint8ArraySlice};
use napi::{Env, Error, Result, Status};
use napi_derive::napi;
use pptx_engine::{Document, EngineError, ErrorKind, Host, Renderer};

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

fn map_error(env: Env, error: EngineError) -> Error {
    let result = (|| -> Result<()> {
        let (name, code) = match error.kind {
            ErrorKind::Message => return env.throw(&error.message),
            ErrorKind::Peer { code } => ("PresentationPeerError", Some(code)),
            ErrorKind::Decode => ("PptxEditDecodeError", None),
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
}

#[napi]
pub struct PptxRenderer {
    inner: Option<Renderer>,
}

impl PptxDocument {
    fn document(&self) -> Result<&Document> {
        self.inner
            .as_ref()
            .ok_or_else(|| Error::from_reason("PptxDocument was freed"))
    }

    fn document_mut(&mut self) -> Result<&mut Document> {
        self.inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("PptxDocument was freed"))
    }
}

impl PptxRenderer {
    fn renderer(&self) -> Result<&Renderer> {
        self.inner
            .as_ref()
            .ok_or_else(|| Error::from_reason("PptxRenderer was freed"))
    }

    fn renderer_mut(&mut self) -> Result<&mut Renderer> {
        self.inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("PptxRenderer was freed"))
    }
}

impl Default for PptxRenderer {
    fn default() -> Self {
        Self {
            inner: Some(Renderer::new()),
        }
    }
}

#[napi]
impl PptxDocument {
    #[napi]
    pub fn free(&mut self) {
        self.inner = None;
    }

    #[napi(factory, js_name = "openReplayBaseline")]
    pub fn open_replay_baseline(
        env: Env,
        source: Uint8Array,
        client_id: Option<f64>,
        initial_update: Option<Uint8Array>,
    ) -> Result<Self> {
        let source = source.to_vec();
        let initial_update = initial_update.map(|bytes| bytes.to_vec());
        Document::open_replay_baseline(&source, client_id, initial_update, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi(factory, js_name = "openPeerDeckJson")]
    pub fn open_peer_deck_json(
        env: Env,
        source: Uint8Array,
        identity_json: String,
        initial_update: Option<Uint8Array>,
    ) -> Result<Self> {
        let source = source.to_vec();
        let initial_update = initial_update.map(|bytes| bytes.to_vec());
        Document::open_peer_deck_json(&source, &identity_json, initial_update, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi(factory, js_name = "openCollaborative")]
    pub fn open_collaborative(env: Env, bytes: Uint8Array, client_id: f64) -> Result<Self> {
        let bytes = bytes.to_vec();
        Document::open_collaborative(&bytes, client_id, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi(factory, js_name = "openCollaborativeFromUpdate")]
    pub fn open_collaborative_from_update(
        env: Env,
        update: Uint8Array,
        client_id: f64,
        source: Option<Uint8Array>,
    ) -> Result<Self> {
        let update = update.to_vec();
        let source = source.map(|bytes| bytes.to_vec());
        Document::open_collaborative_from_update(&update, client_id, source, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn peer_hydration_json(&self, env: Env, fonts_json: String) -> Result<String> {
        self.document()?
            .peer_hydration_json(&fonts_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn register_peer_fonts_json(&self, env: Env, fonts_json: String) -> Result<()> {
        self.document()?
            .register_peer_fonts_json(&fonts_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn adopt_peer_identity(&self, env: Env) -> Result<()> {
        self.document()?
            .adopt_peer_identity()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn replay_json(&self, env: Env, envelope: String) -> Result<String> {
        self.document()?
            .replay_json(&envelope)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn propose_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .propose_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn list_proposals_json(&self, env: Env) -> Result<String> {
        self.document()?
            .list_proposals_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn preview_proposal_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .preview_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn accept_proposal_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .accept_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn reject_proposal_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .reject_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi(getter, js_name = "clientId")]
    pub fn client_id(&self) -> Result<f64> {
        Ok(self.document()?.client_id())
    }

    #[napi]
    pub fn snapshot_json(&self, env: Env) -> Result<String> {
        self.document()?
            .snapshot_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn session_metadata_json(&self, env: Env) -> Result<String> {
        self.document()?
            .session_metadata_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn document_version(&self) -> Result<String> {
        Ok(self.document()?.document_version())
    }

    #[napi]
    pub fn read_content_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .read_content_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn read_command_content_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .read_command_content_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn find_text_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .find_text_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn validate_edits_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .validate_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn apply_edits_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .apply_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn export_structured_json(&self, env: Env, options: String) -> Result<String> {
        self.document()?
            .export_structured_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn export_markdown_json(&self, env: Env, options: String) -> Result<String> {
        self.document()?
            .export_markdown_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn search_text_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .search_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn story_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .story_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn story_formatting_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .story_formatting_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn anchor_caret_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .anchor_caret_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn resolve_caret_anchor_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .resolve_caret_anchor_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn media_bytes(&self, env: Env, part_path: String) -> Result<Uint8Array> {
        self.document()?
            .media_bytes(&part_path)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn save_bytes(&self, env: Env) -> Result<Uint8Array> {
        self.document()?
            .save_bytes()
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn encode_state_vector(&self, env: Env) -> Result<Uint8Array> {
        byte_result(env, self.document()?.encode_state_vector())
    }

    #[napi]
    pub fn encode_state_as_update(&self, env: Env) -> Result<Uint8Array> {
        byte_result(env, self.document()?.encode_state_as_update())
    }

    #[napi]
    pub fn encode_diff(&self, env: Env, remote_state_vector: Uint8Array) -> Result<Uint8Array> {
        let remote_state_vector = remote_state_vector.to_vec();
        self.document()?
            .encode_diff(&remote_state_vector)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn set_collaborative(&self, env: Env) -> Result<()> {
        self.document()?
            .set_collaborative()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn apply_update_json(&self, env: Env, update: Uint8Array) -> Result<String> {
        let update = update.to_vec();
        self.document()?
            .apply_update_json(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn start_update_observation(&mut self, env: Env) -> Result<()> {
        self.document_mut()?
            .start_update_observation()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn clear_update_observation(&mut self) -> Result<()> {
        self.document_mut()?.clear_update_observation();
        Ok(())
    }

    #[napi]
    pub fn drain_update_event(&self, env: Env) -> Result<Uint8Array> {
        byte_result(env, self.document()?.drain_update_event())
    }

    #[napi]
    pub fn insert_text_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .insert_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn insert_text_profiled_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .insert_text_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn delete_text_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .delete_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn delete_text_profiled_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .delete_text_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn format_text_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .format_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_paragraph_alignment_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_paragraph_alignment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn add_comment_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .add_comment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn reply_to_comment_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .reply_to_comment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_comment_position_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_comment_position_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_comment_status_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_comment_status_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn remove_comment_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .remove_comment_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_comment_flavor_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_comment_flavor_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn comments_json(&self, env: Env) -> Result<String> {
        self.document()?
            .comments_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn insert_paragraph_break_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .insert_paragraph_break_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_story_paragraphs_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_story_paragraphs_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn insert_table_row_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .insert_table_row_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn check_table_edit_json(&self, env: Env, args: String) -> Result<()> {
        self.document()?
            .check_table_edit_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn delete_table_row_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .delete_table_row_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn insert_table_column_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .insert_table_column_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn delete_table_column_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .delete_table_column_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn insert_slide_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .insert_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn duplicate_slide_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .duplicate_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn insert_slide_profiled_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .insert_slide_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn delete_slide_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .delete_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_slide_layout_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_slide_layout_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_slide_notes_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_slide_notes_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn move_slide_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .move_slide_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn add_text_box_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .add_text_box_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn add_text_box_profiled_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .add_text_box_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn add_shape_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .add_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn add_picture_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .add_picture_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn remove_shape_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .remove_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn bring_shape_to_front_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .bring_shape_to_front_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn send_shape_to_back_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .send_shape_to_back_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn bring_shape_forward_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .bring_shape_forward_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn send_shape_backward_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .send_shape_backward_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn move_shape_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .move_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn move_shape_profiled_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .move_shape_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn resize_shape_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .resize_shape_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_shape_rect_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_shape_rect_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_shape_fill_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_shape_fill_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_shape_stroke_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_shape_stroke_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_shape_adjust_json(&self, env: Env, args: String) -> Result<String> {
        self.document()?
            .set_shape_adjust_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn undo_capture_mode(&self) -> Result<String> {
        Ok(self.document()?.undo_capture_mode())
    }

    #[napi]
    pub fn set_undo_capture_mode(&self, env: Env, mode: String) -> Result<()> {
        self.document()?
            .set_undo_capture_mode(&mode)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn add_undo_boundary(&self) -> Result<()> {
        self.document()?.add_undo_boundary();
        Ok(())
    }

    #[napi]
    pub fn undo_json(&self, env: Env) -> Result<String> {
        self.document()?
            .undo_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn undo_receipt_json(&self, env: Env) -> Result<String> {
        self.document()?
            .undo_receipt_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn redo_receipt_json(&self, env: Env) -> Result<String> {
        self.document()?
            .redo_receipt_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn undo_profiled_json(&self, env: Env) -> Result<String> {
        self.document()?
            .undo_profiled_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn redo_json(&self, env: Env) -> Result<String> {
        self.document()?
            .redo_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn can_undo(&self) -> Result<bool> {
        Ok(self.document()?.can_undo())
    }

    #[napi]
    pub fn can_redo(&self) -> Result<bool> {
        Ok(self.document()?.can_redo())
    }

    #[napi]
    pub fn version() -> String {
        pptx_engine::version().to_owned()
    }
}

#[napi]
impl PptxRenderer {
    #[napi(constructor)]
    pub fn new() -> Self {
        Self::default()
    }

    #[napi]
    pub fn free(&mut self) {
        self.inner = None;
    }

    #[napi]
    pub fn register_font(
        &mut self,
        env: Env,
        family: String,
        bold: bool,
        italic: bool,
        bytes: Uint8Array,
    ) -> Result<u32> {
        let bytes = bytes.to_vec();
        self.renderer_mut()?
            .register_font(&family, bold, italic, &bytes)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn register_fallback_font(
        &mut self,
        env: Env,
        family: String,
        bold: bool,
        italic: bool,
        bytes: Uint8Array,
    ) -> Result<u32> {
        let bytes = bytes.to_vec();
        self.renderer_mut()?
            .register_fallback_font(&family, bold, italic, &bytes)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn save_bytes(&self, env: Env, document: &PptxDocument) -> Result<Uint8Array> {
        self.renderer()?
            .save_bytes(document.document()?.session())
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn layout_slide_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String> {
        self.renderer_mut()?
            .layout_slide_json(document.document()?.session(), slide_index)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn layout_slide_profiled_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String> {
        self.renderer_mut()?
            .layout_slide_profiled_json(
                document.document()?.session(),
                slide_index,
                &mut native_timer_ms,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn hit_test_json(&self, env: Env, x: f64, y: f64) -> Result<String> {
        self.renderer()?
            .hit_test_json(x as f32, y as f32)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_active_slide(
        &mut self,
        document: &PptxDocument,
        id: String,
        key: String,
    ) -> Result<bool> {
        Ok(self
            .renderer_mut()?
            .set_active_slide(document.document()?.session(), &id, &key))
    }

    #[napi]
    pub fn snapshot_with_layout_keys_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
    ) -> Result<String> {
        self.renderer_mut()?
            .snapshot_with_layout_keys_json(document.document()?.session())
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn slide_layout_key(
        &mut self,
        env: Env,
        document: &PptxDocument,
        slide_index: u32,
    ) -> Result<String> {
        self.renderer_mut()?
            .slide_layout_key(document.document()?.session(), slide_index)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn hit_test_slide_json(
        &mut self,
        env: Env,
        document: &PptxDocument,
        id: String,
        x: f64,
        y: f64,
    ) -> Result<String> {
        self.renderer_mut()?
            .hit_test_slide_json(document.document()?.session(), &id, x as f32, y as f32)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "requestedFontFacesJson")]
    pub fn requested_font_faces_json(
        &self,
        env: Env,
        document: &PptxDocument,
        slide_id: Option<String>,
    ) -> Result<String> {
        self.renderer()?
            .requested_font_faces_json(document.document()?.session(), slide_id.as_deref())
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn layout_proposal_slide_json(
        &self,
        env: Env,
        document: &PptxDocument,
        id: String,
        slide_index: u32,
    ) -> Result<String> {
        self.renderer()?
            .layout_proposal_slide_json(document.document()?.session(), &id, slide_index)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn layout_proposal_diff_slide_json(
        &self,
        env: Env,
        document: &PptxDocument,
        id: String,
        slide_index: u32,
    ) -> Result<String> {
        self.renderer()?
            .layout_proposal_diff_slide_json(document.document()?.session(), &id, slide_index)
            .map_err(|error| map_error(env, error))
    }
}

#[napi(js_name = "parsePptxJson")]
pub fn parse_pptx_json(env: Env, data: Uint8Array) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::parse_pptx_json(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parsePptxJsonWithoutMedia")]
pub fn parse_pptx_json_without_media(env: Env, data: Uint8Array) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::parse_pptx_json_without_media(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "exportPptxStructuredJson")]
pub fn export_pptx_structured_json(env: Env, data: Uint8Array, options: String) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::export_pptx_structured_json(&data, &options).map_err(|error| map_error(env, error))
}

#[napi(js_name = "exportPptxMarkdownJson")]
pub fn export_pptx_markdown_json(env: Env, data: Uint8Array, options: String) -> Result<String> {
    let data = data.to_vec();
    pptx_engine::export_pptx_markdown_json(&data, &options).map_err(|error| map_error(env, error))
}

#[napi(js_name = "renderPptxMarkdownJson")]
pub fn render_pptx_markdown_json(env: Env, content: String, options: String) -> Result<String> {
    pptx_engine::render_pptx_markdown_json(&content, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "compileSlideJson")]
pub fn compile_slide_json(env: Env, slide_json: String) -> Result<String> {
    pptx_engine::compile_slide_json(&slide_json).map_err(|error| map_error(env, error))
}

#[napi(js_name = "rendererVersion")]
pub fn renderer_version() -> String {
    pptx_engine::renderer_version().to_owned()
}

#[napi(js_name = "decodeTiffPng")]
pub fn decode_tiff_png(env: Env, data: Uint8Array) -> Result<Uint8Array> {
    let data = data.to_vec();
    pptx_engine::decode_tiff_png(&data)
        .map_err(|error| map_error(env, error))
        .and_then(|bytes| byte_result(env, bytes))
}

#[napi(js_name = "engineVersion")]
pub fn engine_version() -> String {
    env!("CARGO_PKG_VERSION").to_owned()
}
