use std::cell::Cell;
use std::sync::OnceLock;
use std::time::Instant;

use napi::bindgen_prelude::{
    ArrayBuffer, Either, Function, JsObjectValue, Uint8Array, Uint8ArraySlice, Undefined,
};
use napi::{Env, Error, JsValue, Result, Status, Unknown, ValueType};
use napi_derive::napi;
use xlsx_engine::{Document, EngineError, ErrorKind, Host, SnapshotBuilder};

#[napi(object)]
pub struct DecodedFontFace {
    pub font: Uint8Array,
    pub payload: Option<Uint8Array>,
}

/// Decodes one font pack face using this engine's linked decoder.
#[napi(
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
    static CURRENT_ENV: Cell<Option<Env>> = const { Cell::new(None) };
}

struct EngineCall {
    previous: Option<Env>,
}

impl EngineCall {
    fn new(env: Env) -> Self {
        Self {
            previous: CURRENT_ENV.replace(Some(env)),
        }
    }
}

impl Drop for EngineCall {
    fn drop(&mut self) {
        CURRENT_ENV.set(self.previous);
    }
}

fn native_progress() {
    CURRENT_ENV.with(|current| {
        let Some(env) = current.get() else {
            return;
        };
        let result = (|| -> Result<()> {
            let global = env.get_global()?;
            let hook: Unknown = global.get_named_property("__betterofficeEngineProgress")?;
            if hook.get_type()? == ValueType::Function {
                let hook: Function<(), Unknown> = unsafe { hook.cast()? };
                let _ = hook.call(());
            }
            Ok(())
        })();
        if result.is_err() {
            let mut exception = std::ptr::null_mut();
            unsafe { napi::sys::napi_get_and_clear_last_exception(env.raw(), &mut exception) };
        }
    });
}

fn native_now_serial() -> std::result::Result<Option<f64>, EngineError> {
    CURRENT_ENV.with(|current| {
        let Some(env) = current.get() else {
            return Ok(None);
        };
        let result = (|| -> Result<Option<f64>> {
            let date: Function<f64> = env.get_global()?.get_named_property("Date")?;
            let now: Function<(), f64> = date.get_named_property("now")?;
            let ms = now.apply(date, ())?;
            let instance = date.new_instance(ms)?.coerce_to_object()?;
            let timezone_offset: Function<(), f64> =
                instance.get_named_property("getTimezoneOffset")?;
            let offset = timezone_offset.apply(instance, ())?;
            Ok(Some((ms - offset * 60_000.0) / 86_400_000.0 + 25_569.0))
        })();
        result.map_err(|error| {
            if error.status == Status::PendingException {
                EngineError::pending_exception()
            } else {
                EngineError::error(error.reason)
            }
        })
    })
}

fn byte_result(env: Env, bytes: Vec<u8>) -> Result<Uint8Array> {
    let mut buffer = ArrayBuffer::copy_from(&env, &bytes)?;
    unsafe { buffer.as_mut() }.copy_from_slice(&bytes);
    Uint8ArraySlice::from_arraybuffer(&buffer, 0, bytes.len())?.into_typed_array(&env)
}

fn native_timer_ms() -> f64 {
    static START: OnceLock<Instant> = OnceLock::new();
    START.get_or_init(Instant::now).elapsed().as_secs_f64() * 1_000.0
}

fn native_host() -> Host {
    Host {
        now_serial: native_now_serial,
        timer_ms: native_timer_ms,
        progress: native_progress,
        version: xlsx_engine::Session::version(),
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

fn map_error(env: Env, error: EngineError) -> Error {
    let result = match error.kind {
        ErrorKind::Message => env.throw(&error.message),
        ErrorKind::Error => return throw_error(&env, &error.message),
        ErrorKind::TypeError => env.throw_type_error(&error.message, None),
        ErrorKind::PendingException => return Error::new(Status::PendingException, error.message),
    };
    match result {
        Ok(()) => Error::new(Status::PendingException, error.message),
        Err(error) => error,
    }
}

#[napi]
pub struct XlsxDocument {
    inner: Option<Document>,
}

#[napi]
pub struct XlsxSnapshotBuilder {
    inner: Option<SnapshotBuilder>,
}

impl XlsxDocument {
    fn document(&self, env: Env) -> Result<&Document> {
        self.inner
            .as_ref()
            .ok_or_else(|| map_error(env, EngineError::error("XlsxDocument was freed")))
    }

    fn document_mut(&mut self, env: Env) -> Result<&mut Document> {
        self.inner
            .as_mut()
            .ok_or_else(|| map_error(env, EngineError::error("XlsxDocument was freed")))
    }
}

impl Default for XlsxSnapshotBuilder {
    fn default() -> Self {
        Self {
            inner: Some(SnapshotBuilder::new(native_host())),
        }
    }
}

#[napi]
impl XlsxSnapshotBuilder {
    #[napi(constructor)]
    pub fn new(env: Env) -> Self {
        let _call = EngineCall::new(env);
        Self::default()
    }

    #[napi]
    pub fn push(&mut self, env: Env, chunk: Uint8Array) -> Result<()> {
        let chunk = chunk.to_vec();
        let _call = EngineCall::new(env);
        let inner = self
            .inner
            .as_mut()
            .ok_or_else(|| map_error(env, EngineError::error("XlsxSnapshotBuilder was freed")))?;
        inner.push(&chunk).map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn advance(&mut self, env: Env, records: u32, bytes: u32) -> Result<bool> {
        let _call = EngineCall::new(env);
        let inner = self
            .inner
            .as_mut()
            .ok_or_else(|| map_error(env, EngineError::error("XlsxSnapshotBuilder was freed")))?;
        inner
            .advance(records as usize, bytes as usize)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn finish(&mut self, env: Env) -> Result<XlsxDocument> {
        let _call = EngineCall::new(env);
        let inner = self
            .inner
            .take()
            .ok_or_else(|| map_error(env, EngineError::error("XlsxSnapshotBuilder was freed")))?;
        inner
            .finish()
            .map(|document| XlsxDocument {
                inner: Some(document),
            })
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn free(&mut self, env: Env) {
        let _call = EngineCall::new(env);
        self.inner = None;
    }
}

#[napi]
impl XlsxDocument {
    #[napi]
    pub fn free(&mut self, env: Env) {
        let _call = EngineCall::new(env);
        self.inner = None;
    }

    #[napi(getter, js_name = "clientId")]
    pub fn client_id(&self, env: Env) -> Result<f64> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.client_id())
    }

    #[napi]
    pub fn recovery_checkpoint(
        &self,
        env: Env,
        through_seq: f64,
        preserve_history: Option<bool>,
    ) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .recovery_checkpoint(through_seq, preserve_history)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn begin_peer_snapshot(&mut self, env: Env, records: u32, bytes: u32) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .begin_peer_snapshot(records as usize, bytes as usize)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn next_peer_snapshot_chunk(&mut self, env: Env) -> Result<Either<Uint8Array, Undefined>> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .next_peer_snapshot_chunk()
            .map_err(|error| map_error(env, error))
            .and_then(|chunk| match chunk {
                Some(bytes) => byte_result(env, bytes).map(Either::A),
                None => Ok(Either::B(())),
            })
    }

    #[napi]
    pub fn end_peer_snapshot(&mut self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?.end_peer_snapshot();
        Ok(())
    }

    #[napi]
    pub fn adopt_peer_version(&mut self, env: Env, version: String) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .adopt_peer_version(&version)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn peer_hydration_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .peer_hydration_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_calculation_context_json(&mut self, env: Env, context: String) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .set_calculation_context_json(&context)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn encode_state_vector(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        byte_result(env, self.document(env)?.encode_state_vector())
    }

    #[napi]
    pub fn encode_state_as_update(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        byte_result(env, self.document(env)?.encode_state_as_update())
    }

    #[napi]
    pub fn encode_diff(&self, env: Env, remote_state_vector: Uint8Array) -> Result<Uint8Array> {
        let remote_state_vector = remote_state_vector.to_vec();
        let _call = EngineCall::new(env);
        self.document(env)?
            .encode_diff(&remote_state_vector)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn apply_update_json(&mut self, env: Env, update: Uint8Array) -> Result<String> {
        let update = update.to_vec();
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .apply_update_json(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn start_update_observation(&mut self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .start_update_observation()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn clear_update_observation(&mut self, env: Env) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?.clear_update_observation();
        Ok(())
    }

    #[napi]
    pub fn drain_update_event(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .drain_update_event()
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn print_display_list_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .print_display_list_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn display_list_json(&self, env: Env, viewport_json: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .display_list_json(&viewport_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn display_list_for_json(
        &self,
        env: Env,
        sheet: u32,
        viewport_json: String,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .display_list_for_json(sheet, &viewport_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn png_display_list_json(&self, env: Env, viewport_json: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .png_display_list_json(&viewport_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn range_png_display_list_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .range_png_display_list_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn display_list_profiled_json(&self, env: Env, viewport_json: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .display_list_profiled_json(&viewport_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn chart_at_point_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .chart_at_point_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn move_chart_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .move_chart_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn render_png(&self, env: Env, viewport_json: String) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .render_png(&viewport_json)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn render_range_png(&self, env: Env, args: String) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .render_range_png(&args)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi]
    pub fn sheet_info_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .sheet_info_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn sheet_count(&self, env: Env) -> Result<u32> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.sheet_count() as u32)
    }

    #[napi]
    pub fn sheet_info_for_json(&self, env: Env, sheet: u32) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .sheet_info_for_json(sheet)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn selection_data_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .selection_data_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn cell_rect_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .cell_rect_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn calculation_status_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .calculation_status_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_active_sheet(&mut self, env: Env, index: u32) -> Result<()> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .set_active_sheet(index)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn edit_cell_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .edit_cell_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn edit_cell_profiled_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .edit_cell_profiled_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn edit_cells_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .edit_cells_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn move_range_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .move_range_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn apply_ops_json(&mut self, env: Env, transaction_json: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .apply_ops_json(&transaction_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn charts_json(&self, env: Env, sheet: u32) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .charts_json(sheet)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn edit_charts_json(&mut self, env: Env, edits: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .edit_charts_json(&edits)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn conditional_formats_json(&self, env: Env, sheet: u32) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .conditional_formats_json(sheet)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn apply_ops_profiled_json(
        &mut self,
        env: Env,
        transaction_json: String,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .apply_ops_profiled_json(&transaction_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn undo_json(&mut self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .undo_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn redo_json(&mut self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .redo_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn cell_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .cell_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn search_text_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .search_text_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn cell_position_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .cell_position_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn range_cells_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .range_cells_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn patch_range_style_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .patch_range_style_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn set_range_number_format_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .set_range_number_format_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn selection_formatting_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .selection_formatting_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn capture_format_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .capture_format_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn apply_format_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .apply_format_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn merged_ranges_json(&self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .merged_ranges_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn visible_merged_ranges_json(
        &self,
        env: Env,
        sheet: u32,
        viewport_json: String,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .visible_merged_ranges_json(sheet, &viewport_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn history_state_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .history_state_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn propose_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .propose_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn list_proposals_json(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .list_proposals_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn accept_proposal_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .accept_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn reject_proposal_json(&mut self, env: Env, args: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .reject_proposal_json(&args)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn document_version(&self, env: Env) -> Result<String> {
        let _call = EngineCall::new(env);
        Ok(self.document(env)?.document_version())
    }

    #[napi]
    pub fn read_cells_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .read_cells_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn list_defined_names_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .list_defined_names_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn find_text_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .find_text_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn validate_edits_json(&self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .validate_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn apply_edits_json(&mut self, env: Env, request: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document_mut(env)?
            .apply_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn export_structured_json(&self, env: Env, options: String) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .export_structured_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn export_markdown_json(
        &self,
        env: Env,
        options: String,
        markdown_options: String,
    ) -> Result<String> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .export_markdown_json(&options, &markdown_options)
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn save_bytes(&self, env: Env) -> Result<Uint8Array> {
        let _call = EngineCall::new(env);
        self.document(env)?
            .save_bytes()
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(factory)]
    pub fn open(env: Env, bytes: Uint8Array) -> Result<Self> {
        let bytes = bytes.to_vec();
        let _call = EngineCall::new(env);
        Document::open(&bytes, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi(factory, js_name = "openWithCalculationJson")]
    pub fn open_with_calculation_json(
        env: Env,
        bytes: Uint8Array,
        context: String,
    ) -> Result<Self> {
        let bytes = bytes.to_vec();
        let _call = EngineCall::new(env);
        Document::open_with_calculation_json(&bytes, &context, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi(factory, js_name = "openRecoveryCheckpoint")]
    pub fn open_recovery_checkpoint(env: Env, bytes: Uint8Array, through_seq: f64) -> Result<Self> {
        let bytes = bytes.to_vec();
        let _call = EngineCall::new(env);
        Document::open_recovery_checkpoint(&bytes, through_seq, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi(factory, js_name = "openCollaborative")]
    pub fn open_collaborative(env: Env, bytes: Uint8Array, client_id: f64) -> Result<Self> {
        let bytes = bytes.to_vec();
        let _call = EngineCall::new(env);
        Document::open_collaborative(&bytes, client_id, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi(factory, js_name = "openWithPeerHydrationJson")]
    pub fn open_with_peer_hydration_json(
        env: Env,
        bytes: Uint8Array,
        hydration: String,
    ) -> Result<Self> {
        let bytes = bytes.to_vec();
        let _call = EngineCall::new(env);
        Document::open_with_peer_hydration_json(&bytes, &hydration, native_host())
            .map(|inner| Self { inner: Some(inner) })
            .map_err(|error| map_error(env, error))
    }

    #[napi]
    pub fn version(env: Env) -> String {
        let _call = EngineCall::new(env);
        xlsx_engine::Session::version().to_owned()
    }
}

#[napi(js_name = "allocationFailure")]
pub fn allocation_failure(env: Env) -> f64 {
    let _call = EngineCall::new(env);
    0.0
}

#[napi(js_name = "displayListCellLimit")]
pub fn display_list_cell_limit(env: Env) -> f64 {
    let _call = EngineCall::new(env);
    xlsx_engine::display_list_cell_limit()
}

#[napi(js_name = "exportXlsxStructuredJson")]
pub fn export_xlsx_structured_json(env: Env, bytes: Uint8Array, options: String) -> Result<String> {
    let bytes = bytes.to_vec();
    let _call = EngineCall::new(env);
    xlsx_engine::export_xlsx_structured_json(&bytes, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "exportXlsxMarkdownJson")]
pub fn export_xlsx_markdown_json(
    env: Env,
    bytes: Uint8Array,
    options: String,
    markdown_options: String,
) -> Result<String> {
    let bytes = bytes.to_vec();
    let _call = EngineCall::new(env);
    xlsx_engine::export_xlsx_markdown_json(&bytes, &options, &markdown_options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "renderXlsxMarkdownJson")]
pub fn render_xlsx_markdown_json(env: Env, content: String, options: String) -> Result<String> {
    let _call = EngineCall::new(env);
    xlsx_engine::render_xlsx_markdown_json(&content, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "engineVersion")]
pub fn engine_version(env: Env) -> String {
    let _call = EngineCall::new(env);
    env!("CARGO_PKG_VERSION").to_owned()
}
