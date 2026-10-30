#![allow(clippy::too_many_arguments)]

use std::cell::{Cell, Ref, RefCell, RefMut};
use std::collections::VecDeque;
use std::ops::{Deref, DerefMut};
use std::rc::Rc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, OnceLock};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use docx_engine::{Document, EngineError, EngineErrorKind, Host};
use napi::bindgen_prelude::{
    Array, ArrayBuffer, Either, Float64Array, FnArgs, Function, FunctionRef, JsObjectValue, Null,
    Object, Uint8Array, Uint8ArraySlice, Uint32Array, Undefined,
};
use napi::{Env, Error, Result, Status, Unknown, ValueType};
use napi_derive::napi;

fn byte_result(env: Env, bytes: Vec<u8>) -> Result<Uint8Array> {
    let mut buffer = ArrayBuffer::copy_from(&env, &bytes)?;
    unsafe { buffer.as_mut() }.copy_from_slice(&bytes);
    Uint8ArraySlice::from_arraybuffer(&buffer, 0, bytes.len())?.into_typed_array(&env)
}

fn optional_result<T>(value: Option<T>) -> Either<T, Undefined> {
    match value {
        Some(value) => Either::A(value),
        None => Either::B(()),
    }
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

fn native_clock_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis() as u64
}

fn native_host() -> Host {
    Host::new(
        Arc::new(native_clock_ms),
        Arc::new(native_entropy),
        Arc::new(native_timer_ms),
    )
}

fn map_error(env: Env, error: EngineError) -> Error {
    let result = (|| -> Result<()> {
        match error.kind {
            EngineErrorKind::Message => env.throw(&error.message),
            EngineErrorKind::PeerMetadata { code } => {
                let mut exception = env.create_error(Error::from_reason(error.message.clone()))?;
                exception.set_named_property("name", "PeerMetadataError")?;
                exception.set_named_property("code", code)?;
                env.throw(exception)
            }
        }
    })();
    match result {
        Ok(()) => Error::new(Status::PendingException, error.message),
        Err(error) => error,
    }
}

fn borrow_error() -> Error {
    Error::from_reason(
        "recursive use of an object detected which would lead to unsafe aliasing in rust",
    )
}

type UpdateCallback = FunctionRef<FnArgs<(Uint8Array, u32)>, ()>;
type UpdateQueue = Rc<RefCell<VecDeque<(Vec<u8>, bool)>>>;

struct UpdateObserver {
    _subscription: docx_engine::UpdateSubscription,
    callback: Rc<UpdateCallback>,
    env: Env,
    pending: UpdateQueue,
}

struct PendingUpdate {
    callback: Rc<UpdateCallback>,
    env: Env,
    bytes: Vec<u8>,
    remote: bool,
}

fn take_pending_exception(env: Env) -> Option<napi::sys::napi_value> {
    let mut pending = false;
    let status = unsafe { napi::sys::napi_is_exception_pending(env.raw(), &mut pending) };
    if status != napi::sys::Status::napi_ok || !pending {
        return None;
    }
    let mut exception = std::ptr::null_mut();
    let status = unsafe { napi::sys::napi_get_and_clear_last_exception(env.raw(), &mut exception) };
    if status == napi::sys::Status::napi_ok && !exception.is_null() {
        Some(exception)
    } else {
        None
    }
}

struct DeliveringUpdates<'a> {
    running: &'a Cell<bool>,
    env: Env,
    exception: Option<napi::sys::napi_value>,
}

impl Drop for DeliveringUpdates<'_> {
    fn drop(&mut self) {
        self.running.set(false);
        if let Some(exception) = self.exception {
            unsafe { napi::sys::napi_throw(self.env.raw(), exception) };
        }
    }
}

struct UpdateDelivery<'a>(&'a EditSession);

impl Drop for UpdateDelivery<'_> {
    fn drop(&mut self) {
        self.0.deliver_updates();
    }
}

struct DocumentGuard<'a, T> {
    core: T,
    _delivery: UpdateDelivery<'a>,
}

impl<T: Deref<Target = Document>> Deref for DocumentGuard<'_, T> {
    type Target = Document;

    fn deref(&self) -> &Self::Target {
        &self.core
    }
}

impl<T: DerefMut<Target = Document>> DerefMut for DocumentGuard<'_, T> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.core
    }
}

#[napi(js_name = "EditSession")]
pub struct EditSession {
    update_observer: RefCell<Option<UpdateObserver>>,
    delivering: Cell<bool>,
    core: RefCell<Option<Document>>,
}

impl EditSession {
    fn document(&self) -> Result<DocumentGuard<'_, Ref<'_, Document>>> {
        let core = self.core.try_borrow().map_err(|_| borrow_error())?;
        let core = Ref::filter_map(core, Option::as_ref)
            .map_err(|_| Error::from_reason("EditSession was freed"))?;
        Ok(DocumentGuard {
            core,
            _delivery: UpdateDelivery(self),
        })
    }

    fn document_mut(&self) -> Result<DocumentGuard<'_, RefMut<'_, Document>>> {
        let core = self.core.try_borrow_mut().map_err(|_| borrow_error())?;
        let core = RefMut::filter_map(core, Option::as_mut)
            .map_err(|_| Error::from_reason("EditSession was freed"))?;
        Ok(DocumentGuard {
            core,
            _delivery: UpdateDelivery(self),
        })
    }

    fn clear_observer(&self) -> Result<()> {
        let observer = self
            .update_observer
            .try_borrow_mut()
            .map_err(|_| borrow_error())?
            .take();
        drop(observer);
        Ok(())
    }

    fn next_update(&self) -> Option<PendingUpdate> {
        let observer = self.update_observer.try_borrow().ok()?;
        let observer = observer.as_ref()?;
        let (bytes, remote) = observer.pending.try_borrow_mut().ok()?.pop_front()?;
        Some(PendingUpdate {
            callback: Rc::clone(&observer.callback),
            env: observer.env,
            bytes,
            remote,
        })
    }

    fn deliver_updates(&self) {
        if self.delivering.get() {
            return;
        }
        let Ok(core) = self.core.try_borrow_mut() else {
            return;
        };
        if core.is_none() {
            return;
        }
        drop(core);
        let Some(update) = self.next_update() else {
            return;
        };
        self.delivering.set(true);
        let _delivery = DeliveringUpdates {
            running: &self.delivering,
            env: update.env,
            exception: take_pending_exception(update.env),
        };
        let mut next = Some(update);
        while let Some(update) = next {
            let PendingUpdate {
                callback,
                env,
                bytes,
                remote,
            } = update;
            let _ = (|| -> Result<()> {
                let bytes = byte_result(env, bytes)?;
                callback
                    .borrow_back(&env)?
                    .apply(Null, (bytes, u32::from(remote)).into())
            })();
            let _ = take_pending_exception(env);
            next = self.next_update();
        }
    }
}

#[napi]
impl EditSession {
    #[napi(js_name = "free")]
    pub fn free(&self) -> Result<()> {
        let mut core = self.core.try_borrow_mut().map_err(|_| borrow_error())?;
        self.clear_observer()?;
        drop(core.take());
        Ok(())
    }

    #[napi(js_name = "register_measure_font")]
    pub fn register_measure_font(&self, env: Env, bytes: Uint8Array) -> Result<u32> {
        self.document()?
            .register_measure_font(&bytes)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "decode_font")]
    pub fn decode_font(&self, env: Env, bytes: Uint8Array) -> Result<Uint8Array> {
        self.document()?
            .decode_font(&bytes)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "register_substitute_measure_font")]
    pub fn register_substitute_measure_font(
        &self,
        env: Env,
        base: u32,
        requested_family: String,
    ) -> Result<u32> {
        self.document()?
            .register_substitute_measure_font(base, &requested_family)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "clear_measure_fonts")]
    pub fn clear_measure_fonts(&self) -> Result<()> {
        self.document()?.clear_measure_fonts();
        Ok(())
    }

    #[napi(js_name = "measure_paragraph_json")]
    pub fn measure_paragraph_json(&self, env: Env, input: String) -> Result<String> {
        self.document()?
            .measure_paragraph_json(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_document_json")]
    pub fn layout_document_json(&self, env: Env, input: String) -> Result<String> {
        self.document()?
            .layout_document_json(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_font_requirements_json")]
    pub fn layout_font_requirements_json(&self, env: Env, input: String) -> Result<String> {
        self.document()?
            .layout_font_requirements_json(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_document_with_regions_json")]
    pub fn layout_document_with_regions_json(&self, env: Env, input: String) -> Result<String> {
        self.document()?
            .layout_document_with_regions_json(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_document_with_regions_retained_json")]
    pub fn layout_document_with_regions_retained_json(
        &self,
        env: Env,
        input: String,
    ) -> Result<String> {
        self.document()?
            .layout_document_with_regions_retained_json(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_document_with_regions_retained_meta")]
    pub fn layout_document_with_regions_retained_meta(
        &self,
        env: Env,
        input: String,
    ) -> Result<RetainedLayoutMeta> {
        self.document()?
            .layout_document_with_regions_retained_meta(&input)
            .map_err(|error| map_error(env, error))
            .map(|inner| RetainedLayoutMeta { inner: Some(inner) })
    }

    #[napi(js_name = "retained_layout_json")]
    pub fn retained_layout_json(&self, env: Env) -> Result<String> {
        self.document()?
            .retained_layout_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_document_with_regions_retained")]
    pub fn layout_document_with_regions_retained(&self, env: Env, input: String) -> Result<()> {
        self.document()?
            .layout_document_with_regions_retained(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_document_with_regions_after_edit")]
    pub fn layout_document_with_regions_after_edit(&self, env: Env, input: String) -> Result<()> {
        self.document()?
            .layout_document_with_regions_after_edit(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "retained_headers_footers_json")]
    pub fn retained_headers_footers_json(&self, env: Env) -> Result<Either<String, Undefined>> {
        self.document()?
            .retained_headers_footers_json()
            .map_err(|error| map_error(env, error))
            .map(optional_result)
    }

    #[napi(js_name = "layout_document_with_regions_prefix_retained_json")]
    pub fn layout_document_with_regions_prefix_retained_json(
        &self,
        env: Env,
        input: String,
        pages: u32,
    ) -> Result<String> {
        self.document()?
            .layout_document_with_regions_prefix_retained_json(&input, pages)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "layout_document_with_regions_exact_prefix_retained")]
    pub fn layout_document_with_regions_exact_prefix_retained(
        &self,
        env: Env,
        input: String,
        pages: u32,
    ) -> Result<String> {
        self.document()?
            .layout_document_with_regions_exact_prefix_retained(&input, pages)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "begin_region_layout")]
    pub fn begin_region_layout(&self, env: Env, input: String) -> Result<String> {
        self.document()?
            .begin_region_layout(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "resume_region_layout")]
    pub fn resume_region_layout(&self, env: Env, blocks: u32) -> Result<String> {
        self.document()?
            .resume_region_layout(blocks)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "retained_kernel_inputs_json")]
    pub fn retained_kernel_inputs_json(&self, env: Env) -> Result<String> {
        self.document()?
            .retained_kernel_inputs_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "reset_frame_base")]
    pub fn reset_frame_base(&self) -> Result<()> {
        self.document()?.reset_frame_base();
        Ok(())
    }

    #[napi(js_name = "build_display_list_json")]
    pub fn build_display_list_json(&self, env: Env, input: String) -> Result<String> {
        self.document()?
            .build_display_list_json(&input)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "build_display_list_frame")]
    pub fn build_display_list_frame(
        &self,
        env: Env,
        input: String,
        expected_frame_epoch: f64,
    ) -> Result<Uint8Array> {
        self.document()?
            .build_display_list_frame(&input, expected_frame_epoch)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "set_display_window")]
    pub fn set_display_window(&self, start: u32, end: u32) -> Result<()> {
        self.document()?.set_display_window(start, end);
        Ok(())
    }

    #[napi(js_name = "set_display_retain_built_pages")]
    pub fn set_display_retain_built_pages(&self, retain: bool) -> Result<()> {
        self.document()?.set_display_retain_built_pages(retain);
        Ok(())
    }

    #[napi(js_name = "set_local_lowering")]
    pub fn set_local_lowering(&self, enabled: bool) -> Result<()> {
        self.document()?.set_local_lowering(enabled);
        Ok(())
    }

    #[napi(js_name = "set_direct_batches")]
    pub fn set_direct_batches(&self, enabled: bool) -> Result<()> {
        self.document()?.set_direct_batches(enabled);
        Ok(())
    }

    #[napi(js_name = "direct_batches_applied")]
    pub fn direct_batches_applied(&self) -> Result<f64> {
        Ok(self.document()?.direct_batches_applied())
    }

    #[napi(js_name = "set_windowed_incremental_builds")]
    pub fn set_windowed_incremental_builds(&self, enabled: bool) -> Result<()> {
        self.document()?.set_windowed_incremental_builds(enabled);
        Ok(())
    }

    #[napi(js_name = "build_display_pages_frame")]
    pub fn build_display_pages_frame(
        &self,
        env: Env,
        pages: Uint32Array,
        expected_frame_epoch: f64,
    ) -> Result<Uint8Array> {
        self.document()?
            .build_display_pages_frame(pages.to_vec(), expected_frame_epoch)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "release_display_pages_frame")]
    pub fn release_display_pages_frame(
        &self,
        env: Env,
        pages: Uint32Array,
        expected_frame_epoch: f64,
    ) -> Result<Uint8Array> {
        self.document()?
            .release_display_pages_frame(pages.to_vec(), expected_frame_epoch)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "resident_caret_snapshot_json")]
    pub fn resident_caret_snapshot_json(&self, env: Env) -> Result<String> {
        self.document()?
            .resident_caret_snapshot_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_input")]
    pub fn apply_input(
        &self,
        env: Env,
        text: String,
        expected_frame_epoch: f64,
    ) -> Result<Uint8Array> {
        self.document()?
            .apply_input(&text, expected_frame_epoch)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "apply_input_profiled")]
    pub fn apply_input_profiled(
        &self,
        env: Env,
        text: String,
        expected_frame_epoch: f64,
    ) -> Result<Uint8Array> {
        self.document()?
            .apply_input_profiled(&text, expected_frame_epoch)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "apply_delete")]
    pub fn apply_delete(
        &self,
        env: Env,
        direction: String,
        expected_frame_epoch: f64,
        count: u32,
    ) -> Result<Uint8Array> {
        self.document()?
            .apply_delete(&direction, expected_frame_epoch, count)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "resident_deleted_units")]
    pub fn resident_deleted_units(&self) -> Result<u32> {
        Ok(self.document()?.resident_deleted_units())
    }

    #[napi(js_name = "apply_delete_profiled")]
    pub fn apply_delete_profiled(
        &self,
        env: Env,
        direction: String,
        expected_frame_epoch: f64,
        count: u32,
    ) -> Result<Uint8Array> {
        self.document()?
            .apply_delete_profiled(&direction, expected_frame_epoch, count)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "apply_input_profile_json")]
    pub fn apply_input_profile_json(&self) -> Result<String> {
        Ok(self.document()?.apply_input_profile_json())
    }

    #[napi(js_name = "display_hit_test_regions_json")]
    pub fn display_hit_test_regions_json(
        &self,
        env: Env,
        page_index: u32,
        x: f64,
        y: f64,
    ) -> Result<String> {
        self.document()?
            .display_hit_test_regions_json(page_index, x, y)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "display_vertical_move_json")]
    pub fn display_vertical_move_json(
        &self,
        env: Env,
        position: f64,
        direction: String,
        goal_x: f64,
    ) -> Result<String> {
        self.document()?
            .display_vertical_move_json(position, &direction, goal_x)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "display_range_rects_json")]
    pub fn display_range_rects_json(&self, env: Env, from: f64, to: f64) -> Result<String> {
        self.document()?
            .display_range_rects_json(from, to)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "display_range_rects_on_pages_json")]
    pub fn display_range_rects_on_pages_json(
        &self,
        env: Env,
        from: f64,
        to: f64,
        first_page: f64,
        last_page: f64,
    ) -> Result<String> {
        self.document()?
            .display_range_rects_on_pages_json(from, to, first_page, last_page)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "display_range_rects_region_json")]
    pub fn display_range_rects_region_json(
        &self,
        env: Env,
        region: String,
        part_id: String,
        from: f64,
        to: f64,
    ) -> Result<String> {
        self.document()?
            .display_range_rects_region_json(&region, &part_id, from, to)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "display_range_rects_batch")]
    pub fn display_range_rects_batch(
        &self,
        env: Env,
        region: String,
        part_id: String,
        ranges: Float64Array,
    ) -> Result<Float64Array> {
        self.document()?
            .display_range_rects_batch(&region, &part_id, ranges.to_vec())
            .map_err(|error| map_error(env, error))
            .map(Float64Array::new)
    }

    #[napi(js_name = "outline_glyph_json")]
    pub fn outline_glyph_json(&self, env: Env, font_id: u32, glyph_id: u32) -> Result<String> {
        self.document()?
            .outline_glyph_json(font_id, glyph_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "echo_glyph_set_json")]
    pub fn echo_glyph_set_json(
        &self,
        env: Env,
        font_id: u32,
        size: f64,
        kern: bool,
        alphabet: String,
    ) -> Result<String> {
        self.document()?
            .echo_glyph_set_json(font_id, size as f32, kern, &alphabet)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "accept_change")]
    pub fn accept_change(&self, env: Env, target_json: String) -> Result<String> {
        self.document()?
            .accept_change(&target_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "add_comment")]
    pub fn add_comment(
        &self,
        env: Env,
        ranges_json: String,
        author: String,
        date: String,
        body_json: String,
    ) -> Result<String> {
        self.document()?
            .add_comment(&ranges_json, &author, &date, &body_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "add_undo_boundary")]
    pub fn add_undo_boundary(&self) -> Result<()> {
        self.document()?.add_undo_boundary();
        Ok(())
    }

    #[napi(js_name = "apply_edits_json")]
    pub fn apply_edits_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .apply_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_host_update")]
    pub fn apply_host_update(&self, env: Env, update: Uint8Array) -> Result<()> {
        self.document()?
            .apply_host_update(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_local_update")]
    pub fn apply_local_update(&self, env: Env, update: Uint8Array) -> Result<()> {
        self.document()?
            .apply_local_update(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_paragraph_style")]
    pub fn apply_paragraph_style(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        style_id: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<()> {
        self.document()?
            .apply_paragraph_style(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                &style_id,
                author_name,
                author_date,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_raw_ops")]
    pub fn apply_raw_ops(&self, env: Env, story: String, ops_json: String) -> Result<()> {
        self.document()?
            .apply_raw_ops(&story, &ops_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_seed_raw_ops")]
    pub fn apply_seed_raw_ops(&self, env: Env, story: String, ops_json: String) -> Result<()> {
        self.document()?
            .apply_seed_raw_ops(&story, &ops_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_update_chunks", ts_args_type = "updates: Array<any>")]
    pub fn apply_update_chunks(&self, env: Env, updates: Array<'_>) -> Result<()> {
        let mut chunks = Vec::with_capacity(updates.len() as usize);
        for index in 0..updates.len() {
            let bytes = updates
                .get::<Uint8ArraySlice<'_>>(index)
                .map_err(|_| {
                    map_error(
                        env,
                        EngineError::message("sync updates must be Uint8Arrays"),
                    )
                })?
                .ok_or_else(|| {
                    map_error(
                        env,
                        EngineError::message("sync updates must be Uint8Arrays"),
                    )
                })?;
            chunks.push(bytes.to_vec());
        }
        self.document()?
            .apply_update_chunks(&chunks)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_update")]
    pub fn apply_update(&self, env: Env, update: Uint8Array) -> Result<()> {
        self.document()?
            .apply_update(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "apply_update_with_inference")]
    pub fn apply_update_with_inference(&self, env: Env, update: Uint8Array) -> Result<String> {
        self.document()?
            .apply_update_with_inference(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "begin_opening")]
    pub fn begin_opening(&self, generation: Option<String>) -> Result<()> {
        self.document()?.begin_opening(generation);
        Ok(())
    }

    #[napi(js_name = "begin_shared_reads")]
    pub fn begin_shared_reads(&self) -> Result<()> {
        self.document()?.begin_shared_reads();
        Ok(())
    }

    #[napi(js_name = "bootstrap_peer")]
    pub fn bootstrap_peer(
        &self,
        env: Env,
        state: Uint8Array,
        metadata: Uint8Array,
        source: Option<Uint8Array>,
    ) -> Result<()> {
        self.document()?
            .bootstrap_peer(&state, &metadata, source.map(|bytes| bytes.to_vec()))
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "can_redo")]
    pub fn can_redo(&self) -> Result<bool> {
        Ok(self.document()?.can_redo())
    }

    #[napi(js_name = "can_undo")]
    pub fn can_undo(&self) -> Result<bool> {
        Ok(self.document()?.can_undo())
    }

    #[napi(js_name = "cell_selection")]
    pub fn cell_selection(&self, env: Env) -> Result<String> {
        self.document()?
            .cell_selection()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "clear_cell_selection")]
    pub fn clear_cell_selection(&self) -> Result<Either<String, Undefined>> {
        Ok(optional_result(self.document()?.clear_cell_selection()))
    }

    #[napi(js_name = "clear_content_control_value")]
    pub fn clear_content_control_value(&self, env: Env, embed_id: String) -> Result<()> {
        self.document()?
            .clear_content_control_value(&embed_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "clear_formatting")]
    pub fn clear_formatting(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
    ) -> Result<()> {
        self.document()?
            .clear_formatting(&story, &start_para, start_offset, &end_para, end_offset)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "clear_undo_input_time")]
    pub fn clear_undo_input_time(&self) -> Result<()> {
        self.document()?.clear_undo_input_time();
        Ok(())
    }

    #[napi(js_name = "clear_update_event_observation")]
    pub fn clear_update_event_observation(&self) -> Result<()> {
        self.document_mut()?.clear_update_event_observation();
        Ok(())
    }

    #[napi(js_name = "client_id")]
    pub fn client_id(&self) -> Result<f64> {
        Ok(self.document()?.client_id())
    }

    #[napi(js_name = "comment_text_target_json")]
    pub fn comment_text_target_json(
        &self,
        env: Env,
        target_json: String,
        comment_json: String,
    ) -> Result<String> {
        self.document()?
            .comment_text_target_json(&target_json, &comment_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "compare_docx_json")]
    pub fn compare_docx_json(
        &self,
        env: Env,
        original: Uint8Array,
        revised: Uint8Array,
        options: String,
    ) -> Result<String> {
        self.document()?
            .compare_docx_json(&original, &revised, &options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "content_control_targets_json")]
    pub fn content_control_targets_json(&self, env: Env) -> Result<String> {
        self.document()?
            .content_control_targets_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "create_story")]
    pub fn create_story(
        &self,
        env: Env,
        story_id: String,
        initial_text: String,
        p_style: String,
        alignment: String,
    ) -> Result<String> {
        self.document()?
            .create_story(&story_id, &initial_text, &p_style, &alignment)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "delete_column")]
    pub fn delete_column(&self, env: Env, range_json: String) -> Result<String> {
        self.document()?
            .delete_column(&range_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "delete_range")]
    pub fn delete_range(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .delete_range(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                author_name,
                author_date,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "delete_row")]
    pub fn delete_row(
        &self,
        env: Env,
        range_json: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .delete_row(&range_json, author_name, author_date)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "delete_story")]
    pub fn delete_story(&self, env: Env, story_id: String) -> Result<()> {
        self.document()?
            .delete_story(&story_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "delete_table")]
    pub fn delete_table(&self, env: Env, table_json: String) -> Result<String> {
        self.document()?
            .delete_table(&table_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "document_properties_json")]
    pub fn document_properties_json(&self) -> Result<String> {
        Ok(self.document()?.document_properties_json())
    }

    #[napi(js_name = "drain_update_event")]
    pub fn drain_update_event(&self, env: Env) -> Result<Uint8Array> {
        byte_result(env, self.document()?.drain_update_event())
    }

    #[napi(js_name = "edit_text_resident")]
    pub fn edit_text_resident(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        text: Option<String>,
    ) -> Result<String> {
        self.document()?
            .edit_text_resident(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                text,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "encode_diff_chunks", ts_return_type = "Array<any>")]
    pub fn encode_diff_chunks(
        &self,
        env: Env,
        remote_state_vector: Uint8Array,
        max_chunk_bytes: u32,
    ) -> Result<Array<'_>> {
        let chunks = self
            .document()?
            .encode_diff_chunks(&remote_state_vector, max_chunk_bytes)
            .map_err(|error| map_error(env, error))?;
        let chunks = chunks
            .into_iter()
            .map(|bytes| byte_result(env, bytes))
            .collect::<Result<Vec<_>>>()?;
        Array::from_vec(&env, chunks)
    }

    #[napi(js_name = "encode_diff")]
    pub fn encode_diff(&self, env: Env, remote_state_vector: Uint8Array) -> Result<Uint8Array> {
        self.document()?
            .encode_diff(&remote_state_vector)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "encode_peer_metadata")]
    pub fn encode_peer_metadata(&self, env: Env) -> Result<Uint8Array> {
        self.document()?
            .encode_peer_metadata()
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "encode_state")]
    pub fn encode_state(&self, env: Env) -> Result<Uint8Array> {
        byte_result(env, self.document()?.encode_state())
    }

    #[napi(js_name = "encode_state_vector")]
    pub fn encode_state_vector(&self, env: Env) -> Result<Uint8Array> {
        byte_result(env, self.document()?.encode_state_vector())
    }

    #[napi(js_name = "encode_sticky_position")]
    pub fn encode_sticky_position(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
    ) -> Result<Uint8Array> {
        self.document()?
            .encode_sticky_position(&story, &para_id, offset)
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "encoded_selection")]
    pub fn encoded_selection(&self, env: Env) -> Result<String> {
        self.document()?
            .encoded_selection()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "end_shared_reads")]
    pub fn end_shared_reads(&self) -> Result<()> {
        self.document()?.end_shared_reads();
        Ok(())
    }

    #[napi(js_name = "export_markdown_json")]
    pub fn export_markdown_json(&self, env: Env, options: String) -> Result<String> {
        self.document()?
            .export_markdown_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "export_snapshot_with_private_fonts_json")]
    pub fn export_snapshot_with_private_fonts_json(
        &self,
        env: Env,
        fonts: Uint8Array,
        font_lengths: Uint32Array,
        request: String,
        options: String,
    ) -> Result<String> {
        self.document()?
            .export_snapshot_with_private_fonts_json(&fonts, &font_lengths, &request, &options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "export_structured_json")]
    pub fn export_structured_json(&self, env: Env, options: String) -> Result<String> {
        self.document()?
            .export_structured_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "export_structured_with_pages_json")]
    pub fn export_structured_with_pages_json(
        &self,
        env: Env,
        options: String,
        current_request: Option<String>,
    ) -> Result<String> {
        self.document()?
            .export_structured_with_pages_json(&options, current_request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "fail_compared_docx_json")]
    pub fn fail_compared_docx_json(&self, env: Env, message: String) -> Result<String> {
        self.document()?
            .fail_compared_docx_json(&message)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "find_content_controls_json")]
    pub fn find_content_controls_json(
        &self,
        env: Env,
        query: String,
        options: String,
    ) -> Result<String> {
        self.document()?
            .find_content_controls_json(&query, &options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "find_text_json")]
    pub fn find_text_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .find_text_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "finish_compared_docx_json")]
    pub fn finish_compared_docx_json(&self, env: Env, bytes: Uint8Array) -> Result<String> {
        self.document()?
            .finish_compared_docx_json(&bytes)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "format_range")]
    pub fn format_range(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        delta_json: String,
    ) -> Result<()> {
        self.document()?
            .format_range(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                &delta_json,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "format_text_target_json")]
    pub fn format_text_target_json(
        &self,
        env: Env,
        target_json: String,
        delta_json: String,
    ) -> Result<String> {
        self.document()?
            .format_text_target_json(&target_json, &delta_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "geometry_position_outline_json")]
    pub fn geometry_position_outline_json(&self, env: Env, root: String) -> Result<String> {
        self.document()?
            .geometry_position_outline_json(&root)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "has_pending_update")]
    pub fn has_pending_update(&self) -> Result<bool> {
        Ok(self.document()?.has_pending_update())
    }

    #[napi(js_name = "has_story")]
    pub fn has_story(&self, story: String) -> Result<bool> {
        Ok(self.document()?.has_story(&story))
    }

    #[napi(js_name = "header_footer_aliases_json")]
    pub fn header_footer_aliases_json(&self) -> Result<String> {
        Ok(self.document()?.header_footer_aliases_json())
    }

    #[napi(js_name = "headings_json")]
    pub fn headings_json(&self, env: Env, story: String) -> Result<String> {
        self.document()?
            .headings_json(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "history_stories", ts_return_type = "string[]")]
    pub fn history_stories(&self) -> Result<Vec<String>> {
        Ok(self.document()?.history_stories())
    }

    #[napi(js_name = "insert_column")]
    pub fn insert_column(&self, env: Env, at_json: String, after: bool) -> Result<String> {
        self.document()?
            .insert_column(&at_json, after)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_image")]
    pub fn insert_image(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        payload_json: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .insert_image(
                &story,
                &para_id,
                offset,
                &payload_json,
                author_name,
                author_date,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_line_break")]
    pub fn insert_line_break(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .insert_line_break(&story, &para_id, offset, author_name, author_date)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_page_break")]
    pub fn insert_page_break(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
    ) -> Result<()> {
        self.document()?
            .insert_page_break(&story, &para_id, offset)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_row")]
    pub fn insert_row(
        &self,
        env: Env,
        at_json: String,
        after: bool,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .insert_row(&at_json, after, author_name, author_date)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_section_break")]
    pub fn insert_section_break(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        break_type: String,
    ) -> Result<()> {
        self.document()?
            .insert_section_break(&story, &para_id, offset, &break_type)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_table")]
    pub fn insert_table(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        rows: u32,
        columns: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .insert_table(
                &story,
                &para_id,
                offset,
                rows,
                columns,
                author_name,
                author_date,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_table_fragment")]
    pub fn insert_table_fragment(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        fragment_json: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .insert_table_fragment(
                &story,
                &para_id,
                offset,
                &fragment_json,
                author_name,
                author_date,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_text")]
    pub fn insert_text(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        text: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .insert_text(&story, &para_id, offset, &text, author_name, author_date)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "insert_watermark")]
    pub fn insert_watermark(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        watermark_json: String,
    ) -> Result<()> {
        self.document()?
            .insert_watermark(&story, &para_id, offset, &watermark_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "list_comments")]
    pub fn list_comments(&self, env: Env) -> Result<String> {
        self.document()?
            .list_comments()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "list_content_controls_json")]
    pub fn list_content_controls_json(&self, env: Env, options: String) -> Result<String> {
        self.document()?
            .list_content_controls_json(&options)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "list_revisions")]
    pub fn list_revisions(&self, env: Env) -> Result<String> {
        self.document()?
            .list_revisions()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "list_sections_json")]
    pub fn list_sections_json(&self, env: Env) -> Result<String> {
        self.document()?
            .list_sections_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "list_stories_json")]
    pub fn list_stories_json(&self, env: Env) -> Result<String> {
        self.document()?
            .list_stories_json()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "load")]
    pub fn load(&self, env: Env, update: Uint8Array) -> Result<()> {
        self.document()?
            .load(&update)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "load_json")]
    pub fn load_json(&self, env: Env, stories_json: String) -> Result<String> {
        self.document()?
            .load_json(&stories_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "load_media_sources")]
    pub fn load_media_sources(&self, env: Env, json: String) -> Result<()> {
        self.document()?
            .load_media_sources(&json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "load_note_separators")]
    pub fn load_note_separators(&self, env: Env, state: Uint8Array) -> Result<()> {
        self.document()?
            .load_note_separators(&state)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "locate_paragraph")]
    pub fn locate_paragraph(&self, env: Env, story: String, para_id: String) -> Result<String> {
        self.document()?
            .locate_paragraph(&story, &para_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "materialize_docx")]
    pub fn materialize_docx(&self, env: Env) -> Result<Either<String, Undefined>> {
        self.document()?
            .materialize_docx()
            .map_err(|error| map_error(env, error))
            .map(optional_result)
    }

    #[napi(js_name = "media_bytes")]
    pub fn media_bytes(&self, env: Env, token: String) -> Result<Either<Uint8Array, Undefined>> {
        match self.document()?.media_bytes(&token) {
            Some(bytes) => byte_result(env, bytes).map(Either::A),
            None => Ok(Either::B(())),
        }
    }

    #[napi(js_name = "media_data_url")]
    pub fn media_data_url(&self, token: String) -> Result<Either<String, Undefined>> {
        Ok(optional_result(self.document()?.media_data_url(&token)))
    }

    #[napi(js_name = "media_sources_json")]
    pub fn media_sources_json(&self) -> Result<String> {
        Ok(self.document()?.media_sources_json())
    }

    #[napi(js_name = "media_type")]
    pub fn media_type(&self, token: String) -> Result<Either<String, Undefined>> {
        Ok(optional_result(self.document()?.media_type(&token)))
    }

    #[napi(js_name = "merge_cells")]
    pub fn merge_cells(&self, env: Env, range_json: String) -> Result<String> {
        self.document()?
            .merge_cells(&range_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "merge_paragraphs")]
    pub fn merge_paragraphs(
        &self,
        env: Env,
        story: String,
        para_id: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .merge_paragraphs(&story, &para_id, author_name, author_date)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "move_embed")]
    pub fn move_embed(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        to_story: String,
        to_para_id: String,
        to_offset: u32,
    ) -> Result<String> {
        self.document()?
            .move_embed(&story, &para_id, offset, &to_story, &to_para_id, to_offset)
            .map_err(|error| map_error(env, error))
    }

    #[napi(constructor, js_name = "new")]
    pub fn new(env: Env, client_id: f64) -> Result<Self> {
        const MAX_SAFE_INTEGER: f64 = 9_007_199_254_740_991.0;
        if !(client_id.is_finite()
            && client_id >= 0.0
            && client_id.fract() == 0.0
            && client_id <= MAX_SAFE_INTEGER)
        {
            return Err(map_error(
                env,
                EngineError::message("client_id must be a non-negative safe integer"),
            ));
        }
        Ok(Self {
            update_observer: RefCell::new(None),
            delivering: Cell::new(false),
            core: RefCell::new(Some(Document::new(client_id as u64, native_host()))),
        })
    }

    #[napi(js_name = "next_paragraph_style")]
    pub fn next_paragraph_style(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
    ) -> Result<Either<String, Undefined>> {
        self.document()?
            .next_paragraph_style(&story, &para_id, offset)
            .map_err(|error| map_error(env, error))
            .map(optional_result)
    }

    #[napi(js_name = "note_separators_state")]
    pub fn note_separators_state(&self, env: Env) -> Result<Uint8Array> {
        self.document()?
            .note_separators_state()
            .map_err(|error| map_error(env, error))
            .and_then(|bytes| byte_result(env, bytes))
    }

    #[napi(js_name = "open_docx")]
    pub fn open_docx_owned(
        &self,
        env: Env,
        bytes: Uint8Array,
        seed_stories: bool,
        generation: Option<String>,
        digest: Option<String>,
    ) -> Result<String> {
        self.document()?
            .open_docx_owned(bytes.to_vec(), seed_stories, generation, digest)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "open_docx_preview")]
    pub fn open_docx_preview_owned(
        &self,
        env: Env,
        bytes: Uint8Array,
        blocks: u32,
    ) -> Result<Either<String, Undefined>> {
        self.document()?
            .open_docx_preview_owned(bytes.to_vec(), blocks)
            .map_err(|error| map_error(env, error))
            .map(optional_result)
    }

    #[napi(js_name = "open_docx_preview_with_budget")]
    pub fn open_docx_preview_with_budget_owned(
        &self,
        env: Env,
        bytes: Uint8Array,
        blocks: u32,
        paragraph_budget: Option<u32>,
    ) -> Result<Either<String, Undefined>> {
        self.document()?
            .open_docx_preview_with_budget_owned(bytes.to_vec(), blocks, paragraph_budget)
            .map_err(|error| map_error(env, error))
            .map(optional_result)
    }

    #[napi(js_name = "paragraph_id_count")]
    pub fn paragraph_id_count(&self, env: Env, story: String, para_id: String) -> Result<u32> {
        self.document()?
            .paragraph_id_count(&story, &para_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "paragraph_identities")]
    pub fn paragraph_identities(&self, env: Env) -> Result<String> {
        self.document()?
            .paragraph_identities()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "paragraph_save_plan")]
    pub fn paragraph_save_plan(&self, env: Env) -> Result<String> {
        self.document()?
            .paragraph_save_plan()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "paragraph_spans")]
    pub fn paragraph_spans(&self, env: Env, story: String) -> Result<String> {
        self.document()?
            .paragraph_spans(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "paragraph_unit")]
    pub fn paragraph_unit(&self, env: Env, story: String, para_id: String) -> Result<i32> {
        self.document()?
            .paragraph_unit(&story, &para_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "paragraphs")]
    pub fn paragraphs(&self, env: Env, story: String) -> Result<String> {
        self.document()?
            .paragraphs(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "part_metadata_json")]
    pub fn part_metadata_json(&self) -> Result<String> {
        Ok(self.document()?.part_metadata_json())
    }

    #[napi(js_name = "persist_paragraph_ids")]
    pub fn persist_paragraph_ids(&self, env: Env) -> Result<String> {
        self.document()?
            .persist_paragraph_ids()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "proposal_revision_ranges_json")]
    pub fn proposal_revision_ranges_json(&self, env: Env, ids_json: String) -> Result<String> {
        self.document()?
            .proposal_revision_ranges_json(&ids_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "read_paragraphs_json")]
    pub fn read_paragraphs_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .read_paragraphs_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "read_stories_json")]
    pub fn read_stories_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .read_stories_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "record_saved_paragraph_ids")]
    pub fn record_saved_paragraph_ids(&self, env: Env, saved_json: String) -> Result<String> {
        self.document()?
            .record_saved_paragraph_ids(&saved_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "redo")]
    pub fn redo(&self) -> Result<bool> {
        Ok(self.document()?.redo())
    }

    #[napi(js_name = "reject_change")]
    pub fn reject_change(&self, env: Env, target_json: String) -> Result<String> {
        self.document()?
            .reject_change(&target_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "repeat_content_control_at")]
    pub fn repeat_content_control_at(
        &self,
        env: Env,
        story: String,
        index: u32,
        remove: bool,
    ) -> Result<bool> {
        self.document()?
            .repeat_content_control_at(&story, index, remove)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "replace_range")]
    pub fn replace_range(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        text: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .replace_range(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                &text,
                author_name,
                author_date,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "resident_geometry_json")]
    pub fn resident_geometry_json(
        &self,
        env: Env,
        root: String,
        request: String,
    ) -> Result<String> {
        self.document()?
            .resident_geometry_json(&root, &request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "resident_section_revision")]
    pub fn resident_section_revision(&self) -> Result<String> {
        Ok(self.document()?.resident_section_revision())
    }

    #[napi(js_name = "resolve_comment")]
    pub fn resolve_comment(&self, env: Env, comment_id: String) -> Result<String> {
        self.document()?
            .resolve_comment(&comment_id)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "resolve_encoded_selection")]
    pub fn resolve_encoded_selection(
        &self,
        env: Env,
        story: String,
        anchor: Uint8Array,
        head: Uint8Array,
    ) -> Result<String> {
        self.document()?
            .resolve_encoded_selection(&story, &anchor, &head)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "resolve_paragraph_anchor")]
    pub fn resolve_paragraph_anchor(&self, env: Env, anchor_json: String) -> Result<String> {
        self.document()?
            .resolve_paragraph_anchor(&anchor_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "resolve_sticky_position")]
    pub fn resolve_sticky_position(
        &self,
        env: Env,
        story: String,
        position: Uint8Array,
    ) -> Result<String> {
        self.document()?
            .resolve_sticky_position(&story, &position)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "revision_ids_at")]
    pub fn revision_ids_at(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
    ) -> Result<String> {
        self.document()?
            .revision_ids_at(&story, &para_id, offset)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "revision_stamps_json")]
    pub fn revision_stamps_json(&self, env: Env, ids_json: String) -> Result<String> {
        self.document()?
            .revision_stamps_json(&ids_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "run_inheritance")]
    pub fn run_inheritance(&self, env: Env, styles_json: String) -> Result<RunInheritance> {
        let _core = self.document()?;
        let styles = serde_json::from_str(&styles_json)
            .map_err(|error| map_error(env, EngineError::message(error)))?;
        Ok(RunInheritance {
            inner: Some(docx_engine::RunInheritance::new(Some(&styles))),
        })
    }

    #[napi(js_name = "search_text")]
    pub fn search_text(
        &self,
        env: Env,
        query: String,
        case_sensitive: bool,
        limit: Option<u32>,
    ) -> Result<String> {
        self.document()?
            .search_text(&query, case_sensitive, limit)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "section_margins_json")]
    pub fn section_margins_json(&self) -> Result<String> {
        Ok(self.document()?.section_margins_json())
    }

    #[napi(js_name = "seed_comment_metadata_json")]
    pub fn seed_comment_metadata_json(
        &self,
        env: Env,
        comment_id: String,
        patch_json: String,
    ) -> Result<()> {
        self.document()?
            .seed_comment_metadata_json(&comment_id, &patch_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "seed_from_docx")]
    pub fn seed_from_docx_owned(
        &self,
        env: Env,
        bytes: Uint8Array,
        generation: Option<String>,
    ) -> Result<String> {
        self.document()?
            .seed_from_docx_owned(bytes.to_vec(), generation)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "seed_opaque_sequences")]
    pub fn seed_opaque_sequences(&self, env: Env, names_json: String) -> Result<()> {
        self.document()?
            .seed_opaque_sequences(&names_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "select_story")]
    pub fn select_story(&self, story: String) -> Result<()> {
        self.document()?.select_story(&story);
        Ok(())
    }

    #[napi(js_name = "selection_context")]
    pub fn selection_context(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
    ) -> Result<String> {
        self.document()?
            .selection_context(&story, &start_para, start_offset, &end_para, end_offset)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "selection")]
    pub fn selection(&self, env: Env) -> Result<String> {
        self.document()?
            .selection()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "selection_text_json")]
    pub fn selection_text_json(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
    ) -> Result<String> {
        self.document()?
            .selection_text_json(&story, &start_para, start_offset, &end_para, end_offset)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_cell_borders")]
    pub fn set_cell_borders(
        &self,
        env: Env,
        range_json: String,
        borders_json: String,
    ) -> Result<String> {
        self.document()?
            .set_cell_borders(&range_json, &borders_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_cell_selection")]
    pub fn set_cell_selection(&self, env: Env, range_json: String) -> Result<()> {
        self.document()?
            .set_cell_selection(&range_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_cell_shading")]
    pub fn set_cell_shading(
        &self,
        env: Env,
        range_json: String,
        color: Option<String>,
    ) -> Result<String> {
        self.document()?
            .set_cell_shading(&range_json, color)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_cell_text_format")]
    pub fn set_cell_text_format(
        &self,
        env: Env,
        range_json: String,
        patch_json: String,
    ) -> Result<String> {
        self.document()?
            .set_cell_text_format(&range_json, &patch_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_column_width")]
    pub fn set_column_width(&self, env: Env, at_json: String, width_twips: f64) -> Result<String> {
        self.document()?
            .set_column_width(&at_json, width_twips)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_comment_ranges")]
    pub fn set_comment_ranges(
        &self,
        env: Env,
        comment_id: String,
        ranges_json: String,
    ) -> Result<()> {
        self.document()?
            .set_comment_ranges(&comment_id, &ranges_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_content_control_value_at")]
    pub fn set_content_control_value_at(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        value_json: String,
    ) -> Result<()> {
        self.document()?
            .set_content_control_value_at(&story, &para_id, offset, &value_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_content_control_value")]
    pub fn set_content_control_value(
        &self,
        env: Env,
        embed_id: String,
        value_json: String,
    ) -> Result<()> {
        self.document()?
            .set_content_control_value(&embed_id, &value_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_content_control_value_index")]
    pub fn set_content_control_value_index(
        &self,
        env: Env,
        story: String,
        index: u32,
        json: String,
    ) -> Result<()> {
        self.document()?
            .set_content_control_value_index(&story, index, &json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_document_property")]
    pub fn set_document_property(&self, env: Env, key: String, json: String) -> Result<bool> {
        self.document()?
            .set_document_property(&key, &json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_header_footer_aliases")]
    pub fn set_header_footer_aliases(&self, env: Env, json: String) -> Result<()> {
        self.document()?
            .set_header_footer_aliases(&json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_hyperlink")]
    pub fn set_hyperlink(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        hyperlink_json: String,
    ) -> Result<()> {
        self.document()?
            .set_hyperlink(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                &hyperlink_json,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_image_geometry_at")]
    pub fn set_image_geometry_at(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        geometry_json: String,
    ) -> Result<()> {
        self.document()?
            .set_image_geometry_at(&story, &para_id, offset, &geometry_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_image_geometry")]
    pub fn set_image_geometry(
        &self,
        env: Env,
        embed_id: String,
        geometry_json: String,
    ) -> Result<()> {
        self.document()?
            .set_image_geometry(&embed_id, &geometry_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_media_tokens")]
    pub fn set_media_tokens(&self, enabled: bool) -> Result<()> {
        self.document()?.set_media_tokens(enabled);
        Ok(())
    }

    #[napi(js_name = "set_paragraph_attr")]
    pub fn set_paragraph_attr(
        &self,
        env: Env,
        para_id: String,
        key: String,
        value_json: String,
    ) -> Result<()> {
        self.document()?
            .set_paragraph_attr(&para_id, &key, &value_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_paragraph_attrs")]
    pub fn set_paragraph_attrs(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        attrs_json: String,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<()> {
        self.document()?
            .set_paragraph_attrs(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                &attrs_json,
                author_name,
                author_date,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_part_metadata")]
    pub fn set_part_metadata(&self, env: Env, json: String, remove_json: String) -> Result<bool> {
        self.document()?
            .set_part_metadata(&json, &remove_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_partial_document")]
    pub fn set_partial_document(&self, partial: bool) -> Result<()> {
        self.document()?.set_partial_document(partial);
        Ok(())
    }

    #[napi(js_name = "set_section_margins")]
    pub fn set_section_margins(&self, env: Env, section: String, json: String) -> Result<bool> {
        self.document()?
            .set_section_margins(&section, &json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_selection")]
    pub fn set_selection(
        &self,
        env: Env,
        story: String,
        anchor_para: String,
        anchor_offset: u32,
        head_para: String,
        head_offset: u32,
    ) -> Result<()> {
        self.document()?
            .set_selection(&story, &anchor_para, anchor_offset, &head_para, head_offset)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_table_properties")]
    pub fn set_table_properties(
        &self,
        env: Env,
        table_json: String,
        patch_json: String,
    ) -> Result<String> {
        self.document()?
            .set_table_properties(&table_json, &patch_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_table_width")]
    pub fn set_table_width(
        &self,
        env: Env,
        table_json: String,
        width_twips: f64,
    ) -> Result<String> {
        self.document()?
            .set_table_width(&table_json, width_twips)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_undo_capture_mode")]
    pub fn set_undo_capture_mode(&self, env: Env, mode: String) -> Result<()> {
        self.document()?
            .set_undo_capture_mode(&mode)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "set_undo_input_time")]
    pub fn set_undo_input_time(&self, first: f64, last: f64) -> Result<()> {
        self.document()?.set_undo_input_time(first, last);
        Ok(())
    }

    #[napi(js_name = "settle_revisions_json")]
    pub fn settle_revisions_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .settle_revisions_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "split_cell")]
    pub fn split_cell(
        &self,
        env: Env,
        at_json: String,
        rows: Option<u32>,
        columns: Option<u32>,
    ) -> Result<String> {
        self.document()?
            .split_cell(&at_json, rows, columns)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "split_paragraph")]
    pub fn split_paragraph(
        &self,
        env: Env,
        story: String,
        para_id: String,
        offset: u32,
        author_name: Option<String>,
        author_date: Option<String>,
    ) -> Result<String> {
        self.document()?
            .split_paragraph(&story, &para_id, offset, author_name, author_date)
            .map_err(|error| map_error(env, error))
    }

    /// Runs callbacks before each call returns, after its transactions commit, allowing session reads and edits.
    /// For calls with multiple commits, every callback sees the state after the call.
    #[napi(js_name = "set_update_observer", ts_args_type = "callback: Function")]
    pub fn set_update_observer(
        &self,
        env: Env,
        callback: Function<'_, FnArgs<(Uint8Array, u32)>, ()>,
    ) -> Result<()> {
        let core = self.document_mut()?;
        let callback = Rc::new(callback.create_ref()?);
        let pending = Rc::new(RefCell::new(VecDeque::new()));
        let observed = Rc::clone(&pending);
        let subscription = core
            .observe_updates(move |update, remote| {
                if let Ok(mut pending) = observed.try_borrow_mut() {
                    pending.push_back((update.to_vec(), remote));
                }
            })
            .map_err(|error| map_error(env, error))?;
        let observer = UpdateObserver {
            _subscription: subscription,
            callback,
            env,
            pending,
        };
        let previous = self
            .update_observer
            .try_borrow_mut()
            .map_err(|_| borrow_error())?
            .replace(observer);
        drop(previous);
        Ok(())
    }

    #[napi(js_name = "clear_update_observer")]
    pub fn clear_update_observer(&self) -> Result<()> {
        let _core = self.document_mut()?;
        self.clear_observer()
    }

    #[napi(js_name = "start_update_event_observation")]
    pub fn start_update_event_observation(&self, env: Env) -> Result<()> {
        self.document_mut()?
            .start_update_event_observation()
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "stories_changed_since")]
    pub fn stories_changed_since(&self, since: f64) -> Result<String> {
        Ok(self.document()?.stories_changed_since(since))
    }

    #[napi(js_name = "story_checksum")]
    pub fn story_checksum(&self, env: Env, story: String) -> Result<String> {
        self.document()?
            .story_checksum(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "story_ids", ts_return_type = "string[]")]
    pub fn story_ids(&self) -> Result<Vec<String>> {
        Ok(self.document()?.story_ids())
    }

    #[napi(js_name = "story_len")]
    pub fn story_len(&self, env: Env, story: String) -> Result<u32> {
        self.document()?
            .story_len(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "story_paragraph_ids")]
    pub fn story_paragraph_ids(&self, env: Env, story: String) -> Result<String> {
        self.document()?
            .story_paragraph_ids(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "story_segment_unit_digests")]
    pub fn story_segment_unit_digests(&self, env: Env, story: String) -> Result<String> {
        self.document()?
            .story_segment_unit_digests(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "story_segment_units")]
    pub fn story_segment_units(
        &self,
        env: Env,
        story: String,
        units: Uint32Array,
    ) -> Result<String> {
        self.document()?
            .story_segment_units(&story, units.to_vec())
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "story_segments")]
    pub fn story_segments(&self, env: Env, story: String) -> Result<String> {
        self.document()?
            .story_segments(&story)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "table_payload")]
    pub fn table_payload(
        &self,
        env: Env,
        story: String,
        table_index: u32,
    ) -> Result<Either<String, Undefined>> {
        self.document()?
            .table_payload(&story, table_index)
            .map_err(|error| map_error(env, error))
            .map(optional_result)
    }

    #[napi(js_name = "toggle_mark")]
    pub fn toggle_mark(
        &self,
        env: Env,
        story: String,
        start_para: String,
        start_offset: u32,
        end_para: String,
        end_offset: u32,
        mark_json: String,
    ) -> Result<()> {
        self.document()?
            .toggle_mark(
                &story,
                &start_para,
                start_offset,
                &end_para,
                end_offset,
                &mark_json,
            )
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "track_undo")]
    pub fn track_undo(&self) -> Result<()> {
        self.document()?.track_undo();
        Ok(())
    }

    #[napi(js_name = "undo_capture_mode")]
    pub fn undo_capture_mode(&self) -> Result<String> {
        Ok(self.document()?.undo_capture_mode())
    }

    #[napi(js_name = "undo")]
    pub fn undo(&self) -> Result<bool> {
        Ok(self.document()?.undo())
    }

    #[napi(js_name = "update_comment_metadata_json")]
    pub fn update_comment_metadata_json(
        &self,
        env: Env,
        comment_id: String,
        patch_json: String,
    ) -> Result<()> {
        self.document()?
            .update_comment_metadata_json(&comment_id, &patch_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "validate_edits_json")]
    pub fn validate_edits_json(&self, env: Env, request: String) -> Result<String> {
        self.document()?
            .validate_edits_json(&request)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "validate_table_fragment")]
    pub fn validate_table_fragment(&self, env: Env, fragment_json: String) -> Result<()> {
        self.document()?
            .validate_table_fragment(&fragment_json)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "version")]
    pub fn version(&self) -> Result<String> {
        Ok(self.document()?.version())
    }

    #[napi(js_name = "written_paragraph_ids")]
    pub fn written_paragraph_ids(&self, env: Env, bytes: Uint8Array) -> Result<String> {
        self.document()?
            .written_paragraph_ids(&bytes)
            .map_err(|error| map_error(env, error))
    }

    #[napi(js_name = "yrs_blocks_for_story")]
    pub fn yrs_blocks_for_story(
        &self,
        env: Env,
        story: String,
        env_json: String,
    ) -> Result<String> {
        self.document()?
            .yrs_blocks_for_story(&story, &env_json)
            .map_err(|error| map_error(env, error))
    }
}

#[napi(js_name = "RetainedLayoutMeta")]
pub struct RetainedLayoutMeta {
    inner: Option<docx_engine::RetainedLayoutMeta>,
}

impl RetainedLayoutMeta {
    fn meta(&self) -> Result<&docx_engine::RetainedLayoutMeta> {
        self.inner
            .as_ref()
            .ok_or_else(|| Error::from_reason("RetainedLayoutMeta was freed"))
    }
}

#[napi]
impl RetainedLayoutMeta {
    #[napi(js_name = "free")]
    pub fn free(&mut self) {
        self.inner = None;
    }

    #[napi(getter, js_name = "page_count")]
    pub fn page_count(&self) -> Result<u32> {
        Ok(self.meta()?.page_count as u32)
    }

    #[napi(getter, js_name = "partial")]
    pub fn partial(&self) -> Result<bool> {
        Ok(self.meta()?.partial)
    }

    #[napi(getter, js_name = "provisional")]
    pub fn provisional(&self) -> Result<bool> {
        Ok(self.meta()?.provisional)
    }

    #[napi(getter, js_name = "notes_converged")]
    pub fn notes_converged(&self) -> Result<bool> {
        Ok(self.meta()?.notes_converged)
    }

    #[napi(js_name = "page_sizes")]
    pub fn page_sizes(&self) -> Result<Float64Array> {
        Ok(Float64Array::new(self.meta()?.page_sizes.clone()))
    }

    #[napi(js_name = "layout_shell_json")]
    pub fn layout_shell_json(&self) -> Result<String> {
        Ok(self.meta()?.layout_shell_json.clone())
    }
}

#[napi(js_name = "RunInheritance")]
pub struct RunInheritance {
    inner: Option<docx_engine::RunInheritance>,
}

impl RunInheritance {
    fn inheritance(&self) -> Result<&docx_engine::RunInheritance> {
        self.inner
            .as_ref()
            .ok_or_else(|| Error::from_reason("RunInheritance was freed"))
    }

    fn inheritance_mut(&mut self) -> Result<&mut docx_engine::RunInheritance> {
        self.inner
            .as_mut()
            .ok_or_else(|| Error::from_reason("RunInheritance was freed"))
    }
}

#[napi]
impl RunInheritance {
    #[napi(js_name = "free")]
    pub fn free(&mut self) {
        self.inner = None;
    }

    #[napi(js_name = "attributes")]
    pub fn attributes(
        &mut self,
        env: Env,
        paragraph_style: Option<String>,
        run_style: Option<String>,
        table_run_json: Option<String>,
    ) -> Result<String> {
        let attrs = if let Some(table_run_json) = table_run_json {
            let table_run = serde_json::from_str(&table_run_json)
                .map_err(|error| map_error(env, EngineError::message(error)))?;
            self.inheritance_mut()?.table_run_attrs(
                paragraph_style.as_deref(),
                run_style.as_deref(),
                table_run,
            )
        } else {
            self.inheritance()?
                .run_attrs(paragraph_style.as_deref(), run_style.as_deref())
        };
        serde_json::to_string(&attrs).map_err(|error| map_error(env, EngineError::message(error)))
    }

    #[napi(
        js_name = "paragraph_formatting",
        ts_args_type = "paragraph_style: string | null | undefined, table_formatting_json: string"
    )]
    pub fn paragraph_formatting(
        &mut self,
        env: Env,
        paragraph_style: Option<String>,
        table_formatting_json: String,
    ) -> Result<String> {
        let table_formatting = serde_json::from_str(&table_formatting_json)
            .map_err(|error| map_error(env, EngineError::message(error)))?;
        let formatting = self
            .inheritance_mut()?
            .paragraph_formatting(paragraph_style.as_deref(), table_formatting);
        serde_json::to_string(&formatting)
            .map_err(|error| map_error(env, EngineError::message(error)))
    }

    #[napi(js_name = "table_attributes")]
    pub fn table_attributes(&self, env: Env, table_json: String) -> Result<String> {
        let table = serde_json::from_str(&table_json)
            .map_err(|error| map_error(env, EngineError::message(error)))?;
        serde_json::to_string(&self.inheritance()?.table_attrs(&table))
            .map_err(|error| map_error(env, EngineError::message(error)))
    }
}

#[napi(js_name = "export_docx_structured_json")]
pub fn export_docx_structured_json(env: Env, bytes: Uint8Array, options: String) -> Result<String> {
    docx_engine::export_docx_structured_json(&bytes, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "export_docx_markdown_json")]
pub fn export_docx_markdown_json(env: Env, bytes: Uint8Array, options: String) -> Result<String> {
    docx_engine::export_docx_markdown_json(&bytes, &options).map_err(|error| map_error(env, error))
}

#[napi(js_name = "list_docx_content_controls_json")]
pub fn list_docx_content_controls_json(
    env: Env,
    bytes: Uint8Array,
    options: String,
) -> Result<String> {
    docx_engine::list_docx_content_controls_json(&bytes, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "find_docx_content_controls_json")]
pub fn find_docx_content_controls_json(
    env: Env,
    bytes: Uint8Array,
    query: String,
    options: String,
) -> Result<String> {
    docx_engine::find_docx_content_controls_json(&bytes, &query, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "render_docx_markdown_json")]
pub fn render_docx_markdown_json(env: Env, content: String, options: String) -> Result<String> {
    docx_engine::render_docx_markdown_json(&content, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "render_docx_markdown_with_pages_json")]
pub fn render_docx_markdown_with_pages_json(
    env: Env,
    content: String,
    options: String,
) -> Result<String> {
    docx_engine::render_docx_markdown_with_pages_json(&content, &options)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "build_display_list_json")]
pub fn build_display_list_json(env: Env, input: String) -> Result<String> {
    docx_engine::build_display_list_json(&input).map_err(|error| map_error(env, error))
}

#[napi(js_name = "layout_document_json")]
pub fn layout_document_json(env: Env, input: String) -> Result<String> {
    docx_engine::layout_document_json(&input).map_err(|error| map_error(env, error))
}

#[napi(js_name = "measure_paragraph_json")]
pub fn measure_paragraph_json(env: Env, input: String) -> Result<String> {
    docx_engine::measure_paragraph_json(&input).map_err(|error| map_error(env, error))
}

#[napi(js_name = "register_measure_font")]
pub fn register_measure_font(env: Env, bytes: Uint8Array) -> Result<u32> {
    docx_engine::register_measure_font(&bytes).map_err(|error| map_error(env, error))
}

#[napi(js_name = "register_substitute_measure_font")]
pub fn register_substitute_measure_font(
    env: Env,
    base: u32,
    requested_family: String,
) -> Result<u32> {
    docx_engine::register_substitute_measure_font(base, &requested_family)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "clear_measure_fonts")]
pub fn clear_measure_fonts(env: Env) -> Result<()> {
    docx_engine::clear_measure_fonts().map_err(|error| map_error(env, error))
}

#[napi(js_name = "outline_glyph_json")]
pub fn outline_glyph_json(env: Env, font_id: u32, glyph_id: u32) -> Result<String> {
    docx_engine::outline_glyph_json(font_id, glyph_id).map_err(|error| map_error(env, error))
}

#[napi(js_name = "open_display_list")]
pub fn open_display_list(env: Env, display_list: String) -> Result<u32> {
    docx_engine::open_display_list(&display_list).map_err(|error| map_error(env, error))
}

#[napi(js_name = "close_display_list")]
pub fn close_display_list(handle: u32) {
    docx_engine::close_display_list(handle)
}

#[napi(js_name = "update_display_list")]
pub fn update_display_list(env: Env, handle: u32, update: String) -> Result<()> {
    docx_engine::update_display_list(handle, &update).map_err(|error| map_error(env, error))
}

#[napi(js_name = "hit_test_json")]
pub fn hit_test_json(
    env: Env,
    display_list: String,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String> {
    docx_engine::hit_test_json(&display_list, page_index, x, y)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "hit_test_regions_json")]
pub fn hit_test_regions_json(
    env: Env,
    display_list: String,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String> {
    docx_engine::hit_test_regions_json(&display_list, page_index, x, y)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "hit_test_regions_by_handle")]
pub fn hit_test_regions_by_handle(
    env: Env,
    handle: u32,
    page_index: u32,
    x: f64,
    y: f64,
) -> Result<String> {
    docx_engine::hit_test_regions_by_handle(handle, page_index, x, y)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "range_rects_json")]
pub fn range_rects_json(env: Env, display_list: String, from: f64, to: f64) -> Result<String> {
    docx_engine::range_rects_json(&display_list, from, to).map_err(|error| map_error(env, error))
}

#[napi(js_name = "range_rects_by_handle")]
pub fn range_rects_by_handle(env: Env, handle: u32, from: f64, to: f64) -> Result<String> {
    docx_engine::range_rects_by_handle(handle, from, to).map_err(|error| map_error(env, error))
}

#[napi(js_name = "range_rects_on_pages_json")]
pub fn range_rects_on_pages_json(
    env: Env,
    display_list: String,
    from: f64,
    to: f64,
    first_page: f64,
    last_page: f64,
) -> Result<String> {
    docx_engine::range_rects_on_pages_json(&display_list, from, to, first_page, last_page)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "range_rects_on_pages_by_handle")]
pub fn range_rects_on_pages_by_handle(
    env: Env,
    handle: u32,
    from: f64,
    to: f64,
    first_page: f64,
    last_page: f64,
) -> Result<String> {
    docx_engine::range_rects_on_pages_by_handle(handle, from, to, first_page, last_page)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "range_rects_region_json")]
pub fn range_rects_region_json(
    env: Env,
    display_list: String,
    region: String,
    part_id: String,
    from: f64,
    to: f64,
) -> Result<String> {
    docx_engine::range_rects_region_json(&display_list, &region, &part_id, from, to)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "range_rects_region_by_handle")]
pub fn range_rects_region_by_handle(
    env: Env,
    handle: u32,
    region: String,
    part_id: String,
    from: f64,
    to: f64,
) -> Result<String> {
    docx_engine::range_rects_region_by_handle(handle, &region, &part_id, from, to)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "vertical_move_json")]
pub fn vertical_move_json(
    env: Env,
    display_list: String,
    position: f64,
    direction: String,
    goal_x: f64,
) -> Result<String> {
    docx_engine::vertical_move_json(&display_list, position, &direction, goal_x)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "vertical_move_by_handle")]
pub fn vertical_move_by_handle(
    env: Env,
    handle: u32,
    position: f64,
    direction: String,
    goal_x: f64,
) -> Result<String> {
    docx_engine::vertical_move_by_handle(handle, position, &direction, goal_x)
        .map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_relationships")]
pub fn parse_docx_relationships(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_relationships(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s2")]
pub fn parse_docx_s2(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_s2(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s3")]
pub fn parse_docx_s3(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_s3(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s4")]
pub fn parse_docx_s4(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_s4(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s5")]
pub fn parse_docx_s5(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_s5(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s6")]
pub fn parse_docx_s6(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_s6(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s7")]
pub fn parse_docx_s7(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_s7(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s8")]
pub fn parse_docx_s8(env: Env, data: Uint8Array) -> Result<String> {
    docx_engine::parse_docx_s8(&data).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_docx_s9")]
pub fn parse_docx_s9(env: Env, data: Uint8Array, options_json: String) -> Result<String> {
    docx_engine::parse_docx_s9(&data, &options_json).map_err(|error| map_error(env, error))
}

#[napi(js_name = "parse_relationships_xml")]
pub fn parse_relationships_xml(env: Env, xml: Uint8Array, part_path: String) -> Result<String> {
    docx_engine::parse_relationships_xml(&xml, &part_path).map_err(|error| map_error(env, error))
}

#[napi(js_name = "serialize_docx_s10")]
pub fn serialize_docx_s10(env: Env, request_json: String) -> Result<String> {
    docx_engine::serialize_docx_s10(&request_json).map_err(|error| map_error(env, error))
}

#[napi(js_name = "serialize_docx_s11")]
pub fn serialize_docx_s11(env: Env, request_json: String) -> Result<String> {
    docx_engine::serialize_docx_s11(&request_json).map_err(|error| map_error(env, error))
}

#[napi(js_name = "serialize_docx_s12")]
pub fn serialize_docx_s12(env: Env, request_json: String) -> Result<String> {
    docx_engine::serialize_docx_s12(&request_json).map_err(|error| map_error(env, error))
}

#[napi(js_name = "paragraphPatchWriterVersion")]
pub fn paragraph_patch_writer_version() -> u32 {
    docx_engine::paragraph_patch_writer_version()
}

#[napi(js_name = "write_docx_s13_wasm")]
pub fn write_docx_s13_wasm(
    env: Env,
    request_json: String,
    original_docx: Uint8Array,
) -> Result<Uint8Array> {
    docx_engine::write_docx_s13_wasm(&request_json, &original_docx)
        .map_err(|error| map_error(env, error))
        .and_then(|bytes| byte_result(env, bytes))
}

#[napi(js_name = "write_docx_s13_wasm_with_warnings")]
pub fn write_docx_s13_wasm_with_warnings(
    env: Env,
    request_json: String,
    original_docx: Uint8Array,
) -> Result<Uint8Array> {
    docx_engine::write_docx_s13_wasm_with_warnings(&request_json, &original_docx)
        .map_err(|error| map_error(env, error))
        .and_then(|bytes| byte_result(env, bytes))
}

#[napi(js_name = "decodeTiffPng")]
pub fn decode_tiff_png(env: Env, data: Uint8Array) -> Result<Uint8Array> {
    docx_engine::decode_tiff_png(&data)
        .map_err(|error| map_error(env, error))
        .and_then(|bytes| byte_result(env, bytes))
}

#[napi(js_name = "sanitizeOoxml")]
pub fn sanitize_ooxml(env: Env, data: Uint8Array, expected_format: String) -> Result<Uint8Array> {
    docx_engine::sanitize_ooxml(&data, &expected_format)
        .map_err(|error| map_error(env, error))
        .and_then(|bytes| byte_result(env, bytes))
}

#[napi(js_name = "unzip_docx", ts_return_type = "any")]
pub fn unzip_docx<'env>(env: Env, data: Uint8Array) -> Result<Object<'env>> {
    let parts = docx_engine::unzip_docx(&data).map_err(|error| map_error(env, error))?;
    let mut output = Object::new(&env)?;
    for (name, bytes) in parts {
        output.set(name, byte_result(env, bytes)?)?;
    }
    Ok(output)
}

#[napi(js_name = "rezip_docx", ts_args_type = "entries: any")]
pub fn rezip_docx(env: Env, entries: Unknown<'_>) -> Result<Uint8Array> {
    if !matches!(entries.get_type()?, ValueType::Object | ValueType::Function) {
        return Err(map_error(
            env,
            EngineError::message("rezip_docx: expected an object"),
        ));
    }
    let object: Object<'_> = unsafe { entries.cast()? };
    let global = env.get_global()?;
    let constructor: Function<'_, (), ()> = global.get_named_property("Object")?;
    let keys: Function<'_, Unknown<'_>, Array<'_>> = constructor.get_named_property("keys")?;
    let keys = keys.call(entries)?;
    let array: Function<'_, Unknown<'_>, ()> = global.get_named_property("Uint8Array")?;
    let mut collected = Vec::with_capacity(keys.len() as usize);
    for index in 0..keys.len() {
        let key = keys
            .get::<Unknown<'_>>(index)?
            .ok_or_else(|| map_error(env, EngineError::message("rezip_docx: non-string key")))?;
        if key.get_type()? != ValueType::String {
            return Err(map_error(
                env,
                EngineError::message("rezip_docx: non-string key"),
            ));
        }
        let name: String = unsafe { key.cast()? };
        let value: Unknown<'_> = object.get_property(key)?;
        let bytes: Uint8Array = unsafe { array.new_instance(value)?.cast()? };
        collected.push((name, bytes.to_vec()));
    }
    docx_engine::rezip_docx(collected)
        .map_err(|error| map_error(env, error))
        .and_then(|bytes| byte_result(env, bytes))
}

/// Native no-op; native allocations are not counted.
#[napi(js_name = "wasm_heap_counted")]
pub fn wasm_heap_counted() -> bool {
    false
}

/// Native no-op; native allocations are not counted.
#[napi(js_name = "wasm_live_bytes")]
pub fn wasm_live_bytes() -> f64 {
    0.0
}

/// Native no-op; native allocations are not counted.
#[napi(js_name = "wasm_peak_bytes")]
pub fn wasm_peak_bytes() -> f64 {
    0.0
}

/// Native no-op; native allocations are not counted.
#[napi(js_name = "wasm_failed_allocation_bytes")]
pub fn wasm_failed_allocation_bytes() -> f64 {
    0.0
}

/// Native no-op; native allocations are not counted.
#[napi(js_name = "reset_wasm_peak_bytes")]
pub fn reset_wasm_peak_bytes() {}

/// Native no-op; native allocations are not limited.
#[napi(js_name = "set_wasm_heap_limit")]
pub fn set_wasm_heap_limit(_bytes: f64) {}

#[napi(js_name = "install_panic_hook")]
pub fn install_panic_hook() {}
