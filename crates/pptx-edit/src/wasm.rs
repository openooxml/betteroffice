use std::collections::VecDeque;
use std::sync::{Arc, Mutex};

use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use yrs::Subscription;

use crate::{
    DeckSession, DeckSnapshot, EditCtx, ShapeDraft, TextStyle, TextStylePatch, UpdateEvent,
    UpdateOrigin,
};

#[wasm_bindgen]
pub struct PptxDocument {
    session: DeckSession,
    update_observer: Option<UpdateObserver>,
}

struct UpdateObserver {
    pending: Arc<Mutex<VecDeque<UpdateEvent>>>,
    _subscription: Subscription,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoryArgs {
    story_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct InsertTextArgs {
    story_id: String,
    index: u32,
    text: String,
    #[serde(default)]
    style: TextStyle,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct DeleteTextArgs {
    story_id: String,
    start: u32,
    end: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct FormatTextArgs {
    story_id: String,
    start: u32,
    end: u32,
    #[serde(default)]
    patch: TextStylePatch,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ParagraphBreakArgs {
    story_id: String,
    index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct InsertSlideArgs {
    index: u32,
    layout_part_path: Option<String>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct SlideArgs {
    slide_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MoveSlideArgs {
    slide_id: String,
    to_index: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct AddTextBoxArgs {
    slide_id: String,
    draft: ShapeDraft,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ShapeArgs {
    slide_id: String,
    shape_id: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct MoveShapeArgs {
    slide_id: String,
    shape_id: String,
    x: i64,
    y: i64,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ResizeShapeArgs {
    slide_id: String,
    shape_id: String,
    width: i64,
    height: i64,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct HistoryResult {
    applied: bool,
    snapshot: DeckSnapshot,
}

#[wasm_bindgen]
impl PptxDocument {
    #[wasm_bindgen(js_name = openCollaborative)]
    pub fn open_collaborative(bytes: &[u8], client_id: f64) -> Result<PptxDocument, JsValue> {
        let client_id = parse_client_id(client_id)?;
        DeckSession::open(bytes, client_id)
            .map(|session| Self {
                session,
                update_observer: None,
            })
            .map_err(js_error)
    }

    #[wasm_bindgen(getter, js_name = clientId)]
    pub fn client_id(&self) -> f64 {
        self.session.client_id() as f64
    }

    #[wasm_bindgen(js_name = snapshotJson)]
    pub fn snapshot_json(&self) -> Result<String, JsValue> {
        json(self.session.snapshot().map_err(js_error)?)
    }

    #[wasm_bindgen(js_name = storyJson)]
    pub fn story_json(&self, args: &str) -> Result<String, JsValue> {
        let args: StoryArgs = parse_args(args)?;
        json(self.session.story(&args.story_id).map_err(js_error)?)
    }

    #[wasm_bindgen(js_name = encodeStateVector)]
    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.session.encode_state_vector_v1()
    }

    #[wasm_bindgen(js_name = encodeStateAsUpdate)]
    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.session.encode_state_as_update_v1()
    }

    #[wasm_bindgen(js_name = encodeDiff)]
    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, JsValue> {
        self.session
            .encode_diff_v1(remote_state_vector)
            .map_err(js_error)
    }

    #[wasm_bindgen(js_name = applyUpdateJson)]
    pub fn apply_update_json(&self, update: &[u8]) -> Result<String, JsValue> {
        json(self.session.apply_update_v1(update).map_err(js_error)?)
    }

    #[wasm_bindgen(js_name = startUpdateObservation)]
    pub fn start_update_observation(&mut self) -> Result<(), JsValue> {
        if self.update_observer.is_some() {
            return Ok(());
        }
        let pending = Arc::new(Mutex::new(VecDeque::new()));
        let observed = Arc::clone(&pending);
        let subscription = self
            .session
            .observe_update_v1(move |event| {
                observed
                    .lock()
                    .unwrap_or_else(|poisoned| poisoned.into_inner())
                    .push_back(event);
            })
            .map_err(js_error)?;
        self.update_observer = Some(UpdateObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    #[wasm_bindgen(js_name = clearUpdateObservation)]
    pub fn clear_update_observation(&mut self) {
        self.update_observer = None;
    }

    #[wasm_bindgen(js_name = drainUpdateEvent)]
    pub fn drain_update_event(&self) -> Vec<u8> {
        let Some(observer) = &self.update_observer else {
            return Vec::new();
        };
        let event = observer
            .pending
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .pop_front();
        let Some(event) = event else {
            return Vec::new();
        };
        let mut encoded = Vec::with_capacity(event.update.len() + 1);
        encoded.push(match event.origin {
            UpdateOrigin::Local => 0,
            UpdateOrigin::Remote => 1,
        });
        encoded.extend_from_slice(&event.update);
        encoded
    }

    #[wasm_bindgen(js_name = insertTextJson)]
    pub fn insert_text_json(&self, args: &str) -> Result<String, JsValue> {
        let args: InsertTextArgs = parse_args(args)?;
        json(
            self.session
                .insert_text(
                    &local_context(),
                    &args.story_id,
                    args.index,
                    &args.text,
                    &args.style,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = deleteTextJson)]
    pub fn delete_text_json(&self, args: &str) -> Result<String, JsValue> {
        let args: DeleteTextArgs = parse_args(args)?;
        json(
            self.session
                .delete_text(&local_context(), &args.story_id, args.start, args.end)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = formatTextJson)]
    pub fn format_text_json(&self, args: &str) -> Result<String, JsValue> {
        let args: FormatTextArgs = parse_args(args)?;
        json(
            self.session
                .format_text(
                    &local_context(),
                    &args.story_id,
                    args.start,
                    args.end,
                    &args.patch,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = insertParagraphBreakJson)]
    pub fn insert_paragraph_break_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ParagraphBreakArgs = parse_args(args)?;
        json(
            self.session
                .insert_paragraph_break(&local_context(), &args.story_id, args.index)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = insertSlideJson)]
    pub fn insert_slide_json(&self, args: &str) -> Result<String, JsValue> {
        let args: InsertSlideArgs = parse_args(args)?;
        json(
            self.session
                .insert_slide(
                    &local_context(),
                    args.index,
                    args.layout_part_path.as_deref(),
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = deleteSlideJson)]
    pub fn delete_slide_json(&self, args: &str) -> Result<String, JsValue> {
        let args: SlideArgs = parse_args(args)?;
        json(
            self.session
                .delete_slide(&local_context(), &args.slide_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = moveSlideJson)]
    pub fn move_slide_json(&self, args: &str) -> Result<String, JsValue> {
        let args: MoveSlideArgs = parse_args(args)?;
        json(
            self.session
                .move_slide(&local_context(), &args.slide_id, args.to_index)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = addTextBoxJson)]
    pub fn add_text_box_json(&self, args: &str) -> Result<String, JsValue> {
        let args: AddTextBoxArgs = parse_args(args)?;
        json(
            self.session
                .add_text_box(&local_context(), &args.slide_id, &args.draft)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = removeShapeJson)]
    pub fn remove_shape_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ShapeArgs = parse_args(args)?;
        json(
            self.session
                .remove_shape(&local_context(), &args.slide_id, &args.shape_id)
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = moveShapeJson)]
    pub fn move_shape_json(&self, args: &str) -> Result<String, JsValue> {
        let args: MoveShapeArgs = parse_args(args)?;
        json(
            self.session
                .move_shape(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    args.x,
                    args.y,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = resizeShapeJson)]
    pub fn resize_shape_json(&self, args: &str) -> Result<String, JsValue> {
        let args: ResizeShapeArgs = parse_args(args)?;
        json(
            self.session
                .resize_shape(
                    &local_context(),
                    &args.slide_id,
                    &args.shape_id,
                    args.width,
                    args.height,
                )
                .map_err(js_error)?,
        )
    }

    #[wasm_bindgen(js_name = undoJson)]
    pub fn undo_json(&self) -> Result<String, JsValue> {
        json(HistoryResult {
            applied: self.session.undo(),
            snapshot: self.session.snapshot().map_err(js_error)?,
        })
    }

    #[wasm_bindgen(js_name = redoJson)]
    pub fn redo_json(&self) -> Result<String, JsValue> {
        json(HistoryResult {
            applied: self.session.redo(),
            snapshot: self.session.snapshot().map_err(js_error)?,
        })
    }

    #[wasm_bindgen(js_name = canUndo)]
    pub fn can_undo(&self) -> bool {
        self.session.can_undo()
    }

    #[wasm_bindgen(js_name = canRedo)]
    pub fn can_redo(&self) -> bool {
        self.session.can_redo()
    }

    pub fn version() -> String {
        env!("CARGO_PKG_VERSION").to_owned()
    }
}

fn local_context() -> EditCtx {
    EditCtx::local("wasm")
}

fn parse_args<T: serde::de::DeserializeOwned>(args: &str) -> Result<T, JsValue> {
    serde_json::from_str(args).map_err(js_error)
}

fn json(value: impl Serialize) -> Result<String, JsValue> {
    serde_json::to_string(&value).map_err(js_error)
}

fn parse_client_id(client_id: f64) -> Result<u64, JsValue> {
    if !client_id.is_finite()
        || client_id.fract() != 0.0
        || client_id < 1.0
        || client_id > super::MAX_SAFE_CLIENT_ID as f64
    {
        return Err(JsValue::from_str(
            "client ID must be a positive safe integer below Number.MAX_SAFE_INTEGER",
        ));
    }
    Ok(client_id as u64)
}

fn js_error(error: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&error.to_string())
}
