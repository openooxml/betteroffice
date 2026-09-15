#![deny(clippy::all)]

use std::collections::BTreeMap;
use std::sync::mpsc::{self, Receiver, Sender};

use betteroffice_pptx::{
    Background, CommentFlavor, EditCtx, EditOrigin, MAX_COLLABORATION_CLIENT_ID,
    Presentation as CorePresentation, RenderOptions,
};
use napi::bindgen_prelude::{AsyncTask, Buffer, Task, ToNapiValue, TypeName};
use napi::{Env, Error, Result, ValueType, sys};
use napi_derive::napi;
use serde::Serialize;

fn error(reason: impl ToString) -> Error {
    Error::from_reason(reason.to_string())
}

type Job = Box<dyn FnOnce(&mut CorePresentation) + Send>;

#[derive(Clone)]
pub struct Worker {
    sender: Sender<Job>,
}

impl Worker {
    fn submit<T, F>(&self, operation: F) -> Result<AsyncTask<PendingTask<T>>>
    where
        T: ToNapiValue + TypeName + Send + 'static,
        F: FnOnce(&mut CorePresentation) -> Result<T> + Send + 'static,
    {
        let (reply, receive) = mpsc::channel();
        self.sender
            .send(Box::new(move |presentation| {
                let _ = reply.send(operation(presentation));
            }))
            .map_err(|_| error("presentation worker stopped"))?;
        Ok(AsyncTask::new(PendingTask { receive }))
    }
}

pub struct PendingTask<T> {
    receive: Receiver<Result<T>>,
}

impl<T> Task for PendingTask<T>
where
    T: ToNapiValue + TypeName + Send + 'static,
{
    type Output = T;
    type JsValue = T;

    fn compute(&mut self) -> Result<Self::Output> {
        self.receive
            .recv()
            .map_err(|_| error("presentation worker stopped"))?
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }
}

pub struct JsonValue(serde_json::Value);

impl TypeName for JsonValue {
    fn type_name() -> &'static str {
        "any"
    }

    fn value_type() -> ValueType {
        ValueType::Unknown
    }
}

impl ToNapiValue for JsonValue {
    unsafe fn to_napi_value(env: sys::napi_env, value: Self) -> Result<sys::napi_value> {
        unsafe { serde_json::Value::to_napi_value(env, value.0) }
    }
}

fn json<T: Serialize>(value: T) -> Result<JsonValue> {
    serde_json::to_value(value).map(JsonValue).map_err(error)
}

fn parse<T: serde::de::DeserializeOwned>(value: serde_json::Value) -> Result<T> {
    serde_json::from_value(value).map_err(error)
}

fn client_id(value: f64) -> Result<u64> {
    if !value.is_finite()
        || value.fract() != 0.0
        || value < 1.0
        || value > MAX_COLLABORATION_CLIENT_ID as f64
    {
        return Err(error("clientId must be a positive safe integer"));
    }
    Ok(value as u64)
}

fn max_shadow_pixels(value: f64) -> Result<u64> {
    if !value.is_finite() || value.fract() != 0.0 || value < 0.0 || value > 9_007_199_254_740_991.0
    {
        return Err(error("maxShadowPixels must be a non-negative safe integer"));
    }
    Ok(value as u64)
}

fn origin(value: &str) -> Result<EditOrigin> {
    match value {
        "local" => Ok(EditOrigin::Local),
        "agent" => Ok(EditOrigin::Agent),
        "remote" => Ok(EditOrigin::Remote),
        "system" => Ok(EditOrigin::System),
        _ => Err(error("origin must be local, agent, remote, or system")),
    }
}

fn origin_name(value: EditOrigin) -> &'static str {
    match value {
        EditOrigin::Local => "local",
        EditOrigin::Agent => "agent",
        EditOrigin::Remote => "remote",
        EditOrigin::System => "system",
    }
}

fn comment_flavor(value: &str) -> Result<CommentFlavor> {
    match value {
        "legacy" => Ok(CommentFlavor::Legacy),
        "modern" => Ok(CommentFlavor::Modern),
        _ => Err(error("commentFlavor must be legacy or modern")),
    }
}

fn comment_flavor_name(value: CommentFlavor) -> &'static str {
    match value {
        CommentFlavor::Legacy => "legacy",
        CommentFlavor::Modern => "modern",
    }
}

#[napi(object)]
pub struct OpenPresentationOptions {
    pub client_id: Option<f64>,
    pub author: Option<String>,
    pub origin: Option<String>,
}

#[napi(object)]
pub struct FontFace {
    pub family: String,
    pub data: Buffer,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
}

#[napi(object)]
pub struct RenderSlideOptions {
    pub scale: Option<f64>,
    pub transparent: Option<bool>,
    pub background: Option<String>,
    pub max_shadow_pixels: Option<f64>,
}

#[napi(object)]
pub struct RenderedSlide {
    pub data: Buffer,
    pub width: u32,
    pub height: u32,
    pub skipped_images: u32,
}

#[napi(object)]
pub struct MediaResource {
    pub path: String,
    pub content_type: String,
    pub data: Buffer,
}

#[napi(object)]
pub struct CommentInput {
    pub slide_id: String,
    pub text: String,
    pub author: String,
    pub initials: Option<String>,
    pub created: String,
    pub x: Option<i64>,
    pub y: Option<i64>,
}

#[napi(object)]
pub struct CommentReplyInput {
    pub comment_id: String,
    pub text: String,
    pub author: String,
    pub initials: Option<String>,
    pub created: String,
}

pub struct OpenTask {
    bytes: Vec<u8>,
    client_id: Option<u64>,
    author: String,
    origin: EditOrigin,
}

impl Task for OpenTask {
    type Output = Worker;
    type JsValue = PptxPresentation;

    fn compute(&mut self) -> Result<Self::Output> {
        let bytes = std::mem::take(&mut self.bytes);
        let client_id = self.client_id;
        let (sender, receive) = mpsc::channel::<Job>();
        let (ready, opened) = mpsc::channel();
        std::thread::Builder::new()
            .name("betteroffice-pptx".to_owned())
            .spawn(move || {
                let presentation = match client_id {
                    Some(client_id) => CorePresentation::open_collaborative(&bytes, client_id),
                    None => CorePresentation::open(&bytes),
                };
                let mut presentation = match presentation {
                    Ok(value) => {
                        let _ = ready.send(Ok(()));
                        value
                    }
                    Err(value) => {
                        let _ = ready.send(Err(value.to_string()));
                        return;
                    }
                };
                for job in receive {
                    job(&mut presentation);
                }
            })
            .map_err(error)?;
        opened
            .recv()
            .map_err(|_| error("presentation worker failed to start"))?
            .map_err(error)?;
        Ok(Worker { sender })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(PptxPresentation {
            worker: output,
            author: self.author.clone(),
            origin: self.origin,
            collaborative: self.client_id.is_some(),
        })
    }
}

#[napi(js_name = "Presentation")]
pub struct PptxPresentation {
    worker: Worker,
    author: String,
    origin: EditOrigin,
    collaborative: bool,
}

impl PptxPresentation {
    fn context(&self) -> EditCtx {
        EditCtx {
            origin: self.origin,
            author: self.author.clone(),
        }
    }
}

#[napi]
impl PptxPresentation {
    #[napi(getter)]
    pub fn author(&self) -> String {
        self.author.clone()
    }

    #[napi(setter)]
    pub fn set_author(&mut self, author: String) {
        self.author = author;
    }

    #[napi(getter)]
    pub fn origin(&self) -> &'static str {
        origin_name(self.origin)
    }

    #[napi(setter)]
    pub fn set_origin(&mut self, value: String) -> Result<()> {
        self.origin = origin(&value)?;
        Ok(())
    }

    #[napi(getter)]
    pub fn collaborative(&self) -> bool {
        self.collaborative
    }

    #[napi(getter, ts_return_type = "Promise<number>")]
    pub fn client_id(&self) -> Result<AsyncTask<PendingTask<f64>>> {
        self.worker
            .submit(|presentation| Ok(presentation.client_id() as f64))
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn snapshot(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker
            .submit(|presentation| presentation.snapshot().map_err(error).and_then(json))
    }

    #[napi(getter, ts_return_type = "Promise<number>")]
    pub fn slide_count(&self) -> Result<AsyncTask<PendingTask<u32>>> {
        self.worker.submit(|presentation| {
            presentation
                .snapshot()
                .map(|snapshot| snapshot.slides.len() as u32)
                .map_err(error)
        })
    }

    #[napi(getter, ts_return_type = "Promise<string[]>")]
    pub fn slide_ids(&self) -> Result<AsyncTask<PendingTask<Vec<String>>>> {
        self.worker.submit(|presentation| {
            presentation
                .snapshot()
                .map(|snapshot| {
                    snapshot
                        .slides
                        .into_iter()
                        .map(|slide| slide.id)
                        .collect::<Vec<_>>()
                })
                .map_err(error)
        })
    }

    #[napi(getter, ts_return_type = "Promise<number>")]
    pub fn width_emu(&self) -> Result<AsyncTask<PendingTask<i64>>> {
        self.worker.submit(|presentation| {
            presentation
                .snapshot()
                .map(|snapshot| snapshot.width_emu)
                .map_err(error)
        })
    }

    #[napi(getter, ts_return_type = "Promise<number>")]
    pub fn height_emu(&self) -> Result<AsyncTask<PendingTask<i64>>> {
        self.worker.submit(|presentation| {
            presentation
                .snapshot()
                .map(|snapshot| snapshot.height_emu)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn slide(&self, slide: u32) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(move |presentation| {
            let snapshot = presentation.snapshot().map_err(error)?;
            snapshot
                .slides
                .get(slide as usize)
                .ok_or_else(|| error(format!("slide index {slide} is out of range")))
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn story(&self, story_id: String) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker
            .submit(move |presentation| presentation.story(&story_id).map_err(error).and_then(json))
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn layouts(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(|presentation| {
            json(
                presentation
                    .layouts()
                    .iter()
                    .map(|layout| layout.part_path.clone())
                    .collect::<Vec<_>>(),
            )
        })
    }

    #[napi(ts_return_type = "Promise<MediaResource[]>")]
    pub fn media(&self) -> Result<AsyncTask<PendingTask<Vec<MediaResource>>>> {
        self.worker.submit(|presentation| {
            Ok(presentation
                .media()
                .iter()
                .map(|part| MediaResource {
                    path: part.part_path.clone(),
                    content_type: part.content_type.clone(),
                    data: part.bytes.clone().into(),
                })
                .collect())
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn insert_slide(
        &self,
        index: u32,
        layout: Option<String>,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .insert_slide(&context, index, layout.as_deref())
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn delete_slide(&self, slide_id: String) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .delete_slide(&context, &slide_id)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn move_slide(
        &self,
        slide_id: String,
        index: u32,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .move_slide(&context, &slide_id, index)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<void>")]
    pub fn set_slide_notes(
        &self,
        slide_id: String,
        text: String,
    ) -> Result<AsyncTask<PendingTask<()>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .set_slide_notes(&context, &slide_id, &text)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn add_text_box(
        &self,
        slide_id: String,
        draft: serde_json::Value,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        let draft = parse(draft)?;
        self.worker.submit(move |presentation| {
            presentation
                .add_text_box(&context, &slide_id, &draft)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn add_shape(
        &self,
        slide_id: String,
        draft: serde_json::Value,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        let draft = parse(draft)?;
        self.worker.submit(move |presentation| {
            presentation
                .add_shape(&context, &slide_id, &draft)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn remove_shape(
        &self,
        slide_id: String,
        shape_id: String,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .remove_shape(&context, &slide_id, &shape_id)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn set_shape_fill(
        &self,
        slide_id: String,
        shape_id: String,
        color: Option<String>,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .set_shape_fill(&context, &slide_id, &shape_id, color.as_deref())
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn set_shape_stroke(
        &self,
        slide_id: String,
        shape_id: String,
        stroke: serde_json::Value,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        let stroke = parse(stroke)?;
        self.worker.submit(move |presentation| {
            presentation
                .set_shape_stroke(&context, &slide_id, &shape_id, &stroke)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn set_shape_adjust(
        &self,
        slide_id: String,
        shape_id: String,
        adjustments: serde_json::Value,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        let adjustments: BTreeMap<String, f64> = parse(adjustments)?;
        self.worker.submit(move |presentation| {
            presentation
                .set_shape_adjust(&context, &slide_id, &shape_id, &adjustments)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn move_shape(
        &self,
        slide_id: String,
        shape_id: String,
        x: i64,
        y: i64,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .move_shape(&context, &slide_id, &shape_id, x, y)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn resize_shape(
        &self,
        slide_id: String,
        shape_id: String,
        width: i64,
        height: i64,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .resize_shape(&context, &slide_id, &shape_id, width, height)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn set_shape_rect(
        &self,
        slide_id: String,
        shape_id: String,
        rect: serde_json::Value,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        let rect = parse(rect)?;
        self.worker.submit(move |presentation| {
            presentation
                .set_shape_rect(&context, &slide_id, &shape_id, rect)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn insert_text(
        &self,
        story_id: String,
        index: u32,
        text: String,
        style: Option<serde_json::Value>,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        let style = style.map(parse).transpose()?.unwrap_or_default();
        self.worker.submit(move |presentation| {
            presentation
                .insert_text(&context, &story_id, index, &text, &style)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn delete_text(
        &self,
        story_id: String,
        start: u32,
        end: u32,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .delete_text(&context, &story_id, start, end)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn format_text(
        &self,
        story_id: String,
        start: u32,
        end: u32,
        patch: serde_json::Value,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        let patch = parse(patch)?;
        self.worker.submit(move |presentation| {
            presentation
                .format_text(&context, &story_id, start, end, &patch)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn set_paragraph_alignment(
        &self,
        story_id: String,
        start: u32,
        end: u32,
        alignment: Option<String>,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .set_paragraph_alignment(&context, &story_id, start, end, alignment.as_deref())
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn insert_paragraph_break(
        &self,
        story_id: String,
        index: u32,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .insert_paragraph_break(&context, &story_id, index)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn delete_paragraph_break(
        &self,
        story_id: String,
        index: u32,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .delete_paragraph_break(&context, &story_id, index)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn add_comment(&self, comment: CommentInput) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .add_comment(
                    &context,
                    &comment.slide_id,
                    &comment.author,
                    comment.initials.as_deref().unwrap_or_default(),
                    &comment.text,
                    &comment.created,
                    comment.x.unwrap_or(0),
                    comment.y.unwrap_or(0),
                )
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn reply_to_comment(
        &self,
        reply: CommentReplyInput,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .reply_to_comment(
                    &context,
                    &reply.comment_id,
                    &reply.author,
                    reply.initials.as_deref().unwrap_or_default(),
                    &reply.text,
                    &reply.created,
                )
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn set_comment_status(
        &self,
        comment_id: String,
        resolved: Option<bool>,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .set_comment_status(&context, &comment_id, resolved.unwrap_or(true))
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn remove_comment(&self, comment_id: String) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let context = self.context();
        self.worker.submit(move |presentation| {
            presentation
                .remove_comment(&context, &comment_id)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(getter, ts_return_type = "Promise<any>")]
    pub fn comments(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker
            .submit(|presentation| presentation.comments().map_err(error).and_then(json))
    }

    #[napi(getter, ts_return_type = "Promise<string>")]
    pub fn comment_flavor(&self) -> Result<AsyncTask<PendingTask<String>>> {
        self.worker.submit(|presentation| {
            presentation
                .comment_flavor()
                .map(|value| comment_flavor_name(value).to_owned())
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<string>")]
    pub fn set_comment_flavor(&self, flavor: String) -> Result<AsyncTask<PendingTask<String>>> {
        let context = self.context();
        let flavor = comment_flavor(&flavor)?;
        self.worker.submit(move |presentation| {
            presentation
                .set_comment_flavor(&context, flavor)
                .map(|value| comment_flavor_name(value).to_owned())
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<number>")]
    pub fn register_font(&self, face: FontFace) -> Result<AsyncTask<PendingTask<u32>>> {
        let data = face.data.to_vec();
        self.worker.submit(move |presentation| {
            presentation
                .register_font(
                    &face.family,
                    face.bold.unwrap_or(false),
                    face.italic.unwrap_or(false),
                    &data,
                )
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<RenderedSlide>")]
    pub fn render_slide(
        &self,
        slide: u32,
        options: Option<RenderSlideOptions>,
    ) -> Result<AsyncTask<PendingTask<RenderedSlide>>> {
        let options = options.unwrap_or(RenderSlideOptions {
            scale: None,
            transparent: None,
            background: None,
            max_shadow_pixels: None,
        });
        let background = if options.transparent.unwrap_or(false) {
            Background::Transparent
        } else if let Some(color) = options.background {
            Background::Color(color)
        } else {
            Background::Slide
        };
        self.worker.submit(move |presentation| {
            let defaults = RenderOptions::default();
            let options = RenderOptions {
                scale: options.scale.unwrap_or(1.0) as f32,
                background,
                max_shadow_pixels: options
                    .max_shadow_pixels
                    .map(max_shadow_pixels)
                    .transpose()?
                    .unwrap_or(defaults.max_shadow_pixels),
            };
            let rendered = presentation
                .render_png(slide as usize, &options)
                .map_err(error)?;
            Ok(RenderedSlide {
                data: rendered.bytes.into(),
                width: rendered.width,
                height: rendered.height,
                skipped_images: rendered.skipped_images as u32,
            })
        })
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn encode_state_vector(&self) -> Result<AsyncTask<PendingTask<Buffer>>> {
        let collaborative = self.collaborative;
        self.worker.submit(move |presentation| {
            if !collaborative {
                return Err(error("operation requires a collaborative presentation"));
            }
            Ok(presentation.encode_state_vector_v1().into())
        })
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn encode_state_as_update(&self) -> Result<AsyncTask<PendingTask<Buffer>>> {
        let collaborative = self.collaborative;
        self.worker.submit(move |presentation| {
            if !collaborative {
                return Err(error("operation requires a collaborative presentation"));
            }
            Ok(presentation.encode_state_as_update_v1().into())
        })
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn encode_diff(&self, state_vector: Buffer) -> Result<AsyncTask<PendingTask<Buffer>>> {
        let collaborative = self.collaborative;
        let state_vector = state_vector.to_vec();
        self.worker.submit(move |presentation| {
            if !collaborative {
                return Err(error("operation requires a collaborative presentation"));
            }
            presentation
                .encode_diff_v1(&state_vector)
                .map(Buffer::from)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn apply_update(&self, update: Buffer) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let collaborative = self.collaborative;
        let update = update.to_vec();
        self.worker.submit(move |presentation| {
            if !collaborative {
                return Err(error("operation requires a collaborative presentation"));
            }
            presentation
                .apply_update_v1(&update)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn propose(&self, request: serde_json::Value) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        let request = parse(request)?;
        self.worker
            .submit(move |presentation| presentation.propose(request).map_err(error).and_then(json))
    }

    #[napi(getter, ts_return_type = "Promise<any>")]
    pub fn proposals(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker
            .submit(|presentation| presentation.proposals().map_err(error).and_then(json))
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn preview_proposal(
        &self,
        proposal_id: String,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(move |presentation| {
            presentation
                .preview_proposal(&proposal_id)
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn accept_proposal(
        &self,
        proposal_id: String,
        force: Option<bool>,
    ) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(move |presentation| {
            presentation
                .accept_proposal(&proposal_id, force.unwrap_or(false))
                .map_err(error)
                .and_then(json)
        })
    }

    #[napi(ts_return_type = "Promise<boolean>")]
    pub fn reject_proposal(&self, proposal_id: String) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker
            .submit(move |presentation| Ok(presentation.reject_proposal(&proposal_id)))
    }

    #[napi(getter, ts_return_type = "Promise<boolean>")]
    pub fn can_undo(&self) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker
            .submit(|presentation| Ok(presentation.can_undo()))
    }

    #[napi(getter, ts_return_type = "Promise<boolean>")]
    pub fn can_redo(&self) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker
            .submit(|presentation| Ok(presentation.can_redo()))
    }

    #[napi(ts_return_type = "Promise<boolean>")]
    pub fn undo(&self) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker.submit(|presentation| Ok(presentation.undo()))
    }

    #[napi(ts_return_type = "Promise<boolean>")]
    pub fn redo(&self) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker.submit(|presentation| Ok(presentation.redo()))
    }

    #[napi(ts_return_type = "Promise<void>")]
    pub fn add_undo_barrier(&self) -> Result<AsyncTask<PendingTask<()>>> {
        self.worker.submit(|presentation| {
            presentation.add_undo_barrier();
            Ok(())
        })
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn save(&self) -> Result<AsyncTask<PendingTask<Buffer>>> {
        self.worker
            .submit(|presentation| presentation.save().map(Buffer::from).map_err(error))
    }
}

#[napi(ts_return_type = "Promise<Presentation>")]
pub fn open_presentation(
    data: Buffer,
    options: Option<OpenPresentationOptions>,
) -> Result<AsyncTask<OpenTask>> {
    let options = options.unwrap_or(OpenPresentationOptions {
        client_id: None,
        author: None,
        origin: None,
    });
    Ok(AsyncTask::new(OpenTask {
        bytes: data.to_vec(),
        client_id: options.client_id.map(client_id).transpose()?,
        author: options.author.unwrap_or_else(|| "node".to_owned()),
        origin: options
            .origin
            .as_deref()
            .map(origin)
            .transpose()?
            .unwrap_or_default(),
    }))
}
