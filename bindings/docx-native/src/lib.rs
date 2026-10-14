#![deny(clippy::all)]

use std::sync::mpsc::{self, Receiver, Sender};

use betteroffice_docx::{Document, EditCtx, EditOrigin, ImageScope, NoteKind, SaveOptions};
use napi::bindgen_prelude::{AsyncTask, Buffer, Task, ToNapiValue, TypeName};
use napi::{Env, Error, Result, ValueType, sys};
use napi_derive::napi;

fn error(reason: impl ToString) -> Error {
    Error::from_reason(reason.to_string())
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

#[napi(object)]
pub struct OpenDocumentOptions {
    pub author: Option<String>,
    pub origin: Option<String>,
    pub timestamp: Option<String>,
}

#[napi(object)]
pub struct FontFace {
    pub family: String,
    pub data: Buffer,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
}

#[napi(object)]
pub struct ImageResource {
    pub relationship_id: String,
    pub data: Buffer,
    pub scope: Option<String>,
    pub part: Option<String>,
}

#[napi(object)]
pub struct SaveDocumentOptions {
    pub timestamp: Option<String>,
    pub update_modified_date: Option<bool>,
    pub modified_by: Option<String>,
}

#[napi(object)]
pub struct RenderedPage {
    pub data: Buffer,
    pub skipped_images: u32,
}

#[napi(object)]
pub struct DocumentStructure {
    pub body_paragraphs: u32,
    pub body_tables: u32,
    pub sections: u32,
    pub headers: u32,
    pub footers: u32,
    pub footnotes: u32,
    pub endnotes: u32,
}

#[napi(object)]
pub struct EditReceipt {
    pub paragraph_id: Option<String>,
    pub story: Option<String>,
    pub start: Option<u32>,
    pub end: Option<u32>,
    pub new_paragraph_ids: Vec<String>,
    pub revision_ids: Vec<String>,
}

#[napi(object)]
pub struct LayoutResult {
    pub layout: serde_json::Value,
    pub display_list: serde_json::Value,
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

type Job = Box<dyn FnOnce(&mut Document) + Send>;

#[derive(Clone)]
pub struct Worker {
    sender: Sender<Job>,
}

impl Worker {
    fn start(mut document: Document) -> Result<Self> {
        let (sender, receive) = mpsc::channel::<Job>();
        std::thread::Builder::new()
            .name("betteroffice-docx".to_owned())
            .spawn(move || {
                for job in receive {
                    job(&mut document);
                }
            })
            .map_err(error)?;
        Ok(Self { sender })
    }

    fn submit<T, F>(&self, operation: F) -> Result<AsyncTask<PendingTask<T>>>
    where
        T: ToNapiValue + TypeName + Send + 'static,
        F: FnOnce(&mut Document) -> Result<T> + Send + 'static,
    {
        let (reply, receive) = mpsc::channel();
        self.sender
            .send(Box::new(move |document| {
                let _ = reply.send(operation(document));
            }))
            .map_err(|_| error("document worker stopped"))?;
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
            .map_err(|_| error("document worker stopped"))?
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }
}

pub struct OpenTask {
    bytes: Vec<u8>,
    author: String,
    origin: EditOrigin,
    timestamp: String,
}

impl Task for OpenTask {
    type Output = Worker;
    type JsValue = DocxDocument;

    fn compute(&mut self) -> Result<Self::Output> {
        Worker::start(Document::open(&self.bytes).map_err(error)?)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(DocxDocument {
            worker: output,
            author: self.author.clone(),
            origin: self.origin,
            timestamp: self.timestamp.clone(),
        })
    }
}

#[napi(js_name = "Document")]
pub struct DocxDocument {
    worker: Worker,
    author: String,
    origin: EditOrigin,
    timestamp: String,
}

#[napi]
impl DocxDocument {
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
    pub fn timestamp(&self) -> String {
        self.timestamp.clone()
    }

    #[napi(setter)]
    pub fn set_timestamp(&mut self, timestamp: String) {
        self.timestamp = timestamp;
    }

    #[napi(getter, ts_return_type = "Promise<Array<string | null>>")]
    pub fn paragraph_ids(&self) -> Result<AsyncTask<PendingTask<Vec<Option<String>>>>> {
        self.worker.submit(|document| {
            Ok(document
                .paragraphs()
                .into_iter()
                .map(|paragraph| paragraph.para_id.clone())
                .collect())
        })
    }

    #[napi(getter, ts_return_type = "Promise<string[]>")]
    pub fn warnings(&self) -> Result<AsyncTask<PendingTask<Vec<String>>>> {
        self.worker
            .submit(|document| Ok(document.model().warnings.clone()))
    }

    #[napi(getter, ts_return_type = "Promise<string[]>")]
    pub fn template_variables(&self) -> Result<AsyncTask<PendingTask<Vec<String>>>> {
        self.worker
            .submit(|document| Ok(document.model().template_variables.clone()))
    }

    #[napi(getter, ts_return_type = "Promise<string>")]
    pub fn text(&self) -> Result<AsyncTask<PendingTask<String>>> {
        self.worker.submit(|document| {
            Ok(document
                .paragraphs()
                .into_iter()
                .map(betteroffice_docx::get_paragraph_text)
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>()
                .join("\n"))
        })
    }

    #[napi(getter, ts_return_type = "Promise<DocumentStructure>")]
    pub fn structure(&self) -> Result<AsyncTask<PendingTask<DocumentStructure>>> {
        self.worker.submit(|document| {
            let value = document.structure();
            Ok(DocumentStructure {
                body_paragraphs: value.body_paragraphs as u32,
                body_tables: value.body_tables as u32,
                sections: value.sections as u32,
                headers: value.headers as u32,
                footers: value.footers as u32,
                footnotes: value.footnotes as u32,
                endnotes: value.endnotes as u32,
            })
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn body(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(|document| {
            serde_json::to_value(document.body())
                .map(JsonValue)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn headers(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(|document| {
            serde_json::to_value(document.headers())
                .map(JsonValue)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn footers(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(|document| {
            serde_json::to_value(document.footers())
                .map(JsonValue)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn sections(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(|document| {
            serde_json::to_value(document.sections())
                .map(JsonValue)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn paragraphs(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(|document| {
            serde_json::to_value(document.paragraphs())
                .map(JsonValue)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any>")]
    pub fn tables(&self) -> Result<AsyncTask<PendingTask<JsonValue>>> {
        self.worker.submit(|document| {
            serde_json::to_value(document.tables())
                .map(JsonValue)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<any | null>")]
    pub fn paragraph(
        &self,
        paragraph_id: String,
    ) -> Result<AsyncTask<PendingTask<Option<JsonValue>>>> {
        self.worker.submit(move |document| {
            document
                .paragraph(&paragraph_id)
                .map(serde_json::to_value)
                .transpose()
                .map(|value| value.map(JsonValue))
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<EditReceipt>")]
    pub fn replace_paragraph_text(
        &self,
        paragraph_id: String,
        text: String,
    ) -> Result<AsyncTask<PendingTask<EditReceipt>>> {
        let context = EditCtx {
            author: self.author.clone(),
            origin: self.origin,
            suggesting: None,
            now_iso: self.timestamp.clone(),
        };
        self.worker.submit(move |document| {
            let receipt = document
                .replace_paragraph_text_with(&paragraph_id, &text, 1, &context)
                .map_err(error)?;
            let range = receipt.range;
            Ok(EditReceipt {
                paragraph_id: range.as_ref().map(|value| value.start.para.clone()),
                story: range.as_ref().map(|value| value.start.story.clone()),
                start: range.as_ref().map(|value| value.start.offset),
                end: range.as_ref().map(|value| value.end.offset),
                new_paragraph_ids: receipt.new_para_ids,
                revision_ids: receipt.revision_ids,
            })
        })
    }

    #[napi(ts_return_type = "Promise<LayoutResult>")]
    pub fn layout(&self, input: serde_json::Value) -> Result<AsyncTask<PendingTask<LayoutResult>>> {
        let input = serde_json::from_value(input).map_err(error)?;
        self.worker.submit(move |document| {
            let result = document.layout(input).map_err(error)?;
            Ok(LayoutResult {
                layout: serde_json::to_value(result.layout).map_err(error)?,
                display_list: serde_json::to_value(result.display_list).map_err(error)?,
            })
        })
    }

    #[napi(ts_return_type = "Promise<number>")]
    pub fn register_font(&self, face: FontFace) -> Result<AsyncTask<PendingTask<u32>>> {
        let data = face.data.to_vec();
        self.worker.submit(move |document| {
            document
                .register_font(
                    &face.family,
                    face.bold.unwrap_or(false),
                    face.italic.unwrap_or(false),
                    &data,
                )
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<void>")]
    pub fn register_image(&self, image: ImageResource) -> Result<AsyncTask<PendingTask<()>>> {
        let data = image.data.to_vec();
        self.worker.submit(move |document| {
            let scope = match image.scope.as_deref().unwrap_or("body") {
                "body" => ImageScope::Body,
                "headerFooter" => ImageScope::HeaderFooter(
                    image
                        .part
                        .as_deref()
                        .ok_or_else(|| error("headerFooter images require part"))?,
                ),
                "footnotes" => ImageScope::Notes(NoteKind::Footnote),
                "endnotes" => ImageScope::Notes(NoteKind::Endnote),
                _ => {
                    return Err(error(
                        "scope must be body, headerFooter, footnotes, or endnotes",
                    ));
                }
            };
            document
                .register_image(scope, &image.relationship_id, &data)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<RenderedPage>")]
    pub fn render_page(
        &self,
        display_list: serde_json::Value,
        page: Option<u32>,
    ) -> Result<AsyncTask<PendingTask<RenderedPage>>> {
        let display_list = serde_json::from_value(display_list).map_err(error)?;
        let page = page.unwrap_or(0) as usize;
        self.worker.submit(move |document| {
            let rendered = document.render_png(&display_list, page).map_err(error)?;
            Ok(RenderedPage {
                data: rendered.bytes.into(),
                skipped_images: rendered.skipped_images as u32,
            })
        })
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn save(
        &self,
        options: Option<SaveDocumentOptions>,
    ) -> Result<AsyncTask<PendingTask<Buffer>>> {
        let options = options.unwrap_or(SaveDocumentOptions {
            timestamp: None,
            update_modified_date: None,
            modified_by: None,
        });
        let options = SaveOptions {
            now: options.timestamp.unwrap_or_else(|| self.timestamp.clone()),
            update_modified_date: options.update_modified_date.unwrap_or(false),
            modified_by: options.modified_by,
        };
        self.worker.submit(move |document| {
            document
                .save_with_options(options)
                .map(Buffer::from)
                .map_err(error)
        })
    }
}

#[napi(ts_return_type = "Promise<Document>")]
pub fn open_document(
    data: Buffer,
    options: Option<OpenDocumentOptions>,
) -> Result<AsyncTask<OpenTask>> {
    let options = options.unwrap_or(OpenDocumentOptions {
        author: None,
        origin: None,
        timestamp: None,
    });
    Ok(AsyncTask::new(OpenTask {
        bytes: data.to_vec(),
        author: options.author.unwrap_or_else(|| "node".to_owned()),
        origin: options
            .origin
            .as_deref()
            .map(origin)
            .transpose()?
            .unwrap_or(EditOrigin::Local),
        timestamp: options
            .timestamp
            .unwrap_or_else(|| SaveOptions::default().now),
    }))
}
