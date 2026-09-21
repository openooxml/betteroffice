#![deny(clippy::all)]

use std::sync::mpsc::{self, Receiver, Sender};

use betteroffice_xlsx::{
    CalculationOptions, CellEdit, CellRange, CellRef, CellValue as CoreCellValue,
    HorizontalAlignment, MAX_COLLABORATION_CLIENT_ID, MutationResult as CoreMutationResult,
    NumberFormatMutation, Proposal as CoreProposal, ProposalEditInput as CoreProposalEditInput,
    ProposalRequest, RenderOptions, SheetId, StylePatch, TextWrapping, VerticalAlignment, Workbook,
};
use napi::bindgen_prelude::{AsyncTask, Buffer, Task, ToNapiValue, TypeName};
use napi::{Env, Error, Result};
use napi_derive::napi;

fn error(reason: impl ToString) -> Error {
    Error::from_reason(reason.to_string())
}

fn cell_ref(address: &str) -> Result<CellRef> {
    CellRef::parse_a1(&address.to_ascii_uppercase()).map_err(error)
}

fn cell_range(range: &str) -> Result<CellRange> {
    CellRange::parse_a1(&range.to_ascii_uppercase()).map_err(error)
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

#[napi(object)]
pub struct OpenWorkbookOptions {
    pub client_id: Option<f64>,
    pub read_only: Option<bool>,
    pub recalculate: Option<bool>,
    pub now_serial: Option<f64>,
}

#[napi(object)]
pub struct RenderSheetOptions {
    pub sheet: Option<u32>,
    pub range: Option<String>,
    pub scale: Option<f64>,
    pub max_width: Option<u32>,
    pub max_height: Option<u32>,
}

#[napi(object)]
pub struct RenderedSheet {
    pub data: Buffer,
    pub width: u32,
    pub height: u32,
}

#[napi(object)]
pub struct CellValue {
    pub address: String,
    pub input: String,
    pub formula: bool,
}

#[napi(object)]
pub struct CellInput {
    pub address: String,
    pub input: String,
}

#[napi(object)]
pub struct ProposalInput {
    pub agent_id: String,
    pub note: Option<String>,
    pub edits: Vec<ProposalCellInput>,
}

#[napi(object)]
pub struct ProposalCellInput {
    pub sheet: u32,
    pub address: String,
    pub input: String,
}

#[napi(object)]
pub struct Proposal {
    pub id: String,
    pub agent_id: String,
    pub note: Option<String>,
    pub edits: Vec<ProposedCell>,
}

#[napi(object)]
pub struct ProposedCell {
    pub sheet: u32,
    pub address: String,
    pub input: String,
    pub before: String,
    pub after: String,
}

#[napi(object)]
pub struct CellComputedValue {
    pub kind: String,
    pub number: Option<f64>,
    pub text: Option<String>,
    pub boolean: Option<bool>,
    pub error: Option<String>,
}

#[napi(object)]
pub struct StyleInput {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strikethrough: Option<bool>,
    pub font_family: Option<String>,
    pub font_size: Option<f64>,
    pub text_color: Option<String>,
    pub fill_color: Option<String>,
    pub horizontal_alignment: Option<String>,
    pub vertical_alignment: Option<String>,
    pub text_wrapping: Option<String>,
}

#[napi(object)]
pub struct MutationResult {
    pub applied: bool,
    pub changed: Vec<String>,
    pub cycle_cells: Vec<String>,
    pub limited_cells: Vec<String>,
}

#[napi(object)]
pub struct CalculationResult {
    pub changed: Vec<String>,
    pub cycle_cells: Vec<String>,
    pub limited_cells: Vec<String>,
}

#[napi(object)]
pub struct HistoryState {
    pub can_undo: bool,
    pub can_redo: bool,
    pub undo_depth: u32,
    pub redo_depth: u32,
}

#[napi(object)]
pub struct SheetInfo {
    pub ids: Vec<String>,
    pub names: Vec<String>,
    pub active_sheet: u32,
    pub content_width: f64,
    pub content_height: f64,
    pub frozen_rows: u32,
    pub frozen_columns: u32,
    pub initial_scroll_x: f64,
    pub initial_scroll_y: f64,
}

fn map_cell(address: String, value: CellEdit) -> CellValue {
    CellValue {
        address,
        input: value.input,
        formula: value.is_formula,
    }
}

fn addresses(workbook: &Workbook, values: Vec<betteroffice_xlsx::CellAddress>) -> Vec<String> {
    values
        .into_iter()
        .map(|address| workbook.format_address(address))
        .collect()
}

fn map_mutation(workbook: &Workbook, value: CoreMutationResult) -> MutationResult {
    MutationResult {
        applied: value.applied,
        changed: addresses(workbook, value.changed),
        cycle_cells: addresses(workbook, value.cycle_cells),
        limited_cells: addresses(workbook, value.limited_cells),
    }
}

fn map_computed_value(value: &CoreCellValue) -> CellComputedValue {
    match value {
        CoreCellValue::Empty => CellComputedValue {
            kind: "empty".to_owned(),
            number: None,
            text: None,
            boolean: None,
            error: None,
        },
        CoreCellValue::Number { value } => CellComputedValue {
            kind: "number".to_owned(),
            number: Some(*value),
            text: None,
            boolean: None,
            error: None,
        },
        CoreCellValue::Text { value } => CellComputedValue {
            kind: "text".to_owned(),
            number: None,
            text: Some(value.clone()),
            boolean: None,
            error: None,
        },
        CoreCellValue::Bool { value } => CellComputedValue {
            kind: "boolean".to_owned(),
            number: None,
            text: None,
            boolean: Some(*value),
            error: None,
        },
        CoreCellValue::Error { value } => CellComputedValue {
            kind: "error".to_owned(),
            number: None,
            text: None,
            boolean: None,
            error: Some(value.as_str().to_owned()),
        },
    }
}

fn map_proposal(value: &CoreProposal) -> Proposal {
    Proposal {
        id: value.id.clone(),
        agent_id: value.agent_id.clone(),
        note: value.note.clone(),
        edits: value
            .edits
            .iter()
            .map(|edit| ProposedCell {
                sheet: edit.sheet,
                address: CellRef::new(edit.row, edit.col).to_a1(),
                input: edit.input.clone(),
                before: edit.old_text.clone(),
                after: edit.new_text.clone(),
            })
            .collect(),
    }
}

fn horizontal_alignment(value: &str) -> Result<HorizontalAlignment> {
    match value {
        "left" => Ok(HorizontalAlignment::Left),
        "center" => Ok(HorizontalAlignment::Center),
        "right" => Ok(HorizontalAlignment::Right),
        _ => Err(error("horizontalAlignment must be left, center, or right")),
    }
}

fn vertical_alignment(value: &str) -> Result<VerticalAlignment> {
    match value {
        "top" => Ok(VerticalAlignment::Top),
        "middle" => Ok(VerticalAlignment::Middle),
        "bottom" => Ok(VerticalAlignment::Bottom),
        _ => Err(error("verticalAlignment must be top, middle, or bottom")),
    }
}

fn text_wrapping(value: &str) -> Result<TextWrapping> {
    match value {
        "overflow" => Ok(TextWrapping::Overflow),
        "wrap" => Ok(TextWrapping::Wrap),
        "clip" => Ok(TextWrapping::Clip),
        _ => Err(error("textWrapping must be overflow, wrap, or clip")),
    }
}

type Job = Box<dyn FnOnce(&mut Workbook) + Send>;

#[derive(Clone)]
pub struct Worker {
    sender: Sender<Job>,
}

impl Worker {
    fn start(mut workbook: Workbook) -> Result<Self> {
        let (sender, receive) = mpsc::channel::<Job>();
        std::thread::Builder::new()
            .name("betteroffice-xlsx".to_owned())
            .spawn(move || {
                for job in receive {
                    job(&mut workbook);
                }
            })
            .map_err(error)?;
        Ok(Self { sender })
    }

    fn submit<T, F>(&self, operation: F) -> Result<AsyncTask<PendingTask<T>>>
    where
        T: ToNapiValue + TypeName + Send + 'static,
        F: FnOnce(&mut Workbook) -> Result<T> + Send + 'static,
    {
        let (reply, receive) = mpsc::channel();
        self.sender
            .send(Box::new(move |workbook| {
                let _ = reply.send(operation(workbook));
            }))
            .map_err(|_| error("workbook worker stopped"))?;
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
            .map_err(|_| error("workbook worker stopped"))?
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }
}

pub struct OpenTask {
    bytes: Vec<u8>,
    options: OpenWorkbookOptions,
}

impl Task for OpenTask {
    type Output = Worker;
    type JsValue = XlsxWorkbook;

    fn compute(&mut self) -> Result<Self::Output> {
        let calculation = CalculationOptions {
            now_serial: self.options.now_serial,
        };
        let client_id = self.options.client_id.map(client_id).transpose()?;
        let workbook = match (
            client_id,
            self.options.read_only.unwrap_or(false),
            self.options.recalculate.unwrap_or(false),
        ) {
            (Some(client_id), _, true) => {
                Workbook::open_collaborative_recalculated(&self.bytes, client_id, calculation)
            }
            (Some(client_id), _, false) => Workbook::open_collaborative(&self.bytes, client_id),
            (None, true, _) => Workbook::open_for_read(&self.bytes),
            (None, false, true) => Workbook::open_recalculated(&self.bytes, calculation),
            (None, false, false) => Workbook::open(&self.bytes),
        }
        .map_err(error)?;
        Worker::start(workbook)
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(XlsxWorkbook { worker: output })
    }
}

#[napi(js_name = "Workbook")]
pub struct XlsxWorkbook {
    worker: Worker,
}

#[napi]
impl XlsxWorkbook {
    #[napi(getter, ts_return_type = "Promise<number>")]
    pub fn client_id(&self) -> Result<AsyncTask<PendingTask<f64>>> {
        self.worker
            .submit(|workbook| Ok(workbook.client_id() as f64))
    }

    #[napi(getter, ts_return_type = "Promise<boolean>")]
    pub fn collaborative(&self) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker
            .submit(|workbook| Ok(workbook.is_collaborative()))
    }

    #[napi(getter, ts_return_type = "Promise<number>")]
    pub fn sheet_count(&self) -> Result<AsyncTask<PendingTask<u32>>> {
        self.worker
            .submit(|workbook| Ok(workbook.sheet_count() as u32))
    }

    #[napi(getter, ts_return_type = "Promise<number>")]
    pub fn active_sheet(&self) -> Result<AsyncTask<PendingTask<u32>>> {
        self.worker.submit(|workbook| Ok(workbook.active_sheet().0))
    }

    #[napi(ts_return_type = "Promise<void>")]
    pub fn set_active_sheet(&self, sheet: u32) -> Result<AsyncTask<PendingTask<()>>> {
        self.worker
            .submit(move |workbook| workbook.set_active_sheet(SheetId(sheet)).map_err(error))
    }

    #[napi(ts_return_type = "Promise<SheetInfo>")]
    pub fn sheet_info(&self) -> Result<AsyncTask<PendingTask<SheetInfo>>> {
        self.worker.submit(|workbook| {
            let value = workbook.sheet_info().map_err(error)?;
            Ok(SheetInfo {
                ids: value.sheet_ids,
                names: value.sheet_names,
                active_sheet: value.active_sheet.0,
                content_width: f64::from(value.content_width),
                content_height: f64::from(value.content_height),
                frozen_rows: value.frozen_rows,
                frozen_columns: value.frozen_cols,
                initial_scroll_x: f64::from(value.initial_scroll_x),
                initial_scroll_y: f64::from(value.initial_scroll_y),
            })
        })
    }

    #[napi(ts_return_type = "Promise<CellValue>")]
    pub fn cell(&self, sheet: u32, address: String) -> Result<AsyncTask<PendingTask<CellValue>>> {
        let cell = cell_ref(&address)?;
        self.worker.submit(move |workbook| {
            let value = workbook.cell(SheetId(sheet), cell).map_err(error)?;
            Ok(map_cell(address, value))
        })
    }

    #[napi(ts_return_type = "Promise<CellComputedValue>")]
    pub fn value(
        &self,
        sheet: u32,
        address: String,
    ) -> Result<AsyncTask<PendingTask<CellComputedValue>>> {
        let cell = cell_ref(&address)?;
        self.worker.submit(move |workbook| {
            let sheet = workbook.sheet(SheetId(sheet)).map_err(error)?;
            Ok(sheet
                .cell(cell)
                .map(|cell| map_computed_value(&cell.value))
                .unwrap_or_else(|| map_computed_value(&CoreCellValue::Empty)))
        })
    }

    #[napi(ts_return_type = "Promise<string | null>")]
    pub fn formula(
        &self,
        sheet: u32,
        address: String,
    ) -> Result<AsyncTask<PendingTask<Option<String>>>> {
        let cell = cell_ref(&address)?;
        self.worker.submit(move |workbook| {
            let sheet = workbook.sheet(SheetId(sheet)).map_err(error)?;
            Ok(sheet.cell(cell).and_then(|cell| cell.formula.clone()))
        })
    }

    #[napi(ts_return_type = "Promise<CellValue[][]>")]
    pub fn range(
        &self,
        sheet: u32,
        range: String,
    ) -> Result<AsyncTask<PendingTask<Vec<Vec<CellValue>>>>> {
        let parsed = cell_range(&range)?;
        self.worker.submit(move |workbook| {
            let cells = workbook
                .range_cells(SheetId(sheet), parsed)
                .map_err(error)?;
            Ok(cells
                .into_iter()
                .enumerate()
                .map(|(row, values)| {
                    values
                        .into_iter()
                        .enumerate()
                        .map(|(column, value)| {
                            let address = CellRef::new(
                                parsed.start.row + row as u32,
                                parsed.start.col + column as u32,
                            )
                            .to_a1();
                            map_cell(address, value)
                        })
                        .collect()
                })
                .collect())
        })
    }

    #[napi(ts_return_type = "Promise<MutationResult>")]
    pub fn set(
        &self,
        sheet: u32,
        address: String,
        input: String,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<MutationResult>>> {
        let cell = cell_ref(&address)?;
        self.worker.submit(move |workbook| {
            let result = workbook
                .edit_cell(
                    SheetId(sheet),
                    cell,
                    &input,
                    CalculationOptions { now_serial },
                )
                .map_err(error)?;
            Ok(map_mutation(workbook, result))
        })
    }

    #[napi(ts_return_type = "Promise<MutationResult>")]
    pub fn set_many(
        &self,
        sheet: u32,
        edits: Vec<CellInput>,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<MutationResult>>> {
        let edits = edits
            .into_iter()
            .map(|edit| {
                Ok(betteroffice_xlsx::CellInput {
                    cell: cell_ref(&edit.address)?,
                    input: edit.input,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        self.worker.submit(move |workbook| {
            let result = workbook
                .edit_cells(SheetId(sheet), &edits, CalculationOptions { now_serial })
                .map_err(error)?;
            Ok(map_mutation(workbook, result))
        })
    }

    #[napi(ts_return_type = "Promise<CalculationResult>")]
    pub fn recalculate(
        &self,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<CalculationResult>>> {
        self.worker.submit(move |workbook| {
            let value = workbook.recalculate_all(CalculationOptions { now_serial });
            Ok(CalculationResult {
                changed: addresses(workbook, value.changed),
                cycle_cells: addresses(workbook, value.cycle_cells),
                limited_cells: addresses(workbook, value.limited_cells),
            })
        })
    }

    #[napi(getter, ts_return_type = "Promise<HistoryState>")]
    pub fn history(&self) -> Result<AsyncTask<PendingTask<HistoryState>>> {
        self.worker.submit(|workbook| {
            let value = workbook.history_state();
            Ok(HistoryState {
                can_undo: value.can_undo,
                can_redo: value.can_redo,
                undo_depth: value.undo_depth as u32,
                redo_depth: value.redo_depth as u32,
            })
        })
    }

    #[napi(getter, ts_return_type = "Promise<boolean>")]
    pub fn can_undo(&self) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker.submit(|workbook| Ok(workbook.can_undo()))
    }

    #[napi(getter, ts_return_type = "Promise<boolean>")]
    pub fn can_redo(&self) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker.submit(|workbook| Ok(workbook.can_redo()))
    }

    #[napi(ts_return_type = "Promise<MutationResult>")]
    pub fn undo(&self, now_serial: Option<f64>) -> Result<AsyncTask<PendingTask<MutationResult>>> {
        self.worker.submit(move |workbook| {
            let result = workbook
                .undo(CalculationOptions { now_serial })
                .map_err(error)?;
            Ok(map_mutation(workbook, result))
        })
    }

    #[napi(ts_return_type = "Promise<MutationResult>")]
    pub fn redo(&self, now_serial: Option<f64>) -> Result<AsyncTask<PendingTask<MutationResult>>> {
        self.worker.submit(move |workbook| {
            let result = workbook
                .redo(CalculationOptions { now_serial })
                .map_err(error)?;
            Ok(map_mutation(workbook, result))
        })
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn encode_state_vector(&self) -> Result<AsyncTask<PendingTask<Buffer>>> {
        self.worker
            .submit(|workbook| Ok(workbook.encode_state_vector_v1().into()))
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn encode_state_as_update(&self) -> Result<AsyncTask<PendingTask<Buffer>>> {
        self.worker
            .submit(|workbook| Ok(workbook.encode_state_as_update_v1().into()))
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn encode_diff(&self, state_vector: Buffer) -> Result<AsyncTask<PendingTask<Buffer>>> {
        let state_vector = state_vector.to_vec();
        self.worker.submit(move |workbook| {
            workbook
                .encode_diff_v1(&state_vector)
                .map(Buffer::from)
                .map_err(error)
        })
    }

    #[napi(ts_return_type = "Promise<CalculationResult>")]
    pub fn apply_update(
        &self,
        update: Buffer,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<CalculationResult>>> {
        let update = update.to_vec();
        self.worker.submit(move |workbook| {
            let value = workbook
                .apply_update_v1(&update, CalculationOptions { now_serial })
                .map_err(error)?;
            Ok(CalculationResult {
                changed: addresses(workbook, value.changed),
                cycle_cells: addresses(workbook, value.cycle_cells),
                limited_cells: addresses(workbook, value.limited_cells),
            })
        })
    }

    #[napi(ts_return_type = "Promise<Proposal>")]
    pub fn propose(
        &self,
        proposal: ProposalInput,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<Proposal>>> {
        let edits = proposal
            .edits
            .into_iter()
            .map(|edit| {
                Ok(CoreProposalEditInput {
                    sheet: SheetId(edit.sheet),
                    cell: cell_ref(&edit.address)?,
                    input: edit.input,
                    number_format: None,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        self.worker.submit(move |workbook| {
            let value = workbook
                .propose(
                    ProposalRequest {
                        agent_id: proposal.agent_id,
                        note: proposal.note,
                        edits,
                    },
                    CalculationOptions { now_serial },
                )
                .map_err(error)?;
            Ok(map_proposal(&value))
        })
    }

    #[napi(getter, ts_return_type = "Promise<Proposal[]>")]
    pub fn proposals(&self) -> Result<AsyncTask<PendingTask<Vec<Proposal>>>> {
        self.worker
            .submit(|workbook| Ok(workbook.proposals().iter().map(map_proposal).collect()))
    }

    #[napi(ts_return_type = "Promise<MutationResult>")]
    pub fn accept_proposal(
        &self,
        proposal_id: String,
        force: Option<bool>,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<MutationResult>>> {
        self.worker.submit(move |workbook| {
            let value = workbook
                .accept_proposal(
                    &proposal_id,
                    force.unwrap_or(false),
                    CalculationOptions { now_serial },
                )
                .map_err(error)?;
            Ok(map_mutation(workbook, value.mutation))
        })
    }

    #[napi(ts_return_type = "Promise<boolean>")]
    pub fn reject_proposal(&self, proposal_id: String) -> Result<AsyncTask<PendingTask<bool>>> {
        self.worker
            .submit(move |workbook| Ok(workbook.reject_proposal(&proposal_id)))
    }

    #[napi(ts_return_type = "Promise<string[]>")]
    pub fn merged_ranges(
        &self,
        sheet: u32,
        range: String,
    ) -> Result<AsyncTask<PendingTask<Vec<String>>>> {
        let range = cell_range(&range)?;
        self.worker.submit(move |workbook| {
            Ok(workbook
                .merged_ranges(SheetId(sheet), range)
                .map_err(error)?
                .into_iter()
                .map(|value| value.to_a1())
                .collect())
        })
    }

    #[napi(getter, ts_return_type = "Promise<CalculationResult>")]
    pub fn last_calculation(&self) -> Result<AsyncTask<PendingTask<CalculationResult>>> {
        self.worker.submit(|workbook| {
            let value = workbook.last_calculation();
            Ok(CalculationResult {
                changed: addresses(workbook, value.changed.clone()),
                cycle_cells: addresses(workbook, value.cycle_cells.clone()),
                limited_cells: addresses(workbook, value.limited_cells.clone()),
            })
        })
    }

    #[napi(ts_return_type = "Promise<MutationResult>")]
    pub fn set_number_format(
        &self,
        sheet: u32,
        range: String,
        format: String,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<MutationResult>>> {
        let range = cell_range(&range)?;
        let format = match format.to_ascii_lowercase().as_str() {
            "automatic" => NumberFormatMutation::Automatic,
            "text" => NumberFormatMutation::PlainText,
            "number" => NumberFormatMutation::Number,
            "percent" => NumberFormatMutation::Percent,
            "scientific" => NumberFormatMutation::Scientific,
            "currency" => NumberFormatMutation::Currency,
            "date" => NumberFormatMutation::Date,
            "time" => NumberFormatMutation::Time,
            _ => NumberFormatMutation::Custom { pattern: format },
        };
        self.worker.submit(move |workbook| {
            let value = workbook
                .set_range_number_format(
                    SheetId(sheet),
                    range,
                    format,
                    CalculationOptions { now_serial },
                )
                .map_err(error)?;
            Ok(map_mutation(workbook, value))
        })
    }

    #[napi(ts_return_type = "Promise<MutationResult>")]
    pub fn set_style(
        &self,
        sheet: u32,
        range: String,
        style: StyleInput,
        now_serial: Option<f64>,
    ) -> Result<AsyncTask<PendingTask<MutationResult>>> {
        let range = cell_range(&range)?;
        let patch = StylePatch {
            bold: style.bold,
            italic: style.italic,
            strikethrough: style.strikethrough,
            font_family: style.font_family,
            font_size: style.font_size,
            text_color: style.text_color,
            fill_color: style.fill_color,
            border: None,
            horizontal_alignment: style
                .horizontal_alignment
                .as_deref()
                .map(horizontal_alignment)
                .transpose()?,
            vertical_alignment: style
                .vertical_alignment
                .as_deref()
                .map(vertical_alignment)
                .transpose()?,
            text_wrapping: style
                .text_wrapping
                .as_deref()
                .map(text_wrapping)
                .transpose()?,
            clear: Vec::new(),
        };
        self.worker.submit(move |workbook| {
            let value = workbook
                .patch_range_style(
                    SheetId(sheet),
                    range,
                    patch,
                    CalculationOptions { now_serial },
                )
                .map_err(error)?;
            Ok(map_mutation(workbook, value))
        })
    }

    #[napi(ts_return_type = "Promise<RenderedSheet>")]
    pub fn render_sheet(
        &self,
        options: Option<RenderSheetOptions>,
    ) -> Result<AsyncTask<PendingTask<RenderedSheet>>> {
        let options = options.unwrap_or(RenderSheetOptions {
            sheet: None,
            range: None,
            scale: None,
            max_width: None,
            max_height: None,
        });
        let range = options.range.as_deref().map(cell_range).transpose()?;
        self.worker.submit(move |workbook| {
            let sheet = SheetId(options.sheet.unwrap_or(workbook.active_sheet().0));
            let rendered = workbook
                .render_sheet(
                    sheet,
                    &RenderOptions {
                        range,
                        scale: options.scale.unwrap_or(1.0) as f32,
                        max_width: options.max_width,
                        max_height: options.max_height,
                    },
                )
                .map_err(error)?;
            Ok(RenderedSheet {
                data: rendered.bytes.into(),
                width: rendered.width,
                height: rendered.height,
            })
        })
    }

    #[napi(ts_return_type = "Promise<Buffer>")]
    pub fn save(&self) -> Result<AsyncTask<PendingTask<Buffer>>> {
        self.worker
            .submit(|workbook| workbook.save().map(Buffer::from).map_err(error))
    }
}

#[napi(ts_return_type = "Promise<Workbook>")]
pub fn open_workbook(data: Buffer, options: Option<OpenWorkbookOptions>) -> AsyncTask<OpenTask> {
    let options = options.unwrap_or(OpenWorkbookOptions {
        client_id: None,
        read_only: None,
        recalculate: None,
        now_serial: None,
    });
    AsyncTask::new(OpenTask {
        bytes: data.to_vec(),
        options,
    })
}
