use std::cell::{Cell as StateCell, RefCell};
use std::collections::VecDeque;
use std::future::{Future, poll_fn};
use std::pin::Pin;
use std::rc::Rc;
use std::task::{Context, Poll, Waker};

use ooxml_opc::WorkBudget;
use xlsx_calc::graph::SnapshotGraphBuilder;

use super::*;

pub const PEER_HYDRATION_CHUNK_CELLS: usize = 512;
pub const PEER_HYDRATION_CHUNK_BYTES: usize = 64 * 1024;

async fn append_text(target: &mut String, source: &str, work: &WorkBudget) {
    let mut offset = 0;
    while offset < source.len() {
        let count = work
            .take((source.len() - offset).div_ceil(64).min(256))
            .await
            * 64;
        let mut end = (offset + count).min(source.len());
        while !source.is_char_boundary(end) {
            end -= 1;
        }
        target.push_str(&source[offset..end]);
        offset = end;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OpenerState {
    Parsing,
    NeedsHydration,
    Ready,
}

#[derive(Serialize, Deserialize)]
pub struct PeerHydrationHeader {
    sheet_count: usize,
    delta: bool,
    recalculated_since_open: bool,
    calculations_since_open: u64,
    rand_seed: Option<u32>,
    active_sheet: SheetId,
    version_nonce: String,
    committed_changes: u64,
    client_id: Option<u64>,
    proposal_id_counter: u64,
}

#[derive(Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case")]
pub enum PeerHydrationChunk {
    Header {
        header: PeerHydrationHeader,
    },
    Cells {
        sheet: usize,
        cells: Vec<PeerHydrationCell>,
    },
    CellPart {
        sheet: usize,
        cell: PeerHydrationCell,
        text: String,
        formula: String,
        last: bool,
    },
    DeletedCells {
        sheet: usize,
        cells: Vec<CellRef>,
    },
    Arrays {
        sheet: usize,
        arrays: Vec<(CellRef, CellRange)>,
    },
    Changed {
        cells: Vec<crate::CellAddress>,
    },
    CycleCells {
        cells: Vec<crate::CellAddress>,
    },
    LimitedCells {
        cells: Vec<crate::CellAddress>,
    },
    End,
}

impl PeerHydration {
    pub fn into_chunks(self) -> Vec<PeerHydrationChunk> {
        let mut chunks = vec![PeerHydrationChunk::Header {
            header: PeerHydrationHeader {
                sheet_count: self.cells.len(),
                delta: self.delta,
                recalculated_since_open: self.recalculated_since_open,
                calculations_since_open: self.calculations_since_open,
                rand_seed: self.rand_seed,
                active_sheet: self.active_sheet,
                version_nonce: self.version_nonce,
                committed_changes: self.committed_changes,
                client_id: self.client_id,
                proposal_id_counter: self.proposal_id_counter,
            },
        }];
        fn split<T: Serialize>(values: Vec<T>, mut emit: impl FnMut(Vec<T>)) {
            let mut batch = Vec::new();
            let mut bytes = 256;
            for value in values {
                let size = serde_json::to_vec(&value).expect("hydration record").len() + 1;
                if !batch.is_empty()
                    && (batch.len() == PEER_HYDRATION_CHUNK_CELLS
                        || bytes + size > PEER_HYDRATION_CHUNK_BYTES)
                {
                    emit(std::mem::take(&mut batch));
                    bytes = 256;
                }
                bytes += size;
                batch.push(value);
            }
            if !batch.is_empty() {
                emit(batch);
            }
        }
        for (sheet, cells) in self.cells.into_iter().enumerate() {
            let mut batch = Vec::new();
            let mut bytes = 256;
            for mut cell in cells {
                let size = serde_json::to_vec(&cell).expect("hydration cell").len() + 1;
                if !batch.is_empty()
                    && (batch.len() == PEER_HYDRATION_CHUNK_CELLS
                        || bytes + size > PEER_HYDRATION_CHUNK_BYTES)
                {
                    chunks.push(PeerHydrationChunk::Cells {
                        sheet,
                        cells: std::mem::take(&mut batch),
                    });
                    bytes = 256;
                }
                if size + 256 <= PEER_HYDRATION_CHUNK_BYTES {
                    bytes += size;
                    batch.push(cell);
                    continue;
                }
                let text = match &mut cell.1 {
                    CellValue::Text { value } => std::mem::take(value),
                    _ => String::new(),
                };
                let formula = cell.2.as_mut().map(std::mem::take).unwrap_or_default();
                let mut text = text.chars().peekable();
                let mut formula = formula.chars().peekable();
                loop {
                    let mut text_part = String::new();
                    let mut formula_part = String::new();
                    let mut size = 512;
                    for (source, target) in [
                        (&mut text, &mut text_part),
                        (&mut formula, &mut formula_part),
                    ] {
                        while let Some(&ch) = source.peek() {
                            let count = match ch {
                                '"' | '\\' => 2,
                                '\u{0}'..='\u{1f}' => 6,
                                _ => ch.len_utf8(),
                            };
                            if size + count > PEER_HYDRATION_CHUNK_BYTES {
                                break;
                            }
                            size += count;
                            target.push(source.next().unwrap());
                        }
                    }
                    let last = text.peek().is_none() && formula.peek().is_none();
                    chunks.push(PeerHydrationChunk::CellPart {
                        sheet,
                        cell: cell.clone(),
                        text: text_part,
                        formula: formula_part,
                        last,
                    });
                    if last {
                        break;
                    }
                }
            }
            if !batch.is_empty() {
                chunks.push(PeerHydrationChunk::Cells {
                    sheet,
                    cells: batch,
                });
            }
        }
        for (sheet, cells) in self.deleted_cells.into_iter().enumerate() {
            split(cells, |cells| {
                chunks.push(PeerHydrationChunk::DeletedCells { sheet, cells })
            });
        }
        for (sheet, arrays) in self.arrays.into_iter().enumerate() {
            split(arrays, |arrays| {
                chunks.push(PeerHydrationChunk::Arrays { sheet, arrays })
            });
        }
        split(self.last_calculation.changed, |cells| {
            chunks.push(PeerHydrationChunk::Changed { cells })
        });
        split(self.last_calculation.cycle_cells, |cells| {
            chunks.push(PeerHydrationChunk::CycleCells { cells })
        });
        split(self.last_calculation.limited_cells, |cells| {
            chunks.push(PeerHydrationChunk::LimitedCells { cells })
        });
        chunks.push(PeerHydrationChunk::End);
        chunks
    }
}

pub struct WorkbookPeerOpener {
    work: WorkBudget,
    state: Rc<StateCell<OpenerState>>,
    chunks: Rc<RefCell<VecDeque<PeerHydrationChunk>>>,
    opening: Option<Pin<Box<dyn Future<Output = Result<Workbook>>>>>,
    workbook: Option<Workbook>,
    ended: bool,
    header: bool,
    failed: bool,
}

impl WorkbookPeerOpener {
    pub fn new(bytes: Vec<u8>, client_id: Option<u64>) -> Self {
        let work = WorkBudget::default();
        let state = Rc::new(StateCell::new(OpenerState::Parsing));
        let chunks = Rc::new(RefCell::new(VecDeque::new()));
        let budget = work.clone();
        let progress = state.clone();
        let incoming = chunks.clone();
        let opening = Box::pin(async move {
            let parts = ooxml_opc::unzip_parts_sliced(&bytes, &budget)
                .await
                .map_err(Error::Package)?;
            let mut parsed =
                xlsx_parse::parse_workbook_with_owned_package_sliced(parts, &budget).await?;
            let mut workbook = Workbook::from_source_sliced(
                parsed.workbook,
                Some(parsed.package),
                parsed.active_sheet,
                false,
                client_id,
                &parsed.legacy_dimensions,
                parsed.legacy_styles.as_ref(),
                &budget,
            )
            .await?;
            for legacy in &mut parsed.legacy_dimensions {
                while !legacy.col_widths.is_empty() {
                    budget.step().await;
                    legacy.col_widths.pop_first();
                }
                while !legacy.row_heights.is_empty() {
                    budget.step().await;
                    legacy.row_heights.pop_first();
                }
            }
            if let Some(styles) = parsed.legacy_styles.take() {
                let model = WorkbookModel {
                    styles,
                    ..Default::default()
                };
                xlsx_parse::sliced::retire_workbook(model, &budget).await;
            }
            let mut initializer =
                ooxml_opc::SourceContainerInitializer::new(bytes.len()).map_err(Error::Package)?;
            while !initializer.is_ready() {
                let count = budget.take(256).await * 64;
                initializer.advance(count).map_err(Error::Package)?;
            }
            let mut source = initializer.finish().map_err(Error::Package)?;
            let mut offset = 0;
            while offset < bytes.len() {
                let count = (budget.take(256).await * 64).min(bytes.len() - offset);
                source
                    .push(&bytes[offset..offset + count])
                    .map_err(Error::Package)?;
                offset += count;
            }
            workbook.source_container = Some(source.finish().map_err(Error::Package)?);
            drop(bytes);
            progress.set(OpenerState::NeedsHydration);
            let mut header = None;
            let mut partial: Option<(usize, PeerHydrationCell)> = None;
            loop {
                let chunk = poll_fn(|_| match incoming.borrow_mut().pop_front() {
                    Some(chunk) => {
                        progress.set(OpenerState::Parsing);
                        Poll::Ready(chunk)
                    }
                    None => {
                        progress.set(OpenerState::NeedsHydration);
                        Poll::Pending
                    }
                })
                .await;
                budget.step().await;
                if partial.is_some() && !matches!(&chunk, PeerHydrationChunk::CellPart { .. }) {
                    return Err(Error::InvalidRequest(
                        "Incomplete peer hydration cell".into(),
                    ));
                }
                match chunk {
                    PeerHydrationChunk::Header { header: next } => {
                        if next.sheet_count != workbook.model.sheets.len()
                            || next.client_id != client_id
                        {
                            return Err(Error::InvalidRequest(
                                "Peer hydration sheet count or client differs".into(),
                            ));
                        }
                        for sheet in &mut workbook.model.sheets {
                            if !next.delta {
                                loop {
                                    let at = sheet.iter_cells().next().map(|(at, _)| at);
                                    let Some(at) = at else {
                                        break;
                                    };
                                    budget.step().await;
                                    sheet.set_cell(at, xlsx_model::Cell::default());
                                }
                            }
                            loop {
                                let at = sheet.array_formulas().next().map(|(at, _)| at);
                                let Some(at) = at else {
                                    break;
                                };
                                budget.step().await;
                                sheet.clear_array_formula(at);
                            }
                        }
                        header = Some(next);
                    }
                    PeerHydrationChunk::Cells { sheet, cells } => {
                        let sheet = workbook
                            .model
                            .sheets
                            .get_mut(sheet)
                            .ok_or(Error::SheetOutOfRange(SheetId(sheet as u32)))?;
                        for (at, value, formula, style) in cells {
                            budget.step().await;
                            sheet.set_cell(
                                at,
                                xlsx_model::Cell {
                                    value,
                                    formula,
                                    style,
                                },
                            );
                        }
                    }
                    PeerHydrationChunk::CellPart {
                        sheet,
                        cell,
                        text,
                        formula,
                        last,
                    } => {
                        let (owner, current) = partial.get_or_insert((sheet, cell.clone()));
                        if *owner != sheet || current.0 != cell.0 {
                            return Err(Error::InvalidRequest(
                                "Peer hydration cell parts differ".into(),
                            ));
                        }
                        if let CellValue::Text { value } = &mut current.1 {
                            append_text(value, &text, &budget).await;
                        } else if !text.is_empty() {
                            return Err(Error::InvalidRequest(
                                "Unexpected peer hydration text".into(),
                            ));
                        }
                        if let Some(value) = &mut current.2 {
                            append_text(value, &formula, &budget).await;
                        } else if !formula.is_empty() {
                            return Err(Error::InvalidRequest(
                                "Unexpected peer hydration formula".into(),
                            ));
                        }
                        if last {
                            let (_, (at, value, formula, style)) = partial.take().unwrap();
                            let target = workbook
                                .model
                                .sheets
                                .get_mut(sheet)
                                .ok_or(Error::SheetOutOfRange(SheetId(sheet as u32)))?;
                            target.set_cell(
                                at,
                                xlsx_model::Cell {
                                    value,
                                    formula,
                                    style,
                                },
                            );
                        }
                    }
                    PeerHydrationChunk::DeletedCells { sheet, cells } => {
                        if !header
                            .as_ref()
                            .is_some_and(|header: &PeerHydrationHeader| header.delta)
                        {
                            continue;
                        }
                        let sheet = workbook
                            .model
                            .sheets
                            .get_mut(sheet)
                            .ok_or(Error::SheetOutOfRange(SheetId(sheet as u32)))?;
                        for at in cells {
                            budget.step().await;
                            sheet.set_cell(at, xlsx_model::Cell::default());
                        }
                    }
                    PeerHydrationChunk::Arrays { sheet, arrays } => {
                        let sheet = workbook
                            .model
                            .sheets
                            .get_mut(sheet)
                            .ok_or(Error::SheetOutOfRange(SheetId(sheet as u32)))?;
                        for (at, range) in arrays {
                            budget.step().await;
                            sheet.set_array_formula(at, range);
                        }
                    }
                    PeerHydrationChunk::Changed { cells } => {
                        for cell in cells {
                            budget.step().await;
                            workbook.last_calculation.changed.push(cell);
                        }
                    }
                    PeerHydrationChunk::CycleCells { cells } => {
                        for cell in cells {
                            budget.step().await;
                            workbook.last_calculation.cycle_cells.push(cell);
                        }
                    }
                    PeerHydrationChunk::LimitedCells { cells } => {
                        for cell in cells {
                            budget.step().await;
                            workbook.last_calculation.limited_cells.push(cell);
                        }
                    }
                    PeerHydrationChunk::End => break,
                }
            }
            let header = header
                .ok_or_else(|| Error::InvalidRequest("Peer hydration header is missing".into()))?;
            validate_model_sheets_sliced(&workbook.model, &budget).await?;
            workbook.recalculated_since_open = header.recalculated_since_open;
            workbook.calculations_since_open = header.calculations_since_open;
            workbook.rand_seed = header.rand_seed;
            workbook.set_active_sheet(header.active_sheet)?;
            workbook.version_nonce = header.version_nonce;
            workbook.committed_changes = header.committed_changes;
            workbook.proposals = ProposalSet::with_id_counter(header.proposal_id_counter);
            progress.set(OpenerState::Parsing);
            let mut graph = SnapshotGraphBuilder::new();
            loop {
                budget.step().await;
                if graph
                    .advance(&workbook.model, usize::MAX)
                    .map_err(Error::InvalidOperation)?
                    .0
                {
                    break;
                }
            }
            workbook.graph = Some(graph.finish().ok_or_else(|| {
                Error::InvalidOperation("Peer dependency graph is incomplete".into())
            })?);
            Ok(workbook)
        });
        Self {
            work,
            state,
            chunks,
            opening: Some(opening),
            workbook: None,
            ended: false,
            header: false,
            failed: false,
        }
    }

    pub fn advance(&mut self, units: usize) -> Result<OpenerState> {
        if self.failed {
            return Err(Error::InvalidOperation(
                "Workbook peer opener failed".into(),
            ));
        }
        self.work.reset(units);
        if let Some(opening) = &mut self.opening {
            let mut context = Context::from_waker(Waker::noop());
            if let Poll::Ready(result) = opening.as_mut().poll(&mut context) {
                self.opening = None;
                match result {
                    Ok(workbook) => {
                        self.workbook = Some(workbook);
                        self.state.set(OpenerState::Ready);
                    }
                    Err(error) => {
                        self.failed = true;
                        return Err(error);
                    }
                }
            }
        }
        Ok(self.state.get())
    }

    pub fn push_hydration(&mut self, chunk: PeerHydrationChunk) -> Result<()> {
        if self.ended || self.failed {
            return Err(Error::InvalidRequest("Peer hydration is closed".into()));
        }
        let count = match &chunk {
            PeerHydrationChunk::Header { .. } => {
                if self.header {
                    return Err(Error::InvalidRequest(
                        "Duplicate peer hydration header".into(),
                    ));
                }
                self.header = true;
                0
            }
            PeerHydrationChunk::Cells { cells, .. } => cells.len(),
            PeerHydrationChunk::CellPart { .. } => 1,
            PeerHydrationChunk::DeletedCells { cells, .. } => cells.len(),
            PeerHydrationChunk::Arrays { arrays, .. } => arrays.len(),
            PeerHydrationChunk::Changed { cells }
            | PeerHydrationChunk::CycleCells { cells }
            | PeerHydrationChunk::LimitedCells { cells } => cells.len(),
            PeerHydrationChunk::End => {
                self.ended = true;
                0
            }
        };
        if !self.header || count > PEER_HYDRATION_CHUNK_CELLS {
            return Err(Error::InvalidRequest("Invalid peer hydration chunk".into()));
        }
        self.chunks.borrow_mut().push_back(chunk);
        Ok(())
    }

    pub fn finish(mut self) -> Result<Workbook> {
        self.workbook
            .take()
            .ok_or_else(|| Error::InvalidOperation("Workbook peer opener is not ready".into()))
    }

    #[doc(hidden)]
    pub fn touched_units(&self) -> usize {
        self.work.touched()
    }
}
