//! pure core behind the wasm boundary: `&str`/`&[u8]` in, `Result<_, String>` out.
//! native tests call these fns directly, never the wasm wrappers (JsValue aborts).

use std::collections::HashSet;

use serde::{Deserialize, Serialize};

use xlsx_calc::graph::DepGraph;
use xlsx_calc::{RecalcResult, rebuild_and_recalc_all, recalc_after};
use xlsx_model::{CellRange, CellRef, CellValue, SheetId, Workbook};
use xlsx_ops::{
    CellState, Op, Proposal, ProposalSet, ProposedEdit, Provenance, Transaction, UndoStack,
    cell_state_for_input_no_eval,
};
use xlsx_render::{GridGeometry, Viewport, build_display_list, display_text};

/// cap on the cells one `range_cells_json` call may materialize.
const MAX_RANGE_CELLS: u64 = 100_000;

/// an open workbook, the active sheet, edit history, dependency graph, and
/// pending proposals; one per `XlsxDocument` handed to js.
pub struct Session {
    workbook: Workbook,
    sheet: SheetId,
    undo: UndoStack,
    graph: DepGraph,
    proposals: ProposalSet,
}

/// chrome data: sheet tabs plus the active sheet's scrollable extent.
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct SheetInfo {
    sheet_names: Vec<String>,
    active_sheet: u32,
    content_width: f32,
    content_height: f32,
}

#[derive(Deserialize)]
struct EditArgs {
    sheet: u32,
    row: u32,
    col: u32,
    input: String,
}

#[derive(Deserialize)]
struct EditBatchArgs {
    sheet: u32,
    edits: Vec<CellEditInput>,
}

#[derive(Deserialize)]
struct CellEditInput {
    row: u32,
    col: u32,
    input: String,
}

#[derive(Deserialize)]
struct OpsArgs {
    ops: Vec<Op>,
}

#[derive(Deserialize)]
struct CellArgs {
    sheet: u32,
    row: u32,
    col: u32,
}

#[derive(Deserialize)]
struct RangeArgs {
    sheet: u32,
    range: String,
}

/// mutation result: applied flag, refreshed extent, and the a1 addresses of
/// other cells changed via recalc (`Sheet!A1`-prefixed off the active sheet).
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct EditResult {
    applied: bool,
    sheet_info: SheetInfo,
    changed: Vec<String>,
}

/// editable view of one cell: a1 address, exact editor string, formula flag.
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct CellEdit {
    a1: String,
    input: String,
    is_formula: bool,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct RangeCells {
    cells: Vec<Vec<CellEdit>>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct ProposeArgs {
    agent_id: String,
    #[serde(default)]
    note: Option<String>,
    edits: Vec<ProposeEditInput>,
}

#[derive(Deserialize)]
struct ProposeEditInput {
    sheet: u32,
    row: u32,
    col: u32,
    input: String,
}

#[derive(Deserialize)]
struct AcceptArgs {
    id: String,
    #[serde(default)]
    force: bool,
}

#[derive(Deserialize)]
struct IdArgs {
    id: String,
}

/// wire form of `list_proposals_json`.
#[derive(Serialize)]
struct ProposalList<'a> {
    proposals: &'a [Proposal],
}

/// `reject_proposal_json`'s reply.
#[derive(Serialize)]
struct RejectResult {
    removed: bool,
}

/// `accept_proposal_json`'s reply: the edit envelope plus the proposal id.
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct AcceptResult {
    applied: bool,
    sheet_info: SheetInfo,
    changed: Vec<String>,
    proposal_id: String,
}

/// editor string for a stored (non-formula) value; text that would re-parse
/// as something else gets a leading `'` so a round-trip preserves it.
fn value_to_input(v: &CellValue) -> String {
    match v {
        CellValue::Empty => String::new(),
        CellValue::Number { value } => value.to_string(),
        CellValue::Bool { value } => if *value { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::Text { value } => {
            if text_needs_quote(value) {
                format!("'{value}")
            } else {
                value.clone()
            }
        }
        CellValue::Error { value } => value.as_str().to_string(),
    }
}

/// whether re-entering `text` verbatim would fail to yield the same text.
fn text_needs_quote(text: &str) -> bool {
    !matches!(xlsx_ops::parse_input(text), xlsx_ops::ParsedInput::Text(t) if t == text)
}

/// the `CellState` a proposed/accepted edit stores, carrying over the target
/// cell's existing style so a value edit keeps its number format, as excel does.
fn edit_cell_state(wb: &Workbook, sheet: SheetId, at: CellRef, input: &str) -> CellState {
    let mut cell = cell_state_for_input_no_eval(input);
    cell.style = wb
        .sheet(sheet)
        .and_then(|s| s.cell(at))
        .and_then(|c| c.style);
    cell
}

/// number-format-aware display text for a cell, matching what the grid paints;
/// empty or missing cells yield "".
fn display_text_at(wb: &Workbook, sheet: SheetId, at: CellRef) -> Result<String, String> {
    let s = wb
        .sheet(sheet)
        .ok_or_else(|| format!("sheet {} out of range", sheet.0))?;
    Ok(match s.cell(at) {
        Some(cell) => display_text(&wb.styles, wb.date_system, cell),
        None => String::new(),
    })
}

impl Session {
    /// open a workbook from raw `.xlsx` bytes, rebuild the dep graph, and recalc
    /// every formula, replacing file-shipped cached values with our engine's.
    pub fn open(bytes: &[u8], now_serial: Option<f64>) -> Result<Session, String> {
        let parts = ooxml_opc::unzip_parts(bytes)?;
        let mut workbook = xlsx_parse::parse_workbook(&parts).map_err(|e| e.to_string())?;
        if workbook.sheets.is_empty() {
            return Err("workbook has no sheets".to_string());
        }
        let (graph, _) = rebuild_and_recalc_all(&mut workbook, now_serial);
        Ok(Session {
            workbook,
            sheet: SheetId(0),
            undo: UndoStack::new(),
            graph,
            proposals: ProposalSet::new(),
        })
    }

    /// serialized `DisplayList` for a serialized `Viewport`.
    pub fn display_list_json(&self, viewport_json: &str) -> Result<String, String> {
        let viewport: Viewport =
            serde_json::from_str(viewport_json).map_err(|e| format!("bad viewport: {e}"))?;
        let dl = build_display_list(&self.workbook, self.sheet, &viewport);
        serde_json::to_string(&dl).map_err(|e| e.to_string())
    }

    /// render the current sheet viewport to png bytes via the tiny-skia backend.
    #[cfg(feature = "raster")]
    pub fn render_png(&self, viewport_json: &str) -> Result<Vec<u8>, String> {
        let viewport: Viewport =
            serde_json::from_str(viewport_json).map_err(|e| format!("bad viewport: {e}"))?;
        let dl = build_display_list(&self.workbook, self.sheet, &viewport);
        xlsx_raster::render_png(&dl)
    }

    /// render an a1 range (default: used range) at an optional scale to png;
    /// dimensions are capped so a hostile range cannot force an unbounded pixmap.
    #[cfg(feature = "raster")]
    pub fn render_range_png(&self, args: &str) -> Result<Vec<u8>, String> {
        #[derive(serde::Deserialize)]
        struct RangeArgs {
            range: Option<String>,
            scale: Option<f32>,
        }
        const MAX_DIM_PX: f32 = 16_384.0;

        let a: RangeArgs = serde_json::from_str(args).map_err(|e| format!("bad args: {e}"))?;
        let sheet = self
            .workbook
            .sheet(self.sheet)
            .ok_or_else(|| "active sheet out of range".to_string())?;
        let viewport = match &a.range {
            Some(r) => {
                let range =
                    xlsx_model::CellRange::parse_a1(r).map_err(|e| format!("bad range: {e}"))?;
                xlsx_render::viewport_for_range(sheet, range)
            }
            None => xlsx_render::viewport_for_used_range(sheet),
        };
        let scale = a.scale.unwrap_or(1.0);
        if !(scale > 0.0 && scale.is_finite()) {
            return Err("scale must be a positive number".to_string());
        }
        if viewport.width * scale > MAX_DIM_PX || viewport.height * scale > MAX_DIM_PX {
            return Err(format!(
                "requested render exceeds the {MAX_DIM_PX}px per-side cap; narrow the range or lower scale"
            ));
        }
        let dl = build_display_list(&self.workbook, self.sheet, &viewport);
        xlsx_raster::render_png(&xlsx_render::scaled(dl, scale))
    }

    /// sheet names, active index, and the pixel extent of the active sheet's
    /// used range plus one blank row/col of slack.
    pub fn sheet_info_json(&self) -> Result<String, String> {
        serde_json::to_string(&self.sheet_info()?).map_err(|e| e.to_string())
    }

    fn sheet_info(&self) -> Result<SheetInfo, String> {
        let sheet = self
            .workbook
            .sheet(self.sheet)
            .ok_or_else(|| "active sheet out of range".to_string())?;
        let geom = GridGeometry::new(sheet);
        let (content_width, content_height) = match sheet.used_range() {
            Some(range) => (geom.col_x(range.end.col + 2), geom.row_y(range.end.row + 2)),
            None => (geom.col_x(26), geom.row_y(50)),
        };
        Ok(SheetInfo {
            sheet_names: self
                .workbook
                .sheets
                .iter()
                .map(|s| s.name.clone())
                .collect(),
            active_sheet: self.sheet.0,
            content_width,
            content_height,
        })
    }

    /// enter one cell edit as a user `SetCell`, update the graph, and recalc
    /// the touched cell and its dependents; returns the edit envelope.
    pub fn edit_cell_json(
        &mut self,
        args: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let a: EditArgs = serde_json::from_str(args).map_err(|e| format!("bad edit args: {e}"))?;
        let sheet = SheetId(a.sheet);
        let at = CellRef::new(a.row, a.col);
        let cell = cell_state_for_input_no_eval(&a.input);
        let formula = cell.formula.clone();
        self.commit_user(vec![Op::SetCell { sheet, at, cell }])?;
        self.graph.set_formula(sheet, at, formula.as_deref());
        let seeds = [(sheet, at)];
        let result = recalc_after(&mut self.workbook, &mut self.graph, &seeds, now_serial);
        let changed = self.changed_list(&result, &seeds);
        self.edit_result(true, changed)
    }

    /// enter a batch of cell edits (the paste path) as a single undo step, with
    /// one recalc seeded by all of them so intra-batch references chain correctly.
    pub fn edit_cells_json(
        &mut self,
        args: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let a: EditBatchArgs =
            serde_json::from_str(args).map_err(|e| format!("bad edit args: {e}"))?;
        let sheet = SheetId(a.sheet);
        let mut touched: Vec<(SheetId, CellRef, Option<String>)> =
            Vec::with_capacity(a.edits.len());
        let ops = a
            .edits
            .iter()
            .map(|e| {
                let at = CellRef::new(e.row, e.col);
                let cell = cell_state_for_input_no_eval(&e.input);
                touched.push((sheet, at, cell.formula.clone()));
                Op::SetCell { sheet, at, cell }
            })
            .collect();
        self.commit_user(ops)?;
        for (s, at, formula) in &touched {
            self.graph.set_formula(*s, *at, formula.as_deref());
        }
        let seeds: Vec<(SheetId, CellRef)> = touched.iter().map(|(s, at, _)| (*s, *at)).collect();
        let result = recalc_after(&mut self.workbook, &mut self.graph, &seeds, now_serial);
        let changed = self.changed_list(&result, &seeds);
        self.edit_result(true, changed)
    }

    /// raw op-list escape hatch for structural edits, applied as one user
    /// transaction; the graph is rebuilt and every formula recalculated.
    pub fn apply_ops_json(
        &mut self,
        transaction_json: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let a: OpsArgs =
            serde_json::from_str(transaction_json).map_err(|e| format!("bad ops: {e}"))?;
        self.commit_user(a.ops)?;
        let changed = self.rebuild_and_recalc(now_serial);
        self.edit_result(true, changed)
    }

    /// register an agent proposal without touching the workbook: a clone is
    /// edited and recalculated so `oldText`/`newText` match what the grid paints.
    pub fn propose_json(&mut self, args: &str, now_serial: Option<f64>) -> Result<String, String> {
        let a: ProposeArgs =
            serde_json::from_str(args).map_err(|e| format!("bad propose args: {e}"))?;

        let mut preview = self.workbook.clone();
        for e in &a.edits {
            let sheet = SheetId(e.sheet);
            let at = CellRef::new(e.row, e.col);
            let cell = edit_cell_state(&preview, sheet, at, &e.input);
            let s = preview
                .sheet_mut(sheet)
                .ok_or_else(|| format!("sheet {} out of range", e.sheet))?;
            s.set_cell(at, cell.into());
        }
        rebuild_and_recalc_all(&mut preview, now_serial);

        let mut edits = Vec::with_capacity(a.edits.len());
        for e in &a.edits {
            let sheet = SheetId(e.sheet);
            let at = CellRef::new(e.row, e.col);
            edits.push(ProposedEdit {
                sheet: e.sheet,
                row: e.row,
                col: e.col,
                input: e.input.clone(),
                a1: at.to_a1(),
                old_text: display_text_at(&self.workbook, sheet, at)?,
                new_text: display_text_at(&preview, sheet, at)?,
            });
        }

        let id = self.proposals.next_id();
        let proposal = Proposal {
            id,
            agent_id: a.agent_id,
            note: a.note,
            edits,
        };
        let json = serde_json::to_string(&proposal).map_err(|e| e.to_string())?;
        self.proposals.add(proposal);
        Ok(json)
    }

    /// the pending proposals under a `proposals` key.
    pub fn list_proposals_json(&self) -> Result<String, String> {
        serde_json::to_string(&ProposalList {
            proposals: self.proposals.list(),
        })
        .map_err(|e| e.to_string())
    }

    /// accept a proposal as one undo-able agent transaction and recalc. unless
    /// `force`, errors `stale: <a1 list>` when any target's display text moved.
    pub fn accept_proposal_json(
        &mut self,
        args: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let a: AcceptArgs =
            serde_json::from_str(args).map_err(|e| format!("bad accept args: {e}"))?;
        let proposal = self
            .proposals
            .list()
            .iter()
            .find(|p| p.id == a.id)
            .cloned()
            .ok_or_else(|| format!("no proposal {}", a.id))?;

        if !a.force {
            let stale: Vec<String> = proposal
                .edits
                .iter()
                .filter_map(|e| {
                    let cur = display_text_at(
                        &self.workbook,
                        SheetId(e.sheet),
                        CellRef::new(e.row, e.col),
                    )
                    .ok()?;
                    (cur != e.old_text).then(|| e.a1.clone())
                })
                .collect();
            if !stale.is_empty() {
                return Err(format!("stale: {}", stale.join(", ")));
            }
        }

        let mut touched: Vec<(SheetId, CellRef, Option<String>)> =
            Vec::with_capacity(proposal.edits.len());
        let ops = proposal
            .edits
            .iter()
            .map(|e| {
                let sheet = SheetId(e.sheet);
                let at = CellRef::new(e.row, e.col);
                let cell = edit_cell_state(&self.workbook, sheet, at, &e.input);
                touched.push((sheet, at, cell.formula.clone()));
                Op::SetCell { sheet, at, cell }
            })
            .collect();
        self.commit_agent(ops, proposal.agent_id.clone())?;
        for (s, at, formula) in &touched {
            self.graph.set_formula(*s, *at, formula.as_deref());
        }
        let seeds: Vec<(SheetId, CellRef)> = touched.iter().map(|(s, at, _)| (*s, *at)).collect();
        let result = recalc_after(&mut self.workbook, &mut self.graph, &seeds, now_serial);
        let changed = self.changed_list(&result, &seeds);
        self.proposals.remove(&a.id);

        serde_json::to_string(&AcceptResult {
            applied: true,
            sheet_info: self.sheet_info()?,
            changed,
            proposal_id: a.id,
        })
        .map_err(|e| e.to_string())
    }

    /// reject a proposal by id: `{"id":string}` -> `{"removed":bool}`.
    pub fn reject_proposal_json(&mut self, args: &str) -> Result<String, String> {
        let a: IdArgs = serde_json::from_str(args).map_err(|e| format!("bad reject args: {e}"))?;
        let removed = self.proposals.remove(&a.id);
        serde_json::to_string(&RejectResult { removed }).map_err(|e| e.to_string())
    }

    /// apply `ops` as one user transaction on the undo stack.
    fn commit_user(&mut self, ops: Vec<Op>) -> Result<(), String> {
        let tx = Transaction::new(ops, Provenance::User);
        self.undo
            .commit(&mut self.workbook, &tx)
            .map_err(|e| e.to_string())
    }

    /// apply `ops` as one agent transaction; undo-able like any other.
    fn commit_agent(&mut self, ops: Vec<Op>, agent_id: String) -> Result<(), String> {
        let tx = Transaction::new(ops, Provenance::Agent { id: agent_id });
        self.undo
            .commit(&mut self.workbook, &tx)
            .map_err(|e| e.to_string())
    }

    /// reverse the most recent transaction, then rebuild the graph and recalc;
    /// `applied` is false when there was nothing to undo.
    pub fn undo_json(&mut self, now_serial: Option<f64>) -> Result<String, String> {
        let applied = self
            .undo
            .undo(&mut self.workbook)
            .map_err(|e| e.to_string())?
            .is_some();
        let changed = if applied {
            self.rebuild_and_recalc(now_serial)
        } else {
            Vec::new()
        };
        self.edit_result(applied, changed)
    }

    /// re-apply the most recently undone transaction; same shape as `undo_json`.
    pub fn redo_json(&mut self, now_serial: Option<f64>) -> Result<String, String> {
        let applied = self
            .undo
            .redo(&mut self.workbook)
            .map_err(|e| e.to_string())?
            .is_some();
        let changed = if applied {
            self.rebuild_and_recalc(now_serial)
        } else {
            Vec::new()
        };
        self.edit_result(applied, changed)
    }

    /// rebuild the graph and recalc every formula; returns changed a1 addresses.
    fn rebuild_and_recalc(&mut self, now_serial: Option<f64>) -> Vec<String> {
        let (graph, result) = rebuild_and_recalc_all(&mut self.workbook, now_serial);
        self.graph = graph;
        self.changed_list(&result, &[])
    }

    /// the recalc's changed cells as a1 strings, excluding the edit's own seeds.
    fn changed_list(&self, result: &RecalcResult, seeds: &[(SheetId, CellRef)]) -> Vec<String> {
        let seed_set: HashSet<(u32, u32, u32)> =
            seeds.iter().map(|(s, c)| (s.0, c.row, c.col)).collect();
        result
            .changed
            .iter()
            .filter(|(s, c)| !seed_set.contains(&(s.0, c.row, c.col)))
            .map(|(s, c)| self.a1_with_sheet(*s, *c))
            .collect()
    }

    /// `A1` on the active sheet, else `Sheet!A1`.
    fn a1_with_sheet(&self, sheet: SheetId, cell: CellRef) -> String {
        if sheet == self.sheet {
            cell.to_a1()
        } else {
            let name = self
                .workbook
                .sheet(sheet)
                .map(|s| s.name.as_str())
                .unwrap_or_default();
            format!("{name}!{}", cell.to_a1())
        }
    }

    fn edit_result(&self, applied: bool, changed: Vec<String>) -> Result<String, String> {
        let result = EditResult {
            applied,
            sheet_info: self.sheet_info()?,
            changed,
        };
        serde_json::to_string(&result).map_err(|e| e.to_string())
    }

    /// the editable representation of one cell.
    pub fn cell_json(&self, args: &str) -> Result<String, String> {
        let a: CellArgs = serde_json::from_str(args).map_err(|e| format!("bad cell args: {e}"))?;
        let cell = self.cell_edit(SheetId(a.sheet), CellRef::new(a.row, a.col))?;
        serde_json::to_string(&cell).map_err(|e| e.to_string())
    }

    /// a rectangular block of cells for clipboard copy; rejects ranges over
    /// `MAX_RANGE_CELLS`.
    pub fn range_cells_json(&self, args: &str) -> Result<String, String> {
        let a: RangeArgs =
            serde_json::from_str(args).map_err(|e| format!("bad range args: {e}"))?;
        let sheet = SheetId(a.sheet);
        let range = CellRange::parse_a1(&a.range).map_err(|e| format!("bad range: {e}"))?;
        let rows = u64::from(range.end.row - range.start.row + 1);
        let cols = u64::from(range.end.col - range.start.col + 1);
        if rows * cols > MAX_RANGE_CELLS {
            return Err(format!(
                "range {rows}x{cols} exceeds the {MAX_RANGE_CELLS}-cell copy cap"
            ));
        }
        let mut grid = Vec::with_capacity(rows as usize);
        for r in range.start.row..=range.end.row {
            let mut row_cells = Vec::with_capacity(cols as usize);
            for c in range.start.col..=range.end.col {
                row_cells.push(self.cell_edit(sheet, CellRef::new(r, c))?);
            }
            grid.push(row_cells);
        }
        serde_json::to_string(&RangeCells { cells: grid }).map_err(|e| e.to_string())
    }

    fn cell_edit(&self, sheet: SheetId, at: CellRef) -> Result<CellEdit, String> {
        let s = self
            .workbook
            .sheet(sheet)
            .ok_or_else(|| format!("sheet {} out of range", sheet.0))?;
        let (input, is_formula) = match s.cell(at) {
            Some(c) => match &c.formula {
                Some(f) => (format!("={f}"), true),
                None => (value_to_input(&c.value), false),
            },
            None => (String::new(), false),
        };
        Ok(CellEdit {
            a1: at.to_a1(),
            input,
            is_formula,
        })
    }

    /// serialize the current workbook back to `.xlsx` bytes.
    pub fn save(&self) -> Result<Vec<u8>, String> {
        let parts = xlsx_parse::serialize_workbook(&self.workbook).map_err(|e| e.to_string())?;
        ooxml_opc::rezip_parts(&parts)
    }

    /// switch the active sheet by index.
    pub fn set_active_sheet(&mut self, index: u32) -> Result<(), String> {
        if (index as usize) >= self.workbook.sheets.len() {
            return Err(format!("sheet index {index} out of range"));
        }
        self.sheet = SheetId(index);
        Ok(())
    }

    /// crate version, for wasm/js parity checks.
    pub fn version() -> &'static str {
        env!("CARGO_PKG_VERSION")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::CellRef;
    use xlsx_model::value::CellValue;
    use xlsx_model::workbook::{Cell, Sheet};

    /// real `.xlsx` bytes via our own serializer + container.
    fn sample_xlsx() -> Vec<u8> {
        let mut sheet = Sheet::new("Data");
        sheet.set_cell(
            CellRef::parse_a1("A1").unwrap(),
            Cell {
                value: CellValue::Text {
                    value: "Hello".into(),
                },
                ..Cell::default()
            },
        );
        sheet.set_cell(
            CellRef::parse_a1("B2").unwrap(),
            Cell {
                value: CellValue::Number { value: 42.0 },
                ..Cell::default()
            },
        );
        let mut second = Sheet::new("Empty");
        second.name = "Empty".into();
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        wb.sheets.push(second);
        let parts = xlsx_parse::serialize_workbook(&wb).unwrap();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    #[test]
    fn opens_real_bytes_and_renders() {
        let session = Session::open(&sample_xlsx(), None).unwrap();
        let vp = r#"{"x":0,"y":0,"width":300,"height":120}"#;
        let json = session.display_list_json(vp).unwrap();
        assert!(json.contains("\"commands\""));
        assert!(json.contains("\"op\":\"fillRect\""));
        assert!(json.contains("Hello"));
    }

    #[test]
    fn sheet_info_and_switching() {
        let mut session = Session::open(&sample_xlsx(), None).unwrap();
        let info = session.sheet_info_json().unwrap();
        assert!(info.contains("\"sheetNames\":[\"Data\",\"Empty\"]"));
        assert!(info.contains("\"activeSheet\":0"));
        assert!(info.contains("\"contentWidth\""));

        session.set_active_sheet(1).unwrap();
        let info = session.sheet_info_json().unwrap();
        assert!(info.contains("\"activeSheet\":1"));
        assert!(session.set_active_sheet(9).is_err());
    }

    #[test]
    fn garbage_bytes_error_not_panic() {
        assert!(Session::open(b"not a zip", None).is_err());
        assert!(Session::open(&[], None).is_err());
    }

    #[test]
    fn bad_viewport_is_an_error_not_a_panic() {
        let session = Session::open(&sample_xlsx(), None).unwrap();
        let err = session.display_list_json("not json").unwrap_err();
        assert!(err.contains("bad viewport"));
    }

    #[test]
    fn version_is_populated() {
        assert!(!Session::version().is_empty());
    }

    #[test]
    fn edit_cell_round_trips_through_cell_json() {
        let mut s = Session::open(&sample_xlsx(), None).unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"123"}"#, None)
            .unwrap();
        let c = s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#).unwrap();
        assert!(c.contains(r#""a1":"A1""#));
        assert!(c.contains(r#""input":"123""#));
        assert!(c.contains(r#""isFormula":false"#));

        // text that would re-parse as a number round-trips with a leading quote.
        s.edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"'42"}"#, None)
            .unwrap();
        let c = s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#).unwrap();
        assert!(c.contains(r#""input":"'42""#), "got {c}");
    }

    #[test]
    fn formula_entry_evaluates_and_reads_back() {
        let mut s = Session::open(&sample_xlsx(), None).unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"1"}"#, None)
            .unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":1,"col":0,"input":"2"}"#, None)
            .unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":2,"col":0,"input":"=SUM(A1:A2)"}"#, None)
            .unwrap();

        let c = s.cell_json(r#"{"sheet":0,"row":2,"col":0}"#).unwrap();
        assert!(c.contains(r#""input":"=SUM(A1:A2)""#), "got {c}");
        assert!(c.contains(r#""isFormula":true"#));

        let dl = session_dl(&s);
        assert!(dl.contains(r#""text":"3""#), "sum should render as 3: {dl}");
    }

    #[test]
    fn batch_edit_is_single_undo_step() {
        let mut s = Session::open(&sample_xlsx(), None).unwrap();
        s.edit_cells_json(
            r#"{"sheet":0,"edits":[{"row":0,"col":0,"input":"x"},{"row":0,"col":1,"input":"y"}]}"#,
            None,
        )
        .unwrap();
        assert!(
            s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"x""#)
        );

        let res = s.undo_json(None).unwrap();
        assert!(res.contains(r#""applied":true"#));
        assert!(
            s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"Hello""#)
        );
        assert!(
            s.cell_json(r#"{"sheet":0,"row":0,"col":1}"#)
                .unwrap()
                .contains(r#""input":"""#)
        );
    }

    #[test]
    fn undo_redo_restores_and_replays() {
        let mut s = Session::open(&sample_xlsx(), None).unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"changed"}"#, None)
            .unwrap();

        s.undo_json(None).unwrap();
        assert!(
            s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"Hello""#),
            "undo restores the prior value"
        );

        let res = s.redo_json(None).unwrap();
        assert!(res.contains(r#""applied":true"#));
        assert!(
            s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"changed""#),
            "redo replays the edit"
        );
    }

    #[test]
    fn undo_with_empty_history_reports_not_applied() {
        let mut s = Session::open(&sample_xlsx(), None).unwrap();
        assert!(s.undo_json(None).unwrap().contains(r#""applied":false"#));
    }

    #[test]
    fn structural_op_shifts_cell() {
        let mut s = Session::open(&sample_xlsx(), None).unwrap();
        s.apply_ops_json(
            r#"{"ops":[{"type":"insertRows","sheet":0,"at":0,"count":1}]}"#,
            None,
        )
        .unwrap();
        assert!(
            s.cell_json(r#"{"sheet":0,"row":2,"col":1}"#)
                .unwrap()
                .contains(r#""input":"42""#),
            "42 shifted down to B3"
        );
        assert!(
            s.cell_json(r#"{"sheet":0,"row":1,"col":1}"#)
                .unwrap()
                .contains(r#""input":"""#),
            "B2 is now empty"
        );
    }

    #[test]
    fn save_reopens_with_edits_intact() {
        let mut s = Session::open(&sample_xlsx(), None).unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":4,"col":4,"input":"persisted"}"#, None)
            .unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":5,"col":4,"input":"777"}"#, None)
            .unwrap();

        let bytes = s.save().unwrap();
        let reopened = Session::open(&bytes, None).unwrap();
        assert!(
            reopened
                .cell_json(r#"{"sheet":0,"row":4,"col":4}"#)
                .unwrap()
                .contains(r#""input":"persisted""#)
        );
        assert!(
            reopened
                .cell_json(r#"{"sheet":0,"row":5,"col":4}"#)
                .unwrap()
                .contains(r#""input":"777""#)
        );
    }

    #[test]
    fn range_cells_shape_and_cap() {
        let s = Session::open(&sample_xlsx(), None).unwrap();
        let out = s
            .range_cells_json(r#"{"sheet":0,"range":"A1:B2"}"#)
            .unwrap();
        assert!(out.contains(r#""input":"Hello""#));
        assert!(out.contains(r#""input":"42""#));
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        let rows = v["cells"].as_array().unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].as_array().unwrap().len(), 2);

        let err = s
            .range_cells_json(r#"{"sheet":0,"range":"A1:D100000"}"#)
            .unwrap_err();
        assert!(err.contains("cap"), "got {err}");
    }

    /// `.xlsx` bytes where B1 = SUM(A1:A2) ships a deliberately wrong cached
    /// value (999 instead of 15), so the open-time recalc has something to correct.
    fn xlsx_with_stale_formula() -> Vec<u8> {
        let mut sheet = Sheet::new("Data");
        sheet.set_cell(
            CellRef::parse_a1("A1").unwrap(),
            Cell {
                value: CellValue::Number { value: 10.0 },
                ..Cell::default()
            },
        );
        sheet.set_cell(
            CellRef::parse_a1("A2").unwrap(),
            Cell {
                value: CellValue::Number { value: 5.0 },
                ..Cell::default()
            },
        );
        sheet.set_cell(
            CellRef::parse_a1("B1").unwrap(),
            Cell {
                value: CellValue::Number { value: 999.0 },
                formula: Some("SUM(A1:A2)".into()),
                style: None,
            },
        );
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        let parts = xlsx_parse::serialize_workbook(&wb).unwrap();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    #[test]
    fn open_corrects_stale_cached_formula_value() {
        let s = Session::open(&xlsx_with_stale_formula(), None).unwrap();
        let dl = session_dl(&s);
        assert!(
            dl.contains(r#""text":"15""#),
            "B1 should recompute to 15: {dl}"
        );
        assert!(!dl.contains(r#""text":"999""#), "stale cache must be gone");
    }

    #[test]
    fn edit_input_recalcs_dependent_and_reports_changed() {
        let mut s = Session::open(&xlsx_with_stale_formula(), None).unwrap();
        let res = s
            .edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"20"}"#, None)
            .unwrap();
        assert!(res.contains(r#""changed":["B1"]"#), "got {res}");
        let dl = session_dl(&s);
        assert!(dl.contains(r#""text":"25""#), "B1 should be 25: {dl}");
    }

    #[test]
    fn round_trip_persists_recalculated_value() {
        let mut s = Session::open(&xlsx_with_stale_formula(), None).unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"20"}"#, None)
            .unwrap();
        let bytes = s.save().unwrap();
        let reopened = Session::open(&bytes, None).unwrap();
        let dl = session_dl(&reopened);
        assert!(
            dl.contains(r#""text":"25""#),
            "persisted B1 should be 25: {dl}"
        );
    }

    #[test]
    fn structural_op_remaps_formula_and_undo_restores() {
        let mut s = Session::open(&xlsx_with_stale_formula(), None).unwrap();
        s.apply_ops_json(
            r#"{"ops":[{"type":"insertRows","sheet":0,"at":0,"count":1}]}"#,
            None,
        )
        .unwrap();
        let c = s.cell_json(r#"{"sheet":0,"row":1,"col":1}"#).unwrap();
        assert!(
            c.contains(r#""input":"=SUM(A2:A3)""#),
            "formula should remap to SUM(A2:A3): {c}"
        );

        s.undo_json(None).unwrap();
        let c = s.cell_json(r#"{"sheet":0,"row":0,"col":1}"#).unwrap();
        assert!(
            c.contains(r#""input":"=SUM(A1:A2)""#),
            "undo restores SUM(A1:A2): {c}"
        );
    }

    fn session_dl(s: &Session) -> String {
        s.display_list_json(r#"{"x":0,"y":0,"width":400,"height":200}"#)
            .unwrap()
    }

    /// `.xlsx` bytes with a currency-formatted A1 (`$#,##0.00`, value 1000) and
    /// a plain A2 (250), for format-aware proposal previews.
    fn xlsx_with_currency() -> Vec<u8> {
        use xlsx_model::styles::Xf;
        let mut sheet = Sheet::new("Data");
        sheet.set_cell(
            CellRef::parse_a1("A1").unwrap(),
            Cell {
                value: CellValue::Number { value: 1000.0 },
                formula: None,
                style: Some(1),
            },
        );
        sheet.set_cell(
            CellRef::parse_a1("A2").unwrap(),
            Cell {
                value: CellValue::Number { value: 250.0 },
                ..Cell::default()
            },
        );
        let mut wb = Workbook::default();
        wb.styles.num_fmts.push((164, "$#,##0.00".into()));
        wb.styles.cell_xfs.push(Xf::default());
        wb.styles.cell_xfs.push(Xf {
            num_fmt_id: Some(164),
            ..Xf::default()
        });
        wb.sheets.push(sheet);
        let parts = xlsx_parse::serialize_workbook(&wb).unwrap();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    #[test]
    fn propose_leaves_workbook_untouched_with_formatted_texts() {
        let mut s = Session::open(&xlsx_with_currency(), None).unwrap();
        let json = s
            .propose_json(
                r#"{"agentId":"agent-1","note":"bump it","edits":[{"sheet":0,"row":0,"col":0,"input":"2000"}]}"#,
                None,
            )
            .unwrap();
        assert!(json.contains(r#""oldText":"$1,000.00""#), "got {json}");
        assert!(json.contains(r#""newText":"$2,000.00""#), "got {json}");
        assert!(json.contains(r#""a1":"A1""#));
        assert!(json.contains(r#""agentId":"agent-1""#));
        assert!(json.contains(r#""note":"bump it""#));
        assert!(
            json.contains(r#""cells":[{"#),
            "edits serialize as cells: {json}"
        );

        assert!(
            s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"1000""#),
            "workbook must not change on propose"
        );
    }

    #[test]
    fn propose_previews_evaluated_formula_value() {
        let mut s = Session::open(&xlsx_with_currency(), None).unwrap();
        let json = s
            .propose_json(
                r#"{"agentId":"a","note":null,"edits":[{"sheet":0,"row":0,"col":1,"input":"=A1+A2"}]}"#,
                None,
            )
            .unwrap();
        assert!(
            json.contains(r#""newText":"1250""#),
            "formula preview: {json}"
        );
        assert!(json.contains(r#""oldText":"""#), "b1 starts empty: {json}");
        assert!(
            json.contains(r#""note":null"#),
            "null note stays null: {json}"
        );
    }

    #[test]
    fn accept_applies_recalcs_lands_on_undo_stack() {
        let mut s = Session::open(&xlsx_with_currency(), None).unwrap();
        s.edit_cell_json(r#"{"sheet":0,"row":0,"col":1,"input":"=A1+A2"}"#, None)
            .unwrap();
        let p = s
            .propose_json(
                r#"{"agentId":"agent-1","note":null,"edits":[{"sheet":0,"row":0,"col":0,"input":"3000"}]}"#,
                None,
            )
            .unwrap();
        let id: serde_json::Value = serde_json::from_str(&p).unwrap();
        let id = id["id"].as_str().unwrap();

        let res = s
            .accept_proposal_json(&format!(r#"{{"id":"{id}"}}"#), None)
            .unwrap();
        assert!(
            res.contains(&format!(r#""proposalId":"{id}""#)),
            "got {res}"
        );
        assert!(
            res.contains(r#""changed":["B1"]"#),
            "B1 recalculated: {res}"
        );

        let dl = session_dl(&s);
        assert!(dl.contains(r#""text":"$3,000.00""#), "A1 formatted: {dl}");
        assert!(dl.contains(r#""text":"3250""#), "B1 = 3000+250: {dl}");

        assert!(
            s.list_proposals_json()
                .unwrap()
                .contains(r#""proposals":[]"#)
        );
        s.undo_json(None).unwrap();
        let dl = session_dl(&s);
        assert!(
            dl.contains(r#""text":"$1,000.00""#),
            "undo reverts the agent edit: {dl}"
        );
    }

    #[test]
    fn reject_drops_proposal() {
        let mut s = Session::open(&xlsx_with_currency(), None).unwrap();
        let p = s
            .propose_json(
                r#"{"agentId":"a","note":null,"edits":[{"sheet":0,"row":0,"col":0,"input":"9"}]}"#,
                None,
            )
            .unwrap();
        let id: serde_json::Value = serde_json::from_str(&p).unwrap();
        let id = id["id"].as_str().unwrap().to_string();

        assert!(
            s.reject_proposal_json(&format!(r#"{{"id":"{id}"}}"#))
                .unwrap()
                .contains(r#""removed":true"#)
        );
        assert!(
            s.reject_proposal_json(&format!(r#"{{"id":"{id}"}}"#))
                .unwrap()
                .contains(r#""removed":false"#)
        );
        assert!(
            s.list_proposals_json()
                .unwrap()
                .contains(r#""proposals":[]"#)
        );
        assert!(
            s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"1000""#)
        );
    }

    #[test]
    fn accept_is_stale_when_base_moved_and_force_overrides() {
        let mut s = Session::open(&xlsx_with_currency(), None).unwrap();
        let p = s
            .propose_json(
                r#"{"agentId":"a","note":null,"edits":[{"sheet":0,"row":0,"col":0,"input":"2000"}]}"#,
                None,
            )
            .unwrap();
        let id: serde_json::Value = serde_json::from_str(&p).unwrap();
        let id = id["id"].as_str().unwrap().to_string();

        s.edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"1500"}"#, None)
            .unwrap();
        let err = s
            .accept_proposal_json(&format!(r#"{{"id":"{id}"}}"#), None)
            .unwrap_err();
        assert!(err.starts_with("stale: "), "got {err}");
        assert!(err.contains("A1"), "stale list names A1: {err}");
        assert!(
            s.list_proposals_json()
                .unwrap()
                .contains(&format!(r#""id":"{id}""#))
        );

        let res = s
            .accept_proposal_json(&format!(r#"{{"id":"{id}","force":true}}"#), None)
            .unwrap();
        assert!(res.contains(r#""applied":true"#), "got {res}");
        let c = s.cell_json(r#"{"sheet":0,"row":0,"col":0}"#).unwrap();
        assert!(c.contains(r#""input":"2000""#), "force applied: {c}");
    }

    #[test]
    fn proposal_ids_increment_and_list_shape() {
        let mut s = Session::open(&xlsx_with_currency(), None).unwrap();
        let p1 = s
            .propose_json(
                r#"{"agentId":"a","note":null,"edits":[{"sheet":0,"row":0,"col":0,"input":"1"}]}"#,
                None,
            )
            .unwrap();
        let p2 = s
            .propose_json(
                r#"{"agentId":"a","note":null,"edits":[{"sheet":0,"row":1,"col":0,"input":"2"}]}"#,
                None,
            )
            .unwrap();
        assert!(p1.contains(r#""id":"p1""#), "got {p1}");
        assert!(p2.contains(r#""id":"p2""#), "got {p2}");

        let list = s.list_proposals_json().unwrap();
        let v: serde_json::Value = serde_json::from_str(&list).unwrap();
        let arr = v["proposals"].as_array().unwrap();
        assert_eq!(arr.len(), 2);
        assert_eq!(arr[0]["id"], "p1");
        assert_eq!(arr[1]["id"], "p2");
    }

    #[cfg(feature = "raster")]
    #[test]
    fn render_png_produces_png_bytes() {
        let s = Session::open(&sample_xlsx(), None).unwrap();
        let png = s
            .render_png(r#"{"x":0,"y":0,"width":240,"height":120}"#)
            .unwrap();
        assert!(png.len() > 8);
        assert_eq!(
            &png[0..8],
            &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]
        );
    }

    #[cfg(feature = "raster")]
    #[test]
    fn render_range_png_honors_range_scale_and_caps() {
        let s = Session::open(&sample_xlsx(), None).unwrap();
        let one = s.render_range_png(r#"{"range":"A1:B2"}"#).unwrap();
        let two = s
            .render_range_png(r#"{"range":"A1:B2","scale":2}"#)
            .unwrap();
        assert_eq!(&one[0..4], &[0x89, b'P', b'N', b'G']);
        assert!(two.len() > one.len());
        assert!(s.render_range_png("{}").is_ok());
        assert!(s.render_range_png(r#"{"range":"A1:XFD1048576"}"#).is_err());
        assert!(
            s.render_range_png(r#"{"range":"A1:B2","scale":0}"#)
                .is_err()
        );
        assert!(s.render_range_png(r#"{"range":"nope"}"#).is_err());
    }

    #[cfg(feature = "raster")]
    #[test]
    fn render_png_bad_viewport_is_an_error() {
        let s = Session::open(&sample_xlsx(), None).unwrap();
        assert!(
            s.render_png("not json")
                .unwrap_err()
                .contains("bad viewport")
        );
    }
}
