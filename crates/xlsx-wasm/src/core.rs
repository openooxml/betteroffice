#[cfg(feature = "raster")]
use betteroffice_xlsx::RenderOptions;
use betteroffice_xlsx::{
    CalculationOptions, CellAddress, CellInput as WorkbookCellInput, CellRange, CellRef,
    MutationResult, Op, Proposal, ProposalEditInput as WorkbookProposalEditInput, ProposalRequest,
    SheetId, Viewport, Workbook,
};
use serde::{Deserialize, Serialize};

pub struct Session {
    workbook: Workbook,
}

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

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct EditResult {
    applied: bool,
    sheet_info: SheetInfo,
    changed: Vec<String>,
    limited_cells: Vec<String>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct CalculationStatus {
    limited_cells: Vec<String>,
}

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

#[derive(Serialize)]
struct ProposalList<'a> {
    proposals: &'a [Proposal],
}

#[derive(Serialize)]
struct RejectResult {
    removed: bool,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct AcceptResult {
    applied: bool,
    sheet_info: SheetInfo,
    changed: Vec<String>,
    limited_cells: Vec<String>,
    proposal_id: String,
}

impl Session {
    pub fn open(bytes: &[u8], now_serial: Option<f64>) -> Result<Self, String> {
        Workbook::open_recalculated(bytes, calculation_options(now_serial))
            .map(|workbook| Self { workbook })
            .map_err(|error| error.to_string())
    }

    pub fn display_list_json(&self, viewport_json: &str) -> Result<String, String> {
        let viewport: Viewport = serde_json::from_str(viewport_json)
            .map_err(|error| format!("bad viewport: {error}"))?;
        let display_list = self
            .workbook
            .display_list(&viewport)
            .map_err(|error| error.to_string())?;
        serde_json::to_string(&display_list).map_err(|error| error.to_string())
    }

    #[cfg(feature = "raster")]
    pub fn render_png(&self, viewport_json: &str) -> Result<Vec<u8>, String> {
        let viewport: Viewport = serde_json::from_str(viewport_json)
            .map_err(|error| format!("bad viewport: {error}"))?;
        self.workbook
            .render_png(&viewport)
            .map(|rendered| rendered.bytes)
            .map_err(|error| error.to_string())
    }

    #[cfg(feature = "raster")]
    pub fn render_range_png(&self, args: &str) -> Result<Vec<u8>, String> {
        #[derive(Deserialize)]
        struct Args {
            range: Option<String>,
            scale: Option<f32>,
        }

        let args: Args =
            serde_json::from_str(args).map_err(|error| format!("bad args: {error}"))?;
        let range = args
            .range
            .map(|range| CellRange::parse_a1(&range).map_err(|error| format!("bad range: {error}")))
            .transpose()?;
        self.workbook
            .render_sheet(
                self.workbook.active_sheet(),
                &RenderOptions {
                    range,
                    scale: args.scale.unwrap_or(1.0),
                    ..RenderOptions::default()
                },
            )
            .map(|rendered| rendered.bytes)
            .map_err(|error| error.to_string())
    }

    pub fn sheet_info_json(&self) -> Result<String, String> {
        serde_json::to_string(&self.sheet_info()?).map_err(|error| error.to_string())
    }

    pub fn calculation_status_json(&self) -> Result<String, String> {
        serde_json::to_string(&CalculationStatus {
            limited_cells: self.changed_list(&self.workbook.last_calculation().limited_cells),
        })
        .map_err(|error| error.to_string())
    }

    pub fn set_active_sheet(&mut self, index: u32) -> Result<(), String> {
        self.workbook
            .set_active_sheet(SheetId(index))
            .map_err(|error| error.to_string())
    }

    pub fn edit_cell_json(
        &mut self,
        args: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let args: EditArgs =
            serde_json::from_str(args).map_err(|error| format!("bad edit args: {error}"))?;
        let result = self
            .workbook
            .edit_cell(
                SheetId(args.sheet),
                CellRef::new(args.row, args.col),
                &args.input,
                calculation_options(now_serial),
            )
            .map_err(|error| error.to_string())?;
        self.edit_result(result)
    }

    pub fn edit_cells_json(
        &mut self,
        args: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let args: EditBatchArgs =
            serde_json::from_str(args).map_err(|error| format!("bad edit args: {error}"))?;
        let edits = args
            .edits
            .into_iter()
            .map(|edit| WorkbookCellInput {
                cell: CellRef::new(edit.row, edit.col),
                input: edit.input,
            })
            .collect::<Vec<_>>();
        let result = self
            .workbook
            .edit_cells(SheetId(args.sheet), &edits, calculation_options(now_serial))
            .map_err(|error| error.to_string())?;
        self.edit_result(result)
    }

    pub fn apply_ops_json(
        &mut self,
        transaction_json: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let args: OpsArgs =
            serde_json::from_str(transaction_json).map_err(|error| format!("bad ops: {error}"))?;
        let result = self
            .workbook
            .apply_ops(args.ops, calculation_options(now_serial))
            .map_err(|error| error.to_string())?;
        self.edit_result(result)
    }

    pub fn undo_json(&mut self, now_serial: Option<f64>) -> Result<String, String> {
        let result = self
            .workbook
            .undo(calculation_options(now_serial))
            .map_err(|error| error.to_string())?;
        self.edit_result(result)
    }

    pub fn redo_json(&mut self, now_serial: Option<f64>) -> Result<String, String> {
        let result = self
            .workbook
            .redo(calculation_options(now_serial))
            .map_err(|error| error.to_string())?;
        self.edit_result(result)
    }

    pub fn cell_json(&self, args: &str) -> Result<String, String> {
        let args: CellArgs =
            serde_json::from_str(args).map_err(|error| format!("bad cell args: {error}"))?;
        let cell = self
            .workbook
            .cell(SheetId(args.sheet), CellRef::new(args.row, args.col))
            .map_err(|error| error.to_string())?;
        serde_json::to_string(&CellEdit {
            a1: cell.cell.to_a1(),
            input: cell.input,
            is_formula: cell.is_formula,
        })
        .map_err(|error| error.to_string())
    }

    pub fn range_cells_json(&self, args: &str) -> Result<String, String> {
        let args: RangeArgs =
            serde_json::from_str(args).map_err(|error| format!("bad range args: {error}"))?;
        let range =
            CellRange::parse_a1(&args.range).map_err(|error| format!("bad range: {error}"))?;
        let cells = self
            .workbook
            .range_cells(SheetId(args.sheet), range)
            .map_err(|error| error.to_string())?
            .into_iter()
            .map(|row| {
                row.into_iter()
                    .map(|cell| CellEdit {
                        a1: cell.cell.to_a1(),
                        input: cell.input,
                        is_formula: cell.is_formula,
                    })
                    .collect()
            })
            .collect();
        serde_json::to_string(&RangeCells { cells }).map_err(|error| error.to_string())
    }

    pub fn propose_json(&mut self, args: &str, now_serial: Option<f64>) -> Result<String, String> {
        let args: ProposeArgs =
            serde_json::from_str(args).map_err(|error| format!("bad propose args: {error}"))?;
        let proposal = self
            .workbook
            .propose(
                ProposalRequest {
                    agent_id: args.agent_id,
                    note: args.note,
                    edits: args
                        .edits
                        .into_iter()
                        .map(|edit| WorkbookProposalEditInput {
                            sheet: SheetId(edit.sheet),
                            cell: CellRef::new(edit.row, edit.col),
                            input: edit.input,
                        })
                        .collect(),
                },
                calculation_options(now_serial),
            )
            .map_err(|error| error.to_string())?;
        serde_json::to_string(&proposal).map_err(|error| error.to_string())
    }

    pub fn list_proposals_json(&self) -> Result<String, String> {
        serde_json::to_string(&ProposalList {
            proposals: self.workbook.proposals(),
        })
        .map_err(|error| error.to_string())
    }

    pub fn accept_proposal_json(
        &mut self,
        args: &str,
        now_serial: Option<f64>,
    ) -> Result<String, String> {
        let args: AcceptArgs =
            serde_json::from_str(args).map_err(|error| format!("bad accept args: {error}"))?;
        let result = self
            .workbook
            .accept_proposal(&args.id, args.force, calculation_options(now_serial))
            .map_err(|error| error.to_string())?;
        serde_json::to_string(&AcceptResult {
            applied: result.mutation.applied,
            sheet_info: self.sheet_info()?,
            changed: self.changed_list(&result.mutation.changed),
            limited_cells: self.changed_list(&result.mutation.limited_cells),
            proposal_id: result.proposal_id,
        })
        .map_err(|error| error.to_string())
    }

    pub fn reject_proposal_json(&mut self, args: &str) -> Result<String, String> {
        let args: IdArgs =
            serde_json::from_str(args).map_err(|error| format!("bad reject args: {error}"))?;
        serde_json::to_string(&RejectResult {
            removed: self.workbook.reject_proposal(&args.id),
        })
        .map_err(|error| error.to_string())
    }

    pub fn save(&self) -> Result<Vec<u8>, String> {
        self.workbook.save().map_err(|error| error.to_string())
    }

    pub fn version() -> &'static str {
        env!("CARGO_PKG_VERSION")
    }

    fn sheet_info(&self) -> Result<SheetInfo, String> {
        self.workbook
            .sheet_info()
            .map(|info| SheetInfo {
                sheet_names: info.sheet_names,
                active_sheet: info.active_sheet.0,
                content_width: info.content_width,
                content_height: info.content_height,
            })
            .map_err(|error| error.to_string())
    }

    fn edit_result(&self, result: MutationResult) -> Result<String, String> {
        serde_json::to_string(&EditResult {
            applied: result.applied,
            sheet_info: self.sheet_info()?,
            changed: self.changed_list(&result.changed),
            limited_cells: self.changed_list(&result.limited_cells),
        })
        .map_err(|error| error.to_string())
    }

    fn changed_list(&self, changed: &[CellAddress]) -> Vec<String> {
        changed
            .iter()
            .map(|address| self.workbook.format_address(*address))
            .collect()
    }
}

fn calculation_options(now_serial: Option<f64>) -> CalculationOptions {
    CalculationOptions { now_serial }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellValue, Sheet, Workbook as WorkbookModel};

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
        let mut model = WorkbookModel::default();
        model.sheets.push(sheet);
        model.sheets.push(Sheet::new("Empty"));
        let parts = xlsx_parse::serialize_workbook(&model).unwrap();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn formula_xlsx() -> Vec<u8> {
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
                ..Cell::default()
            },
        );
        let mut model = WorkbookModel::default();
        model.sheets.push(sheet);
        let parts = xlsx_parse::serialize_workbook(&model).unwrap();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn currency_xlsx() -> Vec<u8> {
        let mut sheet = Sheet::new("Data");
        sheet.set_cell(
            CellRef::parse_a1("A1").unwrap(),
            Cell {
                value: CellValue::Number { value: 1000.0 },
                style: Some(1),
                ..Cell::default()
            },
        );
        let mut model = WorkbookModel::default();
        model.styles.num_fmts.push((164, "$#,##0.00".into()));
        model.styles.cell_xfs.push(Default::default());
        model.styles.cell_xfs.push(xlsx_model::Xf {
            num_fmt_id: Some(164),
            ..Default::default()
        });
        model.sheets.push(sheet);
        let parts = xlsx_parse::serialize_workbook(&model).unwrap();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    #[test]
    fn preserves_display_and_sheet_info_wire_shapes() {
        let mut session = Session::open(&sample_xlsx(), None).unwrap();
        let display = session
            .display_list_json(r#"{"x":0,"y":0,"width":300,"height":120}"#)
            .unwrap();
        assert!(display.contains(r#""op":"fillRect""#));
        assert!(display.contains("Hello"));

        let info = session.sheet_info_json().unwrap();
        assert!(info.contains(r#""sheetNames":["Data","Empty"]"#));
        assert!(info.contains(r#""activeSheet":0"#));
        assert_eq!(
            session.calculation_status_json().unwrap(),
            r#"{"limitedCells":[]}"#
        );
        session.set_active_sheet(1).unwrap();
        assert!(
            session
                .sheet_info_json()
                .unwrap()
                .contains(r#""activeSheet":1"#)
        );
    }

    #[test]
    fn edits_recalculate_undo_and_save() {
        let mut session = Session::open(&sample_xlsx(), None).unwrap();
        session
            .edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"1"}"#, None)
            .unwrap();
        session
            .edit_cell_json(r#"{"sheet":0,"row":1,"col":0,"input":"2"}"#, None)
            .unwrap();
        session
            .edit_cell_json(r#"{"sheet":0,"row":2,"col":0,"input":"=SUM(A1:A2)"}"#, None)
            .unwrap();
        assert!(
            session
                .cell_json(r#"{"sheet":0,"row":2,"col":0}"#)
                .unwrap()
                .contains(r#""input":"=SUM(A1:A2)""#)
        );
        assert!(
            session
                .undo_json(None)
                .unwrap()
                .contains(r#""applied":true"#)
        );
        assert!(
            session
                .redo_json(None)
                .unwrap()
                .contains(r#""applied":true"#)
        );

        let bytes = session.save().unwrap();
        let reopened = Session::open(&bytes, None).unwrap();
        assert!(
            reopened
                .cell_json(r#"{"sheet":0,"row":2,"col":0}"#)
                .unwrap()
                .contains(r#""isFormula":true"#)
        );
    }

    #[test]
    fn quoted_text_and_batch_undo_keep_wire_behavior() {
        let mut session = Session::open(&sample_xlsx(), None).unwrap();
        session
            .edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"'42"}"#, None)
            .unwrap();
        assert!(
            session
                .cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"'42""#)
        );

        session
            .edit_cells_json(
                r#"{"sheet":0,"edits":[{"row":0,"col":0,"input":"x"},{"row":0,"col":1,"input":"y"}]}"#,
                None,
            )
            .unwrap();
        session.undo_json(None).unwrap();
        assert!(
            session
                .cell_json(r#"{"sheet":0,"row":0,"col":0}"#)
                .unwrap()
                .contains(r#""input":"'42""#)
        );
        assert!(
            session
                .cell_json(r#"{"sheet":0,"row":0,"col":1}"#)
                .unwrap()
                .contains(r#""input":"""#)
        );
    }

    #[test]
    fn changed_addresses_and_range_shape_stay_stable() {
        let mut session = Session::open(&formula_xlsx(), None).unwrap();
        let result = session
            .edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"20"}"#, None)
            .unwrap();
        assert!(result.contains(r#""changed":["B1"]"#), "{result}");
        let range = session
            .range_cells_json(r#"{"sheet":0,"range":"A1:B2"}"#)
            .unwrap();
        let value: serde_json::Value = serde_json::from_str(&range).unwrap();
        assert_eq!(value["cells"].as_array().unwrap().len(), 2);
        assert_eq!(value["cells"][0].as_array().unwrap().len(), 2);
    }

    #[test]
    fn calculation_limits_are_visible_on_the_wire() {
        let mut session = Session::open(&sample_xlsx(), None).unwrap();
        let result = session
            .edit_cell_json(
                r#"{"sheet":0,"row":2,"col":0,"input":"=SUM(Empty!A1:XFD1048576)"}"#,
                None,
            )
            .unwrap();
        assert!(result.contains(r#""limitedCells":["A3"]"#), "{result}");
        assert_eq!(
            session.calculation_status_json().unwrap(),
            r#"{"limitedCells":["A3"]}"#
        );
    }

    #[test]
    fn structural_ops_remap_and_undo_across_the_json_boundary() {
        let mut session = Session::open(&formula_xlsx(), None).unwrap();
        let result = session
            .apply_ops_json(
                r#"{"ops":[{"type":"insertRows","sheet":0,"at":0,"count":1}]}"#,
                None,
            )
            .unwrap();
        assert!(result.contains(r#""applied":true"#));
        assert!(
            session
                .cell_json(r#"{"sheet":0,"row":1,"col":1}"#)
                .unwrap()
                .contains(r#""input":"=SUM(A2:A3)""#)
        );

        session.undo_json(None).unwrap();
        assert!(
            session
                .cell_json(r#"{"sheet":0,"row":0,"col":1}"#)
                .unwrap()
                .contains(r#""input":"=SUM(A1:A2)""#)
        );
    }

    #[test]
    fn proposals_preserve_wire_behavior() {
        let mut session = Session::open(&sample_xlsx(), None).unwrap();
        let proposal = session
            .propose_json(
                r#"{"agentId":"agent","note":null,"edits":[{"sheet":0,"row":0,"col":0,"input":"new"}]}"#,
                None,
            )
            .unwrap();
        assert!(proposal.contains(r#""id":"p1""#));
        assert!(proposal.contains(r#""cells":["#));
        session
            .edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"moved"}"#, None)
            .unwrap();
        let error = session
            .accept_proposal_json(r#"{"id":"p1"}"#, None)
            .unwrap_err();
        assert!(error.starts_with("stale: A1"));
        assert!(
            session
                .accept_proposal_json(r#"{"id":"p1","force":true}"#, None)
                .unwrap()
                .contains(r#""proposalId":"p1""#)
        );
    }

    #[test]
    fn proposal_previews_remain_number_format_aware() {
        let mut session = Session::open(&currency_xlsx(), None).unwrap();
        let first = session
            .propose_json(
                r#"{"agentId":"agent","note":null,"edits":[{"sheet":0,"row":0,"col":0,"input":"2000"}]}"#,
                None,
            )
            .unwrap();
        assert!(first.contains(r#""id":"p1""#));
        assert!(first.contains(r#""oldText":"$1,000.00""#), "{first}");
        assert!(first.contains(r#""newText":"$2,000.00""#), "{first}");
        let second = session
            .propose_json(
                r#"{"agentId":"agent","note":null,"edits":[{"sheet":0,"row":0,"col":1,"input":"1"}]}"#,
                None,
            )
            .unwrap();
        assert!(second.contains(r#""id":"p2""#));
        session
            .edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"3000"}"#, None)
            .unwrap();
        let display = session
            .display_list_json(r#"{"x":0,"y":0,"width":200,"height":80}"#)
            .unwrap();
        assert!(display.contains("$3,000.00"), "{display}");
    }

    #[test]
    fn range_cap_and_bad_input_are_errors() {
        let session = Session::open(&sample_xlsx(), None).unwrap();
        assert!(session.display_list_json("not json").is_err());
        assert!(
            session
                .range_cells_json(r#"{"sheet":0,"range":"A1:D100000"}"#)
                .unwrap_err()
                .contains("cap")
        );
    }

    #[cfg(feature = "raster")]
    #[test]
    fn raster_methods_produce_png() {
        let session = Session::open(&sample_xlsx(), None).unwrap();
        let png = session
            .render_png(r#"{"x":0,"y":0,"width":240,"height":120}"#)
            .unwrap();
        assert_eq!(&png[..8], &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a]);
        let one = session.render_range_png(r#"{"range":"A1:B2"}"#).unwrap();
        let two = session
            .render_range_png(r#"{"range":"A1:B2","scale":2}"#)
            .unwrap();
        assert!(two.len() > one.len());
        assert!(
            session
                .render_range_png(r#"{"range":"A1:B2","scale":0}"#)
                .is_err()
        );
        assert!(session.render_range_png(r#"{"range":"nope"}"#).is_err());
    }
}
