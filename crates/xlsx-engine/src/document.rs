use std::collections::VecDeque;
use std::sync::{Arc, Mutex};

use betteroffice_xlsx::{
    MAX_COLLABORATION_BYTES, MAX_COLLABORATION_CLIENT_ID, SnapshotBudget, UpdateEvent,
    UpdateOrigin, UpdateSubscription, WorkbookSnapshotBuilder,
};

use crate::EngineError;
use crate::session::Session;

const MAX_QUEUED_UPDATE_EVENTS: usize = 4_097;
const MAX_QUEUED_UPDATE_BYTES: usize = MAX_COLLABORATION_BYTES * 2 + MAX_QUEUED_UPDATE_EVENTS;

#[derive(Clone, Copy)]
pub struct Host {
    pub now_serial: fn() -> Result<Option<f64>, EngineError>,
    pub timer_ms: fn() -> f64,
    pub progress: fn(),
    pub version: &'static str,
}

fn message(error: String) -> EngineError {
    EngineError::message(error)
}

fn error_kind(error: String) -> EngineError {
    EngineError::error(error)
}

fn type_error(error: String) -> EngineError {
    EngineError::type_error(error)
}

pub fn parse_recovery_sequence(value: f64) -> Result<u64, EngineError> {
    if !value.is_finite()
        || value.fract() != 0.0
        || !(0.0..=9_007_199_254_740_991.0).contains(&value)
    {
        return Err(EngineError::type_error("Invalid recovery sequence"));
    }
    Ok(value as u64)
}

pub fn parse_client_id(client_id: f64) -> Result<u64, EngineError> {
    if !client_id.is_finite()
        || client_id.fract() != 0.0
        || client_id < 1.0
        || client_id > MAX_COLLABORATION_CLIENT_ID as f64
    {
        return Err(EngineError::message(
            "client ID must be a nonzero integer no greater than Number.MAX_SAFE_INTEGER",
        ));
    }
    Ok(client_id as u64)
}

pub fn display_list_cell_limit() -> f64 {
    betteroffice_xlsx::MAX_DISPLAY_CELLS as f64
}

struct UpdateObserver {
    pending: Arc<Mutex<PendingUpdateEvents>>,
    _subscription: UpdateSubscription,
}

#[derive(Default)]
struct PendingUpdateEvents {
    events: VecDeque<UpdateEvent>,
    bytes: usize,
    overflowed: bool,
}

pub struct Document {
    session: Session,
    update_observer: Option<UpdateObserver>,
    host: Host,
}

impl Document {
    fn opened(host: Host, open: impl FnOnce() -> Result<Session, String>) -> Result<Self, String> {
        betteroffice_xlsx::set_progress_hook(Some(host.progress));
        open().map(|session| Self {
            session,
            update_observer: None,
            host,
        })
    }

    pub fn host(&self) -> Host {
        self.host
    }

    pub fn recovery_checkpoint(
        &self,
        through_seq: f64,
        preserve_history: Option<bool>,
    ) -> Result<Vec<u8>, EngineError> {
        self.session
            .recovery_checkpoint(
                parse_recovery_sequence(through_seq)?,
                preserve_history.unwrap_or(false),
            )
            .map_err(error_kind)
    }

    pub fn open_recovery_checkpoint(
        bytes: &[u8],
        through_seq: f64,
        host: Host,
    ) -> Result<Self, EngineError> {
        let through_seq = parse_recovery_sequence(through_seq)?;
        Self::opened(host, || {
            Session::open_recovery_checkpoint(bytes, through_seq)
        })
        .map_err(error_kind)
    }

    pub fn begin_peer_snapshot(&mut self, records: usize, bytes: usize) -> Result<(), EngineError> {
        self.session
            .begin_peer_snapshot(records, bytes)
            .map_err(error_kind)
    }

    pub fn next_peer_snapshot_chunk(&mut self) -> Result<Option<Vec<u8>>, EngineError> {
        self.session.next_peer_snapshot_chunk().map_err(error_kind)
    }

    pub fn end_peer_snapshot(&mut self) {
        self.session.end_peer_snapshot();
    }

    pub fn adopt_peer_version(&mut self, version: &str) -> Result<(), EngineError> {
        self.session.adopt_peer_version(version).map_err(message)
    }

    pub fn peer_hydration_json(&self) -> Result<String, EngineError> {
        self.session.peer_hydration_json().map_err(message)
    }

    pub fn open_with_peer_hydration_json(
        bytes: &[u8],
        hydration: &str,
        host: Host,
    ) -> Result<Self, EngineError> {
        Self::opened(host, || {
            Session::open_with_peer_hydration_json(bytes, hydration)
        })
        .map_err(message)
    }

    pub fn open(bytes: &[u8], host: Host) -> Result<Self, EngineError> {
        let now_serial = (host.now_serial)()?;
        Self::opened(host, || Session::open(bytes, now_serial)).map_err(message)
    }

    pub fn open_with_calculation_json(
        bytes: &[u8],
        context: &str,
        host: Host,
    ) -> Result<Self, EngineError> {
        Self::opened(host, || Session::open_with_calculation_json(bytes, context))
            .map_err(type_error)
    }

    pub fn set_calculation_context_json(&mut self, context: &str) -> Result<(), EngineError> {
        self.session
            .set_calculation_context_json(context)
            .map_err(type_error)
    }

    pub fn open_collaborative(
        bytes: &[u8],
        client_id: f64,
        host: Host,
    ) -> Result<Self, EngineError> {
        let client_id = parse_client_id(client_id)?;
        let now_serial = (host.now_serial)()?;
        Self::opened(host, || {
            Session::open_collaborative(bytes, client_id, now_serial)
        })
        .map_err(message)
    }

    pub fn client_id(&self) -> f64 {
        self.session.client_id() as f64
    }

    pub fn encode_state_vector(&self) -> Vec<u8> {
        self.session.encode_state_vector()
    }

    pub fn encode_state_as_update(&self) -> Vec<u8> {
        self.session.encode_state_as_update()
    }

    pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Result<Vec<u8>, EngineError> {
        self.session
            .encode_diff(remote_state_vector)
            .map_err(message)
    }

    pub fn apply_update_json(&mut self, update: &[u8]) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .apply_update_json(update, now_serial)
            .map_err(message)
    }

    pub fn start_update_observation(&mut self) -> Result<(), EngineError> {
        if self.update_observer.is_some() {
            return Ok(());
        }
        let pending = Arc::new(Mutex::new(PendingUpdateEvents::default()));
        let observed = Arc::clone(&pending);
        let subscription = self
            .session
            .observe_update_v1(move |event| {
                if event.origin == UpdateOrigin::Recalculation {
                    return;
                }
                let mut observed = observed
                    .lock()
                    .unwrap_or_else(|poisoned| poisoned.into_inner());
                if observed.overflowed {
                    return;
                }
                let bytes = event.update.len().saturating_add(1);
                if observed.events.len() >= MAX_QUEUED_UPDATE_EVENTS
                    || bytes > MAX_QUEUED_UPDATE_BYTES.saturating_sub(observed.bytes)
                {
                    observed.events.clear();
                    observed.bytes = 0;
                    observed.overflowed = true;
                    return;
                }
                observed.bytes += bytes;
                observed.events.push_back(event);
            })
            .map_err(message)?;
        self.update_observer = Some(UpdateObserver {
            pending,
            _subscription: subscription,
        });
        Ok(())
    }

    pub fn clear_update_observation(&mut self) {
        self.update_observer = None;
    }

    pub fn drain_update_event(&self) -> Result<Vec<u8>, EngineError> {
        let Some(observer) = &self.update_observer else {
            return Ok(Vec::new());
        };
        let mut pending = observer
            .pending
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if pending.overflowed {
            pending.overflowed = false;
            return Err(EngineError::message(
                "xlsx update observation queue exceeded its limit",
            ));
        }
        let event = pending.events.pop_front();
        let Some(event) = event else {
            return Ok(Vec::new());
        };
        pending.bytes = pending
            .bytes
            .saturating_sub(event.update.len().saturating_add(1));
        let mut encoded = Vec::with_capacity(event.update.len() + 1);
        encoded.push(match event.origin {
            UpdateOrigin::Local | UpdateOrigin::Recalculation => 0,
            UpdateOrigin::Remote => 1,
        });
        encoded.extend_from_slice(&event.update);
        Ok(encoded)
    }

    pub fn print_display_list_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.print_display_list_json(args).map_err(message)
    }

    pub fn display_list_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.display_list_json(args).map_err(message)
    }

    pub fn display_list_for_json(
        &self,
        sheet: u32,
        viewport_json: &str,
    ) -> Result<String, EngineError> {
        self.session
            .display_list_for_json(sheet, viewport_json)
            .map_err(message)
    }

    pub fn display_list_profiled_json(&self, viewport_json: &str) -> Result<String, EngineError> {
        let timer_fn = self.host.timer_ms;
        let mut timer = move || timer_fn();
        self.session
            .display_list_profiled_json(viewport_json, &mut timer)
            .map_err(message)
    }

    pub fn chart_at_point_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.chart_at_point_json(args).map_err(message)
    }

    pub fn move_chart_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .move_chart_json(args, now_serial)
            .map_err(message)
    }

    #[cfg(feature = "raster")]
    pub fn render_png(&self, viewport_json: &str) -> Result<Vec<u8>, EngineError> {
        self.session.render_png(viewport_json).map_err(message)
    }

    #[cfg(feature = "raster")]
    pub fn render_range_png(&self, args: &str) -> Result<Vec<u8>, EngineError> {
        self.session.render_range_png(args).map_err(message)
    }

    pub fn sheet_info_json(&self) -> Result<String, EngineError> {
        self.session.sheet_info_json().map_err(message)
    }

    pub fn sheet_count(&self) -> usize {
        self.session.sheet_count()
    }

    pub fn sheet_info_for_json(&self, sheet: u32) -> Result<String, EngineError> {
        self.session.sheet_info_for_json(sheet).map_err(message)
    }

    pub fn selection_data_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.selection_data_json(args).map_err(message)
    }

    pub fn cell_rect_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.cell_rect_json(args).map_err(message)
    }

    pub fn calculation_status_json(&self) -> Result<String, EngineError> {
        self.session.calculation_status_json().map_err(message)
    }

    pub fn set_active_sheet(&mut self, index: u32) -> Result<(), EngineError> {
        self.session.set_active_sheet(index).map_err(message)
    }

    pub fn edit_cell_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .edit_cell_json(args, now_serial)
            .map_err(message)
    }

    pub fn edit_cell_profiled_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        let timer_fn = self.host.timer_ms;
        let mut timer = move || timer_fn();
        self.session
            .edit_cell_profiled_json(args, now_serial, &mut timer)
            .map_err(message)
    }

    pub fn edit_cells_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .edit_cells_json(args, now_serial)
            .map_err(message)
    }

    pub fn move_range_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .move_range_json(args, now_serial)
            .map_err(message)
    }

    pub fn apply_ops_json(&mut self, transaction_json: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .apply_ops_json(transaction_json, now_serial)
            .map_err(message)
    }

    pub fn charts_json(&self, sheet: u32) -> Result<String, EngineError> {
        self.session.charts_json(sheet).map_err(message)
    }

    pub fn edit_charts_json(&mut self, edits: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .edit_charts_json(edits, now_serial)
            .map_err(message)
    }

    pub fn conditional_formats_json(&self, sheet: u32) -> Result<String, EngineError> {
        self.session
            .conditional_formats_json(sheet)
            .map_err(message)
    }

    pub fn apply_ops_profiled_json(
        &mut self,
        transaction_json: &str,
    ) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        let timer_fn = self.host.timer_ms;
        let mut timer = move || timer_fn();
        self.session
            .apply_ops_profiled_json(transaction_json, now_serial, &mut timer)
            .map_err(message)
    }

    pub fn undo_json(&mut self) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session.undo_json(now_serial).map_err(message)
    }

    pub fn redo_json(&mut self) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session.redo_json(now_serial).map_err(message)
    }

    pub fn cell_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.cell_json(args).map_err(message)
    }

    pub fn search_text_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.search_text_json(args).map_err(message)
    }

    pub fn cell_position_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.cell_position_json(args).map_err(message)
    }

    pub fn range_cells_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.range_cells_json(args).map_err(message)
    }

    pub fn selection_formatting_json(&self, args: &str) -> Result<String, EngineError> {
        self.session
            .selection_formatting_json(args)
            .map_err(message)
    }

    pub fn capture_format_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.capture_format_json(args).map_err(message)
    }

    pub fn merged_ranges_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.merged_ranges_json(args).map_err(message)
    }

    pub fn history_state_json(&self) -> Result<String, EngineError> {
        self.session.history_state_json().map_err(message)
    }

    pub fn list_proposals_json(&self) -> Result<String, EngineError> {
        self.session.list_proposals_json().map_err(message)
    }

    pub fn read_cells_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.read_cells_json(args).map_err(message)
    }

    pub fn list_defined_names_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.list_defined_names_json(args).map_err(message)
    }

    pub fn find_text_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.find_text_json(args).map_err(message)
    }

    pub fn validate_edits_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.validate_edits_json(args).map_err(message)
    }

    pub fn export_structured_json(&self, args: &str) -> Result<String, EngineError> {
        self.session.export_structured_json(args).map_err(message)
    }

    pub fn patch_range_style_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .patch_range_style_json(args, now_serial)
            .map_err(message)
    }

    pub fn set_range_number_format_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .set_range_number_format_json(args, now_serial)
            .map_err(message)
    }

    pub fn apply_format_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .apply_format_json(args, now_serial)
            .map_err(message)
    }

    pub fn visible_merged_ranges_json(
        &self,
        sheet: u32,
        viewport_json: &str,
    ) -> Result<String, EngineError> {
        self.session
            .visible_merged_ranges_json(sheet, viewport_json)
            .map_err(message)
    }

    pub fn propose_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session.propose_json(args, now_serial).map_err(message)
    }

    pub fn accept_proposal_json(&mut self, args: &str) -> Result<String, EngineError> {
        let now_serial = (self.host.now_serial)()?;
        self.session
            .accept_proposal_json(args, now_serial)
            .map_err(message)
    }

    pub fn reject_proposal_json(&mut self, args: &str) -> Result<String, EngineError> {
        self.session.reject_proposal_json(args).map_err(message)
    }

    pub fn document_version(&self) -> String {
        self.session.document_version()
    }

    pub fn apply_edits_json(&mut self, request: &str) -> Result<String, EngineError> {
        self.session.apply_edits_json(request).map_err(message)
    }

    pub fn export_markdown_json(
        &self,
        options: &str,
        markdown_options: &str,
    ) -> Result<String, EngineError> {
        self.session
            .export_markdown_json(options, markdown_options)
            .map_err(message)
    }

    pub fn save_bytes(&self) -> Result<Vec<u8>, EngineError> {
        self.session.save().map_err(message)
    }

    pub fn version(&self) -> &'static str {
        self.host.version
    }
}

pub struct SnapshotBuilder {
    builder: WorkbookSnapshotBuilder,
    host: Host,
}

impl SnapshotBuilder {
    pub fn new(host: Host) -> Self {
        Self {
            builder: WorkbookSnapshotBuilder::new(),
            host,
        }
    }

    pub fn push(&mut self, chunk: &[u8]) -> Result<(), EngineError> {
        self.builder
            .push(chunk)
            .map(|_| ())
            .map_err(|error| EngineError::error(error.to_string()))
    }

    pub fn advance(&mut self, records: usize, bytes: usize) -> Result<bool, EngineError> {
        let budget = SnapshotBudget::new(records, bytes)
            .map_err(|error| EngineError::error(error.to_string()))?;
        self.builder
            .advance(budget)
            .map(|progress| progress.is_ready())
            .map_err(|error| EngineError::error(error.to_string()))
    }

    pub fn finish(self) -> Result<Document, EngineError> {
        let host = self.host;
        Document::opened(host, || {
            self.builder
                .finish()
                .map(Session::from_hydrated_workbook)
                .map_err(|error| error.to_string())
        })
        .map_err(EngineError::error)
    }
}

impl Default for SnapshotBuilder {
    fn default() -> Self {
        fn none() -> Result<Option<f64>, EngineError> {
            Ok(None)
        }
        fn zero() -> f64 {
            0.0
        }
        fn noop() {}
        Self::new(Host {
            now_serial: none,
            timer_ms: zero,
            progress: noop,
            version: "",
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ErrorKind;

    fn host() -> Host {
        fn none() -> Result<Option<f64>, EngineError> {
            Ok(None)
        }
        fn zero() -> f64 {
            0.0
        }
        fn noop() {}
        Host {
            now_serial: none,
            timer_ms: zero,
            progress: noop,
            version: "0.0.0",
        }
    }

    #[test]
    fn bad_client_id_uses_message_kind() {
        let error = parse_client_id(0.0).unwrap_err();
        assert_eq!(error.kind, ErrorKind::Message);
        assert_eq!(
            error.message,
            "client ID must be a nonzero integer no greater than Number.MAX_SAFE_INTEGER"
        );
        let error = Document::open_collaborative(&[], 0.0, host())
            .err()
            .expect("bad client id must fail");
        assert_eq!(error.kind, ErrorKind::Message);
    }

    #[test]
    fn bad_recovery_sequence_uses_type_error_kind() {
        let error = parse_recovery_sequence(f64::NAN).unwrap_err();
        assert_eq!(error.kind, ErrorKind::TypeError);
        assert_eq!(error.message, "Invalid recovery sequence");
    }

    #[test]
    fn bad_calculation_json_uses_type_error_kind() {
        let error = Document::open_with_calculation_json(&[], "not json", host())
            .err()
            .expect("bad calculation json must fail");
        assert_eq!(error.kind, ErrorKind::TypeError);
    }

    #[test]
    fn bad_json_args_use_message_kind() {
        let mut sheet = xlsx_model::Sheet::new("Data");
        sheet.set_cell(
            betteroffice_xlsx::CellRef::parse_a1("A1").unwrap(),
            xlsx_model::Cell {
                value: xlsx_model::CellValue::Text {
                    value: "Hello".into(),
                },
                ..xlsx_model::Cell::default()
            },
        );
        let mut model = xlsx_model::Workbook::default();
        model.sheets.push(sheet);
        let parts = xlsx_parse::serialize_workbook(&model).unwrap();
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let mut document = Document::open(&bytes, host()).unwrap();
        let error = document.edit_cell_json("not json").unwrap_err();
        assert_eq!(error.kind, ErrorKind::Message);
        assert!(error.message.contains("bad edit args"));
    }

    #[test]
    fn clock_failure_precedes_open() {
        let mut failing_host = host();
        failing_host.now_serial = || Err(EngineError::pending_exception());
        let error = Document::open(&[], failing_host)
            .err()
            .expect("clock failure must abort open");
        assert_eq!(error.kind, ErrorKind::PendingException);
    }

    #[test]
    fn clock_failure_leaves_edits_untouched() {
        let bytes = include_bytes!("../../ooxml-opc/tests/fixtures/sample.xlsx");
        let mut document = Document::open(bytes, host()).unwrap();
        let request = r#"{"ranges":[{"sheetId":"sheet:0","range":{"kind":"a1","a1":"A1"}}]}"#;
        let before = document.read_cells_json(request).unwrap();
        let version = document.document_version();
        document.host.now_serial = || Err(EngineError::pending_exception());
        let error = document
            .edit_cell_json(r#"{"sheet":0,"row":0,"col":0,"input":"changed"}"#)
            .unwrap_err();
        assert_eq!(error.kind, ErrorKind::PendingException);
        assert_eq!(document.read_cells_json(request).unwrap(), before);
        assert_eq!(document.document_version(), version);
    }

    #[test]
    fn update_queue_covers_the_largest_pending_resolution_burst() {
        assert_eq!(
            MAX_QUEUED_UPDATE_BYTES,
            MAX_COLLABORATION_BYTES * 2 + MAX_QUEUED_UPDATE_EVENTS
        );
    }
}
