use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use crate::sheet_json::{decode_charts, decode_hyperlinks};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use xlsx_model::{
    AnchorEditAs, AnchorExtent, AnchorPos, Cell, CellFormat, CellRange, CellRef, CellValue,
    ChartAnchor, ChartRef, ColStyle, ConditionalFormatting, DateSystem, DefinedName, ErrorValue,
    FreezePane, Hyperlink, MAX_COLS, MAX_ROWS, Sheet, SheetChart, SheetFormat, SheetId, Stylesheet,
    Table, Workbook as WorkbookModel,
};
use xlsx_ops::{CellState, Op, PreservedCellFormat as StoredCellFormat};
use yrs::block::{
    BLOCK_GC_REF_NUMBER, BLOCK_ITEM_ANY_REF_NUMBER, BLOCK_ITEM_DELETED_REF_NUMBER,
    BLOCK_ITEM_TYPE_REF_NUMBER, BLOCK_SKIP_REF_NUMBER, ClientID,
};
use yrs::encoding::read::{Error as DecodeError, Read};
use yrs::encoding::write::Write;
use yrs::sync::time::Clock;
use yrs::types::{TYPE_REFS_ARRAY, TYPE_REFS_MAP};
use yrs::undo::{Options as UndoOptions, StackItem, UndoManager};
use yrs::updates::decoder::{Decode, Decoder, DecoderV1};
use yrs::updates::encoder::{Encode, Encoder, EncoderV1};
use yrs::{
    Any, Array, ArrayRef, BranchID, Doc, ID, Map, MapPrelim, MapRef, Options, Origin, Out, ReadTxn,
    StateVector, Transact, TransactionMut, Update, WriteTxn,
};

#[cfg(test)]
mod cascade_composition_tests;
#[cfg(test)]
mod fingerprint_composition_tests;
mod hyperlinks;
mod merges;
#[cfg_attr(not(test), allow(dead_code))]
pub(crate) mod snapshot;
pub(crate) mod snapshot_validation;
pub(crate) use hyperlinks::validate_hyperlink_edit_baseline;
use hyperlinks::{
    HYPERLINK_PREFIX, HyperlinkBaseline, HyperlinkEdits, HyperlinkProjection,
    clear_hyperlink_edits, insert_hyperlink_edit, project_hyperlinks, sync_hyperlinks,
};
use merges::{
    MERGE_PREFIX, MergeEdits, MergeProjection, clear_merge_edits, clear_sheet_edits,
    insert_merge_edit, project_merges,
};

const META: &str = "xlsx";
const CELL_FORMATS: &str = "xlsx:cell-formats";
const SHEET_ORDER: &str = "xlsx:sheet-order";
const SHEETS: &str = "xlsx:sheets";
const MIN_SUPPORTED_SCHEMA_VERSION: i64 = 3;
/// The schema each feature first appeared in. These are frozen: gating a
/// feature on the current version instead would silently reclassify the
/// newest schema as predating it the moment the current version moves on.
const FREEZE_PANE_SCHEMA_VERSION: i64 = 4;
const HYPERLINK_SCHEMA_VERSION: i64 = 5;
const CHARTS_SCHEMA_VERSION: i64 = 6;
const SCHEMA_VERSION: i64 = 6;
const BASE_FINGERPRINT: &str = "baseFingerprint";
const STRUCTURE_GENERATION: &str = "structureGeneration";
const CHARTS: &str = "charts";
const CONTENTS: &str = "contents";
const COL_WIDTHS: &str = "colWidths";
const COL_STYLES: &str = "colStyles";
const ROW_STYLES: &str = "rowStyles";
const FREEZE_PANE: &str = "freezePane";
const HYPERLINKS: &str = "hyperlinks";
const CONDITIONAL_FORMATS: &str = "conditionalFormats";
const AUTO_FILTER: &str = "autoFilter";
const MERGES: &str = "merges";
const NAME: &str = "name";
const ROW_HEIGHTS: &str = "rowHeights";
const STYLES: &str = "styles";
const BOOTSTRAP_ORIGIN: &str = "xlsx:bootstrap";
const HYDRATE_ORIGIN: &str = "xlsx:hydrate";
const REMOTE_ORIGIN: &str = "xlsx:remote";
/// Local adoptions no undo manager captures.
const UNTRACKED_ORIGIN: &str = "xlsx:untracked";
const MAX_SAFE_CLIENT_ID: u64 = (1_u64 << 53) - 1;
const MAX_SAFE_CLOCK: u32 = i32::MAX as u32;
const MAX_UPDATE_BLOCKS: usize = 1_000_000;
const MAX_UPDATE_VALUES: usize = 1_000_000;
const MAX_UPDATE_DELETE_RANGES: usize = 1_000_000;
/// upper bound on one encoded cell format. the canonical form of any format a
/// workbook can hold is three orders of magnitude smaller.
const MAX_CELL_FORMAT_BYTES: usize = 64 * 1024;
const UNDO_CAPTURE_TIMEOUT_MS: u64 = 500;
pub(crate) const MAX_STATE_VECTOR_ENTRIES: u32 = 65_536;

#[cfg(test)]
thread_local! {
    static FORCE_FULL_MATERIALIZATION: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    static FAST_SET_CELL_COUNT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static FINGERPRINT_PASSES: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// Full passes over every cell made to fingerprint a workbook base.
#[cfg(test)]
pub(crate) fn fingerprint_passes() -> usize {
    FINGERPRINT_PASSES.get()
}

#[cfg(test)]
pub(crate) fn force_full_materialization() -> bool {
    FORCE_FULL_MATERIALIZATION.get()
}

#[cfg(test)]
pub(crate) fn with_full_materialization<T>(f: impl FnOnce() -> T) -> T {
    struct Reset(bool);

    impl Drop for Reset {
        fn drop(&mut self) {
            FORCE_FULL_MATERIALIZATION.set(self.0);
        }
    }

    let _reset = Reset(FORCE_FULL_MATERIALIZATION.replace(true));
    f()
}

#[cfg(test)]
pub(crate) fn fast_set_cell_count() -> usize {
    FAST_SET_CELL_COUNT.get()
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SyncOrigin {
    User,
    Agent,
    Undo,
    Redo,
}

impl SyncOrigin {
    fn as_str(self) -> &'static str {
        match self {
            Self::User => "xlsx:user",
            Self::Agent => "xlsx:agent",
            Self::Undo => "xlsx:undo",
            Self::Redo => "xlsx:redo",
        }
    }
}

#[derive(Debug)]
pub(crate) enum AuthorityError {
    ClientIdConflict(u64),
    InvalidStateVector(String),
    InvalidUpdate(String),
    InvalidState(String),
}

#[derive(Clone)]
struct WorkbookBase {
    bootstrap_client_id: u64,
    date_system: DateSystem,
    defined_names: Vec<DefinedName>,
    fingerprint: String,
    fingerprints: BTreeMap<i64, Vec<String>>,
    freeze_panes: Vec<Option<FreezePane>>,
    formats: Vec<SheetFormat>,
    col_styles: Vec<Vec<ColStyle>>,
    row_styles: Vec<BTreeMap<u32, u32>>,
    hyperlinks: Vec<Vec<Hyperlink>>,
    conditional_formats: Vec<Vec<ConditionalFormatting>>,
    auto_filters: Vec<Option<xlsx_model::AutoFilter>>,
    charts: Vec<Vec<SheetChart>>,
    hidden_dimensions: Vec<HiddenDimensions>,
    shared_strings: Vec<String>,
    styles: Stylesheet,
    /// table parts; read-only reference data, not shared state.
    tables: Vec<Table>,
}

impl WorkbookBase {
    /// A legacy fingerprint hashes no chart state, so a charted workbook pairs
    /// with one on its other content alone and keeps the charts it parsed. It
    /// is the only reading of a state written before charts were shared, and
    /// refusing it would strand every snapshot an earlier release persisted.
    #[cfg(test)]
    fn from_model(model: &WorkbookModel) -> Result<Self, String> {
        Self::from_model_with_legacy_dimensions(model, &[], None)
    }

    /// `legacy_dimensions` are the row heights and column widths releases
    /// before hidden rows and columns read as zero stored, per sheet. Those
    /// maps are hashed at every schema, so a peer that persisted its state
    /// under one of those releases is only recognisable against them.
    ///
    /// The current fingerprint hashes the format facets under its own domain,
    /// where no reading drops one. Releases before facets wrote the others,
    /// over their reading of the styles.
    #[cfg(test)]
    fn from_model_with_legacy_dimensions(
        model: &WorkbookModel,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
    ) -> Result<Self, String> {
        Self::from_model_with_style_readings(model, legacy_dimensions, legacy_styles, &[])
    }

    fn from_model_with_style_readings(
        model: &WorkbookModel,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
        legacy_style_readings: &[Stylesheet],
    ) -> Result<Self, String> {
        let mut builder = FingerprintBuilder::new(model, legacy_dimensions);
        let ((fingerprint, bootstrap_client_id), fingerprints) = builder.batched(|builder| {
            let styles = FingerprintStyles::new(&model.styles);
            let current = builder.current(if released_styles(&model.styles) {
                styles
            } else {
                styles.with_written_fills()
            })?;
            let mut fingerprints = builder.style_fingerprints(&model.styles, legacy_styles)?;
            for reading in legacy_style_readings {
                for (version, accepted) in builder.style_fingerprints(reading, None)? {
                    for fingerprint in accepted {
                        append_fingerprint(&mut fingerprints, version, fingerprint);
                    }
                }
            }
            Ok((current, fingerprints))
        })?;
        Ok(Self {
            bootstrap_client_id,
            date_system: model.date_system,
            defined_names: model.defined_names.clone(),
            fingerprint,
            fingerprints,
            freeze_panes: model.sheets.iter().map(|sheet| sheet.freeze_pane).collect(),
            formats: model.sheets.iter().map(|sheet| sheet.format).collect(),
            col_styles: model
                .sheets
                .iter()
                .map(|sheet| sheet.col_styles.clone())
                .collect(),
            row_styles: base_row_styles(model),
            conditional_formats: model
                .sheets
                .iter()
                .map(|sheet| sheet.conditional_formats.clone())
                .collect(),
            auto_filters: model
                .sheets
                .iter()
                .map(|sheet| sheet.auto_filter.clone())
                .collect(),
            hyperlinks: model
                .sheets
                .iter()
                .map(|sheet| sheet.hyperlinks.clone())
                .collect(),
            charts: model
                .sheets
                .iter()
                .map(|sheet| sheet.charts.clone())
                .collect(),
            hidden_dimensions: hidden_dimensions(model, legacy_dimensions),
            shared_strings: model.shared_strings.clone(),
            styles: model.styles.clone(),
            tables: model.tables.clone(),
        })
    }

    #[cfg(test)]
    fn from_model_with_legacy_dimensions_oracle(
        model: &WorkbookModel,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
    ) -> Result<Self, String> {
        let mut released = model.clone();
        released.styles = pre_facet_styles(&model.styles);
        let released_legacy = legacy_styles.map(pre_facet_styles);
        let mut base =
            Self::released_base_oracle(&released, legacy_dimensions, released_legacy.as_ref())?;
        let (fingerprint, bootstrap_client_id) = fingerprint_model_current_oracle(model)?;
        let mut accepted = Self::current_fingerprints_oracle(model, legacy_styles)?;
        if !released_styles(&model.styles) {
            for reading in [FingerprintReading::Facets, FingerprintReading::Patterns] {
                for fingerprint in Self::reading_fingerprints_oracle(model, legacy_styles, reading)?
                {
                    if !accepted.contains(&fingerprint) {
                        accepted.push(fingerprint);
                    }
                }
            }
        }
        for released in base
            .fingerprints
            .remove(&SCHEMA_VERSION)
            .unwrap_or_default()
        {
            if !accepted.contains(&released) {
                accepted.push(released);
            }
        }
        base.fingerprints.insert(SCHEMA_VERSION, accepted);
        base.fingerprint = fingerprint;
        base.bootstrap_client_id = bootstrap_client_id;
        base.styles = model.styles.clone();
        Ok(base)
    }

    #[cfg(test)]
    fn current_fingerprints_oracle(
        model: &WorkbookModel,
        legacy_styles: Option<&Stylesheet>,
    ) -> Result<Vec<String>, String> {
        Self::reading_fingerprints_oracle(
            model,
            legacy_styles,
            FingerprintReading::of(true, !released_styles(&model.styles)),
        )
    }

    #[cfg(test)]
    fn reading_fingerprints_oracle(
        model: &WorkbookModel,
        legacy_styles: Option<&Stylesheet>,
        reading: FingerprintReading,
    ) -> Result<Vec<String>, String> {
        let mut accepted = Vec::new();
        for (include_col_styles, include_row_styles) in
            [(true, true), (false, false), (true, false)]
        {
            let defaults = include_col_styles && include_row_styles && has_style_defaults(model);
            let current = reading.with_defaults(defaults);
            for reading in std::iter::once(current).chain(current.legacy()) {
                let (fingerprint, _) = fingerprint_model_with_row_domain_oracle(
                    model,
                    SCHEMA_VERSION,
                    true,
                    include_col_styles,
                    include_row_styles,
                    fingerprint_domain(SCHEMA_VERSION, reading),
                )?;
                accepted.push(fingerprint);
            }
        }
        if let Some(styles) = legacy_styles.filter(|styles| **styles != model.styles) {
            let mut legacy_model = model.clone();
            legacy_model.styles = styles.clone();
            accepted.extend(Self::reading_fingerprints_oracle(
                &legacy_model,
                None,
                reading,
            )?);
        }
        if !model.styles.indexed_colors.is_empty() {
            let mut legacy_model = model.clone();
            legacy_model.styles.indexed_colors.clear();
            let colorless =
                Self::reading_fingerprints_oracle(&legacy_model, legacy_styles, reading)?;
            accepted.extend(colorless);
        }
        let mut unique = Vec::new();
        for fingerprint in accepted {
            if !unique.contains(&fingerprint) {
                unique.push(fingerprint);
            }
        }
        Ok(unique)
    }

    /// The base a release before format facets built, over its reading of
    /// the styles.
    #[cfg(test)]
    fn released_base_oracle(
        model: &WorkbookModel,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
    ) -> Result<Self, String> {
        let (fingerprint, bootstrap_client_id) =
            fingerprint_model_for_schema_oracle(model, SCHEMA_VERSION)?;
        let mut fingerprints = BTreeMap::new();
        for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
            let (version_fingerprint, _) = fingerprint_model_for_schema_oracle(model, version)?;
            fingerprints.insert(version, vec![version_fingerprint]);
        }
        if let Some(version_3) = fingerprints.get_mut(&3) {
            let (defined_names_v3, _) = fingerprint_model_with_schema_oracle(model, 3, true, true)?;
            if !version_3.contains(&defined_names_v3) {
                version_3.push(defined_names_v3);
            }
        }
        for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
            let (without_col_styles, _) =
                fingerprint_model_with_schema_oracle(model, version, version >= 4, false)?;
            let (without_row_styles, _) = fingerprint_model_with_row_styles_oracle(
                model,
                version,
                version >= 4,
                true,
                false,
            )?;
            let accepted = fingerprints.entry(version).or_default();
            for fingerprint in [without_col_styles, without_row_styles] {
                if !accepted.contains(&fingerprint) {
                    accepted.push(fingerprint);
                }
            }
        }
        if let Some(legacy) = model_with_legacy_dimensions(model, legacy_dimensions) {
            for version in MIN_SUPPORTED_SCHEMA_VERSION..SCHEMA_VERSION {
                let (legacy_fingerprint, _) =
                    fingerprint_model_for_schema_oracle(&legacy, version)?;
                let (without_col_styles, _) =
                    fingerprint_model_with_schema_oracle(&legacy, version, version >= 4, false)?;
                let accepted = fingerprints.entry(version).or_default();
                for fingerprint in [legacy_fingerprint, without_col_styles] {
                    if !accepted.contains(&fingerprint) {
                        accepted.push(fingerprint);
                    }
                }
            }
            if let Some(version_3) = fingerprints.get_mut(&3) {
                let (defined_names_v3, _) =
                    fingerprint_model_with_schema_oracle(&legacy, 3, true, true)?;
                if !version_3.contains(&defined_names_v3) {
                    version_3.push(defined_names_v3);
                }
            }
        }
        if let Some(styles) = legacy_styles.filter(|styles| **styles != model.styles) {
            let mut legacy_model = model.clone();
            legacy_model.styles = styles.clone();
            let legacy_base = Self::released_base_oracle(&legacy_model, legacy_dimensions, None)?;
            for (version, legacy_fingerprints) in legacy_base.fingerprints {
                let accepted = fingerprints.entry(version).or_default();
                for fingerprint in legacy_fingerprints {
                    if !accepted.contains(&fingerprint) {
                        accepted.push(fingerprint);
                    }
                }
            }
        }
        if !model.styles.indexed_colors.is_empty() {
            let mut legacy_model = model.clone();
            legacy_model.styles.indexed_colors.clear();
            let legacy_base =
                Self::released_base_oracle(&legacy_model, legacy_dimensions, legacy_styles)?;
            for (version, legacy_fingerprints) in legacy_base.fingerprints {
                let accepted = fingerprints.entry(version).or_default();
                for fingerprint in legacy_fingerprints {
                    if !accepted.contains(&fingerprint) {
                        accepted.push(fingerprint);
                    }
                }
            }
        }
        Ok(Self {
            bootstrap_client_id,
            date_system: model.date_system,
            defined_names: model.defined_names.clone(),
            fingerprint,
            fingerprints,
            freeze_panes: model.sheets.iter().map(|sheet| sheet.freeze_pane).collect(),
            formats: model.sheets.iter().map(|sheet| sheet.format).collect(),
            col_styles: model
                .sheets
                .iter()
                .map(|sheet| sheet.col_styles.clone())
                .collect(),
            row_styles: base_row_styles(model),
            conditional_formats: model
                .sheets
                .iter()
                .map(|sheet| sheet.conditional_formats.clone())
                .collect(),
            auto_filters: model
                .sheets
                .iter()
                .map(|sheet| sheet.auto_filter.clone())
                .collect(),
            hyperlinks: model
                .sheets
                .iter()
                .map(|sheet| sheet.hyperlinks.clone())
                .collect(),
            charts: model
                .sheets
                .iter()
                .map(|sheet| sheet.charts.clone())
                .collect(),
            hidden_dimensions: hidden_dimensions(model, legacy_dimensions),
            shared_strings: model.shared_strings.clone(),
            styles: model.styles.clone(),
            tables: model.tables.clone(),
        })
    }

    fn accepts_fingerprint(&self, version: i64, fingerprint: &str) -> bool {
        self.fingerprints
            .get(&version)
            .is_some_and(|accepted| accepted.iter().any(|value| value == fingerprint))
    }

    fn workbook(&self) -> WorkbookModel {
        WorkbookModel {
            sheets: Vec::new(),
            date_system: self.date_system,
            defined_names: self.defined_names.clone(),
            shared_strings: self.shared_strings.clone(),
            styles: self.styles.clone(),
            tables: self.tables.clone(),
        }
    }
}

/// The part of an anchor no move may rewrite: its kind, plus whatever the
/// drawing writer cannot patch. Only the grid markers a save writes whole stay
/// free, so this mirrors `only_grid_position_moved` in the writer.
#[derive(Clone, Debug, PartialEq, Eq)]
enum AnchorShape {
    TwoCell {
        edit_as: AnchorEditAs,
    },
    OneCell {
        extent: AnchorExtent,
    },
    Absolute {
        pos: AnchorPos,
        extent: AnchorExtent,
    },
}

impl AnchorShape {
    fn of(anchor: &ChartAnchor) -> Self {
        match anchor {
            ChartAnchor::TwoCell { edit_as, .. } => Self::TwoCell { edit_as: *edit_as },
            ChartAnchor::OneCell { extent, .. } => Self::OneCell { extent: *extent },
            ChartAnchor::Absolute { pos, extent } => Self::Absolute {
                pos: *pos,
                extent: *extent,
            },
        }
    }
}

/// What the freeze pins about a chart: which drawing anchors which part, what
/// the part reads, and the shape of that anchor. Only where a grid-anchored
/// chart sits is replicated content, so a peer may slide one without changing
/// the structure this describes.
#[derive(Clone, Debug, PartialEq, Eq)]
struct ChartIdentity {
    part: String,
    drawing: String,
    anchor_index: usize,
    refs: Vec<ChartRef>,
    anchor: AnchorShape,
}

impl ChartIdentity {
    fn of(chart: &SheetChart) -> Self {
        Self {
            part: chart.part.clone(),
            drawing: chart.drawing.clone(),
            anchor_index: chart.anchor_index,
            refs: chart.refs.clone(),
            anchor: AnchorShape::of(&chart.anchor),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct WorkbookStructure {
    generation: i64,
    pub(crate) sheet_keys: Vec<String>,
    sheet_names: Vec<String>,
    freeze_panes: Vec<Option<FreezePane>>,
    hyperlinks: Vec<Vec<Hyperlink>>,
    baseline_hyperlinks: BTreeMap<String, Vec<Hyperlink>>,
    conditional_formats: Vec<Vec<ConditionalFormatting>>,
    auto_filters: Vec<Option<xlsx_model::AutoFilter>>,
    charts: Vec<Vec<ChartIdentity>>,
    merges: Vec<Vec<CellRange>>,
    baseline_merges: BTreeMap<String, Vec<CellRange>>,
    shared_types: BTreeMap<String, SheetSharedTypes>,
}

impl WorkbookStructure {
    /// Permit sheet-local layout changes while keeping structural identities pinned.
    pub(crate) fn same_except_local_layout(&self, other: &Self) -> bool {
        self.generation == other.generation
            && self.sheet_keys == other.sheet_keys
            && self.sheet_names == other.sheet_names
            && self.baseline_hyperlinks == other.baseline_hyperlinks
            && self.conditional_formats == other.conditional_formats
            && self.auto_filters == other.auto_filters
            && self.charts == other.charts
            && self.baseline_merges == other.baseline_merges
            && self.shared_types == other.shared_types
    }

    pub(crate) fn snapshot_same_except_local_layout(&self, other: &Self) -> bool {
        self.generation == other.generation
            && self.sheet_keys == other.sheet_keys
            && self.sheet_names == other.sheet_names
            && self.baseline_hyperlinks == other.baseline_hyperlinks
            && self.conditional_formats == other.conditional_formats
            && self.auto_filters == other.auto_filters
            && self.charts == other.charts
            && self.baseline_merges == other.baseline_merges
    }

    /// Whether two structures describe the same workbook, disregarding the Yrs
    /// branch identities. Replacing a bootstrap rebuilds every shared type, so
    /// those always differ and cannot say whether the structure itself did.
    pub(crate) fn describes_same_workbook(&self, other: &Self) -> bool {
        self.sheet_keys == other.sheet_keys
            && self.sheet_names == other.sheet_names
            && self.freeze_panes == other.freeze_panes
            && self.hyperlinks == other.hyperlinks
            && self.baseline_hyperlinks == other.baseline_hyperlinks
            && self.conditional_formats == other.conditional_formats
            && self.auto_filters == other.auto_filters
            && self.charts == other.charts
            && self.merges == other.merges
            && self.baseline_merges == other.baseline_merges
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SheetSharedTypes {
    sheet: BranchID,
    col_widths: BranchID,
    contents: BranchID,
    row_heights: BranchID,
    styles: BranchID,
}

/// What a replica can do with an update offered as its whole state.
pub(crate) enum SnapshotAdoption {
    /// Not a whole document, or this replica already holds more than its own
    /// bootstrap. The update is an ordinary one; merge it.
    NotApplicable,
    /// A whole document this replica cannot take on.
    Incompatible(String),
    Replacement(Box<WorkbookAuthority>),
}

pub(crate) struct StagedUpdate {
    pub(crate) commit_update: Vec<u8>,
    pub(crate) effective: bool,
    pub(crate) model: WorkbookModel,
    pub(crate) pending: bool,
    pub(crate) state_bytes: usize,
    pub(crate) state_vector_entries: usize,
    pub(crate) structure: WorkbookStructure,
    pub(crate) update: Vec<u8>,
}

pub(crate) struct StagedLocalUpdate {
    pub(crate) model: WorkbookModel,
    pub(crate) state_bytes: usize,
    pub(crate) state_vector_entries: usize,
    pub(crate) structure: WorkbookStructure,
    pub(crate) update: Vec<u8>,
    carried: CarriedState,
}

/// What staging changed beside the document, handed to the live authority on adoption.
struct CarriedState {
    base: Arc<WorkbookBase>,
    next_sheet_id: u64,
    sheet_order: Option<SheetOrderEntry>,
}

/// Which history an adopted local update enters.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum LocalHistory {
    /// One collaborative undo item.
    Undo,
    /// The sheet-order entry a standalone undo replays against.
    SheetOrder,
    /// Neither; existing undo and redo entries stay.
    None,
}

pub(crate) struct AuthorityCheckpoint {
    state: Vec<u8>,
    undo_stack: Vec<StackItem<()>>,
    redo_stack: Vec<StackItem<()>>,
}

pub(crate) struct HistoryUpdate {
    pub(crate) model: WorkbookModel,
    pub(crate) structure: WorkbookStructure,
    pub(crate) update: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct SheetOrderEntry {
    before: Vec<String>,
    after: Vec<String>,
}

#[derive(Clone, Default, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct SheetOrderHistory {
    undo: Vec<SheetOrderEntry>,
    redo: Vec<SheetOrderEntry>,
}

enum HistoryAction {
    Push(SheetOrderEntry),
    Undo(SheetOrderEntry),
    Redo(SheetOrderEntry),
}

#[cfg(test)]
thread_local! {
    pub(crate) static SNAPSHOT_VECTOR_ENCODINGS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    static SCHEMA_SIX_KEY_VALIDATION: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    pub(crate) static LOCAL_UPDATE_ENCODINGS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[cfg(test)]
fn with_schema_six_key_validation<T>(body: impl FnOnce() -> T) -> T {
    struct Reset(bool);
    impl Drop for Reset {
        fn drop(&mut self) {
            SCHEMA_SIX_KEY_VALIDATION.set(self.0);
        }
    }
    let _reset = Reset(SCHEMA_SIX_KEY_VALIDATION.replace(true));
    body()
}

pub(crate) struct WorkbookAuthority {
    doc: Doc,
    projection_valid: Arc<AtomicBool>,
    snapshot_revision: Arc<AtomicU64>,
    base: Arc<WorkbookBase>,
    history: SheetOrderHistory,
    next_sheet_id: u64,
    undo_stack: Vec<StackItem<()>>,
    redo_stack: Vec<StackItem<()>>,
}

struct SetCellSync {
    keys: Vec<String>,
    sheet_map: MapRef,
    cell_formats: MapRef,
    styles: Stylesheet,
    sheet: Sheet,
    at: CellRef,
    formatted: bool,
}

impl WorkbookAuthority {
    pub(crate) fn recovery_sheet_order_history(&self) -> SheetOrderHistory {
        self.history.clone()
    }

    pub(crate) fn restore_recovery_sheet_order_history(
        &mut self,
        history: SheetOrderHistory,
        undo_depth: usize,
        redo_depth: usize,
    ) -> crate::snapshot::SnapshotResult<()> {
        if history.undo.len() != undo_depth || history.redo.len() != redo_depth {
            return Err(crate::snapshot::SnapshotError::new(
                "Recovery sheet-order history depth mismatch",
            ));
        }
        self.history = history;
        Ok(())
    }

    #[cfg(test)]
    pub(crate) fn encode_recovery_state(&self) -> crate::snapshot::SnapshotResult<Vec<u8>> {
        Ok(self.recovery_state_parts()?.concat())
    }

    /// The recovery state in pieces whose concatenation is the state, so the
    /// caller frames the Yrs update without copying it into a buffer of its own.
    pub(crate) fn recovery_state_parts(&self) -> crate::snapshot::SnapshotResult<[Vec<u8>; 3]> {
        use crate::snapshot::wire::Writer;
        use snapshot::Codec;
        let mut writer = Writer::new();
        let base = &self.base;
        base.bootstrap_client_id.write(&mut writer);
        base.date_system.write(&mut writer);
        base.defined_names.write(&mut writer);
        base.fingerprint.write(&mut writer);
        base.fingerprints.write(&mut writer);
        base.freeze_panes.write(&mut writer);
        base.formats.write(&mut writer);
        base.col_styles.write(&mut writer);
        base.hyperlinks.write(&mut writer);
        base.charts.write(&mut writer);
        base.hidden_dimensions.write(&mut writer);
        base.shared_strings.write(&mut writer);
        base.tables.write(&mut writer);
        let styles = serde_json::to_vec(&base.styles)
            .map_err(|error| crate::snapshot::SnapshotError::new(error.to_string()))?;
        writer.bytes(&styles);
        let update = self.encode_state_as_update_v1();
        writer.var_usize(update.len());
        let mut tail = Writer::new();
        base.row_styles.write(&mut tail);
        if base.auto_filters.iter().any(Option::is_some)
            || base
                .conditional_formats
                .iter()
                .any(|formats| !formats.is_empty())
        {
            let formats = serde_json::to_vec(&base.conditional_formats)
                .map_err(|error| crate::snapshot::SnapshotError::new(error.to_string()))?;
            tail.bytes(&formats);
            if base.auto_filters.iter().any(Option::is_some) {
                let filters = serde_json::to_vec(&base.auto_filters)
                    .map_err(|error| crate::snapshot::SnapshotError::new(error.to_string()))?;
                tail.bytes(&filters);
            }
        }
        Ok([writer.into_bytes(), update, tail.into_bytes()])
    }

    /// The recovery state as releases before row defaults encoded it.
    #[cfg(test)]
    pub(crate) fn encode_released_recovery_state(&self) -> Vec<u8> {
        use crate::snapshot::wire::Writer;
        use snapshot::Codec;
        let mut writer = Writer::new();
        let base = &self.base;
        base.bootstrap_client_id.write(&mut writer);
        base.date_system.write(&mut writer);
        base.defined_names.write(&mut writer);
        base.fingerprint.write(&mut writer);
        base.fingerprints.write(&mut writer);
        base.freeze_panes.write(&mut writer);
        base.formats.write(&mut writer);
        base.col_styles.write(&mut writer);
        base.hyperlinks.write(&mut writer);
        base.charts.write(&mut writer);
        base.hidden_dimensions.write(&mut writer);
        base.shared_strings.write(&mut writer);
        base.tables.write(&mut writer);
        writer.bytes(&serde_json::to_vec(&base.styles).unwrap());
        writer.bytes(&self.encode_state_as_update_v1());
        writer.into_bytes()
    }

    /// `row_styles` reads the row defaults that follow the Yrs state; a state
    /// releases before them encoded ends there.
    pub(crate) fn restore_recovery_state(
        bytes: &[u8],
        client_id: u64,
        next_sheet_id: u64,
        row_styles: bool,
    ) -> crate::snapshot::SnapshotResult<Self> {
        use crate::snapshot::SnapshotError;
        use crate::snapshot::wire::Reader;
        use snapshot::Codec;
        if client_id > MAX_SAFE_CLIENT_ID {
            return Err(SnapshotError::new("Recovery client id is out of range"));
        }
        let mut reader = Reader::new(bytes);
        let mut base = WorkbookBase {
            bootstrap_client_id: Codec::read(&mut reader)?,
            date_system: Codec::read(&mut reader)?,
            defined_names: Codec::read(&mut reader)?,
            fingerprint: Codec::read(&mut reader)?,
            fingerprints: Codec::read(&mut reader)?,
            freeze_panes: Codec::read(&mut reader)?,
            formats: Codec::read(&mut reader)?,
            col_styles: Codec::read(&mut reader)?,
            row_styles: Vec::new(),
            conditional_formats: Vec::new(),
            auto_filters: Vec::new(),
            hyperlinks: Codec::read(&mut reader)?,
            charts: Codec::read(&mut reader)?,
            hidden_dimensions: Codec::read(&mut reader)?,
            shared_strings: Codec::read(&mut reader)?,
            tables: Codec::read(&mut reader)?,
            styles: serde_json::from_slice(reader.bytes()?)
                .map_err(|error| SnapshotError::new(error.to_string()))?,
        };
        if client_id == base.bootstrap_client_id {
            return Err(SnapshotError::new(
                "Recovery client id conflicts with the base",
            ));
        }
        let update = reader.bytes()?;
        if row_styles {
            base.row_styles = Codec::read(&mut reader)?;
            if !reader.is_empty() {
                base.conditional_formats = serde_json::from_slice(reader.bytes()?)
                    .map_err(|error| SnapshotError::new(error.to_string()))?;
            }
        }
        if !reader.is_empty() {
            base.auto_filters = serde_json::from_slice(reader.bytes()?)
                .map_err(|error| SnapshotError::new(error.to_string()))?;
        }
        reader.finish()?;
        let doc = loop {
            let candidate = Doc::new();
            let id = candidate.client_id();
            if id.get() == client_id || id.get() == base.bootstrap_client_id {
                continue;
            }
            hydrate_local_doc(&candidate, update).map_err(SnapshotError::new)?;
            let existing_client = candidate.transact().state_vector().contains_client(&id);
            if !existing_client {
                break candidate;
            }
        };
        if has_pending(&doc) {
            return Err(SnapshotError::new(
                "Recovery authority has unresolved Yrs state",
            ));
        }
        let authority = Self::hydrated(doc, Arc::new(base), next_sheet_id);
        authority.strict_materialize().map_err(SnapshotError::new)?;
        Ok(authority)
    }

    fn hydrated(doc: Doc, base: Arc<WorkbookBase>, next_sheet_id: u64) -> Self {
        let snapshot_revision = Arc::new(AtomicU64::new(0));
        let projection_valid = observe_projection(&doc, snapshot_revision.clone());
        Self {
            doc,
            projection_valid,
            snapshot_revision,
            base,
            history: SheetOrderHistory::default(),
            next_sheet_id,
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
        }
    }

    #[cfg(test)]
    fn from_model(model: &WorkbookModel) -> Result<Self, AuthorityError> {
        Self::from_source(model, None, &[], None)
    }

    #[cfg(test)]
    fn from_model_with_client_id(
        model: &WorkbookModel,
        client_id: u64,
    ) -> Result<Self, AuthorityError> {
        Self::from_source(model, Some(client_id), &[], None)
    }

    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn from_source(
        model: &WorkbookModel,
        client_id: Option<u64>,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
    ) -> Result<Self, AuthorityError> {
        Self::from_source_with_projection(model, client_id, legacy_dimensions, legacy_styles, &[])
            .map(|(authority, _, _)| authority)
    }

    pub(crate) fn from_source_with_projection(
        model: &WorkbookModel,
        client_id: Option<u64>,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
        legacy_style_readings: &[Stylesheet],
    ) -> Result<(Self, WorkbookModel, WorkbookStructure), AuthorityError> {
        Self::from_source_with_seed_chunk(
            model,
            client_id,
            legacy_dimensions,
            legacy_styles,
            legacy_style_readings,
            SEED_CHUNK,
        )
    }

    fn from_source_with_seed_chunk(
        model: &WorkbookModel,
        client_id: Option<u64>,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
        legacy_style_readings: &[Stylesheet],
        chunk: usize,
    ) -> Result<(Self, WorkbookModel, WorkbookStructure), AuthorityError> {
        let base = WorkbookBase::from_model_with_style_readings(
            model,
            legacy_dimensions,
            legacy_styles,
            legacy_style_readings,
        )
        .map_err(AuthorityError::InvalidState)?;
        if client_id == Some(base.bootstrap_client_id) {
            return Err(AuthorityError::ClientIdConflict(base.bootstrap_client_id));
        }

        let doc = match client_id {
            Some(client_id) => Doc::with_client_id(client_id),
            None => loop {
                let candidate = Doc::new();
                if candidate.client_id().get() != base.bootstrap_client_id {
                    break candidate;
                }
            },
        };
        let bootstrap = Doc::with_client_id(base.bootstrap_client_id);
        let keys = (0..model.sheets.len())
            .map(|index| format!("sheet:{index}"))
            .collect::<Vec<_>>();
        let mut updates = Vec::new();
        let flush = |update: Vec<u8>| {
            updates.push(update);
            Ok(())
        };
        seed_in_chunks(&bootstrap, &base, model, &keys, style_entry, chunk, flush)
            .map_err(AuthorityError::InvalidState)?;
        drop(bootstrap);
        for update in updates {
            hydrate_local_doc(&doc, &update).map_err(AuthorityError::InvalidState)?;
            xlsx_model::progress::pulse();
        }
        let authority = Self::hydrated(doc, Arc::new(base), 0);
        let (model, structure) = authority
            .strict_materialize()
            .map_err(AuthorityError::InvalidState)?;
        Ok((authority, model, structure))
    }

    pub(crate) fn client_id(&self) -> u64 {
        self.doc.client_id().get()
    }

    pub(crate) fn snapshot_revision(&self) -> u64 {
        self.snapshot_revision.load(Ordering::Relaxed)
    }

    pub(crate) fn snapshot_projection_valid(&self) -> crate::snapshot::SnapshotResult<bool> {
        let _transaction = self.doc.try_transact().map_err(|_| {
            crate::snapshot::SnapshotError::new("snapshot authority has an active transaction")
        })?;
        if !self.history.undo.is_empty()
            || !self.history.redo.is_empty()
            || !self.undo_stack.is_empty()
            || !self.redo_stack.is_empty()
            || self.next_sheet_id != 0
            || has_pending(&self.doc)
        {
            return Err(crate::snapshot::SnapshotError::new(
                "snapshot authority is not at the initial boundary",
            ));
        }
        self.base.styles.snapshot_field_counts();
        Ok(self.projection_valid.load(Ordering::Relaxed))
    }

    pub(crate) fn set_snapshot_projection_valid(&mut self, valid: bool) {
        self.projection_valid.store(valid, Ordering::Relaxed);
    }

    #[cfg(test)]
    pub(crate) fn snapshot_transaction_for_test(&self) -> TransactionMut<'_> {
        self.doc.transact_mut()
    }

    #[cfg(test)]
    pub(crate) fn snapshot_skip_gc_for_test(&mut self) {
        let mut options = Options::with_guid_and_client_id(self.doc.guid(), self.doc.client_id());
        options.skip_gc = true;
        let doc = Doc::with_options(options);
        hydrate_local_doc(&doc, &self.encode_state_as_update_v1()).unwrap();
        self.doc = doc;
        self.snapshot_revision.fetch_add(1, Ordering::Relaxed);
        self.projection_valid = observe_projection(&self.doc, self.snapshot_revision.clone());
    }

    #[cfg(test)]
    #[doc(hidden)]
    pub(crate) fn snapshot_checkpoint_for_test(&mut self) {
        let mut undo = build_undo_manager(
            &self.doc,
            std::mem::take(&mut self.undo_stack),
            std::mem::take(&mut self.redo_stack),
        )
        .unwrap();
        undo.clear_all();
        drop(undo);
        self.doc.transact_mut_with(HYDRATE_ORIGIN).gc(None);
        self.clear_history();
        assert_eq!(self.next_sheet_id, 0);
    }

    pub(crate) fn state_vector_entries(&self) -> usize {
        self.doc.transact().state_vector().len()
    }

    pub(crate) fn materialize(&self) -> Result<WorkbookModel, AuthorityError> {
        self.materialize_internal(false)
            .map(|(model, _)| model)
            .map_err(AuthorityError::InvalidState)
    }

    pub(crate) fn structure(&self) -> Result<WorkbookStructure, AuthorityError> {
        self.materialize_internal(false)
            .map(|(_, structure)| structure)
            .map_err(AuthorityError::InvalidState)
    }

    /// Applies `ops`, whose style indices index `styles`, the caller's style table.
    #[cfg(test)]
    pub(crate) fn apply_ops(
        &mut self,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
    ) -> Result<Option<Vec<u8>>, AuthorityError> {
        self.apply_ops_with_update(ops, origin, styles, true)
    }

    pub(crate) fn apply_ops_with_update(
        &mut self,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
        encode_update: bool,
    ) -> Result<Option<Vec<u8>>, AuthorityError> {
        self.apply_ops_with_merge_edits(ops, origin, styles, encode_update, false)
    }

    fn apply_ops_with_merge_edits(
        &mut self,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
        encode_update: bool,
        collaborative: bool,
    ) -> Result<Option<Vec<u8>>, AuthorityError> {
        let state_vector = encode_update.then(|| self.doc.transact().state_vector());
        if let Some(prepared) = self.prepare_set_cell_sync(ops, styles)
            && let Ok((keys, history)) =
                self.plan_sheet_keys(&prepared.keys, ops, prepared.keys.len(), origin)
            && keys == prepared.keys
            && self.validate_sync_state(&prepared.keys, &keys).is_ok()
        {
            let mut txn = self.doc.transact_mut_with(origin.as_str());
            sync_cell_formats(&prepared.cell_formats, &mut txn, &prepared.styles)
                .map_err(AuthorityError::InvalidState)?;
            sync_authored_cell(&prepared.sheet_map, &mut txn, &prepared.sheet, prepared.at)
                .map_err(AuthorityError::InvalidState)?;
            if prepared.formatted {
                sync_cell_format(
                    &prepared.sheet_map,
                    &mut txn,
                    &prepared.sheet,
                    &prepared.styles,
                    prepared.at,
                )
                .map_err(AuthorityError::InvalidState)?;
            }
            drop(txn);
            self.apply_history(history);
            self.projection_valid.store(true, Ordering::Relaxed);
            #[cfg(test)]
            FAST_SET_CELL_COUNT.set(FAST_SET_CELL_COUNT.get() + 1);
            return Ok(self.encode_ops_update(state_vector));
        }
        let mut model = self.materialize()?;
        model.styles.bind_plain_xf();
        let ops =
            remap_styles(ops, styles, &mut model.styles).map_err(AuthorityError::InvalidState)?;
        let ops = ops.as_ref();
        let authored_styles = ops
            .iter()
            .filter_map(|op| match op {
                Op::SetCell { sheet, at, .. } => Some((
                    (sheet.0, *at),
                    model
                        .sheet(*sheet)
                        .and_then(|sheet| sheet.cell(*at))
                        .and_then(|cell| cell.style),
                )),
                _ => None,
            })
            .collect::<HashMap<_, _>>();
        for op in ops {
            xlsx_ops::apply_in_place(&mut model, op).map_err(|error| {
                AuthorityError::InvalidState(format!(
                    "cannot apply local operation to authored state: {error}"
                ))
            })?;
        }
        if self.base.defined_names != model.defined_names {
            Arc::make_mut(&mut self.base).defined_names = model.defined_names.clone();
        }
        if self.base.tables != model.tables {
            Arc::make_mut(&mut self.base).tables = model.tables.clone();
        }
        self.sync_model(&model, ops, origin, &authored_styles, collaborative)
            .map_err(AuthorityError::InvalidState)?;
        if let [Op::SetCell { at, cell, .. }] = ops
            && at.row < MAX_ROWS
            && at.col < MAX_COLS
            && !matches!(cell.value, CellValue::Number { value } if !value.is_finite())
        {
            self.projection_valid.store(true, Ordering::Relaxed);
        }
        Ok(self.encode_ops_update(state_vector))
    }

    fn encode_ops_update(&self, state_vector: Option<StateVector>) -> Option<Vec<u8>> {
        let state_vector = state_vector?;
        #[cfg(test)]
        LOCAL_UPDATE_ENCODINGS.set(LOCAL_UPDATE_ENCODINGS.get() + 1);
        let update = self.doc.transact().encode_diff_v1(&state_vector);
        (update.as_slice() != Update::EMPTY_V1).then_some(update)
    }

    fn prepare_set_cell_sync(&self, ops: &[Op], source: &Stylesheet) -> Option<SetCellSync> {
        #[cfg(test)]
        if force_full_materialization() {
            return None;
        }
        let [Op::SetCell { sheet, at, cell }] = ops else {
            return None;
        };
        if !self.projection_valid.load(Ordering::Relaxed)
            || at.row >= MAX_ROWS
            || at.col >= MAX_COLS
            || matches!(cell.value, CellValue::Number { value } if !value.is_finite())
            || self.schema_version().ok()? != SCHEMA_VERSION
            || !self.has_current_base_fingerprint()
        {
            return None;
        }
        let txn = self.doc.transact();
        let cell_formats = txn.get_map(CELL_FORMATS)?;
        let (mut styles, _) =
            materialize_cell_formats(&cell_formats, &txn, &self.base.styles).ok()?;
        styles.bind_plain_xf();
        let order = txn.get_array(SHEET_ORDER)?;
        let keys = sheet_keys(&order, &txn).ok()?;
        let key = keys.get(sheet.0 as usize)?;
        let sheets = txn.get_map(SHEETS)?;
        let sheet_map = sheets.get(&txn, key)?.cast::<MapRef>().ok()?;
        let contents = nested_map(&sheet_map, &txn, CONTENTS).ok()?;
        let key = cell_key(*at);
        match contents.get(&txn, &key) {
            Some(Out::Any(value)) => {
                content_from_any(&value).ok()?;
            }
            Some(_) => return None,
            None => {}
        }
        let authored_styles = nested_map(&sheet_map, &txn, STYLES).ok()?;
        let authored_style = match authored_styles.get(&txn, &key) {
            Some(value) => Some(value.cast::<String>().ok()?),
            None => None,
        };
        let remapped = remap_styles(ops, source, &mut styles).ok()?;
        let Op::SetCell { cell, .. } = &remapped[0] else {
            return None;
        };
        cell_format_entry(&CellFormat::default()).ok()?;
        for index in 0..styles.cell_xfs.len() {
            cell_format_entry(&styles.cell_format(Some(u32::try_from(index).ok()?))).ok()?;
        }
        let style = cell
            .style
            .map(|style| style_key(&styles, style))
            .transpose()
            .ok()?;
        let formatted = style != authored_style;
        let mut target = Sheet::default();
        target.set_cell(*at, cell.clone().into());
        Some(SetCellSync {
            keys,
            sheet_map,
            cell_formats,
            styles,
            sheet: target,
            at: *at,
            formatted,
        })
    }

    pub(crate) fn encode_state_vector_v1(&self) -> Vec<u8> {
        #[cfg(test)]
        SNAPSHOT_VECTOR_ENCODINGS.set(SNAPSHOT_VECTOR_ENCODINGS.get() + 1);
        let state_vector = self.doc.transact().state_vector();
        let mut entries = state_vector
            .iter()
            .map(|(client, clock)| (client.get(), *clock))
            .collect::<Vec<_>>();
        entries.sort_unstable();
        let mut encoder = EncoderV1::new();
        encoder.write_var(entries.len());
        for (client, clock) in entries {
            encoder.write_var(client);
            encoder.write_var(clock);
        }
        encoder.to_vec()
    }

    pub(crate) fn encode_state_as_update_v1(&self) -> Vec<u8> {
        self.doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default())
    }

    pub(crate) fn encode_diff_v1(
        &self,
        remote_state_vector: &[u8],
    ) -> Result<Vec<u8>, AuthorityError> {
        let state_vector = decode_state_vector_v1(remote_state_vector)
            .map_err(AuthorityError::InvalidStateVector)?;
        Ok(self.doc.transact().encode_diff_v1(&state_vector))
    }

    /// The replica this one would become by taking a persisted snapshot as its
    /// whole state and upgrading it to the current schema.
    ///
    /// Merging cannot restore a snapshot an earlier release wrote. The
    /// bootstrap client ID is the head of the base fingerprint, which hashes
    /// the schema version, so a bootstrap this build seeds never dedupes
    /// against the one such a snapshot carries: the two bases double up and
    /// the survivor is whichever client ID sorts higher. A snapshot supersedes
    /// an untouched bootstrap of the same workbook whatever schema it was
    /// written at, and where the bootstraps do agree the two are the same
    /// document, so adopting is never worse than merging.
    #[cfg(test)]
    pub(crate) fn snapshot_replacement(&self, update: &[u8]) -> SnapshotAdoption {
        self.snapshot_replacement_with_legacy_styles(update, || None)
    }

    pub(crate) fn snapshot_replacement_with_legacy_styles(
        &self,
        update: &[u8],
        legacy_styles: impl FnOnce() -> Option<xlsx_parse::LegacyStyleReadings>,
    ) -> SnapshotAdoption {
        if !self.is_pristine() {
            return SnapshotAdoption::NotApplicable;
        }
        let doc = Doc::with_client_id(self.client_id());
        if hydrate_doc(&doc, update).is_err() {
            return SnapshotAdoption::NotApplicable;
        }
        let mut candidate = Self::hydrated(doc, self.base.clone(), self.next_sheet_id);
        if !candidate.is_whole_document() {
            return SnapshotAdoption::NotApplicable;
        }
        let accepted = candidate
            .schema_version()
            .ok()
            .zip(candidate.state_fingerprint())
            .is_some_and(|(version, fingerprint)| {
                self.base.accepts_fingerprint(version, &fingerprint)
            });
        if !accepted
            && let Some(legacy) = legacy_styles()
            && let Ok(base) = WorkbookBase::from_model_with_style_readings(
                &legacy.workbook,
                &legacy.dimensions,
                None,
                &legacy.styles,
            )
        {
            let target = Arc::make_mut(&mut candidate.base);
            for (version, accepted) in base.fingerprints {
                for fingerprint in accepted {
                    append_fingerprint(&mut target.fingerprints, version, fingerprint);
                }
            }
        }
        let released = candidate.holds_released_reading();
        if let Err(error) = candidate.upgrade_schema_from(Some(&self.doc)) {
            return SnapshotAdoption::Incompatible(error);
        }
        if released && let Err(error) = candidate.keep_released_source_formats(self) {
            return SnapshotAdoption::Incompatible(error);
        }
        if let Err(error) = candidate.keep_source_italic(self) {
            return SnapshotAdoption::Incompatible(error);
        }
        match candidate.strict_materialize() {
            Err(error) => SnapshotAdoption::Incompatible(error),
            Ok(_) => SnapshotAdoption::Replacement(Box::new(candidate)),
        }
    }

    fn holds_released_reading(&self) -> bool {
        fingerprint_reading(&self.base.fingerprint).as_written()
            && !self
                .state_fingerprint()
                .is_some_and(|fingerprint| fingerprint_reading(&fingerprint).as_written())
    }

    /// Points each sheet's conditional formats back at the source's where a
    /// release before italic wrote them without it, so the state reads as the
    /// source does.
    fn keep_source_italic(&self, source: &Self) -> Result<(), String> {
        let mut txn = self.doc.transact_mut_with(HYDRATE_ORIGIN);
        let source_txn = source.doc.transact();
        let source_sheets = source_txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        for key in sheets.keys(&txn).map(str::to_string).collect::<Vec<_>>() {
            let Some(sheet) = sheets
                .get(&txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
            else {
                continue;
            };
            let Some(Out::Any(Any::String(restored))) = sheet.get(&txn, CONDITIONAL_FORMATS) else {
                continue;
            };
            let Some(Out::Any(Any::String(own))) = source_sheets
                .get(&source_txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .and_then(|sheet| sheet.get(&source_txn, CONDITIONAL_FORMATS))
            else {
                continue;
            };
            if restored == own {
                continue;
            }
            let read = |json: &str| serde_json::from_str::<Vec<ConditionalFormatting>>(json).ok();
            if let (Some(restored), Some(original)) = (read(&restored), read(&own))
                && xlsx_parse::keeps_source_conditional_formats(&restored, &original)
            {
                sheet.try_update(&mut txn, CONDITIONAL_FORMATS, own.to_string());
            }
        }
        Ok(())
    }

    /// Points each cell of a state a release before patterned fills wrote back
    /// at its own source format where that release held the format's released
    /// reading. The reading keeps its catalog entry, which that release's
    /// writers still reference.
    fn keep_released_source_formats(&self, source: &Self) -> Result<(), String> {
        use snapshot::ReleasedLayout;
        let base = &self.base.styles;
        let mut sources = HashMap::new();
        for index in 0..base.cell_xfs.len() {
            let index =
                u32::try_from(index).map_err(|_| "cell format table is too large".to_string())?;
            let format = base.cell_format(Some(index));
            if format.fill.released_layout() && format.border.released_layout() {
                continue;
            }
            let (key, payload) = style_entry(base, index)?;
            sources.entry(key).or_insert((format, payload));
        }
        if sources.is_empty() {
            return Ok(());
        }
        let mut txn = self.doc.transact_mut_with(HYDRATE_ORIGIN);
        let catalog = txn
            .get_map(CELL_FORMATS)
            .ok_or_else(|| "missing cell format catalog".to_string())?;
        let source_txn = source.doc.transact();
        let source_sheets = source_txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        let mut restored_formats = HashMap::new();
        for key in sheets.keys(&txn).map(str::to_string).collect::<Vec<_>>() {
            let styles = sheets
                .get(&txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("sheet {key} is not a map"))
                .and_then(|sheet| nested_map(&sheet, &txn, STYLES))?;
            let source_styles = source_sheets
                .get(&source_txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .and_then(|sheet| nested_map(&sheet, &source_txn, STYLES).ok());
            let mut kept = Vec::new();
            for (cell, value) in styles.iter(&txn) {
                let restored = value
                    .cast::<String>()
                    .map_err(|_| format!("cell style {cell} is not a string"))?;
                let own = source_styles
                    .as_ref()
                    .and_then(|styles| styles.get(&source_txn, cell))
                    .and_then(|value| value.cast::<String>().ok())
                    .filter(|own| *own != restored);
                let Some((own, (format, _))) =
                    own.and_then(|own| sources.get(&own).map(|source| (own, source)))
                else {
                    continue;
                };
                let current = restored_formats.entry(restored.clone()).or_insert_with(|| {
                    match catalog.get(&txn, &restored) {
                        Some(Out::Any(Any::String(payload))) => {
                            serde_json::from_str::<CellFormat>(&payload).ok()
                        }
                        _ => None,
                    }
                });
                let keeps = current.as_ref().is_some_and(|current| {
                    xlsx_parse::keeps_source(current, format, true, released_format)
                });
                if keeps {
                    kept.push((cell.to_string(), own));
                }
            }
            for (cell, own) in kept {
                catalog.try_update(&mut txn, own.clone(), sources[&own].1.clone());
                styles.try_update(&mut txn, cell, own);
            }
        }
        Ok(())
    }

    /// Whether `update` holds a whole workbook seeded by another bootstrap
    /// than the one this replica's document grew from, which only
    /// `snapshot_replacement` may take into a pristine replica. Merged into an
    /// edited one, its seed would race this replica's seed for every shared
    /// key.
    pub(crate) fn seeds_another_bootstrap(&self, update: &[u8]) -> bool {
        let Ok(incoming) = decode_update_v1(update) else {
            return false;
        };
        let known = self.doc.transact().state_vector();
        if incoming
            .state_vector()
            .iter()
            .all(|(client, _)| known.contains_client(client))
        {
            return false;
        }
        let seeds = self.seed_clients(&known);
        let doc = Doc::with_client_id(self.client_id());
        if hydrate_doc(&doc, update).is_err() {
            return false;
        }
        let candidate = Self::hydrated(doc, self.base.clone(), self.next_sheet_id);
        candidate.is_whole_document()
            && !candidate
                .doc
                .transact()
                .state_vector()
                .iter()
                .any(|(client, _)| seeds.contains(&client.get()))
    }

    /// The bootstraps `state_vector` holds: this replica's own, or the one
    /// that seeded a snapshot it adopted, under any reading of the base it
    /// accepts.
    fn seed_clients(&self, state_vector: &StateVector) -> Vec<u64> {
        let bootstraps = self
            .base
            .fingerprints
            .values()
            .flatten()
            .filter_map(|fingerprint| fingerprint_client_id(fingerprint.as_str()))
            .chain([self.base.bootstrap_client_id])
            .collect::<HashSet<_>>();
        state_vector
            .iter()
            .map(|(client, _)| client.get())
            .filter(|client| bootstraps.contains(client))
            .collect()
    }

    /// True while the replica still holds nothing but its own bootstrap.
    fn is_pristine(&self) -> bool {
        let txn = self.doc.transact();
        let state_vector = txn.state_vector();
        state_vector.len() == 1
            && state_vector
                .iter()
                .all(|(client, _)| client.get() == self.base.bootstrap_client_id)
            && txn.snapshot().delete_set.is_empty()
    }

    /// True when the document stands on its own rather than being the tail of
    /// someone else's — an incremental update hydrates into neither the roots
    /// nor the metadata a whole workbook carries.
    fn is_whole_document(&self) -> bool {
        let txn = self.doc.transact();
        if require_root_keys(&txn, &[CELL_FORMATS, META, SHEET_ORDER, SHEETS]).is_err() {
            return false;
        }
        txn.get_map(META).is_some_and(|meta| {
            require_map_keys(
                &meta,
                &txn,
                &[BASE_FINGERPRINT, "schemaVersion", STRUCTURE_GENERATION],
                "workbook metadata",
            )
            .is_ok()
        })
    }

    /// `baseline` may carry this replica's already-encoded current state.
    pub(crate) fn stage_updates_v1(
        &self,
        updates: &[&[u8]],
        baseline: Option<&[u8]>,
    ) -> Result<StagedUpdate, AuthorityError> {
        if updates.is_empty() {
            return Err(AuthorityError::InvalidUpdate(
                "no updates were provided".to_string(),
            ));
        }
        let merged;
        let (bytes, local) = match updates {
            [update] => (*update, false),
            _ => {
                let decoded = updates
                    .iter()
                    .map(|update| decode_update_v1(update).map_err(AuthorityError::InvalidUpdate))
                    .collect::<Result<Vec<_>, _>>()?;
                merged = Update::merge_updates(decoded).encode_v1();
                (merged.as_slice(), true)
            }
        };
        let (incoming, layout) =
            decode_update_layout_v1(bytes, local).map_err(AuthorityError::InvalidUpdate)?;
        let before_vector = self.doc.transact().state_vector();
        let owned;
        let before = match baseline {
            Some(baseline) => baseline,
            None => {
                owned = self.encode_state_as_update_v1();
                &owned
            }
        };
        let staged_doc = Doc::with_client_id(self.client_id());
        hydrate_local_doc(&staged_doc, before).map_err(AuthorityError::InvalidState)?;
        let known = staged_doc.transact().state_vector();
        let (held_back, applied_deletes) = match hold_back_unreachable_v1(bytes, &layout, &known) {
            Some(reachable) => {
                let reachable =
                    decode_local_update_v1(&reachable).map_err(AuthorityError::InvalidUpdate)?;
                (true, apply_remote(&staged_doc, reachable)?)
            }
            None => (false, apply_remote(&staged_doc, incoming)?),
        };

        let staged = Self::hydrated(staged_doc, self.base.clone(), self.next_sheet_id);
        let pending = held_back || has_pending(&staged.doc);
        if let Err(error) = staged.upgrade_schema_from(Some(&self.doc))
            && !pending
        {
            return Err(AuthorityError::InvalidState(error));
        }
        let after = staged.encode_state_as_update_v1();
        let integrated = staged.doc.transact().encode_diff_v1(&before_vector);
        let after_vector = staged.doc.transact().state_vector();
        let state_vector_entries = after_vector.len();
        if pending {
            let mut current = None;
            if integrated.as_slice() != Update::EMPTY_V1
                && let Ok((model, structure)) = staged.strict_materialize()
            {
                let effective = after_vector != before_vector || applied_deletes || {
                    let state = self
                        .strict_materialize()
                        .map_err(AuthorityError::InvalidState)?;
                    let effective = model != state.0 || structure != state.1;
                    current = Some(state);
                    effective
                };
                if effective {
                    return Ok(StagedUpdate {
                        commit_update: integrated.clone(),
                        effective: true,
                        model,
                        pending: true,
                        state_bytes: after.len(),
                        state_vector_entries,
                        structure,
                        update: integrated,
                    });
                }
            }
            let (current_model, current_structure) = match current {
                Some(state) => state,
                None => self
                    .strict_materialize()
                    .map_err(AuthorityError::InvalidState)?,
            };
            return Ok(StagedUpdate {
                commit_update: Update::EMPTY_V1.to_vec(),
                effective: false,
                model: current_model,
                pending: true,
                state_bytes: before.len(),
                state_vector_entries: self.state_vector_entries(),
                structure: current_structure,
                update: Update::EMPTY_V1.to_vec(),
            });
        }
        let (model, structure) = staged
            .strict_materialize()
            .map_err(AuthorityError::InvalidState)?;
        let effective = after_vector != before_vector || applied_deletes || {
            let (current_model, current_structure) = self
                .strict_materialize()
                .map_err(AuthorityError::InvalidState)?;
            model != current_model || structure != current_structure
        };
        Ok(StagedUpdate {
            commit_update: integrated.clone(),
            effective,
            model,
            pending,
            state_bytes: after.len(),
            state_vector_entries,
            structure,
            update: integrated,
        })
    }

    pub(crate) fn stage_local_ops_v1(
        &self,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
    ) -> Result<StagedLocalUpdate, AuthorityError> {
        self.stage_local_ops_from_validated_v1(
            ops,
            origin,
            styles,
            &self.encode_state_as_update_v1(),
            self.projection_valid.load(Ordering::Relaxed),
            true,
        )
    }

    /// [`Self::stage_local_ops_v1`] from this replica's already-encoded state.
    pub(crate) fn stage_local_ops_from_v1(
        &self,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
        baseline: &[u8],
    ) -> Result<StagedLocalUpdate, AuthorityError> {
        self.stage_local_ops_from_validated_v1(ops, origin, styles, baseline, false, true)
    }

    pub(crate) fn stage_standalone_ops_v1(
        &self,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
        baseline: Option<&[u8]>,
    ) -> Result<StagedLocalUpdate, AuthorityError> {
        let projection_valid = baseline.is_none() && self.projection_valid.load(Ordering::Relaxed);
        let encoded;
        let baseline = match baseline {
            Some(baseline) => baseline,
            None => {
                encoded = self.encode_state_as_update_v1();
                &encoded
            }
        };
        self.stage_local_ops_from_validated_v1(
            ops,
            origin,
            styles,
            baseline,
            projection_valid,
            false,
        )
    }

    fn stage_local_ops_from_validated_v1(
        &self,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
        baseline: &[u8],
        projection_valid: bool,
        collaborative: bool,
    ) -> Result<StagedLocalUpdate, AuthorityError> {
        let staged_doc = Doc::with_client_id(self.client_id());
        hydrate_local_doc(&staged_doc, baseline).map_err(AuthorityError::InvalidState)?;
        let mut staged = Self::hydrated(staged_doc, self.base.clone(), self.next_sheet_id);
        staged
            .projection_valid
            .store(projection_valid, Ordering::Relaxed);
        // `apply_ops` already encoded the same diff: the staged doc is the
        // hydrated baseline, so its pre-op state vector is this replica's own.
        let update = staged
            .apply_ops_with_merge_edits(ops, origin, styles, true, collaborative)?
            .unwrap_or_else(|| Update::EMPTY_V1.to_vec());
        let state = staged.encode_state_as_update_v1();
        let state_vector_entries = staged.doc.transact().state_vector().len();
        let (model, structure) = staged
            .materialize_internal(false)
            .map_err(AuthorityError::InvalidState)?;
        Ok(StagedLocalUpdate {
            model,
            state_bytes: state.len(),
            state_vector_entries,
            structure,
            update,
            carried: CarriedState {
                sheet_order: staged.history.undo.pop(),
                base: staged.base,
                next_sheet_id: staged.next_sheet_id,
            },
        })
    }

    /// Adopts a staged local update, returning its bytes unless it was empty. They are decoded
    /// before anything changes, so a failure leaves this replica as it was.
    pub(crate) fn adopt_local_update(
        &mut self,
        staged: StagedLocalUpdate,
        history: LocalHistory,
    ) -> Result<Option<Vec<u8>>, AuthorityError> {
        let StagedLocalUpdate {
            update: bytes,
            carried,
            ..
        } = staged;
        let update = decode_local_update_v1(&bytes).map_err(AuthorityError::InvalidUpdate)?;
        let undo = match history {
            LocalHistory::Undo => Some(
                build_undo_manager(&self.doc, self.undo_stack.clone(), self.redo_stack.clone())
                    .map_err(AuthorityError::InvalidState)?,
            ),
            LocalHistory::SheetOrder | LocalHistory::None => None,
        };
        match undo {
            Some(mut undo) => {
                undo.reset();
                self.doc
                    .transact_mut_with(self.client_id())
                    .apply_update(update)
                    .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))?;
                self.undo_stack = undo.undo_stack().to_vec();
                self.redo_stack = undo.redo_stack().to_vec();
            }
            None => self
                .doc
                .transact_mut_with(UNTRACKED_ORIGIN)
                .apply_update(update)
                .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))?,
        }
        let CarriedState {
            base,
            next_sheet_id,
            sheet_order,
        } = carried;
        self.base = base;
        self.next_sheet_id = next_sheet_id;
        if history == LocalHistory::SheetOrder
            && let Some(entry) = sheet_order
        {
            self.apply_history(HistoryAction::Push(entry));
        }
        Ok((bytes.as_slice() != Update::EMPTY_V1).then_some(bytes))
    }

    /// Whether `update` integrates completely into a fresh replica of `baseline`, this
    /// replica's encoded state.
    pub(crate) fn rehearse_local_update(
        &self,
        baseline: &[u8],
        update: &[u8],
    ) -> Result<bool, AuthorityError> {
        let doc = Doc::with_client_id(self.client_id());
        hydrate_local_doc(&doc, baseline).map_err(AuthorityError::InvalidState)?;
        let update = decode_local_update_v1(update).map_err(AuthorityError::InvalidUpdate)?;
        doc.transact_mut_with(UNTRACKED_ORIGIN)
            .apply_update(update)
            .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))?;
        Ok(!has_pending(&doc))
    }

    /// Whether this replica holds update structs it has not integrated yet.
    pub(crate) fn has_pending_updates(&self) -> bool {
        has_pending(&self.doc)
    }

    pub(crate) fn apply_local_update_v1(
        &mut self,
        update: &[u8],
        origin: SyncOrigin,
    ) -> Result<(), AuthorityError> {
        let update = decode_local_update_v1(update).map_err(AuthorityError::InvalidUpdate)?;
        if origin == SyncOrigin::User {
            let mut undo =
                build_undo_manager(&self.doc, self.undo_stack.clone(), self.redo_stack.clone())
                    .map_err(AuthorityError::InvalidState)?;
            undo.reset();
            self.doc
                .transact_mut_with(self.client_id())
                .apply_update(update)
                .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))?;
            self.undo_stack = undo.undo_stack().to_vec();
            self.redo_stack = undo.redo_stack().to_vec();
            Ok(())
        } else {
            self.doc
                .transact_mut_with(origin.as_str())
                .apply_update(update)
                .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))
        }
    }

    pub(crate) fn apply_staged_update_v1(&self, update: &[u8]) -> Result<(), AuthorityError> {
        let update = decode_local_update_v1(update).map_err(AuthorityError::InvalidUpdate)?;
        self.doc
            .transact_mut_with(REMOTE_ORIGIN)
            .apply_update(update)
            .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))
    }

    pub(crate) fn can_undo(&self) -> bool {
        !self.undo_stack.is_empty()
    }

    pub(crate) fn can_redo(&self) -> bool {
        !self.redo_stack.is_empty()
    }

    pub(crate) fn undo_depth(&self) -> usize {
        self.undo_stack.len()
    }

    pub(crate) fn redo_depth(&self) -> usize {
        self.redo_stack.len()
    }

    /// The document and history as they stand, so a caller that rejects what a
    /// history step produced can put the authority back. The step itself is
    /// only half the change: the caller validates the other half.
    pub(crate) fn checkpoint(&self) -> AuthorityCheckpoint {
        AuthorityCheckpoint {
            state: self.encode_state_as_update_v1(),
            undo_stack: self.undo_stack.clone(),
            redo_stack: self.redo_stack.clone(),
        }
    }

    pub(crate) fn restore(
        &mut self,
        checkpoint: AuthorityCheckpoint,
    ) -> Result<(), AuthorityError> {
        self.rollback(
            &checkpoint.state,
            checkpoint.undo_stack,
            checkpoint.redo_stack,
        )
    }

    pub(crate) fn undo(&mut self) -> Result<Option<HistoryUpdate>, AuthorityError> {
        self.apply_history_change(false)
    }

    pub(crate) fn redo(&mut self) -> Result<Option<HistoryUpdate>, AuthorityError> {
        self.apply_history_change(true)
    }

    /// Moves the document and both stacks. Every failure here leaves that half
    /// done, so the caller holds a checkpoint and puts it back — one checkpoint
    /// for the whole step rather than one here and another there, since this
    /// runs on a keystroke and each costs a copy of the document.
    fn apply_history_change(
        &mut self,
        redo: bool,
    ) -> Result<Option<HistoryUpdate>, AuthorityError> {
        let state_vector = self.doc.transact().state_vector();
        let mut undo =
            build_undo_manager(&self.doc, self.undo_stack.clone(), self.redo_stack.clone())
                .map_err(AuthorityError::InvalidState)?;
        let applied = if redo {
            undo.redo_blocking()
        } else {
            undo.undo_blocking()
        };
        self.undo_stack = undo.undo_stack().to_vec();
        self.redo_stack = undo.redo_stack().to_vec();
        drop(undo);
        if !applied {
            return Ok(None);
        }
        let (model, structure) = self
            .strict_materialize()
            .map_err(AuthorityError::InvalidState)?;
        let update = self.doc.transact().encode_diff_v1(&state_vector);
        Ok(Some(HistoryUpdate {
            model,
            structure,
            update,
        }))
    }

    /// Rebuilds the document from a checkpoint. The GUID is carried over
    /// because a history entry names the document it came from: a fresh one
    /// orphans every restored entry, and the undo manager discards orphans
    /// silently, one whole stack at a time.
    fn rollback(
        &mut self,
        restore: &[u8],
        undo_stack: Vec<StackItem<()>>,
        redo_stack: Vec<StackItem<()>>,
    ) -> Result<(), AuthorityError> {
        let doc = Doc::with_options(Options::with_guid_and_client_id(
            self.doc.guid().clone(),
            self.doc.client_id(),
        ));
        hydrate_local_doc(&doc, restore).map_err(AuthorityError::InvalidState)?;
        self.doc = doc;
        self.snapshot_revision.fetch_add(1, Ordering::Relaxed);
        self.projection_valid = observe_projection(&self.doc, self.snapshot_revision.clone());
        self.undo_stack = undo_stack;
        self.redo_stack = redo_stack;
        Ok(())
    }

    pub(crate) fn clear_history(&mut self) {
        self.history = SheetOrderHistory::default();
    }

    fn strict_materialize(&self) -> Result<(WorkbookModel, WorkbookStructure), String> {
        self.materialize_internal(true)
    }

    #[cfg(test)]
    fn upgrade_schema(&self) -> Result<bool, String> {
        self.upgrade_schema_from(None)
    }

    /// `source` is the replica adopting this state, whose cells hold the
    /// formats a state persisted before format facets keyed only by its
    /// reading of them.
    fn upgrade_schema_from(&self, source: Option<&Doc>) -> Result<bool, String> {
        let version = self.schema_version()?;
        validate_schema_version(version)?;
        self.deduplicate_sheet_order()?;
        let fingerprint = self.state_fingerprint();
        let current = fingerprint.as_deref() == Some(self.base.fingerprint.as_str());
        if version == SCHEMA_VERSION && current {
            return Ok(false);
        }
        let (model, structure) = self.materialize_internal(false)?;
        let features = structure
            .sheet_keys
            .iter()
            .zip(&model.sheets)
            .map(|(key, sheet)| {
                let hyperlinks = serde_json::to_string(&sheet.hyperlinks)
                    .map_err(|error| format!("cannot encode sheet hyperlinks: {error}"))?;
                let charts = serde_json::to_string(&sheet.charts)
                    .map_err(|error| format!("cannot encode sheet charts: {error}"))?;
                Ok((
                    key.clone(),
                    (
                        sheet.freeze_pane,
                        hyperlinks,
                        charts,
                        sheet.col_widths.clone(),
                        sheet.row_heights.clone(),
                    ),
                ))
            })
            .collect::<Result<BTreeMap<_, _>, String>>()?;
        let mut txn = self.doc.transact_mut_with(HYDRATE_ORIGIN);
        if !current && fingerprint.as_deref().is_some_and(released_fingerprint) {
            migrate_pre_facet_formats(
                &mut txn,
                &self.base.styles,
                source,
                fingerprint_reading(fingerprint.as_deref().unwrap_or_default()),
            )?;
        }
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        let keys = sheets.keys(&txn).map(str::to_string).collect::<Vec<_>>();
        for key in keys {
            let sheet = sheets
                .get(&txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("sheet {key} is not a map"))?;
            if let Some((.., col_widths, row_heights)) = features.get(&key) {
                let widths: MapRef = sheet.get_or_init(&mut txn, COL_WIDTHS);
                sync_numbers(&widths, &mut txn, col_widths);
                let heights: MapRef = sheet.get_or_init(&mut txn, ROW_HEIGHTS);
                sync_numbers(&heights, &mut txn, row_heights);
            }
            let (freeze_pane, hyperlinks, charts) = features
                .get(&key)
                .map(|(freeze_pane, hyperlinks, charts, ..)| {
                    (*freeze_pane, hyperlinks.as_str(), charts.as_str())
                })
                .unwrap_or((None, "[]", "[]"));
            if version < FREEZE_PANE_SCHEMA_VERSION {
                sheet.try_update(&mut txn, FREEZE_PANE, freeze_pane_to_any(freeze_pane));
            }
            if version < HYPERLINK_SCHEMA_VERSION {
                sheet.try_update(&mut txn, HYPERLINKS, hyperlinks);
            }
            if version < CHARTS_SCHEMA_VERSION {
                sheet.try_update(&mut txn, CHARTS, charts);
            }
        }
        let meta = txn
            .get_map(META)
            .ok_or_else(|| "missing workbook metadata".to_string())?;
        meta.try_update(&mut txn, BASE_FINGERPRINT, self.base.fingerprint.as_str());
        meta.try_update(&mut txn, "schemaVersion", SCHEMA_VERSION);
        Ok(true)
    }

    fn deduplicate_sheet_order(&self) -> Result<(), String> {
        let mut txn = self.doc.transact_mut_with(HYDRATE_ORIGIN);
        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        let keys = sheet_keys(&order, &txn)?;
        let mut seen = HashSet::with_capacity(keys.len());
        let duplicates = keys
            .iter()
            .enumerate()
            .filter_map(|(index, key)| (!seen.insert(key)).then_some(index))
            .collect::<Vec<_>>();
        for index in duplicates.into_iter().rev() {
            order.remove(&mut txn, yrs_index(index)?);
        }
        Ok(())
    }

    fn schema_version(&self) -> Result<i64, String> {
        let txn = self.doc.transact();
        txn.get_map(META)
            .and_then(|meta| meta.get(&txn, "schemaVersion"))
            .and_then(|value| value.cast::<i64>().ok())
            .ok_or_else(|| "missing schema version".to_string())
    }

    fn has_current_base_fingerprint(&self) -> bool {
        self.state_fingerprint()
            .is_some_and(|fingerprint| fingerprint == self.base.fingerprint)
    }

    fn state_fingerprint(&self) -> Option<String> {
        let txn = self.doc.transact();
        txn.get_map(META)
            .and_then(|meta| meta.get(&txn, BASE_FINGERPRINT))
            .and_then(|value| value.cast::<String>().ok())
    }

    fn materialize_internal(
        &self,
        strict: bool,
    ) -> Result<(WorkbookModel, WorkbookStructure), String> {
        let txn = self.doc.transact();
        if strict {
            require_root_keys(&txn, &[CELL_FORMATS, META, SHEET_ORDER, SHEETS])?;
        }
        let meta = txn
            .get_map(META)
            .ok_or_else(|| "missing workbook metadata".to_string())?;
        if strict {
            require_map_keys(
                &meta,
                &txn,
                &[BASE_FINGERPRINT, "schemaVersion", STRUCTURE_GENERATION],
                "workbook metadata",
            )?;
        }
        let version = meta
            .get(&txn, "schemaVersion")
            .and_then(|value| value.cast::<i64>().ok())
            .ok_or_else(|| "missing schema version".to_string())?;
        validate_schema_version(version)?;
        let fingerprint = meta
            .get(&txn, BASE_FINGERPRINT)
            .and_then(|value| value.cast::<String>().ok())
            .ok_or_else(|| "missing workbook base fingerprint".to_string())?;
        if !self.base.accepts_fingerprint(version, &fingerprint) {
            return Err("workbook base fingerprint does not match shared state".to_string());
        }
        let generation = structure_generation(&meta, &txn)?;
        let cell_formats = txn
            .get_map(CELL_FORMATS)
            .ok_or_else(|| "missing cell format catalog".to_string())?;
        let (styles, style_indices) =
            materialize_cell_formats(&cell_formats, &txn, &self.base.styles)?;

        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        let keys = sheet_keys(&order, &txn)?;
        let mut seen = HashSet::with_capacity(keys.len());
        let mut model = self.base.workbook();
        model.styles = styles;
        let mut plain = None;
        let expected_sheet_keys = sheet_schema_keys(version);
        let optional_sheet_keys = sheet_schema_optional_keys(version);
        for key in keys.iter() {
            if !seen.insert(key.clone()) {
                return Err(format!("duplicate sheet key {key}"));
            }
            let sheet_map = sheets
                .get(&txn, key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("missing sheet {key}"))?;
            if strict {
                require_sheet_keys(
                    &sheet_map,
                    &txn,
                    version,
                    expected_sheet_keys,
                    optional_sheet_keys,
                    &format!("sheet {key}"),
                )?;
            }
            let base_sheet = base_sheet_index(key);
            let freeze_pane = base_sheet
                .and_then(|base| self.base.freeze_panes.get(base))
                .copied()
                .flatten();
            let format = base_sheet
                .and_then(|base| self.base.formats.get(base))
                .copied()
                .unwrap_or_default();
            let col_styles = base_sheet
                .and_then(|base| self.base.col_styles.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let row_styles = base_sheet
                .and_then(|base| self.base.row_styles.get(base))
                .unwrap_or(&EMPTY_ROW_STYLES);
            let hyperlinks = base_sheet
                .and_then(|base| self.base.hyperlinks.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let auto_filter = base_sheet
                .and_then(|base| self.base.auto_filters.get(base))
                .and_then(Option::as_ref);
            let conditional_formats = base_sheet
                .and_then(|base| self.base.conditional_formats.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let charts = base_sheet
                .and_then(|base| self.base.charts.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let hidden_dimensions = base_sheet
                .and_then(|base| self.base.hidden_dimensions.get(base))
                .unwrap_or(&EMPTY_HIDDEN_DIMENSIONS);
            model.sheets.push(materialize_sheet(
                &sheet_map,
                &txn,
                &style_indices,
                &mut || *plain.get_or_insert_with(|| model.styles.plain_xf()),
                version,
                SheetFallbacks {
                    freeze_pane,
                    format,
                    col_styles,
                    row_styles,
                    hyperlinks,
                    charts,
                    hidden_dimensions,
                    conditional_formats,
                    auto_filter,
                },
            )?);
        }
        project_shared_frame_anchors(&mut model.sheets);
        let active = keys.iter().cloned().collect::<BTreeSet<_>>();
        let all_keys = sheets
            .keys(&txn)
            .map(str::to_string)
            .collect::<BTreeSet<_>>();
        let mut shared_types = BTreeMap::new();
        let mut baseline_hyperlinks = BTreeMap::new();
        let mut baseline_merges = BTreeMap::new();
        for key in all_keys {
            let sheet_map = sheets
                .get(&txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("sheet {key} is not a map"))?;
            if strict && !active.contains(&key) {
                require_sheet_keys(
                    &sheet_map,
                    &txn,
                    version,
                    expected_sheet_keys,
                    optional_sheet_keys,
                    &format!("inactive sheet {key}"),
                )?;
                materialize_sheet(
                    &sheet_map,
                    &txn,
                    &style_indices,
                    &mut || 0,
                    version,
                    SheetFallbacks::default(),
                )?;
            }
            let Some(Out::Any(baseline)) = sheet_map.get(&txn, MERGES) else {
                return Err(format!("sheet {key} is missing merges"));
            };
            let link_baseline = match sheet_map.get(&txn, HYPERLINKS) {
                Some(Out::Any(Any::String(json))) => decode_hyperlinks(&json)?,
                None if version < HYPERLINK_SCHEMA_VERSION => base_sheet_index(&key)
                    .and_then(|index| self.base.hyperlinks.get(index))
                    .cloned()
                    .unwrap_or_default(),
                _ => return Err(format!("sheet {key} is missing hyperlinks")),
            };
            baseline_hyperlinks.insert(key.clone(), link_baseline);
            baseline_merges.insert(key.clone(), merges_from_any(&baseline)?);
            shared_types.insert(key, sheet_shared_types(&sheet_map, &txn)?);
        }
        let structure = WorkbookStructure {
            generation,
            sheet_keys: keys,
            sheet_names: model
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect(),
            freeze_panes: model.sheets.iter().map(|sheet| sheet.freeze_pane).collect(),
            hyperlinks: model
                .sheets
                .iter()
                .map(|sheet| sheet.hyperlinks.clone())
                .collect(),
            baseline_hyperlinks,
            conditional_formats: model
                .sheets
                .iter()
                .map(|sheet| sheet.conditional_formats.clone())
                .collect(),
            auto_filters: model
                .sheets
                .iter()
                .map(|sheet| sheet.auto_filter.clone())
                .collect(),
            charts: model
                .sheets
                .iter()
                .map(|sheet| sheet.charts.iter().map(ChartIdentity::of).collect())
                .collect(),
            merges: model
                .sheets
                .iter()
                .map(|sheet| sheet.merges.clone())
                .collect(),
            baseline_merges,
            shared_types,
        };
        self.projection_valid.store(true, Ordering::Relaxed);
        Ok((model, structure))
    }

    /// `authored_styles` holds the pre-batch style of every `SetCell` target.
    fn sync_model(
        &mut self,
        model: &WorkbookModel,
        ops: &[Op],
        origin: SyncOrigin,
        authored_styles: &HashMap<(u32, CellRef), Option<u32>>,
        collaborative: bool,
    ) -> Result<(), String> {
        let current_keys = self.current_sheet_keys()?;
        let (keys, history) =
            self.plan_sheet_keys(&current_keys, ops, model.sheets.len(), origin)?;
        self.validate_sync_state(&current_keys, &keys)?;

        let topology_changed = current_keys != keys;
        // a SetCell after AddSheet targets the post-insert index, so its
        // pre-batch baseline lookup would read the wrong sheet.
        let full_sync = ops.iter().any(requires_full_semantic_sync)
            || (ops.iter().any(|op| matches!(op, Op::AddSheet { .. }))
                && ops.iter().any(|op| matches!(op, Op::SetCell { .. })))
            || (ops
                .iter()
                .any(|op| matches!(op, Op::AddSheet { .. } | Op::RemoveSheet { .. }))
                && ops.iter().any(|op| {
                    matches!(
                        op,
                        Op::MergeCells { .. }
                            | Op::UnmergeCells { .. }
                            | Op::SetHyperlink { .. }
                            | Op::RemoveHyperlink { .. }
                    )
                }));
        let structure_delta = i64::try_from(
            ops.iter()
                .filter(|op| {
                    is_structural_op(op)
                        && (!is_collaborative_local_layout_op(op)
                            || (!collaborative
                                && matches!(
                                    op,
                                    Op::MergeCells { .. }
                                        | Op::UnmergeCells { .. }
                                        | Op::SetHyperlink { .. }
                                        | Op::RemoveHyperlink { .. }
                                )))
                })
                .count(),
        )
        .map_err(|_| "too many structural operations".to_string())?;
        let mut authored_cells = HashSet::new();
        let mut formatted_cells = HashSet::new();
        let mut col_widths = HashSet::new();
        let mut row_heights = HashSet::new();
        let mut freeze_panes = HashSet::new();
        let mut merges = HashSet::new();
        let mut hyperlinks = HashSet::new();
        let mut conditional_formats = HashSet::new();
        if !full_sync {
            let targets = targeted_sheet_keys(&current_keys, &keys, ops)?;
            for (op, target) in ops.iter().zip(targets) {
                match (op, target) {
                    (Op::SetCell { sheet, at, cell }, Some(key)) => {
                        let current_style = authored_styles.get(&(sheet.0, *at)).copied().flatten();
                        if cell.style != current_style {
                            formatted_cells.insert((key.clone(), *at));
                        }
                        authored_cells.insert((key, *at));
                    }
                    (
                        Op::SetColWidth { col, .. }
                        | Op::SetColVisibility { col, .. }
                        | Op::RestoreColDimension { col, .. },
                        Some(key),
                    ) => {
                        col_widths.insert((key, *col));
                    }
                    (
                        Op::SetRowHeight { row, .. }
                        | Op::SetRowVisibility { row, .. }
                        | Op::RestoreRowDimension { row, .. },
                        Some(key),
                    ) => {
                        row_heights.insert((key, *row));
                    }
                    (Op::SetFreezePane { .. }, Some(key)) => {
                        freeze_panes.insert(key);
                    }
                    (Op::MergeCells { .. } | Op::UnmergeCells { .. }, Some(key)) => {
                        merges.insert(key);
                    }
                    (Op::SetHyperlink { .. } | Op::RemoveHyperlink { .. }, Some(key)) => {
                        hyperlinks.insert(key);
                    }
                    (
                        Op::AddConditionalFormat { .. }
                        | Op::RemoveConditionalFormat { .. }
                        | Op::SetConditionalFormats { .. }
                        | Op::SetAutoFilter { .. }
                        | Op::RestoreAutoFilter { .. },
                        Some(key),
                    ) => {
                        conditional_formats.insert(key);
                    }
                    (
                        Op::PatchRangeStyle { range, .. }
                        | Op::SetRangeNumberFormat { range, .. }
                        | Op::ApplyRangeFormat { range, .. },
                        Some(key),
                    ) => {
                        for row in range.start.row..=range.end.row {
                            for col in range.start.col..=range.end.col {
                                formatted_cells.insert((key.clone(), CellRef::new(row, col)));
                            }
                        }
                    }
                    (Op::AddSheet { .. } | Op::RemoveSheet { .. }, None) => {}
                    (_, None) => {}
                    _ => return Err("semantic operation requires a full sync".to_string()),
                }
            }
        }
        let current_key_set = current_keys
            .iter()
            .map(String::as_str)
            .collect::<HashSet<_>>();
        let newly_active = keys
            .iter()
            .filter(|key| !current_key_set.contains(key.as_str()))
            .cloned()
            .collect::<HashSet<_>>();
        if !topology_changed
            && !full_sync
            && structure_delta == 0
            && authored_cells.is_empty()
            && formatted_cells.is_empty()
            && col_widths.is_empty()
            && row_heights.is_empty()
            && freeze_panes.is_empty()
            && merges.is_empty()
            && hyperlinks.is_empty()
            && conditional_formats.is_empty()
        {
            self.apply_history(history);
            return Ok(());
        }

        let mut txn = self.doc.transact_mut_with(origin.as_str());
        let cell_formats = txn
            .get_map(CELL_FORMATS)
            .ok_or_else(|| "missing cell format catalog".to_string())?;
        sync_cell_formats(&cell_formats, &mut txn, &model.styles)?;
        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        if structure_delta != 0 {
            let meta = txn
                .get_map(META)
                .ok_or_else(|| "missing workbook metadata".to_string())?;
            let next = structure_generation(&meta, &txn)?
                .checked_add(structure_delta)
                .ok_or_else(|| "structure generation overflow".to_string())?;
            meta.try_update(&mut txn, STRUCTURE_GENERATION, next);
        }
        if topology_changed {
            patch_sheet_order(&order, &mut txn, &current_keys, &keys)?;
        }
        if full_sync {
            for (key, sheet) in keys.iter().zip(&model.sheets) {
                let sheet_map = sheet_map_for_sync(&sheets, &mut txn, key)?;
                sync_sheet(&sheet_map, &mut txn, sheet, &model.styles)?;
            }
        } else {
            for (key, sheet) in keys.iter().zip(&model.sheets) {
                if newly_active.contains(key) {
                    let sheet_map = sheet_map_for_sync(&sheets, &mut txn, key)?;
                    sync_sheet(&sheet_map, &mut txn, sheet, &model.styles)?;
                }
            }
            let order_cells = |targets: HashSet<(String, CellRef)>| {
                let mut targets = targets.into_iter().collect::<Vec<_>>();
                targets.sort_unstable_by(|(left_key, left), (right_key, right)| {
                    (left_key, left.row, left.col, left.abs_row, left.abs_col).cmp(&(
                        right_key,
                        right.row,
                        right.col,
                        right.abs_row,
                        right.abs_col,
                    ))
                });
                targets
            };
            let mut col_widths = col_widths.into_iter().collect::<Vec<_>>();
            col_widths.sort_unstable();
            let mut row_heights = row_heights.into_iter().collect::<Vec<_>>();
            row_heights.sort_unstable();
            let mut merges = merges.into_iter().collect::<Vec<_>>();
            merges.sort_unstable();
            for (key, at) in order_cells(authored_cells) {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                sync_authored_cell(&sheet_map, &mut txn, sheet_model, at)?;
            }
            for (key, at) in order_cells(formatted_cells) {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                sync_cell_format(&sheet_map, &mut txn, sheet_model, &model.styles, at)?;
            }
            for (key, col) in col_widths {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                let map: MapRef = sheet_map.get_or_init(&mut txn, COL_WIDTHS);
                sync_number(
                    &map,
                    &mut txn,
                    col,
                    sheet_model.col_widths.get(&col).copied(),
                );
            }
            for (key, row) in row_heights {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                let map: MapRef = sheet_map.get_or_init(&mut txn, ROW_HEIGHTS);
                sync_number(
                    &map,
                    &mut txn,
                    row,
                    sheet_model.row_heights.get(&row).copied(),
                );
            }
            for key in freeze_panes {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                sheet_map.try_update(
                    &mut txn,
                    FREEZE_PANE,
                    freeze_pane_to_any(sheet_model.freeze_pane),
                );
            }
            for key in merges {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                if !collaborative {
                    sheet_map.try_update(&mut txn, MERGES, merges_to_any(&sheet_model.merges));
                    clear_merge_edits(&sheet_map, &mut txn);
                    continue;
                }
                let Some(Out::Any(baseline)) = sheet_map.get(&txn, MERGES) else {
                    return Err(format!("sheet {key} is missing merges"));
                };
                let baseline_keys = merges_from_any(&baseline)?
                    .iter()
                    .map(CellRange::to_a1)
                    .collect::<BTreeSet<_>>();
                let current_keys = sheet_model
                    .merges
                    .iter()
                    .map(CellRange::to_a1)
                    .collect::<BTreeSet<_>>();
                let mut edit_keys = sheet_map
                    .keys(&txn)
                    .filter_map(|key| key.strip_prefix(MERGE_PREFIX).map(str::to_string))
                    .collect::<BTreeSet<_>>();
                edit_keys.extend(baseline_keys.iter().cloned());
                edit_keys.extend(current_keys.iter().cloned());
                for range_key in edit_keys {
                    let edit_key = format!("{MERGE_PREFIX}{range_key}");
                    match (
                        baseline_keys.contains(&range_key),
                        current_keys.contains(&range_key),
                    ) {
                        (true, true) | (false, false) => {
                            sheet_map.remove(&mut txn, &edit_key);
                        }
                        (_, present) => {
                            sheet_map.try_update(&mut txn, edit_key, present);
                        }
                    }
                }
            }
            for key in hyperlinks {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                sync_hyperlinks(&sheet_map, &mut txn, sheet_model, collaborative)?;
            }
            for key in conditional_formats {
                let (sheet_map, sheet_model) =
                    sheet_parts_by_key(&sheets, &txn, &keys, model, &key)?;
                let json = serde_json::to_string(&sheet_model.conditional_formats)
                    .map_err(|error| format!("cannot encode sheet conditional formats: {error}"))?;
                sheet_map.try_update(&mut txn, CONDITIONAL_FORMATS, json);
                let json = serde_json::to_string(&sheet_model.auto_filter)
                    .map_err(|error| format!("cannot encode sheet auto filter: {error}"))?;
                sheet_map.try_update(&mut txn, AUTO_FILTER, json);
            }
        }
        drop(txn);
        self.apply_history(history);
        Ok(())
    }

    fn current_sheet_keys(&self) -> Result<Vec<String>, String> {
        let txn = self.doc.transact();
        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        sheet_keys(&order, &txn)
    }

    fn validate_sync_state(&self, current: &[String], desired: &[String]) -> Result<(), String> {
        let txn = self.doc.transact();
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        for key in current {
            match sheets.get(&txn, key) {
                Some(Out::YMap(_)) => {}
                Some(_) => return Err(format!("sheet {key} is not a map")),
                None => return Err(format!("missing sheet {key}")),
            }
        }
        for key in desired {
            if let Some(value) = sheets.get(&txn, key)
                && !matches!(value, Out::YMap(_))
            {
                return Err(format!("sheet {key} is not a map"));
            }
        }
        Ok(())
    }

    fn plan_sheet_keys(
        &mut self,
        current: &[String],
        ops: &[Op],
        final_len: usize,
        origin: SyncOrigin,
    ) -> Result<(Vec<String>, HistoryAction), String> {
        match origin {
            SyncOrigin::User | SyncOrigin::Agent => {
                let keys = self.reconcile_sheet_keys(current.to_vec(), ops, final_len)?;
                let entry = SheetOrderEntry {
                    before: current.to_vec(),
                    after: keys.clone(),
                };
                Ok((keys, HistoryAction::Push(entry)))
            }
            SyncOrigin::Undo => {
                let entry = self
                    .history
                    .undo
                    .last()
                    .cloned()
                    .ok_or_else(|| "sheet-order undo history is empty".to_string())?;
                if entry.after != current {
                    return Err("sheet-order undo history does not match current state".to_string());
                }
                if entry.before.len() != final_len {
                    return Err("sheet-order undo result does not match workbook".to_string());
                }
                Ok((entry.before.clone(), HistoryAction::Undo(entry)))
            }
            SyncOrigin::Redo => {
                let entry = self
                    .history
                    .redo
                    .last()
                    .cloned()
                    .ok_or_else(|| "sheet-order redo history is empty".to_string())?;
                if entry.before != current {
                    return Err("sheet-order redo history does not match current state".to_string());
                }
                if entry.after.len() != final_len {
                    return Err("sheet-order redo result does not match workbook".to_string());
                }
                Ok((entry.after.clone(), HistoryAction::Redo(entry)))
            }
        }
    }

    fn reconcile_sheet_keys(
        &mut self,
        mut keys: Vec<String>,
        ops: &[Op],
        final_len: usize,
    ) -> Result<Vec<String>, String> {
        for op in ops {
            match op {
                Op::AddSheet { index, .. } => {
                    if *index > keys.len() {
                        return Err(format!("sheet insertion index {index} is out of range"));
                    }
                    let key = self.allocate_sheet_key();
                    keys.insert(*index, key);
                }
                Op::MoveSheet { from, to } => {
                    let key = keys.remove(*from);
                    keys.insert(*to, key);
                }
                Op::RemoveSheet { index } => {
                    if *index >= keys.len() {
                        return Err(format!("sheet removal index {index} is out of range"));
                    }
                    keys.remove(*index);
                }
                _ => {}
            }
        }
        if keys.len() != final_len {
            return Err("sheet order does not match workbook projection".to_string());
        }
        Ok(keys)
    }

    fn apply_history(&mut self, action: HistoryAction) {
        match action {
            HistoryAction::Push(entry) => {
                self.history.undo.push(entry);
                self.history.redo.clear();
            }
            HistoryAction::Undo(entry) => {
                self.history.undo.pop();
                self.history.redo.push(entry);
            }
            HistoryAction::Redo(entry) => {
                self.history.redo.pop();
                self.history.undo.push(entry);
            }
        }
    }

    fn allocate_sheet_key(&mut self) -> String {
        let key = format!("replica:{}:{}", self.client_id(), self.next_sheet_id);
        self.next_sheet_id += 1;
        key
    }
}

fn observe_projection(doc: &Doc, revision: Arc<AtomicU64>) -> Arc<AtomicBool> {
    let valid = Arc::new(AtomicBool::new(false));
    let observed = valid.clone();
    doc.observe_after_transaction_with("projection", move |transaction| {
        observed.store(false, Ordering::Relaxed);
        if !transaction.insert_set().is_empty() || !transaction.delete_set().is_empty() {
            revision.fetch_add(1, Ordering::Relaxed);
        }
    })
    .expect("authority document has no active transaction");
    valid
}

/// The base-model sheet a stable key names. Bootstrap mints `sheet:N` from the
/// base order and a replica mints `replica:...`, so a legacy state whose sheets
/// were reordered still reads its own fallback features rather than whichever
/// sheet now sits at that position.
fn base_sheet_index(key: &str) -> Option<usize> {
    key.strip_prefix("sheet:")?.parse().ok()
}

pub(crate) fn is_structural_op(op: &Op) -> bool {
    matches!(
        op,
        Op::InsertRows { .. }
            | Op::DeleteRows { .. }
            | Op::InsertCols { .. }
            | Op::DeleteCols { .. }
            | Op::SetFreezePane { .. }
            | Op::SetHyperlinks { .. }
            | Op::SetHyperlink { .. }
            | Op::RemoveHyperlink { .. }
            | Op::AddConditionalFormat { .. }
            | Op::RemoveConditionalFormat { .. }
            | Op::SetConditionalFormats { .. }
            | Op::SetAutoFilter { .. }
            | Op::RestoreAutoFilter { .. }
            | Op::RestoreColStyles { .. }
            | Op::RestoreRowStyles { .. }
            | Op::MergeCells { .. }
            | Op::UnmergeCells { .. }
            | Op::AddSheet { .. }
            | Op::RemoveSheet { .. }
            | Op::MoveSheet { .. }
            | Op::RenameSheet { .. }
            | Op::RestoreSheet { .. }
            | Op::SetCharts { .. }
            | Op::SetDefinedName { .. }
            | Op::DeleteDefinedName { .. }
            | Op::SetDefinedNames { .. }
    )
}

pub(crate) fn is_collaborative_local_layout_op(op: &Op) -> bool {
    matches!(
        op,
        Op::SetFreezePane { .. }
            | Op::MergeCells { .. }
            | Op::UnmergeCells { .. }
            | Op::SetHyperlink { .. }
            | Op::RemoveHyperlink { .. }
    )
}

#[cfg(test)]
fn seed(
    doc: &Doc,
    base: &WorkbookBase,
    model: &WorkbookModel,
    keys: &[String],
) -> Result<(), String> {
    seed_with_format_entries(doc, base, model, keys, style_entry)
}

type FormatEntryWriter = fn(&Stylesheet, u32) -> Result<(String, String), String>;

#[cfg(test)]
fn seed_with_format_entries(
    doc: &Doc,
    base: &WorkbookBase,
    model: &WorkbookModel,
    keys: &[String],
    entry: FormatEntryWriter,
) -> Result<(), String> {
    seed_in_chunks(doc, base, model, keys, entry, usize::MAX, |_| Ok(()))
}

/// Cell entries one bootstrap transaction holds at most when an open seeds in
/// chunks, so no single local apply grows with the workbook.
const SEED_CHUNK: usize = 200_000;

/// Makes the insertions of one bootstrap transaction, in the same order, but
/// commits after every `chunk` cell entries and hands `flush` the update of
/// each committed transaction, the last one included.
fn seed_in_chunks(
    doc: &Doc,
    base: &WorkbookBase,
    model: &WorkbookModel,
    keys: &[String],
    entry: FormatEntryWriter,
    chunk: usize,
    mut flush: impl FnMut(Vec<u8>) -> Result<(), String>,
) -> Result<(), String> {
    let mut txn = doc.transact_mut_with(BOOTSTRAP_ORIGIN);
    let cell_formats = txn.get_or_insert_map(CELL_FORMATS);
    let (key, payload) = cell_format_entry(&CellFormat::default())?;
    cell_formats.try_update(&mut txn, key, payload);
    let mut style_keys = Vec::with_capacity(model.styles.cell_xfs.len());
    for index in 0..model.styles.cell_xfs.len() {
        let index =
            u32::try_from(index).map_err(|_| "cell format table is too large".to_string())?;
        let (key, payload) = entry(&model.styles, index)?;
        style_keys.push(key.clone());
        cell_formats.try_update(&mut txn, key, payload);
    }
    let meta = txn.get_or_insert_map(META);
    meta.insert(&mut txn, BASE_FINGERPRINT, base.fingerprint.as_str());
    meta.insert(&mut txn, "schemaVersion", SCHEMA_VERSION);
    meta.insert(&mut txn, STRUCTURE_GENERATION, 0_i64);
    let order = txn.get_or_insert_array(SHEET_ORDER);
    order.insert_range(&mut txn, 0, keys.iter().cloned());
    let sheets = txn.get_or_insert_map(SHEETS);
    let style_key = |style: u32| {
        style_keys
            .get(style as usize)
            .cloned()
            .ok_or_else(|| format!("cell style index {style} is out of range"))
    };
    let mut room = chunk;
    for (key, sheet) in keys.iter().zip(&model.sheets) {
        let sheet_map = sheets.insert(&mut txn, key.as_str(), MapPrelim::default());
        let maps = sync_sheet_frame(&sheet_map, &mut txn, sheet)?;
        for (part, map) in maps.iter().enumerate() {
            let entries: Box<dyn Iterator<Item = (String, Any)> + '_> = match part {
                0 => Box::new(number_entries(&sheet.col_widths)),
                1 => Box::new(desired_contents(sheet).into_iter()),
                2 => Box::new(number_entries(&sheet.row_heights)),
                _ => Box::new(desired_styles(sheet, style_key)?.into_iter()),
            };
            for (key, value) in entries {
                map.try_update(&mut txn, key, value);
                xlsx_model::progress::tick();
                room -= 1;
                if room == 0 {
                    let update = txn.encode_update_v1();
                    drop(txn);
                    flush(update)?;
                    txn = doc.transact_mut_with(BOOTSTRAP_ORIGIN);
                    room = chunk;
                }
            }
        }
    }
    let update = txn.encode_update_v1();
    drop(txn);
    flush(update)
}

fn hydrate_doc(doc: &Doc, update: &[u8]) -> Result<(), String> {
    let update = decode_update_v1(update)?;
    doc.transact_mut_with(HYDRATE_ORIGIN)
        .apply_update(update)
        .map_err(|error| error.to_string())
}

fn hydrate_local_doc(doc: &Doc, update: &[u8]) -> Result<(), String> {
    let update = decode_local_update_v1(update)?;
    doc.transact_mut_with(HYDRATE_ORIGIN)
        .apply_update(update)
        .map_err(|error| error.to_string())
}

#[cfg(test)]
pub(crate) fn hydrate_snapshot_part(doc: &Doc, update: &[u8]) -> Result<(), String> {
    hydrate_local_doc(doc, update)
}

fn apply_remote(doc: &Doc, update: Update) -> Result<bool, AuthorityError> {
    let mut txn = doc.transact_mut_with(REMOTE_ORIGIN);
    txn.apply_update(update)
        .map_err(|error| AuthorityError::InvalidUpdate(error.to_string()))?;
    Ok(!txn.delete_set().is_empty())
}

/// Re-encodes `bytes` without each client's blocks from the first one past
/// what `known` and the update's own earlier blocks reach, as Yjs holds them,
/// keeping the whole delete set; `None` when every block is reachable.
fn hold_back_unreachable_v1(
    bytes: &[u8],
    layout: &UpdateLayout,
    known: &StateVector,
) -> Option<Vec<u8>> {
    let mut frontiers = HashMap::new();
    let kept = layout
        .clients
        .iter()
        .map(|section| {
            let frontier = frontiers
                .entry(section.client)
                .or_insert_with(|| Some(known.get(&section.client)));
            let mut kept = 0;
            for (index, block) in section.blocks.iter().enumerate() {
                match *frontier {
                    Some(_) if block.skip => {}
                    Some(clock) if block.clock <= clock => {
                        *frontier = Some(clock.max(block.end_clock));
                        kept = index + 1;
                    }
                    _ => {
                        *frontier = None;
                        break;
                    }
                }
            }
            kept
        })
        .collect::<Vec<_>>();
    let held = layout
        .clients
        .iter()
        .zip(&kept)
        .any(|(section, &kept)| section.blocks[kept..].iter().any(|block| !block.skip));
    if !held {
        return None;
    }
    let mut encoder = EncoderV1::new();
    encoder.write_var(kept.iter().filter(|&&kept| kept > 0).count() as u32);
    for (section, &kept) in layout.clients.iter().zip(&kept) {
        if kept == 0 {
            continue;
        }
        encoder.write_var(kept as u32);
        encoder.write_client(section.client);
        encoder.write_var(section.blocks[0].clock);
        encoder.write_all(&bytes[section.blocks[0].start..section.blocks[kept - 1].end]);
    }
    encoder.write_all(&bytes[layout.deletes..]);
    Some(encoder.to_vec())
}

fn has_pending(doc: &Doc) -> bool {
    let txn = doc.transact();
    txn.store().pending_update().is_some() || txn.store().pending_ds().is_some()
}

fn build_undo_manager(
    doc: &Doc,
    undo_stack: Vec<StackItem<()>>,
    redo_stack: Vec<StackItem<()>>,
) -> Result<UndoManager<()>, String> {
    let txn = doc.transact();
    let sheets = txn
        .get_map(SHEETS)
        .ok_or_else(|| "missing sheet map".to_string())?;
    drop(txn);
    let options = UndoOptions {
        capture_timeout_millis: UNDO_CAPTURE_TIMEOUT_MS,
        tracked_origins: HashSet::from([Origin::from(doc.client_id().get())]),
        capture_transaction: None,
        timestamp: undo_clock(),
        init_undo_stack: undo_stack,
        init_redo_stack: redo_stack,
    };
    let mut undo = UndoManager::with_options(options);
    undo.expand_scope(doc, &sheets);
    Ok(undo)
}

fn undo_clock() -> Arc<dyn Clock> {
    #[cfg(not(all(target_family = "wasm", target_os = "unknown")))]
    {
        Arc::new(|| {
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|elapsed| elapsed.as_millis() as u64)
                .unwrap_or_default()
        })
    }
    #[cfg(all(target_family = "wasm", target_os = "unknown"))]
    {
        use std::sync::atomic::{AtomicU64, Ordering};
        let ticks = AtomicU64::new(0);
        Arc::new(move || ticks.fetch_add(UNDO_CAPTURE_TIMEOUT_MS + 1, Ordering::Relaxed))
    }
}

fn decode_state_vector_v1(bytes: &[u8]) -> Result<StateVector, String> {
    validate_state_vector_entry_count(bytes)?;
    let mut decoder = CheckedDecoderV1::new(bytes);
    let entries = decoder
        .read_var_u32_checked()
        .map_err(|error| error.to_string())?;
    let mut state_vector = StateVector::default();
    for _ in 0..entries {
        let client = decoder
            .read_var_u64_checked()
            .map_err(|error| error.to_string())?;
        let client = checked_client_id(client).map_err(|error| error.to_string())?;
        if state_vector.contains_client(&client) {
            return Err("state vector contains a duplicate client".to_string());
        }
        let clock = decoder
            .read_var_u32_checked()
            .map_err(|error| error.to_string())?;
        checked_clock(clock).map_err(|error| error.to_string())?;
        state_vector.set_max(client, clock);
    }
    if !decoder
        .read_to_end()
        .map_err(|error| error.to_string())?
        .is_empty()
    {
        return Err("state vector contains trailing bytes".to_string());
    }
    Ok(state_vector)
}

fn decode_update_v1(bytes: &[u8]) -> Result<Update, String> {
    decode_checked_update(&mut CheckedDecoderV1::new_update(bytes))
}

/// Decodes an update along with where each block sits; `local` bounds counts
/// by the encoded size, for bytes merged from already-validated updates.
fn decode_update_layout_v1(bytes: &[u8], local: bool) -> Result<(Update, UpdateLayout), String> {
    let mut decoder = CheckedDecoderV1::new_update(bytes);
    if local {
        decoder.max_blocks = bytes.len();
        decoder.max_values = bytes.len();
        decoder.max_delete_ranges = bytes.len();
    }
    decoder.layout = Some(UpdateLayout::default());
    let update = decode_checked_update(&mut decoder)?;
    Ok((update, decoder.layout.unwrap_or_default()))
}

/// Decodes authority-generated bytes with counts bounded by their encoded size.
fn decode_local_update_v1(bytes: &[u8]) -> Result<Update, String> {
    let mut decoder = CheckedDecoderV1::new_update(bytes);
    decoder.max_blocks = bytes.len();
    decoder.max_values = bytes.len();
    decoder.max_delete_ranges = bytes.len();
    decode_checked_update(&mut decoder)
}

fn decode_checked_update(decoder: &mut CheckedDecoderV1<'_>) -> Result<Update, String> {
    catch_unwind(AssertUnwindSafe(|| {
        let update = Update::decode(&mut *decoder).map_err(|error| error.to_string())?;
        if !decoder
            .read_to_end()
            .map_err(|error| error.to_string())?
            .is_empty()
        {
            return Err("update contains trailing bytes".to_string());
        }
        Ok(update)
    }))
    .map_err(|_| "update decoder panicked".to_string())?
}

#[derive(Clone, Copy)]
enum CaptureKind {
    BlockClock,
    BlockCount,
    ClientCount,
    DeleteClient,
    DeleteClientCount,
    DeleteRangeCount,
    SkipLength,
}

struct VarCapture {
    kind: CaptureKind,
    bytes: u8,
    shift: u32,
    value: u64,
}

impl VarCapture {
    fn new(kind: CaptureKind) -> Self {
        Self {
            kind,
            bytes: 0,
            shift: 0,
            value: 0,
        }
    }
}

#[derive(Clone, Copy)]
enum PendingLength {
    Any,
    Deleted,
    Gc,
    Type,
}

#[derive(Default)]
struct UpdateLayout {
    clients: Vec<ClientLayout>,
    deletes: usize,
}

struct ClientLayout {
    client: ClientID,
    blocks: Vec<BlockLayout>,
}

struct BlockLayout {
    start: usize,
    end: usize,
    clock: u32,
    end_clock: u32,
    skip: bool,
}

struct CheckedDecoderV1<'a> {
    inner: DecoderV1<'a>,
    len: usize,
    remaining: usize,
    layout: Option<UpdateLayout>,
    max_blocks: usize,
    max_values: usize,
    max_delete_ranges: usize,
    clock: Option<u32>,
    delete_clock: Option<u32>,
    capture: Option<VarCapture>,
    pending_length: Option<PendingLength>,
    update_mode: bool,
    update_clients_remaining: u32,
    blocks_remaining: u32,
    total_blocks: usize,
    any_items_remaining: u32,
    any_block_len: u32,
    declared_any_items: usize,
    decoded_any_values: usize,
    delete_clients_remaining: u32,
    delete_ranges_remaining: u32,
    total_delete_ranges: usize,
}

impl<'a> CheckedDecoderV1<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self {
            inner: DecoderV1::from(bytes),
            len: bytes.len(),
            remaining: bytes.len(),
            layout: None,
            max_blocks: MAX_UPDATE_BLOCKS,
            max_values: MAX_UPDATE_VALUES,
            max_delete_ranges: MAX_UPDATE_DELETE_RANGES,
            clock: None,
            delete_clock: None,
            capture: None,
            pending_length: None,
            update_mode: false,
            update_clients_remaining: 0,
            blocks_remaining: 0,
            total_blocks: 0,
            any_items_remaining: 0,
            any_block_len: 0,
            declared_any_items: 0,
            decoded_any_values: 0,
            delete_clients_remaining: 0,
            delete_ranges_remaining: 0,
            total_delete_ranges: 0,
        }
    }

    fn new_update(bytes: &'a [u8]) -> Self {
        let mut decoder = Self::new(bytes);
        decoder.update_mode = true;
        decoder.capture = Some(VarCapture::new(CaptureKind::ClientCount));
        decoder
    }

    fn offset(&self) -> usize {
        self.len - self.remaining
    }

    fn begin_capture(&mut self, kind: CaptureKind) -> Result<(), DecodeError> {
        if self.capture.is_some() {
            return Err(decode_error("overlapping update varints"));
        }
        let offset = self.offset();
        if let CaptureKind::DeleteClientCount = kind
            && let Some(layout) = &mut self.layout
        {
            layout.deletes = offset;
        }
        self.capture = Some(VarCapture::new(kind));
        Ok(())
    }

    fn feed_capture(&mut self, byte: u8) -> Result<(), DecodeError> {
        let Some(mut capture) = self.capture.take() else {
            return Ok(());
        };
        capture.bytes = capture
            .bytes
            .checked_add(1)
            .ok_or(DecodeError::InvalidVarInt)?;
        let payload = u64::from(byte & 0x7f);
        if capture.shift >= 64 || payload > (u64::MAX >> capture.shift) {
            return Err(DecodeError::InvalidVarInt);
        }
        capture.value |= payload << capture.shift;
        if byte & 0x80 != 0 {
            if capture.shift >= 63 {
                return Err(DecodeError::InvalidVarInt);
            }
            capture.shift += 7;
            self.capture = Some(capture);
            return Ok(());
        }
        if capture.bytes > 1 && payload == 0 {
            return Err(DecodeError::InvalidVarInt);
        }

        match capture.kind {
            CaptureKind::BlockClock => {
                let clock = u32::try_from(capture.value).map_err(|_| DecodeError::InvalidVarInt)?;
                checked_clock(clock)?;
                self.clock = Some(clock);
            }
            CaptureKind::BlockCount => {
                let blocks =
                    u32::try_from(capture.value).map_err(|_| DecodeError::InvalidVarInt)?;
                if self.update_clients_remaining == 0 || blocks == 0 {
                    return Err(decode_error("update contains an empty client block set"));
                }
                let blocks = blocks as usize;
                self.total_blocks = self
                    .total_blocks
                    .checked_add(blocks)
                    .ok_or_else(|| decode_error("update block count overflow"))?;
                if self.total_blocks > self.max_blocks || blocks > self.remaining {
                    return Err(decode_error("update contains too many blocks"));
                }
                self.blocks_remaining = blocks as u32;
            }
            CaptureKind::ClientCount => {
                let clients =
                    u32::try_from(capture.value).map_err(|_| DecodeError::InvalidVarInt)?;
                if clients > MAX_STATE_VECTOR_ENTRIES {
                    return Err(decode_error("update contains too many clients"));
                }
                self.update_clients_remaining = clients;
                self.begin_capture(if clients == 0 {
                    CaptureKind::DeleteClientCount
                } else {
                    CaptureKind::BlockCount
                })?;
            }
            CaptureKind::DeleteClient => {
                checked_client_id(capture.value)?;
                self.begin_capture(CaptureKind::DeleteRangeCount)?;
            }
            CaptureKind::DeleteClientCount => {
                let clients =
                    u32::try_from(capture.value).map_err(|_| DecodeError::InvalidVarInt)?;
                if clients > MAX_STATE_VECTOR_ENTRIES {
                    return Err(decode_error("update delete set contains too many clients"));
                }
                self.delete_clients_remaining = clients;
            }
            CaptureKind::DeleteRangeCount => {
                let ranges =
                    u32::try_from(capture.value).map_err(|_| DecodeError::InvalidVarInt)?;
                let total = self
                    .total_delete_ranges
                    .checked_add(ranges as usize)
                    .ok_or_else(|| decode_error("update delete range count overflow"))?;
                if total > self.max_delete_ranges || ranges as usize > self.remaining {
                    return Err(decode_error("update contains too many delete ranges"));
                }
                self.total_delete_ranges = total;
                self.delete_ranges_remaining = ranges;
                if ranges == 0 {
                    self.finish_delete_client()?;
                }
            }
            CaptureKind::SkipLength => {
                let len = u32::try_from(capture.value).map_err(|_| DecodeError::InvalidVarInt)?;
                if len == 0 {
                    return Err(decode_error("update contains an empty skip block"));
                }
                self.advance_clock(len)?;
                self.finish_block()?;
            }
        }
        Ok(())
    }

    fn advance_clock(&mut self, len: u32) -> Result<(), DecodeError> {
        let clock = self
            .clock
            .ok_or_else(|| decode_error("update block is missing its initial clock"))?;
        let next = clock
            .checked_add(len)
            .ok_or_else(|| decode_error("update block clock overflows u32"))?;
        checked_clock(next)?;
        self.clock = Some(next);
        Ok(())
    }

    fn finish_block(&mut self) -> Result<(), DecodeError> {
        if self.blocks_remaining == 0 {
            return Err(decode_error("update contains more blocks than declared"));
        }
        let (end, end_clock) = (self.offset(), self.clock.unwrap_or_default());
        if let Some(block) = self
            .layout
            .as_mut()
            .and_then(|layout| layout.clients.last_mut())
            .and_then(|section| section.blocks.last_mut())
        {
            block.end = end;
            block.end_clock = end_clock;
        }
        self.blocks_remaining -= 1;
        if self.blocks_remaining == 0 {
            self.clock = None;
            self.update_clients_remaining = self
                .update_clients_remaining
                .checked_sub(1)
                .ok_or_else(|| decode_error("update client count underflow"))?;
            self.begin_capture(if self.update_clients_remaining == 0 {
                CaptureKind::DeleteClientCount
            } else {
                CaptureKind::BlockCount
            })?;
        }
        Ok(())
    }

    fn finish_delete_client(&mut self) -> Result<(), DecodeError> {
        self.delete_clients_remaining = self
            .delete_clients_remaining
            .checked_sub(1)
            .ok_or_else(|| decode_error("update delete client count underflow"))?;
        Ok(())
    }

    fn read_var_u64_checked(&mut self) -> Result<u64, DecodeError> {
        if self.capture.is_some() {
            return Err(decode_error(
                "unexpected checked varint during update framing",
            ));
        }
        let mut value = 0_u64;
        let mut shift = 0_u32;
        let mut bytes = 0_u8;
        loop {
            let byte = self.read_u8()?;
            bytes += 1;
            let payload = u64::from(byte & 0x7f);
            if shift >= 64 || payload > (u64::MAX >> shift) {
                return Err(DecodeError::InvalidVarInt);
            }
            value |= payload << shift;
            if byte & 0x80 == 0 {
                if bytes > 1 && payload == 0 {
                    return Err(DecodeError::InvalidVarInt);
                }
                return Ok(value);
            }
            if shift >= 63 {
                return Err(DecodeError::InvalidVarInt);
            }
            shift += 7;
        }
    }

    fn read_var_u32_checked(&mut self) -> Result<u32, DecodeError> {
        u32::try_from(self.read_var_u64_checked()?).map_err(|_| DecodeError::InvalidVarInt)
    }

    fn read_var_usize_checked(&mut self) -> Result<usize, DecodeError> {
        usize::try_from(self.read_var_u64_checked()?).map_err(|_| DecodeError::InvalidVarInt)
    }

    fn read_var_i64_checked(&mut self) -> Result<i64, DecodeError> {
        let first = self.read_u8()?;
        let negative = first & 0x40 != 0;
        let mut magnitude = u64::from(first & 0x3f);
        let mut byte = first;
        let mut shift = 6_u32;
        while byte & 0x80 != 0 {
            byte = self.read_u8()?;
            let payload = u64::from(byte & 0x7f);
            if shift >= 63 || payload > (i64::MAX as u64 >> shift) {
                return Err(DecodeError::InvalidVarInt);
            }
            magnitude |= payload << shift;
            if byte & 0x80 == 0 {
                if payload == 0 {
                    return Err(DecodeError::InvalidVarInt);
                }
                break;
            }
            shift += 7;
        }
        if negative && magnitude == 0 {
            return Err(DecodeError::InvalidVarInt);
        }
        let value = i64::try_from(magnitude).map_err(|_| DecodeError::InvalidVarInt)?;
        Ok(if negative { -value } else { value })
    }

    fn read_id(&mut self) -> Result<ID, DecodeError> {
        let client = checked_client_id(self.read_var_u64_checked()?)?;
        let clock = self.read_var_u32_checked()?;
        checked_clock(clock)?;
        Ok(ID::new(client, clock))
    }

    fn decode_any(&mut self, depth: u8) -> Result<Any, DecodeError> {
        if depth >= 64 {
            return Err(decode_error("update value nesting exceeds 64 levels"));
        }
        self.decoded_any_values = self
            .decoded_any_values
            .checked_add(1)
            .ok_or_else(|| decode_error("update value count overflow"))?;
        if self.decoded_any_values > self.max_values {
            return Err(decode_error("update contains too many values"));
        }
        Ok(match self.read_u8()? {
            127 => Any::Undefined,
            126 => Any::Null,
            125 => Any::Number(self.read_var_i64_checked()? as f64),
            124 => Any::Number(self.read_f32()? as f64),
            123 => Any::Number(self.read_f64()?),
            122 => Any::BigInt(self.read_i64()?),
            121 => Any::Bool(false),
            120 => Any::Bool(true),
            119 => Any::String(Arc::from(self.read_string()?)),
            118 => {
                let len = self.read_var_usize_checked()?;
                if len > self.remaining
                    || len > self.max_values.saturating_sub(self.decoded_any_values)
                {
                    return Err(decode_error("update map length exceeds its payload"));
                }
                let mut map = HashMap::new();
                map.try_reserve(len)?;
                for _ in 0..len {
                    let key = self.read_string()?.to_owned();
                    map.insert(key, self.decode_any(depth + 1)?);
                }
                Any::Map(Arc::new(map))
            }
            117 => {
                let len = self.read_var_usize_checked()?;
                if len > self.remaining
                    || len > self.max_values.saturating_sub(self.decoded_any_values)
                {
                    return Err(decode_error("update array length exceeds its payload"));
                }
                let mut values = Vec::new();
                values.try_reserve(len)?;
                for _ in 0..len {
                    values.push(self.decode_any(depth + 1)?);
                }
                Any::Array(Arc::from(values))
            }
            116 => {
                let len = self.read_var_u32_checked()? as usize;
                Any::Buffer(Arc::from(self.read_exact(len)?))
            }
            _ => return Err(DecodeError::UnexpectedValue),
        })
    }
}

impl Read for CheckedDecoderV1<'_> {
    fn read_exact(&mut self, len: usize) -> Result<&[u8], DecodeError> {
        if len > self.remaining {
            return Err(DecodeError::EndOfBuffer(len));
        }
        let bytes = self.inner.read_exact(len)?;
        self.remaining -= len;
        Ok(bytes)
    }

    fn read_u8(&mut self) -> Result<u8, DecodeError> {
        if self.remaining == 0 {
            return Err(DecodeError::EndOfBuffer(1));
        }
        let byte = self.inner.read_u8()?;
        self.remaining -= 1;
        self.feed_capture(byte)?;
        Ok(byte)
    }

    fn read_string(&mut self) -> Result<&str, DecodeError> {
        let len = self.read_var_u32_checked()? as usize;
        let bytes = self.read_exact(len)?;
        std::str::from_utf8(bytes).map_err(|error| decode_error(error.to_string()))
    }
}

impl Decoder for CheckedDecoderV1<'_> {
    fn reset_ds_cur_val(&mut self) {
        self.delete_clock = None;
        if self.capture.is_none()
            && self.delete_clients_remaining > 0
            && self.delete_ranges_remaining == 0
        {
            self.capture = Some(VarCapture::new(CaptureKind::DeleteClient));
        }
    }

    fn read_ds_clock(&mut self) -> Result<u32, DecodeError> {
        if self.delete_ranges_remaining == 0 {
            return Err(decode_error("update contains an undeclared delete range"));
        }
        let clock = self.read_var_u32_checked()?;
        checked_clock(clock)?;
        self.delete_clock = Some(clock);
        Ok(clock)
    }

    fn read_ds_len(&mut self) -> Result<u32, DecodeError> {
        let len = self.read_var_u32_checked()?;
        if len == 0 {
            return Err(decode_error("update contains an empty delete range"));
        }
        let end = self
            .delete_clock
            .ok_or_else(|| decode_error("delete range is missing its clock"))?
            .checked_add(len)
            .ok_or_else(|| decode_error("delete range clock overflows u32"))?;
        checked_clock(end)?;
        self.delete_ranges_remaining -= 1;
        if self.delete_ranges_remaining == 0 {
            self.finish_delete_client()?;
        }
        Ok(len)
    }

    fn read_left_id(&mut self) -> Result<ID, DecodeError> {
        self.read_id()
    }

    fn read_right_id(&mut self) -> Result<ID, DecodeError> {
        self.read_id()
    }

    fn read_client(&mut self) -> Result<ClientID, DecodeError> {
        if self.blocks_remaining == 0 {
            return Err(decode_error("update client is missing its block count"));
        }
        let client = checked_client_id(self.read_var_u64_checked()?)?;
        if let Some(layout) = &mut self.layout {
            layout.clients.push(ClientLayout {
                client,
                blocks: Vec::new(),
            });
        }
        self.begin_capture(CaptureKind::BlockClock)?;
        Ok(client)
    }

    fn read_info(&mut self) -> Result<u8, DecodeError> {
        if self.pending_length.is_some() {
            return Err(decode_error("update block is missing its content length"));
        }
        let (start, clock) = (self.offset(), self.clock.unwrap_or_default());
        let info = self.read_u8()?;
        if let Some(section) = self
            .layout
            .as_mut()
            .and_then(|layout| layout.clients.last_mut())
        {
            section.blocks.push(BlockLayout {
                start,
                end: start,
                clock,
                end_clock: clock,
                skip: info == BLOCK_SKIP_REF_NUMBER,
            });
        }
        if info == BLOCK_SKIP_REF_NUMBER {
            self.begin_capture(CaptureKind::SkipLength)?;
        } else if info == BLOCK_GC_REF_NUMBER {
            self.pending_length = Some(PendingLength::Gc);
        } else {
            match info & 0x0f {
                BLOCK_ITEM_DELETED_REF_NUMBER => {
                    self.pending_length = Some(PendingLength::Deleted);
                }
                BLOCK_ITEM_TYPE_REF_NUMBER => {
                    self.pending_length = Some(PendingLength::Type);
                }
                BLOCK_ITEM_ANY_REF_NUMBER => {
                    self.pending_length = Some(PendingLength::Any);
                }
                _ => return Err(decode_error("update contains an unsupported block type")),
            }
        }
        Ok(info)
    }

    fn read_parent_info(&mut self) -> Result<bool, DecodeError> {
        match self.read_var_u32_checked()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(decode_error("update contains invalid parent info")),
        }
    }

    fn read_type_ref(&mut self) -> Result<u8, DecodeError> {
        if !matches!(self.pending_length.take(), Some(PendingLength::Type)) {
            return Err(decode_error("update contains an unexpected shared type"));
        }
        let type_ref = self.read_u8()?;
        if type_ref != TYPE_REFS_ARRAY && type_ref != TYPE_REFS_MAP {
            return Err(decode_error("update contains an unsupported shared type"));
        }
        self.advance_clock(1)?;
        self.finish_block()?;
        Ok(type_ref)
    }

    fn read_len(&mut self) -> Result<u32, DecodeError> {
        let len = self.read_var_u32_checked()?;
        let pending = self
            .pending_length
            .take()
            .ok_or_else(|| decode_error("update contains an unexpected block length"))?;
        match pending {
            PendingLength::Gc => {
                if len == 0 {
                    return Err(decode_error("update contains an empty GC block"));
                }
                self.advance_clock(len)?;
                self.finish_block()?;
            }
            PendingLength::Deleted => {
                self.advance_clock(len)?;
                self.finish_block()?;
            }
            PendingLength::Any => {
                self.declared_any_items = self
                    .declared_any_items
                    .checked_add(len as usize)
                    .ok_or_else(|| decode_error("update value count overflow"))?;
                if self.declared_any_items > self.max_values || len as usize > self.remaining {
                    return Err(decode_error("update contains too many values"));
                }
                self.any_items_remaining = len;
                self.any_block_len = len;
                if len == 0 {
                    self.finish_block()?;
                }
            }
            PendingLength::Type => {
                return Err(decode_error("shared type block contains a length"));
            }
        }
        Ok(len)
    }

    fn read_any(&mut self) -> Result<Any, DecodeError> {
        if self.any_items_remaining == 0 {
            return Err(decode_error("update contains an undeclared value"));
        }
        let value = self.decode_any(0)?;
        self.any_items_remaining -= 1;
        if self.any_items_remaining == 0 {
            self.advance_clock(self.any_block_len)?;
            self.any_block_len = 0;
            self.finish_block()?;
        }
        Ok(value)
    }

    fn read_json(&mut self) -> Result<Any, DecodeError> {
        Any::from_json(self.read_string()?)
    }

    fn read_key(&mut self) -> Result<Arc<str>, DecodeError> {
        Ok(Arc::from(self.read_string()?))
    }

    fn read_to_end(&mut self) -> Result<&[u8], DecodeError> {
        if self.update_mode
            && (self.capture.is_some()
                || self.update_clients_remaining != 0
                || self.blocks_remaining != 0
                || self.pending_length.is_some()
                || self.any_items_remaining != 0
                || self.delete_clients_remaining != 0
                || self.delete_ranges_remaining != 0)
        {
            return Err(decode_error("update ended before its declared content"));
        }
        let bytes = self.inner.read_to_end()?;
        if bytes.len() != self.remaining {
            return Err(decode_error("update decoder length mismatch"));
        }
        Ok(bytes)
    }
}

fn checked_client_id(client: u64) -> Result<ClientID, DecodeError> {
    if client > MAX_SAFE_CLIENT_ID {
        Err(decode_error("client ID exceeds the 53-bit Yjs limit"))
    } else {
        Ok(ClientID::new(client))
    }
}

fn checked_clock(clock: u32) -> Result<(), DecodeError> {
    if clock > MAX_SAFE_CLOCK {
        Err(decode_error("clock exceeds the supported i32 range"))
    } else {
        Ok(())
    }
}

fn decode_error(message: impl Into<String>) -> DecodeError {
    DecodeError::Custom(message.into())
}

fn validate_state_vector_entry_count(bytes: &[u8]) -> Result<(), String> {
    let Some((&first, _)) = bytes.split_first() else {
        return Err("state vector is empty".to_string());
    };
    let mut value = u32::from(first & 0x7f);
    let mut shift = 7;
    let mut used = 1;
    let mut byte = first;
    while byte & 0x80 != 0 {
        if used == 5 || used >= bytes.len() {
            return Err("invalid state vector entry count".to_string());
        }
        byte = bytes[used];
        if used == 4 && byte > 0x0f {
            return Err("invalid state vector entry count".to_string());
        }
        value |= u32::from(byte & 0x7f) << shift;
        shift += 7;
        used += 1;
    }
    if value > MAX_STATE_VECTOR_ENTRIES {
        return Err(format!(
            "state vector contains {value} entries, exceeds the {MAX_STATE_VECTOR_ENTRIES}-entry limit"
        ));
    }
    if value as usize > bytes.len().saturating_sub(used) / 2 {
        return Err("state vector entry count exceeds its payload".to_string());
    }
    Ok(())
}

fn requires_full_semantic_sync(op: &Op) -> bool {
    matches!(
        op,
        Op::InsertRows { .. }
            | Op::DeleteRows { .. }
            | Op::InsertCols { .. }
            | Op::DeleteCols { .. }
            | Op::SetHyperlinks { .. }
            | Op::RestoreColStyles { .. }
            | Op::RestoreRowStyles { .. }
            | Op::SetCharts { .. }
            | Op::SetChartAnchor { .. }
            | Op::RemoveSheet { .. }
            | Op::MoveSheet { .. }
            | Op::RenameSheet { .. }
            | Op::RestoreSheet { .. }
    )
}

#[derive(Clone)]
enum SheetToken {
    Existing(String),
    Added(usize),
}

fn targeted_sheet_keys(
    current: &[String],
    desired: &[String],
    ops: &[Op],
) -> Result<Vec<Option<String>>, String> {
    let mut tokens = current
        .iter()
        .cloned()
        .map(SheetToken::Existing)
        .collect::<Vec<_>>();
    let mut targets = Vec::with_capacity(ops.len());
    let mut next_added = 0;
    for op in ops {
        match op {
            Op::AddSheet { index, .. } => {
                if *index > tokens.len() {
                    return Err(format!("sheet insertion index {index} is out of range"));
                }
                tokens.insert(*index, SheetToken::Added(next_added));
                next_added += 1;
                targets.push(None);
            }
            Op::RemoveSheet { index } => {
                if *index >= tokens.len() {
                    return Err(format!("sheet removal index {index} is out of range"));
                }
                tokens.remove(*index);
                targets.push(None);
            }
            Op::MoveSheet { from, to } => {
                let token = tokens.remove(*from);
                tokens.insert(*to, token);
                targets.push(None);
            }
            Op::SetDefinedNames { .. }
            | Op::SetDefinedName { .. }
            | Op::DeleteDefinedName { .. } => targets.push(None),
            op => {
                let sheet = op_sheet(op)
                    .ok_or_else(|| "operation has no sheet target".to_string())?
                    .0 as usize;
                let token = tokens
                    .get(sheet)
                    .cloned()
                    .ok_or_else(|| format!("sheet {sheet} is out of range"))?;
                targets.push(Some(token));
            }
        }
    }
    if tokens.len() != desired.len() {
        return Err("sheet operation plan does not match final order".to_string());
    }
    let mut added = HashMap::new();
    for (token, key) in tokens.iter().zip(desired) {
        match token {
            SheetToken::Existing(existing) if existing != key => {
                return Err("sheet operation plan changed retained identity".to_string());
            }
            SheetToken::Existing(_) => {}
            SheetToken::Added(id) => {
                added.insert(*id, key.clone());
            }
        }
    }
    let active = desired.iter().map(String::as_str).collect::<HashSet<_>>();
    Ok(targets
        .into_iter()
        .map(|target| match target {
            Some(SheetToken::Existing(key)) if active.contains(key.as_str()) => Some(key),
            Some(SheetToken::Added(id)) => added.get(&id).cloned(),
            _ => None,
        })
        .collect())
}

fn op_sheet(op: &Op) -> Option<SheetId> {
    match op {
        Op::SetCell { sheet, .. }
        | Op::InsertRows { sheet, .. }
        | Op::DeleteRows { sheet, .. }
        | Op::InsertCols { sheet, .. }
        | Op::DeleteCols { sheet, .. }
        | Op::SetColWidth { sheet, .. }
        | Op::SetColVisibility { sheet, .. }
        | Op::RestoreColDimension { sheet, .. }
        | Op::SetRowHeight { sheet, .. }
        | Op::SetRowVisibility { sheet, .. }
        | Op::RestoreRowDimension { sheet, .. }
        | Op::SetFreezePane { sheet, .. }
        | Op::SetHyperlinks { sheet, .. }
        | Op::SetHyperlink { sheet, .. }
        | Op::RemoveHyperlink { sheet, .. }
        | Op::AddConditionalFormat { sheet, .. }
        | Op::RemoveConditionalFormat { sheet, .. }
        | Op::SetConditionalFormats { sheet, .. }
        | Op::SetAutoFilter { sheet, .. }
        | Op::RestoreAutoFilter { sheet, .. }
        | Op::RestoreColStyles { sheet, .. }
        | Op::RestoreRowStyles { sheet, .. }
        | Op::SetCharts { sheet, .. }
        | Op::SetChartAnchor { sheet, .. }
        | Op::MergeCells { sheet, .. }
        | Op::UnmergeCells { sheet, .. }
        | Op::PatchRangeStyle { sheet, .. }
        | Op::SetRangeNumberFormat { sheet, .. }
        | Op::ApplyRangeFormat { sheet, .. }
        | Op::RenameSheet { sheet, .. }
        | Op::RestoreSheet { sheet, .. } => Some(*sheet),
        Op::AddSheet { .. }
        | Op::RemoveSheet { .. }
        | Op::MoveSheet { .. }
        | Op::SetDefinedName { .. }
        | Op::DeleteDefinedName { .. }
        | Op::SetDefinedNames { .. } => None,
    }
}

fn patch_sheet_order(
    order: &ArrayRef,
    txn: &mut TransactionMut<'_>,
    existing: &[String],
    desired: &[String],
) -> Result<(), String> {
    let mut working = existing.to_vec();
    let mut index = 0;
    while index < desired.len() {
        if working.get(index) == desired.get(index) {
            index += 1;
            continue;
        }
        if let Some(offset) = working[index..]
            .iter()
            .position(|key| key == &desired[index])
        {
            order.remove_range(txn, yrs_index(index)?, yrs_index(offset)?);
            working.drain(index..index + offset);
        } else {
            order.insert(txn, yrs_index(index)?, desired[index].clone());
            working.insert(index, desired[index].clone());
            index += 1;
        }
    }
    if working.len() > desired.len() {
        order.remove_range(
            txn,
            yrs_index(desired.len())?,
            yrs_index(working.len() - desired.len())?,
        );
    }
    Ok(())
}

fn yrs_index(index: usize) -> Result<u32, String> {
    u32::try_from(index).map_err(|_| "sheet order exceeds Yrs index range".to_string())
}

fn sheet_map_for_sync(
    sheets: &MapRef,
    txn: &mut TransactionMut<'_>,
    key: &str,
) -> Result<MapRef, String> {
    match sheets.get(txn, key) {
        Some(Out::YMap(map)) => Ok(map),
        Some(_) => Err(format!("sheet {key} is not a map")),
        None => Ok(sheets.insert(txn, key, MapPrelim::default())),
    }
}

fn sheet_parts_by_key<'a, T: ReadTxn>(
    sheets: &MapRef,
    txn: &T,
    keys: &[String],
    model: &'a WorkbookModel,
    key: &str,
) -> Result<(MapRef, &'a Sheet), String> {
    let index = keys
        .iter()
        .position(|candidate| candidate == key)
        .ok_or_else(|| format!("sheet {key} is not active"))?;
    let sheet_map = sheets
        .get(txn, key)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| format!("missing sheet {key}"))?;
    let sheet_model = model
        .sheets
        .get(index)
        .ok_or_else(|| format!("sheet {key} is missing from the projection"))?;
    Ok((sheet_map, sheet_model))
}

fn sync_sheet(
    sheet_map: &MapRef,
    txn: &mut TransactionMut<'_>,
    sheet: &Sheet,
    stylesheet: &Stylesheet,
) -> Result<(), String> {
    sync_sheet_with_style_key(sheet_map, txn, sheet, |style| style_key(stylesheet, style))
}

fn sync_sheet_with_style_key(
    sheet_map: &MapRef,
    txn: &mut TransactionMut<'_>,
    sheet: &Sheet,
    style_key: impl Fn(u32) -> Result<String, String>,
) -> Result<(), String> {
    let [col_widths, contents, row_heights, styles] = sync_sheet_frame(sheet_map, txn, sheet)?;
    sync_numbers(&col_widths, txn, &sheet.col_widths);
    sync_map(&contents, txn, desired_contents(sheet));
    sync_numbers(&row_heights, txn, &sheet.row_heights);
    sync_map(&styles, txn, desired_styles(sheet, style_key)?);
    Ok(())
}

/// Writes every key of a sheet but its per-cell entries, and returns the maps
/// those go in: column widths, contents, row heights and styles.
fn sync_sheet_frame(
    sheet_map: &MapRef,
    txn: &mut TransactionMut<'_>,
    sheet: &Sheet,
) -> Result<[MapRef; 4], String> {
    let col_widths: MapRef = sheet_map.get_or_init(txn, COL_WIDTHS);
    let contents: MapRef = sheet_map.get_or_init(txn, CONTENTS);
    if !sheet.row_styles.is_empty() || sheet_map.get(txn, ROW_STYLES).is_some() {
        let rows = sheet
            .row_styles
            .iter()
            .map(|(&row, &style)| Any::Array(vec![Any::from(row), Any::from(style)].into()))
            .collect::<Vec<_>>();
        sheet_map.try_update(txn, ROW_STYLES, Any::Array(rows.into()));
    }
    if !sheet.col_styles.is_empty() || sheet_map.get(txn, COL_STYLES).is_some() {
        let columns = sheet
            .col_styles
            .iter()
            .map(|run| {
                Any::Array(
                    vec![Any::from(run.first), Any::from(run.last), Any::from(run.xf)].into(),
                )
            })
            .collect::<Vec<_>>();
        sheet_map.try_update(txn, COL_STYLES, Any::Array(columns.into()));
    }
    sheet_map.try_update(txn, FREEZE_PANE, freeze_pane_to_any(sheet.freeze_pane));
    let hyperlinks = serde_json::to_string(&sheet.hyperlinks)
        .map_err(|error| format!("cannot encode sheet hyperlinks: {error}"))?;
    sheet_map.try_update(txn, HYPERLINKS, hyperlinks);
    clear_hyperlink_edits(sheet_map, txn);
    if sheet.auto_filter.is_some() || sheet_map.get(txn, AUTO_FILTER).is_some() {
        let json = serde_json::to_string(&sheet.auto_filter)
            .map_err(|error| format!("cannot encode sheet auto filter: {error}"))?;
        sheet_map.try_update(txn, AUTO_FILTER, json);
    }
    if !sheet.conditional_formats.is_empty() || sheet_map.get(txn, CONDITIONAL_FORMATS).is_some() {
        let conditional_formats = serde_json::to_string(&sheet.conditional_formats)
            .map_err(|error| format!("cannot encode sheet conditional formats: {error}"))?;
        sheet_map.try_update(txn, CONDITIONAL_FORMATS, conditional_formats);
    }
    let charts = serde_json::to_string(&sheet.charts)
        .map_err(|error| format!("cannot encode sheet charts: {error}"))?;
    sheet_map.try_update(txn, CHARTS, charts);
    sheet_map.try_update(txn, MERGES, merges_to_any(&sheet.merges));
    clear_merge_edits(sheet_map, txn);
    sheet_map.try_update(txn, NAME, sheet.name.as_str());
    let row_heights: MapRef = sheet_map.get_or_init(txn, ROW_HEIGHTS);
    let styles: MapRef = sheet_map.get_or_init(txn, STYLES);
    Ok([col_widths, contents, row_heights, styles])
}

fn desired_contents(sheet: &Sheet) -> BTreeMap<String, Any> {
    sheet
        .iter_cells()
        .filter_map(|(at, cell)| content_to_any(cell).map(|value| (cell_key(at), value)))
        .collect()
}

fn desired_styles(
    sheet: &Sheet,
    style_key: impl Fn(u32) -> Result<String, String>,
) -> Result<BTreeMap<String, Any>, String> {
    sheet
        .iter_cells()
        .filter_map(|(at, cell)| cell.style.map(|style| (at, style)))
        .map(|(at, style)| style_key(style).map(|key| (cell_key(at), Any::from(key))))
        .collect()
}

fn number_entries(values: &BTreeMap<u32, f64>) -> impl Iterator<Item = (String, Any)> + '_ {
    values
        .iter()
        .map(|(index, &value)| (index.to_string(), Any::from(value)))
}

fn sync_map(map: &MapRef, txn: &mut TransactionMut<'_>, desired: BTreeMap<String, Any>) {
    let mut stale = map
        .keys(txn)
        .filter(|key| !desired.contains_key(*key))
        .map(str::to_string)
        .collect::<Vec<_>>();
    stale.sort();
    for key in stale {
        map.remove(txn, &key);
    }
    for (key, value) in desired {
        map.try_update(txn, key, value);
    }
}

fn sync_authored_cell(
    sheet_map: &MapRef,
    txn: &mut TransactionMut<'_>,
    sheet: &Sheet,
    at: CellRef,
) -> Result<(), String> {
    let contents: MapRef = sheet_map.get_or_init(txn, CONTENTS);
    let key = cell_key(at);
    match sheet.cell(at) {
        Some(cell) => {
            let current = match contents.get(txn, &key) {
                Some(Out::Any(value)) => Some(content_from_any(&value)?),
                Some(_) => return Err(format!("cell content {key} is not an atomic value")),
                None => None,
            };
            if !authored_content_equal(current.as_ref(), Some(cell)) {
                sync_optional(&contents, txn, &key, content_to_any(cell));
            }
        }
        None => {
            contents.remove(txn, &key);
        }
    }
    Ok(())
}

fn sync_cell_format(
    sheet_map: &MapRef,
    txn: &mut TransactionMut<'_>,
    sheet: &Sheet,
    stylesheet: &Stylesheet,
    at: CellRef,
) -> Result<(), String> {
    let styles: MapRef = sheet_map.get_or_init(txn, STYLES);
    let key = cell_key(at);
    let style = sheet
        .cell(at)
        .and_then(|cell| cell.style)
        .map(|style| style_key(stylesheet, style).map(Any::from))
        .transpose()?;
    sync_optional(&styles, txn, &key, style);
    Ok(())
}

fn sync_cell_formats(
    map: &MapRef,
    txn: &mut TransactionMut<'_>,
    stylesheet: &Stylesheet,
) -> Result<(), String> {
    let (key, payload) = cell_format_entry(&CellFormat::default())?;
    map.try_update(txn, key, payload);
    for index in 0..stylesheet.cell_xfs.len() {
        let index =
            u32::try_from(index).map_err(|_| "cell format table is too large".to_string())?;
        let (key, payload) = style_entry(stylesheet, index)?;
        map.try_update(txn, key, payload);
    }
    Ok(())
}

fn materialize_cell_formats<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    base: &Stylesheet,
) -> Result<(Stylesheet, BTreeMap<String, Option<u32>>), String> {
    let mut catalog = BTreeMap::new();
    for (key, value) in map.iter(txn) {
        let Out::Any(Any::String(payload)) = value else {
            return Err(format!("cell format {key} is not a string"));
        };
        if payload.len() > MAX_CELL_FORMAT_BYTES {
            return Err(format!("cell format {key} exceeds its size limit"));
        }
        let stored = serde_json::from_str::<StoredCellFormat>(&payload)
            .map_err(|error| format!("invalid cell format {key}: {error}"))?;
        let (expected, canonical) = stored_format_entry(&stored)?;
        if key != expected || *payload != canonical {
            return Err(format!("cell format {key} is not canonical"));
        }
        catalog.insert(key.to_string(), stored);
    }

    let mut styles = base.clone();
    styles.preserve_style_identity();
    let mut known = BTreeMap::new();
    for index in 0..base.cell_xfs.len() {
        let index =
            u32::try_from(index).map_err(|_| "cell format table is too large".to_string())?;
        known.entry(style_key(base, index)?).or_insert(Some(index));
    }
    let mut indices = BTreeMap::new();
    for (key, stored) in catalog {
        let index = match known.get(&key) {
            Some(index) => *index,
            None => {
                let index = match stored.derived_from {
                    Some(source) => Some(
                        styles
                            .intern_derived_cell_format(source, &stored.format)
                            .map_err(|_| "number format table is full".to_string())?,
                    ),
                    None => styles
                        .intern_cell_format(&stored.format)
                        .map_err(|_| "number format table is full".to_string())?,
                };
                if stored.source_xf.is_none()
                    && stored.derived_from.is_none()
                    && let Some(index) = index
                {
                    styles.mark_generated_format(index);
                }
                known.insert(key.clone(), index);
                index
            }
        };
        indices.insert(key, index);
    }
    Ok((styles, indices))
}

/// Maps authored style indices between tables.
fn remap_styles<'a>(
    ops: &'a [Op],
    source: &Stylesheet,
    target: &mut Stylesheet,
) -> Result<Cow<'a, [Op]>, String> {
    let mut mapped: HashMap<u32, Option<u32>> = HashMap::new();
    let mut resolve = |style: Option<u32>| -> Result<Option<u32>, String> {
        let Some(index) = style else {
            return Ok(None);
        };
        if let Some(&known) = mapped.get(&index) {
            return Ok(known);
        }
        if index != 0 && source.xf(index).is_none() {
            return Err(format!("cell style index {index} is out of range"));
        }
        let format = source.cell_format(Some(index));
        let source_key = style_key(source, index)?;
        let resolved = if target.xf(index).is_some() && style_key(target, index)? == source_key {
            Some(index)
        } else if target.binds_plain() && format == CellFormat::default() {
            Some(target.plain_edit_xf())
        } else {
            let found = match (0..target.cell_xfs.len() as u32)
                .find(|&candidate| style_key(target, candidate).is_ok_and(|key| key == source_key))
            {
                Some(found) => Some(found),
                None => match source.format_source(index).1 {
                    Some(source) => Some(
                        target
                            .intern_derived_cell_format(source, &format)
                            .map_err(|_| "number format table is full".to_string())?,
                    ),
                    None if format == CellFormat::default() => (0..target.cell_xfs.len() as u32)
                        .find(|&candidate| target.cell_format(Some(candidate)) == format),
                    None => target
                        .intern_cell_format(&format)
                        .map_err(|_| "number format table is full".to_string())?,
                },
            };
            Some(found.unwrap_or_else(|| target.plain_xf()))
        };
        mapped.insert(index, resolved);
        Ok(resolved)
    };
    let mut remapped: Option<Vec<Op>> = None;
    for (position, op) in ops.iter().enumerate() {
        let rewritten = match op {
            Op::SetCell { sheet, at, cell } => {
                let style = resolve(cell.style)?;
                (style != cell.style).then(|| Op::SetCell {
                    sheet: *sheet,
                    at: *at,
                    cell: CellState {
                        style,
                        ..cell.clone()
                    },
                })
            }
            Op::RestoreSheet {
                sheet,
                name,
                formulas,
            } => {
                let mut changed = false;
                let mut restored = Vec::with_capacity(formulas.len());
                for (formula_sheet, at, state) in formulas {
                    let style = resolve(state.style)?;
                    changed |= style != state.style;
                    restored.push((
                        *formula_sheet,
                        *at,
                        CellState {
                            style,
                            ..state.clone()
                        },
                    ));
                }
                changed.then(|| Op::RestoreSheet {
                    sheet: *sheet,
                    name: name.clone(),
                    formulas: restored,
                })
            }
            _ => None,
        };
        if let Some(rewritten) = rewritten {
            remapped
                .get_or_insert_with(|| ops[..position].to_vec())
                .push(rewritten);
        } else if let Some(remapped) = &mut remapped {
            remapped.push(op.clone());
        }
    }
    Ok(match remapped {
        Some(remapped) => Cow::Owned(remapped),
        None => Cow::Borrowed(ops),
    })
}

/// Whether `format` fits the size a shared cell format may take.
pub(crate) fn cell_format_fits(format: &CellFormat) -> bool {
    cell_format_entry(format).is_ok_and(|(_, payload)| payload.len() <= MAX_CELL_FORMAT_BYTES)
}

fn style_key(stylesheet: &Stylesheet, style: u32) -> Result<String, String> {
    if style != 0 && stylesheet.xf(style).is_none() {
        return Err(format!("cell style index {style} is out of range"));
    }
    style_entry(stylesheet, style).map(|(key, _)| key)
}

fn style_entry(stylesheet: &Stylesheet, index: u32) -> Result<(String, String), String> {
    let (source_xf, derived_from) = stylesheet.format_source(index);
    stored_format_entry(&StoredCellFormat {
        format: stylesheet.cell_format(Some(index)),
        source_xf,
        derived_from,
    })
}

fn cell_format_entry(format: &CellFormat) -> Result<(String, String), String> {
    stored_format_entry(&StoredCellFormat {
        format: format.clone(),
        source_xf: None,
        derived_from: None,
    })
}

fn stored_format_entry(stored: &StoredCellFormat) -> Result<(String, String), String> {
    let payload = serde_json::to_string(stored)
        .map_err(|error| format!("cannot encode cell format: {error}"))?;
    let digest = Sha256::digest(payload.as_bytes());
    Ok((format!("{digest:x}"), payload))
}

/// `format` as releases before format facets resolved it.
fn pre_facet_format(format: &CellFormat) -> CellFormat {
    CellFormat {
        font: format.font.clone(),
        fill: format.fill.clone(),
        border: format.border.clone(),
        number_format: format.number_format.clone(),
        alignment: xlsx_model::Alignment {
            h: format.alignment.h,
            v: format.alignment.v,
            wrap_text: format.alignment.wrap_text,
            shrink_to_fit: format.alignment.shrink_to_fit,
            ..xlsx_model::Alignment::default()
        },
        ..CellFormat::default()
    }
}

/// Moves cells keyed by a release before format facets onto the formats of
/// `styles`. A cell takes its counterpart's format in `source` when that
/// reads the same, and otherwise the first one that does, as such a release
/// resolved it.
fn migrate_pre_facet_formats(
    txn: &mut TransactionMut<'_>,
    styles: &Stylesheet,
    source: Option<&Doc>,
    reading: FingerprintReading,
) -> Result<(), String> {
    let mut entries = Vec::with_capacity(styles.cell_xfs.len());
    let mut first = HashMap::new();
    let mut readings = HashMap::new();
    for index in 0..styles.cell_xfs.len() {
        let index =
            u32::try_from(index).map_err(|_| "cell format table is too large".to_string())?;
        let format = styles.cell_format(Some(index));
        let pre_facet = pre_facet_format(&format);
        let released = if reading.as_written() {
            pre_facet
        } else {
            released_format(&pre_facet)
        };
        let (released, _) = cell_format_entry(&released)?;
        let entry = style_entry(styles, index)?;
        first.entry(released.clone()).or_insert(entries.len());
        readings.insert(entry.0.clone(), released);
        entries.push(entry);
    }
    if readings.iter().all(|(key, released)| key == released) {
        return Ok(());
    }
    let source = source.map(Doc::transact);
    let sheets = txn
        .get_map(SHEETS)
        .ok_or_else(|| "missing sheet map".to_string())?;
    let catalog = txn
        .get_map(CELL_FORMATS)
        .ok_or_else(|| "missing cell format catalog".to_string())?;
    let legacy_readings = catalog
        .iter(&*txn)
        .map(|(key, value)| {
            let payload = value
                .cast::<String>()
                .map_err(|_| format!("cell format {key} is not a string"))?;
            let stored = serde_json::from_str::<StoredCellFormat>(&payload)
                .map_err(|error| format!("invalid cell format {key}: {error}"))?;
            let (reading, _) = cell_format_entry(&stored.format)?;
            Ok((key.to_string(), reading))
        })
        .collect::<Result<HashMap<_, _>, String>>()?;
    let keys = sheets.keys(&*txn).map(str::to_string).collect::<Vec<_>>();
    let mut used = HashSet::new();
    for key in keys {
        let sheet = sheets
            .get(&*txn, &key)
            .and_then(|value| value.cast::<MapRef>().ok())
            .ok_or_else(|| format!("sheet {key} is not a map"))?;
        let cell_styles = nested_map(&sheet, &*txn, STYLES)?;
        let counterparts = source.as_ref().and_then(|source| {
            let sheet = source
                .get_map(SHEETS)?
                .get(source, &key)?
                .cast::<MapRef>()
                .ok()?;
            nested_map(&sheet, source, STYLES).ok()
        });
        let cells = cell_styles
            .iter(&*txn)
            .filter_map(|(cell, value)| Some((cell.to_string(), value.cast::<String>().ok()?)))
            .collect::<Vec<_>>();
        for (cell, released) in cells {
            let released_reading = legacy_readings.get(&released).unwrap_or(&released);
            let counterpart = counterparts
                .as_ref()
                .zip(source.as_ref())
                .and_then(|(styles, source)| styles.get(source, &cell))
                .and_then(|value| value.cast::<String>().ok());
            let target = match counterpart {
                Some(counterpart) if readings.get(&counterpart) == Some(released_reading) => {
                    counterpart
                }
                _ => match first.get(released_reading) {
                    Some(&index) => entries[index].0.clone(),
                    None => continue,
                },
            };
            if target != released {
                cell_styles.try_update(txn, cell, Any::from(target.clone()));
            }
            used.insert(target);
        }
    }
    for (key, payload) in &entries {
        if used.contains(key) {
            catalog.try_update(txn, key.clone(), payload.clone());
        }
    }
    let (default, _) = cell_format_entry(&CellFormat::default())?;
    for (key, reading) in &legacy_readings {
        if first.contains_key(reading) && !readings.contains_key(key) && *key != default {
            catalog.remove(txn, key);
        }
    }
    Ok(())
}

fn authored_content_equal(left: Option<&Cell>, right: Option<&Cell>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => match (&left.formula, &right.formula) {
            (Some(left), Some(right)) => left == right,
            (None, None) => left.value == right.value,
            _ => false,
        },
        _ => false,
    }
}

fn sync_optional(map: &MapRef, txn: &mut TransactionMut<'_>, key: &str, value: Option<Any>) {
    if let Some(value) = value {
        map.try_update(txn, key, value);
    } else {
        map.remove(txn, key);
    }
}

fn sync_numbers(map: &MapRef, txn: &mut TransactionMut<'_>, values: &BTreeMap<u32, f64>) {
    let retained = values.keys().map(u32::to_string).collect::<HashSet<_>>();
    let mut stale = map
        .keys(txn)
        .filter(|key| !retained.contains(*key))
        .map(str::to_string)
        .collect::<Vec<_>>();
    stale.sort();
    for key in stale {
        map.remove(txn, &key);
    }
    for (&index, &value) in values {
        map.try_update(txn, index.to_string(), value);
    }
}

fn sync_number(map: &MapRef, txn: &mut TransactionMut<'_>, index: u32, value: Option<f64>) {
    let key = index.to_string();
    match value {
        Some(value) => {
            map.try_update(txn, key, value);
        }
        None => {
            map.remove(txn, &key);
        }
    }
}

/// What a sheet reads from the workbook it was parsed from rather than from
/// the shared document, because the document's schema cannot express it.
#[derive(Clone, Copy)]
struct SheetFallbacks<'a> {
    freeze_pane: Option<FreezePane>,
    format: SheetFormat,
    col_styles: &'a [ColStyle],
    row_styles: &'a BTreeMap<u32, u32>,
    hyperlinks: &'a [Hyperlink],
    charts: &'a [SheetChart],
    hidden_dimensions: &'a HiddenDimensions,
    conditional_formats: &'a [ConditionalFormatting],
    auto_filter: Option<&'a xlsx_model::AutoFilter>,
}

impl Default for SheetFallbacks<'_> {
    fn default() -> Self {
        Self {
            freeze_pane: None,
            format: SheetFormat::default(),
            col_styles: &[],
            row_styles: &EMPTY_ROW_STYLES,
            hyperlinks: &[],
            charts: &[],
            hidden_dimensions: &EMPTY_HIDDEN_DIMENSIONS,
            conditional_formats: &[],
            auto_filter: None,
        }
    }
}

fn materialize_sheet<T: ReadTxn>(
    sheet_map: &MapRef,
    txn: &T,
    style_indices: &BTreeMap<String, Option<u32>>,
    plain: &mut dyn FnMut() -> u32,
    version: i64,
    fallbacks: SheetFallbacks<'_>,
) -> Result<Sheet, String> {
    let name = sheet_map
        .get(txn, NAME)
        .and_then(|value| value.cast::<String>().ok())
        .ok_or_else(|| "sheet is missing its name".to_string())?;
    let mut sheet = Sheet::new(name);
    let mut cells = BTreeMap::<(u32, u32), Cell>::new();
    let contents = nested_map(sheet_map, txn, CONTENTS)?;
    for (key, value) in contents.iter(txn) {
        xlsx_model::progress::tick();
        let at = parse_cell_key(key)?;
        let Out::Any(value) = value else {
            return Err(format!("cell content {key} is not an atomic value"));
        };
        cells.insert((at.row, at.col), content_from_any(&value)?);
    }
    let styles = nested_map(sheet_map, txn, STYLES)?;
    for (key, value) in styles.iter(txn) {
        xlsx_model::progress::tick();
        let at = parse_cell_key(key)?;
        let style_key = value
            .cast::<String>()
            .map_err(|_| format!("cell style {key} is not a string"))?;
        let style = style_indices
            .get(&style_key)
            .ok_or_else(|| format!("cell style {key} references an unknown format"))?;
        let style = match *style {
            Some(style) => style,
            None => plain(),
        };
        cells.entry((at.row, at.col)).or_default().style = Some(style);
    }
    sheet.adopt_cells(cells);
    sheet.col_widths = materialize_numbers(
        &nested_map(sheet_map, txn, COL_WIDTHS)?,
        txn,
        MAX_COLS,
        "column width",
    )?;
    sheet.row_heights = materialize_numbers(
        &nested_map(sheet_map, txn, ROW_HEIGHTS)?,
        txn,
        MAX_ROWS,
        "row height",
    )?;
    if version < SCHEMA_VERSION {
        for (&at, &size) in &fallbacks.hidden_dimensions.col_widths {
            sheet.col_widths.entry(at).or_insert(size);
        }
        for (&at, &size) in &fallbacks.hidden_dimensions.row_heights {
            sheet.row_heights.entry(at).or_insert(size);
        }
    }
    sheet.format = fallbacks.format;
    sheet.col_styles = fallbacks.col_styles.to_vec();
    sheet.row_styles = fallbacks.row_styles.clone();
    materialize_style_defaults(&mut sheet, sheet_map, txn)?;
    sheet.freeze_pane = match (version, sheet_map.get(txn, FREEZE_PANE)) {
        (FREEZE_PANE_SCHEMA_VERSION.., Some(Out::Any(value))) => freeze_pane_from_any(&value)?,
        (FREEZE_PANE_SCHEMA_VERSION.., _) => {
            return Err("sheet is missing freeze pane".to_string());
        }
        _ => fallbacks.freeze_pane,
    };
    let baseline_hyperlinks = match (version, sheet_map.get(txn, HYPERLINKS)) {
        (HYPERLINK_SCHEMA_VERSION.., Some(Out::Any(Any::String(json)))) => {
            decode_hyperlinks(&json)?
        }
        (HYPERLINK_SCHEMA_VERSION.., Some(_)) => {
            return Err("sheet hyperlinks are not a string".to_string());
        }
        (HYPERLINK_SCHEMA_VERSION.., None) => {
            return Err("sheet is missing hyperlinks".to_string());
        }
        _ => fallbacks.hyperlinks.to_vec(),
    };
    sheet.hyperlinks = project_hyperlinks(baseline_hyperlinks, sheet_map, txn)?;
    sheet.auto_filter = match sheet_map.get(txn, AUTO_FILTER) {
        Some(Out::Any(Any::String(json))) => serde_json::from_str(&json)
            .map_err(|error| format!("sheet auto filter is invalid: {error}"))?,
        Some(_) => return Err("sheet auto filter is not a string".to_string()),
        None => fallbacks.auto_filter.cloned(),
    };
    sheet.conditional_formats = match sheet_map.get(txn, CONDITIONAL_FORMATS) {
        Some(Out::Any(Any::String(json))) => serde_json::from_str(&json)
            .map_err(|error| format!("sheet conditional formats are invalid: {error}"))?,
        Some(_) => return Err("sheet conditional formats are not a string".to_string()),
        None => fallbacks.conditional_formats.to_vec(),
    };
    sheet.charts = match (version, sheet_map.get(txn, CHARTS)) {
        (CHARTS_SCHEMA_VERSION.., Some(Out::Any(Any::String(json)))) => decode_charts(&json)?,
        (CHARTS_SCHEMA_VERSION.., Some(_)) => {
            return Err("sheet charts are not a string".to_string());
        }
        _ => fallbacks.charts.to_vec(),
    };
    let baseline_merges = match sheet_map.get(txn, MERGES) {
        Some(Out::Any(value)) => merges_from_any(&value)?,
        _ => return Err("sheet is missing merges".to_string()),
    };
    sheet.merges = project_merges(baseline_merges, sheet_map, txn)?;
    Ok(sheet)
}

/// One drawing anchor is one element however many sheets point at it, but each
/// sheet's chart state is assigned on its own. Two replicas can each write a
/// legal value for a different sheet and only disagree once the two are merged,
/// so the disagreement is settled on the way out rather than refused on the way
/// in: refusing it would make integration depend on delivery order, and the
/// replicas would never meet again. Sheet order is shared, so first-in-order
/// wins is the same answer everywhere.
///
/// Removing a sheet is structural and refused while collaborative. Were that
/// ever allowed, dropping the donor would hand the frame to whatever the
/// surviving sheet's blob holds — which may be a value this has been quietly
/// covering for, and which nothing has checked since it arrived.
fn project_shared_frame_anchors(sheets: &mut [Sheet]) {
    let mut chosen = BTreeMap::new();
    for sheet in sheets.iter() {
        for chart in &sheet.charts {
            chosen.entry(chart.frame_id()).or_insert(chart.anchor);
        }
    }
    for sheet in sheets.iter_mut() {
        for chart in &mut sheet.charts {
            if let Some(anchor) = chosen.get(&chart.frame_id()) {
                chart.anchor = *anchor;
            }
        }
    }
}

fn nested_map<T: ReadTxn>(parent: &MapRef, txn: &T, key: &str) -> Result<MapRef, String> {
    parent
        .get(txn, key)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| format!("sheet is missing {key}"))
}

fn sheet_shared_types<T: ReadTxn>(sheet_map: &MapRef, txn: &T) -> Result<SheetSharedTypes, String> {
    Ok(SheetSharedTypes {
        sheet: sheet_map.as_ref().id(),
        col_widths: nested_map(sheet_map, txn, COL_WIDTHS)?.as_ref().id(),
        contents: nested_map(sheet_map, txn, CONTENTS)?.as_ref().id(),
        row_heights: nested_map(sheet_map, txn, ROW_HEIGHTS)?.as_ref().id(),
        styles: nested_map(sheet_map, txn, STYLES)?.as_ref().id(),
    })
}

fn materialize_numbers<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    limit: u32,
    label: &str,
) -> Result<BTreeMap<u32, f64>, String> {
    let mut values = BTreeMap::new();
    for (key, value) in map.iter(txn) {
        let index = key
            .parse::<u32>()
            .map_err(|_| format!("invalid numeric key {key}"))?;
        if key != index.to_string() {
            return Err(format!("noncanonical numeric key {key}"));
        }
        if index >= limit {
            return Err(format!("{label} key {key} is out of bounds"));
        }
        let value = value
            .cast::<f64>()
            .map_err(|_| format!("invalid numeric value at {key}"))?;
        if !value.is_finite() {
            return Err(format!("nonfinite {label} at {key}"));
        }
        values.insert(index, value);
    }
    Ok(values)
}

fn sheet_keys<T: ReadTxn>(order: &ArrayRef, txn: &T) -> Result<Vec<String>, String> {
    order
        .iter(txn)
        .map(|value| {
            value
                .cast::<String>()
                .map_err(|_| "sheet order contains a non-string key".to_string())
        })
        .collect()
}

fn structure_generation<T: ReadTxn>(meta: &MapRef, txn: &T) -> Result<i64, String> {
    let generation = meta
        .get(txn, STRUCTURE_GENERATION)
        .and_then(|value| value.cast::<i64>().ok())
        .ok_or_else(|| "missing structure generation".to_string())?;
    if generation < 0 {
        return Err("structure generation is negative".to_string());
    }
    Ok(generation)
}

fn validate_schema_version(version: i64) -> Result<(), String> {
    if (MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION).contains(&version) {
        Ok(())
    } else {
        Err(format!(
            "unsupported schema version {version}; supported versions are {MIN_SUPPORTED_SCHEMA_VERSION} through {SCHEMA_VERSION}"
        ))
    }
}

fn sheet_schema_keys(version: i64) -> &'static [&'static str] {
    const V3: &[&str] = &[COL_WIDTHS, CONTENTS, MERGES, NAME, ROW_HEIGHTS, STYLES];
    const V4: &[&str] = &[
        COL_WIDTHS,
        CONTENTS,
        FREEZE_PANE,
        MERGES,
        NAME,
        ROW_HEIGHTS,
        STYLES,
    ];
    const V5: &[&str] = &[
        COL_WIDTHS,
        CONTENTS,
        FREEZE_PANE,
        HYPERLINKS,
        MERGES,
        NAME,
        ROW_HEIGHTS,
        STYLES,
    ];
    const V6: &[&str] = &[
        COL_WIDTHS,
        CONTENTS,
        FREEZE_PANE,
        HYPERLINKS,
        MERGES,
        NAME,
        ROW_HEIGHTS,
        STYLES,
    ];
    match version {
        MIN_SUPPORTED_SCHEMA_VERSION => V3,
        FREEZE_PANE_SCHEMA_VERSION => V4,
        HYPERLINK_SCHEMA_VERSION => V5,
        _ => V6,
    }
}

/// Chart state is the one sheet key two replicas can assign at once, because
/// repinning a chart is an ordinary edit. Undoing the assignment that won
/// deletes the key outright, so it reads as absent and falls back to what the
/// source package holds rather than failing the whole workbook.
fn sheet_schema_optional_keys(version: i64) -> &'static [&'static str] {
    const NONE: &[&str] = &[];
    const V6: &[&str] = &[
        CHARTS,
        ROW_STYLES,
        COL_STYLES,
        CONDITIONAL_FORMATS,
        AUTO_FILTER,
    ];
    if version >= CHARTS_SCHEMA_VERSION {
        V6
    } else {
        NONE
    }
}

fn require_root_keys<T: ReadTxn>(txn: &T, expected: &[&str]) -> Result<(), String> {
    let actual = txn
        .root_refs()
        .map(|(key, _)| key.to_string())
        .collect::<BTreeSet<_>>();
    let expected = expected
        .iter()
        .map(|key| (*key).to_string())
        .collect::<BTreeSet<_>>();
    if actual == expected {
        Ok(())
    } else {
        Err("collaborative document roots do not match the schema".to_string())
    }
}

fn require_map_keys<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    expected: &[&str],
    label: &str,
) -> Result<(), String> {
    require_map_keys_with_optional(map, txn, expected, &[], label)
}

/// `optional` keys may be present or absent. A key two replicas assign at once
/// is deleted outright when the winning assignment is undone, so any such key
/// has to read as absent rather than as a broken document.
fn require_map_keys_with_optional<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    expected: &[&str],
    optional: &[&str],
    label: &str,
) -> Result<(), String> {
    let actual = map.keys(txn).map(str::to_string).collect::<BTreeSet<_>>();
    let expected = expected
        .iter()
        .map(|key| (*key).to_string())
        .collect::<BTreeSet<_>>();
    let optional = optional
        .iter()
        .map(|key| (*key).to_string())
        .collect::<BTreeSet<_>>();
    if expected.is_subset(&actual)
        && actual
            .difference(&expected)
            .all(|key| optional.contains(key))
    {
        Ok(())
    } else {
        Err(format!("{label} keys do not match the schema"))
    }
}

fn require_sheet_keys<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    version: i64,
    expected: &[&str],
    optional: &[&str],
    label: &str,
) -> Result<(), String> {
    #[cfg(test)]
    if SCHEMA_SIX_KEY_VALIDATION.get() {
        return require_map_keys_with_optional(map, txn, expected, optional, label);
    }
    if expected.iter().any(|key| map.get(txn, key).is_none()) {
        return Err(format!("{label} keys do not match the schema"));
    }
    let mut edits = MergeEdits::new();
    let mut hyperlinks = HyperlinkEdits::new();
    for (key, value) in map.iter(txn) {
        if key.starts_with(MERGE_PREFIX) {
            insert_merge_edit(&mut edits, key, &value)?;
        } else if key.starts_with(HYPERLINK_PREFIX) {
            if version < SCHEMA_VERSION {
                return Err(format!("{label} keys do not match the schema"));
            }
            insert_hyperlink_edit(&mut hyperlinks, key, &value)?;
        } else if !expected.contains(&key) && !optional.contains(&key) {
            return Err(format!("{label} keys do not match the schema"));
        }
    }
    Ok(())
}

fn cell_key(at: CellRef) -> String {
    format!("{}:{}", at.row, at.col)
}

fn parse_cell_key(key: &str) -> Result<CellRef, String> {
    let (row, col) = key
        .split_once(':')
        .ok_or_else(|| format!("invalid cell key {key}"))?;
    let row = row
        .parse::<u32>()
        .map_err(|_| format!("invalid cell row in {key}"))?;
    let col = col
        .parse::<u32>()
        .map_err(|_| format!("invalid cell column in {key}"))?;
    let at = CellRef::new(row, col);
    if key != cell_key(at) {
        return Err(format!("noncanonical cell key {key}"));
    }
    if row >= MAX_ROWS || col >= MAX_COLS {
        return Err(format!("cell key {key} is out of bounds"));
    }
    Ok(at)
}

fn content_to_any(cell: &Cell) -> Option<Any> {
    if let Some(formula) = &cell.formula {
        Some(any_array(vec![
            Any::BigInt(1),
            Any::from(formula.as_str()),
            value_to_any(&cell.value),
        ]))
    } else if !matches!(&cell.value, CellValue::Empty) {
        Some(any_array(vec![Any::BigInt(0), value_to_any(&cell.value)]))
    } else {
        None
    }
}

fn content_from_any(value: &Any) -> Result<Cell, String> {
    let values = any_values(value, "cell content")?;
    let kind = values
        .first()
        .ok_or_else(|| "cell content is empty".to_string())?;
    match any_i64(kind, "cell content kind")? {
        0 if values.len() == 2 => {
            let value = value_from_any(&values[1])?;
            if matches!(&value, CellValue::Empty) {
                return Err("empty literal cell content must be omitted".to_string());
            }
            Ok(Cell {
                value,
                ..Cell::default()
            })
        }
        1 if values.len() == 3 => {
            let Any::String(formula) = &values[1] else {
                return Err("formula cell content is missing formula text".to_string());
            };
            Ok(Cell {
                value: value_from_any(&values[2])?,
                formula: Some(formula.to_string()),
                ..Cell::default()
            })
        }
        0 | 1 => Err("cell content has the wrong payload length".to_string()),
        kind => Err(format!("unsupported cell content kind {kind}")),
    }
}

fn value_to_any(value: &CellValue) -> Any {
    match value {
        CellValue::Empty => any_array(vec![Any::BigInt(0)]),
        CellValue::Number { value } => any_array(vec![Any::BigInt(1), Any::Number(*value)]),
        CellValue::Text { value } => any_array(vec![Any::BigInt(2), Any::from(value.as_str())]),
        CellValue::Bool { value } => any_array(vec![Any::BigInt(3), Any::Bool(*value)]),
        CellValue::Error { value } => any_array(vec![Any::BigInt(4), Any::from(value.as_str())]),
    }
}

fn value_from_any(value: &Any) -> Result<CellValue, String> {
    let values = any_values(value, "cell value")?;
    let kind = values
        .first()
        .ok_or_else(|| "cell value is empty".to_string())?;
    match any_i64(kind, "cell value kind")? {
        0 if values.len() == 1 => Ok(CellValue::Empty),
        1 if values.len() == 2 => match &values[1] {
            Any::Number(value) if value.is_finite() => Ok(CellValue::Number { value: *value }),
            _ => Err("numeric cell has a non-number value".to_string()),
        },
        2 if values.len() == 2 => match &values[1] {
            Any::String(value) => Ok(CellValue::Text {
                value: value.to_string(),
            }),
            _ => Err("text cell has a non-string value".to_string()),
        },
        3 if values.len() == 2 => match &values[1] {
            Any::Bool(value) => Ok(CellValue::Bool { value: *value }),
            _ => Err("boolean cell has a non-boolean value".to_string()),
        },
        4 if values.len() == 2 => match &values[1] {
            Any::String(value) => Ok(CellValue::Error {
                value: error_from_str(value)?,
            }),
            _ => Err("error cell has a non-string value".to_string()),
        },
        0..=4 => Err("cell value has the wrong payload length".to_string()),
        kind => Err(format!("unsupported cell value kind {kind}")),
    }
}

fn any_array(values: Vec<Any>) -> Any {
    Any::Array(Arc::from(values))
}

fn any_values<'a>(value: &'a Any, label: &str) -> Result<&'a [Any], String> {
    match value {
        Any::Array(values) => Ok(values),
        _ => Err(format!("{label} is not an array")),
    }
}

fn any_i64(value: &Any, label: &str) -> Result<i64, String> {
    match value {
        Any::BigInt(value) => Ok(*value),
        _ => Err(format!("{label} is not an integer")),
    }
}

fn error_from_str(value: &str) -> Result<ErrorValue, String> {
    match value {
        "#DIV/0!" => Ok(ErrorValue::Div0),
        "#N/A" => Ok(ErrorValue::NA),
        "#NAME?" => Ok(ErrorValue::Name),
        "#NULL!" => Ok(ErrorValue::Null),
        "#NUM!" => Ok(ErrorValue::Num),
        "#REF!" => Ok(ErrorValue::Ref),
        "#VALUE!" => Ok(ErrorValue::Value),
        "#SPILL!" => Ok(ErrorValue::Spill),
        "#CALC!" => Ok(ErrorValue::Calc),
        _ => Err(format!("unsupported cell error {value}")),
    }
}

fn merges_to_any(merges: &[CellRange]) -> Any {
    Any::Array(Arc::from(
        merges
            .iter()
            .map(|range| {
                any_array(vec![
                    Any::from(range.start.row),
                    Any::from(range.start.col),
                    Any::Bool(range.start.abs_row),
                    Any::Bool(range.start.abs_col),
                    Any::from(range.end.row),
                    Any::from(range.end.col),
                    Any::Bool(range.end.abs_row),
                    Any::Bool(range.end.abs_col),
                ])
            })
            .collect::<Vec<_>>(),
    ))
}

fn freeze_pane_to_any(pane: Option<FreezePane>) -> Any {
    match pane {
        None => Any::Null,
        Some(pane) => any_array(vec![
            Any::from(pane.rows),
            Any::from(pane.cols),
            Any::from(pane.top_left.row),
            Any::from(pane.top_left.col),
            Any::Bool(pane.top_left.abs_row),
            Any::Bool(pane.top_left.abs_col),
        ]),
    }
}

fn freeze_pane_from_any(value: &Any) -> Result<Option<FreezePane>, String> {
    let Any::Array(values) = value else {
        return if matches!(value, Any::Null) {
            Ok(None)
        } else {
            Err("sheet freeze pane is not an array".to_string())
        };
    };
    if values.len() != 6 {
        return Err("sheet freeze pane must contain six values".to_string());
    }
    let pane = FreezePane::new(
        any_u32(&values[0], "frozen row count")?,
        any_u32(&values[1], "frozen column count")?,
        CellRef {
            row: any_u32(&values[2], "freeze pane top row")?,
            col: any_u32(&values[3], "freeze pane left column")?,
            abs_row: any_bool(&values[4], "freeze pane absolute row")?,
            abs_col: any_bool(&values[5], "freeze pane absolute column")?,
        },
    );
    if pane.rows > MAX_ROWS
        || pane.cols > MAX_COLS
        || pane.top_left.row >= MAX_ROWS
        || pane.top_left.col >= MAX_COLS
    {
        return Err("sheet freeze pane is out of bounds".to_string());
    }
    Ok(Some(pane))
}

fn merges_from_any(value: &Any) -> Result<Vec<CellRange>, String> {
    let Any::Array(merges) = value else {
        return Err("sheet merges are not an array".to_string());
    };
    let mut projection = MergeProjection::new(MergeEdits::new());
    merges
        .iter()
        .map(|merge| {
            let Any::Array(values) = merge else {
                return Err("merge entry is not an array".to_string());
            };
            if values.len() != 8 {
                return Err("merge entry must contain eight values".to_string());
            }
            let range = CellRange {
                start: CellRef {
                    row: any_u32(&values[0], "merge start row")?,
                    col: any_u32(&values[1], "merge start column")?,
                    abs_row: any_bool(&values[2], "merge start absolute row")?,
                    abs_col: any_bool(&values[3], "merge start absolute column")?,
                },
                end: CellRef {
                    row: any_u32(&values[4], "merge end row")?,
                    col: any_u32(&values[5], "merge end column")?,
                    abs_row: any_bool(&values[6], "merge end absolute row")?,
                    abs_col: any_bool(&values[7], "merge end absolute column")?,
                },
            };
            let _ = projection.baseline(range)?;
            Ok(range)
        })
        .collect()
}

fn any_u32(value: &Any, label: &str) -> Result<u32, String> {
    match value {
        Any::Number(value)
            if value.is_finite()
                && *value >= 0.0
                && *value <= u32::MAX as f64
                && value.fract() == 0.0 =>
        {
            Ok(*value as u32)
        }
        Any::BigInt(value) if *value >= 0 && *value <= i64::from(u32::MAX) => Ok(*value as u32),
        _ => Err(format!("{label} is not a u32")),
    }
}

fn any_bool(value: &Any, label: &str) -> Result<bool, String> {
    match value {
        Any::Bool(value) => Ok(*value),
        _ => Err(format!("{label} is not a boolean")),
    }
}

/// Dimensions only this build's parser records: a hidden row or column with no
/// authored size. A state written before it did carries no entry at all, so
/// materializing one has to put them back or the row silently unhides.
#[derive(Clone, Debug, Default)]
struct HiddenDimensions {
    col_widths: BTreeMap<u32, f64>,
    row_heights: BTreeMap<u32, f64>,
}

static EMPTY_ROW_STYLES: BTreeMap<u32, u32> = BTreeMap::new();

/// Each sheet's row defaults, or none at all when no sheet has any, as a base
/// releases before them recorded.
fn base_row_styles(model: &WorkbookModel) -> Vec<BTreeMap<u32, u32>> {
    if model.sheets.iter().all(|sheet| sheet.row_styles.is_empty()) {
        return Vec::new();
    }
    model
        .sheets
        .iter()
        .map(|sheet| sheet.row_styles.clone())
        .collect()
}

fn materialize_style_defaults<T: ReadTxn>(
    sheet: &mut Sheet,
    map: &MapRef,
    txn: &T,
) -> Result<(), String> {
    if let Some(Out::Any(value)) = map.get(txn, ROW_STYLES) {
        let rows = any_values(&value, "row styles")?;
        if rows.len() > xlsx_parse::MAX_ROW_STYLES {
            return Err("row style count exceeded cap".into());
        }
        sheet.row_styles.clear();
        for value in rows {
            let row = any_values(value, "row style")?;
            if row.len() != 2 {
                return Err("invalid row style".into());
            }
            let index = any_u32(&row[0], "row index")?;
            let style = any_u32(&row[1], "row style")?;
            if index >= MAX_ROWS
                || sheet
                    .row_styles
                    .last_key_value()
                    .is_some_and(|(&last, _)| index <= last)
            {
                return Err("row styles are out of range or reordered".into());
            }
            sheet.row_styles.insert(index, style);
        }
    } else if map.get(txn, ROW_STYLES).is_some() {
        return Err("invalid row styles".into());
    }
    if let Some(Out::Any(value)) = map.get(txn, COL_STYLES) {
        let columns = any_values(&value, "column styles")?;
        if columns.len() > xlsx_parse::MAX_COL_STYLES {
            return Err("column style count exceeded cap".into());
        }
        sheet.col_styles.clear();
        for value in columns {
            let run = any_values(value, "column style")?;
            if run.len() != 3 {
                return Err("invalid column style".into());
            }
            let first = any_u32(&run[0], "first column")?;
            let last = any_u32(&run[1], "last column")?;
            let xf = any_u32(&run[2], "column style")?;
            if first > last || last >= MAX_COLS {
                return Err("column styles are out of range".into());
            }
            sheet.col_styles.push(ColStyle { first, last, xf });
        }
    } else if map.get(txn, COL_STYLES).is_some() {
        return Err("invalid column styles".into());
    }
    Ok(())
}

static EMPTY_HIDDEN_DIMENSIONS: HiddenDimensions = HiddenDimensions {
    col_widths: BTreeMap::new(),
    row_heights: BTreeMap::new(),
};

fn hidden_dimensions(
    model: &WorkbookModel,
    legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
) -> Vec<HiddenDimensions> {
    model
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| {
            let Some(legacy) = legacy_dimensions.get(index) else {
                return HiddenDimensions::default();
            };
            let only_current = |current: &BTreeMap<u32, f64>, before: &BTreeMap<u32, f64>| {
                current
                    .iter()
                    .filter(|(at, _)| !before.contains_key(at))
                    .map(|(at, size)| (*at, *size))
                    .collect()
            };
            HiddenDimensions {
                col_widths: only_current(&sheet.col_widths, &legacy.col_widths),
                row_heights: only_current(&sheet.row_heights, &legacy.row_heights),
            }
        })
        .collect()
}

/// The same workbook as an earlier release would have modelled it, or `None`
/// when no sheet's dimensions changed meaning and the two agree already.
#[cfg(test)]
fn model_with_legacy_dimensions(
    model: &WorkbookModel,
    legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
) -> Option<WorkbookModel> {
    let differs = model
        .sheets
        .iter()
        .zip(legacy_dimensions)
        .any(|(sheet, legacy)| {
            sheet.col_widths != legacy.col_widths || sheet.row_heights != legacy.row_heights
        });
    if !differs {
        return None;
    }
    let mut legacy_model = model.clone();
    for (sheet, legacy) in legacy_model.sheets.iter_mut().zip(legacy_dimensions) {
        sheet.col_widths.clone_from(&legacy.col_widths);
        sheet.row_heights.clone_from(&legacy.row_heights);
    }
    Some(legacy_model)
}

/// The collaboration state a release before format facets persisted for
/// `model`, which must already hold that release's reading of its styles.
#[cfg(test)]
pub(crate) fn pre_facet_state_for_test(model: &WorkbookModel) -> Vec<u8> {
    let mut base = WorkbookBase::from_model(model).unwrap();
    (base.fingerprint, base.bootstrap_client_id) =
        fingerprint_model_with_schema(model, 6, true, true).unwrap();
    let doc = Doc::with_client_id(base.bootstrap_client_id);
    let keys = (0..model.sheets.len())
        .map(|index| format!("sheet:{index}"))
        .collect::<Vec<_>>();
    seed_with_format_entries(&doc, &base, model, &keys, |styles, index| {
        cell_format_entry(&styles.cell_format(Some(index)))
    })
    .unwrap();
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

#[cfg(test)]
fn fingerprint_model_for_schema(
    model: &WorkbookModel,
    schema_version: i64,
) -> Result<(String, u64), String> {
    fingerprint_model_with_schema(model, schema_version, schema_version >= 4, true)
}

#[cfg(test)]
fn fingerprint_model_with_schema(
    model: &WorkbookModel,
    schema_version: i64,
    include_defined_names: bool,
    include_col_styles: bool,
) -> Result<(String, u64), String> {
    fingerprint_model_with_schema_oracle(
        model,
        schema_version,
        include_defined_names,
        include_col_styles,
    )
}

/// The replica schema 6 seeded for `model`, which read builtin ids below 164
/// as their builtins whatever the workbook declared for them.
#[cfg(test)]
pub(crate) fn schema_6_authority(model: &WorkbookModel, client_id: u64) -> WorkbookAuthority {
    const VERSION: i64 = 6;
    let (fingerprint, bootstrap_client_id) =
        fingerprint_model_with_schema(model, VERSION, true, true).unwrap();
    let mut base = WorkbookBase::from_model(model).unwrap();
    base.fingerprint = fingerprint;
    base.bootstrap_client_id = bootstrap_client_id;
    base.fingerprints.retain(|version, _| *version <= VERSION);
    let mut builtin = model.clone();
    builtin.styles.num_fmts.retain(|(id, _)| *id >= 164);
    let doc = Doc::with_client_id(bootstrap_client_id);
    let keys = (0..model.sheets.len())
        .map(|index| format!("sheet:{index}"))
        .collect::<Vec<_>>();
    seed(&doc, &base, &builtin, &keys).unwrap();
    {
        let mut txn = doc.transact_mut_with("test:schema-6");
        let meta = txn.get_map(META).unwrap();
        meta.try_update(&mut txn, "schemaVersion", VERSION);
    }
    let update = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let replica = Doc::with_client_id(client_id);
    hydrate_doc(&replica, &update).unwrap();
    let authority = WorkbookAuthority::hydrated(replica, Arc::new(base), 0);
    authority.materialize().unwrap();
    authority
}

#[derive(Clone, Copy)]
struct FingerprintStyles<'a> {
    styles: &'a Stylesheet,
    indexed_colors: &'a [String],
    xf_facets: bool,
    as_written: bool,
}

impl<'a> FingerprintStyles<'a> {
    fn new(styles: &'a Stylesheet) -> Self {
        Self {
            styles,
            indexed_colors: &styles.indexed_colors,
            xf_facets: true,
            as_written: false,
        }
    }

    fn with_written_fills(self) -> Self {
        Self {
            as_written: true,
            ..self
        }
    }

    fn without_indexed_colors(self) -> Self {
        Self {
            indexed_colors: &[],
            ..self
        }
    }

    fn without_xf_facets(self) -> Self {
        Self {
            xf_facets: false,
            ..self
        }
    }

    fn cell_xfs(&self) -> Cow<'a, [xlsx_model::Xf]> {
        if self.xf_facets {
            Cow::Borrowed(&self.styles.cell_xfs)
        } else {
            let xfs = &self.styles.cell_xfs;
            Cow::Owned(xfs.iter().map(xlsx_model::Xf::released_reading).collect())
        }
    }
}

#[cfg(test)]
fn pre_facet_styles(styles: &Stylesheet) -> Stylesheet {
    let mut pre_facet = styles.clone();
    let xfs = &styles.cell_xfs;
    pre_facet.cell_xfs = xfs.iter().map(xlsx_model::Xf::released_reading).collect();
    pre_facet
}

impl PartialEq for FingerprintStyles<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.styles.fonts == other.styles.fonts
            && self.styles.fills == other.styles.fills
            && self.styles.borders == other.styles.borders
            && self.cell_xfs() == other.cell_xfs()
            && self.styles.num_fmts == other.styles.num_fmts
            && self.styles.theme == other.styles.theme
            && self.indexed_colors == other.indexed_colors
            && self.as_written == other.as_written
    }
}

impl serde::Serialize for FingerprintStyles<'_> {
    fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct(
            "Stylesheet",
            if self.indexed_colors.is_empty() { 6 } else { 7 },
        )?;
        state.serialize_field("fonts", &self.styles.fonts)?;
        if self.as_written {
            state.serialize_field("fills", &self.styles.fills)?;
            state.serialize_field("borders", &self.styles.borders)?;
        } else {
            let fills: Vec<_> = self.styles.fills.iter().map(released_fill).collect();
            state.serialize_field("fills", &fills)?;
            let borders: Vec<_> = self.styles.borders.iter().map(released_border).collect();
            state.serialize_field("borders", &borders)?;
        }
        state.serialize_field("cell_xfs", &*self.cell_xfs())?;
        state.serialize_field("num_fmts", &self.styles.num_fmts)?;
        state.serialize_field("theme", &self.styles.theme)?;
        if !self.indexed_colors.is_empty() {
            state.serialize_field("indexed_colors", self.indexed_colors)?;
        }
        state.end()
    }
}

/// The domain of the current fingerprint of styles holding patterned fills or
/// extended border styles, hashed as written. Releases before them hashed their
/// reading under the per-schema domains.
const PATTERN_FILLS_FINGERPRINT_DOMAIN: &[u8] = b"betteroffice-xlsx-yrs-v7-pattern-fills";

/// Heads the text of a fingerprint hashed under that domain, so a state says
/// its writer read the fills and borders as written.
const PATTERN_FILLS_FINGERPRINT_TAG: &str = "pattern-fills:";

/// Whether releases before patterned fills read `styles` as written.
pub(crate) fn released_styles(styles: &Stylesheet) -> bool {
    use snapshot::ReleasedLayout;
    styles.fills.iter().all(ReleasedLayout::released_layout)
        && styles.borders.iter().all(ReleasedLayout::released_layout)
}

/// A format as releases before patterned fills read it.
fn released_format(format: &CellFormat) -> CellFormat {
    CellFormat {
        fill: released_fill(&format.fill).into_owned(),
        border: released_border(&format.border).into_owned(),
        ..format.clone()
    }
}

/// A fill as releases before patterned fills read it: a pattern as its solid
/// foreground. Base fingerprints hash fills this way, so a source keeps the
/// fingerprint and bootstrap client that those releases persisted.
fn released_fill(fill: &xlsx_model::Fill) -> Cow<'_, xlsx_model::Fill> {
    match fill {
        xlsx_model::Fill::Pattern { fg, .. } => Cow::Owned(xlsx_model::Fill::Solid(
            fg.clone().unwrap_or(xlsx_model::Color::Auto),
        )),
        fill => Cow::Borrowed(fill),
    }
}

/// A border as releases before the extended dash styles read it.
fn released_border(border: &xlsx_model::Border) -> Cow<'_, xlsx_model::Border> {
    let edges = [&border.left, &border.right, &border.top, &border.bottom];
    if edges
        .into_iter()
        .flatten()
        .all(|edge| released_border_style(edge.style) == edge.style)
    {
        return Cow::Borrowed(border);
    }
    let edge = |edge: &Option<xlsx_model::BorderEdge>| {
        edge.as_ref().map(|edge| xlsx_model::BorderEdge {
            style: released_border_style(edge.style),
            color: edge.color.clone(),
        })
    };
    Cow::Owned(xlsx_model::Border {
        left: edge(&border.left),
        right: edge(&border.right),
        top: edge(&border.top),
        bottom: edge(&border.bottom),
    })
}

fn released_border_style(style: xlsx_model::BorderStyle) -> xlsx_model::BorderStyle {
    use xlsx_model::BorderStyle;
    match style {
        BorderStyle::MediumDashed | BorderStyle::MediumDashDot | BorderStyle::MediumDashDotDot => {
            BorderStyle::Medium
        }
        BorderStyle::DashDot | BorderStyle::DashDotDot | BorderStyle::SlantDashDot => {
            BorderStyle::Dashed
        }
        style => style,
    }
}

type FingerprintKey = (usize, i64, bool, bool, bool, bool, FingerprintReading);

struct FingerprintBuilder<'a> {
    model: &'a WorkbookModel,
    legacy_dimensions: &'a [xlsx_parse::LegacySheetDimensions],
    has_style_defaults: bool,
    styles: Vec<(FingerprintStyles<'a>, usize)>,
    style_bytes: Vec<Vec<u8>>,
    hashes: BTreeMap<FingerprintKey, (String, u64)>,
    planned: Option<Vec<(FingerprintKey, FingerprintStyles<'a>)>>,
}

impl<'a> FingerprintBuilder<'a> {
    fn new(
        model: &'a WorkbookModel,
        legacy_dimensions: &'a [xlsx_parse::LegacySheetDimensions],
    ) -> Self {
        let differs = model
            .sheets
            .iter()
            .zip(legacy_dimensions)
            .any(|(sheet, legacy)| {
                sheet.col_widths != legacy.col_widths || sheet.row_heights != legacy.row_heights
            });
        Self {
            model,
            legacy_dimensions: if differs { legacy_dimensions } else { &[] },
            has_style_defaults: has_style_defaults(model),
            styles: Vec::new(),
            style_bytes: Vec::new(),
            hashes: BTreeMap::new(),
            planned: None,
        }
    }

    fn style_fingerprints(
        &mut self,
        stylesheet: &'a Stylesheet,
        legacy_styles: Option<&'a Stylesheet>,
    ) -> Result<BTreeMap<i64, Vec<String>>, String> {
        let styles = FingerprintStyles::new(stylesheet);
        let current_styles = if released_styles(stylesheet) {
            styles
        } else {
            styles.with_written_fills()
        };
        let mut fingerprints =
            self.accepted_fingerprints(styles.without_xf_facets(), legacy_styles)?;
        let released = fingerprints.remove(&SCHEMA_VERSION).unwrap_or_default();
        let mut accepted = self.current_fingerprints(current_styles, legacy_styles)?;
        if current_styles.as_written {
            for reading in [styles, current_styles.without_xf_facets()] {
                for fingerprint in self.current_fingerprints(reading, legacy_styles)? {
                    if !accepted.contains(&fingerprint) {
                        accepted.push(fingerprint);
                    }
                }
            }
        }
        for fingerprint in released {
            if !accepted.contains(&fingerprint) {
                accepted.push(fingerprint);
            }
        }
        fingerprints.insert(SCHEMA_VERSION, accepted);
        Ok(fingerprints)
    }

    /// Runs `build` once to learn the fingerprints it needs, hashes them all
    /// in one pass over the cells, then runs it again on the results.
    fn batched<T>(
        &mut self,
        mut build: impl FnMut(&mut Self) -> Result<T, String>,
    ) -> Result<T, String> {
        self.planned = Some(Vec::new());
        let planning = build(self);
        let planned = self.planned.take().unwrap_or_default();
        planning?;
        self.compute(planned)?;
        build(self)
    }

    fn compute(
        &mut self,
        planned: Vec<(FingerprintKey, FingerprintStyles<'a>)>,
    ) -> Result<(), String> {
        if planned.is_empty() {
            return Ok(());
        }
        let specs = planned
            .iter()
            .map(|&(key, styles)| {
                let (
                    _,
                    schema_version,
                    include_defined_names,
                    include_col_styles,
                    include_row_styles,
                    legacy_dimensions,
                    reading,
                ) = key;
                FingerprintSpec {
                    styles,
                    schema_version,
                    include_defined_names,
                    include_col_styles,
                    include_row_styles,
                    legacy_dimensions: if legacy_dimensions {
                        self.legacy_dimensions
                    } else {
                        &[]
                    },
                    reading,
                }
            })
            .collect::<Vec<_>>();
        let hashes = fingerprint_model(self.model, &specs)?;
        for ((key, _), hash) in planned.into_iter().zip(hashes) {
            self.hashes.insert(key, hash);
        }
        Ok(())
    }

    fn current(&mut self, styles: FingerprintStyles<'a>) -> Result<(String, u64), String> {
        self.hash(styles, SCHEMA_VERSION, true, true, true, false)
    }

    fn fingerprint(
        &mut self,
        styles: FingerprintStyles<'a>,
        version: i64,
        include_defined_names: bool,
        include_col_styles: bool,
        include_row_styles: bool,
        legacy_dimensions: bool,
    ) -> Result<(String, u64), String> {
        let defaults = include_col_styles
            && include_row_styles
            && version >= CHARTS_SCHEMA_VERSION
            && self.has_style_defaults;
        let reading =
            FingerprintReading::of(styles.xf_facets, styles.as_written).with_defaults(defaults);
        self.hash_as(
            styles,
            version,
            include_defined_names,
            (include_col_styles, include_row_styles, legacy_dimensions),
            reading.legacy().unwrap_or(reading),
        )
    }

    fn hash(
        &mut self,
        styles: FingerprintStyles<'a>,
        version: i64,
        include_defined_names: bool,
        include_col_styles: bool,
        include_row_styles: bool,
        legacy_dimensions: bool,
    ) -> Result<(String, u64), String> {
        let defaults = include_col_styles
            && include_row_styles
            && version >= CHARTS_SCHEMA_VERSION
            && self.has_style_defaults;
        let reading =
            FingerprintReading::of(styles.xf_facets, styles.as_written).with_defaults(defaults);
        self.hash_as(
            styles,
            version,
            include_defined_names,
            (include_col_styles, include_row_styles, legacy_dimensions),
            reading,
        )
    }

    fn hash_as(
        &mut self,
        styles: FingerprintStyles<'a>,
        version: i64,
        include_defined_names: bool,
        (include_col_styles, include_row_styles, legacy_dimensions): (bool, bool, bool),
        reading: FingerprintReading,
    ) -> Result<(String, u64), String> {
        let style_index = match self.styles.iter().find(|(known, _)| {
            std::ptr::eq(known.styles, styles.styles)
                && std::ptr::eq(known.indexed_colors, styles.indexed_colors)
                && known.xf_facets == styles.xf_facets
                && known.as_written == styles.as_written
        }) {
            Some((_, index)) => *index,
            None => {
                let bytes = serde_json::to_vec(&styles)
                    .map_err(|error| format!("cannot fingerprint workbook base: {error}"))?;
                let index = match self.style_bytes.iter().position(|known| *known == bytes) {
                    Some(index) => index,
                    None => {
                        self.style_bytes.push(bytes);
                        self.style_bytes.len() - 1
                    }
                };
                self.styles.push((styles, index));
                index
            }
        };
        let include_col_styles =
            include_col_styles && version >= CHARTS_SCHEMA_VERSION && self.has_style_defaults;
        let include_row_styles = include_row_styles && include_col_styles;
        let legacy_dimensions = legacy_dimensions && !self.legacy_dimensions.is_empty();
        let key = (
            style_index,
            version,
            include_defined_names,
            include_col_styles,
            include_row_styles,
            legacy_dimensions,
            reading,
        );
        if let Some(hash) = self.hashes.get(&key) {
            return Ok(hash.clone());
        }
        if let Some(planned) = &mut self.planned {
            if !planned.iter().any(|(known, _)| *known == key) {
                planned.push((key, styles));
            }
            return Ok((String::new(), 0));
        }
        self.compute(vec![(key, styles)])?;
        Ok(self.hashes[&key].clone())
    }

    /// The current schema's readings for column styles, applyX flags and the
    /// palette, under the current domain: the exact fingerprint comes first.
    fn current_fingerprints(
        &mut self,
        styles: FingerprintStyles<'a>,
        legacy_styles: Option<&'a Stylesheet>,
    ) -> Result<Vec<String>, String> {
        let mut accepted = Vec::new();
        let mut readings = Vec::new();
        for (columns, rows) in [(true, true), (false, false), (true, false)] {
            let (current, _) = self.hash(styles, SCHEMA_VERSION, true, columns, rows, false)?;
            if !accepted.contains(&current) {
                accepted.push(current);
            }
            let defaults = columns && rows && self.has_style_defaults;
            if let Some(legacy) = FingerprintReading::of(styles.xf_facets, styles.as_written)
                .with_defaults(defaults)
                .legacy()
            {
                let (fingerprint, _) =
                    self.hash_as(styles, SCHEMA_VERSION, true, (columns, rows, false), legacy)?;
                if !accepted.contains(&fingerprint) {
                    accepted.push(fingerprint);
                }
            }
        }
        if let Some(legacy) = legacy_styles
            .map(|legacy| FingerprintStyles {
                styles: legacy,
                indexed_colors: &legacy.indexed_colors,
                xf_facets: styles.xf_facets,
                as_written: styles.as_written,
            })
            .filter(|legacy| *legacy != styles)
        {
            readings.extend(self.current_fingerprints(legacy, None)?);
        }
        if !styles.indexed_colors.is_empty() {
            let colorless =
                self.current_fingerprints(styles.without_indexed_colors(), legacy_styles)?;
            readings.extend(colorless);
        }
        for fingerprint in readings {
            if !accepted.contains(&fingerprint) {
                accepted.push(fingerprint);
            }
        }
        Ok(accepted)
    }

    fn accepted_fingerprints(
        &mut self,
        styles: FingerprintStyles<'a>,
        legacy_styles: Option<&'a Stylesheet>,
    ) -> Result<BTreeMap<i64, Vec<String>>, String> {
        let mut fingerprints = BTreeMap::new();
        for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
            let (fingerprint, _) =
                self.fingerprint(styles, version, version >= 4, true, true, false)?;
            fingerprints.insert(version, vec![fingerprint]);
        }
        let (defined_names_v3, _) = self.fingerprint(styles, 3, true, true, true, false)?;
        append_fingerprint(&mut fingerprints, 3, defined_names_v3);
        for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
            let (without_col_styles, _) =
                self.fingerprint(styles, version, version >= 4, false, false, false)?;
            append_fingerprint(&mut fingerprints, version, without_col_styles);
            let (without_row_styles, _) =
                self.fingerprint(styles, version, version >= 4, true, false, false)?;
            append_fingerprint(&mut fingerprints, version, without_row_styles);
        }
        if !self.legacy_dimensions.is_empty() {
            for version in MIN_SUPPORTED_SCHEMA_VERSION..SCHEMA_VERSION {
                let (legacy_fingerprint, _) =
                    self.fingerprint(styles, version, version >= 4, true, true, true)?;
                let (without_col_styles, _) =
                    self.fingerprint(styles, version, version >= 4, false, false, true)?;
                for fingerprint in [legacy_fingerprint, without_col_styles] {
                    append_fingerprint(&mut fingerprints, version, fingerprint);
                }
            }
            let (defined_names_v3, _) = self.fingerprint(styles, 3, true, true, true, true)?;
            append_fingerprint(&mut fingerprints, 3, defined_names_v3);
        }
        if let Some(legacy) = legacy_styles
            .map(|legacy| FingerprintStyles::new(legacy).without_xf_facets())
            .filter(|legacy| *legacy != styles)
        {
            let legacy_fingerprints = self.accepted_fingerprints(legacy, None)?;
            for (version, accepted) in legacy_fingerprints {
                for fingerprint in accepted {
                    append_fingerprint(&mut fingerprints, version, fingerprint);
                }
            }
        }
        if !styles.indexed_colors.is_empty() {
            let legacy_fingerprints =
                self.accepted_fingerprints(styles.without_indexed_colors(), legacy_styles)?;
            for (version, accepted) in legacy_fingerprints {
                for fingerprint in accepted {
                    append_fingerprint(&mut fingerprints, version, fingerprint);
                }
            }
        }
        Ok(fingerprints)
    }
}

/// The domain of a base fingerprint that vouches for a shared state carrying
/// row and column defaults. Releases before them hashed the same workbooks under
/// the per-schema domains, which stay accepted for the states they persisted.
const LEGACY_STYLE_DEFAULTS_FINGERPRINT_DOMAIN: &[u8] = b"betteroffice-xlsx-yrs-v6-style-defaults";
const LEGACY_FACETS_FINGERPRINT_DOMAIN: &[u8] = b"betteroffice-xlsx-yrs-v6-xf-facets";
const LEGACY_PATTERN_FILLS_FINGERPRINT_DOMAIN: &[u8] = b"betteroffice-xlsx-yrs-v6-pattern-fills";
const LEGACY_FACETS_PATTERN_FILLS_FINGERPRINT_DOMAIN: &[u8] =
    b"betteroffice-xlsx-yrs-v6-xf-facets-pattern-fills";
const STYLE_DEFAULTS_FINGERPRINT_DOMAIN: &[u8] = b"betteroffice-xlsx-yrs-v7-style-defaults";
const FACETS_STYLE_DEFAULTS_FINGERPRINT_DOMAIN: &[u8] =
    b"betteroffice-xlsx-yrs-v7-xf-facets-style-defaults";
const PATTERNS_STYLE_DEFAULTS_FINGERPRINT_DOMAIN: &[u8] =
    b"betteroffice-xlsx-yrs-v7-pattern-fills-style-defaults";
const FACETS_PATTERNS_STYLE_DEFAULTS_FINGERPRINT_DOMAIN: &[u8] =
    b"betteroffice-xlsx-yrs-v7-xf-facets-pattern-fills-style-defaults";

fn has_style_defaults(model: &WorkbookModel) -> bool {
    model
        .sheets
        .iter()
        .any(|sheet| !sheet.col_styles.is_empty() || !sheet.row_styles.is_empty())
}

const FACETS_FINGERPRINT_DOMAIN: &[u8] = b"betteroffice-xlsx-yrs-v7-xf-facets";
const FACETS_FINGERPRINT_TAG: &str = "xf-facets:";
const FACETS_PATTERN_FILLS_FINGERPRINT_DOMAIN: &[u8] =
    b"betteroffice-xlsx-yrs-v7-xf-facets-pattern-fills";
const FACETS_PATTERN_FILLS_FINGERPRINT_TAG: &str = "xf-facets-pattern-fills:";

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum FingerprintReading {
    Released,
    Facets,
    Patterns,
    FacetsPatterns,
    Defaults,
    FacetsDefaults,
    PatternsDefaults,
    FacetsPatternsDefaults,
    LegacyDefaults,
    LegacyFacets,
    LegacyPatterns,
    LegacyFacetsPatterns,
}

impl FingerprintReading {
    fn of(facets: bool, patterns: bool) -> Self {
        match (facets, patterns) {
            (false, false) => Self::Released,
            (true, false) => Self::Facets,
            (false, true) => Self::Patterns,
            (true, true) => Self::FacetsPatterns,
        }
    }

    fn with_defaults(self, defaults: bool) -> Self {
        match (self, defaults) {
            (Self::Released, true) => Self::Defaults,
            (Self::Facets, true) => Self::FacetsDefaults,
            (Self::Patterns, true) => Self::PatternsDefaults,
            (Self::FacetsPatterns, true) => Self::FacetsPatternsDefaults,
            _ => self,
        }
    }

    fn legacy(self) -> Option<Self> {
        match self {
            Self::Released => Some(Self::Released),
            Self::Defaults => Some(Self::LegacyDefaults),
            Self::Facets => Some(Self::LegacyFacets),
            Self::Patterns => Some(Self::LegacyPatterns),
            Self::FacetsPatterns => Some(Self::LegacyFacetsPatterns),
            _ => None,
        }
    }

    fn has_facets(self) -> bool {
        matches!(
            self,
            Self::Facets
                | Self::FacetsPatterns
                | Self::FacetsDefaults
                | Self::FacetsPatternsDefaults
                | Self::LegacyFacets
                | Self::LegacyFacetsPatterns
        )
    }

    fn as_written(self) -> bool {
        matches!(
            self,
            Self::Patterns
                | Self::FacetsPatterns
                | Self::PatternsDefaults
                | Self::FacetsPatternsDefaults
                | Self::LegacyPatterns
                | Self::LegacyFacetsPatterns
        )
    }

    fn tag(self) -> &'static str {
        match self {
            Self::Released | Self::LegacyDefaults => "",
            Self::Facets | Self::LegacyFacets => FACETS_FINGERPRINT_TAG,
            Self::Patterns | Self::LegacyPatterns => PATTERN_FILLS_FINGERPRINT_TAG,
            Self::FacetsPatterns | Self::LegacyFacetsPatterns => {
                FACETS_PATTERN_FILLS_FINGERPRINT_TAG
            }
            Self::Defaults => "style-defaults:",
            Self::FacetsDefaults => "xf-facets-style-defaults:",
            Self::PatternsDefaults => "pattern-fills-style-defaults:",
            Self::FacetsPatternsDefaults => "xf-facets-pattern-fills-style-defaults:",
        }
    }
}

const TAGGED_FINGERPRINT_READINGS: [FingerprintReading; 7] = [
    FingerprintReading::FacetsPatternsDefaults,
    FingerprintReading::FacetsDefaults,
    FingerprintReading::PatternsDefaults,
    FingerprintReading::Defaults,
    FingerprintReading::FacetsPatterns,
    FingerprintReading::Facets,
    FingerprintReading::Patterns,
];

fn fingerprint_reading(fingerprint: &str) -> FingerprintReading {
    TAGGED_FINGERPRINT_READINGS
        .into_iter()
        .find(|reading| fingerprint.starts_with(reading.tag()))
        .unwrap_or(FingerprintReading::Released)
}

fn fingerprint_text(reading: FingerprintReading, digest: String) -> String {
    format!("{}{digest}", reading.tag())
}

fn released_fingerprint(fingerprint: &str) -> bool {
    !fingerprint_reading(fingerprint).has_facets()
}

fn fingerprint_domain(schema_version: i64, reading: FingerprintReading) -> &'static [u8] {
    match reading {
        FingerprintReading::Facets => FACETS_FINGERPRINT_DOMAIN,
        FingerprintReading::Patterns => PATTERN_FILLS_FINGERPRINT_DOMAIN,
        FingerprintReading::FacetsPatterns => FACETS_PATTERN_FILLS_FINGERPRINT_DOMAIN,
        FingerprintReading::Defaults => STYLE_DEFAULTS_FINGERPRINT_DOMAIN,
        FingerprintReading::FacetsDefaults => FACETS_STYLE_DEFAULTS_FINGERPRINT_DOMAIN,
        FingerprintReading::PatternsDefaults => PATTERNS_STYLE_DEFAULTS_FINGERPRINT_DOMAIN,
        FingerprintReading::FacetsPatternsDefaults => {
            FACETS_PATTERNS_STYLE_DEFAULTS_FINGERPRINT_DOMAIN
        }
        FingerprintReading::LegacyDefaults => LEGACY_STYLE_DEFAULTS_FINGERPRINT_DOMAIN,
        FingerprintReading::LegacyFacets => LEGACY_FACETS_FINGERPRINT_DOMAIN,
        FingerprintReading::LegacyPatterns => LEGACY_PATTERN_FILLS_FINGERPRINT_DOMAIN,
        FingerprintReading::LegacyFacetsPatterns => LEGACY_FACETS_PATTERN_FILLS_FINGERPRINT_DOMAIN,
        FingerprintReading::Released => match schema_version {
            3 => b"betteroffice-xlsx-yrs-v3",
            4 => b"betteroffice-xlsx-yrs-v4",
            5 => b"betteroffice-xlsx-yrs-v5",
            _ => b"betteroffice-xlsx-yrs-v6",
        },
    }
}

fn append_fingerprint(fingerprints: &mut BTreeMap<i64, Vec<String>>, version: i64, hash: String) {
    let accepted = fingerprints.entry(version).or_default();
    if !accepted.contains(&hash) {
        accepted.push(hash);
    }
}

/// One base fingerprint: the styles reading, schema and readings it hashes.
struct FingerprintSpec<'a> {
    styles: FingerprintStyles<'a>,
    schema_version: i64,
    include_defined_names: bool,
    include_col_styles: bool,
    include_row_styles: bool,
    legacy_dimensions: &'a [xlsx_parse::LegacySheetDimensions],
    reading: FingerprintReading,
}

const FINGERPRINT_CHUNK_BYTES: usize = 64 * 1024;

/// Hashes every spec in a single pass over the cells, which all of them hash
/// alike.
fn fingerprint_model(
    model: &WorkbookModel,
    specs: &[FingerprintSpec<'_>],
) -> Result<Vec<(String, u64)>, String> {
    #[cfg(test)]
    FINGERPRINT_PASSES.set(FINGERPRINT_PASSES.get() + 1);
    let mut hashers = specs
        .iter()
        .map(|spec| {
            validate_schema_version(spec.schema_version)?;
            let mut hasher = Sha256::new();
            hasher.update(fingerprint_domain(spec.schema_version, spec.reading));
            let base = if spec.include_defined_names {
                serde_json::to_vec(&(
                    model.date_system,
                    &model.defined_names,
                    &model.shared_strings,
                    &spec.styles,
                ))
            } else {
                serde_json::to_vec(&(model.date_system, &model.shared_strings, &spec.styles))
            }
            .map_err(|error| format!("cannot fingerprint workbook base: {error}"))?;
            hash_bytes(&mut hasher, &base);
            hash_u64(&mut hasher, model.sheets.len() as u64);
            Ok(hasher)
        })
        .collect::<Result<Vec<_>, String>>()?;
    let has_col_styles = model
        .sheets
        .iter()
        .any(|sheet| !sheet.col_styles.is_empty());
    let hashes_col_styles = |spec: &FingerprintSpec<'_>| {
        spec.include_col_styles && spec.schema_version >= CHARTS_SCHEMA_VERSION && has_col_styles
    };
    let mut shared = Vec::new();
    for (index, sheet) in model.sheets.iter().enumerate() {
        hash_bytes(&mut shared, sheet.name.as_bytes());
        hash_u64(&mut shared, sheet.iter_cells().count() as u64);
        for (at, cell) in sheet.iter_cells() {
            xlsx_model::progress::tick();
            hash_u32(&mut shared, at.row);
            hash_u32(&mut shared, at.col);
            hash_cell_value(&mut shared, &cell.value);
            match &cell.formula {
                Some(formula) => {
                    shared.push(1);
                    hash_bytes(&mut shared, formula.as_bytes());
                }
                None => shared.push(0),
            }
            match cell.style {
                Some(style) => {
                    shared.push(1);
                    hash_u32(&mut shared, style);
                }
                None => shared.push(0),
            }
            if shared.len() >= FINGERPRINT_CHUNK_BYTES {
                feed_all(&mut hashers, &mut shared);
            }
        }
        hash_u64(&mut shared, sheet.merges.len() as u64);
        for range in &sheet.merges {
            hash_cell_ref(&mut shared, range.start);
            hash_cell_ref(&mut shared, range.end);
        }
        feed_all(&mut hashers, &mut shared);
        let hyperlinks = sheet_json(
            specs
                .iter()
                .any(|spec| spec.schema_version >= HYPERLINK_SCHEMA_VERSION),
            &sheet.hyperlinks,
            "hyperlinks",
        )?;
        let charts = sheet_json(
            specs
                .iter()
                .any(|spec| spec.schema_version >= CHARTS_SCHEMA_VERSION),
            &sheet.charts,
            "charts",
        )?;
        let col_styles = sheet_json(
            specs.iter().any(hashes_col_styles),
            &sheet.col_styles,
            "column styles",
        )?;
        let row_styles = sheet_json(
            specs.iter().any(|spec| spec.include_row_styles)
                && model
                    .sheets
                    .iter()
                    .any(|sheet| !sheet.row_styles.is_empty()),
            &sheet.row_styles,
            "row styles",
        )?;
        for (hasher, spec) in hashers.iter_mut().zip(specs) {
            let (col_widths, row_heights) = match spec.legacy_dimensions.get(index) {
                Some(legacy) => (&legacy.col_widths, &legacy.row_heights),
                None => (&sheet.col_widths, &sheet.row_heights),
            };
            hash_u64(hasher, col_widths.len() as u64);
            for (&column, &width) in col_widths {
                hash_u32(hasher, column);
                hash_u64(hasher, width.to_bits());
            }
            hash_u64(hasher, row_heights.len() as u64);
            for (&row, &height) in row_heights {
                hash_u32(hasher, row);
                hash_u64(hasher, height.to_bits());
            }
            if spec.schema_version >= FREEZE_PANE_SCHEMA_VERSION {
                match sheet.freeze_pane {
                    Some(pane) => {
                        hasher.update([1]);
                        hash_u32(hasher, pane.rows);
                        hash_u32(hasher, pane.cols);
                        hash_cell_ref(hasher, pane.top_left);
                    }
                    None => hasher.update([0]),
                }
            }
            if let Some(bytes) = hyperlinks
                .as_ref()
                .filter(|_| spec.schema_version >= HYPERLINK_SCHEMA_VERSION)
            {
                hash_bytes(hasher, bytes);
            }
            if let Some(bytes) = charts
                .as_ref()
                .filter(|_| spec.schema_version >= CHARTS_SCHEMA_VERSION)
            {
                hash_bytes(hasher, bytes);
            }
            if let Some(bytes) = col_styles.as_ref().filter(|_| hashes_col_styles(spec)) {
                hash_bytes(hasher, bytes);
            }
            if let Some(bytes) = row_styles.as_ref().filter(|_| spec.include_row_styles) {
                hash_bytes(hasher, bytes);
            }
        }
    }
    Ok(hashers
        .into_iter()
        .zip(specs)
        .map(|(hasher, spec)| {
            let digest = hasher.finalize();
            let fingerprint = fingerprint_text(spec.reading, format!("{digest:x}"));
            let mut client_bytes = [0_u8; 8];
            client_bytes.copy_from_slice(&digest[..8]);
            let client = client_id_of_digest(u64::from_be_bytes(client_bytes));
            (fingerprint, client)
        })
        .collect())
}

fn client_id_of_digest(head: u64) -> u64 {
    match head & MAX_SAFE_CLIENT_ID {
        0 => 1,
        client => client,
    }
}

/// The bootstrap client a base fingerprint seeds: the head of its digest.
fn fingerprint_client_id(fingerprint: &str) -> Option<u64> {
    let digest = fingerprint.strip_prefix(fingerprint_reading(fingerprint).tag())?;
    let head = u64::from_str_radix(digest.get(..16)?, 16).ok()?;
    Some(client_id_of_digest(head))
}

fn sheet_json<T: Serialize>(
    needed: bool,
    value: &T,
    what: &str,
) -> Result<Option<Vec<u8>>, String> {
    needed
        .then(|| serde_json::to_vec(value))
        .transpose()
        .map_err(|error| format!("cannot fingerprint sheet {what}: {error}"))
}

fn feed_all(hashers: &mut [Sha256], bytes: &mut Vec<u8>) {
    for hasher in hashers.iter_mut() {
        hasher.update(bytes.as_slice());
    }
    bytes.clear();
}

/// Where fingerprint bytes go: a hasher, or a buffer several hashers share.
trait FingerprintSink {
    fn feed(&mut self, bytes: &[u8]);
}

impl FingerprintSink for Sha256 {
    fn feed(&mut self, bytes: &[u8]) {
        self.update(bytes);
    }
}

impl FingerprintSink for Vec<u8> {
    fn feed(&mut self, bytes: &[u8]) {
        self.extend_from_slice(bytes);
    }
}

fn hash_cell_value(sink: &mut impl FingerprintSink, value: &CellValue) {
    match value {
        CellValue::Empty => sink.feed(&[0]),
        CellValue::Number { value } => {
            sink.feed(&[1]);
            hash_u64(sink, value.to_bits());
        }
        CellValue::Text { value } => {
            sink.feed(&[2]);
            hash_bytes(sink, value.as_bytes());
        }
        CellValue::Bool { value } => sink.feed(&[3, u8::from(*value)]),
        CellValue::Error { value } => {
            sink.feed(&[4]);
            hash_bytes(sink, value.as_str().as_bytes());
        }
    }
}

fn hash_cell_ref(sink: &mut impl FingerprintSink, cell: CellRef) {
    hash_u32(sink, cell.row);
    hash_u32(sink, cell.col);
    sink.feed(&[u8::from(cell.abs_row), u8::from(cell.abs_col)]);
}

fn hash_bytes(sink: &mut impl FingerprintSink, bytes: &[u8]) {
    hash_u64(sink, bytes.len() as u64);
    sink.feed(bytes);
}

fn hash_u32(sink: &mut impl FingerprintSink, value: u32) {
    sink.feed(&value.to_le_bytes());
}

fn hash_u64(sink: &mut impl FingerprintSink, value: u64) {
    sink.feed(&value.to_le_bytes());
}

#[cfg(test)]
impl WorkbookAuthority {
    pub(crate) fn from_source_oracle(
        model: &WorkbookModel,
        client_id: Option<u64>,
        legacy_dimensions: &[xlsx_parse::LegacySheetDimensions],
        legacy_styles: Option<&Stylesheet>,
        legacy_style_readings: &[Stylesheet],
    ) -> Result<Self, AuthorityError> {
        let mut base = WorkbookBase::from_model_with_legacy_dimensions_oracle(
            model,
            legacy_dimensions,
            legacy_styles,
        )
        .map_err(AuthorityError::InvalidState)?;
        for reading in legacy_style_readings {
            let mut legacy_model = model.clone();
            legacy_model.styles = reading.clone();
            let projected = WorkbookBase::from_model_with_legacy_dimensions_oracle(
                &legacy_model,
                legacy_dimensions,
                None,
            )
            .map_err(AuthorityError::InvalidState)?;
            for (version, accepted) in projected.fingerprints {
                for fingerprint in accepted {
                    append_fingerprint(&mut base.fingerprints, version, fingerprint);
                }
            }
        }
        if client_id == Some(base.bootstrap_client_id) {
            return Err(AuthorityError::ClientIdConflict(base.bootstrap_client_id));
        }

        let bootstrap = Doc::with_client_id(base.bootstrap_client_id);
        let keys = (0..model.sheets.len())
            .map(|index| format!("sheet:{index}"))
            .collect::<Vec<_>>();
        seed(&bootstrap, &base, model, &keys).map_err(AuthorityError::InvalidState)?;
        let bootstrap_update = bootstrap
            .transact()
            .encode_state_as_update_v1(&StateVector::default());

        let doc = match client_id {
            Some(client_id) => Doc::with_client_id(client_id),
            None => loop {
                let candidate = Doc::new();
                if candidate.client_id().get() != base.bootstrap_client_id {
                    break candidate;
                }
            },
        };
        hydrate_local_doc(&doc, &bootstrap_update).map_err(AuthorityError::InvalidState)?;
        let authority = Self::hydrated(doc, Arc::new(base), 0);
        authority
            .materialize_internal_oracle(true)
            .map_err(AuthorityError::InvalidState)?;
        Ok(authority)
    }

    pub(crate) fn materialize_oracle(&self) -> Result<WorkbookModel, AuthorityError> {
        self.materialize_internal_oracle(false)
            .map(|(model, _)| model)
            .map_err(AuthorityError::InvalidState)
    }

    pub(crate) fn structure_oracle(&self) -> Result<WorkbookStructure, AuthorityError> {
        self.materialize_internal_oracle(false)
            .map(|(_, structure)| structure)
            .map_err(AuthorityError::InvalidState)
    }

    fn materialize_internal_oracle(
        &self,
        strict: bool,
    ) -> Result<(WorkbookModel, WorkbookStructure), String> {
        let txn = self.doc.transact();
        if strict {
            require_root_keys(&txn, &[CELL_FORMATS, META, SHEET_ORDER, SHEETS])?;
        }
        let meta = txn
            .get_map(META)
            .ok_or_else(|| "missing workbook metadata".to_string())?;
        if strict {
            require_map_keys(
                &meta,
                &txn,
                &[BASE_FINGERPRINT, "schemaVersion", STRUCTURE_GENERATION],
                "workbook metadata",
            )?;
        }
        let version = meta
            .get(&txn, "schemaVersion")
            .and_then(|value| value.cast::<i64>().ok())
            .ok_or_else(|| "missing schema version".to_string())?;
        validate_schema_version(version)?;
        let fingerprint = meta
            .get(&txn, BASE_FINGERPRINT)
            .and_then(|value| value.cast::<String>().ok())
            .ok_or_else(|| "missing workbook base fingerprint".to_string())?;
        if !self.base.accepts_fingerprint(version, &fingerprint) {
            return Err("workbook base fingerprint does not match shared state".to_string());
        }
        let generation = structure_generation(&meta, &txn)?;
        let cell_formats = txn
            .get_map(CELL_FORMATS)
            .ok_or_else(|| "missing cell format catalog".to_string())?;
        let (styles, style_indices) =
            materialize_cell_formats(&cell_formats, &txn, &self.base.styles)?;

        let order = txn
            .get_array(SHEET_ORDER)
            .ok_or_else(|| "missing sheet order".to_string())?;
        let sheets = txn
            .get_map(SHEETS)
            .ok_or_else(|| "missing sheet map".to_string())?;
        let keys = sheet_keys(&order, &txn)?;
        let mut seen = HashSet::with_capacity(keys.len());
        let mut model = self.base.workbook();
        model.styles = styles;
        let mut plain = None;
        let expected_sheet_keys = sheet_schema_keys(version);
        let optional_sheet_keys = sheet_schema_optional_keys(version);
        for key in keys.iter() {
            if !seen.insert(key.clone()) {
                return Err(format!("duplicate sheet key {key}"));
            }
            let sheet_map = sheets
                .get(&txn, key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("missing sheet {key}"))?;
            if strict {
                require_map_keys_with_optional(
                    &sheet_map,
                    &txn,
                    expected_sheet_keys,
                    optional_sheet_keys,
                    &format!("sheet {key}"),
                )?;
            }
            let base_sheet = base_sheet_index(key);
            let freeze_pane = base_sheet
                .and_then(|base| self.base.freeze_panes.get(base))
                .copied()
                .flatten();
            let format = base_sheet
                .and_then(|base| self.base.formats.get(base))
                .copied()
                .unwrap_or_default();
            let col_styles = base_sheet
                .and_then(|base| self.base.col_styles.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let row_styles = base_sheet
                .and_then(|base| self.base.row_styles.get(base))
                .unwrap_or(&EMPTY_ROW_STYLES);
            let hyperlinks = base_sheet
                .and_then(|base| self.base.hyperlinks.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let auto_filter = base_sheet
                .and_then(|base| self.base.auto_filters.get(base))
                .and_then(Option::as_ref);
            let conditional_formats = base_sheet
                .and_then(|base| self.base.conditional_formats.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let charts = base_sheet
                .and_then(|base| self.base.charts.get(base))
                .map(Vec::as_slice)
                .unwrap_or_default();
            let hidden_dimensions = base_sheet
                .and_then(|base| self.base.hidden_dimensions.get(base))
                .unwrap_or(&EMPTY_HIDDEN_DIMENSIONS);
            model.sheets.push(materialize_sheet_oracle(
                &sheet_map,
                &txn,
                &style_indices,
                &mut || *plain.get_or_insert_with(|| model.styles.plain_xf()),
                version,
                SheetFallbacks {
                    freeze_pane,
                    format,
                    col_styles,
                    row_styles,
                    hyperlinks,
                    charts,
                    hidden_dimensions,
                    conditional_formats,
                    auto_filter,
                },
            )?);
        }
        project_shared_frame_anchors(&mut model.sheets);
        let active = keys.iter().cloned().collect::<BTreeSet<_>>();
        let all_keys = sheets
            .keys(&txn)
            .map(str::to_string)
            .collect::<BTreeSet<_>>();
        let mut shared_types = BTreeMap::new();
        let mut baseline_hyperlinks = BTreeMap::new();
        let mut baseline_merges = BTreeMap::new();
        for key in all_keys {
            let sheet_map = sheets
                .get(&txn, &key)
                .and_then(|value| value.cast::<MapRef>().ok())
                .ok_or_else(|| format!("sheet {key} is not a map"))?;
            if strict && !active.contains(&key) {
                require_map_keys_with_optional(
                    &sheet_map,
                    &txn,
                    expected_sheet_keys,
                    optional_sheet_keys,
                    &format!("inactive sheet {key}"),
                )?;
                materialize_sheet_oracle(
                    &sheet_map,
                    &txn,
                    &style_indices,
                    &mut || 0,
                    version,
                    SheetFallbacks::default(),
                )?;
            }
            let Some(Out::Any(baseline)) = sheet_map.get(&txn, MERGES) else {
                return Err(format!("sheet {key} is missing merges"));
            };
            let link_baseline = match sheet_map.get(&txn, HYPERLINKS) {
                Some(Out::Any(Any::String(json))) => decode_hyperlinks(&json)?,
                None if version < HYPERLINK_SCHEMA_VERSION => base_sheet_index(&key)
                    .and_then(|index| self.base.hyperlinks.get(index))
                    .cloned()
                    .unwrap_or_default(),
                _ => return Err(format!("sheet {key} is missing hyperlinks")),
            };
            baseline_hyperlinks.insert(key.clone(), link_baseline);
            baseline_merges.insert(key.clone(), merges_from_any(&baseline)?);
            shared_types.insert(key, sheet_shared_types(&sheet_map, &txn)?);
        }
        let structure = WorkbookStructure {
            generation,
            sheet_keys: keys,
            sheet_names: model
                .sheets
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect(),
            freeze_panes: model.sheets.iter().map(|sheet| sheet.freeze_pane).collect(),
            hyperlinks: model
                .sheets
                .iter()
                .map(|sheet| sheet.hyperlinks.clone())
                .collect(),
            baseline_hyperlinks,
            conditional_formats: model
                .sheets
                .iter()
                .map(|sheet| sheet.conditional_formats.clone())
                .collect(),
            auto_filters: model
                .sheets
                .iter()
                .map(|sheet| sheet.auto_filter.clone())
                .collect(),
            charts: model
                .sheets
                .iter()
                .map(|sheet| sheet.charts.iter().map(ChartIdentity::of).collect())
                .collect(),
            merges: model
                .sheets
                .iter()
                .map(|sheet| sheet.merges.clone())
                .collect(),
            baseline_merges,
            shared_types,
        };
        self.projection_valid.store(true, Ordering::Relaxed);
        Ok((model, structure))
    }
}

#[cfg(test)]
fn materialize_sheet_oracle<T: ReadTxn>(
    sheet_map: &MapRef,
    txn: &T,
    style_indices: &BTreeMap<String, Option<u32>>,
    plain: &mut dyn FnMut() -> u32,
    version: i64,
    fallbacks: SheetFallbacks<'_>,
) -> Result<Sheet, String> {
    let name = sheet_map
        .get(txn, NAME)
        .and_then(|value| value.cast::<String>().ok())
        .ok_or_else(|| "sheet is missing its name".to_string())?;
    let mut sheet = Sheet::new(name);
    let mut cells = BTreeMap::<(u32, u32), Cell>::new();
    let contents = nested_map(sheet_map, txn, CONTENTS)?;
    for (key, value) in contents.iter(txn) {
        let at = parse_cell_key(key)?;
        let Out::Any(value) = value else {
            return Err(format!("cell content {key} is not an atomic value"));
        };
        cells.insert((at.row, at.col), content_from_any(&value)?);
    }
    let styles = nested_map(sheet_map, txn, STYLES)?;
    for (key, value) in styles.iter(txn) {
        let at = parse_cell_key(key)?;
        let style_key = value
            .cast::<String>()
            .map_err(|_| format!("cell style {key} is not a string"))?;
        let style = style_indices
            .get(&style_key)
            .ok_or_else(|| format!("cell style {key} references an unknown format"))?;
        let style = match *style {
            Some(style) => style,
            None => plain(),
        };
        cells.entry((at.row, at.col)).or_default().style = Some(style);
    }
    for ((row, col), cell) in cells {
        sheet.set_cell(CellRef::new(row, col), cell);
    }
    sheet.col_widths = materialize_numbers(
        &nested_map(sheet_map, txn, COL_WIDTHS)?,
        txn,
        MAX_COLS,
        "column width",
    )?;
    sheet.row_heights = materialize_numbers(
        &nested_map(sheet_map, txn, ROW_HEIGHTS)?,
        txn,
        MAX_ROWS,
        "row height",
    )?;
    if version < SCHEMA_VERSION {
        for (&at, &size) in &fallbacks.hidden_dimensions.col_widths {
            sheet.col_widths.entry(at).or_insert(size);
        }
        for (&at, &size) in &fallbacks.hidden_dimensions.row_heights {
            sheet.row_heights.entry(at).or_insert(size);
        }
    }
    sheet.format = fallbacks.format;
    sheet.col_styles = fallbacks.col_styles.to_vec();
    sheet.row_styles = fallbacks.row_styles.clone();
    materialize_style_defaults(&mut sheet, sheet_map, txn)?;
    sheet.freeze_pane = match (version, sheet_map.get(txn, FREEZE_PANE)) {
        (FREEZE_PANE_SCHEMA_VERSION.., Some(Out::Any(value))) => freeze_pane_from_any(&value)?,
        (FREEZE_PANE_SCHEMA_VERSION.., _) => {
            return Err("sheet is missing freeze pane".to_string());
        }
        _ => fallbacks.freeze_pane,
    };
    sheet.hyperlinks = match (version, sheet_map.get(txn, HYPERLINKS)) {
        (HYPERLINK_SCHEMA_VERSION.., Some(Out::Any(Any::String(json)))) => {
            decode_hyperlinks(&json)?
        }
        (HYPERLINK_SCHEMA_VERSION.., Some(_)) => {
            return Err("sheet hyperlinks are not a string".to_string());
        }
        (HYPERLINK_SCHEMA_VERSION.., None) => {
            return Err("sheet is missing hyperlinks".to_string());
        }
        _ => fallbacks.hyperlinks.to_vec(),
    };
    sheet.auto_filter = match sheet_map.get(txn, AUTO_FILTER) {
        Some(Out::Any(Any::String(json))) => serde_json::from_str(&json)
            .map_err(|error| format!("sheet auto filter is invalid: {error}"))?,
        Some(_) => return Err("sheet auto filter is not a string".to_string()),
        None => fallbacks.auto_filter.cloned(),
    };
    sheet.conditional_formats = match sheet_map.get(txn, CONDITIONAL_FORMATS) {
        Some(Out::Any(Any::String(json))) => serde_json::from_str(&json)
            .map_err(|error| format!("sheet conditional formats are invalid: {error}"))?,
        Some(_) => return Err("sheet conditional formats are not a string".to_string()),
        None => fallbacks.conditional_formats.to_vec(),
    };
    sheet.charts = match (version, sheet_map.get(txn, CHARTS)) {
        (CHARTS_SCHEMA_VERSION.., Some(Out::Any(Any::String(json)))) => decode_charts(&json)?,
        (CHARTS_SCHEMA_VERSION.., Some(_)) => {
            return Err("sheet charts are not a string".to_string());
        }
        _ => fallbacks.charts.to_vec(),
    };
    sheet.merges = match sheet_map.get(txn, MERGES) {
        Some(Out::Any(value)) => merges_from_any(&value)?,
        _ => return Err("sheet is missing merges".to_string()),
    };
    Ok(sheet)
}

#[cfg(test)]
fn fingerprint_model_with_schema_oracle(
    model: &WorkbookModel,
    schema_version: i64,
    include_defined_names: bool,
    include_col_styles: bool,
) -> Result<(String, u64), String> {
    fingerprint_model_with_row_styles_oracle(
        model,
        schema_version,
        include_defined_names,
        include_col_styles,
        include_col_styles,
    )
}

#[cfg(test)]
fn fingerprint_model_with_row_styles_oracle(
    model: &WorkbookModel,
    schema_version: i64,
    include_defined_names: bool,
    include_col_styles: bool,
    include_row_styles: bool,
) -> Result<(String, u64), String> {
    let style_defaults = include_col_styles
        && include_row_styles
        && schema_version >= CHARTS_SCHEMA_VERSION
        && model
            .sheets
            .iter()
            .any(|sheet| !sheet.col_styles.is_empty() || !sheet.row_styles.is_empty());
    let domain = match schema_version {
        _ if style_defaults => LEGACY_STYLE_DEFAULTS_FINGERPRINT_DOMAIN,
        3 => b"betteroffice-xlsx-yrs-v3".as_slice(),
        4 => b"betteroffice-xlsx-yrs-v4".as_slice(),
        5 => b"betteroffice-xlsx-yrs-v5".as_slice(),
        _ => b"betteroffice-xlsx-yrs-v6".as_slice(),
    };
    fingerprint_model_with_row_domain_oracle(
        model,
        schema_version,
        include_defined_names,
        include_col_styles,
        include_row_styles,
        domain,
    )
}

#[cfg(test)]
fn fingerprint_model_current_oracle(model: &WorkbookModel) -> Result<(String, u64), String> {
    fingerprint_model_with_row_domain_oracle(
        model,
        SCHEMA_VERSION,
        true,
        true,
        has_style_defaults(model),
        fingerprint_domain(
            SCHEMA_VERSION,
            FingerprintReading::of(true, !released_styles(&model.styles))
                .with_defaults(has_style_defaults(model)),
        ),
    )
}

#[cfg(test)]
fn fingerprint_model_with_domain_oracle(
    model: &WorkbookModel,
    schema_version: i64,
    include_defined_names: bool,
    include_col_styles: bool,
    domain: &[u8],
) -> Result<(String, u64), String> {
    fingerprint_model_with_row_domain_oracle(
        model,
        schema_version,
        include_defined_names,
        include_col_styles,
        false,
        domain,
    )
}

#[cfg(test)]
fn fingerprint_model_with_row_domain_oracle(
    model: &WorkbookModel,
    schema_version: i64,
    include_defined_names: bool,
    include_col_styles: bool,
    include_row_styles: bool,
    domain: &[u8],
) -> Result<(String, u64), String> {
    validate_schema_version(schema_version)?;
    let mut hasher = Sha256::new();
    hasher.update(domain);
    let reading = TAGGED_FINGERPRINT_READINGS
        .into_iter()
        .chain([
            FingerprintReading::LegacyDefaults,
            FingerprintReading::LegacyFacets,
            FingerprintReading::LegacyPatterns,
            FingerprintReading::LegacyFacetsPatterns,
        ])
        .find(|reading| fingerprint_domain(schema_version, *reading) == domain)
        .unwrap_or(FingerprintReading::Released);
    let mut styles = model.styles.clone();
    if !reading.has_facets() {
        styles = pre_facet_styles(&styles);
    }
    if !reading.as_written() {
        styles.fills = styles
            .fills
            .iter()
            .map(|fill| released_fill(fill).into_owned())
            .collect();
        styles.borders = styles
            .borders
            .iter()
            .map(|border| released_border(border).into_owned())
            .collect();
    }
    let base = if include_defined_names {
        serde_json::to_vec(&(
            model.date_system,
            &model.defined_names,
            &model.shared_strings,
            &styles,
        ))
    } else {
        serde_json::to_vec(&(model.date_system, &model.shared_strings, &styles))
    }
    .map_err(|error| format!("cannot fingerprint workbook base: {error}"))?;
    hash_bytes(&mut hasher, &base);
    hash_u64(&mut hasher, model.sheets.len() as u64);
    // a workbook declaring no column style hashes exactly as releases before
    // they were read did, so only a workbook that declares one needs the
    // pre-change fingerprint carried alongside.
    let hash_col_styles = include_col_styles
        && schema_version >= CHARTS_SCHEMA_VERSION
        && model
            .sheets
            .iter()
            .any(|sheet| !sheet.col_styles.is_empty());
    let hash_row_styles = include_row_styles
        && include_col_styles
        && schema_version >= CHARTS_SCHEMA_VERSION
        && model
            .sheets
            .iter()
            .any(|sheet| !sheet.row_styles.is_empty());
    for sheet in &model.sheets {
        hash_bytes(&mut hasher, sheet.name.as_bytes());
        hash_u64(&mut hasher, sheet.iter_cells().count() as u64);
        for (at, cell) in sheet.iter_cells() {
            hash_u32(&mut hasher, at.row);
            hash_u32(&mut hasher, at.col);
            hash_cell_value(&mut hasher, &cell.value);
            match &cell.formula {
                Some(formula) => {
                    hasher.update([1]);
                    hash_bytes(&mut hasher, formula.as_bytes());
                }
                None => hasher.update([0]),
            }
            match cell.style {
                Some(style) => {
                    hasher.update([1]);
                    hash_u32(&mut hasher, style);
                }
                None => hasher.update([0]),
            }
        }
        hash_u64(&mut hasher, sheet.merges.len() as u64);
        for range in &sheet.merges {
            hash_cell_ref(&mut hasher, range.start);
            hash_cell_ref(&mut hasher, range.end);
        }
        hash_u64(&mut hasher, sheet.col_widths.len() as u64);
        for (&column, &width) in &sheet.col_widths {
            hash_u32(&mut hasher, column);
            hash_u64(&mut hasher, width.to_bits());
        }
        hash_u64(&mut hasher, sheet.row_heights.len() as u64);
        for (&row, &height) in &sheet.row_heights {
            hash_u32(&mut hasher, row);
            hash_u64(&mut hasher, height.to_bits());
        }
        if schema_version >= FREEZE_PANE_SCHEMA_VERSION {
            match sheet.freeze_pane {
                Some(pane) => {
                    hasher.update([1]);
                    hash_u32(&mut hasher, pane.rows);
                    hash_u32(&mut hasher, pane.cols);
                    hash_cell_ref(&mut hasher, pane.top_left);
                }
                None => hasher.update([0]),
            }
        }
        if schema_version >= HYPERLINK_SCHEMA_VERSION {
            let hyperlinks = serde_json::to_vec(&sheet.hyperlinks)
                .map_err(|error| format!("cannot fingerprint sheet hyperlinks: {error}"))?;
            hash_bytes(&mut hasher, &hyperlinks);
        }
        if schema_version >= CHARTS_SCHEMA_VERSION {
            let charts = serde_json::to_vec(&sheet.charts)
                .map_err(|error| format!("cannot fingerprint sheet charts: {error}"))?;
            hash_bytes(&mut hasher, &charts);
        }
        if hash_col_styles {
            let col_styles = serde_json::to_vec(&sheet.col_styles)
                .map_err(|error| format!("cannot fingerprint sheet column styles: {error}"))?;
            hash_bytes(&mut hasher, &col_styles);
        }
        if hash_row_styles {
            let rows = serde_json::to_vec(&sheet.row_styles)
                .map_err(|error| format!("cannot fingerprint sheet row styles: {error}"))?;
            hash_bytes(&mut hasher, &rows);
        }
    }
    let digest = hasher.finalize();
    let fingerprint = fingerprint_text(reading, format!("{digest:x}"));
    let mut client_bytes = [0_u8; 8];
    client_bytes.copy_from_slice(&digest[..8]);
    let mut bootstrap_client_id = u64::from_be_bytes(client_bytes) & MAX_SAFE_CLIENT_ID;
    if bootstrap_client_id == 0 {
        bootstrap_client_id = 1;
    }
    Ok((fingerprint, bootstrap_client_id))
}

#[cfg(test)]
fn fingerprint_model_for_schema_oracle(
    model: &WorkbookModel,
    schema_version: i64,
) -> Result<(String, u64), String> {
    fingerprint_model_with_schema_oracle(model, schema_version, schema_version >= 4, true)
}

#[cfg(test)]
pub(crate) mod open_test_cases {
    use super::*;
    use xlsx_model::{AnchorCell, Border, ChartRefKind, Color, Fill, Font, Xf};

    pub(crate) struct OpenCase {
        pub(crate) label: String,
        pub(crate) model: WorkbookModel,
        pub(crate) legacy_dimensions: Vec<xlsx_parse::LegacySheetDimensions>,
        pub(crate) legacy_styles: Option<Stylesheet>,
    }

    pub(crate) fn matrix() -> Vec<OpenCase> {
        let mut cases = Vec::new();
        for sheet_count in [1, 3] {
            for dimensions in 0..3 {
                for legacy_style in 0..4 {
                    for indexed_colors in [false, true] {
                        for col_styles in [false, true] {
                            for features in [false, true] {
                                let mut model = model(sheet_count, features, col_styles);
                                if indexed_colors {
                                    model.styles.indexed_colors =
                                        vec!["#123456".into(), String::new(), "#abcdef".into()];
                                }
                                let legacy_dimensions = match dimensions {
                                    0 => Vec::new(),
                                    1 => model
                                        .sheets
                                        .iter()
                                        .map(|sheet| xlsx_parse::LegacySheetDimensions {
                                            col_widths: sheet.col_widths.clone(),
                                            row_heights: sheet.row_heights.clone(),
                                        })
                                        .collect(),
                                    _ => vec![xlsx_parse::LegacySheetDimensions {
                                        col_widths: BTreeMap::from([(1, 12.0)]),
                                        row_heights: BTreeMap::from([(2, 15.0)]),
                                    }],
                                };
                                let legacy_styles = match legacy_style {
                                    0 => None,
                                    1 => Some(model.styles.clone()),
                                    _ => {
                                        let mut styles = model.styles.clone();
                                        styles.indexed_colors = if indexed_colors {
                                            Vec::new()
                                        } else {
                                            vec!["#654321".into()]
                                        };
                                        if legacy_style == 2 {
                                            styles.fonts[0].bold = !styles.fonts[0].bold;
                                            styles.cell_xfs[1].num_fmt_id = None;
                                        }
                                        Some(styles)
                                    }
                                };
                                cases.push(OpenCase {
                                    label: format!(
                                        "sheets={sheet_count} dimensions={dimensions} legacy_style={legacy_style} indexed={indexed_colors} columns={col_styles} features={features}"
                                    ),
                                    model,
                                    legacy_dimensions,
                                    legacy_styles,
                                });
                            }
                        }
                    }
                }
            }
        }
        cases
    }

    fn model(sheet_count: usize, features: bool, col_styles: bool) -> WorkbookModel {
        let mut model = WorkbookModel {
            date_system: if features {
                DateSystem::V1904
            } else {
                DateSystem::V1900
            },
            shared_strings: vec!["header".into()],
            ..WorkbookModel::default()
        };
        model.styles.fonts.push(Font {
            name: Some("Calibri".into()),
            color: Some(Color::Indexed(0)),
            ..Font::default()
        });
        model.styles.fills = vec![Fill::None, Fill::Solid(Color::Rgb("#abcdef".into()))];
        model.styles.borders.push(Border::default());
        model.styles.num_fmts.push((164, "0.000".into()));
        model.styles.cell_xfs = vec![
            Xf::default(),
            Xf {
                font: Some(0),
                fill: Some(1),
                border: Some(0),
                num_fmt_id: Some(164),
                ..Xf::default()
            },
        ];
        for index in 0..sheet_count {
            let mut sheet = Sheet::new(format!("Sheet{}", index + 1));
            for (row, col, value, formula, style) in [
                (0, 0, CellValue::Number { value: 2.0 }, None, Some(0)),
                (1, 0, CellValue::Number { value: 3.0 }, Some("A1+1"), None),
                (0, 3, CellValue::Number { value: 3.0 }, Some("A1+1"), None),
                (1, 3, CellValue::Number { value: 4.0 }, Some("A2+1"), None),
                (
                    2,
                    1,
                    CellValue::Number { value: 1.0 },
                    Some("ROW(A1:A2)"),
                    None,
                ),
                (3, 1, CellValue::Number { value: 2.0 }, None, None),
                (4, 0, CellValue::Empty, None, Some(1)),
                (4, 1, CellValue::Empty, None, Some(0)),
                (
                    5,
                    0,
                    CellValue::Text {
                        value: "header".into(),
                    },
                    None,
                    None,
                ),
                (5, 1, CellValue::Bool { value: false }, None, None),
                (
                    5,
                    2,
                    CellValue::Error {
                        value: ErrorValue::Div0,
                    },
                    None,
                    None,
                ),
            ] {
                sheet.set_cell(
                    CellRef::new(row, col),
                    Cell {
                        value,
                        formula: formula.map(str::to_string),
                        style,
                    },
                );
            }
            sheet.set_array_formula(
                CellRef::new(2, 1),
                CellRange::new(CellRef::new(2, 1), CellRef::new(3, 1)),
            );
            sheet.col_widths = BTreeMap::from([(1, 0.0), (5, 18.25)]);
            sheet.row_heights = BTreeMap::from([(2, 0.0), (5, 20.5)]);
            sheet.format = SheetFormat {
                default_row_height_pt: Some(16.0),
                custom_height: true,
                zero_height: false,
            };
            if col_styles {
                sheet.col_styles.push(ColStyle {
                    first: 0,
                    last: 2,
                    xf: 1,
                });
            }
            if features {
                sheet.freeze_pane = Some(FreezePane::new(1, 1, CellRef::new(2, 2)));
                sheet
                    .merges
                    .push(CellRange::new(CellRef::new(7, 0), CellRef::new(7, 2)));
                sheet.hyperlinks.push(Hyperlink {
                    range: CellRange::new(CellRef::new(5, 0), CellRef::new(5, 0)),
                    external_target: Some("https://example.com".into()),
                    location: None,
                    tooltip: Some("Example".into()),
                    display: None,
                });
            }
            model.sheets.push(sheet);
        }
        if features {
            model.defined_names.push(DefinedName {
                name: "Answer".into(),
                formula: "Sheet1!$A$1".into(),
                local_sheet: None,
                hidden: false,
            });
            model.defined_names.push(DefinedName {
                name: "Local".into(),
                formula: "Sheet1!$B$3".into(),
                local_sheet: Some(SheetId(0)),
                hidden: true,
            });
            model.tables.push(Table {
                name: "Data".into(),
                sheet: SheetId(0),
                range: CellRange::new(CellRef::new(5, 0), CellRef::new(6, 1)),
                header_rows: 1,
                totals_rows: 0,
                columns: vec!["First".into(), "Second".into()],
            });
            model.sheets[0].charts.push(SheetChart {
                part: "xl/charts/chart1.xml".into(),
                drawing: "xl/drawings/drawing1.xml".into(),
                anchor_index: 0,
                anchor: ChartAnchor::OneCell {
                    from: AnchorCell::default(),
                    extent: AnchorExtent {
                        cx: 100_000,
                        cy: 200_000,
                    },
                },
                refs: vec![ChartRef {
                    kind: ChartRefKind::Values,
                    formula: "Sheet1!$A$1:$A$2".into(),
                }],
            });
        }
        model
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::Xf;

    fn source_conditional_formats() -> Vec<ConditionalFormatting> {
        serde_json::from_value(serde_json::json!([{
            "sqref": "A1:A5",
            "rules": [{"id": "cf1", "type": "expression", "formulas": ["A1>0"], "priority": 1,
                "rawXml": "<cfRule type=\"expression\" priority=\"1\"><formula>A1&gt;0</formula></cfRule>"}]
        }])).unwrap()
    }

    #[test]
    fn conditional_format_key_is_seeded_only_when_needed() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 101).unwrap();
        {
            let txn = authority.doc.transact();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            assert!(sheet.get(&txn, CONDITIONAL_FORMATS).is_none());
        }
        model.sheets[0].conditional_formats = source_conditional_formats();
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 102).unwrap();
        assert_eq!(authority.materialize().unwrap(), model);
        let mut txn = authority.doc.transact_mut();
        let sheet = txn
            .get_map(SHEETS)
            .unwrap()
            .get(&txn, "sheet:0")
            .unwrap()
            .cast::<MapRef>()
            .unwrap();
        assert!(sheet.get(&txn, CONDITIONAL_FORMATS).is_some());
        model.sheets[0].conditional_formats.clear();
        sync_sheet(&sheet, &mut txn, &model.sheets[0], &model.styles).unwrap();
        assert!(
            matches!(sheet.get(&txn, CONDITIONAL_FORMATS), Some(Out::Any(Any::String(value))) if value.as_ref() == "[]")
        );
    }

    #[test]
    fn yrs_without_conditional_format_key_retains_source_rules() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        let legacy = WorkbookAuthority::from_model_with_client_id(&model, 103).unwrap();
        {
            let mut txn = legacy.doc.transact_mut();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            sheet.remove(&mut txn, CONDITIONAL_FORMATS);
        }
        model.sheets[0].conditional_formats = source_conditional_formats();
        let current = WorkbookAuthority::from_model_with_client_id(&model, 104).unwrap();
        let SnapshotAdoption::Replacement(restored) =
            current.snapshot_replacement(&legacy.encode_state_as_update_v1())
        else {
            panic!("released Yrs snapshot should restore");
        };
        assert_eq!(restored.materialize().unwrap(), model);
        let recovered = WorkbookAuthority::restore_recovery_state(
            &restored.encode_recovery_state().unwrap(),
            105,
            0,
            true,
        )
        .unwrap();
        assert_eq!(recovered.materialize().unwrap(), model);
        {
            let mut txn = restored.doc.transact_mut();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            sheet.try_update(&mut txn, CONDITIONAL_FORMATS, "[]");
        }
        model.sheets[0].conditional_formats.clear();
        assert_eq!(restored.materialize().unwrap(), model);
    }

    #[test]
    fn remote_conditional_format_removal_is_refused_for_updates_and_snapshots() {
        use crate::{CalculationOptions, Workbook};

        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        model.sheets[0].conditional_formats = source_conditional_formats();
        let donor = Workbook::from_model_collaborative(model.clone(), 106).unwrap();
        let peer = Doc::with_client_id(107);
        hydrate_doc(&peer, &donor.encode_state_as_update_v1()).unwrap();
        let vector = peer.transact().state_vector();
        {
            let mut txn = peer.transact_mut();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            sheet.try_update(&mut txn, CONDITIONAL_FORMATS, "[]");
        }
        let diff = peer.transact().encode_diff_v1(&vector);
        let snapshot = peer
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        for touched in [false, true] {
            for update in [&diff, &snapshot] {
                let mut target = Workbook::from_model_collaborative(model.clone(), 108).unwrap();
                if touched {
                    target
                        .edit_cell(
                            SheetId(0),
                            CellRef::new(0, 0),
                            "1",
                            CalculationOptions::default(),
                        )
                        .unwrap();
                }
                let before = target.model().clone();
                let state = target.encode_state_as_update_v1();
                assert!(
                    target
                        .apply_update_v1(update, CalculationOptions::default())
                        .is_err()
                );
                assert_eq!(target.model(), &before);
                assert_eq!(target.encode_state_as_update_v1(), state);
            }
        }
    }

    #[test]
    fn multi_target_twins_emit_identical_updates() {
        use std::sync::Mutex;

        use crate::{CalculationOptions, Workbook};

        let mut source_model = rich_model();
        source_model.styles.fonts.push(xlsx_model::Font {
            bold: true,
            ..xlsx_model::Font::default()
        });
        source_model.styles.cell_xfs.push(Xf {
            font: Some(0),
            ..Xf::default()
        });
        let parts = xlsx_parse::serialize_workbook(&source_model).unwrap();
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let open = || {
            let mut workbook = Workbook::open_collaborative(&bytes, 71).unwrap();
            workbook.set_rand_seed(Some(23));
            workbook
        };
        let mut first = open();
        let mut second = open();
        let first_updates = Arc::new(Mutex::new(Vec::new()));
        let second_updates = Arc::new(Mutex::new(Vec::new()));
        let observed = Arc::clone(&first_updates);
        let _first_subscription = first
            .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
            .unwrap();
        let observed = Arc::clone(&second_updates);
        let _second_subscription = second
            .observe_update_v1(move |event| observed.lock().unwrap().push(event.update))
            .unwrap();
        let mut ops = Vec::new();
        for sheet in [SheetId(1), SheetId(0)] {
            for row in [7, 2, 5, 0] {
                for col in [4, 0, 2] {
                    ops.push(Op::SetCell {
                        sheet,
                        at: CellRef::new(row, col),
                        cell: CellState {
                            value: CellValue::Number {
                                value: f64::from(row * 10 + col),
                            },
                            formula: (row == 7 && col == 4).then(|| "RAND()+NOW()".into()),
                            style: Some(1),
                        },
                    });
                }
                ops.push(Op::SetRowHeight {
                    sheet,
                    row,
                    height: Some(f64::from(row + 20)),
                });
            }
            for col in [5, 0, 3] {
                ops.push(Op::SetColWidth {
                    sheet,
                    col,
                    width: Some(f64::from(col + 15)),
                });
            }
        }
        let options = CalculationOptions {
            now_serial: Some(45_000.25),
        };
        assert_eq!(
            first.apply_ops(ops.clone(), options).unwrap(),
            second.apply_ops(ops.clone(), options).unwrap(),
        );
        assert!(!first_updates.lock().unwrap().is_empty());
        assert_eq!(
            *first_updates.lock().unwrap(),
            *second_updates.lock().unwrap()
        );
        assert_eq!(
            first.encode_state_vector_v1(),
            second.encode_state_vector_v1()
        );
        assert_eq!(
            first.encode_state_as_update_v1(),
            second.encode_state_as_update_v1()
        );
        assert_eq!(first.save().unwrap(), second.save().unwrap());

        let mut first = WorkbookAuthority::from_model_with_client_id(&source_model, 71).unwrap();
        let mut second = WorkbookAuthority::from_model_with_client_id(&source_model, 71).unwrap();
        ops.extend([
            Op::MergeCells {
                sheet: SheetId(1),
                range: CellRange::new(CellRef::new(10, 0), CellRef::new(10, 2)),
            },
            Op::MergeCells {
                sheet: SheetId(0),
                range: CellRange::new(CellRef::new(12, 0), CellRef::new(12, 2)),
            },
        ]);
        assert_eq!(
            first
                .apply_ops(&ops, SyncOrigin::User, &source_model.styles)
                .unwrap(),
            second
                .apply_ops(&ops, SyncOrigin::User, &source_model.styles)
                .unwrap(),
        );
    }

    #[test]
    fn fingerprint_data_matches_clone_oracle_matrix() {
        for case in open_test_cases::matrix() {
            let actual = WorkbookBase::from_model_with_legacy_dimensions(
                &case.model,
                &case.legacy_dimensions,
                case.legacy_styles.as_ref(),
            )
            .unwrap();
            let expected = WorkbookBase::from_model_with_legacy_dimensions_oracle(
                &case.model,
                &case.legacy_dimensions,
                case.legacy_styles.as_ref(),
            )
            .unwrap();
            assert_eq!(actual.fingerprints, expected.fingerprints, "{}", case.label);
            assert_eq!(actual.fingerprint, expected.fingerprint, "{}", case.label);
            assert_eq!(
                actual.bootstrap_client_id, expected.bootstrap_client_id,
                "{}",
                case.label
            );
        }
    }

    #[test]
    fn pattern_details_keep_the_released_base_fingerprint() {
        use xlsx_model::{Border, BorderEdge, BorderStyle, Color, Fill, PatternType};
        let mut patterned = open_test_cases::matrix().remove(0).model;
        let mut released = patterned.clone();
        for (model, fill, style) in [
            (
                &mut patterned,
                Fill::Pattern {
                    pattern: PatternType::LightUp,
                    fg: Some(Color::Rgb("#123456".into())),
                    bg: Some(Color::Rgb("#fedcba".into())),
                },
                BorderStyle::MediumDashDot,
            ),
            (
                &mut released,
                Fill::Solid(Color::Rgb("#123456".into())),
                BorderStyle::Medium,
            ),
        ] {
            model.styles.fills.push(fill);
            model.styles.borders.push(Border {
                top: Some(BorderEdge { style, color: None }),
                ..Border::default()
            });
        }
        let actual = WorkbookBase::from_model(&patterned).unwrap();
        let expected = WorkbookBase::from_model(&released).unwrap();
        assert_ne!(actual.fingerprint, expected.fingerprint);
        assert_ne!(actual.bootstrap_client_id, expected.bootstrap_client_id);
        let current = &actual.fingerprints[&SCHEMA_VERSION];
        assert!(current.contains(&actual.fingerprint));
        assert_eq!(
            released_readings(&actual.fingerprints),
            expected.fingerprints
        );
        let oracle =
            WorkbookBase::from_model_with_legacy_dimensions_oracle(&patterned, &[], None).unwrap();
        assert_eq!(oracle.fingerprints, actual.fingerprints);
        assert_eq!(oracle.fingerprint, actual.fingerprint);
    }

    /// The accepted fingerprints hashed over the released reading of the styles.
    fn released_readings(fingerprints: &BTreeMap<i64, Vec<String>>) -> BTreeMap<i64, Vec<String>> {
        let mut released = fingerprints.clone();
        for accepted in released.values_mut() {
            accepted.retain(|fingerprint| !fingerprint_reading(fingerprint).as_written());
        }
        released
    }

    #[test]
    fn pattern_backgrounds_tell_current_bases_apart() {
        use xlsx_model::{Color, Fill, PatternType};
        let bases = ["#fedcba", "#00ff00"].map(|bg| {
            let mut model = open_test_cases::matrix().remove(0).model;
            model.styles.fills.push(Fill::Pattern {
                pattern: PatternType::LightUp,
                fg: Some(Color::Rgb("#123456".into())),
                bg: Some(Color::Rgb(bg.into())),
            });
            let base = WorkbookBase::from_model(&model).unwrap();
            let oracle =
                WorkbookBase::from_model_with_legacy_dimensions_oracle(&model, &[], None).unwrap();
            assert_eq!(oracle.fingerprint, base.fingerprint);
            assert_eq!(oracle.fingerprints, base.fingerprints);
            assert_eq!(oracle.bootstrap_client_id, base.bootstrap_client_id);
            base
        });
        assert_ne!(bases[0].fingerprint, bases[1].fingerprint);
        assert_ne!(bases[0].bootstrap_client_id, bases[1].bootstrap_client_id);
    }

    #[test]
    fn released_reader_hashes_follow_the_released_collapse_rule() {
        use xlsx_model::{Border, BorderEdge, BorderStyle, Color, Fill, PatternType};
        // released parsers read a pattern as a solid of its foreground (auto without one)
        // and the six extended border styles as medium or dashed
        let collapse = |style: BorderStyle| match style.as_sml() {
            "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" => BorderStyle::Medium,
            "dashDot" | "dashDotDot" | "slantDashDot" => BorderStyle::Dashed,
            _ => style,
        };
        let border = |style, color: &str| Border {
            top: Some(BorderEdge {
                style,
                color: Some(Color::Rgb(color.into())),
            }),
            ..Border::default()
        };
        let mut patterned = open_test_cases::matrix().remove(0).model;
        let mut released = patterned.clone();
        let first_pattern = patterned.styles.fills.len();
        for (index, pattern) in PatternType::ALL.into_iter().enumerate() {
            let fg = (index % 2 == 0).then(|| Color::Rgb(format!("#1234{index:02}")));
            patterned.styles.fills.push(Fill::Pattern {
                pattern,
                fg: fg.clone(),
                bg: Some(Color::Indexed(9)),
            });
            released
                .styles
                .fills
                .push(Fill::Solid(fg.unwrap_or(Color::Auto)));
        }
        for token in [
            "mediumDashed",
            "dashDot",
            "mediumDashDot",
            "dashDotDot",
            "mediumDashDotDot",
            "slantDashDot",
        ] {
            let style = BorderStyle::from_sml(token);
            patterned.styles.borders.push(border(style, "#654321"));
            released
                .styles
                .borders
                .push(border(collapse(style), "#654321"));
        }
        let hashes = |model: &WorkbookModel| {
            let base = WorkbookBase::from_model(model).unwrap();
            let oracle =
                WorkbookBase::from_model_with_legacy_dimensions_oracle(model, &[], None).unwrap();
            (
                released_readings(&base.fingerprints),
                released_readings(&oracle.fingerprints),
            )
        };
        let expected = hashes(&released);
        assert_eq!(hashes(&patterned), expected);

        let mut foreground = patterned.clone();
        foreground.styles.fills[first_pattern] = Fill::Pattern {
            pattern: PatternType::ALL[0],
            fg: Some(Color::Rgb("#000001".into())),
            bg: Some(Color::Indexed(9)),
        };
        let mut edge_color = patterned.clone();
        *edge_color.styles.borders.last_mut().unwrap() =
            border(BorderStyle::SlantDashDot, "#000002");
        for changed in [foreground, edge_color] {
            let (fingerprints, oracle) = hashes(&changed);
            assert_ne!(fingerprints, expected.0);
            assert_ne!(oracle, expected.1);
        }
    }

    #[test]
    fn fingerprint_style_overrides_preserve_signed_zero() {
        let mut case = open_test_cases::matrix().remove(0);
        case.model.styles.fonts[0].size_pt = Some(-0.0);
        case.model.styles.indexed_colors = vec!["#123456".into()];
        let mut legacy_styles = case.model.styles.clone();
        legacy_styles.fonts[0].size_pt = Some(0.0);
        legacy_styles.indexed_colors = vec!["#abcdef".into()];
        let actual =
            WorkbookBase::from_model_with_legacy_dimensions(&case.model, &[], Some(&legacy_styles))
                .unwrap();
        let expected = WorkbookBase::from_model_with_legacy_dimensions_oracle(
            &case.model,
            &[],
            Some(&legacy_styles),
        )
        .unwrap();
        assert_eq!(actual.fingerprints, expected.fingerprints);
        assert_eq!(actual.fingerprint, expected.fingerprint);
        assert_eq!(actual.bootstrap_client_id, expected.bootstrap_client_id);
    }

    #[test]
    fn fingerprints_accept_the_reading_before_format_facets() {
        let mut model = open_test_cases::matrix().remove(0).model;
        let before_facets = model.clone();
        let xf = &mut model.styles.cell_xfs[1];
        xf.xf_id = Some(0);
        xf.quote_prefix = true;
        xf.protection = Some(xlsx_model::Protection {
            locked: false,
            hidden: false,
        });
        xf.alignment = Some(xlsx_model::Alignment {
            text_rotation: Some(90),
            ..Default::default()
        });
        let mut legacy_styles = model.styles.clone();
        legacy_styles.cell_xfs[1].num_fmt_id = None;
        let mut held_back = before_facets.clone();
        held_back.styles.cell_xfs[1].num_fmt_id = None;
        let actual =
            WorkbookBase::from_model_with_legacy_dimensions(&model, &[], Some(&legacy_styles))
                .unwrap();
        let expected = WorkbookBase::from_model_with_legacy_dimensions_oracle(
            &model,
            &[],
            Some(&legacy_styles),
        )
        .unwrap();
        assert_eq!(actual.fingerprints, expected.fingerprints);
        assert_eq!(actual.fingerprint, expected.fingerprint);
        let (current, _) = fingerprint_model_current_oracle(&model).unwrap();
        assert_eq!(actual.fingerprint, current);
        for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
            for reading in [&before_facets, &held_back] {
                let (fingerprint, _) = fingerprint_model_for_schema(reading, version).unwrap();
                assert!(
                    actual.accepts_fingerprint(version, &fingerprint),
                    "{version}"
                );
            }
        }
    }

    #[test]
    fn single_projection_matches_double_projection_matrix() {
        for case in open_test_cases::matrix() {
            let (actual, model, structure) = WorkbookAuthority::from_source_with_projection(
                &case.model,
                Some(11),
                &case.legacy_dimensions,
                case.legacy_styles.as_ref(),
                &[],
            )
            .unwrap();
            let expected = WorkbookAuthority::from_source_oracle(
                &case.model,
                Some(11),
                &case.legacy_dimensions,
                case.legacy_styles.as_ref(),
                &[],
            )
            .unwrap();
            let projection = expected.materialize_internal_oracle(false).unwrap();
            assert_eq!((model, structure), projection, "{}", case.label);
            assert_eq!(
                actual.encode_state_vector_v1(),
                expected.encode_state_vector_v1(),
                "{}",
                case.label
            );
            assert_eq!(
                actual.encode_state_as_update_v1(),
                expected.encode_state_as_update_v1(),
                "{}",
                case.label
            );
        }
    }

    fn rich_model() -> WorkbookModel {
        let mut first = Sheet::new("Data");
        first.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number { value: 42.0 },
                formula: Some("40+2".into()),
                style: Some(0),
            },
        );
        first.set_cell(
            CellRef::new(1, 0),
            Cell {
                value: CellValue::Text {
                    value: "hello".into(),
                },
                ..Cell::default()
            },
        );
        first.col_widths.insert(1, 24.5);
        first.row_heights.insert(2, 30.0);
        first
            .merges
            .push(CellRange::new(CellRef::new(3, 0), CellRef::new(4, 2)));
        first.freeze_pane = Some(FreezePane::new(1, 1, CellRef::new(3, 2)));
        first.hyperlinks.push(Hyperlink {
            range: CellRange::new(CellRef::new(1, 0), CellRef::new(1, 0)),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: Some("Open".into()),
            display: Some("Example".into()),
        });
        let mut model = WorkbookModel {
            date_system: DateSystem::V1904,
            shared_strings: vec!["hello".into()],
            ..WorkbookModel::default()
        };
        model.defined_names.push(DefinedName {
            name: "Answer".into(),
            formula: "Data!A1".into(),
            local_sheet: None,
            hidden: false,
        });
        model.styles.cell_xfs.push(Xf::default());
        model.sheets.push(first);
        model.sheets.push(Sheet::new("Second"));
        model
    }

    fn assert_singleton_lockstep(
        fast: &mut WorkbookAuthority,
        oracle: &mut WorkbookAuthority,
        ops: &[Op],
        origin: SyncOrigin,
        styles: &Stylesheet,
    ) -> Result<Option<Vec<u8>>, String> {
        let result = fast
            .apply_ops(ops, origin, styles)
            .map_err(|error| format!("{error:?}"));
        let expected = with_full_materialization(|| oracle.apply_ops(ops, origin, styles))
            .map_err(|error| format!("{error:?}"));
        assert_eq!(result, expected);
        assert_eq!(
            fast.encode_state_vector_v1(),
            oracle.encode_state_vector_v1()
        );
        assert_eq!(
            fast.encode_state_as_update_v1(),
            oracle.encode_state_as_update_v1()
        );
        assert_eq!(fast.history.undo, oracle.history.undo);
        assert_eq!(fast.history.redo, oracle.history.redo);
        let model = fast.materialize().map_err(|error| format!("{error:?}"));
        let expected_model = oracle.materialize().map_err(|error| format!("{error:?}"));
        assert_eq!(model, expected_model);
        result
    }

    #[test]
    fn singleton_set_cell_sync_matches_authored_state_and_catalog() {
        let model = rich_model();
        let mut fast = WorkbookAuthority::from_model_with_client_id(&model, 73).unwrap();
        let mut oracle = WorkbookAuthority::from_model_with_client_id(&model, 73).unwrap();
        let mut styles = model.styles.clone();
        let mut format = CellFormat::default();
        format.font.bold = true;
        let bold = styles.intern_cell_format(&format).unwrap();
        let count = fast_set_cell_count();
        let cached = Op::SetCell {
            sheet: SheetId(0),
            at: CellRef::new(0, 0),
            cell: CellState {
                value: CellValue::Number { value: 999.0 },
                formula: Some("40+2".into()),
                style: Some(0),
            },
        };
        assert_eq!(
            assert_singleton_lockstep(&mut fast, &mut oracle, &[cached], SyncOrigin::User, &styles)
                .unwrap(),
            None
        );
        assert_eq!(
            fast.materialize().unwrap().sheets[0]
                .cell(CellRef::new(0, 0))
                .unwrap()
                .value,
            CellValue::Number { value: 42.0 }
        );
        for (origin, cell) in [
            (
                SyncOrigin::User,
                CellState {
                    value: CellValue::Number { value: 123.0 },
                    style: bold,
                    ..Default::default()
                },
            ),
            (
                SyncOrigin::Agent,
                CellState {
                    formula: Some("Data!A1+1".into()),
                    style: bold,
                    ..Default::default()
                },
            ),
            (
                SyncOrigin::User,
                CellState {
                    style: bold,
                    ..Default::default()
                },
            ),
            (SyncOrigin::Undo, CellState::default()),
            (
                SyncOrigin::Redo,
                CellState {
                    value: CellValue::Bool { value: true },
                    ..Default::default()
                },
            ),
        ] {
            let op = Op::SetCell {
                sheet: SheetId(1),
                at: CellRef::new(2, 3),
                cell,
            };
            assert_singleton_lockstep(&mut fast, &mut oracle, &[op], origin, &styles).unwrap();
        }
        assert_eq!(fast_set_cell_count(), count + 6);
    }

    #[test]
    fn singleton_set_cell_falls_back_before_writes() {
        let model = rich_model();
        for op in [
            Op::SetCell {
                sheet: SheetId(99),
                at: CellRef::new(0, 0),
                cell: CellState::default(),
            },
            Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    style: Some(u32::MAX),
                    ..Default::default()
                },
            },
        ] {
            let mut fast = WorkbookAuthority::from_model_with_client_id(&model, 73).unwrap();
            let mut oracle = WorkbookAuthority::from_model_with_client_id(&model, 73).unwrap();
            let before = fast.encode_state_as_update_v1();
            let count = fast_set_cell_count();
            assert!(
                assert_singleton_lockstep(
                    &mut fast,
                    &mut oracle,
                    &[op],
                    SyncOrigin::User,
                    &model.styles
                )
                .is_err()
            );
            assert_eq!(fast.encode_state_as_update_v1(), before);
            assert_eq!(fast_set_cell_count(), count);
        }
        let mut fast = WorkbookAuthority::from_model_with_client_id(&model, 73).unwrap();
        let mut oracle = WorkbookAuthority::from_model_with_client_id(&model, 73).unwrap();
        for authority in [&fast, &oracle] {
            let mut txn = authority.doc.transact_mut();
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            let contents = nested_map(&sheet, &txn, CONTENTS).unwrap();
            contents.insert(&mut txn, "1:0", Any::BigInt(42));
        }
        let before = fast.encode_state_as_update_v1();
        let count = fast_set_cell_count();
        let op = Op::SetCell {
            sheet: SheetId(1),
            at: CellRef::new(0, 0),
            cell: CellState::default(),
        };
        assert!(
            assert_singleton_lockstep(
                &mut fast,
                &mut oracle,
                &[op],
                SyncOrigin::User,
                &model.styles
            )
            .is_err()
        );
        assert_eq!(fast.encode_state_as_update_v1(), before);
        assert_eq!(fast_set_cell_count(), count);
    }

    #[test]
    fn singleton_set_cell_legacy_schema_matches_full_materialization() {
        let model = rich_model();
        for version in MIN_SUPPORTED_SCHEMA_VERSION..SCHEMA_VERSION {
            let bytes = legacy_update(&model, version, true);
            let mut fast = authority_from_update(&model, &bytes, 73);
            let mut oracle = authority_from_update(&model, &bytes, 73);
            fast.materialize().unwrap();
            oracle.materialize().unwrap();
            let count = fast_set_cell_count();
            let op = Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(1, 0),
                cell: CellState {
                    value: CellValue::Text {
                        value: "new".into(),
                    },
                    ..Default::default()
                },
            };
            assert_singleton_lockstep(
                &mut fast,
                &mut oracle,
                &[op],
                SyncOrigin::User,
                &model.styles,
            )
            .unwrap();
            assert_eq!(fast_set_cell_count(), count);
        }
    }

    pub(super) fn legacy_update(
        model: &WorkbookModel,
        version: i64,
        include_defined_names: bool,
    ) -> Vec<u8> {
        let base = WorkbookBase::from_model(model).unwrap();
        let (_, client_id) =
            fingerprint_model_with_schema(model, version, include_defined_names, true).unwrap();
        let doc = Doc::with_client_id(client_id);
        let keys = (0..model.sheets.len())
            .map(|index| format!("sheet:{index}"))
            .collect::<Vec<_>>();
        seed(&doc, &base, model, &keys).unwrap();
        {
            let (fingerprint, _) =
                fingerprint_model_with_schema(model, version, include_defined_names, true).unwrap();
            let mut txn = doc.transact_mut_with("test:legacy-schema");
            let meta = txn.get_map(META).unwrap();
            meta.try_update(&mut txn, BASE_FINGERPRINT, fingerprint);
            meta.try_update(&mut txn, "schemaVersion", version);
            let sheets = txn.get_map(SHEETS).unwrap();
            for key in keys {
                let sheet = sheets
                    .get(&txn, &key)
                    .and_then(|value| value.cast::<MapRef>().ok())
                    .unwrap();
                if version < CHARTS_SCHEMA_VERSION {
                    sheet.remove(&mut txn, CHARTS);
                }
                if version < HYPERLINK_SCHEMA_VERSION {
                    sheet.remove(&mut txn, HYPERLINKS);
                }
                if version < FREEZE_PANE_SCHEMA_VERSION {
                    sheet.remove(&mut txn, FREEZE_PANE);
                }
            }
        }
        doc.transact()
            .encode_state_as_update_v1(&StateVector::default())
    }

    pub(super) fn authority_from_update(
        model: &WorkbookModel,
        update: &[u8],
        client_id: u64,
    ) -> WorkbookAuthority {
        let doc = Doc::with_client_id(client_id);
        hydrate_doc(&doc, update).unwrap();
        WorkbookAuthority::hydrated(doc, Arc::new(WorkbookBase::from_model(model).unwrap()), 0)
    }

    #[test]
    fn a_column_styled_workbook_still_accepts_its_pre_change_fingerprint() {
        let mut model = rich_model();
        model.sheets[0].col_styles = vec![ColStyle {
            first: 0,
            last: 3,
            xf: 1,
        }];
        let base = WorkbookBase::from_model(&model).unwrap();
        for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
            let (pre_change, _) =
                fingerprint_model_with_schema(&model, version, version >= 4, false).unwrap();
            assert!(
                base.fingerprints[&version].contains(&pre_change),
                "schema v{version} must still recognise the fingerprint released before \
                 column styles were read"
            );
        }
        let (current, _) = fingerprint_model_for_schema(&model, SCHEMA_VERSION).unwrap();
        let (without, _) =
            fingerprint_model_with_schema(&model, SCHEMA_VERSION, true, false).unwrap();
        assert_ne!(
            current, without,
            "two workbooks differing only in their column styles must not share a fingerprint"
        );

        let mut plain = rich_model();
        plain.sheets[0].col_styles.clear();
        let (plain_current, _) = fingerprint_model_for_schema(&plain, SCHEMA_VERSION).unwrap();
        let (plain_without, _) =
            fingerprint_model_with_schema(&plain, SCHEMA_VERSION, true, false).unwrap();
        assert_eq!(
            plain_current, plain_without,
            "a workbook declaring no column style keeps the fingerprint it always had"
        );
    }

    #[test]
    fn a_row_and_column_styled_workbook_accepts_its_column_only_fingerprint() {
        let mut model = rich_model();
        model.sheets[0].col_styles = vec![ColStyle {
            first: 0,
            last: 3,
            xf: 1,
        }];
        model.sheets[0].row_styles = BTreeMap::from([(4, 1)]);
        let base = WorkbookBase::from_model(&model).unwrap();
        let mut columns_only = model.clone();
        columns_only.sheets[0].row_styles.clear();
        for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
            let (released, _) = fingerprint_model_with_row_styles_oracle(
                &columns_only,
                version,
                version >= 4,
                true,
                false,
            )
            .unwrap();
            assert!(
                base.fingerprints[&version].contains(&released),
                "schema v{version} must recognise a fingerprint hashing column styles only"
            );
        }
        let (current, _) = fingerprint_model_current_oracle(&model).unwrap();
        let (released, _) = fingerprint_model_with_row_styles_oracle(
            &columns_only,
            SCHEMA_VERSION,
            true,
            true,
            false,
        )
        .unwrap();
        assert_ne!(current, released);
        assert_eq!(base.fingerprint, current);
        let oracle =
            WorkbookBase::from_model_with_legacy_dimensions_oracle(&model, &[], None).unwrap();
        assert_eq!(base.fingerprints, oracle.fingerprints);
    }

    #[test]
    fn style_default_fingerprints_take_their_own_domain_and_accept_released_readings() {
        let mut columns = rich_model();
        columns.sheets[0].col_styles = vec![ColStyle {
            first: 0,
            last: 3,
            xf: 1,
        }];
        let mut rows = rich_model();
        rows.sheets[0].row_styles = BTreeMap::from([(4, 1)]);
        for model in [&columns, &rows] {
            let base = WorkbookBase::from_model(model).unwrap();
            let (released, released_client) =
                fingerprint_model_with_row_styles_oracle(model, SCHEMA_VERSION, true, true, false)
                    .unwrap();
            assert!(base.fingerprints[&SCHEMA_VERSION].contains(&released));
            assert_ne!(
                base.fingerprint, released,
                "a state that may carry row and column defaults must not pass for a released one"
            );
            assert_ne!(base.bootstrap_client_id, released_client);
        }
        let plain = rich_model();
        let base = WorkbookBase::from_model(&plain).unwrap();
        let (released, released_client) =
            fingerprint_model_with_row_styles_oracle(&plain, SCHEMA_VERSION, true, true, false)
                .unwrap();
        let (current, current_client) = fingerprint_model_current_oracle(&plain).unwrap();
        assert_eq!(base.fingerprint, current);
        assert_eq!(base.bootstrap_client_id, current_client);
        assert!(base.fingerprints[&SCHEMA_VERSION].contains(&released));
        assert_ne!(base.fingerprint, released);
        assert_ne!(base.bootstrap_client_id, released_client);
    }

    #[test]
    fn remapping_a_plain_style_over_an_italic_xf_zero_keeps_it_plain() {
        let mut target = Stylesheet::default();
        target.fonts.push(xlsx_model::Font {
            italic: true,
            ..Default::default()
        });
        target.cell_xfs.push(Xf {
            font: Some(0),
            ..Xf::default()
        });
        let mut source = target.clone();
        source.cell_xfs.push(Xf::default());
        let ops = [Op::SetCell {
            sheet: SheetId(0),
            at: CellRef::new(0, 0),
            cell: CellState {
                style: Some(1),
                ..CellState::default()
            },
        }];
        let remapped = remap_styles(&ops, &source, &mut target).unwrap();
        let [Op::SetCell { cell, .. }] = &remapped[..] else {
            panic!("expected one remapped cell");
        };
        assert_eq!(target.cell_format(cell.style), CellFormat::default());
        assert!(target.cell_format(Some(0)).font.italic);
    }

    #[test]
    fn remapping_a_plain_style_follows_the_shared_plain_key() {
        let mut target = Stylesheet::default();
        target.fonts.extend([
            xlsx_model::Font {
                italic: true,
                ..Default::default()
            },
            xlsx_model::Font::default(),
        ]);
        target.cell_xfs.extend([
            Xf {
                font: Some(0),
                ..Xf::default()
            },
            Xf {
                font: Some(1),
                ..Xf::default()
            },
        ]);
        let mut source = target.clone();
        source.cell_xfs.push(Xf::default());
        let ops = [Op::SetCell {
            sheet: SheetId(0),
            at: CellRef::new(0, 0),
            cell: CellState {
                style: Some(2),
                ..CellState::default()
            },
        }];
        let remapped = remap_styles(&ops, &source, &mut target).unwrap();
        let [Op::SetCell { cell, .. }] = &remapped[..] else {
            panic!("expected one remapped cell");
        };
        assert_eq!(target.cell_format(cell.style), CellFormat::default());
        assert_eq!(cell.style, Some(1), "the first xf reading plain");
    }

    #[test]
    fn remapping_a_fresh_plain_style_onto_an_unproven_bound_xf_uses_zero() {
        let mut target = Stylesheet::default();
        target.fonts.extend([
            xlsx_model::Font {
                italic: true,
                ..Default::default()
            },
            xlsx_model::Font::default(),
        ]);
        target.cell_xfs.extend([
            Xf {
                font: Some(0),
                ..Xf::default()
            },
            Xf {
                font: Some(1),
                ..Xf::default()
            },
        ]);
        let mut source = target.clone();
        source.cell_xfs.push(Xf::default());
        target.bind_plain_xf();
        let set = |at, style| Op::SetCell {
            sheet: SheetId(0),
            at,
            cell: CellState {
                style,
                ..CellState::default()
            },
        };
        let ops = [
            set(CellRef::new(0, 0), Some(2)),
            set(CellRef::new(0, 1), Some(1)),
        ];
        let remapped = remap_styles(&ops, &source, &mut target).unwrap();
        let [Op::SetCell { cell: a1, .. }, Op::SetCell { cell: b1, .. }] = &remapped[..] else {
            panic!("expected two remapped cells");
        };
        assert_eq!(a1.style, Some(0));
        assert_eq!(b1.style, Some(1));
    }

    #[test]
    fn deterministic_bootstrap_round_trips_formula_fallbacks() {
        let model = rich_model();
        let left = WorkbookAuthority::from_model_with_client_id(&model, 11).unwrap();
        let right = WorkbookAuthority::from_model_with_client_id(&model, 12).unwrap();
        assert_eq!(left.materialize().unwrap(), model);
        assert_eq!(right.materialize().unwrap(), model);
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        assert_eq!(
            left.encode_state_as_update_v1(),
            right.encode_state_as_update_v1()
        );
    }

    #[test]
    fn canonical_cell_formats_preserve_high_precision_theme_tints() {
        for tint in [
            "-0.14996795556505021",
            "-0.24994659260841701",
            "-4.9989318521683403E-2",
            "0.39994506668294322",
        ] {
            let format = CellFormat {
                fill: xlsx_model::Fill::Solid(xlsx_model::Color::Theme {
                    idx: 4,
                    tint: tint.parse().unwrap(),
                }),
                ..CellFormat::default()
            };
            let (key, payload) = cell_format_entry(&format).unwrap();
            let decoded: CellFormat = serde_json::from_str(&payload).unwrap();
            assert_eq!(decoded, format, "{tint}");
            assert_eq!(cell_format_entry(&decoded).unwrap(), (key, payload));

            let mut model = rich_model();
            let style = model.styles.intern_cell_format(&format).unwrap();
            model.sheets[0].set_cell(
                CellRef::new(1, 1),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    style,
                    ..Cell::default()
                },
            );
            let authority = WorkbookAuthority::from_model_with_client_id(&model, 11).unwrap();
            assert_eq!(authority.materialize().unwrap(), model);
            let peer = authority_from_update(&model, &authority.encode_state_as_update_v1(), 12);
            assert_eq!(peer.materialize().unwrap(), model);
        }
    }

    #[test]
    fn large_local_workbooks_stage_edits_and_restore_checkpoints() {
        let mut sheet = Sheet::new("Data");
        for row in 0..=(MAX_UPDATE_VALUES / 5) as u32 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number { value: row as f64 },
                    ..Cell::default()
                },
            );
        }
        let model = WorkbookModel {
            sheets: vec![sheet],
            ..WorkbookModel::default()
        };
        let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 11).unwrap();
        let peer = WorkbookAuthority::from_model_with_client_id(&model, 12).unwrap();
        let snapshot = authority.encode_state_as_update_v1();
        assert!(decode_update_v1(&snapshot).is_err());
        assert!(matches!(
            peer.snapshot_replacement(&snapshot),
            SnapshotAdoption::NotApplicable
        ));
        assert!(authority.materialize().unwrap() == model);
        assert!(matches!(
            peer.stage_updates_v1(&[&snapshot], None),
            Err(AuthorityError::InvalidUpdate(_))
        ));
        let checkpoint = authority.checkpoint();

        let at = CellRef::new(0, 0);
        let staged = authority
            .stage_local_ops_v1(
                &[Op::SetCell {
                    sheet: SheetId(0),
                    at,
                    cell: xlsx_ops::CellState {
                        value: CellValue::Number { value: -1.0 },
                        ..xlsx_ops::CellState::default()
                    },
                }],
                SyncOrigin::User,
                &Stylesheet::default(),
            )
            .unwrap();
        authority
            .apply_local_update_v1(&staged.update, SyncOrigin::User)
            .unwrap();
        let remote = peer.stage_updates_v1(&[&staged.update], None).unwrap();
        peer.apply_staged_update_v1(&remote.commit_update).unwrap();
        assert!(peer.materialize().unwrap() == authority.materialize().unwrap());

        assert!(authority.undo().unwrap().unwrap().model == model);
        assert!(authority.redo().unwrap().unwrap().model == peer.materialize().unwrap());
        authority.restore(checkpoint).unwrap();
        assert!(!authority.can_undo());
        assert!(!authority.can_redo());
        assert!(authority.materialize().unwrap() == model);
    }

    #[test]
    fn indexed_palettes_survive_legacy_snapshots_and_distinguish_current_bases() {
        let mut model = rich_model();
        let legacy = WorkbookAuthority::from_model_with_client_id(&model, 11).unwrap();
        model.styles.indexed_colors = vec!["#123456".into(); 64];
        let current = WorkbookAuthority::from_model_with_client_id(&model, 12).unwrap();
        let SnapshotAdoption::Replacement(mut restored) =
            current.snapshot_replacement(&legacy.encode_state_as_update_v1())
        else {
            panic!("legacy snapshot should retain the source palette");
        };
        assert_eq!(restored.materialize().unwrap(), model);
        let mut other_model = model.clone();
        other_model.styles.indexed_colors[2] = "#abcdef".into();
        let other = WorkbookAuthority::from_model_with_client_id(&other_model, 13).unwrap();
        assert!(matches!(
            current.snapshot_replacement(&other.encode_state_as_update_v1()),
            SnapshotAdoption::Incompatible(_)
        ));
        restored
            .apply_ops(
                &[Op::SetCell {
                    sheet: SheetId(0),
                    at: CellRef::new(0, 0),
                    cell: xlsx_ops::CellState {
                        value: CellValue::Number { value: 99.0 },
                        ..Default::default()
                    },
                }],
                SyncOrigin::User,
                &Stylesheet::default(),
            )
            .unwrap();
        let restored_snapshot = restored.encode_state_as_update_v1();
        assert!(matches!(
            other.snapshot_replacement(&restored_snapshot),
            SnapshotAdoption::Incompatible(_)
        ));
        let SnapshotAdoption::Replacement(same_palette) =
            current.snapshot_replacement(&restored_snapshot)
        else {
            panic!("restored snapshots should match the source palette");
        };
        assert_eq!(
            same_palette.materialize().unwrap(),
            restored.materialize().unwrap()
        );
        assert!(!same_palette.upgrade_schema().unwrap());
    }

    #[test]
    fn a_base_differing_only_in_format_facets_refuses_current_snapshots() {
        let plain = rich_model();
        let current = WorkbookAuthority::from_model_with_client_id(&plain, 11)
            .unwrap()
            .encode_state_as_update_v1();
        for facet in 0..2 {
            let mut faceted = plain.clone();
            if facet == 0 {
                faceted.styles.cell_xfs[0].alignment = Some(xlsx_model::Alignment {
                    text_rotation: Some(90),
                    ..Default::default()
                });
            } else {
                faceted.styles.cell_xfs[0].protection = Some(xlsx_model::Protection {
                    locked: false,
                    hidden: false,
                });
            }
            let base = WorkbookAuthority::from_model_with_client_id(&faceted, 12).unwrap();
            assert!(
                matches!(
                    base.snapshot_replacement(&current),
                    SnapshotAdoption::Incompatible(_)
                ),
                "facet {facet}"
            );
            let SnapshotAdoption::Replacement(restored) =
                base.snapshot_replacement(&pre_facet_state_for_test(&plain))
            else {
                panic!("a released snapshot of the faceted base must still restore");
            };
            assert_eq!(restored.strict_materialize().unwrap().0, faceted);
        }
    }

    #[test]
    fn known_schema_versions_materialize_and_upgrade_to_current() {
        let model = rich_model();
        for (index, (version, include_defined_names)) in
            [(3, false), (3, true), (4, true), (5, true)]
                .into_iter()
                .enumerate()
        {
            let update = legacy_update(&model, version, include_defined_names);
            let authority = authority_from_update(&model, &update, 101 + index as u64);
            assert_eq!(authority.strict_materialize().unwrap().0, model);

            let staged = authority
                .stage_updates_v1(&[Update::EMPTY_V1], None)
                .unwrap();
            assert!(staged.effective);
            authority
                .apply_staged_update_v1(&staged.commit_update)
                .unwrap();
            assert_eq!(authority.schema_version().unwrap(), SCHEMA_VERSION);
            assert_eq!(authority.strict_materialize().unwrap().0, model);
        }
    }

    #[test]
    fn legacy_snapshot_merges_into_current_bootstrap() {
        let model = rich_model();
        for (version, include_defined_names) in [(3, false), (3, true), (4, true), (5, true)] {
            let update = legacy_update(&model, version, include_defined_names);
            let authority = WorkbookAuthority::from_model_with_client_id(&model, 108).unwrap();
            let staged = authority.stage_updates_v1(&[&update], None).unwrap();
            assert_eq!(staged.model, model);
            authority
                .apply_staged_update_v1(&staged.commit_update)
                .unwrap();
            assert_eq!(authority.schema_version().unwrap(), SCHEMA_VERSION);
            assert_eq!(authority.strict_materialize().unwrap().0, model);
        }
    }

    /// The legacy fallback is keyed on `sheet:N`, not on where a sheet sits
    /// now, so a reordered legacy state still reads its own charts.
    #[test]
    fn a_reordered_legacy_state_keeps_each_sheet_its_own_charts() {
        let mut model = WorkbookModel::default();
        model.sheets.push(charted("First", "First!$A$1"));
        model.sheets.push(charted("Second", "Second!$A$1"));
        let base = WorkbookBase::from_model(&model).unwrap();
        let doc = Doc::with_client_id(base.bootstrap_client_id);
        seed(
            &doc,
            &base,
            &model,
            &["sheet:0".to_owned(), "sheet:1".to_owned()],
        )
        .unwrap();
        {
            let mut txn = doc.transact_mut_with("test:reordered-legacy");
            let order = txn.get_array(SHEET_ORDER).unwrap();
            order.remove(&mut txn, 0);
            order.insert(&mut txn, 1, "sheet:0");
            let sheets = txn.get_map(SHEETS).unwrap();
            for key in ["sheet:0", "sheet:1"] {
                let sheet = sheets
                    .get(&txn, key)
                    .and_then(|value| value.cast::<MapRef>().ok())
                    .unwrap();
                sheet.remove(&mut txn, CHARTS);
            }
            let (fingerprint, _) = fingerprint_model_with_schema(&model, 5, true, true).unwrap();
            let meta = txn.get_map(META).unwrap();
            meta.try_update(&mut txn, BASE_FINGERPRINT, fingerprint);
            meta.try_update(&mut txn, "schemaVersion", 5);
        }
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());

        let mut uncharted = model.clone();
        for sheet in &mut uncharted.sheets {
            sheet.charts.clear();
        }
        let mut authority = authority_from_update(&uncharted, &update, 121);
        let fingerprints = authority.base.fingerprints.clone();
        authority.base = Arc::new(WorkbookBase::from_model(&model).unwrap());
        Arc::make_mut(&mut authority.base).fingerprints = fingerprints;

        let materialized = authority.materialize().unwrap();
        assert_eq!(materialized.sheets[0].name, "Second");
        assert_eq!(
            materialized.sheets[0].charts[0].refs[0].formula, "Second!$A$1",
            "the fallback must follow the stable key, not the position"
        );
        assert_eq!(materialized.sheets[1].name, "First");
        assert_eq!(
            materialized.sheets[1].charts[0].refs[0].formula,
            "First!$A$1"
        );
    }

    /// A v3-v5 state carries no chart state at all, so a charted workbook
    /// pairs with it on the rest and keeps the charts it parsed. Refusing
    /// instead would strand every snapshot written before charts were shared.
    #[test]
    fn a_charted_workbook_pairs_with_a_legacy_fingerprint() {
        let mut model = WorkbookModel::default();
        model.sheets.push(charted("Report", "Report!$A$1"));
        for version in MIN_SUPPORTED_SCHEMA_VERSION..SCHEMA_VERSION {
            let update = legacy_update(&model, version, true);
            let authority = authority_from_update(&model, &update, 130 + version as u64);
            let materialized = authority.materialize().unwrap();
            assert_eq!(materialized, model, "version {version}");
            assert!(authority.upgrade_schema().unwrap());
            assert_eq!(authority.strict_materialize().unwrap().0, model);
        }
    }

    /// Charts must stay pinned to the schema that introduced them. Gating them
    /// on the current version instead reads the newest schema as pre-chart the
    /// moment the current version moves past it, which is how a released
    /// snapshot came to be unreadable in the first place.
    ///
    /// The key is optional rather than required, because repinning is an
    /// ordinary edit and undoing the assignment that won a race deletes it. An
    /// absent key falls back to the source package, so the schema gate is what
    /// decides whether the document's own chart state is read at all.
    #[test]
    fn chart_state_is_gated_on_the_schema_that_introduced_it() {
        const { assert!(CHARTS_SCHEMA_VERSION <= SCHEMA_VERSION) };
        const { assert!(HYPERLINK_SCHEMA_VERSION < CHARTS_SCHEMA_VERSION) };

        let mut model = WorkbookModel::default();
        model.sheets.push(charted("Report", "Report!$A$1"));
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 140).unwrap();

        // present: the document's own chart state wins over the fallback.
        let mut moved = model.sheets[0].charts.clone();
        moved[0].refs[0].formula = "Report!$B$9".to_owned();
        {
            let mut txn = authority.doc.transact_mut_with("test:charts-gate");
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            sheet.try_update(
                &mut txn,
                CHARTS,
                serde_json::to_string(&moved).unwrap().as_str(),
            );
        }
        assert_eq!(
            authority.materialize().unwrap().sheets[0].charts[0].refs[0].formula,
            "Report!$B$9",
            "a document at the chart schema must read its own chart state"
        );

        // absent: the source package answers, and the chart is still there.
        {
            let mut txn = authority.doc.transact_mut_with("test:charts-gate");
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            sheet.remove(&mut txn, CHARTS);
        }
        let fallen_back = authority
            .materialize()
            .expect("an absent chart key must read as absent, not as a broken workbook");
        assert_eq!(fallen_back.sheets[0].charts, model.sheets[0].charts);

        let gated = |version: i64| {
            sheet_schema_keys(version).contains(&CHARTS)
                || sheet_schema_optional_keys(version).contains(&CHARTS)
        };
        assert!(gated(CHARTS_SCHEMA_VERSION));
        assert!(!gated(CHARTS_SCHEMA_VERSION - 1));
    }

    #[test]
    fn unknown_schema_version_reports_supported_range() {
        let model = rich_model();
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 110).unwrap();
        {
            let mut txn = authority.doc.transact_mut_with("test:unknown-schema");
            let meta = txn.get_map(META).unwrap();
            meta.try_update(&mut txn, "schemaVersion", SCHEMA_VERSION + 1);
        }
        let AuthorityError::InvalidState(error) = authority.materialize().unwrap_err() else {
            panic!("expected invalid state");
        };
        assert_eq!(
            error,
            "unsupported schema version 7; supported versions are 3 through 6"
        );
    }

    fn charted(name: &str, formula: &str) -> Sheet {
        let mut sheet = Sheet::new(name);
        sheet.charts.push(xlsx_model::SheetChart {
            part: "xl/charts/chart1.xml".to_owned(),
            drawing: "xl/drawings/drawing1.xml".to_owned(),
            anchor_index: 0,
            anchor: xlsx_model::ChartAnchor::Absolute {
                pos: xlsx_model::AnchorPos::default(),
                extent: xlsx_model::AnchorExtent::default(),
            },
            refs: vec![xlsx_model::ChartRef {
                kind: xlsx_model::ChartRefKind::Values,
                formula: formula.to_owned(),
            }],
        });
        sheet
    }

    /// Removing a sheet strands the chart references that named it, so the
    /// remaining sheets' chart state must reach the shared document. A partial
    /// sync would leave a peer reading the pre-removal ranges.
    #[test]
    fn removing_a_sheet_carries_the_stranded_chart_state_into_the_document() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        model.sheets.push(charted("Report", "Data!$A$1:$A$2"));
        let mut authority = WorkbookAuthority::from_model(&model).unwrap();

        authority
            .apply_ops(
                &[Op::RemoveSheet { index: 0 }],
                SyncOrigin::User,
                &Stylesheet::default(),
            )
            .unwrap();

        let shared = authority.materialize().unwrap();
        assert_eq!(shared.sheets.len(), 1);
        assert_eq!(shared.sheets[0].charts[0].refs[0].formula, "#REF!");
    }

    /// `SetCharts` replaces a whole sheet's chart state, so it can only travel
    /// as a full semantic sync.
    #[test]
    fn set_charts_travels_as_a_full_semantic_sync() {
        let mut model = WorkbookModel::default();
        model.sheets.push(charted("Report", "Data!$A$1:$A$2"));
        let mut authority = WorkbookAuthority::from_model(&model).unwrap();

        let mut charts = model.sheets[0].charts.clone();
        charts[0].refs[0].formula = "Data!$A$1:$A$9".to_owned();
        authority
            .apply_ops(
                &[Op::SetCharts {
                    sheet: SheetId(0),
                    charts,
                }],
                SyncOrigin::User,
                &Stylesheet::default(),
            )
            .expect("SetCharts is a full sync, not a rejected partial one");

        let shared = authority.materialize().unwrap();
        assert_eq!(shared.sheets[0].charts[0].refs[0].formula, "Data!$A$1:$A$9");
    }

    #[test]
    fn oversized_peer_chart_state_is_refused_and_leaves_the_authority_intact() {
        const ELEMENT: &str = r#"{"part":"a","drawing":"b","anchorIndex":0,"anchor":{"kind":"absolute","pos":{"x":0,"y":0},"extent":{"cx":0,"cy":0}},"refs":[]}"#;
        let model = rich_model();
        for (client_id, count, expected) in [
            (111, 10_000, "sheet has too many charts"),
            (112, 80_000, "sheet chart state exceeds its size limit"),
        ] {
            let authority =
                WorkbookAuthority::from_model_with_client_id(&model, client_id).unwrap();
            let peer = Doc::with_client_id(client_id + 1);
            hydrate_doc(&peer, &authority.encode_state_as_update_v1()).unwrap();
            let before = peer.transact().state_vector();
            let mut payload = String::with_capacity(count * (ELEMENT.len() + 1) + 2);
            payload.push('[');
            for index in 0..count {
                if index > 0 {
                    payload.push(',');
                }
                payload.push_str(ELEMENT);
            }
            payload.push(']');
            {
                let mut txn = peer.transact_mut_with("test:hostile-charts");
                let sheets = txn.get_map(SHEETS).unwrap();
                let sheet = sheets
                    .get(&txn, "sheet:0")
                    .and_then(|value| value.cast::<MapRef>().ok())
                    .unwrap();
                sheet.try_update(&mut txn, CHARTS, payload.as_str());
            }
            let update = peer.transact().encode_diff_v1(&before);
            let Err(AuthorityError::InvalidState(error)) =
                authority.stage_updates_v1(&[&update], None)
            else {
                panic!("expected invalid state");
            };
            assert!(error.contains(expected), "{error}");
            assert_eq!(authority.strict_materialize().unwrap().0, model);
        }
    }

    #[test]
    fn unused_merge_edits_preserve_schema_six_state_and_identity() {
        for imported in [false, true] {
            let mut model = rich_model();
            if !imported {
                for sheet in &mut model.sheets {
                    sheet.merges.clear();
                }
            }
            let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 119).unwrap();
            let mut oracle =
                WorkbookAuthority::from_source_oracle(&model, Some(119), &[], None, &[]).unwrap();
            assert_eq!(authority.schema_version().unwrap(), 6);
            assert_eq!(authority.base.fingerprint, oracle.base.fingerprint);
            assert_eq!(
                authority.base.bootstrap_client_id,
                oracle.base.bootstrap_client_id
            );
            assert_eq!(
                authority.encode_state_as_update_v1(),
                oracle.encode_state_as_update_v1()
            );
            let state = authority.encode_state_as_update_v1();
            let SnapshotAdoption::Replacement(resumed) = authority.snapshot_replacement(&state)
            else {
                panic!("schema six snapshot refused")
            };
            assert_eq!(resumed.encode_state_as_update_v1(), state);
            assert_eq!(resumed.schema_version().unwrap(), 6);
            let ops = [Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(20, 0),
                cell: CellState {
                    value: CellValue::Text {
                        value: "edited".into(),
                    },
                    ..Default::default()
                },
            }];
            for current in [&mut authority, &mut oracle] {
                let staged = current
                    .stage_local_ops_v1(&ops, SyncOrigin::User, &model.styles)
                    .unwrap();
                current
                    .apply_local_update_v1(&staged.update, SyncOrigin::User)
                    .unwrap();
            }
            assert_eq!(
                authority.encode_state_as_update_v1(),
                oracle.encode_state_as_update_v1()
            );
            for current in [&mut authority, &mut oracle] {
                current.undo().unwrap().unwrap();
                let txn = current.doc.transact();
                let sheets = txn.get_map(SHEETS).unwrap();
                for (_, sheet) in sheets.iter(&txn) {
                    let sheet = sheet.cast::<MapRef>().unwrap();
                    assert!(
                        sheet
                            .keys(&txn)
                            .all(|key| !key.starts_with("merge:") && key != "mergeEdits")
                    );
                }
                assert_eq!(current.schema_version().unwrap(), 6);
            }
            assert_eq!(
                authority.encode_state_as_update_v1(),
                oracle.encode_state_as_update_v1()
            );
        }
    }

    #[test]
    fn flat_merge_keys_validate_in_snapshots_and_refuse_schema_six_peers() {
        let mut model = rich_model();
        model.sheets[0].merges = vec![
            CellRange::parse_a1("D4:E5").unwrap(),
            CellRange::parse_a1("A1:B2").unwrap(),
        ];
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 130).unwrap();
        {
            let mut txn = authority.doc.transact_mut();
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            for (key, present) in [
                ("merge:C1:C2", true),
                ("merge:G1:H2", true),
                ("merge:J1:K2", false),
                ("merge:D4:E5", true),
            ] {
                sheet.try_update(&mut txn, key, present);
            }
        }
        let (projected, _) = authority.strict_materialize().unwrap();
        assert_eq!(&projected.sheets[0].merges[..2], model.sheets[0].merges);
        assert!(authority.materialize_internal_oracle(true).is_err());
        snapshot_validation::validate_for_test(&authority, projected.clone()).unwrap();
        assert!(snapshot_validation::validate_for_test(&authority, model.clone()).is_err());
        let peer = WorkbookAuthority::from_model_with_client_id(&projected, 131).unwrap();
        assert_ne!(authority.base.fingerprint, peer.base.fingerprint);

        let mut old = crate::Workbook::from_model_collaborative(model.clone(), 132).unwrap();
        let mut current = crate::Workbook::from_model_collaborative(model, 133).unwrap();
        let before_vector = old.encode_state_vector_v1();
        let before_bytes = old.encode_state_as_update_v1();
        let before_model = old.model().clone();
        let options = crate::CalculationOptions::default();
        current
            .apply_ops(
                vec![
                    Op::MergeCells {
                        sheet: SheetId(0),
                        range: CellRange::parse_a1("J1:K2").unwrap(),
                    },
                    Op::SetCell {
                        sheet: SheetId(0),
                        at: CellRef::new(20, 0),
                        cell: CellState {
                            value: CellValue::Text {
                                value: "new value".into(),
                            },
                            ..Default::default()
                        },
                    },
                ],
                options,
            )
            .unwrap();
        let update = current.encode_diff_v1(&before_vector).unwrap();
        SCHEMA_SIX_KEY_VALIDATION.set(true);
        let result = old.apply_update_v1(&update, options);
        SCHEMA_SIX_KEY_VALIDATION.set(false);
        let error = result.unwrap_err().to_string();
        assert!(
            error.contains("sheet sheet:0 keys do not match the schema"),
            "{error}"
        );
        assert_eq!(old.encode_state_vector_v1(), before_vector);
        assert_eq!(old.encode_state_as_update_v1(), before_bytes);
        assert_eq!(old.model(), &before_model);
        assert!(!old.has_pending_updates_for_test());
    }

    #[test]
    fn local_merge_edits_remove_shadowed_keys() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Merges"));
        let options = crate::CalculationOptions::default();
        for unmerge in [false, true] {
            let mut left = crate::Workbook::from_model_collaborative(model.clone(), 180).unwrap();
            let mut right = crate::Workbook::from_model_collaborative(model.clone(), 181).unwrap();
            for (peer, range) in [(&mut left, "A1:B2"), (&mut right, "B2:C3")] {
                peer.apply_ops(
                    vec![Op::MergeCells {
                        sheet: SheetId(0),
                        range: CellRange::parse_a1(range).unwrap(),
                    }],
                    options,
                )
                .unwrap();
            }
            let update = right
                .encode_diff_v1(&left.encode_state_vector_v1())
                .unwrap();
            left.apply_update_v1(&update, options).unwrap();
            assert_eq!(
                left.model().sheets[0].merges,
                vec![CellRange::parse_a1("A1:B2").unwrap()]
            );
            let replacement = CellRange::parse_a1("B2:D4").unwrap();
            let op = if unmerge {
                Op::UnmergeCells {
                    sheet: SheetId(0),
                    range: CellRange::parse_a1("A1:B2").unwrap(),
                }
            } else {
                Op::MergeCells {
                    sheet: SheetId(0),
                    range: replacement,
                }
            };
            left.apply_ops(vec![op], options).unwrap();
            let expected = if unmerge { vec![] } else { vec![replacement] };
            assert_eq!(left.model().sheets[0].merges, expected);
            right
                .apply_update_v1(
                    &left
                        .encode_diff_v1(&right.encode_state_vector_v1())
                        .unwrap(),
                    options,
                )
                .unwrap();
            assert_eq!(left.model(), right.model());
            let doc = Doc::new();
            hydrate_doc(&doc, &left.encode_state_as_update_v1()).unwrap();
            let txn = doc.transact();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            let keys = sheet
                .keys(&txn)
                .filter(|key| key.starts_with(MERGE_PREFIX))
                .collect::<BTreeSet<_>>();
            assert_eq!(
                keys,
                if unmerge {
                    BTreeSet::new()
                } else {
                    BTreeSet::from(["merge:B2:D4"])
                }
            );
        }
    }

    #[test]
    fn local_merge_ranges_normalize_anchors() {
        let range = CellRange::parse_a1("A1:B2").unwrap();
        let options = crate::CalculationOptions::default();
        for collaborative in [false, true] {
            for imported in [false, true] {
                let mut model = WorkbookModel::default();
                let mut sheet = Sheet::new("Merges");
                if imported {
                    sheet.merges.push(range);
                }
                model.sheets.push(sheet);
                let mut workbook = if collaborative {
                    crate::Workbook::from_model_collaborative(model, 182).unwrap()
                } else {
                    crate::Workbook::from_model(model).unwrap()
                };
                if !imported {
                    workbook
                        .apply_ops(
                            vec![Op::MergeCells {
                                sheet: SheetId(0),
                                range: CellRange::parse_a1("$A$1:$B$2").unwrap(),
                            }],
                            options,
                        )
                        .unwrap();
                    assert_eq!(workbook.model().sheets[0].merges, vec![range]);
                    let reopened = crate::Workbook::open(&workbook.save().unwrap()).unwrap();
                    assert_eq!(reopened.model().sheets[0].merges, vec![range]);
                    let doc = Doc::new();
                    hydrate_doc(&doc, &workbook.encode_state_as_update_v1()).unwrap();
                    let txn = doc.transact();
                    let sheet = txn
                        .get_map(SHEETS)
                        .unwrap()
                        .get(&txn, "sheet:0")
                        .unwrap()
                        .cast::<MapRef>()
                        .unwrap();
                    assert!(sheet.keys(&txn).all(|key| !key.contains('$')));
                    let Out::Any(baseline) = sheet.get(&txn, MERGES).unwrap() else {
                        panic!("missing merges")
                    };
                    assert!(
                        merges_from_any(&baseline)
                            .unwrap()
                            .iter()
                            .all(|range| !range.to_a1().contains('$'))
                    );
                }
                workbook
                    .apply_ops(
                        vec![Op::UnmergeCells {
                            sheet: SheetId(0),
                            range: CellRange::parse_a1("$A1:B$2").unwrap(),
                        }],
                        options,
                    )
                    .unwrap();
                assert!(workbook.model().sheets[0].merges.is_empty());
            }
        }
    }

    #[test]
    fn overlapping_merge_baselines_are_refused_by_both_validators() {
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Merges");
        sheet.merges.push(CellRange::parse_a1("A1:B2").unwrap());
        model.sheets.push(sheet);
        for edit in [
            None,
            Some(("merge:D1:E2", true)),
            Some(("merge:B2:C3", false)),
        ] {
            let authority = WorkbookAuthority::from_model_with_client_id(&model, 183).unwrap();
            let mut projected = model.clone();
            {
                let mut txn = authority.doc.transact_mut();
                let sheet = txn
                    .get_map(SHEETS)
                    .unwrap()
                    .get(&txn, "sheet:0")
                    .unwrap()
                    .cast::<MapRef>()
                    .unwrap();
                sheet.try_update(
                    &mut txn,
                    MERGES,
                    merges_to_any(&[
                        CellRange::parse_a1("A1:B2").unwrap(),
                        CellRange::parse_a1("B2:C3").unwrap(),
                    ]),
                );
                if let Some((key, present)) = edit {
                    sheet.try_update(&mut txn, key, present);
                    if present {
                        projected.sheets[0]
                            .merges
                            .push(CellRange::parse_a1("D1:E2").unwrap());
                    }
                }
            }
            let strict = authority.strict_materialize();
            let snapshot = snapshot_validation::validate_for_test(&authority, projected);
            assert!(strict.is_err(), "strict accepted {edit:?}");
            assert!(snapshot.is_err(), "snapshot accepted {edit:?}");
            assert!(strict.unwrap_err().contains("merged cell ranges overlap"));
            assert!(
                snapshot
                    .unwrap_err()
                    .to_string()
                    .contains("merged cell ranges overlap")
            );
        }
    }

    #[test]
    fn local_single_cell_merges_are_refused() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Merges"));
        for collaborative in [false, true] {
            let mut workbook = if collaborative {
                crate::Workbook::from_model_collaborative(model.clone(), 184).unwrap()
            } else {
                crate::Workbook::from_model(model.clone()).unwrap()
            };
            let before = workbook.encode_state_as_update_v1();
            for address in ["A1", "$A1:A$1"] {
                let result = workbook.apply_ops(
                    vec![Op::MergeCells {
                        sheet: SheetId(0),
                        range: CellRange::parse_a1(address).unwrap(),
                    }],
                    crate::CalculationOptions::default(),
                );
                assert!(matches!(result, Err(crate::Error::InvalidOperation(_))));
                assert_eq!(workbook.encode_state_as_update_v1(), before);
                assert!(workbook.model().sheets[0].merges.is_empty());
            }
        }
    }

    #[test]
    fn flat_merge_keys_refuse_nested_values_unknown_fields_and_excess_entries() {
        for kind in [0, 1, 2] {
            let model = rich_model();
            let authority = WorkbookAuthority::from_model_with_client_id(&model, 140).unwrap();
            let original = authority.encode_state_as_update_v1();
            let peer = Doc::with_client_id(141);
            hydrate_doc(&peer, &original).unwrap();
            {
                let mut txn = peer.transact_mut();
                let sheets = txn.get_map(SHEETS).unwrap();
                let sheet = sheets
                    .get(&txn, "sheet:0")
                    .unwrap()
                    .cast::<MapRef>()
                    .unwrap();
                match kind {
                    0 => {
                        let _: MapRef = sheet.get_or_init(&mut txn, "merge:A1:B2");
                    }
                    1 => {
                        sheet.try_update(&mut txn, "mergeEdits", false);
                    }
                    _ => {
                        for row in 1..=merges::MAX_MERGE_EDITS_PER_SHEET + 1 {
                            sheet.try_update(&mut txn, format!("merge:Z{row}:AA{row}"), false);
                        }
                    }
                }
            }
            let hostile = WorkbookAuthority::hydrated(peer, authority.base.clone(), 0);
            assert!(hostile.strict_materialize().is_err());
            assert!(snapshot_validation::validate_for_test(&hostile, model).is_err());
            assert!(matches!(
                authority.stage_updates_v1(&[&hostile.encode_state_as_update_v1()], None),
                Err(AuthorityError::InvalidState(_))
            ));
            assert_eq!(authority.encode_state_as_update_v1(), original);
        }
    }

    #[test]
    fn merge_edit_limit_has_identical_strict_and_streaming_projection() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Merges"));
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 150).unwrap();
        {
            let mut txn = authority.doc.transact_mut();
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            for row in 1..=merges::MAX_MERGE_EDITS_PER_SHEET {
                sheet.try_update(&mut txn, format!("merge:A{row}:B{row}"), true);
            }
        }
        let projected = authority.strict_materialize().unwrap().0;
        assert_eq!(
            projected.sheets[0].merges.len(),
            merges::MAX_MERGE_EDITS_PER_SHEET
        );
        snapshot_validation::validate_for_test(&authority, projected).unwrap();
    }

    #[test]
    fn standalone_full_sync_flattens_merge_edits_into_the_baseline() {
        let model = rich_model();
        let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 160).unwrap();
        let range = CellRange::parse_a1("J1:K2").unwrap();
        let merge = [Op::MergeCells {
            sheet: SheetId(0),
            range,
        }];
        let staged = authority
            .stage_local_ops_v1(&merge, SyncOrigin::User, &model.styles)
            .unwrap();
        authority
            .apply_local_update_v1(&staged.update, SyncOrigin::User)
            .unwrap();
        authority
            .apply_ops(
                &[Op::UnmergeCells {
                    sheet: SheetId(0),
                    range,
                }],
                SyncOrigin::User,
                &model.styles,
            )
            .unwrap();
        assert_eq!(
            authority.strict_materialize().unwrap().0.sheets[0].merges,
            model.sheets[0].merges
        );
        let txn = authority.doc.transact();
        let sheets = txn.get_map(SHEETS).unwrap();
        let sheet = sheets
            .get(&txn, "sheet:0")
            .unwrap()
            .cast::<MapRef>()
            .unwrap();
        assert!(sheet.keys(&txn).all(|key| !key.starts_with(MERGE_PREFIX)));
        let Out::Any(baseline) = sheet.get(&txn, MERGES).unwrap() else {
            panic!("missing baseline")
        };
        assert_eq!(merges_from_any(&baseline).unwrap(), model.sheets[0].merges);
    }

    #[test]
    fn concurrent_merge_and_unmerge_of_one_baseline_key_converge() {
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Baseline");
        sheet.merges.push(CellRange::parse_a1("A1:B2").unwrap());
        model.sheets.push(sheet);
        let mut left = WorkbookAuthority::from_model_with_client_id(&model, 170).unwrap();
        let mut right = WorkbookAuthority::from_model_with_client_id(&model, 171).unwrap();
        let vector = left.encode_state_vector_v1();
        for (authority, present) in [(&left, true), (&right, false)] {
            let mut txn = authority.doc.transact_mut();
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            sheet.try_update(&mut txn, "merge:A1:B2", present);
        }
        let left_update = left.encode_diff_v1(&vector).unwrap();
        let right_update = right.encode_diff_v1(&vector).unwrap();
        for (authority, update) in [(&mut left, right_update), (&mut right, left_update)] {
            let staged = authority.stage_updates_v1(&[&update], None).unwrap();
            authority
                .apply_staged_update_v1(&staged.commit_update)
                .unwrap();
        }
        let projection = left.strict_materialize().unwrap().0;
        assert_eq!(projection, right.strict_materialize().unwrap().0);
        for authority in [&left, &right] {
            snapshot_validation::validate_for_test(authority, projection.clone()).unwrap();
            let txn = authority.doc.transact();
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            let Out::Any(baseline) = sheet.get(&txn, MERGES).unwrap() else {
                panic!("missing baseline")
            };
            assert_eq!(merges_from_any(&baseline).unwrap(), model.sheets[0].merges);
        }
    }

    #[test]
    fn malformed_peer_merge_edits_are_refused_without_mutating_the_authority() {
        let model = rich_model();
        for (client_id, key, value) in [
            (113, "A0:B2", Any::Bool(true)),
            (114, "A1:B2", Any::from("true")),
            (115, "a1:b2", Any::Bool(true)),
            (116, "B2:A1", Any::Bool(true)),
            (117, "A1:A1", Any::Bool(true)),
            (118, "A1:XFE2", Any::Bool(true)),
            (119, "A1:B1048577", Any::Bool(true)),
            (120, "", Any::Bool(true)),
            (121, "A1:B2", Any::Null),
            (122, "A1:B2", Any::from(1)),
            (123, "$A$1:$B$2", Any::Bool(true)),
            (124, "$A1:B2", Any::Bool(false)),
            (125, "A1", Any::Bool(true)),
            (126, "A1", Any::Bool(false)),
        ] {
            let authority =
                WorkbookAuthority::from_model_with_client_id(&model, client_id).unwrap();
            let peer = Doc::with_client_id(client_id + 100);
            hydrate_doc(&peer, &authority.encode_state_as_update_v1()).unwrap();
            let before = peer.transact().state_vector();
            {
                let mut txn = peer.transact_mut_with("test:hostile-merges");
                let sheets = txn.get_map(SHEETS).unwrap();
                let sheet = sheets
                    .get(&txn, "sheet:0")
                    .and_then(|value| value.cast::<MapRef>().ok())
                    .unwrap();
                sheet.try_update(&mut txn, format!("{MERGE_PREFIX}{key}"), value);
            }
            let update = peer.transact().encode_diff_v1(&before);
            let original = authority.encode_state_as_update_v1();
            let hostile = WorkbookAuthority::hydrated(peer, authority.base.clone(), 0);
            assert!(snapshot_validation::validate_for_test(&hostile, model.clone()).is_err());
            assert!(matches!(
                authority.stage_updates_v1(&[&update], None),
                Err(AuthorityError::InvalidState(_))
            ));
            assert_eq!(authority.strict_materialize().unwrap().0, model);
            assert_eq!(authority.encode_state_as_update_v1(), original);
        }
    }

    #[test]
    fn unused_hyperlink_edits_preserve_schema_six_state_and_identity() {
        for imported in [false, true] {
            let mut model = rich_model();
            if !imported {
                for sheet in &mut model.sheets {
                    sheet.hyperlinks.clear();
                }
            }
            let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 229).unwrap();
            let mut oracle =
                WorkbookAuthority::from_source_oracle(&model, Some(229), &[], None, &[]).unwrap();
            assert_eq!(authority.schema_version().unwrap(), 6);
            assert_eq!(authority.base.fingerprint, oracle.base.fingerprint);
            assert_eq!(
                authority.base.bootstrap_client_id,
                oracle.base.bootstrap_client_id
            );
            assert_eq!(
                authority.encode_state_as_update_v1(),
                oracle.encode_state_as_update_v1()
            );
            let state = authority.encode_state_as_update_v1();
            let SnapshotAdoption::Replacement(resumed) = authority.snapshot_replacement(&state)
            else {
                panic!("schema six snapshot refused")
            };
            assert_eq!(resumed.encode_state_as_update_v1(), state);
            assert_eq!(resumed.schema_version().unwrap(), 6);
            let ops = [Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(20, 0),
                cell: CellState {
                    value: CellValue::Text {
                        value: "edited".into(),
                    },
                    ..Default::default()
                },
            }];
            for current in [&mut authority, &mut oracle] {
                let staged = current
                    .stage_local_ops_v1(&ops, SyncOrigin::User, &model.styles)
                    .unwrap();
                current
                    .apply_local_update_v1(&staged.update, SyncOrigin::User)
                    .unwrap();
            }
            assert_eq!(
                authority.encode_state_as_update_v1(),
                oracle.encode_state_as_update_v1()
            );
            for current in [&mut authority, &mut oracle] {
                current.undo().unwrap().unwrap();
                let txn = current.doc.transact();
                let sheets = txn.get_map(SHEETS).unwrap();
                for (_, sheet) in sheets.iter(&txn) {
                    let sheet = sheet.cast::<MapRef>().unwrap();
                    assert!(
                        sheet
                            .keys(&txn)
                            .all(|key| !key.starts_with("hyperlink:") && key != "hyperlinkEdits")
                    );
                }
                assert_eq!(current.schema_version().unwrap(), 6);
            }
            assert_eq!(
                authority.encode_state_as_update_v1(),
                oracle.encode_state_as_update_v1()
            );
        }
    }

    fn test_hyperlink(address: &str, location: &str) -> Hyperlink {
        Hyperlink {
            range: CellRange::parse_a1(address).unwrap(),
            external_target: None,
            location: Some(location.to_string()),
            tooltip: None,
            display: None,
        }
    }

    #[test]
    fn hyperlink_schema_six_key_validation_resets_after_unwinding() {
        assert!(!SCHEMA_SIX_KEY_VALIDATION.get());
        let result = catch_unwind(|| {
            with_schema_six_key_validation(|| {
                assert!(SCHEMA_SIX_KEY_VALIDATION.get());
                panic!("reset flag");
            });
        });
        assert!(result.is_err());
        assert!(!SCHEMA_SIX_KEY_VALIDATION.get());
        with_schema_six_key_validation(|| {
            with_schema_six_key_validation(|| assert!(SCHEMA_SIX_KEY_VALIDATION.get()));
            assert!(SCHEMA_SIX_KEY_VALIDATION.get());
        });
        assert!(!SCHEMA_SIX_KEY_VALIDATION.get());
    }

    #[test]
    fn hyperlink_snapshots_refuse_noncanonical_baselines_with_or_without_edits() {
        for edited in [false, true] {
            let link = test_hyperlink("A1", "Data!A1");
            let canonical = serde_json::to_string(std::slice::from_ref(&link)).unwrap();
            let mut extra = serde_json::to_value(&link).unwrap();
            extra
                .as_object_mut()
                .unwrap()
                .insert("unknown".into(), serde_json::Value::Bool(true));
            let mut omitted = serde_json::to_value(&link).unwrap();
            omitted.as_object_mut().unwrap().remove("tooltip");
            for json in [
                format!(" {canonical}"),
                canonical.replace("null", " null"),
                serde_json::to_string(&[extra]).unwrap(),
                serde_json::to_string(&[omitted]).unwrap(),
            ] {
                let mut model = WorkbookModel::default();
                model.sheets.push(Sheet::new("Data"));
                model.sheets[0].hyperlinks = vec![test_hyperlink("A1", "Data!A1")];
                let authority = WorkbookAuthority::from_model_with_client_id(&model, 260).unwrap();
                {
                    let mut txn = authority.doc.transact_mut();
                    let sheet = txn
                        .get_map(SHEETS)
                        .unwrap()
                        .get(&txn, "sheet:0")
                        .unwrap()
                        .cast::<MapRef>()
                        .unwrap();
                    sheet.try_update(&mut txn, HYPERLINKS, json);
                    if edited {
                        sheet.try_update(&mut txn, "hyperlink:Z99", false);
                    }
                }
                assert!(authority.strict_materialize().is_err());
                assert!(snapshot_validation::validate_for_test(&authority, model).is_err());
            }
        }
    }

    #[test]
    fn legacy_snapshots_refuse_hyperlink_keys_in_both_validators() {
        for version in MIN_SUPPORTED_SCHEMA_VERSION..SCHEMA_VERSION {
            let mut model = WorkbookModel::default();
            model.sheets.push(Sheet::new("Data"));
            let source = WorkbookAuthority::from_model_with_client_id(&model, 261).unwrap();
            let doc = Doc::with_client_id(263);
            hydrate_doc(&doc, &legacy_update(&model, version, true)).unwrap();
            let authority = WorkbookAuthority::hydrated(doc, source.base.clone(), 0);
            assert!(authority.strict_materialize().is_ok());
            snapshot_validation::validate_for_test(&authority, model.clone()).unwrap();
            {
                let mut txn = authority.doc.transact_mut();
                let sheet = txn
                    .get_map(SHEETS)
                    .unwrap()
                    .get(&txn, "sheet:0")
                    .unwrap()
                    .cast::<MapRef>()
                    .unwrap();
                sheet.try_update(&mut txn, "hyperlink:A1", false);
            }
            assert!(authority.strict_materialize().is_err());
            assert!(snapshot_validation::validate_for_test(&authority, model).is_err());
        }
    }

    #[test]
    fn empty_optional_hyperlink_fields_are_accepted_only_in_baselines() {
        for baseline in [false, true] {
            for long in [false, true] {
                for field in 0..4 {
                    let mut model = WorkbookModel::default();
                    model.sheets.push(Sheet::new("Data"));
                    let authority =
                        WorkbookAuthority::from_model_with_client_id(&model, 262).unwrap();
                    let mut link = test_hyperlink("A1", "Data!A1");
                    link.external_target = Some("https://example.com".into());
                    if long {
                        link.tooltip = Some("a".repeat(300));
                        link.display = Some("b".repeat(300));
                    }
                    match field {
                        0 => link.external_target = Some(String::new()),
                        1 => link.location = Some(String::new()),
                        2 => link.tooltip = Some(String::new()),
                        _ => link.display = Some(String::new()),
                    }
                    {
                        let mut txn = authority.doc.transact_mut();
                        let sheet = txn
                            .get_map(SHEETS)
                            .unwrap()
                            .get(&txn, "sheet:0")
                            .unwrap()
                            .cast::<MapRef>()
                            .unwrap();
                        if baseline {
                            sheet.try_update(
                                &mut txn,
                                HYPERLINKS,
                                serde_json::to_string(std::slice::from_ref(&link)).unwrap(),
                            );
                        } else {
                            sheet.try_update(
                                &mut txn,
                                "hyperlink:A1",
                                serde_json::to_string(&link).unwrap(),
                            );
                        }
                    }
                    if baseline {
                        model.sheets[0].hyperlinks = vec![link];
                        assert_eq!(authority.strict_materialize().unwrap().0, model);
                        snapshot_validation::validate_for_test(&authority, model).unwrap();
                    } else {
                        assert!(authority.strict_materialize().is_err());
                        assert!(snapshot_validation::validate_for_test(&authority, model).is_err());
                    }
                }
            }
        }
    }

    #[test]
    fn imported_empty_hyperlink_fields_validate_through_edit_undo_and_redo() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
        parts
            .iter_mut()
            .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
            .unwrap()
            .1 =
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><hyperlinks><hyperlink ref="A1" location="Data!D1" tooltip="" display=""/></hyperlinks></worksheet>"#.to_vec();
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        let mut workbook = crate::Workbook::open_collaborative(&bytes, 264).unwrap();
        let base = Arc::new(WorkbookBase::from_model(workbook.model()).unwrap());
        let options = crate::CalculationOptions::default();
        for step in 0..4 {
            match step {
                1 => {
                    workbook
                        .set_hyperlink(SheetId(0), test_hyperlink("B2", "Data!D2"), options)
                        .unwrap();
                }
                2 => {
                    workbook.undo(options).unwrap();
                }
                3 => {
                    workbook.redo(options).unwrap();
                }
                _ => {}
            }
            let doc = Doc::new();
            hydrate_doc(&doc, &workbook.encode_state_as_update_v1()).unwrap();
            let authority = WorkbookAuthority::hydrated(doc, base.clone(), 0);
            assert_eq!(authority.strict_materialize().unwrap().0, *workbook.model());
            snapshot_validation::validate_for_test(&authority, workbook.model().clone()).unwrap();
        }
    }

    #[test]
    fn anchored_hyperlink_baselines_validate_and_refuse_peer_edits_atomically() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        model.sheets[0].hyperlinks = vec![test_hyperlink("$A$1", "Data!D1")];
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 265).unwrap();
        assert_eq!(authority.strict_materialize().unwrap().0, model);
        snapshot_validation::validate_for_test(&authority, model.clone()).unwrap();
        let original = authority.encode_state_as_update_v1();
        let vector = authority.encode_state_vector_v1();
        for address in ["A1", "B2"] {
            for remove in [false, true] {
                let doc = Doc::with_client_id(266);
                hydrate_doc(&doc, &original).unwrap();
                {
                    let mut txn = doc.transact_mut();
                    let sheet = txn
                        .get_map(SHEETS)
                        .unwrap()
                        .get(&txn, "sheet:0")
                        .unwrap()
                        .cast::<MapRef>()
                        .unwrap();
                    let key = format!("hyperlink:{address}");
                    if remove {
                        sheet.try_update(&mut txn, key, false);
                    } else {
                        sheet.try_update(
                            &mut txn,
                            key,
                            serde_json::to_string(&test_hyperlink(address, "Data!D2")).unwrap(),
                        );
                    }
                }
                let hostile = WorkbookAuthority::hydrated(doc, authority.base.clone(), 0);
                assert_eq!(
                    hostile.strict_materialize().unwrap_err(),
                    "cannot edit anchored imported hyperlinks"
                );
                assert_eq!(
                    snapshot_validation::validate_for_test(&hostile, model.clone())
                        .unwrap_err()
                        .to_string(),
                    "cannot edit anchored imported hyperlinks"
                );
                assert!(matches!(
                    authority.stage_updates_v1(&[&hostile.encode_state_as_update_v1()], None),
                    Err(AuthorityError::InvalidState(_))
                ));
                assert_eq!(authority.encode_state_as_update_v1(), original);
                assert_eq!(authority.encode_state_vector_v1(), vector);
                assert_eq!(authority.strict_materialize().unwrap().0, model);
            }
        }
    }

    #[test]
    fn flat_hyperlink_keys_validate_in_snapshots_and_refuse_schema_six_peers() {
        let mut model = rich_model();
        model.sheets[0].hyperlinks = vec![
            test_hyperlink("D4:E5", "Data!D1"),
            test_hyperlink("A1:B2", "Data!A1"),
        ];
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 230).unwrap();
        {
            let mut txn = authority.doc.transact_mut();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            for address in ["C1:C2", "G1:H2"] {
                sheet.try_update(
                    &mut txn,
                    format!("{HYPERLINK_PREFIX}{address}"),
                    serde_json::to_string(&test_hyperlink(address, "Data!A1")).unwrap(),
                );
            }
            sheet.try_update(&mut txn, "hyperlink:J1:K2", false);
        }
        let (projected, _) = authority.strict_materialize().unwrap();
        assert_eq!(
            &projected.sheets[0].hyperlinks[..2],
            model.sheets[0].hyperlinks
        );
        assert!(authority.materialize_internal_oracle(true).is_err());
        with_schema_six_key_validation(|| {
            assert!(snapshot_validation::validate_for_test(&authority, projected.clone()).is_err());
        });
        snapshot_validation::validate_for_test(&authority, projected).unwrap();
        assert!(snapshot_validation::validate_for_test(&authority, model.clone()).is_err());

        let mut old = crate::Workbook::from_model_collaborative(model.clone(), 231).unwrap();
        let mut current = crate::Workbook::from_model_collaborative(model, 232).unwrap();
        let before_vector = old.encode_state_vector_v1();
        let before_bytes = old.encode_state_as_update_v1();
        let before_model = old.model().clone();
        let options = crate::CalculationOptions::default();
        current
            .apply_ops(
                vec![
                    Op::SetHyperlink {
                        sheet: SheetId(0),
                        hyperlink: test_hyperlink("J1:K2", "Data!A1"),
                    },
                    Op::SetCell {
                        sheet: SheetId(0),
                        at: CellRef::new(20, 0),
                        cell: CellState {
                            value: CellValue::Text {
                                value: "new value".into(),
                            },
                            ..Default::default()
                        },
                    },
                ],
                options,
            )
            .unwrap();
        let update = current.encode_diff_v1(&before_vector).unwrap();
        let result = with_schema_six_key_validation(|| old.apply_update_v1(&update, options));
        let error = result.unwrap_err().to_string();
        assert!(
            error.contains("sheet sheet:0 keys do not match the schema"),
            "{error}"
        );
        assert_eq!(old.encode_state_vector_v1(), before_vector);
        assert_eq!(old.encode_state_as_update_v1(), before_bytes);
        assert_eq!(old.model(), &before_model);
        assert!(!old.has_pending_updates_for_test());
    }

    #[test]
    fn flat_hyperlink_keys_refuse_nested_values_unknown_fields_and_excess_entries() {
        for kind in [0, 1, 2] {
            let model = rich_model();
            let authority = WorkbookAuthority::from_model_with_client_id(&model, 240).unwrap();
            let original = authority.encode_state_as_update_v1();
            let vector = authority.encode_state_vector_v1();
            let peer = Doc::with_client_id(241);
            hydrate_doc(&peer, &original).unwrap();
            {
                let mut txn = peer.transact_mut();
                let sheet = txn
                    .get_map(SHEETS)
                    .unwrap()
                    .get(&txn, "sheet:0")
                    .unwrap()
                    .cast::<MapRef>()
                    .unwrap();
                match kind {
                    0 => {
                        sheet.insert(&mut txn, "hyperlink:A1", MapPrelim::default());
                    }
                    1 => {
                        sheet.try_update(&mut txn, "hyperlinkEdits", false);
                    }
                    _ => {
                        for row in 1..=merges::MAX_MERGE_EDITS_PER_SHEET + 1 {
                            sheet.try_update(&mut txn, format!("hyperlink:Z{row}"), false);
                        }
                    }
                }
            }
            let hostile = WorkbookAuthority::hydrated(peer, authority.base.clone(), 0);
            let strict = hostile.strict_materialize().unwrap_err();
            let snapshot = snapshot_validation::validate_for_test(&hostile, model.clone())
                .unwrap_err()
                .to_string();
            if kind != 1 {
                assert_eq!(strict, snapshot);
            }
            if kind == 2 {
                assert_eq!(strict, "too many sheet hyperlink edits");
            }
            assert!(matches!(
                authority.stage_updates_v1(&[&hostile.encode_state_as_update_v1()], None),
                Err(AuthorityError::InvalidState(_))
            ));
            assert_eq!(authority.encode_state_vector_v1(), vector);
            assert_eq!(authority.encode_state_as_update_v1(), original);
            assert_eq!(authority.strict_materialize().unwrap().0, model);
        }
    }

    #[test]
    fn hyperlink_edit_limit_has_identical_strict_and_streaming_projection() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Links"));
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 250).unwrap();
        {
            let mut txn = authority.doc.transact_mut();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            for row in 1..=merges::MAX_MERGE_EDITS_PER_SHEET {
                let address = format!("A{row}");
                let link = test_hyperlink(&address, "Links!A1");
                sheet.try_update(
                    &mut txn,
                    format!("{HYPERLINK_PREFIX}{address}"),
                    serde_json::to_string(&link).unwrap(),
                );
            }
        }
        let projected = authority.strict_materialize().unwrap().0;
        assert_eq!(
            projected.sheets[0].hyperlinks.len(),
            merges::MAX_MERGE_EDITS_PER_SHEET
        );
        snapshot_validation::validate_for_test(&authority, projected).unwrap();
    }

    #[test]
    fn standalone_full_sync_flattens_hyperlink_edits_into_the_baseline() {
        let model = rich_model();
        for full_sync in [false, true] {
            let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 260).unwrap();
            let hyperlink = test_hyperlink("J1:K2", "Data!A1");
            let staged = authority
                .stage_local_ops_v1(
                    &[Op::SetHyperlink {
                        sheet: SheetId(0),
                        hyperlink: hyperlink.clone(),
                    }],
                    SyncOrigin::User,
                    &model.styles,
                )
                .unwrap();
            authority
                .apply_local_update_v1(&staged.update, SyncOrigin::User)
                .unwrap();
            let expected = if full_sync {
                authority.strict_materialize().unwrap().0
            } else {
                model.clone()
            };
            let op = if full_sync {
                Op::SetHyperlinks {
                    sheet: SheetId(0),
                    hyperlinks: expected.sheets[0].hyperlinks.clone(),
                }
            } else {
                Op::RemoveHyperlink {
                    sheet: SheetId(0),
                    range: hyperlink.range,
                }
            };
            let flattened = authority
                .stage_standalone_ops_v1(&[op], SyncOrigin::User, &model.styles, None)
                .unwrap();
            authority
                .apply_local_update_v1(&flattened.update, SyncOrigin::User)
                .unwrap();
            assert_eq!(authority.strict_materialize().unwrap().0, expected);
            let txn = authority.doc.transact();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            assert!(
                sheet
                    .keys(&txn)
                    .all(|key| !key.starts_with(HYPERLINK_PREFIX))
            );
            let Out::Any(Any::String(baseline)) = sheet.get(&txn, HYPERLINKS).unwrap() else {
                panic!("missing baseline")
            };
            assert_eq!(
                decode_hyperlinks(&baseline).unwrap(),
                expected.sheets[0].hyperlinks
            );
        }
    }

    #[test]
    fn concurrent_hyperlink_edit_and_baseline_removal_converge() {
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Baseline");
        sheet.hyperlinks.push(test_hyperlink("A1", "Baseline!B1"));
        model.sheets.push(sheet);
        let mut left = crate::Workbook::from_model_collaborative(model.clone(), 270).unwrap();
        let mut right = crate::Workbook::from_model_collaborative(model.clone(), 271).unwrap();
        let vector = left.encode_state_vector_v1();
        let options = crate::CalculationOptions::default();
        left.set_hyperlink(SheetId(0), test_hyperlink("A1", "Baseline!C1"), options)
            .unwrap();
        right
            .remove_hyperlink(SheetId(0), CellRange::parse_a1("A1").unwrap(), options)
            .unwrap();
        let left_update = left.encode_diff_v1(&vector).unwrap();
        let right_update = right.encode_diff_v1(&vector).unwrap();
        left.apply_update_v1(&right_update, options).unwrap();
        right.apply_update_v1(&left_update, options).unwrap();
        assert_eq!(left.model(), right.model());
        assert_eq!(
            left.encode_state_vector_v1(),
            right.encode_state_vector_v1()
        );
        for workbook in [&left, &right] {
            let doc = Doc::new();
            hydrate_doc(&doc, &workbook.encode_state_as_update_v1()).unwrap();
            let authority = WorkbookAuthority::hydrated(
                doc,
                Arc::new(WorkbookBase::from_model(&model).unwrap()),
                0,
            );
            snapshot_validation::validate_for_test(&authority, workbook.model().clone()).unwrap();
            let txn = authority.doc.transact();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            let Out::Any(Any::String(baseline)) = sheet.get(&txn, HYPERLINKS).unwrap() else {
                panic!("missing baseline")
            };
            assert_eq!(
                decode_hyperlinks(&baseline).unwrap(),
                model.sheets[0].hyperlinks
            );
        }
    }

    #[test]
    fn overlapping_hyperlink_baselines_have_identical_validation() {
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Links");
        sheet.hyperlinks = vec![
            test_hyperlink("A1:B2", "Links!D1"),
            test_hyperlink("B2:C3", "Links!D2"),
        ];
        model.sheets.push(sheet);
        for edited in [false, true] {
            let authority = WorkbookAuthority::from_model_with_client_id(&model, 280).unwrap();
            if edited {
                let mut txn = authority.doc.transact_mut();
                let sheet = txn
                    .get_map(SHEETS)
                    .unwrap()
                    .get(&txn, "sheet:0")
                    .unwrap()
                    .cast::<MapRef>()
                    .unwrap();
                sheet.try_update(&mut txn, "hyperlink:A1:B2", false);
            }
            let strict = authority.strict_materialize();
            let snapshot = snapshot_validation::validate_for_test(&authority, model.clone());
            if edited {
                assert_eq!(
                    strict.unwrap_err(),
                    "cannot edit overlapping imported hyperlinks"
                );
                assert_eq!(
                    snapshot.unwrap_err().to_string(),
                    "cannot edit overlapping imported hyperlinks"
                );
            } else {
                assert_eq!(strict.unwrap().0, model);
                snapshot.unwrap();
            }
        }
    }

    #[test]
    fn local_hyperlink_edits_remove_shadowed_keys() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Links"));
        let options = crate::CalculationOptions::default();
        for remove in [false, true] {
            let mut left = crate::Workbook::from_model_collaborative(model.clone(), 290).unwrap();
            let mut right = crate::Workbook::from_model_collaborative(model.clone(), 291).unwrap();
            for (peer, address) in [(&mut left, "A1:B2"), (&mut right, "B2:C3")] {
                peer.set_hyperlink(SheetId(0), test_hyperlink(address, "Links!E1"), options)
                    .unwrap();
            }
            left.apply_update_v1(
                &right
                    .encode_diff_v1(&left.encode_state_vector_v1())
                    .unwrap(),
                options,
            )
            .unwrap();
            assert_eq!(
                left.model().sheets[0].hyperlinks,
                vec![test_hyperlink("A1:B2", "Links!E1")]
            );
            if remove {
                left.remove_hyperlink(SheetId(0), CellRange::parse_a1("A1:B2").unwrap(), options)
                    .unwrap();
            } else {
                left.set_hyperlink(SheetId(0), test_hyperlink("B2:D4", "Links!E2"), options)
                    .unwrap();
            }
            right
                .apply_update_v1(
                    &left
                        .encode_diff_v1(&right.encode_state_vector_v1())
                        .unwrap(),
                    options,
                )
                .unwrap();
            assert_eq!(left.model(), right.model());
            let expected = if remove {
                vec![]
            } else {
                vec![test_hyperlink("B2:D4", "Links!E2")]
            };
            assert_eq!(left.model().sheets[0].hyperlinks, expected);
            let doc = Doc::new();
            hydrate_doc(&doc, &left.encode_state_as_update_v1()).unwrap();
            let txn = doc.transact();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            let keys = sheet
                .keys(&txn)
                .filter(|key| key.starts_with(HYPERLINK_PREFIX))
                .collect::<BTreeSet<_>>();
            assert_eq!(
                keys,
                if remove {
                    BTreeSet::new()
                } else {
                    BTreeSet::from(["hyperlink:B2:D4"])
                }
            );
        }
    }

    #[test]
    fn malformed_peer_hyperlink_edits_are_refused_without_mutating_the_authority() {
        let model = rich_model();
        let link = serde_json::to_string(&model.sheets[0].hyperlinks[0]).unwrap();
        for (client_id, key, value) in [
            (211, "A0", Any::from(link.as_str())),
            (212, "A1", Any::from("not json")),
            (213, "A1", Any::from(link.as_str())),
            (214, "A1", Any::Bool(true)),
            (215, "a1", Any::Bool(false)),
            (216, "$A$1", Any::Bool(false)),
            (217, "A1:A1", Any::Bool(false)),
            (218, "B2:A1", Any::Bool(false)),
            (219, "XFE1", Any::Bool(false)),
            (220, "A1048577", Any::Bool(false)),
            (223, "", Any::Bool(false)),
            (224, "A1", Any::Null),
            (225, "A1", Any::from(1)),
            (226, "A1", Any::from("[]")),
        ] {
            let authority =
                WorkbookAuthority::from_model_with_client_id(&model, client_id).unwrap();
            let peer = Doc::with_client_id(client_id + 100);
            hydrate_doc(&peer, &authority.encode_state_as_update_v1()).unwrap();
            let original = authority.encode_state_as_update_v1();
            let before = peer.transact().state_vector();
            {
                let mut txn = peer.transact_mut_with("test:hostile-hyperlink");
                let sheets = txn.get_map(SHEETS).unwrap();
                let sheet = sheets
                    .get(&txn, "sheet:0")
                    .and_then(|value| value.cast::<MapRef>().ok())
                    .unwrap();
                sheet.try_update(&mut txn, format!("{HYPERLINK_PREFIX}{key}"), value);
            }
            let update = peer.transact().encode_diff_v1(&before);
            let hostile = WorkbookAuthority::hydrated(peer, authority.base.clone(), 0);
            let strict = hostile.strict_materialize().unwrap_err();
            let snapshot = snapshot_validation::validate_for_test(&hostile, model.clone())
                .unwrap_err()
                .to_string();
            assert_eq!(strict, snapshot);
            assert!(matches!(
                authority.stage_updates_v1(&[&update], None),
                Err(AuthorityError::InvalidState(_))
            ));
            assert_eq!(authority.strict_materialize().unwrap().0, model);
            assert_eq!(authority.encode_state_as_update_v1(), original);
        }
    }

    #[test]
    fn peer_cannot_replace_the_imported_hyperlink_baseline_as_a_layout_edit() {
        let model = rich_model();
        let source = WorkbookAuthority::from_model_with_client_id(&model, 221).unwrap();
        let target = WorkbookAuthority::from_model_with_client_id(&model, 222).unwrap();
        let frozen = target.structure().unwrap();
        let vector = target.encode_state_vector_v1();
        {
            let mut txn = source
                .doc
                .transact_mut_with("test:replace-hyperlink-baseline");
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            sheet.try_update(&mut txn, HYPERLINKS, "[]");
        }
        let staged = target
            .stage_updates_v1(&[&source.encode_diff_v1(&vector).unwrap()], None)
            .unwrap();
        assert!(!staged.structure.same_except_local_layout(&frozen));
        assert_eq!(target.materialize().unwrap(), model);
    }

    #[test]
    fn untouched_imported_merge_order_survives_projection() {
        let doc = Doc::new();
        let edits = doc.transact_mut().get_or_insert_map("edits");
        let baseline = vec![
            CellRange::parse_a1("D4:E5").unwrap(),
            CellRange::parse_a1("A1:B2").unwrap(),
        ];
        assert_eq!(
            project_merges(baseline.clone(), &edits, &doc.transact()).unwrap(),
            baseline
        );
        edits.try_update(&mut doc.transact_mut(), "merge:C1:C2", true);
        let mut expected = baseline;
        expected.push(CellRange::parse_a1("C1:C2").unwrap());
        assert_eq!(
            project_merges(expected[..2].to_vec(), &edits, &doc.transact()).unwrap(),
            expected
        );
    }

    #[test]
    fn formula_content_is_one_atomic_payload() {
        let formula = Cell {
            value: CellValue::Error {
                value: ErrorValue::Ref,
            },
            formula: Some("Missing!A1".into()),
            style: None,
        };
        assert_eq!(
            content_from_any(&content_to_any(&formula).unwrap()).unwrap(),
            formula
        );
    }

    #[test]
    fn strict_decoders_reject_trailing_and_impossible_vectors() {
        assert!(decode_state_vector_v1(&[1]).is_err());
        assert!(decode_state_vector_v1(&[0, 0]).is_err());
        assert!(decode_update_v1(&[0, 0, 0]).is_err());
    }

    #[test]
    fn update_decoder_rejects_overflowing_block_and_delete_clocks() {
        let overflowing_skip = [
            1,
            1,
            1,
            0xff,
            0xff,
            0xff,
            0xff,
            0x0f,
            BLOCK_SKIP_REF_NUMBER,
            1,
            0,
        ];
        let overflowing_gc = [
            1,
            1,
            1,
            0x80,
            0x80,
            0x80,
            0x80,
            0x08,
            BLOCK_GC_REF_NUMBER,
            1,
            0,
        ];
        let overflowing_delete = [0, 1, 1, 1, 0xff, 0xff, 0xff, 0xff, 0x0f, 1];
        assert!(decode_update_v1(&overflowing_skip).is_err());
        assert!(decode_update_v1(&overflowing_gc).is_err());
        assert!(decode_update_v1(&overflowing_delete).is_err());
    }

    #[test]
    fn checked_decoders_bound_counts_and_reject_noncanonical_ids() {
        let too_many_blocks = [1, 0xff, 0xff, 0xff, 0xff, 0x0f];
        let too_many_delete_ranges = [0, 1, 1, 0xff, 0xff, 0xff, 0xff, 0x0f];
        assert!(decode_update_v1(&too_many_blocks).is_err());
        assert!(decode_update_v1(&too_many_delete_ranges).is_err());
        assert!(decode_state_vector_v1(&[1, 0x81, 0x80, 0, 0]).is_err());
    }

    #[test]
    fn checked_decoder_rejects_overlong_signed_varints() {
        for bytes in [
            vec![125, 0x40],
            vec![125, 0x80, 0],
            vec![125, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0],
        ] {
            assert!(CheckedDecoderV1::new(&bytes).decode_any(0).is_err());
        }
        assert_eq!(
            CheckedDecoderV1::new(&[125, 0xc1, 1])
                .decode_any(0)
                .unwrap(),
            Any::Number(-65.0)
        );
    }

    #[test]
    fn state_vector_entry_count_is_bounded_before_decode() {
        let error = decode_state_vector_v1(&[0x81, 0x80, 0x04]).unwrap_err();
        assert!(error.contains("65536-entry limit"), "{error}");
    }

    #[test]
    fn bootstrap_client_id_conflicts_are_rejected_locally() {
        let model = rich_model();
        let base = WorkbookBase::from_model(&model).unwrap();
        assert!(matches!(
            WorkbookAuthority::from_model_with_client_id(&model, base.bootstrap_client_id),
            Err(AuthorityError::ClientIdConflict(_))
        ));
    }

    #[test]
    fn shadowed_deletes_wait_for_missing_items_and_commit_with_pending_blocks() {
        let model = rich_model();
        let source = WorkbookAuthority::from_model_with_client_id(&model, 25).unwrap();
        let target = WorkbookAuthority::from_model_with_client_id(&model, 26).unwrap();
        let sheet = {
            let txn = source.doc.transact();
            txn.get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap()
        };
        let insertions = ["merge:B4:D5", "merge:B4:D6", "merge:B5:D6"].map(|key| {
            let mut txn = source.doc.transact_mut();
            sheet.insert(&mut txn, key, true);
            txn.encode_update_v1()
        });
        let deletion = {
            let mut txn = source.doc.transact_mut();
            sheet.remove(&mut txn, "merge:B4:D5");
            txn.encode_update_v1()
        };
        let before = target.encode_state_as_update_v1();
        let staged = target.stage_updates_v1(&[&deletion], None).unwrap();
        assert!(staged.pending);
        assert!(!staged.effective);
        assert_eq!(staged.commit_update, Update::EMPTY_V1);
        assert_eq!(target.encode_state_as_update_v1(), before);

        let staged = target.stage_updates_v1(&[&insertions[0]], None).unwrap();
        target
            .apply_staged_update_v1(&staged.commit_update)
            .unwrap();
        let vector = target.encode_state_vector_v1();
        let structure = target.structure().unwrap();
        let pending = Update::merge_updates([
            Update::decode_v1(&insertions[2]).unwrap(),
            Update::decode_v1(&deletion).unwrap(),
        ])
        .encode_v1();
        let staged = target.stage_updates_v1(&[&pending], None).unwrap();
        assert!(staged.pending);
        assert!(staged.effective);
        assert_eq!(staged.model, model);
        assert_eq!(staged.structure, structure);
        target
            .apply_staged_update_v1(&staged.commit_update)
            .unwrap();
        assert_eq!(target.encode_state_vector_v1(), vector);
        let staged = target.stage_updates_v1(&[&pending], None).unwrap();
        assert!(staged.pending);
        assert!(!staged.effective);
        assert_eq!(staged.commit_update, Update::EMPTY_V1);

        let staged = target.stage_updates_v1(&[&insertions[1]], None).unwrap();
        target
            .apply_staged_update_v1(&staged.commit_update)
            .unwrap();
        let staged = target.stage_updates_v1(&[&pending], None).unwrap();
        assert!(!staged.pending);
        assert!(staged.effective);
        target
            .apply_staged_update_v1(&staged.commit_update)
            .unwrap();
        assert_eq!(
            target.encode_state_as_update_v1(),
            source.encode_state_as_update_v1()
        );
        assert!(
            !target
                .stage_updates_v1(&[&deletion], None)
                .unwrap()
                .effective
        );
    }

    #[test]
    fn shared_map_replacement_changes_the_frozen_structure() {
        let model = rich_model();
        let source = WorkbookAuthority::from_model_with_client_id(&model, 21).unwrap();
        let target = WorkbookAuthority::from_model_with_client_id(&model, 22).unwrap();
        let target_structure = target.structure().unwrap();
        let target_vector = target.encode_state_vector_v1();

        {
            let mut txn = source.doc.transact_mut_with("test:replace-map");
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:1")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            sheet.insert(&mut txn, CONTENTS, MapPrelim::default());
        }

        let update = source.encode_diff_v1(&target_vector).unwrap();
        let staged = target.stage_updates_v1(&[&update], None).unwrap();
        assert_eq!(staged.model, target.materialize().unwrap());
        assert_ne!(staged.structure, target_structure);
    }

    #[test]
    fn peer_cannot_replace_the_imported_merge_baseline_as_a_layout_edit() {
        let model = rich_model();
        let source = WorkbookAuthority::from_model_with_client_id(&model, 23).unwrap();
        let target = WorkbookAuthority::from_model_with_client_id(&model, 24).unwrap();
        let frozen = target.structure().unwrap();
        let vector = target.encode_state_vector_v1();
        {
            let mut txn = source.doc.transact_mut_with("test:replace-merge-baseline");
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            sheet.try_update(
                &mut txn,
                MERGES,
                merges_to_any(&[CellRange::parse_a1("A1:B2").unwrap()]),
            );
        }
        let staged = target
            .stage_updates_v1(&[&source.encode_diff_v1(&vector).unwrap()], None)
            .unwrap();
        assert!(!staged.structure.same_except_local_layout(&frozen));
        assert_eq!(target.materialize().unwrap(), model);
    }

    #[test]
    fn retained_sheet_maps_stay_valid_and_keep_identity_through_undo() {
        let model = rich_model();
        let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 31).unwrap();
        authority
            .apply_ops(
                &[Op::RemoveSheet { index: 1 }],
                SyncOrigin::User,
                &Stylesheet::default(),
            )
            .unwrap();
        let (_, removed) = authority.strict_materialize().unwrap();
        assert_eq!(removed.sheet_keys, ["sheet:0"]);
        assert_eq!(removed.shared_types.len(), 2);
        let retained = removed.shared_types["sheet:1"].clone();

        authority
            .apply_ops(
                &[Op::AddSheet {
                    index: 1,
                    name: "Second".into(),
                }],
                SyncOrigin::Undo,
                &Stylesheet::default(),
            )
            .unwrap();
        let (restored, structure) = authority.strict_materialize().unwrap();
        assert_eq!(restored, model);
        assert_eq!(structure.shared_types["sheet:1"], retained);
    }

    fn sliding_chart(name: &str) -> Sheet {
        let mut sheet = Sheet::new(name);
        sheet.charts.push(SheetChart {
            part: "xl/charts/chart1.xml".to_owned(),
            drawing: "xl/drawings/drawing1.xml".to_owned(),
            anchor_index: 0,
            anchor: ChartAnchor::TwoCell {
                from: xlsx_model::AnchorCell::default(),
                to: xlsx_model::AnchorCell {
                    col: 4,
                    col_off: 0,
                    row: 8,
                    row_off: 0,
                },
                edit_as: AnchorEditAs::TwoCell,
            },
            refs: vec![ChartRef {
                kind: xlsx_model::ChartRefKind::Values,
                formula: "Data!$A$1:$A$2".to_owned(),
            }],
        });
        sheet
    }

    fn sliding_model() -> WorkbookModel {
        let mut model = WorkbookModel::default();
        model.sheets.push(sliding_chart("Report"));
        model
    }

    /// An update that assigns a sheet's chart state without touching the
    /// structure generation, as a peer forked from `authority`'s state.
    fn peer_chart_update(authority: &WorkbookAuthority, client_id: u64, charts: &str) -> Vec<u8> {
        let peer = Doc::with_client_id(client_id);
        hydrate_doc(&peer, &authority.encode_state_as_update_v1()).unwrap();
        let before = peer.transact().state_vector();
        {
            let mut txn = peer.transact_mut_with("test:peer-charts");
            let sheets = txn.get_map(SHEETS).unwrap();
            let sheet = sheets
                .get(&txn, "sheet:0")
                .and_then(|value| value.cast::<MapRef>().ok())
                .unwrap();
            sheet.try_update(&mut txn, CHARTS, charts);
        }
        peer.transact().encode_diff_v1(&before)
    }

    fn slid_anchor(cols: i64) -> ChartAnchor {
        ChartAnchor::TwoCell {
            from: xlsx_model::AnchorCell {
                col: cols as u32,
                ..xlsx_model::AnchorCell::default()
            },
            to: xlsx_model::AnchorCell {
                col: 4 + cols as u32,
                col_off: 0,
                row: 8,
                row_off: 0,
            },
            edit_as: AnchorEditAs::TwoCell,
        }
    }

    /// The freeze pins what a chart *is*, so a peer cannot pass a remap off as
    /// a move. The structure generation is untouched here, which is what makes
    /// this the identity check rather than the generation counter.
    #[test]
    fn a_peer_cannot_disguise_a_chart_remap_as_a_move() {
        let model = sliding_model();
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 41).unwrap();
        let frozen = authority.structure().unwrap();
        let generation = frozen.generation;

        for (label, charts) in [
            (
                "refs",
                r#"[{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{"kind":"twoCell","from":{"col":0,"colOff":0,"row":0,"rowOff":0},"to":{"col":4,"colOff":0,"row":8,"rowOff":0},"edit_as":"twoCell"},"refs":[{"kind":"values","formula":"Hijacked!$A$1"}]}]"#,
            ),
            (
                "part",
                r#"[{"part":"xl/charts/other.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{"kind":"twoCell","from":{"col":0,"colOff":0,"row":0,"rowOff":0},"to":{"col":4,"colOff":0,"row":8,"rowOff":0},"edit_as":"twoCell"},"refs":[{"kind":"values","formula":"Data!$A$1:$A$2"}]}]"#,
            ),
            (
                "anchorIndex",
                r#"[{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":3,"anchor":{"kind":"twoCell","from":{"col":0,"colOff":0,"row":0,"rowOff":0},"to":{"col":4,"colOff":0,"row":8,"rowOff":0},"edit_as":"twoCell"},"refs":[{"kind":"values","formula":"Data!$A$1:$A$2"}]}]"#,
            ),
        ] {
            let update = peer_chart_update(&authority, 42, charts);
            let staged = authority.stage_updates_v1(&[&update], None).unwrap();
            assert_eq!(
                staged.structure.generation, generation,
                "{label} must not move the generation, or it proves nothing"
            );
            assert_ne!(
                staged.structure, frozen,
                "a rewritten {label} must change the frozen structure"
            );
        }
    }

    /// A move may slide a grid-anchored chart and nothing else. The drawing
    /// writer refuses a changed kind, `editAs` mode or one-cell extent, so the
    /// freeze has to refuse them too rather than accept a workbook that can no
    /// longer be saved.
    #[test]
    fn a_peer_cannot_reshape_an_anchor_a_save_can_only_slide() {
        let model = sliding_model();
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 43).unwrap();
        let frozen = authority.structure().unwrap();

        for (label, charts) in [
            (
                "editAs",
                r#"[{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{"kind":"twoCell","from":{"col":0,"colOff":0,"row":0,"rowOff":0},"to":{"col":4,"colOff":0,"row":8,"rowOff":0},"edit_as":"oneCell"},"refs":[{"kind":"values","formula":"Data!$A$1:$A$2"}]}]"#,
            ),
            (
                "kind",
                r#"[{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{"kind":"oneCell","from":{"col":0,"colOff":0,"row":0,"rowOff":0},"extent":{"cx":100000,"cy":100000}},"refs":[{"kind":"values","formula":"Data!$A$1:$A$2"}]}]"#,
            ),
        ] {
            let update = peer_chart_update(&authority, 44, charts);
            let staged = authority.stage_updates_v1(&[&update], None).unwrap();
            assert_ne!(
                staged.structure, frozen,
                "a rewritten anchor {label} must change the frozen structure"
            );
        }

        // sliding the same anchor across the grid is the one accepted change.
        let slid = r#"[{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{"kind":"twoCell","from":{"col":2,"colOff":0,"row":0,"rowOff":0},"to":{"col":6,"colOff":0,"row":8,"rowOff":0},"edit_as":"twoCell"},"refs":[{"kind":"values","formula":"Data!$A$1:$A$2"}]}]"#;
        let update = peer_chart_update(&authority, 44, slid);
        let staged = authority.stage_updates_v1(&[&update], None).unwrap();
        assert_eq!(staged.structure, frozen);
    }

    /// A checkpoint has to put back everything a history step moves: the
    /// document, both stacks, and the identity the stacks name. Restoring into
    /// a fresh document would look right and leave the history unusable — the
    /// undo manager drops entries belonging to a document it does not know,
    /// silently, a whole stack at a time — so the proof is that undo still
    /// works afterwards.
    #[test]
    fn a_restored_checkpoint_brings_back_a_working_history() {
        let model = sliding_model();
        let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 61).unwrap();
        let commit = |authority: &mut WorkbookAuthority, to| {
            let ops = [Op::SetChartAnchor {
                sheet: SheetId(0),
                frame: "xl/drawings/drawing1.xml#0".to_owned(),
                part: "xl/charts/chart1.xml".to_owned(),
                from: authority.materialize().unwrap().sheets[0].charts[0].anchor,
                to,
            }];
            let staged = authority
                .stage_local_ops_v1(&ops, SyncOrigin::User, &Stylesheet::default())
                .unwrap();
            authority
                .apply_local_update_v1(&staged.update, SyncOrigin::User)
                .unwrap();
        };

        commit(&mut authority, slid_anchor(2));
        let checkpoint = authority.checkpoint();
        let vector = authority.encode_state_vector_v1();
        let anchor = authority.materialize().unwrap().sheets[0].charts[0].anchor;
        let depth = authority.undo_depth();
        assert!(depth > 0);

        commit(&mut authority, slid_anchor(5));
        assert_ne!(
            authority.materialize().unwrap().sheets[0].charts[0].anchor,
            anchor
        );

        authority.restore(checkpoint).unwrap();
        assert_eq!(
            authority.materialize().unwrap().sheets[0].charts[0].anchor,
            anchor,
            "restore must bring the document back"
        );
        assert_eq!(
            authority.encode_state_vector_v1(),
            vector,
            "restore must wind the document back, not forward"
        );
        assert_eq!(
            authority.undo_depth(),
            depth,
            "restore must bring back the stacks"
        );

        // the history still belongs to this document, so it can still be spent
        let undone = authority
            .undo()
            .expect("a restored history must still apply")
            .expect("the entry is on the stack");
        assert_eq!(
            undone.model.sheets[0].charts[0].anchor,
            model.sheets[0].charts[0].anchor
        );
        assert_eq!(authority.undo_depth(), depth - 1);
    }

    /// A peer's chart write can lose the merge to a concurrent local one and
    /// sit in the document unseen. Undoing the local write must not strand the
    /// authority on state it cannot read: the step is refused and the replica
    /// stays exactly as usable as it was.
    #[test]
    fn a_hidden_chart_conflict_leaves_the_replica_usable() {
        let model = sliding_model();
        let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 304).unwrap();
        let frozen = authority.structure().unwrap();
        let hostile = peer_chart_update(
            &authority,
            111,
            r#"[{"part":"xl/charts/chart1.xml","drawing":"xl/drawings/drawing1.xml","anchorIndex":0,"anchor":{"kind":"twoCell","from":{"col":0,"colOff":0,"row":0,"rowOff":0},"to":{"col":4,"colOff":0,"row":8,"rowOff":0},"edit_as":"twoCell"},"refs":[{"kind":"values","formula":"Hijacked!$A$1"}]}]"#,
        );

        let ops = [Op::SetChartAnchor {
            sheet: SheetId(0),
            frame: "xl/drawings/drawing1.xml#0".to_owned(),
            part: "xl/charts/chart1.xml".to_owned(),
            from: model.sheets[0].charts[0].anchor,
            to: slid_anchor(2),
        }];
        let local = authority
            .stage_local_ops_v1(&ops, SyncOrigin::User, &Stylesheet::default())
            .unwrap();
        authority
            .apply_local_update_v1(&local.update, SyncOrigin::User)
            .unwrap();

        // the hostile value loses the merge, so the freeze sees only the move.
        let staged = authority.stage_updates_v1(&[&hostile], None).unwrap();
        assert_eq!(staged.structure, frozen);
        authority
            .apply_staged_update_v1(&staged.commit_update)
            .unwrap();
        let merged = authority.materialize().unwrap();
        assert_eq!(
            merged.sheets[0].charts[0].refs,
            model.sheets[0].charts[0].refs
        );

        // undoing the move succeeds, takes the anchor back, and never lets the
        // hidden remap through.
        let undone = authority
            .undo()
            .expect("a hidden conflict must not fail the undo")
            .expect("the move is on the stack, so undo must do something");
        assert_eq!(undone.structure, frozen);
        let after = authority.materialize().unwrap();
        assert_eq!(
            after.sheets[0].charts[0].refs,
            model.sheets[0].charts[0].refs
        );
        assert_eq!(
            after.sheets[0].charts[0].anchor, model.sheets[0].charts[0].anchor,
            "undo must put the chart back where it started"
        );
        assert_eq!(authority.structure().unwrap(), frozen);
        assert!(
            !authority.can_undo(),
            "one undo must consume exactly the one entry, not drain a longer stack"
        );
        let _ = merged;
    }

    fn tall_sheet(name: &str, rows: u32) -> Sheet {
        let mut sheet = Sheet::new(name);
        for row in 0..rows {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row),
                    },
                    formula: (row % 3 == 0).then(|| format!("{row}+0")),
                    style: Some(0),
                },
            );
        }
        sheet
    }

    #[test]
    fn seeding_the_bootstrap_reports_progress() {
        thread_local! {
            static CALLS: std::cell::Cell<u32> = const { std::cell::Cell::new(0) };
        }
        fn count() {
            CALLS.with(|calls| calls.set(calls.get() + 1));
        }
        let cells = 60_000;
        let mut model = WorkbookModel::default();
        model.styles.cell_xfs.push(Xf::default());
        model.sheets.push(tall_sheet("Data", cells));
        let base = WorkbookBase::from_model_with_style_readings(&model, &[], None, &[]).unwrap();
        let bootstrap = Doc::with_client_id(base.bootstrap_client_id);
        let keys = ["sheet:0".to_string()];

        xlsx_model::progress::set_hook(Some(count));
        let seeded = seed_in_chunks(
            &bootstrap,
            &base,
            &model,
            &keys,
            style_entry,
            usize::MAX,
            |_| Ok(()),
        );
        xlsx_model::progress::set_hook(None);

        assert!(seeded.is_ok());
        assert!(CALLS.with(std::cell::Cell::get) >= cells / xlsx_model::progress::EVERY);
    }

    /// The reference seed: the whole bootstrap in one transaction of
    /// whole-sheet syncs.
    fn single_transaction_bootstrap(model: &WorkbookModel, base: &WorkbookBase) -> Vec<u8> {
        let doc = Doc::with_client_id(base.bootstrap_client_id);
        let keys = (0..model.sheets.len())
            .map(|index| format!("sheet:{index}"))
            .collect::<Vec<_>>();
        {
            let mut txn = doc.transact_mut_with(BOOTSTRAP_ORIGIN);
            let cell_formats = txn.get_or_insert_map(CELL_FORMATS);
            let (key, payload) = cell_format_entry(&CellFormat::default()).unwrap();
            cell_formats.try_update(&mut txn, key, payload);
            let mut style_keys = Vec::new();
            for index in 0..model.styles.cell_xfs.len() {
                let (key, payload) = style_entry(&model.styles, index as u32).unwrap();
                style_keys.push(key.clone());
                cell_formats.try_update(&mut txn, key, payload);
            }
            let meta = txn.get_or_insert_map(META);
            meta.insert(&mut txn, BASE_FINGERPRINT, base.fingerprint.as_str());
            meta.insert(&mut txn, "schemaVersion", SCHEMA_VERSION);
            meta.insert(&mut txn, STRUCTURE_GENERATION, 0_i64);
            let order = txn.get_or_insert_array(SHEET_ORDER);
            order.insert_range(&mut txn, 0, keys.iter().cloned());
            let sheets = txn.get_or_insert_map(SHEETS);
            for (key, sheet) in keys.iter().zip(&model.sheets) {
                let sheet_map = sheets.insert(&mut txn, key.as_str(), MapPrelim::default());
                sync_sheet_with_style_key(&sheet_map, &mut txn, sheet, |style| {
                    Ok(style_keys[style as usize].clone())
                })
                .unwrap();
            }
        }
        doc.transact()
            .encode_state_as_update_v1(&StateVector::default())
    }

    /// Opens `model` seeding `chunk` cell entries per bootstrap transaction and
    /// asserts it lands on the state one bootstrap update gives; returns how
    /// many chunks the seed took.
    fn open_in_chunks_and_compare(model: &WorkbookModel, chunk: usize) -> usize {
        let base = WorkbookBase::from_model_with_style_readings(model, &[], None, &[]).unwrap();
        let single = single_transaction_bootstrap(model, &base);

        let bootstrap = Doc::with_client_id(base.bootstrap_client_id);
        let keys = (0..model.sheets.len())
            .map(|index| format!("sheet:{index}"))
            .collect::<Vec<_>>();
        let mut chunks = 0;
        let count = |_: Vec<u8>| {
            chunks += 1;
            Ok(())
        };
        seed_in_chunks(&bootstrap, &base, model, &keys, style_entry, chunk, count).unwrap();
        let chunked = bootstrap
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        assert_eq!(chunked, single);

        let doc = Doc::with_client_id(101);
        hydrate_local_doc(&doc, &single).unwrap();
        let expected = WorkbookAuthority::hydrated(doc, Arc::new(base), 0);
        let (actual, model_out, structure) =
            WorkbookAuthority::from_source_with_seed_chunk(model, Some(101), &[], None, &[], chunk)
                .unwrap();
        assert_eq!(actual.client_id(), expected.client_id());
        assert_eq!(
            actual.encode_state_vector_v1(),
            expected.encode_state_vector_v1()
        );
        assert_eq!(
            actual.encode_state_as_update_v1(),
            expected.encode_state_as_update_v1()
        );
        let materialized = expected.strict_materialize().unwrap();
        assert_eq!((model_out, structure), materialized);
        chunks
    }

    #[test]
    fn a_chunked_open_matches_one_bootstrap_update_across_sheets() {
        let mut model = rich_model();
        model.sheets.push(tall_sheet("Tall", 40));
        model.sheets[2]
            .merges
            .push(CellRange::new(CellRef::new(50, 0), CellRef::new(51, 1)));
        model.sheets[2].row_heights.insert(5, 18.0);

        let chunks = open_in_chunks_and_compare(&model, 3);
        assert!(chunks > model.sheets.len());
        assert_eq!(open_in_chunks_and_compare(&model, SEED_CHUNK), 1);
    }

    #[test]
    fn a_chunked_open_splits_one_sheet_across_chunks() {
        let mut model = WorkbookModel::default();
        model.styles.cell_xfs.push(Xf::default());
        model.sheets.push(tall_sheet("Data", 50));

        assert!(open_in_chunks_and_compare(&model, 7) >= 100 / 7);
    }
}
