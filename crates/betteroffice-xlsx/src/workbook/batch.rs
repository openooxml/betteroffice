//! Version-checked, all-or-nothing host edit batches.
//!
//! A batch resolves every step against the state its expected version names, stages the
//! resulting operations on a copy of the model and of the authority, rehearses the authority
//! update, and adopts both as one commit, or returns a refusal with nothing changed.

use std::collections::{BTreeSet, HashSet};
use std::fmt;

use serde::de::{DeserializeOwned, Error as _};
use serde::ser::Error as _;
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde_json::Value;
use xlsx_calc::graph::DepGraph;
use xlsx_calc::{RecalcResult, parse_formula, recalc_after};
use xlsx_model::{CellRange, CellRef, CellValue, Sheet, SheetId};
use xlsx_ops::{CellState, NumberFormatMutation, Op, OpError, StylePatch, StyleProperty};
use xlsx_render::display_text;

use super::staging::CommitHistory;
use super::target::{CellTarget, FindRequest, RangeTarget, ReadRequest, Resolved, cell_target};
use super::{
    StagedApply, Workbook, authority_error, calculation_result, cell_states_semantically_equal,
    current_cell_state, edit_cell_state, validate_cell_state, validate_model_sheets, validate_op,
};
use crate::authority::{SyncOrigin, cell_format_fits};
use crate::{Error, Result};

const MAX_STEPS: usize = 128;
/// Most target and guard cells one batch may address.
const MAX_CELL_VISITS: u64 = 100_000;
/// Most UTF-16 units of inputs, formulas and guard text one batch may carry.
const MAX_TEXT_UNITS: u64 = 1_048_576;
/// Largest encoded authority a batch stages a copy of.
const MAX_STAGING_BYTES: usize = 256 * 1024 * 1024;
/// Most cells each calculation diagnostics list reports.
pub(super) const MAX_DIAGNOSTIC_CELLS: usize = 10_000;
/// Largest JSON request the `*_json` entry points decode.
pub const MAX_REQUEST_BYTES: usize = 16 * 1024 * 1024;
/// Largest JSON response a read, search, validation or application returns.
pub const MAX_RESPONSE_BYTES: usize = 64 * 1024 * 1024;

/// An opaque, session-scoped optimistic-concurrency token. Compare for equality only.
#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(transparent)]
pub struct DocumentVersion(String);

impl DocumentVersion {
    pub(super) fn new(nonce: &str, changes: u64) -> Self {
        Self(format!("{nonce}-{changes}"))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl From<String> for DocumentVersion {
    fn from(token: String) -> Self {
        Self(token)
    }
}

impl From<&str> for DocumentVersion {
    fn from(token: &str) -> Self {
        Self(token.to_owned())
    }
}

impl fmt::Display for DocumentVersion {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

/// A random nonce naming one authority, independent of its client id and of other instances.
pub(super) fn mint_nonce() -> String {
    yrs::uuid_v4().replace('-', "")
}

/// Who asked for a batch. Provenance only: it neither grants permission nor selects history.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum EditSource {
    #[default]
    Host,
    Agent,
}

/// How an applied batch enters undo history.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum EditHistory {
    /// Exactly one undo step.
    #[default]
    Separate,
    /// No undo step; existing undo and redo entries stay. Standalone history replays inverse
    /// operations, so undoing an older step that wrote the same cell still overwrites it.
    None,
}

/// Conditions on one cell's current state; an absent field imposes none.
#[derive(Clone, Debug, Default, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CellGuard {
    /// The stored value, a formula's current result included.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<CellValue>,
    /// Formula source without the leading `=`; `Some(None)` asserts there is no formula.
    #[serde(
        default,
        deserialize_with = "present",
        skip_serializing_if = "Option::is_none"
    )]
    pub formula: Option<Option<String>>,
    /// The exact text the engine formats the cell as.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_text: Option<String>,
}

fn present<'de, D: Deserializer<'de>, T: Deserialize<'de>>(
    deserializer: D,
) -> std::result::Result<Option<T>, D::Error> {
    T::deserialize(deserializer).map(Some)
}

/// Row-major guards shaped exactly like the step's target.
#[derive(Clone, Debug, Default, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct StepGuard {
    pub cells: Vec<Vec<CellGuard>>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(
    tag = "op",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum EditOperation {
    /// What a user would type into each cell, parsed against the cell's current number format.
    SetCellInputs {
        target: RangeTarget,
        inputs: Vec<Vec<String>>,
    },
    /// Formula source without the leading `=`, stored as a formula whatever the cell's format.
    SetFormulas {
        target: RangeTarget,
        formulas: Vec<Vec<String>>,
    },
    /// Relative mutations resolve against each cell's current format.
    SetNumberFormat {
        target: RangeTarget,
        #[serde(deserialize_with = "number_format")]
        format: NumberFormatMutation,
    },
    PatchStyle {
        target: RangeTarget,
        patch: StylePatch,
    },
}

impl EditOperation {
    pub fn target(&self) -> &RangeTarget {
        match self {
            Self::SetCellInputs { target, .. }
            | Self::SetFormulas { target, .. }
            | Self::SetNumberFormat { target, .. }
            | Self::PatchStyle { target, .. } => target,
        }
    }
}

/// Accepts `"percent"` as well as `{"type": "percent"}`.
fn number_format<'de, D: Deserializer<'de>>(
    deserializer: D,
) -> std::result::Result<NumberFormatMutation, D::Error> {
    let value = match Value::deserialize(deserializer)? {
        Value::String(kind) => serde_json::json!({ "type": kind }),
        value => value,
    };
    serde_json::from_value(value).map_err(D::Error::custom)
}

/// One batch step; on the wire the operation's fields sit beside `expect`.
#[derive(Clone, Debug, PartialEq)]
pub struct EditStep {
    pub operation: EditOperation,
    pub expect: Option<StepGuard>,
}

impl EditStep {
    pub fn new(operation: EditOperation) -> Self {
        Self {
            operation,
            expect: None,
        }
    }
}

impl Serialize for EditStep {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut value = serde_json::to_value(&self.operation).map_err(S::Error::custom)?;
        if let (Some(expect), Value::Object(object)) = (&self.expect, &mut value) {
            object.insert(
                "expect".to_owned(),
                serde_json::to_value(expect).map_err(S::Error::custom)?,
            );
        }
        value.serialize(serializer)
    }
}

impl<'de> Deserialize<'de> for EditStep {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> std::result::Result<Self, D::Error> {
        let mut value = Value::deserialize(deserializer)?;
        let object = value
            .as_object_mut()
            .ok_or_else(|| D::Error::custom("an edit step must be an object"))?;
        let expect = object.remove("expect").unwrap_or(Value::Null);
        Ok(Self {
            expect: serde_json::from_value(expect).map_err(D::Error::custom)?,
            operation: serde_json::from_value(value).map_err(D::Error::custom)?,
        })
    }
}

#[derive(Clone, Copy, Debug, Default, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CalculationRequest {
    /// The serial date volatile functions see; omitted, they have no clock.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub now_serial: Option<f64>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct EditRequest {
    pub expect_version: DocumentVersion,
    #[serde(default)]
    pub source: EditSource,
    #[serde(default)]
    pub history: EditHistory,
    #[serde(default)]
    pub calculation: CalculationRequest,
    pub steps: Vec<EditStep>,
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum EditFailureCode {
    StaleVersion,
    MissingTarget,
    ContentMismatch,
    OverlappingSteps,
    LockedTarget,
    Unsupported,
    InvalidStep,
    LimitExceeded,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditFailure {
    pub code: EditFailureCode,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub step_index: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub conflicting_step_index: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<RangeTarget>,
    pub message: String,
}

impl EditFailure {
    fn at(mut self, step_index: u32) -> Self {
        self.step_index = Some(step_index);
        self
    }
}

/// A policy refusal; the workbook is untouched at `version`.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditRefusal {
    pub version: DocumentVersion,
    pub failure: EditFailure,
}

pub(super) fn failure(
    code: EditFailureCode,
    message: String,
    target: Option<RangeTarget>,
) -> EditFailure {
    EditFailure {
        code,
        step_index: None,
        conflicting_step_index: None,
        target,
        message,
    }
}

pub(super) fn refusal(version: DocumentVersion, failure: EditFailure) -> EditRefusal {
    EditRefusal { version, failure }
}

/// What one step did, in request order.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditReceipt {
    pub step_index: u32,
    pub changed: bool,
    pub target: RangeTarget,
    /// Row-major.
    pub changed_cells: Vec<CellTarget>,
}

/// What one step would do against the validated state. It reserves nothing.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditPreview {
    pub step_index: u32,
    pub target: RangeTarget,
    pub would_change: bool,
    pub changed_cell_count: u32,
}

/// What recalculating after an applied batch did beyond the cells it wrote.
#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditCalculation {
    /// Dependents whose value moved.
    pub changed: Vec<CellTarget>,
    pub cycle_cells: Vec<CellTarget>,
    pub limited_cells: Vec<CellTarget>,
    /// Whether a list stopped at 10,000 cells.
    pub truncated: bool,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditValidation {
    pub base_version: DocumentVersion,
    pub would_apply: bool,
    pub previews: Vec<EditPreview>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct EditApplication {
    pub base_version: DocumentVersion,
    pub version: DocumentVersion,
    /// False when every step was a no-op: no history, update or version change.
    pub applied: bool,
    pub source: EditSource,
    pub receipts: Vec<EditReceipt>,
    /// Sheets whose cells the batch or its recalculation changed, as ids in sheet order.
    pub changed_sheets: Vec<String>,
    pub calculation: EditCalculation,
}

pub type ValidationOutcome = std::result::Result<EditValidation, EditRefusal>;
pub type EditOutcome = std::result::Result<EditApplication, EditRefusal>;

/// The JSON wire form of an outcome: the success body or the refusal, tagged with `ok`.
pub fn outcome_json<T: Serialize>(
    outcome: &std::result::Result<T, EditRefusal>,
) -> serde_json::Result<String> {
    let mut value = match outcome {
        Ok(body) => serde_json::to_value(body)?,
        Err(refusal) => serde_json::to_value(refusal)?,
    };
    if let Value::Object(object) = &mut value {
        object.insert("ok".to_owned(), Value::Bool(outcome.is_ok()));
    }
    serde_json::to_string(&value)
}

/// The properties a style patch writes, clears included.
fn style_claims(patch: &StylePatch) -> Vec<StyleProperty> {
    let mut claims = Vec::new();
    for (written, property) in [
        (patch.bold.is_some(), StyleProperty::Bold),
        (patch.italic.is_some(), StyleProperty::Italic),
        (patch.strikethrough.is_some(), StyleProperty::Strikethrough),
        (patch.font_family.is_some(), StyleProperty::FontFamily),
        (patch.font_size.is_some(), StyleProperty::FontSize),
        (patch.text_color.is_some(), StyleProperty::TextColor),
        (patch.fill_color.is_some(), StyleProperty::FillColor),
        (patch.border.is_some(), StyleProperty::Borders),
        (
            patch.horizontal_alignment.is_some(),
            StyleProperty::HorizontalAlignment,
        ),
        (
            patch.vertical_alignment.is_some(),
            StyleProperty::VerticalAlignment,
        ),
        (patch.text_wrapping.is_some(), StyleProperty::TextWrapping),
    ] {
        if written || patch.clear.contains(&property) {
            claims.push(property);
        }
    }
    claims
}

/// What a step writes, for conflict checks.
struct Claims {
    content: bool,
    number_format: bool,
    style: Vec<StyleProperty>,
}

impl Claims {
    fn of(operation: &EditOperation) -> Self {
        match operation {
            EditOperation::SetCellInputs { .. } | EditOperation::SetFormulas { .. } => Self {
                content: true,
                number_format: false,
                style: Vec::new(),
            },
            EditOperation::SetNumberFormat { .. } => Self {
                content: false,
                number_format: true,
                style: Vec::new(),
            },
            EditOperation::PatchStyle { patch, .. } => Self {
                content: false,
                number_format: false,
                style: style_claims(patch),
            },
        }
    }

    fn conflicts(&self, other: &Self) -> bool {
        (self.content && other.content)
            || (self.number_format && other.number_format)
            || self
                .style
                .iter()
                .any(|property| other.style.contains(property))
    }
}

enum Effect {
    /// The cells whose content changes, with their new state.
    Content(Vec<(CellRef, CellState)>),
    NumberFormat(NumberFormatMutation),
    /// `None` for a patch that writes nothing.
    Style(Option<StylePatch>),
}

struct Planned {
    index: u32,
    resolved: Resolved,
    target: RangeTarget,
    claims: Claims,
    effect: Effect,
}

struct Plan {
    base_version: DocumentVersion,
    source: EditSource,
    keys: Vec<String>,
    steps: Vec<Planned>,
    /// Per step, the cells it changes, row-major.
    changed: Vec<Vec<CellRef>>,
    /// `None` when no step changes anything.
    staged: Option<super::staging::PreparedCommit>,
}

fn matrix_fits<T>(matrix: &[Vec<T>], resolved: Resolved) -> bool {
    matrix.len() == resolved.rows() && matrix.iter().all(|row| row.len() == resolved.cols())
}

fn utf16_units(text: &str) -> u64 {
    text.encode_utf16().count() as u64
}

fn validate_formula(source: &str) -> std::result::Result<(), String> {
    if source.is_empty() {
        return Err("is empty".to_owned());
    }
    if source.starts_with('=') {
        return Err("starts with `=`; pass the source without it".to_owned());
    }
    if source.len() > xlsx_calc::lexer::MAX_FORMULA_BYTES {
        return Err("exceeds the formula length limit".to_owned());
    }
    parse_formula(source)
        .map(|_| ())
        .map_err(|error| format!("does not parse: {error}"))
}

fn overlap(left: CellRange, right: CellRange) -> Option<CellRange> {
    let start = CellRef::new(
        left.start.row.max(right.start.row),
        left.start.col.max(right.start.col),
    );
    let end = CellRef::new(
        left.end.row.min(right.end.row),
        left.end.col.min(right.end.col),
    );
    (start.row <= end.row && start.col <= end.col).then_some(CellRange { start, end })
}

/// The first cell of `range` that follows a merge or that an array formula fills; writes to
/// it would split it from its owner.
fn write_lock(sheet: &Sheet, range: CellRange) -> Option<(CellRef, &'static str)> {
    for merged in &sheet.merges {
        let Some(shared) = overlap(*merged, range) else {
            continue;
        };
        let owner = CellRef::new(merged.start.row, merged.start.col);
        let first = CellRef::new(shared.start.row, shared.start.col);
        if first != owner {
            return Some((first, "follows a merged cell"));
        }
        if shared.end.col > shared.start.col {
            return Some((
                CellRef::new(first.row, first.col + 1),
                "follows a merged cell",
            ));
        }
        if shared.end.row > shared.start.row {
            return Some((
                CellRef::new(first.row + 1, first.col),
                "follows a merged cell",
            ));
        }
    }
    sheet.array_formulas().find_map(|(_, filled)| {
        overlap(filled, range).map(|shared| {
            (
                CellRef::new(shared.start.row, shared.start.col),
                "is filled by an array formula",
            )
        })
    })
}

const EMPTY: CellValue = CellValue::Empty;

impl Workbook {
    fn guard_mismatch(
        &self,
        sheet: SheetId,
        at: CellRef,
        guard: &CellGuard,
    ) -> Option<&'static str> {
        let stored = self.model.sheet(sheet).and_then(|sheet| sheet.cell(at));
        if let Some(value) = &guard.value
            && stored.map_or(&EMPTY, |cell| &cell.value) != value
        {
            return Some("value");
        }
        if let Some(formula) = &guard.formula
            && stored.and_then(|cell| cell.formula.as_ref()) != formula.as_ref()
        {
            return Some("formula");
        }
        if let Some(text) = &guard.display_text {
            let current = stored.map_or_else(String::new, |cell| {
                display_text(&self.model.styles, self.model.date_system, cell)
            });
            if current != *text {
                return Some("display text");
            }
        }
        None
    }

    /// Resolves one step against the current state; `planned` are the steps before it.
    fn plan_step(
        &self,
        index: u32,
        step: &EditStep,
        keys: &[String],
        planned: &[Planned],
        visits: &mut u64,
        units: &mut u64,
    ) -> std::result::Result<Planned, EditFailure> {
        let requested = step.operation.target();
        let at = |code, message: String| failure(code, message, Some(requested.clone()));
        let resolved = self.resolve_target(keys, requested)?;
        let target = resolved.target(keys);
        let mut addressed = resolved.cells();
        if step.expect.is_some() {
            addressed *= 2;
        }
        *visits = visits.saturating_add(addressed);
        if *visits > MAX_CELL_VISITS {
            return Err(at(
                EditFailureCode::LimitExceeded,
                format!("a batch addresses at most {MAX_CELL_VISITS} target and guard cells"),
            ));
        }
        if let Some((code, message)) = self.sheet_write_refusal(resolved.sheet) {
            return Err(at(code, message));
        }
        let mut text = step
            .expect
            .iter()
            .flat_map(|guard| guard.cells.iter().flatten())
            .fold(0, |total, guard| {
                total
                    + guard.display_text.as_deref().map_or(0, utf16_units)
                    + guard
                        .formula
                        .as_ref()
                        .and_then(Option::as_deref)
                        .map_or(0, utf16_units)
                    + match &guard.value {
                        Some(CellValue::Text { value }) => utf16_units(value),
                        _ => 0,
                    }
            });
        match &step.operation {
            EditOperation::SetCellInputs { inputs: matrix, .. }
            | EditOperation::SetFormulas {
                formulas: matrix, ..
            } => {
                if !matrix_fits(matrix, resolved) {
                    return Err(at(
                        EditFailureCode::InvalidStep,
                        format!(
                            "the values must be {} rows of {} cells, one per target cell",
                            resolved.rows(),
                            resolved.cols()
                        ),
                    ));
                }
                text += matrix
                    .iter()
                    .flatten()
                    .map(|value| utf16_units(value))
                    .sum::<u64>();
            }
            EditOperation::SetNumberFormat {
                format: NumberFormatMutation::Custom { pattern },
                ..
            } if pattern.is_empty() => {
                return Err(at(
                    EditFailureCode::InvalidStep,
                    "a custom number format needs a pattern".to_owned(),
                ));
            }
            EditOperation::SetNumberFormat { .. } | EditOperation::PatchStyle { .. } => {}
        }
        *units = units.saturating_add(text);
        if *units > MAX_TEXT_UNITS {
            return Err(at(
                EditFailureCode::LimitExceeded,
                format!("a batch carries at most {MAX_TEXT_UNITS} UTF-16 units of text"),
            ));
        }
        if let Some(guard) = &step.expect
            && !matrix_fits(&guard.cells, resolved)
        {
            return Err(at(
                EditFailureCode::InvalidStep,
                format!(
                    "the guard must be {} rows of {} cells, one per target cell",
                    resolved.rows(),
                    resolved.cols()
                ),
            ));
        }
        if let EditOperation::SetFormulas { formulas, .. } = &step.operation {
            for (cell, source) in resolved.iter().zip(formulas.iter().flatten()) {
                validate_formula(source).map_err(|problem| {
                    at(
                        EditFailureCode::InvalidStep,
                        format!("the formula for {} {problem}", cell.to_a1()),
                    )
                })?;
            }
        }
        let claims = Claims::of(&step.operation);
        if let Some(earlier) = planned.iter().find(|earlier| {
            earlier.resolved.sheet == resolved.sheet
                && overlap(earlier.resolved.range, resolved.range).is_some()
                && earlier.claims.conflicts(&claims)
        }) {
            let mut conflict = at(
                EditFailureCode::OverlappingSteps,
                format!(
                    "step {index} writes what step {} already writes in the same cells",
                    earlier.index
                ),
            );
            conflict.conflicting_step_index = Some(earlier.index);
            return Err(conflict);
        }
        let sheet = &self.model.sheets[resolved.sheet.0 as usize];
        if let Some((cell, reason)) = write_lock(sheet, resolved.range) {
            return Err(at(
                EditFailureCode::LockedTarget,
                format!("{} {reason}", cell.to_a1()),
            ));
        }
        if let Some(guard) = &step.expect {
            for (cell, guard) in resolved.iter().zip(guard.cells.iter().flatten()) {
                if let Some(field) = self.guard_mismatch(resolved.sheet, cell, guard) {
                    return Err(at(
                        EditFailureCode::ContentMismatch,
                        format!("the {field} of {} does not match its guard", cell.to_a1()),
                    ));
                }
            }
        }
        let effect = match &step.operation {
            EditOperation::SetCellInputs { inputs, .. } => {
                let mut states = Vec::new();
                for (cell, input) in resolved.iter().zip(inputs.iter().flatten()) {
                    let state = edit_cell_state(&self.model, resolved.sheet, cell, input);
                    validate_cell_state(&state).map_err(|error| {
                        at(
                            EditFailureCode::InvalidStep,
                            format!("the input for {}: {error}", cell.to_a1()),
                        )
                    })?;
                    if !cell_states_semantically_equal(
                        &current_cell_state(&self.model, resolved.sheet, cell),
                        &state,
                    ) {
                        states.push((cell, state));
                    }
                }
                Effect::Content(states)
            }
            EditOperation::SetFormulas { formulas, .. } => {
                let mut states = Vec::new();
                for (cell, source) in resolved.iter().zip(formulas.iter().flatten()) {
                    let current = current_cell_state(&self.model, resolved.sheet, cell);
                    let state = CellState {
                        value: CellValue::Empty,
                        formula: Some(source.clone()),
                        style: current.style,
                    };
                    if !cell_states_semantically_equal(&current, &state) {
                        states.push((cell, state));
                    }
                }
                Effect::Content(states)
            }
            EditOperation::SetNumberFormat { format, .. } => Effect::NumberFormat(format.clone()),
            EditOperation::PatchStyle { patch, .. } => {
                Effect::Style((!claims.style.is_empty()).then(|| patch.clone()))
            }
        };
        Ok(Planned {
            index,
            resolved,
            target,
            claims,
            effect,
        })
    }

    /// Resolves, stages and rehearses a batch without adopting it. `Err` is an internal failure.
    fn plan(&self, request: &EditRequest) -> Result<std::result::Result<Plan, EditRefusal>> {
        let version = self.version();
        let refuse = |failure: EditFailure| Ok(Err(refusal(version.clone(), failure)));
        if request.expect_version != version {
            return refuse(failure(
                EditFailureCode::StaleVersion,
                "the workbook changed since the expected version".to_owned(),
                None,
            ));
        }
        if request.steps.len() > MAX_STEPS {
            return refuse(failure(
                EditFailureCode::LimitExceeded,
                format!("a batch has at most {MAX_STEPS} steps"),
                None,
            ));
        }
        if self.authority.has_pending_updates() {
            return refuse(failure(
                EditFailureCode::Unsupported,
                "the workbook holds updates that are not integrated yet".to_owned(),
                None,
            ));
        }
        let keys = self.sheet_keys();
        let mut visits = 0;
        let mut units = 0;
        let mut steps: Vec<Planned> = Vec::with_capacity(request.steps.len());
        for (index, step) in request.steps.iter().enumerate() {
            let index = index as u32;
            match self.plan_step(index, step, &keys, &steps, &mut visits, &mut units) {
                Ok(planned) => steps.push(planned),
                Err(failure) => return refuse(failure.at(index)),
            }
        }
        let mut ops = Vec::new();
        for planned in &steps {
            if let Effect::Content(states) = &planned.effect {
                ops.extend(states.iter().map(|(at, state)| {
                    (
                        planned.index,
                        Op::SetCell {
                            sheet: planned.resolved.sheet,
                            at: *at,
                            cell: state.clone(),
                        },
                    )
                }));
            }
        }
        for planned in &steps {
            if let Effect::NumberFormat(format) = &planned.effect {
                ops.push((
                    planned.index,
                    Op::SetRangeNumberFormat {
                        sheet: planned.resolved.sheet,
                        range: planned.resolved.range,
                        format: format.clone(),
                    },
                ));
            }
        }
        for planned in &steps {
            if let Effect::Style(Some(patch)) = &planned.effect {
                ops.push((
                    planned.index,
                    Op::PatchRangeStyle {
                        sheet: planned.resolved.sheet,
                        range: planned.resolved.range,
                        patch: patch.clone(),
                    },
                ));
            }
        }
        let unchanged = |steps: Vec<Planned>| {
            Ok(Ok(Plan {
                base_version: version.clone(),
                source: request.source,
                changed: vec![Vec::new(); steps.len()],
                keys: keys.clone(),
                steps,
                staged: None,
            }))
        };
        if ops.is_empty() {
            return unchanged(steps);
        }
        let baseline = self.authority.encode_state_as_update_v1();
        if baseline.len() > MAX_STAGING_BYTES {
            return refuse(failure(
                EditFailureCode::LimitExceeded,
                format!(
                    "the workbook is too large to stage a batch: {} bytes",
                    baseline.len()
                ),
                None,
            ));
        }
        let mut preview = self.model.clone();
        let mut changed: Vec<BTreeSet<(u32, u32)>> = vec![BTreeSet::new(); steps.len()];
        let mut per_op = Vec::with_capacity(ops.len());
        for (step, op) in &ops {
            validate_op(&preview, op)?;
            let formats = preview.styles.cell_xfs.len();
            let inverse = match xlsx_ops::apply_in_place(&mut preview, op) {
                Ok(inverse) => inverse.0,
                Err(error @ (OpError::InvalidStyle(_) | OpError::NumFmtTableFull)) => {
                    let code = match error {
                        OpError::NumFmtTableFull => EditFailureCode::LimitExceeded,
                        _ => EditFailureCode::InvalidStep,
                    };
                    let planned = &steps[*step as usize];
                    return refuse(
                        failure(code, error.to_string(), Some(planned.target.clone())).at(*step),
                    );
                }
                Err(error) => return Err(error.into()),
            };
            if (formats..preview.styles.cell_xfs.len())
                .any(|index| !cell_format_fits(&preview.styles.cell_format(Some(index as u32))))
            {
                let planned = &steps[*step as usize];
                return refuse(
                    failure(
                        EditFailureCode::LimitExceeded,
                        "a resulting cell format exceeds the size a workbook can share".to_owned(),
                        Some(planned.target.clone()),
                    )
                    .at(*step),
                );
            }
            for undone in &inverse {
                if let Op::SetCell { at, .. } = undone {
                    changed[*step as usize].insert((at.row, at.col));
                }
            }
            per_op.push(inverse);
        }
        if changed.iter().all(BTreeSet::is_empty) {
            return unchanged(steps);
        }
        validate_model_sheets(&preview)?;
        let mut inverse = Vec::new();
        for chunk in per_op.into_iter().rev() {
            inverse.extend(chunk);
        }
        let history = match request.history {
            EditHistory::Separate => CommitHistory::Separate,
            EditHistory::None => CommitHistory::None,
        };
        let prepared = match self.prepare_commit(
            ops.into_iter().map(|(_, op)| op).collect(),
            StagedApply::new(preview, inverse),
            SyncOrigin::User,
            history,
            Some(&baseline),
        ) {
            Ok(prepared) => prepared,
            Err(error @ Error::CollaborationDataTooLarge { .. }) => {
                return refuse(failure(
                    EditFailureCode::LimitExceeded,
                    error.to_string(),
                    None,
                ));
            }
            Err(error) => return Err(error),
        };
        if !self
            .authority
            .rehearse_local_update(&baseline, prepared.update())
            .map_err(authority_error)?
        {
            return Err(Error::CollaborativeState(
                "a staged edit batch does not integrate into this replica".to_owned(),
            ));
        }
        Ok(Ok(Plan {
            base_version: version,
            source: request.source,
            keys,
            changed: changed
                .into_iter()
                .map(|cells| {
                    cells
                        .into_iter()
                        .map(|(row, col)| CellRef::new(row, col))
                        .collect()
                })
                .collect(),
            steps,
            staged: Some(prepared),
        }))
    }

    /// Resolves, stages and rehearses a batch like [`Workbook::apply_edits`], then discards it.
    /// Nothing changes. `Err` is an internal failure.
    pub fn validate_edits(&self, request: &EditRequest) -> Result<ValidationOutcome> {
        let plan = match self.plan(request)? {
            Ok(plan) => plan,
            Err(refusal) => return Ok(Err(refusal)),
        };
        Ok(Ok(EditValidation {
            base_version: plan.base_version,
            would_apply: plan.staged.is_some(),
            previews: plan
                .steps
                .iter()
                .zip(&plan.changed)
                .map(|(planned, cells)| EditPreview {
                    step_index: planned.index,
                    target: planned.target.clone(),
                    would_change: !cells.is_empty(),
                    changed_cell_count: cells.len() as u32,
                })
                .collect(),
        }))
    }

    /// Applies every step or none. A policy failure comes back as an [`EditRefusal`] with the
    /// workbook, its history and its proposals untouched; `Err` is an internal failure.
    ///
    /// An applied batch is one committed change, recalculated before the update is published:
    /// one undo step for [`EditHistory::Separate`], none for [`EditHistory::None`].
    pub fn apply_edits(&mut self, request: &EditRequest) -> Result<EditOutcome> {
        let plan = match self.plan(request)? {
            Ok(plan) => plan,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let Plan {
            base_version,
            source,
            keys,
            steps,
            changed,
            staged,
        } = plan;
        let receipts = steps
            .iter()
            .zip(&changed)
            .map(|(planned, cells)| EditReceipt {
                step_index: planned.index,
                changed: !cells.is_empty(),
                target: planned.target.clone(),
                changed_cells: cells
                    .iter()
                    .map(|cell| cell_target(&keys, planned.resolved.sheet, *cell))
                    .collect(),
            })
            .collect();
        let Some(mut prepared) = staged else {
            return Ok(Ok(EditApplication {
                version: base_version.clone(),
                base_version,
                applied: false,
                source,
                receipts,
                changed_sheets: Vec::new(),
                calculation: EditCalculation::default(),
            }));
        };
        let seeds = steps
            .iter()
            .zip(&changed)
            .filter(|(planned, _)| planned.claims.content)
            .flat_map(|(planned, cells)| cells.iter().map(|cell| (planned.resolved.sheet, *cell)))
            .collect::<Vec<_>>();
        let mut graph = DepGraph::build(&prepared.model);
        let recalculated = recalc_after(
            &mut prepared.model,
            &mut graph,
            &seeds,
            request.calculation.now_serial,
        );
        let changed_sheets = steps
            .iter()
            .zip(&changed)
            .filter(|(_, cells)| !cells.is_empty())
            .map(|(planned, _)| planned.resolved.sheet)
            .chain(recalculated.changed.iter().map(|(sheet, _)| *sheet))
            .collect::<BTreeSet<_>>()
            .into_iter()
            .filter_map(|sheet| keys.get(sheet.0 as usize).cloned())
            .collect();
        let application = EditApplication {
            base_version,
            version: DocumentVersion::new(&self.version_nonce, self.committed_changes + 1),
            applied: true,
            source,
            receipts,
            changed_sheets,
            calculation: edit_calculation(&keys, &recalculated, &seeds),
        };
        let bytes = serde_json::to_vec(&application)
            .map_err(|error| Error::InvalidOperation(error.to_string()))?
            .len();
        if bytes > MAX_RESPONSE_BYTES {
            return Ok(Err(refusal(
                application.base_version,
                failure(
                    EditFailureCode::LimitExceeded,
                    format!("the batch's result exceeds {MAX_RESPONSE_BYTES} bytes"),
                    None,
                ),
            )));
        }
        if self.version() != application.base_version {
            return Ok(Err(refusal(
                self.version(),
                failure(
                    EditFailureCode::StaleVersion,
                    "the workbook changed while the batch was staged".to_owned(),
                    None,
                ),
            )));
        }
        let update = self.commit_prepared(prepared)?;
        self.graph = Some(graph);
        self.last_calculation = calculation_result(&recalculated);
        self.publish(update);
        Ok(Ok(application))
    }

    /// [`Workbook::read_cells`] over JSON. An oversized request or response is refused with
    /// `limit-exceeded`; a malformed request is [`Error::InvalidRequest`].
    pub fn read_cells_json(&self, request: &str) -> Result<String> {
        if let Some(refused) = self.oversized_request(request) {
            return Ok(refused);
        }
        self.bounded_json(&self.read_cells(&decode::<ReadRequest>(request)?)?)
    }

    /// [`Workbook::find_text`] over JSON, bounded like [`Workbook::read_cells_json`].
    pub fn find_text_json(&self, request: &str) -> Result<String> {
        if let Some(refused) = self.oversized_request(request) {
            return Ok(refused);
        }
        self.bounded_json(&self.find_text(&decode::<FindRequest>(request)?)?)
    }

    /// [`Workbook::validate_edits`] over JSON, bounded like [`Workbook::read_cells_json`].
    pub fn validate_edits_json(&self, request: &str) -> Result<String> {
        if let Some(refused) = self.oversized_request(request) {
            return Ok(refused);
        }
        self.bounded_json(&self.validate_edits(&decode::<EditRequest>(request)?)?)
    }

    /// [`Workbook::apply_edits`] over JSON; the result is bounded before the batch commits.
    pub fn apply_edits_json(&mut self, request: &str) -> Result<String> {
        if let Some(refused) = self.oversized_request(request) {
            return Ok(refused);
        }
        let outcome = self.apply_edits(&decode::<EditRequest>(request)?)?;
        encode(&outcome)
    }

    fn oversized_request(&self, request: &str) -> Option<String> {
        (request.len() > MAX_REQUEST_BYTES).then(|| {
            self.limit_refusal(format!(
                "a request carries at most {MAX_REQUEST_BYTES} bytes"
            ))
        })
    }

    fn bounded_json<T: Serialize>(
        &self,
        outcome: &std::result::Result<T, EditRefusal>,
    ) -> Result<String> {
        let json = encode(outcome)?;
        if json.len() <= MAX_RESPONSE_BYTES {
            return Ok(json);
        }
        Ok(self.limit_refusal(format!("the result exceeds {MAX_RESPONSE_BYTES} bytes")))
    }

    fn limit_refusal(&self, message: String) -> String {
        let refused: std::result::Result<(), EditRefusal> = Err(refusal(
            self.version(),
            failure(EditFailureCode::LimitExceeded, message, None),
        ));
        encode(&refused).expect("a refusal encodes")
    }
}

fn decode<T: DeserializeOwned>(request: &str) -> Result<T> {
    serde_json::from_str(request).map_err(|error| Error::InvalidRequest(error.to_string()))
}

fn encode<T: Serialize>(outcome: &std::result::Result<T, EditRefusal>) -> Result<String> {
    outcome_json(outcome).map_err(|error| Error::InvalidOperation(error.to_string()))
}

/// Calculation diagnostics beyond the written `seeds`, each list capped.
fn edit_calculation(
    keys: &[String],
    result: &RecalcResult,
    seeds: &[(SheetId, CellRef)],
) -> EditCalculation {
    let written = seeds
        .iter()
        .map(|(sheet, cell)| (sheet.0, cell.row, cell.col))
        .collect::<HashSet<_>>();
    let dependents = result
        .changed
        .iter()
        .filter(|(sheet, cell)| !written.contains(&(sheet.0, cell.row, cell.col)))
        .copied()
        .collect::<Vec<_>>();
    let (changed, changed_cut) = capped_targets(keys, &dependents);
    let (cycle_cells, cycles_cut) = capped_targets(keys, &result.cycle_cells);
    let (limited_cells, limited_cut) = capped_targets(keys, &result.limited_cells);
    EditCalculation {
        changed,
        cycle_cells,
        limited_cells,
        truncated: changed_cut || cycles_cut || limited_cut,
    }
}

fn capped_targets(keys: &[String], cells: &[(SheetId, CellRef)]) -> (Vec<CellTarget>, bool) {
    let targets = cells
        .iter()
        .filter(|(sheet, _)| (sheet.0 as usize) < keys.len())
        .take(MAX_DIAGNOSTIC_CELLS)
        .map(|(sheet, cell)| cell_target(keys, *sheet, *cell))
        .collect();
    (targets, cells.len() > MAX_DIAGNOSTIC_CELLS)
}
