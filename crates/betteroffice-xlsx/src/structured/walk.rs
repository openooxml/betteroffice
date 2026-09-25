//! The one walker behind live and bytes exports. Every record is charged against the
//! budgets before it joins the content, so a stop leaves a complete prefix.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::io;
use std::ops::ControlFlow;
use std::sync::Mutex;

use serde::Serialize;
use sha2::{Digest, Sha256};
use xlsx_model::numfmt::format_is_approximate;
use xlsx_model::{
    Cell, CellRange, CellRef, CellValue, ChartAnchor, DateSystem, Sheet, SheetId,
    Workbook as WorkbookModel,
};
use xlsx_parse::{
    DrawingObjectKind, InspectionBudget, PreservedPackage, SharedStringCells, SheetAxes,
    SheetInventory, SheetVisibility, SourceCellFacts, SourceObject, SourceSheetKind,
};
use xlsx_render::{display_text, format_code_for_cell};

use super::*;
use crate::CalculationResult;
use crate::workbook::target::{RangeAddress, parse_range};

/// Records other than cells: hidden spans, merges, tables, hyperlinks, objects, names.
const MAX_METADATA_RECORDS: usize = 100_000;
const MAX_DIAGNOSTICS: usize = 10_000;
/// Stored cells, map seeks and records one export may visit, exported or not.
const MAX_VISITED: u64 = 20_000_000;
/// XML nodes and part bytes one export may read from the retained package.
const MAX_INSPECTED_NODES: u64 = 20_000_000;
const MAX_INSPECTED_BYTES: u64 = 256 * 1024 * 1024;
/// The smallest `maxBytes` accepted, whatever the workbook.
const MIN_EXPORT_BYTES: u32 = 1_024;

/// What an export reads: the current model with the package it was opened from.
pub(crate) struct ExportSource<'a> {
    pub(crate) model: &'a WorkbookModel,
    pub(crate) package: Option<&'a PreservedPackage>,
    /// The source sheet each current sheet came from.
    pub(crate) origins: &'a [Option<usize>],
    pub(crate) shared_string_cells: &'a [SharedStringCells],
    pub(crate) axes: &'a [Option<SheetAxes>],
    pub(crate) calculation: &'a CalculationResult,
    /// Whether each sheet was added in this session.
    pub(crate) created: &'a [bool],
    /// Whether anything was calculated or edited since the package was read.
    pub(crate) edited: bool,
    /// SHA-256 of retained parts, which never change.
    pub(crate) part_hashes: &'a Mutex<BTreeMap<String, String>>,
}

impl ExportSource<'_> {
    fn identity(&self, index: usize) -> XlsxSheetIdentity {
        XlsxSheetIdentity {
            index: index as u32,
            name: self.model.sheets[index].name.clone(),
        }
    }

    fn origin(&self, index: usize) -> Option<(usize, &PreservedPackage)> {
        let origin = self.origins.get(index).copied().flatten()?;
        Some((origin, self.package?))
    }

    fn facts(&self, index: usize) -> Option<&SourceCellFacts> {
        let (origin, package) = self.origin(index)?;
        package.source_cell_facts(origin)
    }

    /// The source cell current cell `at` of sheet `index` was read from; `Err` when the
    /// sheet's row and column mappings are gone.
    fn source_cell(
        &self,
        index: usize,
        at: CellRef,
    ) -> std::result::Result<Option<(u32, u32)>, ()> {
        let axes = self.axes.get(index).and_then(Option::as_ref).ok_or(())?;
        Ok(axes.rows.source(at.row).zip(axes.cols.source(at.col)))
    }

    fn part_hash(&self, part: &str) -> Option<String> {
        let bytes = self.package?.part_bytes(part)?;
        let mut hashes = self
            .part_hashes
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        Some(
            hashes
                .entry(part.to_owned())
                .or_insert_with(|| hex(&Sha256::digest(bytes)))
                .clone(),
        )
    }

    fn source_part(&self, part: &str, path: Vec<u32>) -> Option<XlsxSourcePart> {
        Some(XlsxSourcePart {
            part: part.to_owned(),
            part_sha256: self.part_hash(part)?,
            path,
        })
    }
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02x}")).collect()
}

struct Limits {
    max_cells: u32,
    max_bytes: usize,
}

fn refused(
    code: XlsxExportFailureCode,
    target: Option<XlsxAnchor>,
    message: String,
) -> XlsxExportFailure {
    XlsxExportFailure {
        code,
        target,
        message,
    }
}

fn limits(options: &XlsxExportOptions) -> std::result::Result<Limits, XlsxExportFailure> {
    let max_cells = options.max_cells.unwrap_or(DEFAULT_EXPORT_MAX_CELLS);
    if max_cells == 0 {
        return Err(refused(
            XlsxExportFailureCode::InvalidOptions,
            None,
            "maxCells must be positive".to_owned(),
        ));
    }
    if max_cells > MAX_EXPORT_CELLS {
        return Err(refused(
            XlsxExportFailureCode::LimitExceeded,
            None,
            format!("maxCells is at most {MAX_EXPORT_CELLS}"),
        ));
    }
    let max_bytes = options.max_bytes.unwrap_or(DEFAULT_EXPORT_MAX_BYTES);
    if max_bytes > MAX_EXPORT_BYTES {
        return Err(refused(
            XlsxExportFailureCode::LimitExceeded,
            None,
            format!("maxBytes is at most {MAX_EXPORT_BYTES}"),
        ));
    }
    if max_bytes < MIN_EXPORT_BYTES {
        return Err(refused(
            XlsxExportFailureCode::LimitExceeded,
            None,
            format!("maxBytes must be at least {MIN_EXPORT_BYTES}"),
        ));
    }
    Ok(Limits {
        max_cells,
        max_bytes: max_bytes as usize,
    })
}

/// The sheets to walk in sheet order, each with its requested range.
fn selections(
    source: &ExportSource<'_>,
    options: &XlsxExportOptions,
) -> std::result::Result<Vec<(usize, Option<CellRange>)>, XlsxExportFailure> {
    let count = source.model.sheets.len();
    let Some(scope) = &options.scope else {
        return Ok((0..count).map(|index| (index, None)).collect());
    };
    if scope.is_empty() {
        return Err(refused(
            XlsxExportFailureCode::InvalidScope,
            None,
            "scope must name at least one sheet".to_owned(),
        ));
    }
    let mut selected = BTreeMap::new();
    for entry in scope {
        let index = entry.sheet as usize;
        if index >= count {
            return Err(refused(
                XlsxExportFailureCode::InvalidScope,
                None,
                format!("sheet {index} does not exist; the workbook has {count} sheets"),
            ));
        }
        let anchor = XlsxAnchor::Sheet {
            sheet: source.identity(index),
        };
        let range = match &entry.range {
            Some(a1) => Some(parse_range(&RangeAddress::A1 { a1: a1.clone() }).map_err(
                |message| {
                    refused(
                        XlsxExportFailureCode::InvalidScope,
                        Some(anchor.clone()),
                        message,
                    )
                },
            )?),
            None => None,
        };
        if selected.insert(index, range).is_some() {
            return Err(refused(
                XlsxExportFailureCode::InvalidScope,
                Some(anchor),
                format!("sheet {index} appears more than once in scope"),
            ));
        }
    }
    Ok(selected.into_iter().collect())
}

/// Work budgets beside the ones options set.
#[derive(Clone, Copy)]
pub(super) struct Budgets {
    pub(super) visited: u64,
    pub(super) inspection: InspectionBudget,
}

impl Default for Budgets {
    fn default() -> Self {
        Self {
            visited: MAX_VISITED,
            inspection: InspectionBudget {
                nodes: MAX_INSPECTED_NODES,
                bytes: MAX_INSPECTED_BYTES,
            },
        }
    }
}

/// Exports `source`, or refuses options it cannot honor.
pub(super) fn export(
    source: &ExportSource<'_>,
    options: &XlsxExportOptions,
    anchor_scope: XlsxAnchorScope,
) -> std::result::Result<XlsxStructuredContent, XlsxExportFailure> {
    export_within(source, options, anchor_scope, Budgets::default())
}

pub(super) fn export_within(
    source: &ExportSource<'_>,
    options: &XlsxExportOptions,
    anchor_scope: XlsxAnchorScope,
    budgets: Budgets,
) -> std::result::Result<XlsxStructuredContent, XlsxExportFailure> {
    let limits = limits(options)?;
    let selections = selections(source, options)?;
    let included = XlsxExportIncluded {
        hidden_sheets: options.include_hidden_sheets.unwrap_or(false),
        hidden_rows: options.include_hidden_rows.unwrap_or(false),
        hidden_columns: options.include_hidden_columns.unwrap_or(false),
        defined_names: options.include_defined_names.unwrap_or(true),
        hidden_names: options.include_hidden_names.unwrap_or(false),
    };
    let content = XlsxStructuredContent {
        schema_version: XlsxSchemaVersion,
        anchor_scope,
        date_system: match source.model.date_system {
            DateSystem::V1900 => XlsxDateSystem::V1900,
            DateSystem::V1904 => XlsxDateSystem::V1904,
        },
        calculation: XlsxCalculationState::default(),
        included,
        sheets: Vec::new(),
        defined_names: Vec::new(),
        diagnostics: Vec::new(),
        truncated: false,
    };
    let reserve = (0..source.model.sheets.len())
        .map(|index| {
            Some(XlsxAnchor::Sheet {
                sheet: source.identity(index),
            })
        })
        .chain([None])
        .flat_map(|anchor| {
            Stop::ALL.map(|stop| {
                measure(&truncation(
                    anchor.clone(),
                    stop,
                    Some(CellRef::new(
                        xlsx_model::MAX_ROWS - 1,
                        xlsx_model::MAX_COLS - 1,
                    )),
                )) + 1
            })
        })
        .max()
        .unwrap_or(0);
    let floor = measure(&content) + reserve;
    if limits.max_bytes < floor {
        return Err(refused(
            XlsxExportFailureCode::LimitExceeded,
            None,
            format!("maxBytes must be at least {floor} for this workbook"),
        ));
    }
    let mut walk = Walk {
        source,
        used: floor - reserve,
        max_bytes: limits.max_bytes - reserve,
        cells: 0,
        max_cells: limits.max_cells,
        metadata: 0,
        visited: 0,
        max_visited: budgets.visited,
        inspection: budgets.inspection,
        content,
        failures: failure_cells(source.calculation),
    };
    let finished = walk.defined_names().and_then(|()| {
        selections
            .into_iter()
            .try_for_each(|(index, range)| walk.sheet(index, range))
    });
    if let Err(stopped) = finished {
        walk.content.truncated = true;
        if let Some(sheet) = walk.content.sheets.last_mut()
            && stopped.sheet == Some(sheet.anchor.clone())
        {
            sheet.truncated = true;
        }
        walk.content
            .diagnostics
            .push(truncation(stopped.sheet, stopped.limit, stopped.after));
    }
    debug_assert!(measure(&walk.content) <= limits.max_bytes);
    Ok(walk.content)
}

#[derive(Clone, Copy)]
enum Stop {
    Cells,
    Bytes,
    Metadata,
    Diagnostics,
    Visited,
    Inspection,
}

impl Stop {
    const ALL: [Stop; 6] = [
        Stop::Cells,
        Stop::Bytes,
        Stop::Metadata,
        Stop::Diagnostics,
        Stop::Visited,
        Stop::Inspection,
    ];
}

/// Where and why a walk stopped.
struct Stopped {
    sheet: Option<XlsxAnchor>,
    limit: Stop,
    after: Option<CellRef>,
}

fn truncation(
    anchor: Option<XlsxAnchor>,
    limit: Stop,
    after: Option<CellRef>,
) -> XlsxExportDiagnostic {
    let limit = match limit {
        Stop::Cells => "maxCells",
        Stop::Bytes => "maxBytes",
        Stop::Metadata => "metadata record",
        Stop::Diagnostics => "diagnostic",
        Stop::Visited => "visited-cell",
        Stop::Inspection => "source-inspection",
    };
    let position = match (&anchor, after) {
        (Some(_), Some(cell)) => format!("after cell {}", cell.to_a1()),
        (Some(_), None) => "at this sheet".to_owned(),
        (None, _) => "before the first sheet".to_owned(),
    };
    XlsxExportDiagnostic {
        code: XlsxExportDiagnosticCode::Truncated,
        severity: XlsxExportSeverity::Warning,
        anchor,
        message: format!(
            "Export stopped at the {limit} limit {position}; nothing after that point, \
             including later sheets, is covered."
        ),
    }
}

fn failure_cells(calculation: &CalculationResult) -> [HashSet<(u32, u32, u32)>; 2] {
    let keys = |cells: &[crate::CellAddress]| {
        cells
            .iter()
            .map(|address| (address.sheet.0, address.cell.row, address.cell.col))
            .collect()
    };
    [
        keys(&calculation.cycle_cells),
        keys(&calculation.limited_cells),
    ]
}

struct Walk<'a> {
    source: &'a ExportSource<'a>,
    content: XlsxStructuredContent,
    used: usize,
    max_bytes: usize,
    cells: u32,
    max_cells: u32,
    metadata: usize,
    visited: u64,
    max_visited: u64,
    /// What reading the retained package may still spend.
    inspection: InspectionBudget,
    /// Cycle cells, then limited cells, of the last calculation.
    failures: [HashSet<(u32, u32, u32)>; 2],
}

/// Diagnostics each sheet raises once, at the first cell that needs it.
#[derive(Clone, Copy, Default)]
struct Raised {
    formula: bool,
    missing: bool,
    failure: bool,
    rich: bool,
    approximate: bool,
    unavailable: bool,
}

/// Per-sheet object bookkeeping: diagnostics raised and records admitted.
#[derive(Default)]
struct ObjectNotes {
    placeholders: bool,
    unresolved: bool,
    emitted: usize,
}

struct Counter(usize);

impl io::Write for Counter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.0 += bytes.len();
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Compact JSON bytes of `value`.
fn measure(value: &impl Serialize) -> usize {
    let mut counter = Counter(0);
    serde_json::to_writer(&mut counter, value).expect("export records serialize");
    counter.0
}

fn a1(cell: CellRef) -> String {
    CellRef::new(cell.row, cell.col).to_a1()
}

fn a1_range(range: CellRange) -> String {
    CellRange::new(
        CellRef::new(range.start.row, range.start.col),
        CellRef::new(range.end.row, range.end.col),
    )
    .to_a1()
}

fn within(outer: CellRange, inner: CellRange) -> bool {
    outer.contains(inner.start) && outer.contains(inner.end)
}

impl Walk<'_> {
    fn stop(&self, sheet: Option<&XlsxAnchor>, limit: Stop, after: Option<CellRef>) -> Stopped {
        Stopped {
            sheet: sheet.cloned(),
            limit,
            after,
        }
    }

    /// Charges `record` plus its separator, or reports that it does not fit.
    fn charge(&mut self, record: &impl Serialize) -> bool {
        let cost = measure(record) + 1;
        if self.used + cost > self.max_bytes {
            return false;
        }
        self.used += cost;
        true
    }

    fn metadata(
        &mut self,
        record: &impl Serialize,
        sheet: Option<&XlsxAnchor>,
    ) -> std::result::Result<(), Stopped> {
        if self.metadata == MAX_METADATA_RECORDS {
            return Err(self.stop(sheet, Stop::Metadata, None));
        }
        if !self.charge(record) {
            return Err(self.stop(sheet, Stop::Bytes, None));
        }
        self.metadata += 1;
        Ok(())
    }

    fn diagnose(
        &mut self,
        code: XlsxExportDiagnosticCode,
        severity: XlsxExportSeverity,
        anchor: Option<XlsxAnchor>,
        message: String,
        sheet: Option<&XlsxAnchor>,
    ) -> std::result::Result<(), Stopped> {
        let diagnostic = XlsxExportDiagnostic {
            code,
            severity,
            anchor,
            message,
        };
        if self.content.diagnostics.len() == MAX_DIAGNOSTICS {
            return Err(self.stop(sheet, Stop::Diagnostics, None));
        }
        if !self.charge(&diagnostic) {
            return Err(self.stop(sheet, Stop::Bytes, None));
        }
        self.content.diagnostics.push(diagnostic);
        Ok(())
    }

    fn info(
        &mut self,
        code: XlsxExportDiagnosticCode,
        anchor: XlsxAnchor,
        message: String,
    ) -> std::result::Result<(), Stopped> {
        let sheet = sheet_of(&anchor);
        self.diagnose(
            code,
            XlsxExportSeverity::Info,
            Some(anchor),
            message,
            sheet.as_ref(),
        )
    }

    fn warn(
        &mut self,
        code: XlsxExportDiagnosticCode,
        anchor: XlsxAnchor,
        message: String,
    ) -> std::result::Result<(), Stopped> {
        let sheet = sheet_of(&anchor);
        self.diagnose(
            code,
            XlsxExportSeverity::Warning,
            Some(anchor),
            message,
            sheet.as_ref(),
        )
    }

    fn defined_names(&mut self) -> std::result::Result<(), Stopped> {
        let model = self.source.model;
        if model.defined_names.is_empty() {
            return Ok(());
        }
        if !self.content.included.defined_names {
            return self.diagnose(
                XlsxExportDiagnosticCode::HiddenContentExcluded,
                XlsxExportSeverity::Info,
                None,
                format!(
                    "Defined names are excluded ({}); set includeDefinedNames to list them.",
                    model.defined_names.len()
                ),
                None,
            );
        }
        let (mut hidden, mut on_hidden_sheets) = (0, 0);
        for (ordinal, defined) in model.defined_names.iter().enumerate() {
            let local_sheet = defined
                .local_sheet
                .map(|sheet| sheet.0 as usize)
                .filter(|&index| index < model.sheets.len());
            if defined.hidden && !self.content.included.hidden_names {
                hidden += 1;
                continue;
            }
            if local_sheet.is_some_and(|index| !self.sheet_visible_enough(index)) {
                on_hidden_sheets += 1;
                continue;
            }
            let local_sheet = local_sheet.map(|index| self.source.identity(index));
            let record = XlsxExportDefinedName {
                id: format!("n{ordinal}"),
                anchor: XlsxAnchor::DefinedName {
                    name: defined.name.clone(),
                    local_sheet: local_sheet.clone(),
                    ordinal: ordinal as u32,
                },
                name: defined.name.clone(),
                formula: defined.formula.clone(),
                hidden: defined.hidden,
                local_sheet,
            };
            self.metadata(&record, None)?;
            self.content.defined_names.push(record);
        }
        if hidden > 0 {
            self.diagnose(
                XlsxExportDiagnosticCode::HiddenContentExcluded,
                XlsxExportSeverity::Info,
                None,
                format!(
                    "Hidden defined names are excluded ({hidden}); set includeHiddenNames to list them."
                ),
                None,
            )?;
        }
        if on_hidden_sheets > 0 {
            self.diagnose(
                XlsxExportDiagnosticCode::HiddenContentExcluded,
                XlsxExportSeverity::Info,
                None,
                format!(
                    "Defined names scoped to excluded hidden sheets are excluded ({on_hidden_sheets})."
                ),
                None,
            )?;
        }
        Ok(())
    }

    /// A sheet read from a package has the visibility its entry declares, and one added in
    /// this session is visible. Any other sheet, from a model handed in or a replica's
    /// peers, has unknown visibility.
    fn visibility(&self, index: usize) -> XlsxSheetVisibility {
        let source = self.source;
        let Some((origin, package)) = source.origin(index) else {
            return if source.created.get(index).copied().unwrap_or(false) {
                XlsxSheetVisibility::Visible
            } else {
                XlsxSheetVisibility::Unknown
            };
        };
        match package.source_sheet_visibility(origin) {
            Some(SheetVisibility::Visible) => XlsxSheetVisibility::Visible,
            Some(SheetVisibility::Hidden) => XlsxSheetVisibility::Hidden,
            Some(SheetVisibility::VeryHidden) => XlsxSheetVisibility::VeryHidden,
            None | Some(SheetVisibility::Unknown) => XlsxSheetVisibility::Unknown,
        }
    }

    fn sheet_visible_enough(&self, index: usize) -> bool {
        self.content.included.hidden_sheets
            || self.visibility(index) == XlsxSheetVisibility::Visible
    }

    fn sheet(
        &mut self,
        index: usize,
        requested: Option<CellRange>,
    ) -> std::result::Result<(), Stopped> {
        let source = self.source;
        let sheet = &source.model.sheets[index];
        let anchor = XlsxAnchor::Sheet {
            sheet: source.identity(index),
        };
        let visibility = self.visibility(index);
        if !self.sheet_visible_enough(index) {
            return match visibility {
                XlsxSheetVisibility::Unknown => self.warn(
                    XlsxExportDiagnosticCode::VisibilityUnknown,
                    anchor,
                    "The sheet's visibility is unavailable or not a known state, so it is excluded as hidden; set includeHiddenSheets to export it.".to_owned(),
                ),
                _ => self.info(
                    XlsxExportDiagnosticCode::HiddenContentExcluded,
                    anchor,
                    "The sheet is hidden and excluded; set includeHiddenSheets to export it.".to_owned(),
                ),
            };
        }
        let origin = source.origin(index);
        let kind = match origin.and_then(|(origin, package)| package.source_sheet_kind(origin)) {
            None | Some(SourceSheetKind::Worksheet) => XlsxSheetKind::Worksheet,
            Some(SourceSheetKind::Chartsheet) => XlsxSheetKind::Chartsheet,
            Some(SourceSheetKind::Dialogsheet) => XlsxSheetKind::Dialogsheet,
            Some(SourceSheetKind::Macrosheet) => XlsxSheetKind::Macrosheet,
            Some(SourceSheetKind::Other) => XlsxSheetKind::Other,
        };
        let used = sheet.used_range();
        let selected = requested.or(used);
        let record = XlsxExportSheet {
            id: format!("s{index}"),
            anchor: anchor.clone(),
            kind,
            visibility,
            used_range: used.map(a1_range),
            selected_range: selected.map(a1_range),
            hidden_rows: Vec::new(),
            hidden_columns: Vec::new(),
            merges: Vec::new(),
            tables: Vec::new(),
            hyperlinks: Vec::new(),
            objects: Vec::new(),
            cells: Vec::new(),
            source: origin.and_then(|(origin, package)| {
                source.source_part(package.source_sheet_part(origin)?, Vec::new())
            }),
            truncated: false,
        };
        if !self.charge(&record) {
            return Err(self.stop(Some(&anchor), Stop::Bytes, None));
        }
        self.content.sheets.push(record);
        let at = Some(&anchor);
        if kind != XlsxSheetKind::Worksheet {
            self.warn(
                XlsxExportDiagnosticCode::UnsupportedSheet,
                anchor.clone(),
                "This is not a worksheet; only its drawing objects are exported.".to_owned(),
            )?;
        }
        let inventory = match origin {
            Some((origin, package)) => {
                let mut budget = self.inspection;
                let inventory = package.source_sheet_inventory(origin, &mut budget);
                self.inspection = budget;
                inventory
            }
            None => SheetInventory::default(),
        };
        self.inventory_diagnostics(&anchor, &inventory)?;
        if let Some(part) = &inventory.limited {
            return Err(self.limited(&anchor, part));
        }
        let in_scope = |range: CellRange| selected.is_none_or(|selected| selected.overlaps(&range));
        let clipped = |range: CellRange| selected.is_some_and(|selected| !within(selected, range));

        if let Some(range) = selected {
            let letters = xlsx_model::addr::col_to_letters;
            let rows = sheet
                .hidden_row_spans(range.start.row..=range.end.row)
                .map(|(first, last)| format!("{}:{}", first + 1, last + 1));
            let columns = sheet
                .hidden_col_spans(range.start.col..=range.end.col)
                .map(|(first, last)| format!("{}:{}", letters(first), letters(last)));
            self.hidden_spans(rows, true, &anchor)?;
            self.hidden_spans(columns, false, &anchor)?;
        }

        let identity = source.identity(index);
        let range_anchor = |range: CellRange| XlsxAnchor::Range {
            sheet: identity.clone(),
            a1: a1_range(range),
        };
        let mut merges = Vec::new();
        for &range in &sheet.merges {
            self.visit(at, 1)?;
            if !in_scope(range) {
                continue;
            }
            let record = XlsxExportMerge {
                id: format!("s{index}:m{}", merges.len()),
                anchor: range_anchor(range),
                clipped: clipped(range),
            };
            self.metadata(&record, at)?;
            self.current().merges.push(record);
            merges.push(range);
        }
        let mut ordinal = 0;
        for table in &source.model.tables {
            self.visit(at, 1)?;
            if table.sheet != SheetId(index as u32) || !in_scope(table.range) {
                continue;
            }
            let record = XlsxExportTable {
                id: format!("s{index}:t{ordinal}"),
                anchor: range_anchor(table.range),
                name: table.name.clone(),
                header_rows: table.header_rows,
                totals_rows: table.totals_rows,
                columns: table.columns.clone(),
                clipped: clipped(table.range),
            };
            self.metadata(&record, at)?;
            self.current().tables.push(record);
            ordinal += 1;
        }
        let mut ordinal = 0;
        for link in &sheet.hyperlinks {
            self.visit(at, 1)?;
            if !in_scope(link.range) {
                continue;
            }
            let record = XlsxExportHyperlink {
                id: format!("s{index}:h{ordinal}"),
                anchor: range_anchor(link.range),
                external_target: link.external_target.clone(),
                location: link.location.clone(),
                tooltip: link.tooltip.clone(),
                display: link.display.clone(),
                clipped: clipped(link.range),
            };
            self.metadata(&record, at)?;
            self.current().hyperlinks.push(record);
            ordinal += 1;
        }
        self.objects(index, sheet, requested)?;

        if kind == XlsxSheetKind::Worksheet
            && let Some(range) = selected
        {
            self.cells(index, sheet, range, merges, &anchor)?;
        }
        Ok(())
    }

    /// Lists hidden spans as they are found, noting an exclusion before the first.
    fn hidden_spans(
        &mut self,
        spans: impl Iterator<Item = String>,
        rows: bool,
        anchor: &XlsxAnchor,
    ) -> std::result::Result<(), Stopped> {
        let (include, what, option) = if rows {
            (
                self.content.included.hidden_rows,
                "rows",
                "includeHiddenRows",
            )
        } else {
            (
                self.content.included.hidden_columns,
                "columns",
                "includeHiddenColumns",
            )
        };
        for (ordinal, span) in spans.enumerate() {
            if ordinal == 0 && !include {
                self.info(
                    XlsxExportDiagnosticCode::HiddenContentExcluded,
                    anchor.clone(),
                    format!("Cells in hidden {what} are excluded; set {option} to export them."),
                )?;
            }
            self.metadata(&span, Some(anchor))?;
            let sheet = self.current();
            if rows {
                sheet.hidden_rows.push(span);
            } else {
                sheet.hidden_columns.push(span);
            }
        }
        Ok(())
    }

    fn current(&mut self) -> &mut XlsxExportSheet {
        self.content
            .sheets
            .last_mut()
            .expect("a sheet is being exported")
    }

    fn inventory_diagnostics(
        &mut self,
        anchor: &XlsxAnchor,
        inventory: &SheetInventory,
    ) -> std::result::Result<(), Stopped> {
        let comments = inventory.comments + inventory.threaded_comments;
        if comments > 0 {
            self.info(
                XlsxExportDiagnosticCode::CommentsOmitted,
                anchor.clone(),
                format!("Comments on this sheet are not exported ({comments})."),
            )?;
        }
        if !inventory.features.is_empty() {
            self.warn(
                XlsxExportDiagnosticCode::UnsupportedContent,
                anchor.clone(),
                format!(
                    "This sheet carries content the export does not represent: {}.",
                    inventory.features.join(", ")
                ),
            )?;
        }
        for part in &inventory.unreadable_parts {
            self.unreadable(anchor, part)?;
        }
        Ok(())
    }

    /// Stops at a part too large to read within the parser caps or the inspection budget,
    /// naming it first when that still fits.
    fn limited(&mut self, anchor: &XlsxAnchor, part: &str) -> Stopped {
        let named = self.warn(
            XlsxExportDiagnosticCode::UnreadableObject,
            anchor.clone(),
            format!(
                "The part {part} exceeds what the export may read; it and everything after it are not exported."
            ),
        );
        match named {
            Ok(()) => self.stop(Some(anchor), Stop::Inspection, None),
            Err(stopped) => stopped,
        }
    }

    fn unreadable(&mut self, anchor: &XlsxAnchor, part: &str) -> std::result::Result<(), Stopped> {
        self.warn(
            XlsxExportDiagnosticCode::UnreadableObject,
            anchor.clone(),
            format!("The part {part} could not be read; what it holds is not exported."),
        )
    }

    /// Counts `steps` against the visited budget.
    fn visit(
        &mut self,
        sheet: Option<&XlsxAnchor>,
        steps: u64,
    ) -> std::result::Result<(), Stopped> {
        self.visited += steps;
        if self.visited > self.max_visited {
            return Err(self.stop(sheet, Stop::Visited, None));
        }
        Ok(())
    }

    /// Drawing objects as the source inventory lists them, one at a time, then modelled
    /// charts it does not list.
    fn objects(
        &mut self,
        index: usize,
        sheet: &Sheet,
        requested: Option<CellRange>,
    ) -> std::result::Result<(), Stopped> {
        let source = self.source;
        let charts = sheet
            .charts
            .iter()
            .enumerate()
            .map(|(position, chart)| ((chart.drawing.as_str(), chart.anchor_index), position))
            .collect::<HashMap<_, _>>();
        let mut modelled = vec![false; sheet.charts.len()];
        let mut notes = ObjectNotes::default();
        if let Some((origin, package)) = source.origin(index) {
            let mut outcome = Ok(());
            let mut budget = self.inspection;
            let _ = package.visit_source_sheet_objects(origin, &mut budget, |item| {
                outcome = self.source_object(
                    index,
                    sheet,
                    item,
                    &charts,
                    &mut modelled,
                    requested,
                    &mut notes,
                );
                if outcome.is_ok() {
                    ControlFlow::Continue(())
                } else {
                    ControlFlow::Break(())
                }
            });
            self.inspection = budget;
            outcome?;
        }
        let identity = source.identity(index);
        for (chart, _) in sheet
            .charts
            .iter()
            .zip(&modelled)
            .filter(|(_, seen)| !**seen)
        {
            let record = XlsxExportObject {
                id: String::new(),
                kind: XlsxObjectKind::Chart,
                anchor: grid_anchor(&identity, &chart.anchor),
                name: None,
                alt_text: None,
                title: None,
                hidden: false,
                part: Some(chart.part.clone()),
                source: None,
            };
            self.object(index, record, requested, &mut notes)?;
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn source_object(
        &mut self,
        index: usize,
        sheet: &Sheet,
        item: SourceObject,
        charts: &HashMap<(&str, usize), usize>,
        modelled: &mut [bool],
        requested: Option<CellRange>,
        notes: &mut ObjectNotes,
    ) -> std::result::Result<(), Stopped> {
        let source = self.source;
        let identity = source.identity(index);
        let anchor = XlsxAnchor::Sheet {
            sheet: identity.clone(),
        };
        self.visit(Some(&anchor), 1)?;
        let object = match item {
            SourceObject::Object(object) => object,
            SourceObject::Unreadable(part) => return self.unreadable(&anchor, &part),
            SourceObject::Limited(part) => return Err(self.limited(&anchor, &part)),
        };
        let chart = object
            .anchor_index
            .and_then(|anchor_index| charts.get(&(object.drawing.as_str(), anchor_index)))
            .copied();
        let provenance = source.source_part(&object.drawing, vec![object.ordinal]);
        let (placed, part) = match chart {
            Some(position) => {
                modelled[position] = true;
                let chart = &sheet.charts[position];
                (
                    grid_anchor(&identity, &chart.anchor),
                    Some(chart.part.clone()),
                )
            }
            None => {
                let axes = source.axes.get(index).and_then(Option::as_ref);
                let placed = match object.anchor.as_ref() {
                    Some(authored) if grid_anchor(&identity, authored).is_some() => {
                        match axes.filter(|axes| unmoved(axes, authored)) {
                            Some(_) => grid_anchor(&identity, authored),
                            None => {
                                if !notes.unresolved {
                                    notes.unresolved = true;
                                    self.info(
                                        XlsxExportDiagnosticCode::ProvenanceUnavailable,
                                        anchor.clone(),
                                        "Rows or columns under some drawing objects of this sheet moved since it was read, and a save keeps those objects where they were, so they have no grid anchor.".to_owned(),
                                    )?;
                                }
                                None
                            }
                        }
                    }
                    _ => None,
                };
                (placed, object.target.clone())
            }
        };
        if object.kind == DrawingObjectKind::Chart && chart.is_none() {
            let located = provenance.clone().map(|source| XlsxAnchor::SourcePart {
                part: source.part,
                part_sha256: source.part_sha256,
                path: source.path,
            });
            self.diagnose(
                XlsxExportDiagnosticCode::UnreadableObject,
                XlsxExportSeverity::Warning,
                located.or_else(|| Some(anchor.clone())),
                format!(
                    "A chart in {} could not be read; it is exported as a placeholder only.",
                    object.drawing
                ),
                Some(&anchor),
            )?;
        }
        let record = XlsxExportObject {
            id: String::new(),
            kind: object_kind(object.kind),
            anchor: placed,
            name: object.name,
            alt_text: object.description,
            title: object.title,
            hidden: object.hidden,
            part,
            source: provenance,
        };
        self.object(index, record, requested, notes)
    }

    /// Admits one object record unless the requested range excludes it.
    fn object(
        &mut self,
        index: usize,
        record: XlsxExportObject,
        requested: Option<CellRange>,
        notes: &mut ObjectNotes,
    ) -> std::result::Result<(), Stopped> {
        let outside = match (&record.anchor, requested) {
            (Some(XlsxAnchor::Range { a1, .. } | XlsxAnchor::Cell { a1, .. }), Some(range)) => {
                CellRange::parse_a1(a1).is_ok_and(|placed| !placed.overlaps(&range))
            }
            _ => false,
        };
        if outside {
            return Ok(());
        }
        let anchor = XlsxAnchor::Sheet {
            sheet: self.source.identity(index),
        };
        if !notes.placeholders {
            notes.placeholders = true;
            self.info(
                XlsxExportDiagnosticCode::ObjectPlaceholder,
                anchor.clone(),
                "Drawing objects on this sheet are exported as placeholders without chart data, image bytes or shape text.".to_owned(),
            )?;
        }
        let record = XlsxExportObject {
            id: format!("s{index}:o{}", notes.emitted),
            ..record
        };
        self.metadata(&record, Some(&anchor))?;
        notes.emitted += 1;
        self.current().objects.push(record);
        Ok(())
    }

    fn cells(
        &mut self,
        index: usize,
        sheet: &Sheet,
        range: CellRange,
        mut merges: Vec<CellRange>,
        anchor: &XlsxAnchor,
    ) -> std::result::Result<(), Stopped> {
        let identity = self.source.identity(index);
        let include_rows = self.content.included.hidden_rows;
        let include_columns = self.content.included.hidden_columns;
        merges.sort_by_key(|range| (range.start.row, range.start.col));
        let mut merge_index = MergeSweep::new(merges);
        let mut raised = Raised::default();
        let mut last = None;
        let mut cursor = sheet.cells_in_range(range);
        let mut previous_seeks = 0;
        while let Some((at, cell)) = cursor.next() {
            self.visited += 1 + cursor.seeks() - previous_seeks;
            previous_seeks = cursor.seeks();
            if self.visited > self.max_visited {
                return Err(self.stop(Some(anchor), Stop::Visited, last));
            }
            if !include_rows && sheet.row_hidden(at.row) {
                cursor.skip_row(at.row);
                continue;
            }
            if !include_columns && sheet.col_hidden(at.col) {
                continue;
            }
            if self.cells == self.max_cells {
                return Err(self.stop(Some(anchor), Stop::Cells, last));
            }
            let cell_anchor = XlsxAnchor::Cell {
                sheet: identity.clone(),
                a1: a1(at),
            };
            let result = match &cell.formula {
                Some(_) => Some(self.formula_result(index, sheet, at, cell, anchor, last)?),
                None => None,
            };
            let record = self.cell_record(index, at, cell, result, &mut merge_index, cell_anchor);
            let (pending, updated) = self.cell_diagnostics(index, at, cell, &record, raised);
            if self.content.diagnostics.len() + pending.len() > MAX_DIAGNOSTICS {
                return Err(self.stop(Some(anchor), Stop::Diagnostics, last));
            }
            let cost = pending
                .iter()
                .map(|diagnostic| measure(diagnostic) + 1)
                .sum::<usize>()
                + measure(&record)
                + 1;
            if self.used + cost > self.max_bytes {
                return Err(self.stop(Some(anchor), Stop::Bytes, last));
            }
            self.used += cost;
            self.content.diagnostics.extend(pending);
            raised = updated;
            self.cells += 1;
            last = Some(at);
            self.current().cells.push(record);
        }
        self.visited += cursor.seeks() - previous_seeks;
        Ok(())
    }

    /// What is known about a formula cell's stored result. A result counts as missing only
    /// while nothing has calculated since a package that stored none was read; once
    /// something has, or the cell cannot be traced to its source, it is uncertain.
    fn formula_result(
        &mut self,
        index: usize,
        sheet: &Sheet,
        at: CellRef,
        cell: &Cell,
        anchor: &XlsxAnchor,
        last: Option<CellRef>,
    ) -> std::result::Result<XlsxFormulaResult, Stopped> {
        let key = (index as u32, at.row, at.col);
        if self.failures[0].contains(&key) {
            return Ok(XlsxFormulaResult::Cycle);
        }
        if self.failures[1].contains(&key) {
            return Ok(XlsxFormulaResult::Limited);
        }
        if matches!(cell.value, CellValue::Empty) {
            let mut spilled = false;
            for (_, spill) in sheet.array_formulas() {
                self.visit(Some(anchor), 1).map_err(|stopped| Stopped {
                    after: last,
                    ..stopped
                })?;
                if spill.contains(at) {
                    spilled = true;
                    break;
                }
            }
            return Ok(if spilled {
                XlsxFormulaResult::Uncertain
            } else {
                XlsxFormulaResult::Missing
            });
        }
        let source = self.source;
        let uncached = match source.facts(index) {
            Some(facts) if !facts.uncached_formulas.is_empty() => {
                match source.source_cell(index, at) {
                    Ok(from) => from.map(|from| facts.uncached_formulas.contains(&from)),
                    Err(()) => None,
                }
            }
            _ => Some(false),
        };
        Ok(match uncached {
            Some(false) => XlsxFormulaResult::Unverified,
            Some(true) if !source.edited => XlsxFormulaResult::Missing,
            _ => XlsxFormulaResult::Uncertain,
        })
    }

    fn cell_record(
        &self,
        index: usize,
        at: CellRef,
        cell: &Cell,
        formula_result: Option<XlsxFormulaResult>,
        merges: &mut MergeSweep,
        anchor: XlsxAnchor,
    ) -> XlsxExportCell {
        let model = self.source.model;
        let value = match &cell.value {
            CellValue::Empty => XlsxExportValue::Empty,
            CellValue::Number { value } if value.is_finite() => {
                XlsxExportValue::Number { value: *value }
            }
            CellValue::Number { .. } => XlsxExportValue::Unavailable {
                reason: "non-finite number".to_owned(),
            },
            CellValue::Text { value } => XlsxExportValue::Text {
                value: value.clone(),
            },
            CellValue::Bool { value } => XlsxExportValue::Bool { value: *value },
            CellValue::Error { value } => XlsxExportValue::Error { value: *value },
        };
        XlsxExportCell {
            id: format!("s{index}!{}", a1(at)),
            anchor,
            value,
            formula: cell.formula.clone(),
            display_text: display_text(&model.styles, model.date_system, cell),
            number_format: format_code_for_cell(&model.styles, cell),
            formula_result,
            merge: merges.at(at).map(|range| XlsxCellMerge {
                range: a1_range(range),
                origin: range.start.row == at.row && range.start.col == at.col,
            }),
        }
    }

    /// The once-per-sheet diagnostics `record` raises, with the flags after them.
    fn cell_diagnostics(
        &self,
        index: usize,
        at: CellRef,
        cell: &Cell,
        record: &XlsxExportCell,
        raised: Raised,
    ) -> (Vec<XlsxExportDiagnostic>, Raised) {
        let mut raised = raised;
        let mut pending = Vec::new();
        let mut raise = |code, severity, message: String| {
            pending.push(XlsxExportDiagnostic {
                code,
                severity,
                anchor: Some(record.anchor.clone()),
                message,
            });
        };
        if record.formula_result.is_some() && !raised.formula {
            raised.formula = true;
            raise(
                XlsxExportDiagnosticCode::FormulaCacheUnverified,
                XlsxExportSeverity::Info,
                "Formula results on this sheet are the stored values; the export does not recalculate or verify them.".to_owned(),
            );
        }
        match record.formula_result {
            Some(XlsxFormulaResult::Missing | XlsxFormulaResult::Uncertain) if !raised.missing => {
                raised.missing = true;
                raise(
                    XlsxExportDiagnosticCode::FormulaResultMissing,
                    XlsxExportSeverity::Warning,
                    "This formula has no stored result, or none that can be confirmed; its value is what the engine holds, not a result the file stored. Each cell's formulaResult says which formulas this applies to.".to_owned(),
                );
            }
            Some(XlsxFormulaResult::Cycle | XlsxFormulaResult::Limited) if !raised.failure => {
                raised.failure = true;
                raise(
                    XlsxExportDiagnosticCode::CalculationFailure,
                    XlsxExportSeverity::Warning,
                    "The last calculation left this formula in a circular reference or at an evaluation limit; each cell's formulaResult says which.".to_owned(),
                );
            }
            _ => {}
        }
        if matches!(record.value, XlsxExportValue::Unavailable { .. }) && !raised.unavailable {
            raised.unavailable = true;
            raise(
                XlsxExportDiagnosticCode::ValueUnavailable,
                XlsxExportSeverity::Warning,
                "This cell holds a value JSON cannot carry.".to_owned(),
            );
        }
        if !raised.approximate && format_is_approximate(&cell.value, &record.number_format) {
            raised.approximate = true;
            raise(
                XlsxExportDiagnosticCode::FormattingApproximate,
                XlsxExportSeverity::Info,
                format!(
                    "The number format {:?} is only approximated, so displayText falls back to General here and wherever this sheet uses it.",
                    record.number_format
                ),
            );
        }
        if !raised.rich && self.rich_text(index, at, cell) {
            raised.rich = true;
            raise(
                XlsxExportDiagnosticCode::RichTextOmitted,
                XlsxExportSeverity::Info,
                "Text formatting runs inside cells on this sheet are not exported; the text is."
                    .to_owned(),
            );
        }
        (pending, raised)
    }

    /// Whether `cell` still holds a string authored with formatted runs: a shared string
    /// by its entry, an inline string by its source cell. When the sheet's row and column
    /// mappings are gone, any text on a sheet that had rich inline strings may.
    fn rich_text(&self, index: usize, at: CellRef, cell: &Cell) -> bool {
        let source = self.source;
        let (Some(package), CellValue::Text { value }) = (source.package, &cell.value) else {
            return false;
        };
        let shared = source
            .shared_string_cells
            .get(index)
            .and_then(|cells| cells.get(&(at.row, at.col)))
            .is_some_and(|&entry| {
                package.shared_string_is_rich(entry)
                    && source.model.shared_strings.get(entry) == Some(value)
            });
        shared
            || source.facts(index).is_some_and(|facts| {
                !facts.rich_inline.is_empty()
                    && match source.source_cell(index, at) {
                        Ok(from) => {
                            from.and_then(|from| facts.rich_inline.get(&from)) == Some(value)
                        }
                        Err(()) => true,
                    }
            })
    }
}

fn sheet_of(anchor: &XlsxAnchor) -> Option<XlsxAnchor> {
    match anchor {
        XlsxAnchor::Sheet { sheet }
        | XlsxAnchor::Cell { sheet, .. }
        | XlsxAnchor::Range { sheet, .. } => Some(XlsxAnchor::Sheet {
            sheet: sheet.clone(),
        }),
        _ => None,
    }
}

fn object_kind(kind: DrawingObjectKind) -> XlsxObjectKind {
    match kind {
        DrawingObjectKind::Chart => XlsxObjectKind::Chart,
        DrawingObjectKind::Picture => XlsxObjectKind::Picture,
        DrawingObjectKind::Shape => XlsxObjectKind::Shape,
        DrawingObjectKind::Group => XlsxObjectKind::Group,
        DrawingObjectKind::Connector => XlsxObjectKind::Connector,
        DrawingObjectKind::Diagram => XlsxObjectKind::Diagram,
        DrawingObjectKind::GraphicFrame => XlsxObjectKind::GraphicFrame,
        DrawingObjectKind::ContentPart => XlsxObjectKind::ContentPart,
        DrawingObjectKind::Unknown => XlsxObjectKind::Unknown,
    }
}

/// Whether every row and column an authored anchor names still sits where it was read.
fn unmoved(axes: &SheetAxes, anchor: &ChartAnchor) -> bool {
    let (rows, cols) = match anchor {
        ChartAnchor::TwoCell { from, to, .. } => ([from.row, to.row], [from.col, to.col]),
        ChartAnchor::OneCell { from, .. } => ([from.row, from.row], [from.col, from.col]),
        ChartAnchor::Absolute { .. } => return true,
    };
    rows.iter().all(|&row| axes.rows.current(row) == Some(row))
        && cols.iter().all(|&col| axes.cols.current(col) == Some(col))
}

/// The cells holding an anchor's corners; an absolute anchor has none.
fn grid_anchor(sheet: &XlsxSheetIdentity, anchor: &ChartAnchor) -> Option<XlsxAnchor> {
    match anchor {
        ChartAnchor::TwoCell { from, to, .. } => Some(XlsxAnchor::Range {
            sheet: sheet.clone(),
            a1: a1_range(CellRange::new(
                CellRef::new(from.row, from.col),
                CellRef::new(to.row, to.col),
            )),
        }),
        ChartAnchor::OneCell { from, .. } => Some(XlsxAnchor::Cell {
            sheet: sheet.clone(),
            a1: a1(CellRef::new(from.row, from.col)),
        }),
        ChartAnchor::Absolute { .. } => None,
    }
}

/// Finds the merge covering each cell of a row-major walk. Merges never overlap, so the
/// ones spanning the current row have disjoint column bands.
struct MergeSweep {
    pending: Vec<CellRange>,
    next: usize,
    by_col: BTreeMap<u32, CellRange>,
    by_end: BTreeSet<(u32, u32)>,
}

impl MergeSweep {
    fn new(sorted: Vec<CellRange>) -> Self {
        Self {
            pending: sorted,
            next: 0,
            by_col: BTreeMap::new(),
            by_end: BTreeSet::new(),
        }
    }

    fn at(&mut self, cell: CellRef) -> Option<CellRange> {
        while let Some(&(end, col)) = self.by_end.first()
            && end < cell.row
        {
            self.by_end.pop_first();
            if self
                .by_col
                .get(&col)
                .is_some_and(|range| range.end.row == end)
            {
                self.by_col.remove(&col);
            }
        }
        while let Some(range) = self.pending.get(self.next)
            && range.start.row <= cell.row
        {
            if range.end.row >= cell.row {
                self.by_col.insert(range.start.col, *range);
                self.by_end.insert((range.end.row, range.start.col));
            }
            self.next += 1;
        }
        self.by_col
            .range(..=cell.col)
            .next_back()
            .map(|(_, range)| *range)
            .filter(|range| range.contains(cell))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Workbook;
    use xlsx_model::Cell;

    /// One sheet whose drawing holds `pictures` pictures.
    fn pictured(pictures: usize) -> Vec<u8> {
        let mut model = WorkbookModel::default();
        let mut sheet = Sheet::new("Pics");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
        model.sheets.push(sheet);
        let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
        let anchor = r#"<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/></xdr:nvPicPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"#;
        parts.push((
            "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#.to_vec(),
        ));
        parts.push((
            "xl/drawings/drawing1.xml".to_owned(),
            format!(
                r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">{}</xdr:wsDr>"#,
                anchor.repeat(pictures)
            )
            .into_bytes(),
        ));
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn export_with(workbook: &Workbook, budgets: Budgets) -> XlsxStructuredContent {
        export_within(
            &workbook.export_source(),
            &XlsxExportOptions::default(),
            XlsxAnchorScope::Snapshot,
            budgets,
        )
        .unwrap()
    }

    #[test]
    fn a_small_inspection_budget_stops_reading_the_package_early() {
        let workbook = Workbook::open_for_read(&pictured(200)).unwrap();
        let full = export_with(&workbook, Budgets::default());
        assert!(!full.truncated);
        assert_eq!(full.sheets[0].objects.len(), 200);

        let small = Budgets {
            inspection: InspectionBudget {
                nodes: 50,
                bytes: u64::MAX,
            },
            ..Budgets::default()
        };
        let content = export_with(&workbook, small);
        assert!(content.truncated && content.sheets[0].truncated);
        assert!(content.sheets[0].objects.is_empty());
        let last = content.diagnostics.last().unwrap();
        assert_eq!(last.code, XlsxExportDiagnosticCode::Truncated);
        assert!(
            last.message.contains("source-inspection"),
            "{}",
            last.message
        );

        let few_bytes = Budgets {
            inspection: InspectionBudget {
                nodes: u64::MAX,
                bytes: 64,
            },
            ..Budgets::default()
        };
        assert!(export_with(&workbook, few_bytes).truncated);

        let few_visits = Budgets {
            visited: 20,
            ..Budgets::default()
        };
        let visited = export_with(&workbook, few_visits);
        assert!(visited.truncated);
        assert_eq!(visited.sheets[0].objects.len(), 20);
        assert!(
            visited
                .diagnostics
                .last()
                .unwrap()
                .message
                .contains("visited-cell")
        );
    }
}
