use serde::{Deserialize, Serialize};

use crate::{
    CellEdit, CellRange, CellRef, DifferentialFormat, Error, Result, SheetId, SheetInfo, Viewport,
    Workbook,
};

/// Sheet metadata shared by binding payloads.
#[derive(Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetInfoWire {
    sheet_ids: Vec<String>,
    sheet_names: Vec<String>,
    active_sheet: u32,
    content_width: f32,
    content_height: f32,
    frozen_rows: u32,
    frozen_cols: u32,
    initial_scroll_x: f32,
    initial_scroll_y: f32,
    auto_filter: Option<String>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct CellEditWire {
    a1: String,
    input: String,
    is_formula: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    is_array_formula: Option<bool>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct TextSearchMatchWire {
    sheet: u32,
    sheet_id: String,
    sheet_name: String,
    row: u32,
    col: u32,
    a1: String,
    text: String,
}

fn read_json(value: &impl Serialize) -> Result<String> {
    serde_json::to_string(value).map_err(|error| Error::InvalidOperation(error.to_string()))
}

/// Filters and pagination for [`Workbook::list_defined_names`].
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename = "Request", deny_unknown_fields)]
pub struct DefinedNamesRequest {
    /// Omitted selects all scopes; null selects workbook scope.
    #[serde(
        default,
        deserialize_with = "scope_filter",
        skip_serializing_if = "Option::is_none"
    )]
    pub scope: Option<Option<u32>>,
    /// Number of matching names to skip, defaulting to zero.
    pub offset: Option<usize>,
    /// Page size from 1 to 100, defaulting to 100.
    pub limit: Option<usize>,
}

fn scope_filter<'de, D: serde::Deserializer<'de>>(
    deserializer: D,
) -> std::result::Result<Option<Option<u32>>, D::Error> {
    Option::<u32>::deserialize(deserializer).map(Some)
}

/// One defined name in source order.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DefinedNameInfo {
    /// Whether the name is hidden.
    pub hidden: bool,
    /// Defined name spelling.
    pub name: String,
    /// Formula or reference without a leading equals sign.
    pub refers_to: String,
    /// Local sheet index, or null for workbook scope.
    pub scope: Option<u32>,
}

/// A page of defined names at the current document version.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DefinedNamesResult {
    /// Matching names in source order.
    pub items: Vec<DefinedNameInfo>,
    /// Offset of the next page, or null at the end.
    pub next_offset: Option<usize>,
    /// Total matching names before pagination.
    pub total: usize,
    /// Document version observed by this read.
    pub version: String,
}

/// One conditional-format rule and its sheet range.
#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ConditionalFormatInfo {
    /// Source differential-format index, when present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<u32>,
    /// Whether the rule type supports editing.
    pub editable: bool,
    /// Supported differential formatting properties.
    pub format: DifferentialFormat,
    /// Rule formulas in source order.
    pub formulas: Vec<String>,
    /// Rule identifier.
    pub id: String,
    /// Comparison operator, when present.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    /// Rule evaluation priority.
    pub priority: u32,
    /// Source range or space-separated ranges in A1 notation.
    pub range: String,
    /// Conditional-format rule type.
    #[serde(rename = "type")]
    pub rule_type: String,
}

/// Limits reached by the last recalculation.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculationStatus {
    /// Limited cells formatted relative to the active sheet.
    pub limited_cells: Vec<String>,
    /// Whether recalculation exhausted its shared budget, omitted when false.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub budget_exhausted: bool,
}

impl Workbook {
    /// Serializes metadata for the active sheet, including its auto-filter.
    pub fn sheet_info_json(&self) -> Result<String> {
        self.sheet_info_wire_json(self.sheet_info()?)
    }

    /// Serializes metadata for a sheet without changing the active sheet.
    pub fn sheet_info_for_json(&self, sheet: SheetId) -> Result<String> {
        self.sheet_info_wire_json(self.sheet_info_for(sheet)?)
    }

    fn sheet_info_wire_json(&self, info: SheetInfo) -> Result<String> {
        read_json(&SheetInfoWire {
            auto_filter: self.auto_filter(info.active_sheet)?,
            sheet_ids: info.sheet_ids,
            sheet_names: info.sheet_names,
            active_sheet: info.active_sheet.0,
            content_width: info.content_width,
            content_height: info.content_height,
            frozen_rows: info.frozen_rows,
            frozen_cols: info.frozen_cols,
            initial_scroll_x: info.initial_scroll_x,
            initial_scroll_y: info.initial_scroll_y,
        })
    }

    /// Serializes editable cell input with array-formula membership.
    pub fn cell_json(&self, sheet: SheetId, cell: CellRef) -> Result<String> {
        read_json(&self.cell_edit_wire(sheet, self.cell(sheet, cell)?)?)
    }

    fn cell_edit_wire(&self, sheet: SheetId, cell: CellEdit) -> Result<CellEditWire> {
        Ok(CellEditWire {
            a1: cell.cell.to_a1(),
            is_array_formula: self.is_array_formula(sheet, cell.cell)?.then_some(true),
            input: cell.input,
            is_formula: cell.is_formula,
        })
    }

    /// Serializes row-major editable cell inputs as a JSON array.
    pub fn range_cells_json(&self, sheet: SheetId, range: CellRange) -> Result<String> {
        let cells = self
            .range_cells(sheet, range)?
            .into_iter()
            .map(|row| {
                row.into_iter()
                    .map(|cell| self.cell_edit_wire(sheet, cell))
                    .collect::<Result<Vec<_>>>()
            })
            .collect::<Result<Vec<_>>>()?;
        read_json(&cells)
    }

    /// Serializes formatted-text matches with sheet identities.
    pub fn search_text_json(
        &self,
        query: &str,
        case_sensitive: bool,
        limit: Option<u32>,
    ) -> Result<String> {
        let info = self.sheet_info()?;
        let matches = self
            .search_text(query, case_sensitive, limit.map(|limit| limit as usize))
            .into_iter()
            .map(|result| {
                let sheet = result.address.sheet.0;
                TextSearchMatchWire {
                    sheet,
                    sheet_id: info.sheet_ids[sheet as usize].clone(),
                    sheet_name: info.sheet_names[sheet as usize].clone(),
                    row: result.address.cell.row,
                    col: result.address.cell.col,
                    a1: result.address.cell.to_a1(),
                    text: result.text,
                }
            })
            .collect::<Vec<_>>();
        read_json(&matches)
    }

    /// Serializes uniform formatting properties for a selection.
    pub fn selection_formatting_json(&self, sheet: SheetId, range: CellRange) -> Result<String> {
        read_json(&self.selection_formatting(sheet, range)?)
    }

    /// Serializes overlapping merged ranges as a JSON array.
    pub fn merged_range_list_json(&self, sheet: SheetId, range: CellRange) -> Result<String> {
        read_json(&self.merged_ranges(sheet, range)?)
    }

    /// Serializes merged ranges visible in a JSON viewport as an array.
    pub fn visible_merged_ranges_json(&self, sheet: SheetId, viewport: &str) -> Result<String> {
        let viewport: Viewport = serde_json::from_str(viewport)
            .map_err(|error| Error::InvalidRequest(error.to_string()))?;
        read_json(&self.visible_merged_ranges(sheet, &viewport)?)
    }

    /// Serializes chart metadata for a sheet as a JSON array.
    pub fn charts_json(&self, sheet: SheetId) -> Result<String> {
        read_json(&self.charts(sheet)?)
    }

    /// Serializes limits reached by the last recalculation.
    pub fn calculation_status_json(&self) -> Result<String> {
        read_json(&self.calculation_status())
    }

    /// Returns the sheet's auto-filter range in A1 notation, or None when absent.
    pub fn auto_filter(&self, sheet: SheetId) -> Result<Option<String>> {
        Ok(self
            .sheet(sheet)?
            .auto_filter
            .as_ref()
            .map(|filter| filter.range.to_a1()))
    }

    /// Returns whether the cell belongs to an array formula, including its anchor.
    pub fn is_array_formula(&self, sheet: SheetId, at: CellRef) -> Result<bool> {
        self.validate_cell(at)?;
        Ok(self
            .sheet(sheet)?
            .array_formulas()
            .any(|(_, range)| range.contains(at)))
    }

    /// Lists defined names in source order with optional scope and pagination.
    pub fn list_defined_names(&self, request: &DefinedNamesRequest) -> Result<DefinedNamesResult> {
        let offset = request.offset.unwrap_or(0);
        let limit = request.limit.unwrap_or(100);
        if !(1..=100).contains(&limit) {
            return Err(Error::InvalidOperation(
                "limit must be between 1 and 100".into(),
            ));
        }
        if let Some(Some(scope)) = request.scope
            && self.model().sheet(SheetId(scope)).is_none()
        {
            return Err(Error::InvalidOperation(
                "scope sheet is out of range".into(),
            ));
        }
        let names = self.model().defined_names.iter().filter(|name| {
            request
                .scope
                .is_none_or(|scope| scope == name.local_sheet.map(|sheet| sheet.0))
        });
        let total = names.clone().count();
        let items = names
            .skip(offset)
            .take(limit)
            .map(|name| DefinedNameInfo {
                hidden: name.hidden,
                name: name.name.clone(),
                refers_to: name.formula.clone(),
                scope: name.local_sheet.map(|sheet| sheet.0),
            })
            .collect::<Vec<_>>();
        let next = offset.saturating_add(items.len());
        Ok(DefinedNamesResult {
            items,
            next_offset: (next < total).then_some(next),
            total,
            version: self.version().to_string(),
        })
    }

    /// Lists defined names using the engine's JSON request and response shapes.
    pub fn list_defined_names_json(&self, request: &str) -> Result<String> {
        let request = serde_json::from_str(request)
            .map_err(|error| Error::InvalidOperation(error.to_string()))?;
        let result = self.list_defined_names(&request)?;
        let value = serde_json::to_value(result)
            .map_err(|error| Error::InvalidOperation(error.to_string()))?;
        serde_json::to_string(&value).map_err(|error| Error::InvalidOperation(error.to_string()))
    }

    /// Lists conditional-format rules in group and rule order.
    pub fn conditional_formats(&self, sheet: SheetId) -> Result<Vec<ConditionalFormatInfo>> {
        let sheet_ref = self
            .model()
            .sheet(sheet)
            .ok_or_else(|| Error::InvalidOperation(format!("sheet {} not found", sheet.0)))?;
        Ok(sheet_ref
            .conditional_formats
            .iter()
            .flat_map(|group| {
                group.rules.iter().map(|rule| ConditionalFormatInfo {
                    dxf_id: rule.dxf_id,
                    editable: rule.editable(),
                    format: rule.format.clone(),
                    formulas: rule.formulas.clone(),
                    id: rule.id.clone(),
                    operator: rule.operator.clone(),
                    priority: rule.priority,
                    range: group.sqref.clone(),
                    rule_type: rule.rule_type.clone(),
                })
            })
            .collect())
    }

    /// Lists conditional-format rules using the engine's JSON response shape.
    pub fn conditional_formats_json(&self, sheet: SheetId) -> Result<String> {
        let items = self.conditional_formats(sheet)?;
        let value = serde_json::to_value(items)
            .map_err(|error| Error::InvalidOperation(error.to_string()))?;
        serde_json::to_string(&value).map_err(|error| Error::InvalidOperation(error.to_string()))
    }

    /// Returns the last recalculation's limited cells and budget status.
    pub fn calculation_status(&self) -> CalculationStatus {
        let calculation = self.last_calculation();
        CalculationStatus {
            limited_cells: calculation
                .limited_cells
                .iter()
                .map(|address| self.format_address(*address))
                .collect(),
            budget_exhausted: calculation.budget_exhausted,
        }
    }
}
