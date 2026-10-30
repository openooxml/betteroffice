use serde::{Deserialize, Serialize};

use crate::{CellRef, DifferentialFormat, Error, Result, SheetId, Workbook};

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
