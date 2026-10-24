use serde::{Deserialize, Serialize};
use xlsx_model::{ChartAnchor, ChartEdit, SheetId};
use xlsx_ops::Op;
use xlsx_parse::ChartMetadata;

use super::{
    CalculationScope, CommitHistory, StagedApply, SyncOrigin, Workbook, validate_model_sheets,
    validate_shared_drawings,
};
use crate::{CalculationOptions, Error, MutationResult, Result};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartMutation {
    pub sheet: SheetId,
    pub chart: usize,
    #[serde(flatten)]
    pub edit: ChartEdit,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartInfo {
    pub id: String,
    pub index: usize,
    pub part: String,
    pub anchor_range: Option<String>,
    /// false when the chart part cannot be read; `error` then says why and
    /// the metadata is empty.
    pub editable: bool,
    pub error: Option<String>,
    #[serde(flatten)]
    pub metadata: ChartMetadata,
}

impl std::ops::Deref for ChartInfo {
    type Target = ChartMetadata;
    fn deref(&self) -> &Self::Target {
        &self.metadata
    }
}

impl Workbook {
    pub fn charts(&self, sheet: SheetId) -> Result<Vec<ChartInfo>> {
        let owner = self
            .model
            .sheet(sheet)
            .ok_or(Error::SheetOutOfRange(sheet))?;
        Ok(owner
            .charts
            .iter()
            .enumerate()
            .map(|(index, chart)| {
                let metadata =
                    self.chart_bytes(chart, &self.model, &owner.name)
                        .and_then(|bytes| {
                            Ok(xlsx_parse::chart_metadata(
                                &bytes,
                                &self.model,
                                &owner.name,
                            )?)
                        });
                let anchor_range = match chart.anchor {
                    ChartAnchor::TwoCell { from, to, .. } => Some(
                        xlsx_model::CellRange::new(
                            xlsx_model::CellRef::new(from.row, from.col),
                            xlsx_model::CellRef::new(to.row, to.col),
                        )
                        .to_a1(),
                    ),
                    ChartAnchor::OneCell { from, .. } => {
                        Some(xlsx_model::CellRef::new(from.row, from.col).to_a1())
                    }
                    ChartAnchor::Absolute { .. } => None,
                };
                let (metadata, error) = match metadata {
                    Ok(metadata) => (metadata, None),
                    Err(error) => (
                        ChartMetadata {
                            chart_type: String::new(),
                            title: None,
                            auto_title: false,
                            rich_title: false,
                            series: Vec::new(),
                        },
                        Some(error.to_string()),
                    ),
                };
                ChartInfo {
                    id: chart.frame_id(),
                    index,
                    part: chart.part.clone(),
                    anchor_range,
                    editable: error.is_none(),
                    error,
                    metadata,
                }
            })
            .collect::<Vec<_>>())
    }

    pub(super) fn chart_bytes(
        &self,
        chart: &xlsx_model::SheetChart,
        model: &xlsx_model::Workbook,
        owner: &str,
    ) -> Result<Vec<u8>> {
        let missing =
            || Error::InvalidOperation(format!("chart part {} has no source", chart.part));
        let package = self.source_package.as_ref().ok_or_else(missing)?;
        let bytes = package
            .facts()
            .chart_part_bytes(&chart.part)
            .ok_or_else(missing)?;
        let read = match &chart.raw_xml {
            Some(xml) => xlsx_parse::chart_references(xml.as_bytes())?,
            None => xlsx_parse::chart_references(bytes)?,
        };
        if read == chart.refs || chart.raw_xml.is_some() {
            return Ok(xlsx_parse::chart_bytes(
                bytes,
                chart,
                model,
                owner,
                (model, owner),
            )?);
        }
        let opened = package.materialize()?;
        let original_owner = self
            .model
            .sheets
            .iter()
            .position(|sheet| sheet.name == owner)
            .and_then(|index| self.preserved.origins.get(index).copied().flatten())
            .and_then(|origin| opened.source_sheet_name(origin))
            .unwrap_or(owner);
        Ok(xlsx_parse::chart_bytes(
            bytes,
            chart,
            model,
            owner,
            (opened.original_workbook(), original_owner),
        )?)
    }

    pub fn edit_charts(
        &mut self,
        edits: Vec<ChartMutation>,
        options: CalculationOptions,
    ) -> Result<MutationResult> {
        if self.is_collaborative() {
            return Err(Error::CollaborativeStructureOperation);
        }
        if edits.is_empty() || edits.len() > 32 {
            return Err(Error::InvalidOperation(
                "edits must contain 1 to 32 chart edits".into(),
            ));
        }
        let package = self
            .source_package
            .as_ref()
            .map(|package| package.materialize())
            .transpose()?;
        let cell_sheets = (0..self.model.sheets.len())
            .map(|index| {
                match (
                    self.preserved.origins.get(index).copied().flatten(),
                    package,
                ) {
                    (Some(origin), Some(package)) => matches!(
                        package.source_sheet_kind(origin),
                        Some(xlsx_parse::SourceSheetKind::Worksheet)
                    ),
                    _ => true,
                }
            })
            .collect::<Vec<_>>();
        let mut preview = self.model.clone();
        let mut changed = std::collections::BTreeSet::new();
        for (index, mutation) in edits.iter().enumerate() {
            let result: Result<()> = (|| {
                let sheet = preview
                    .sheet(mutation.sheet)
                    .ok_or(Error::SheetOutOfRange(mutation.sheet))?;
                if let (Some(origin), Some(package)) = (
                    self.preserved
                        .origins
                        .get(mutation.sheet.0 as usize)
                        .copied()
                        .flatten(),
                    package,
                ) && !package.source_sheet_allows_chart_edits(origin)
                {
                    return Err(Error::InvalidOperation(format!(
                        "sheet {} does not allow chart edits; choose an unprotected worksheet or chartsheet from the sheet list",
                        sheet.name
                    )));
                }
                let chart = sheet.charts.get(mutation.chart).ok_or_else(|| {
                    Error::InvalidOperation(format!(
                        "chart {} is out of range; valid chart indices: {}",
                        mutation.chart,
                        (0..sheet.charts.len())
                            .map(|index| index.to_string())
                            .collect::<Vec<_>>()
                            .join(", ")
                    ))
                })?;
                let claims = preview
                    .sheets
                    .iter()
                    .flat_map(|sheet| &sheet.charts)
                    .filter(|other| other.part.eq_ignore_ascii_case(&chart.part))
                    .count();
                if claims > 1 {
                    return Err(Error::InvalidOperation(format!(
                        "chart part {} is shared by {claims} anchors; choose a chart with an unshared part from the chart list",
                        chart.part
                    )));
                }
                let bytes = self.chart_bytes(chart, &preview, &sheet.name)?;
                let (xml, refs) = xlsx_parse::edit_chart(
                    &bytes,
                    &mutation.edit,
                    &preview,
                    &sheet.name,
                    &cell_sheets,
                )?;
                let part = chart.part.clone();
                for (sheet_index, sheet) in preview.sheets.iter_mut().enumerate() {
                    for chart in &mut sheet.charts {
                        if chart.part == part {
                            chart.raw_xml = Some(xml.clone());
                            chart.refs = refs.clone();
                            changed.insert(sheet_index);
                        }
                    }
                }
                Ok(())
            })();
            if let Err(error) = result {
                return Err(Error::InvalidOperation(format!(
                    "edits[{index}]: {error}; no chart edits were applied"
                )));
            }
        }
        if preview == self.model {
            return Ok(MutationResult::default());
        }
        validate_model_sheets(&preview)?;
        validate_shared_drawings(&preview)?;
        let inverse = changed
            .iter()
            .map(|&sheet| Op::SetCharts {
                sheet: SheetId(sheet as u32),
                charts: self.model.sheets[sheet].charts.clone(),
            })
            .collect();
        let ops = changed
            .into_iter()
            .map(|sheet| Op::SetCharts {
                sheet: SheetId(sheet as u32),
                charts: preview.sheets[sheet].charts.clone(),
            })
            .collect();
        let prepared = self.prepare_commit(
            ops,
            StagedApply::new(preview, inverse),
            SyncOrigin::User,
            CommitHistory::Separate,
            None,
        )?;
        let update = self.commit_prepared(prepared)?;
        self.proposals.clear();
        self.recalculate_scope(CalculationScope::Kept, &[], options);
        self.publish(update);
        Ok(MutationResult {
            applied: true,
            ..MutationResult::default()
        })
    }
}
