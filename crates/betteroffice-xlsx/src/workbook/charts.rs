use std::collections::{BTreeSet, HashMap};

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
        Ok(xlsx_parse::chart_bytes(bytes, chart, model, owner)?)
    }

    /// Refuses a batch that leaves a chart part where a save could not write
    /// it back: a reference reshaped around a multi-level or unmodelled cache
    /// that cannot be rebuilt, or a part two sheets anchor that the batch
    /// would give two sets of references. Every part the batch changed is
    /// patched exactly as the save will. What decides that refusal is the
    /// references and the sheets they name, not what the cells read, so the
    /// recalculation after the batch cannot change it, short of cell text no
    /// XML can carry. `inverse` is the batch's inverse, which records
    /// a `SetCharts` for every sheet whose charts it changed, so a batch that
    /// changes none costs nothing here.
    /// The chart parts a save keeps as authored because the workbook holds a
    /// reference for them that no `c:f` can carry, decided as the save decides
    /// it, once per part.
    pub(super) fn chart_save_warnings(&self) -> Vec<crate::SaveWarning> {
        let Some(package) = self.source_package.as_ref() else {
            return Vec::new();
        };
        let facts = package.facts();
        let mut warned = BTreeSet::new();
        let mut warnings = Vec::new();
        for sheet in &self.model.sheets {
            for chart in &sheet.charts {
                if warned.contains(chart.part.as_str()) {
                    continue;
                }
                let Some(source) = facts.chart_part_bytes(&chart.part) else {
                    continue;
                };
                if let Err(error) = xlsx_parse::chart_bytes(source, chart, &self.model, &sheet.name)
                    && xlsx_parse::is_unwritable_chart_reference(&error)
                {
                    warned.insert(chart.part.as_str());
                    warnings.push(crate::SaveWarning {
                        part: chart.part.clone(),
                        reason: "kept as the file wrote it: the workbook holds a chart reference with a character XML cannot carry, which no chart reference can be written with".to_owned(),
                    });
                }
            }
        }
        warnings
    }

    pub(super) fn ensure_charts_follow(
        &self,
        preview: &xlsx_model::Workbook,
        inverse: &[Op],
    ) -> Result<()> {
        let Some(package) = self.source_package.as_ref() else {
            return Ok(());
        };
        if !inverse.iter().any(|op| matches!(op, Op::SetCharts { .. })) {
            return Ok(());
        }
        #[cfg(test)]
        CHART_CHECKS.set(CHART_CHECKS.get() + 1);
        let changed = changed_chart_parts(&self.model, preview);
        let mut claims: HashMap<&str, Vec<(&xlsx_model::SheetChart, &str)>> = HashMap::new();
        for sheet in &preview.sheets {
            for chart in &sheet.charts {
                if changed.contains(chart.part.as_str()) {
                    claims
                        .entry(chart.part.as_str())
                        .or_default()
                        .push((chart, sheet.name.as_str()));
                }
            }
        }
        for part in changed {
            let outcome = package.facts().chart_part_bytes(part).map(|bytes| {
                xlsx_parse::agreed_chart_bytes(
                    bytes,
                    part,
                    claims.get(part).map_or(&[][..], Vec::as_slice),
                    preview,
                )
            });
            match outcome {
                Some(Ok(_)) => {}
                Some(Err(error)) if xlsx_parse::is_unwritable_chart_reference(&error) => {
                    return Err(Error::InvalidOperation(
                        "chart reference contains a character xml cannot carry".to_string(),
                    ));
                }
                _ => {
                    return Err(Error::InvalidOperation(format!(
                        "{part} references cells this edit would move, and it cannot be rewritten"
                    )));
                }
            }
        }
        Ok(())
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
        self.recalculate_scope(CalculationScope::Kept, &[], options);
        self.publish(update);
        Ok(MutationResult {
            applied: true,
            ..MutationResult::default()
        })
    }
}

/// The chart parts of `after` some chart of which reads references, or holds
/// an edit, that no chart of the same part in `before` does: the parts a save
/// would patch differently. Linear in the charts of both.
pub(super) fn changed_chart_parts<'a>(
    before: &xlsx_model::Workbook,
    after: &'a xlsx_model::Workbook,
) -> BTreeSet<&'a str> {
    let mut held: HashMap<&str, Vec<&xlsx_model::SheetChart>> = HashMap::new();
    for chart in before.sheets.iter().flat_map(|sheet| &sheet.charts) {
        held.entry(chart.part.as_str()).or_default().push(chart);
    }
    after
        .sheets
        .iter()
        .flat_map(|sheet| &sheet.charts)
        .filter(|chart| {
            !held.get(chart.part.as_str()).is_some_and(|charts| {
                charts
                    .iter()
                    .any(|kept| kept.refs == chart.refs && kept.raw_xml == chart.raw_xml)
            })
        })
        .map(|chart| chart.part.as_str())
        .collect()
}

#[cfg(test)]
thread_local! {
    static CHART_CHECKS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[cfg(test)]
pub(super) fn chart_check_count() -> usize {
    CHART_CHECKS.get()
}

#[cfg(test)]
mod reference_tests {
    use super::Workbook;

    include!("../../tests/fixtures/chart_control_book.rs");

    fn chart_part(bytes: &[u8]) -> Vec<u8> {
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == CHART)
            .map(|(_, bytes)| bytes)
            .unwrap()
    }

    /// Whatever door a rewritten reference xml cannot carry comes through, a
    /// peer state, a snapshot, a checkpoint or one not yet known, the save
    /// keeps the chart part as authored instead of failing or writing invalid
    /// xml, and says so with one warning naming the part.
    #[test]
    fn a_save_keeps_the_chart_as_authored_whatever_installed_such_a_reference() {
        let source = book();
        for installed in ["'A\u{1}B'!$B$3:$B$5", "Data!$B$2\u{0}"] {
            let mut workbook = Workbook::open(&source).unwrap();
            workbook.model.sheets[0].charts[0].refs[0].formula = installed.to_owned();
            let (saved, warnings) = workbook
                .save_with_warnings()
                .unwrap_or_else(|error| panic!("{installed:?}: {error}"));
            assert_eq!(chart_part(&saved), chart_part(&source), "{installed:?}");
            assert_eq!(warnings.len(), 1, "{installed:?}: {warnings:?}");
            assert_eq!(warnings[0].part, CHART);
            assert!(warnings[0].reason.contains("character XML cannot carry"));
        }
    }
}
