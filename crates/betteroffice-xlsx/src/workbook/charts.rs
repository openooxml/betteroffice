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
    /// A chart reference holding a character xml 1.0 cannot carry is kept
    /// only as the source wrote it: an opened chart's reference at the same
    /// position, which a save leaves in place. One an edit or a peer would
    /// introduce or rewrite could not be written back, so it is refused before
    /// anything changes, and the chart stays as authored.
    pub(super) fn validate_chart_references(&self, model: &xlsx_model::Workbook) -> Result<()> {
        for chart in model.sheets.iter().flat_map(|sheet| &sheet.charts) {
            for (index, reference) in chart.refs.iter().enumerate() {
                if super::is_writable_xml_text(&reference.formula) {
                    continue;
                }
                let authored = self.authority.source_charts().any(|source| {
                    source.part == chart.part
                        && source
                            .refs
                            .get(index)
                            .is_some_and(|opened| opened.formula == reference.formula)
                });
                if !authored {
                    return Err(Error::InvalidOperation(
                        "chart reference contains a character xml cannot carry".to_string(),
                    ));
                }
            }
        }
        Ok(())
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
            let follows = package.facts().chart_part_bytes(part).is_some_and(|bytes| {
                xlsx_parse::agreed_chart_bytes(
                    bytes,
                    part,
                    claims.get(part).map_or(&[][..], Vec::as_slice),
                    preview,
                )
                .is_ok()
            });
            if !follows {
                return Err(Error::InvalidOperation(format!(
                    "{part} references cells this edit would move, and it cannot be rewritten"
                )));
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
        self.validate_chart_references(&preview)?;
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
    use xlsx_model::{
        AnchorCell, AnchorEditAs, ChartAnchor, ChartRef, ChartRefKind, Sheet, SheetChart,
    };

    use super::Workbook;

    fn model(formula: &str) -> xlsx_model::Workbook {
        let mut sheet = Sheet::new("A\u{1}B");
        sheet.charts.push(SheetChart {
            raw_xml: None,
            part: "xl/charts/chart1.xml".to_owned(),
            drawing: "xl/drawings/drawing1.xml".to_owned(),
            anchor_index: 0,
            anchor: ChartAnchor::TwoCell {
                from: AnchorCell::default(),
                to: AnchorCell {
                    col: 4,
                    row: 8,
                    ..AnchorCell::default()
                },
                edit_as: AnchorEditAs::TwoCell,
            },
            refs: vec![ChartRef {
                kind: ChartRefKind::Values,
                formula: formula.to_owned(),
            }],
        });
        let mut model = xlsx_model::Workbook::default();
        model.sheets.push(sheet);
        model
    }

    /// The check every edit, peer state and hydrated snapshot passes: a
    /// reference holding a character xml cannot carry is accepted only as the
    /// opened chart wrote it, at the same position of the same part.
    #[test]
    fn only_an_authored_reference_may_hold_a_character_xml_cannot_carry() {
        let authored = "'A\u{1}B'!$B$2:$B$4";
        let workbook = Workbook::from_model(model(authored)).unwrap();
        workbook
            .validate_chart_references(&model(authored))
            .unwrap();
        workbook
            .validate_chart_references(&model("Data!$B$2:$B$4"))
            .unwrap();
        for rewritten in ["'A\u{1}B'!$B$3:$B$5", "Data!$B$2\u{1}"] {
            let error = workbook
                .validate_chart_references(&model(rewritten))
                .unwrap_err();
            assert!(error.to_string().contains("xml cannot carry"), "{error}");
        }
        let mut other_part = model(authored);
        other_part.sheets[0].charts[0].part = "xl/charts/chart2.xml".to_owned();
        assert!(workbook.validate_chart_references(&other_part).is_err());
    }
}
