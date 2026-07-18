//! Prescan — the pure look-ahead phase of pagination. Port of
//! `packages/core/src/layout/pagination/prescan.ts`.
//!
//! Gathers per-section geometry, section break types, and the keep-with-next
//! scan in one pass so placement reads a plan instead of recomputing
//! look-ahead mid-flight. The keep-with-next scan comes through its feature
//! hook (see `hooks.rs`).

use crate::LayoutError;
use crate::hooks::{self, KeepWithNextScan};
use crate::types::{ColumnLayout, LayoutBlock, MeasuredBlock, PageMargins, SectionBreakType, Size};

/// TS `SectionLayoutConfig` — page-flow geometry of a single section.
#[derive(Debug, Clone)]
pub struct SectionLayoutConfig {
    pub page_size: Size,
    pub margins: PageMargins,
    /// Sections without explicit columns inherit `{ count: 1 }`.
    pub columns: Option<ColumnLayout>,
}

/// TS `DEFAULT_COLUMNS` — single-column fallback.
pub fn default_columns() -> ColumnLayout {
    ColumnLayout {
        count: 1.0,
        gap: 0.0,
        equal_width: None,
        separator: None,
        columns: None,
    }
}

/// TS `collectSectionConfigs` — one config per section break plus a trailing
/// `final_config`; `break_indices` is 1-to-1 with the inner entries.
pub fn collect_section_configs(
    measured: &[MeasuredBlock],
    initial_config: &SectionLayoutConfig,
    final_config: SectionLayoutConfig,
) -> (Vec<SectionLayoutConfig>, Vec<usize>) {
    let mut configs: Vec<SectionLayoutConfig> = Vec::new();
    let mut break_indices: Vec<usize> = Vec::new();
    let mut previous_config = initial_config.clone();
    for (i, mb) in measured.iter().enumerate() {
        let LayoutBlock::SectionBreak(sb) = &mb.block else {
            continue;
        };
        let config = SectionLayoutConfig {
            page_size: sb
                .page_size
                .clone()
                .unwrap_or_else(|| previous_config.page_size.clone()),
            margins: sb
                .margins
                .clone()
                .unwrap_or_else(|| previous_config.margins.clone()),
            columns: sb.columns.clone(),
        };
        configs.push(config.clone());
        break_indices.push(i);
        previous_config = config;
    }
    configs.push(final_config);
    (configs, break_indices)
}

/// TS `LayoutPlan` — the look-ahead the placement walk consumes.
pub struct LayoutPlan {
    /// One config per section break, plus a trailing final-section config.
    pub section_configs: Vec<SectionLayoutConfig>,
    /// Block index of each section break, 1-to-1 with the inner configs.
    pub break_indices: Vec<usize>,
    /// Break type entering section `i + 1`; trailing entry is the body's.
    pub section_break_types: Vec<Option<SectionBreakType>>,
    /// Keep-with-next groups keyed by head, plus their interior-member set.
    pub keep_with_next: KeepWithNextScan,
}

/// TS `prescan` — gather all pagination look-ahead into one plan.
pub fn prescan(
    measured: &[MeasuredBlock],
    body_config: &SectionLayoutConfig,
    final_config: SectionLayoutConfig,
    body_break_type: Option<SectionBreakType>,
) -> Result<LayoutPlan, LayoutError> {
    let (section_configs, break_indices) =
        collect_section_configs(measured, body_config, final_config);
    let mut section_break_types: Vec<Option<SectionBreakType>> = break_indices
        .iter()
        .map(|&i| match &measured[i].block {
            LayoutBlock::SectionBreak(sb) => sb.break_type,
            _ => None,
        })
        .collect();
    section_break_types.push(body_break_type);
    let keep_with_next = hooks::analyze_keep_with_next(measured)?;
    Ok(LayoutPlan {
        section_configs,
        break_indices,
        section_break_types,
        keep_with_next,
    })
}
