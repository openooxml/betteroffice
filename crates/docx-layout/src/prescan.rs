use crate::LayoutError;
use crate::hooks::{self, KeepWithNextScan};
use crate::types::{ColumnLayout, LayoutBlock, MeasuredBlock, PageMargins, SectionBreakType, Size};

#[derive(Debug, Clone)]
pub struct SectionLayoutConfig {
    pub page_size: Size,
    pub margins: PageMargins,
    /// Sections without explicit columns inherit `{ count: 1 }`.
    pub columns: Option<ColumnLayout>,
}

pub fn default_columns() -> ColumnLayout {
    ColumnLayout {
        count: 1.0,
        gap: 0.0,
        equal_width: None,
        separator: None,
    }
}

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
