//! Read-only lowering from the pilcrow-stream editing model to the existing layout contract.
//!
//! The renderer continues to consume [`LayoutBlock`] values. This module is the model-specific
//! seam: it walks one yrs story in UTF-16 units, resolves side-map comment anchors, lowers the
//! authored OOXML properties carried by yrs, and synthesizes the ProseMirror integer positions
//! required during coexistence.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::sync::Arc;

use docx_layout::types::{
    BlockId, BorderStyle, BoxEdges, CellBorderSpec, CellBorders, ChartBlock, ColumnBreakBlock,
    ColumnLayout, FieldRun, FloatingTablePosition, HyperlinkInfo, ImageRun, ImageRunPosition,
    LayoutBlock, LineBreakRun, ListNumPr, PageBreakBlock, PageMargins, ParagraphAttrs,
    ParagraphBlock, ParagraphBorders, ParagraphIndent, ParagraphSpacing, Run, RunFontSlots,
    RunFormatting, RunLanguageSlots, SdtGroup, SectionBreakBlock, SectionBreakType, ShapeBlock,
    Size, SpacingExplicit, TabRun, TabStop, TableBlock, TableCell, TableRow, TextRun,
    UnderlineSpec,
};
use serde_json::{Map as JsonMap, Value};
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, OffsetKind, Out, ReadTxn, Text, Transact};

use super::{COMMENTS, DEL, EditError, EditingDoc, INS, decode_anchor, is_pilcrow, story_ref};

const AUTO_PARAGRAPH_SPACING_PX: f64 = 14.0;

/// Pre-flattened document values needed while lowering a story.
///
/// Styles remain represented on each pilcrow through `pStyle` and `defaultTextFormatting`.
/// `theme_colors` contains raw six-digit RGB values keyed by OOXML theme slot. Missing slots use
/// the same Office default palette as the TypeScript color resolver. `numeric_ids` is the explicit
/// coexistence adapter for the current numeric layout contract: callers that mirror a PM document
/// should map yrs' client-scoped IDs to the PM revision/comment IDs here.
#[derive(Clone, Debug, Default, PartialEq, serde::Deserialize, serde::Serialize)]
#[serde(default, rename_all = "camelCase")]
pub struct RenderEnv {
    pub theme_colors: BTreeMap<String, String>,
    pub default_tab_stop_twips: Option<f64>,
    pub page_content_height: Option<f64>,
    pub numeric_ids: BTreeMap<String, f64>,
}

impl RenderEnv {
    pub fn with_numeric_id(mut self, yrs_id: impl Into<String>, layout_id: f64) -> Self {
        self.numeric_ids.insert(yrs_id.into(), layout_id);
        self
    }
}

/// A malformed story or a model/bridge configuration mismatch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BridgeError {
    Edit(EditError),
    WrongOffsetKind,
    UnterminatedStory(String),
    UnsupportedEmbed {
        story: String,
        index: u32,
    },
    MalformedTable {
        story: String,
        index: u32,
        detail: String,
    },
    RecursiveStory(String),
}

impl fmt::Display for BridgeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Edit(error) => error.fmt(f),
            Self::WrongOffsetKind => write!(f, "render bridge requires OffsetKind::Utf16"),
            Self::UnterminatedStory(story) => {
                write!(f, "story {story:?} does not end in a pilcrow embed")
            }
            Self::UnsupportedEmbed { story, index } => {
                write!(
                    f,
                    "story {story:?} contains an unsupported embed at {index}"
                )
            }
            Self::MalformedTable {
                story,
                index,
                detail,
            } => write!(f, "malformed table in story {story:?} at {index}: {detail}"),
            Self::RecursiveStory(story) => {
                write!(f, "recursive table story reference to {story:?}")
            }
        }
    }
}

impl std::error::Error for BridgeError {}

impl From<EditError> for BridgeError {
    fn from(value: EditError) -> Self {
        Self::Edit(value)
    }
}

/// The coexistence position formula from `render-bridge.md` section 2.1.
///
/// `story_index` counts UTF-16 units and counts every pilcrow as one. A PM paragraph contributes
/// two tag positions, so each already-crossed pilcrow adds the one extra unit not present in yrs.
pub const fn pm_position(story_index: u32, pilcrows_before: u32) -> u64 {
    story_index as u64 + 1 + pilcrows_before as u64
}

/// Lowers one story from an [`EditingDoc`] to the unchanged renderer vocabulary.
pub fn yrs_doc_to_layout_blocks(
    doc: &EditingDoc,
    story_id: &str,
    env: &RenderEnv,
) -> Result<Vec<LayoutBlock>, BridgeError> {
    if doc.yrs_doc().offset_kind() != OffsetKind::Utf16 {
        return Err(BridgeError::WrongOffsetKind);
    }

    let txn = doc.yrs_doc().transact();
    let mut active_stories = BTreeSet::new();
    let mut list_state = ListState::default();
    lower_story(&txn, story_id, env, 0, &mut active_stories, &mut list_state)
        .map(|(blocks, _)| blocks)
}

fn lower_story<T: ReadTxn>(
    txn: &T,
    story_id: &str,
    env: &RenderEnv,
    pm_base: u64,
    active_stories: &mut BTreeSet<String>,
    list_state: &mut ListState,
) -> Result<(Vec<LayoutBlock>, u64), BridgeError> {
    if !active_stories.insert(story_id.to_owned()) {
        return Err(BridgeError::RecursiveStory(story_id.to_owned()));
    }

    let result = (|| {
        let story = story_ref(txn, story_id)?;
        let comments = resolve_comment_intervals(txn, story_id, env)?;
        let mut blocks = Vec::new();
        let mut paragraph_runs = Vec::new();
        let mut paragraph_drawings = Vec::new();
        let mut story_index = 0_u32;
        let mut paragraph_start = 0_u32;
        let mut paragraph_pm_start = pm_base;
        let mut paragraph_pm_units = 0_u32;
        let mut pm_cursor = pm_base;
        let mut at_block_boundary = true;
        // S5a: section-break margin cascade, per story (sections are body-level;
        // cell/HF stories simply never carry section properties).
        let mut section_margins = SectionMarginsTwips::default();

        for diff in story.diff(txn, YChange::identity) {
            let attributes = diff.attributes.as_deref();
            match diff.insert {
                Out::Any(Any::String(text)) => {
                    let text = text.as_ref();
                    push_text_chunks(
                        &mut paragraph_runs,
                        text,
                        story_index,
                        attributes,
                        &comments,
                        env,
                        paragraph_pm_units,
                    );
                    let width = utf16_len(text);
                    story_index += width;
                    paragraph_pm_units += width;
                    at_block_boundary = false;
                }
                Out::YMap(pilcrow) if is_pilcrow(&pilcrow, txn) => {
                    let paragraph_blocks = flush_paragraph_parts(
                        paragraph_runs,
                        paragraph_drawings,
                        &pilcrow,
                        attributes,
                        txn,
                        story_id,
                        env,
                        paragraph_pm_start,
                        paragraph_pm_units,
                        list_state,
                    );
                    pm_cursor = paragraph_pm_start + u64::from(paragraph_pm_units) + 2;
                    blocks.extend(paragraph_blocks);
                    // S5a: a pilcrow carrying section properties ends a section —
                    // emit the section-break block right after its paragraph,
                    // exactly like the PM path (`toLayoutBlocks`).
                    let values = pilcrow_values(&pilcrow, txn);
                    if let Some(section_break) = section_break_block(&values, &mut section_margins)
                    {
                        blocks.push(LayoutBlock::SectionBreak(section_break));
                    }
                    paragraph_runs = Vec::new();
                    paragraph_drawings = Vec::new();
                    story_index += 1;
                    paragraph_start = story_index;
                    paragraph_pm_start = pm_cursor;
                    paragraph_pm_units = 0;
                    at_block_boundary = true;
                }
                Out::YMap(table)
                    if shared_map_string(&table, txn, "_kind").as_deref() == Some("table") =>
                {
                    if !paragraph_runs.is_empty()
                        || !paragraph_drawings.is_empty()
                        || paragraph_start != story_index
                    {
                        return Err(BridgeError::MalformedTable {
                            story: story_id.to_owned(),
                            index: story_index,
                            detail: "table embed interrupts paragraph content".to_owned(),
                        });
                    }
                    let (table, node_size) = lower_table(
                        &table,
                        txn,
                        story_id,
                        story_index,
                        pm_cursor,
                        env,
                        active_stories,
                        list_state,
                    )?;
                    blocks.push(LayoutBlock::Table(table));
                    story_index += 1;
                    paragraph_start = story_index;
                    pm_cursor += node_size;
                    paragraph_pm_start = pm_cursor;
                    paragraph_pm_units = 0;
                    at_block_boundary = true;
                }
                Out::YMap(page_break)
                    if matches!(
                        shared_map_string(&page_break, txn, "_kind").as_deref(),
                        Some("pageBreak" | "columnBreak")
                    ) =>
                {
                    if !at_block_boundary
                        || !paragraph_runs.is_empty()
                        || !paragraph_drawings.is_empty()
                    {
                        return Err(BridgeError::UnsupportedEmbed {
                            story: story_id.to_owned(),
                            index: story_index,
                        });
                    }
                    let kind = shared_map_string(&page_break, txn, "_kind").unwrap_or_default();
                    let id = BlockId::Str(format!("{story_id}:{kind}:{story_index}"));
                    if kind == "columnBreak" {
                        blocks.push(LayoutBlock::ColumnBreak(ColumnBreakBlock {
                            sdt_groups: None,
                            id,
                            pm_start: Some(pm_cursor as f64),
                            pm_end: Some((pm_cursor + 1) as f64),
                        }));
                    } else {
                        blocks.push(LayoutBlock::PageBreak(PageBreakBlock {
                            sdt_groups: None,
                            id,
                            pm_start: Some(pm_cursor as f64),
                            pm_end: Some((pm_cursor + 1) as f64),
                        }));
                    }
                    story_index += 1;
                    paragraph_start = story_index;
                    pm_cursor += 1;
                    paragraph_pm_start = pm_cursor;
                    paragraph_pm_units = 0;
                    at_block_boundary = true;
                }
                Out::YMap(block_sdt)
                    if shared_map_string(&block_sdt, txn, "_kind").as_deref()
                        == Some("blockSdt") =>
                {
                    if !at_block_boundary
                        || !paragraph_runs.is_empty()
                        || !paragraph_drawings.is_empty()
                    {
                        return Err(BridgeError::UnsupportedEmbed {
                            story: story_id.to_owned(),
                            index: story_index,
                        });
                    }
                    let Some(child_story) = shared_map_string(&block_sdt, txn, "story") else {
                        return Err(BridgeError::UnsupportedEmbed {
                            story: story_id.to_owned(),
                            index: story_index,
                        });
                    };
                    let group = lower_sdt_group(&block_sdt, txn, pm_cursor as i64);
                    let (mut child_blocks, content_size) = lower_story(
                        txn,
                        &child_story,
                        env,
                        pm_cursor + 1,
                        active_stories,
                        list_state,
                    )?;
                    stamp_sdt_group(&mut child_blocks, group);
                    blocks.extend(child_blocks);
                    story_index += 1;
                    paragraph_start = story_index;
                    pm_cursor += content_size + 2;
                    paragraph_pm_start = pm_cursor;
                    paragraph_pm_units = 0;
                    at_block_boundary = true;
                }
                Out::YMap(note_ref)
                    if shared_map_string(&note_ref, txn, "_kind").as_deref() == Some("noteRef") =>
                {
                    let footnote_id = shared_any(&note_ref, txn, "footnoteRefId")
                        .as_ref()
                        .and_then(|value| note_ref_id(value, env));
                    let endnote_id = shared_any(&note_ref, txn, "endnoteRefId")
                        .as_ref()
                        .and_then(|value| note_ref_id(value, env));
                    let Some(id) = footnote_id.or(endnote_id) else {
                        return Err(BridgeError::UnsupportedEmbed {
                            story: story_id.to_owned(),
                            index: story_index,
                        });
                    };
                    let mut formatting = lower_run_formatting(attributes, env);
                    formatting.superscript = Some(true);
                    if footnote_id.is_some() {
                        formatting.footnote_ref_id = Some(id);
                    } else {
                        formatting.endnote_ref_id = Some(id);
                    }
                    let mut comment_ids: Vec<f64> = comments
                        .iter()
                        .filter(|interval| {
                            interval.start <= story_index && story_index + 1 <= interval.end
                        })
                        .map(|interval| interval.id)
                        .collect();
                    comment_ids.sort_by(f64::total_cmp);
                    comment_ids.dedup_by(|a, b| a.total_cmp(b).is_eq());
                    if !comment_ids.is_empty() {
                        formatting.comment_ids = Some(comment_ids);
                    }
                    let label = note_ref_label(id);
                    let label_width = utf16_len(&label);
                    paragraph_runs.push(RawRun {
                        kind: RawRunKind::Text(label),
                        formatting,
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: paragraph_pm_units,
                        pm_end: paragraph_pm_units + label_width,
                        inline_sdt_widget: None,
                    });
                    story_index += 1;
                    paragraph_pm_units += label_width;
                    at_block_boundary = false;
                }
                Out::YMap(field)
                    if shared_map_string(&field, txn, "_kind").as_deref() == Some("field") =>
                {
                    let field_type = shared_map_string(&field, txn, "fieldType")
                        .unwrap_or_else(|| "OTHER".to_owned());
                    let mapped_type = match field_type.as_str() {
                        "PAGE" | "NUMPAGES" | "DATE" | "TIME" => field_type.clone(),
                        _ => "OTHER".to_owned(),
                    };
                    paragraph_runs.push(RawRun {
                        kind: RawRunKind::Field {
                            field_type: mapped_type.clone(),
                            raw_type: (field_type != mapped_type).then_some(field_type),
                            instruction: shared_map_string(&field, txn, "instruction")
                                .filter(|value| !value.is_empty()),
                            fallback: Some(
                                shared_map_string(&field, txn, "displayText").unwrap_or_default(),
                            ),
                        },
                        formatting: lower_run_formatting(attributes, env),
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: paragraph_pm_units,
                        pm_end: paragraph_pm_units + 1,
                        inline_sdt_widget: None,
                    });
                    story_index += 1;
                    paragraph_pm_units += 1;
                    at_block_boundary = false;
                }
                Out::YMap(line_break)
                    if shared_map_string(&line_break, txn, "_kind").as_deref() == Some("break") =>
                {
                    paragraph_runs.push(RawRun {
                        kind: RawRunKind::LineBreak,
                        formatting: RunFormatting::default(),
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: paragraph_pm_units,
                        pm_end: paragraph_pm_units + 1,
                        inline_sdt_widget: None,
                    });
                    story_index += 1;
                    paragraph_pm_units += 1;
                    at_block_boundary = false;
                }
                Out::YMap(image)
                    if shared_map_string(&image, txn, "_kind").as_deref() == Some("image") =>
                {
                    let formatting = lower_run_formatting(attributes, env);
                    paragraph_runs.push(RawRun {
                        kind: RawRunKind::Image(lower_image_run(&image, txn, &formatting, env)),
                        formatting,
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: paragraph_pm_units,
                        pm_end: paragraph_pm_units + 1,
                        inline_sdt_widget: None,
                    });
                    story_index += 1;
                    paragraph_pm_units += 1;
                    at_block_boundary = false;
                }
                Out::YMap(math)
                    if shared_map_string(&math, txn, "_kind").as_deref() == Some("math") =>
                {
                    let text = shared_map_string(&math, txn, "plainText")
                        .filter(|value| !value.is_empty())
                        .unwrap_or_else(|| "[equation]".to_owned());
                    paragraph_runs.push(RawRun {
                        kind: RawRunKind::Text(text),
                        formatting: RunFormatting {
                            italic: Some(true),
                            font_family: Some("Cambria Math".to_owned()),
                            // Sentinel consumed by `stamp_logical_order`: the
                            // PM math fallback omits logicalOrder.
                            logical_order: Some(u64::MAX),
                            ..RunFormatting::default()
                        },
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: paragraph_pm_units,
                        pm_end: paragraph_pm_units + 1,
                        inline_sdt_widget: None,
                    });
                    story_index += 1;
                    paragraph_pm_units += 1;
                    at_block_boundary = false;
                }
                Out::YMap(sdt)
                    if shared_map_string(&sdt, txn, "_kind").as_deref() == Some("sdt") =>
                {
                    let node_size = lower_inline_sdt(
                        &sdt,
                        txn,
                        env,
                        story_index,
                        paragraph_pm_start,
                        paragraph_pm_units,
                        None,
                        &mut paragraph_runs,
                    );
                    story_index += 1;
                    paragraph_pm_units += node_size;
                    at_block_boundary = false;
                }
                Out::YMap(shape)
                    if shared_map_string(&shape, txn, "_kind").as_deref() == Some("shape") =>
                {
                    let pm_offset = paragraph_pm_units;
                    let Some(block) = lower_shape_block(
                        &shape,
                        txn,
                        paragraph_pm_start + 1 + u64::from(pm_offset),
                        env,
                    ) else {
                        return Err(BridgeError::UnsupportedEmbed {
                            story: story_id.to_owned(),
                            index: story_index,
                        });
                    };
                    paragraph_drawings.push(DrawingMarker {
                        pm_offset,
                        block: LayoutBlock::Shape(block),
                    });
                    story_index += 1;
                    paragraph_pm_units += 1;
                    at_block_boundary = false;
                }
                Out::YMap(chart)
                    if shared_map_string(&chart, txn, "_kind").as_deref() == Some("chart") =>
                {
                    let pm_offset = paragraph_pm_units;
                    let Some(block) = lower_chart_block(
                        &chart,
                        txn,
                        paragraph_pm_start + 1 + u64::from(pm_offset),
                        env,
                    ) else {
                        return Err(BridgeError::UnsupportedEmbed {
                            story: story_id.to_owned(),
                            index: story_index,
                        });
                    };
                    paragraph_drawings.push(DrawingMarker {
                        pm_offset,
                        block: LayoutBlock::Chart(block),
                    });
                    story_index += 1;
                    paragraph_pm_units += 1;
                    at_block_boundary = false;
                }
                _ => {
                    return Err(BridgeError::UnsupportedEmbed {
                        story: story_id.to_owned(),
                        index: story_index,
                    });
                }
            }
        }

        if !at_block_boundary
            || !paragraph_runs.is_empty()
            || !paragraph_drawings.is_empty()
            || paragraph_start != story_index
        {
            return Err(BridgeError::UnterminatedStory(story_id.to_owned()));
        }

        Ok((blocks, pm_cursor - pm_base))
    })();
    active_stories.remove(story_id);
    result
}

/// Exact migration-spelling alias used by the build specification.
#[allow(non_snake_case)]
pub fn yrsDocToLayoutBlocks(
    doc: &EditingDoc,
    story_id: &str,
    env: &RenderEnv,
) -> Result<Vec<LayoutBlock>, BridgeError> {
    yrs_doc_to_layout_blocks(doc, story_id, env)
}

fn malformed_table(story: &str, index: u32, detail: impl Into<String>) -> BridgeError {
    BridgeError::MalformedTable {
        story: story.to_owned(),
        index,
        detail: detail.into(),
    }
}

fn shared_any<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<Any> {
    match map.get(txn, key) {
        Some(Out::Any(value)) => Some(value),
        _ => None,
    }
}

fn shared_map_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<String> {
    shared_any(map, txn, key)
        .as_ref()
        .and_then(any_str)
        .map(str::to_owned)
}

fn note_ref_id(value: &Any, env: &RenderEnv) -> Option<f64> {
    match value {
        Any::String(value) => Some(numeric_id(value, env)),
        _ => any_number(value),
    }
}

fn note_ref_label(id: f64) -> String {
    if id.fract() == 0.0 {
        format!("{id:.0}")
    } else {
        id.to_string()
    }
}

#[allow(clippy::too_many_arguments)]
fn lower_table<T: ReadTxn>(
    table: &MapRef,
    txn: &T,
    parent_story: &str,
    story_index: u32,
    pm_start: u64,
    env: &RenderEnv,
    active_stories: &mut BTreeSet<String>,
    list_state: &mut ListState,
) -> Result<(TableBlock, u64), BridgeError> {
    let tbl_pr_value = shared_any(table, txn, "tblPr")
        .ok_or_else(|| malformed_table(parent_story, story_index, "missing tblPr"))?;
    let tbl_pr = any_map(&tbl_pr_value)
        .ok_or_else(|| malformed_table(parent_story, story_index, "tblPr must be a map"))?;
    let grid_value = shared_any(table, txn, "grid")
        .ok_or_else(|| malformed_table(parent_story, story_index, "missing grid"))?;
    let grid = match &grid_value {
        Any::Array(values) => values,
        _ => {
            return Err(malformed_table(
                parent_story,
                story_index,
                "grid must be an array",
            ));
        }
    };
    let rows_value = shared_any(table, txn, "rows")
        .ok_or_else(|| malformed_table(parent_story, story_index, "missing rows"))?;
    let row_values = match &rows_value {
        Any::Array(values) => values,
        _ => {
            return Err(malformed_table(
                parent_story,
                story_index,
                "rows must be an array",
            ));
        }
    };
    // A table's absolute story index changes whenever text is inserted before it. Cell story
    // names, however, are persistent document identities, so anchor the table and row IDs to the
    // first cell instead. This keeps retained display-list primitive identities stable when an
    // unrelated earlier paragraph changes length.
    let table_identity = row_values
        .iter()
        .filter_map(any_map)
        .filter_map(|row| row.get("cells"))
        .filter_map(|cells| match cells {
            Any::Array(cells) => cells.first(),
            _ => None,
        })
        .filter_map(any_map)
        .find_map(|cell| map_string(cell, "story"))
        .unwrap_or_else(|| story_index.to_string());
    let table_id = format!("{parent_story}:table:{table_identity}");

    let table_margins = tbl_pr.get("cellMargins").and_then(any_map);
    let mut rows = Vec::with_capacity(row_values.len());
    let mut row_pm_start = pm_start + 1;

    for (row_index, row_value) in row_values.iter().enumerate() {
        let row = any_map(row_value).ok_or_else(|| {
            malformed_table(
                parent_story,
                story_index,
                format!("row {row_index} must be a map"),
            )
        })?;
        let tr_pr = row.get("trPr").and_then(any_map).ok_or_else(|| {
            malformed_table(
                parent_story,
                story_index,
                format!("row {row_index} is missing trPr"),
            )
        })?;
        let cell_values = match row.get("cells") {
            Some(Any::Array(values)) => values,
            _ => {
                return Err(malformed_table(
                    parent_story,
                    story_index,
                    format!("row {row_index} cells must be an array"),
                ));
            }
        };

        let mut cells = Vec::with_capacity(cell_values.len());
        let mut cell_pm_start = row_pm_start + 1;
        for (cell_index, cell_value) in cell_values.iter().enumerate() {
            let cell = any_map(cell_value).ok_or_else(|| {
                malformed_table(
                    parent_story,
                    story_index,
                    format!("row {row_index} cell {cell_index} must be a map"),
                )
            })?;
            let tc_pr = cell.get("tcPr").and_then(any_map).ok_or_else(|| {
                malformed_table(
                    parent_story,
                    story_index,
                    format!("row {row_index} cell {cell_index} is missing tcPr"),
                )
            })?;
            let cell_story = map_string(cell, "story").ok_or_else(|| {
                malformed_table(
                    parent_story,
                    story_index,
                    format!("row {row_index} cell {cell_index} is missing story"),
                )
            })?;
            let (blocks, content_size) = lower_story(
                txn,
                &cell_story,
                env,
                cell_pm_start + 1,
                active_stories,
                list_state,
            )?;

            let width_value = map_number(tc_pr, "width");
            let width_type = map_string(tc_pr, "widthType");
            let width =
                width_value.filter(|value| *value != 0.0).and_then(|value| {
                    match width_type.as_deref() {
                        None | Some("dxa") | Some("auto") => Some(twips_to_pixels(value)),
                        _ => None,
                    }
                });
            cells.push(TableCell {
                id: BlockId::Str(cell_story),
                blocks,
                col_span: map_number(tc_pr, "colspan"),
                row_span: map_number(tc_pr, "rowspan"),
                width,
                width_value,
                width_type,
                preferred_width: None,
                grid_start: None,
                min_content_width: None,
                max_content_width: None,
                vertical_align: map_string(tc_pr, "verticalAlign"),
                background: map_string(tc_pr, "backgroundColor").map(|value| format!("#{value}")),
                borders: lower_cell_borders(tc_pr, env),
                padding: Some(lower_cell_padding(tc_pr, table_margins)),
                no_wrap: (map_bool(tc_pr, "noWrap") == Some(true)).then_some(true),
                tracked_marker: tc_pr.get("cellMarker").and_then(any_json),
            });
            cell_pm_start += content_size + 2;
        }

        let row_node_size = cell_pm_start + 1 - row_pm_start;
        let original = tr_pr.get("_originalFormatting").and_then(any_map);
        rows.push(TableRow {
            id: BlockId::Str(format!("{table_id}:r{row_index}")),
            cells,
            height: map_number(tr_pr, "height")
                .filter(|value| *value != 0.0)
                .map(twips_to_pixels),
            height_rule: map_string(tr_pr, "heightRule"),
            is_header: map_bool(tr_pr, "isHeader"),
            cant_split: original
                .and_then(|value| map_bool(value, "cantSplit"))
                .filter(|value| *value),
            grid_before: None,
            grid_after: None,
            width_before: None,
            width_after: None,
            tracked_ins: tr_pr
                .get("trIns")
                .and_then(|value| paragraph_revision_value(value, env)),
            tracked_del: tr_pr
                .get("trDel")
                .and_then(|value| paragraph_revision_value(value, env)),
        });
        row_pm_start += row_node_size;
    }

    let node_size = row_pm_start + 1 - pm_start;
    let column_widths: Vec<f64> = grid
        .iter()
        .filter_map(any_number)
        .map(twips_to_pixels)
        .collect();
    let original = tbl_pr.get("_originalFormatting").and_then(any_map);
    let indent = original
        .and_then(|value| value.get("indent"))
        .and_then(any_map)
        .filter(|value| map_string(value, "type").as_deref() == Some("dxa"))
        .and_then(|value| map_number(value, "value"))
        .filter(|value| *value != 0.0)
        .map(twips_to_pixels);
    let floating = tbl_pr
        .get("floating")
        .and_then(any_map)
        .map(lower_floating_table);

    Ok((
        TableBlock {
            sdt_groups: None,
            id: BlockId::Str(table_id),
            rows,
            column_widths: (!column_widths.is_empty()).then_some(column_widths),
            grid_widths: None,
            width: map_number(tbl_pr, "width"),
            width_type: map_string(tbl_pr, "widthType"),
            preferred_width: None,
            layout_mode: None,
            width_algorithm: None,
            style_cascade: None,
            background: None,
            justification: map_string(tbl_pr, "justification"),
            bidi: (map_bool(tbl_pr, "bidi") == Some(true)).then_some(true),
            indent,
            floating,
            pm_start: Some(pm_start as f64),
            pm_end: Some((pm_start + node_size) as f64),
        },
        node_size,
    ))
}

fn any_json(value: &Any) -> Option<Value> {
    let value = serde_json::to_value(value).ok()?;
    (!value.is_null()).then_some(value)
}

fn shared_values<T: ReadTxn>(map: &MapRef, txn: &T) -> std::collections::HashMap<String, Any> {
    map.iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_string(), value)),
            _ => None,
        })
        .collect()
}

fn attrs_from_any(value: Option<&Any>) -> Attrs {
    value
        .and_then(any_map)
        .into_iter()
        .flat_map(|map| map.iter())
        .map(|(key, value)| (Arc::<str>::from(key.as_str()), value.clone()))
        .collect()
}

fn constrain_to_page(width: f64, height: f64, env: &RenderEnv) -> (f64, f64) {
    let Some(limit) = env.page_content_height.filter(|limit| *limit > 0.0) else {
        return (width, height);
    };
    if height <= limit {
        return (width, height);
    }
    ((width * (limit / height)).round(), limit)
}

fn image_transform_metrics(
    width: f64,
    height: f64,
    transform: Option<&str>,
) -> (Option<f64>, Option<bool>, Option<bool>, Option<Value>) {
    let rotation = transform
        .and_then(|value| value.split_once("rotate("))
        .and_then(|(_, tail)| tail.split_once("deg)"))
        .and_then(|(value, _)| value.trim().parse::<f64>().ok())
        .filter(|value| value.is_finite())
        .unwrap_or(0.0)
        .rem_euclid(360.0);
    let radians = rotation.to_radians();
    let bounds_width = width * radians.cos().abs() + height * radians.sin().abs();
    let bounds_height = width * radians.sin().abs() + height * radians.cos().abs();
    let bounds = (rotation != 0.0).then(|| {
        serde_json::json!({
            "width": bounds_width,
            "height": bounds_height,
            "offsetX": (bounds_width - width) / 2.0,
            "offsetY": (bounds_height - height) / 2.0,
        })
    });
    (
        (rotation != 0.0).then_some(rotation),
        transform
            .is_some_and(|value| value.contains("scaleX(-1)"))
            .then_some(true),
        transform
            .is_some_and(|value| value.contains("scaleY(-1)"))
            .then_some(true),
        bounds,
    )
}

fn lower_image_values(
    values: &std::collections::HashMap<String, Any>,
    formatting: &RunFormatting,
    env: &RenderEnv,
) -> ImageRun {
    let (width, height) = constrain_to_page(
        map_number(values, "width")
            .filter(|value| *value != 0.0)
            .unwrap_or(100.0),
        map_number(values, "height")
            .filter(|value| *value != 0.0)
            .unwrap_or(100.0),
        env,
    );
    let transform = map_string(values, "transform");
    let (rotation_deg, flip_h, flip_v, rotation_bounds) =
        image_transform_metrics(width, height, transform.as_deref());
    let position = values
        .get("position")
        .and_then(any_json)
        .and_then(|value| serde_json::from_value::<ImageRunPosition>(value).ok());

    ImageRun {
        src: map_string(values, "src").unwrap_or_default(),
        width,
        height,
        alt: map_string(values, "alt"),
        transform,
        position,
        wrap_type: map_string(values, "wrapType"),
        display_mode: map_string(values, "displayMode"),
        css_float: map_string(values, "cssFloat"),
        dist_top: map_number(values, "distTop"),
        dist_bottom: map_number(values, "distBottom"),
        dist_left: map_number(values, "distLeft"),
        dist_right: map_number(values, "distRight"),
        crop_top: map_number(values, "cropTop"),
        crop_right: map_number(values, "cropRight"),
        crop_bottom: map_number(values, "cropBottom"),
        crop_left: map_number(values, "cropLeft"),
        opacity: map_number(values, "opacity"),
        rotation_deg,
        flip_h,
        flip_v,
        rotation_bounds,
        wrap_text: None,
        wrap_polygon: None,
        allow_overlap: None,
        layout_in_cell: None,
        effect_extent: None,
        effects: None,
        outline: None,
        decorative: None,
        hyperlink: None,
        is_insertion: formatting.is_insertion,
        is_deletion: formatting.is_deletion,
        change_author: formatting.change_author.clone(),
        change_date: formatting.change_date.clone(),
        change_revision_id: formatting.change_revision_id,
        pm_start: None,
        pm_end: None,
    }
}

fn lower_image_run<T: ReadTxn>(
    image: &MapRef,
    txn: &T,
    formatting: &RunFormatting,
    env: &RenderEnv,
) -> ImageRun {
    lower_image_values(&shared_values(image, txn), formatting, env)
}

fn preset_geometry(shape_type: &str) -> Option<Vec<Value>> {
    let command = |value: Value| value;
    let close = || serde_json::json!({ "type": "close" });
    Some(match shape_type {
        "rect" => vec![
            command(serde_json::json!({"type":"move","x":0,"y":0})),
            command(serde_json::json!({"type":"line","x":1,"y":0})),
            command(serde_json::json!({"type":"line","x":1,"y":1})),
            command(serde_json::json!({"type":"line","x":0,"y":1})),
            close(),
        ],
        "roundRect" => {
            let r = 1.0 / 6.0;
            vec![
                serde_json::json!({"type":"move","x":r,"y":0}),
                serde_json::json!({"type":"line","x":1.0-r,"y":0}),
                serde_json::json!({"type":"quad","cpx":1,"cpy":0,"x":1,"y":r}),
                serde_json::json!({"type":"line","x":1,"y":1.0-r}),
                serde_json::json!({"type":"quad","cpx":1,"cpy":1,"x":1.0-r,"y":1}),
                serde_json::json!({"type":"line","x":r,"y":1}),
                serde_json::json!({"type":"quad","cpx":0,"cpy":1,"x":0,"y":1.0-r}),
                serde_json::json!({"type":"line","x":0,"y":r}),
                serde_json::json!({"type":"quad","cpx":0,"cpy":0,"x":r,"y":0}),
                close(),
            ]
        }
        "ellipse" => {
            let k = 0.552_284_749_830_793_6 / 2.0;
            vec![
                serde_json::json!({"type":"move","x":1,"y":0.5}),
                serde_json::json!({"type":"cubic","cp1x":1,"cp1y":0.5+k,"cp2x":0.5+k,"cp2y":1,"x":0.5,"y":1}),
                serde_json::json!({"type":"cubic","cp1x":0.5-k,"cp1y":1,"cp2x":0,"cp2y":0.5+k,"x":0,"y":0.5}),
                serde_json::json!({"type":"cubic","cp1x":0,"cp1y":0.5-k,"cp2x":0.5-k,"cp2y":0,"x":0.5,"y":0}),
                serde_json::json!({"type":"cubic","cp1x":0.5+k,"cp1y":0,"cp2x":1,"cp2y":0.5-k,"x":1,"y":0.5}),
                close(),
            ]
        }
        "line" | "straightConnector1" => vec![
            serde_json::json!({"type":"move","x":0,"y":0}),
            serde_json::json!({"type":"line","x":1,"y":1}),
        ],
        "triangle" | "isosTriangle" => vec![
            serde_json::json!({"type":"move","x":0.5,"y":0}),
            serde_json::json!({"type":"line","x":1,"y":1}),
            serde_json::json!({"type":"line","x":0,"y":1}),
            close(),
        ],
        _ => return None,
    })
}

fn shape_fill(values: &std::collections::HashMap<String, Any>) -> Option<Value> {
    let fill_type = map_string(values, "fillType").unwrap_or_else(|| "solid".to_owned());
    let color = map_string(values, "fillColor");
    let mut fill = JsonMap::new();
    fill.insert("type".to_owned(), Value::String(fill_type.clone()));
    if let Some(color) = color {
        fill.insert("color".to_owned(), Value::String(color));
    }
    if fill_type == "gradient" {
        if let Some(value) = map_string(values, "gradientType") {
            fill.insert("gradientType".to_owned(), Value::String(value));
        }
        if let Some(value) = map_number(values, "gradientAngle") {
            fill.insert("gradientAngle".to_owned(), Value::from(value));
        }
        if let Some(stops) = map_string(values, "gradientStops")
            .and_then(|value| serde_json::from_str::<Value>(&value).ok())
            .filter(Value::is_array)
        {
            fill.insert("gradientStops".to_owned(), stops);
        }
    }
    Some(Value::Object(fill))
}

fn shape_stroke(values: &std::collections::HashMap<String, Any>) -> Option<Value> {
    let width = map_number(values, "outlineWidth");
    let color = map_string(values, "outlineColor");
    let dash = map_string(values, "outlineStyle");
    if width.is_none() && color.is_none() && dash.is_none() {
        return None;
    }
    let mut stroke = JsonMap::new();
    if let Some(value) = color {
        stroke.insert("color".to_owned(), Value::String(value));
    }
    if let Some(value) = width {
        stroke.insert("width".to_owned(), Value::from(value));
    }
    if let Some(value) = dash {
        stroke.insert("dash".to_owned(), Value::String(value));
    }
    Some(Value::Object(stroke))
}

fn shape_transform(values: &std::collections::HashMap<String, Any>) -> Option<Value> {
    let rotation = map_number(values, "rotation").or_else(|| {
        map_string(values, "transform")
            .as_deref()
            .and_then(|value| image_transform_metrics(1.0, 1.0, Some(value)).0)
    });
    let flip_h = map_bool(values, "flipH") == Some(true)
        || map_string(values, "transform").is_some_and(|value| value.contains("scaleX(-1)"));
    let flip_v = map_bool(values, "flipV") == Some(true)
        || map_string(values, "transform").is_some_and(|value| value.contains("scaleY(-1)"));
    if rotation.is_none() && !flip_h && !flip_v {
        return None;
    }
    let mut transform = JsonMap::new();
    if let Some(value) = rotation {
        transform.insert("rotation".to_owned(), Value::from(value));
    }
    if flip_h {
        transform.insert("flipH".to_owned(), Value::Bool(true));
    }
    if flip_v {
        transform.insert("flipV".to_owned(), Value::Bool(true));
    }
    Some(Value::Object(transform))
}

fn lower_shape_block<T: ReadTxn>(
    shape: &MapRef,
    txn: &T,
    pm_start: u64,
    env: &RenderEnv,
) -> Option<ShapeBlock> {
    let values = shared_values(shape, txn);
    if let Some(mut block) = map_string(&values, "layoutBlockJson")
        .and_then(|value| serde_json::from_str::<ShapeBlock>(&value).ok())
    {
        (block.width, block.height) = constrain_to_page(block.width, block.height, env);
        block.doc_start = Some(pm_start as f64);
        block.doc_end = Some((pm_start + 1) as f64);
        block.pm_start = Some(pm_start as f64);
        block.pm_end = Some((pm_start + 1) as f64);
        return Some(block);
    }
    let shape_type = map_string(&values, "shapeType").unwrap_or_else(|| "rect".to_owned());
    let geometry_path = values
        .get("geometryPath")
        .and_then(any_json)
        .and_then(|value| value.as_array().cloned())
        .filter(|value| !value.is_empty())
        .or_else(|| preset_geometry(&shape_type))?;
    let (width, height) = constrain_to_page(
        map_number(&values, "width")
            .filter(|value| *value != 0.0)
            .unwrap_or(100.0),
        map_number(&values, "height")
            .filter(|value| *value != 0.0)
            .unwrap_or(80.0),
        env,
    );
    Some(ShapeBlock {
        sdt_groups: None,
        id: BlockId::Str(format!("shape:{pm_start}")),
        shape_type,
        geometry_path,
        fill: shape_fill(&values),
        stroke: shape_stroke(&values),
        transform: shape_transform(&values),
        width,
        height,
        x: None,
        y: None,
        inner_text: None,
        inner_measures: None,
        children: Vec::new(),
        scene: None,
        effects: None,
        text_body_properties: None,
        position: None,
        wrap_type: None,
        wrap_text: None,
        relative_height: None,
        behind_doc: None,
        decorative: None,
        title: None,
        description: None,
        doc_start: Some(pm_start as f64),
        doc_end: Some((pm_start + 1) as f64),
        pm_start: Some(pm_start as f64),
        pm_end: Some((pm_start + 1) as f64),
    })
}

fn lower_chart_block<T: ReadTxn>(
    chart: &MapRef,
    txn: &T,
    pm_start: u64,
    env: &RenderEnv,
) -> Option<ChartBlock> {
    let values = shared_values(chart, txn);
    let chart = map_string(&values, "chartJson")
        .and_then(|value| serde_json::from_str::<Value>(&value).ok())?;
    let (width, height) = constrain_to_page(
        map_number(&values, "width")
            .filter(|value| *value != 0.0)
            .unwrap_or(320.0),
        map_number(&values, "height")
            .filter(|value| *value != 0.0)
            .unwrap_or(220.0),
        env,
    );
    Some(ChartBlock {
        sdt_groups: None,
        id: BlockId::Str(format!("chart:{pm_start}")),
        chart,
        width,
        height,
        position: None,
        wrap_type: None,
        wrap_text: None,
        relative_height: None,
        behind_doc: None,
        doc_start: Some(pm_start as f64),
        doc_end: Some((pm_start + 1) as f64),
        pm_start: Some(pm_start as f64),
        pm_end: Some((pm_start + 1) as f64),
    })
}

fn repeating_sdt_item(raw: Option<&str>) -> Option<bool> {
    let raw = raw?;
    let marker = "<w15:repeatingSectionItem";
    let tail = raw.split_once(marker)?.1;
    tail.chars()
        .next()
        .is_some_and(|ch| ch.is_whitespace() || ch == '/' || ch == '>')
        .then_some(true)
}

fn sdt_group_from_values(values: &std::collections::HashMap<String, Any>, pos: i64) -> SdtGroup {
    SdtGroup {
        id: format!("sdt@{pos}"),
        sdt_type: map_string(values, "sdtType").unwrap_or_else(|| "richText".to_owned()),
        tag: map_string(values, "tag"),
        alias: map_string(values, "alias"),
        lock: map_string(values, "lock"),
        checked: authored_checkbox_value(values).or_else(|| map_bool(values, "checked")),
        bound: values.get("dataBinding").map(|_| true),
        repeating_item: repeating_sdt_item(map_string(values, "rawPropertiesXml").as_deref()),
        control_id: map_number(values, "id").map(|value| value as i64),
        pos: None,
        control_state: None,
        properties: None,
    }
}

fn lower_sdt_group<T: ReadTxn>(sdt: &MapRef, txn: &T, pos: i64) -> SdtGroup {
    sdt_group_from_values(&shared_values(sdt, txn), pos)
}

fn stamp_sdt_group(blocks: &mut [LayoutBlock], group: SdtGroup) {
    for block in blocks {
        let groups = match block {
            LayoutBlock::Paragraph(block) => &mut block.sdt_groups,
            LayoutBlock::Table(block) => &mut block.sdt_groups,
            LayoutBlock::Image(block) => &mut block.sdt_groups,
            LayoutBlock::Shape(block) => &mut block.sdt_groups,
            LayoutBlock::Chart(block) => &mut block.sdt_groups,
            LayoutBlock::TextBox(block) => &mut block.sdt_groups,
            LayoutBlock::SectionBreak(block) => &mut block.sdt_groups,
            LayoutBlock::PageBreak(block) => &mut block.sdt_groups,
            LayoutBlock::ColumnBreak(block) => &mut block.sdt_groups,
            LayoutBlock::Unsupported => continue,
        };
        groups.get_or_insert_with(Vec::new).insert(0, group.clone());
    }
}

fn inline_checkbox_widget(
    values: &std::collections::HashMap<String, Any>,
    pos: i64,
) -> Option<Value> {
    if map_string(values, "sdtType").as_deref() != Some("checkbox") {
        return None;
    }
    if matches!(
        map_string(values, "lock").as_deref(),
        Some("contentLocked" | "sdtContentLocked")
    ) || values.contains_key("dataBinding")
    {
        return None;
    }
    let mut widget = JsonMap::from_iter([
        ("kind".to_owned(), Value::String("checkbox".to_owned())),
        ("groupId".to_owned(), Value::String(format!("sdt@{pos}"))),
        ("pos".to_owned(), Value::from(pos)),
    ]);
    if let Some(value) = map_string(values, "tag") {
        widget.insert("tag".to_owned(), Value::String(value));
    }
    if let Some(value) = map_string(values, "alias") {
        widget.insert("alias".to_owned(), Value::String(value));
    }
    if let Some(value) = map_number(values, "id") {
        widget.insert("controlId".to_owned(), Value::from(value as i64));
    }
    if let Some(value) = authored_checkbox_value(values).or_else(|| map_bool(values, "checked")) {
        widget.insert("checked".to_owned(), Value::Bool(value));
    }
    Some(Value::Object(widget))
}

fn authored_checkbox_value(values: &std::collections::HashMap<String, Any>) -> Option<bool> {
    let value = values.get("value").and_then(any_map)?;
    (map_string(value, "kind").as_deref() == Some("checkbox"))
        .then(|| map_bool(value, "checked"))
        .flatten()
}

#[allow(clippy::too_many_arguments)]
fn lower_inline_sdt<T: ReadTxn>(
    sdt: &MapRef,
    txn: &T,
    env: &RenderEnv,
    story_index: u32,
    paragraph_pm_start: u64,
    node_pm_start: u32,
    inherited_widget: Option<Value>,
    runs: &mut Vec<RawRun>,
) -> u32 {
    lower_inline_sdt_values(
        &shared_values(sdt, txn),
        env,
        story_index,
        paragraph_pm_start,
        node_pm_start,
        inherited_widget,
        runs,
    )
}

#[allow(clippy::too_many_arguments)]
fn lower_inline_sdt_values(
    values: &std::collections::HashMap<String, Any>,
    env: &RenderEnv,
    story_index: u32,
    paragraph_pm_start: u64,
    node_pm_start: u32,
    inherited_widget: Option<Value>,
    runs: &mut Vec<RawRun>,
) -> u32 {
    let absolute_pos = paragraph_pm_start + 1 + u64::from(node_pm_start);
    let widget = inline_checkbox_widget(values, absolute_pos as i64).or(inherited_widget);
    let mut content_size = 0_u32;
    let content = match values.get("content") {
        Some(Any::Array(content)) => content,
        _ => return 2,
    };

    for child in content.iter() {
        let Some(child) = any_map(child) else {
            continue;
        };
        let kind = map_string(child, "kind").unwrap_or_default();
        let attrs = attrs_from_any(child.get("attrs"));
        let formatting = lower_run_formatting(Some(&attrs), env);
        let child_pm_start = node_pm_start + 1 + content_size;
        let payload = child.get("payload").and_then(any_map);
        let child_size = match kind.as_str() {
            "text" => {
                let text = map_string(child, "text").unwrap_or_default();
                let width = utf16_len(&text);
                if width > 0 {
                    runs.push(RawRun {
                        kind: RawRunKind::Text(text),
                        formatting,
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: child_pm_start,
                        pm_end: child_pm_start + width,
                        inline_sdt_widget: widget.clone(),
                    });
                }
                width
            }
            "tab" => {
                runs.push(RawRun {
                    kind: RawRunKind::Tab,
                    formatting,
                    story_start: story_index,
                    story_end: story_index + 1,
                    pm_start: child_pm_start,
                    pm_end: child_pm_start + 1,
                    inline_sdt_widget: None,
                });
                1
            }
            "break" => {
                runs.push(RawRun {
                    kind: RawRunKind::LineBreak,
                    formatting: RunFormatting::default(),
                    story_start: story_index,
                    story_end: story_index + 1,
                    pm_start: child_pm_start,
                    pm_end: child_pm_start + 1,
                    inline_sdt_widget: None,
                });
                1
            }
            "image" => {
                if let Some(payload) = payload {
                    runs.push(RawRun {
                        kind: RawRunKind::Image(lower_image_values(payload, &formatting, env)),
                        formatting,
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: child_pm_start,
                        pm_end: child_pm_start + 1,
                        inline_sdt_widget: None,
                    });
                }
                1
            }
            "field" => {
                let field_type = payload
                    .and_then(|payload| map_string(payload, "fieldType"))
                    .unwrap_or_else(|| "OTHER".to_owned());
                let mapped_type = match field_type.as_str() {
                    "PAGE" | "NUMPAGES" | "DATE" | "TIME" => field_type.clone(),
                    _ => "OTHER".to_owned(),
                };
                runs.push(RawRun {
                    kind: RawRunKind::Field {
                        field_type: mapped_type.clone(),
                        raw_type: (field_type != mapped_type).then_some(field_type),
                        instruction: payload
                            .and_then(|payload| map_string(payload, "instruction"))
                            .filter(|value| !value.is_empty()),
                        fallback: Some(
                            payload
                                .and_then(|payload| map_string(payload, "displayText"))
                                .unwrap_or_default(),
                        ),
                    },
                    formatting,
                    story_start: story_index,
                    story_end: story_index + 1,
                    pm_start: child_pm_start,
                    pm_end: child_pm_start + 1,
                    inline_sdt_widget: None,
                });
                1
            }
            "math" => {
                let text = payload
                    .and_then(|payload| map_string(payload, "plainText"))
                    .filter(|value| !value.is_empty())
                    .unwrap_or_else(|| "[equation]".to_owned());
                runs.push(RawRun {
                    kind: RawRunKind::Text(text),
                    formatting: RunFormatting {
                        italic: Some(true),
                        font_family: Some("Cambria Math".to_owned()),
                        logical_order: Some(u64::MAX),
                        ..RunFormatting::default()
                    },
                    story_start: story_index,
                    story_end: story_index + 1,
                    pm_start: child_pm_start,
                    pm_end: child_pm_start + 1,
                    inline_sdt_widget: None,
                });
                1
            }
            "noteRef" => {
                let footnote_id = payload
                    .and_then(|payload| payload.get("footnoteRefId"))
                    .and_then(|value| note_ref_id(value, env));
                let endnote_id = payload
                    .and_then(|payload| payload.get("endnoteRefId"))
                    .and_then(|value| note_ref_id(value, env));
                if let Some(id) = footnote_id.or(endnote_id) {
                    let mut formatting = formatting;
                    formatting.superscript = Some(true);
                    if footnote_id.is_some() {
                        formatting.footnote_ref_id = Some(id);
                    } else {
                        formatting.endnote_ref_id = Some(id);
                    }
                    let label = note_ref_label(id);
                    let width = utf16_len(&label);
                    runs.push(RawRun {
                        kind: RawRunKind::Text(label),
                        formatting,
                        story_start: story_index,
                        story_end: story_index + 1,
                        pm_start: child_pm_start,
                        pm_end: child_pm_start + width,
                        inline_sdt_widget: widget.clone(),
                    });
                    width
                } else {
                    1
                }
            }
            "sdt" => payload.map_or(2, |payload| {
                lower_inline_sdt_values(
                    payload,
                    env,
                    story_index,
                    paragraph_pm_start,
                    child_pm_start,
                    widget.clone(),
                    runs,
                )
            }),
            // Shape/chart children nested inside an inline SDT are omitted by
            // the current PM run lowering too; all other leaves occupy one PM
            // position even when they have no layout run.
            _ => 1,
        };
        content_size += child_size;
    }
    content_size + 2
}

fn lower_floating_table(value: &std::collections::HashMap<String, Any>) -> FloatingTablePosition {
    let px = |key| map_number(value, key).map(twips_to_pixels);
    FloatingTablePosition {
        horz_anchor: map_string(value, "horzAnchor"),
        tblp_x: px("tblpX"),
        tblp_x_spec: map_string(value, "tblpXSpec"),
        vert_anchor: map_string(value, "vertAnchor"),
        tblp_y: px("tblpY"),
        tblp_y_spec: map_string(value, "tblpYSpec"),
        top_from_text: px("topFromText"),
        right_from_text: px("rightFromText"),
        bottom_from_text: px("bottomFromText"),
        left_from_text: px("leftFromText"),
    }
}

fn lower_cell_padding(
    tc_pr: &std::collections::HashMap<String, Any>,
    table_margins: Option<&std::collections::HashMap<String, Any>>,
) -> BoxEdges {
    let cell_margins = tc_pr.get("margins").and_then(any_map);
    let side = |key: &str| {
        cell_margins
            .and_then(|margins| map_number(margins, key))
            .map(twips_to_pixels)
            .filter(|value| *value > 0.0)
            .or_else(|| {
                table_margins
                    .and_then(|margins| map_number(margins, key))
                    .map(twips_to_pixels)
                    .filter(|value| *value >= 0.0)
            })
            .unwrap_or(0.0)
    };
    BoxEdges {
        top: side("top"),
        right: side("right"),
        bottom: side("bottom"),
        left: side("left"),
    }
}

fn lower_cell_borders(
    tc_pr: &std::collections::HashMap<String, Any>,
    env: &RenderEnv,
) -> Option<CellBorders> {
    let borders = tc_pr.get("borders").and_then(any_map)?;
    let side = |key: &str| {
        borders
            .get(key)
            .and_then(any_map)
            .and_then(|border| lower_cell_border(border, env))
            .or_else(|| {
                Some(CellBorderSpec {
                    width: Some(0.0),
                    color: None,
                    style: Some("none".to_owned()),
                })
            })
    };
    Some(CellBorders {
        top: side("top"),
        right: side("right"),
        bottom: side("bottom"),
        left: side("left"),
    })
}

fn lower_cell_border(
    border: &std::collections::HashMap<String, Any>,
    env: &RenderEnv,
) -> Option<CellBorderSpec> {
    let style = map_string(border, "style")?;
    if matches!(style.as_str(), "none" | "nil") {
        return None;
    }
    let css_style = match style.as_str() {
        "double"
        | "triple"
        | "thinThickSmallGap"
        | "thickThinSmallGap"
        | "thinThickThinSmallGap"
        | "thinThickMediumGap"
        | "thickThinMediumGap"
        | "thinThickThinMediumGap"
        | "thinThickLargeGap"
        | "thickThinLargeGap"
        | "thinThickThinLargeGap"
        | "doubleWave" => "double",
        "dotted" | "dotDotDash" => "dotted",
        "dashed" | "dashSmallGap" | "dotDash" | "dashDotStroked" => "dashed",
        "threeDEmboss" => "ridge",
        "threeDEngrave" => "groove",
        "outset" => "outset",
        "inset" => "inset",
        _ => "solid",
    };
    let width = ((map_number(border, "size").unwrap_or(0.0) / 8.0) * 1.333)
        .round()
        .max(1.0);
    let color = border
        .get("color")
        .and_then(|value| resolve_color(value, env))
        .unwrap_or_else(|| "#000000".to_owned());
    Some(CellBorderSpec {
        width: Some(width),
        color: Some(color),
        style: Some(css_style.to_owned()),
    })
}

#[derive(Clone, Debug)]
struct CommentInterval {
    start: u32,
    end: u32,
    id: f64,
}

#[derive(Clone, Debug)]
enum RawRunKind {
    Text(String),
    Tab,
    Image(ImageRun),
    LineBreak,
    Field {
        field_type: String,
        raw_type: Option<String>,
        instruction: Option<String>,
        fallback: Option<String>,
    },
}

#[derive(Clone, Debug)]
struct RawRun {
    kind: RawRunKind,
    formatting: RunFormatting,
    story_start: u32,
    story_end: u32,
    /// PM offsets relative to the paragraph content start. Usually identical
    /// to story offsets; a one-unit `noteRef` displays its possibly multi-digit id.
    pm_start: u32,
    pm_end: u32,
    /// Checkbox chrome inherited from the nearest editable inline SDT.
    inline_sdt_widget: Option<Value>,
}

#[derive(Clone, Debug)]
struct DrawingMarker {
    /// PM offset relative to the original paragraph content start.
    pm_offset: u32,
    block: LayoutBlock,
}

fn resolve_comment_intervals<T: ReadTxn>(
    txn: &T,
    story_id: &str,
    env: &RenderEnv,
) -> Result<Vec<CommentInterval>, BridgeError> {
    let comments = txn
        .get_map(COMMENTS)
        .expect("comments root is declared by EditingDoc::new");
    let mut intervals = Vec::new();

    for (comment_id, value) in comments.iter(txn) {
        let Out::YMap(comment) = value else {
            continue;
        };
        let Some(Out::Any(Any::Array(anchors))) = comment.get(txn, "anchors") else {
            return Err(EditError::InvalidComment("anchors must be an array".into()).into());
        };
        for encoded in anchors.iter() {
            let anchor = decode_anchor(encoded)?;
            if anchor.story != story_id {
                continue;
            }
            let start = anchor.start.get_offset(txn).ok_or_else(|| {
                EditError::InvalidComment("start anchor no longer resolves".into())
            })?;
            let end = anchor
                .end
                .get_offset(txn)
                .ok_or_else(|| EditError::InvalidComment("end anchor no longer resolves".into()))?;
            if start.index < end.index {
                intervals.push(CommentInterval {
                    start: start.index,
                    end: end.index,
                    id: numeric_id(comment_id, env),
                });
            }
        }
    }

    intervals.sort_by(|a, b| {
        a.start
            .cmp(&b.start)
            .then(a.end.cmp(&b.end))
            .then(a.id.total_cmp(&b.id))
    });
    Ok(intervals)
}

fn push_text_chunks(
    runs: &mut Vec<RawRun>,
    text: &str,
    chunk_start: u32,
    attributes: Option<&Attrs>,
    comments: &[CommentInterval],
    env: &RenderEnv,
    chunk_pm_start: u32,
) {
    let chunk_end = chunk_start + utf16_len(text);
    let mut cuts = BTreeSet::from([chunk_start, chunk_end]);
    for interval in comments {
        if interval.start > chunk_start && interval.start < chunk_end {
            cuts.insert(interval.start);
        }
        if interval.end > chunk_start && interval.end < chunk_end {
            cuts.insert(interval.end);
        }
    }
    let mut offset = chunk_start;
    for ch in text.chars() {
        let next = offset + ch.len_utf16() as u32;
        if ch == '\t' || ch == '\u{000b}' {
            cuts.insert(offset);
            cuts.insert(next);
        }
        offset = next;
    }

    let cuts: Vec<u32> = cuts.into_iter().collect();
    for pair in cuts.windows(2) {
        let start = pair[0];
        let end = pair[1];
        if start == end {
            continue;
        }
        let value = utf16_slice(text, start - chunk_start, end - chunk_start);
        let kind = match value.as_str() {
            "\t" => RawRunKind::Tab,
            "\u{000b}" => RawRunKind::LineBreak,
            _ => RawRunKind::Text(value),
        };
        let mut formatting = lower_run_formatting(attributes, env);
        let mut comment_ids: Vec<f64> = comments
            .iter()
            .filter(|interval| interval.start <= start && end <= interval.end)
            .map(|interval| interval.id)
            .collect();
        comment_ids.sort_by(f64::total_cmp);
        comment_ids.dedup_by(|a, b| a.total_cmp(b).is_eq());
        if !comment_ids.is_empty() {
            formatting.comment_ids = Some(comment_ids);
        }
        runs.push(RawRun {
            kind,
            formatting,
            story_start: start,
            story_end: end,
            pm_start: chunk_pm_start + start - chunk_start,
            pm_end: chunk_pm_start + end - chunk_start,
            inline_sdt_widget: None,
        });
    }
}

#[allow(clippy::too_many_arguments)]
fn flush_paragraph_parts<T: ReadTxn>(
    mut raw_runs: Vec<RawRun>,
    drawings: Vec<DrawingMarker>,
    pilcrow: &MapRef,
    pilcrow_attributes: Option<&Attrs>,
    txn: &T,
    story_id: &str,
    env: &RenderEnv,
    paragraph_pm_start: u64,
    paragraph_pm_units: u32,
    list_state: &mut ListState,
) -> Vec<LayoutBlock> {
    if drawings.is_empty() {
        return vec![LayoutBlock::Paragraph(flush_paragraph(
            raw_runs,
            pilcrow,
            pilcrow_attributes,
            txn,
            story_id,
            env,
            paragraph_pm_start,
            paragraph_pm_units,
            list_state,
        ))];
    }

    // `convertParagraphWithShapes` splits the PM paragraph around direct
    // shape/chart children and omits empty paragraph slices. Reproduce that
    // exact geometry while retaining the original child PM positions.
    let mut blocks = Vec::new();
    let mut segment_start = 0_u32;
    for drawing in drawings {
        let split_at = raw_runs.partition_point(|run| run.pm_end <= drawing.pm_offset);
        let remaining = raw_runs.split_off(split_at);
        let mut segment = std::mem::replace(&mut raw_runs, remaining);
        if !segment.is_empty() {
            for run in &mut segment {
                run.pm_start -= segment_start;
                run.pm_end -= segment_start;
            }
            blocks.push(LayoutBlock::Paragraph(flush_paragraph(
                segment,
                pilcrow,
                pilcrow_attributes,
                txn,
                story_id,
                env,
                paragraph_pm_start + u64::from(segment_start),
                drawing.pm_offset - segment_start,
                list_state,
            )));
        }
        blocks.push(drawing.block);
        segment_start = drawing.pm_offset + 1;
    }

    if !raw_runs.is_empty() {
        for run in &mut raw_runs {
            run.pm_start -= segment_start;
            run.pm_end -= segment_start;
        }
        blocks.push(LayoutBlock::Paragraph(flush_paragraph(
            raw_runs,
            pilcrow,
            pilcrow_attributes,
            txn,
            story_id,
            env,
            paragraph_pm_start + u64::from(segment_start),
            paragraph_pm_units - segment_start,
            list_state,
        )));
    }
    blocks
}

#[allow(clippy::too_many_arguments)]
fn flush_paragraph<T: ReadTxn>(
    mut raw_runs: Vec<RawRun>,
    pilcrow: &MapRef,
    pilcrow_attributes: Option<&Attrs>,
    txn: &T,
    story_id: &str,
    env: &RenderEnv,
    paragraph_pm_start: u64,
    paragraph_pm_units: u32,
    list_state: &mut ListState,
) -> ParagraphBlock {
    let values = pilcrow_values(pilcrow, txn);
    let para_id = value_string(values.get("paraId")).unwrap_or_default();
    let generated_prefix = format!("{story_id}:p");
    let para_id_is_generated = para_id
        .strip_prefix(&generated_prefix)
        .is_some_and(|suffix| suffix.parse::<usize>().is_ok());
    let style_id = paragraph_style_id(&values);
    let defaults = paragraph_run_defaults(&values);

    for run in &mut raw_runs {
        apply_run_defaults(&mut run.formatting, &defaults);
        if style_id.as_deref().is_some_and(is_toc_style) {
            strip_toc_hyperlink_style(&mut run.formatting);
        }
    }
    let raw_runs = coalesce_runs(raw_runs);
    let mut runs: Vec<Run> = raw_runs
        .into_iter()
        .map(|raw| raw_run_to_layout(raw, paragraph_pm_start))
        .collect();
    stamp_logical_order(&mut runs);

    ParagraphBlock {
        sdt_groups: None,
        id: BlockId::Str(para_id.clone()),
        para_id: (!para_id.is_empty() && !para_id_is_generated).then_some(para_id),
        runs,
        attrs: Some(lower_paragraph_attrs(
            &values,
            pilcrow_attributes,
            env,
            list_state,
        )),
        pm_start: Some(paragraph_pm_start as f64),
        pm_end: Some((paragraph_pm_start + u64::from(paragraph_pm_units) + 2) as f64),
    }
}

fn pilcrow_values<T: ReadTxn>(pilcrow: &MapRef, txn: &T) -> BTreeMap<String, Any> {
    pilcrow
        .iter(txn)
        .filter_map(|(key, value)| match value {
            Out::Any(value) => Some((key.to_string(), value)),
            _ => None,
        })
        .collect()
}

/// OOXML defaults used when a `sectPr` overrides only part of the page
/// geometry: US-Letter page size, one-inch margins, half-inch column gap
/// (all in twips) — the same constants the PM path (`toLayoutBlocks`) uses.
const DEFAULT_PAGE_WIDTH_TWIPS: f64 = 12240.0;
const DEFAULT_PAGE_HEIGHT_TWIPS: f64 = 15840.0;
const DEFAULT_SECTION_MARGIN_TWIPS: f64 = 1440.0;
const DEFAULT_COLUMN_GAP_TWIPS: f64 = 720.0;

/// The running per-side margin cascade across section breaks, in twips. A
/// section that overrides any margin emits a FULL margins record; its unset
/// sides inherit from the prior section instead of resetting to the OOXML
/// default. Mirrors `lastSectionMarginsTwips` in `toLayoutBlocks`.
#[derive(Clone, Copy, Debug)]
struct SectionMarginsTwips {
    top: f64,
    bottom: f64,
    left: f64,
    right: f64,
}

impl Default for SectionMarginsTwips {
    fn default() -> Self {
        Self {
            top: DEFAULT_SECTION_MARGIN_TWIPS,
            bottom: DEFAULT_SECTION_MARGIN_TWIPS,
            left: DEFAULT_SECTION_MARGIN_TWIPS,
            right: DEFAULT_SECTION_MARGIN_TWIPS,
        }
    }
}

fn section_break_type(value: &str) -> Option<SectionBreakType> {
    match value {
        "continuous" => Some(SectionBreakType::Continuous),
        "nextPage" => Some(SectionBreakType::NextPage),
        "evenPage" => Some(SectionBreakType::EvenPage),
        "oddPage" => Some(SectionBreakType::OddPage),
        // `nextColumn` has no variant in the layout engine's enum.
        _ => None,
    }
}

/// Lowers a section-boundary pilcrow's `sectPr` sub-map + `sectionBreakType`
/// to the [`SectionBreakBlock`] the renderer consumes, or `None` when the
/// pilcrow carries no section properties. Geometry math mirrors the PM path
/// (`toLayoutBlocks`): twips → px, page size only when a dimension is
/// overridden, the margin cascade above, and columns only when count > 1.
fn section_break_block(
    values: &BTreeMap<String, Any>,
    cascade: &mut SectionMarginsTwips,
) -> Option<SectionBreakBlock> {
    let sect_pr = values.get("sectPr").and_then(any_map);
    let break_attr = value_string(values.get("sectionBreakType"));
    if sect_pr.is_none() && break_attr.is_none() {
        return None;
    }

    let para_id = value_string(values.get("paraId")).unwrap_or_default();
    let mut block = SectionBreakBlock {
        sdt_groups: None,
        id: BlockId::Str(format!("sect:{para_id}")),
        break_type: sect_pr
            .and_then(|map| map_string(map, "sectionStart"))
            .or(break_attr)
            .as_deref()
            .and_then(section_break_type),
        page_size: None,
        orientation: None,
        margins: None,
        columns: None,
    };

    let Some(sect_pr) = sect_pr else {
        return Some(block);
    };

    let page_width = map_number(sect_pr, "pageWidth");
    let page_height = map_number(sect_pr, "pageHeight");
    if page_width.is_some() || page_height.is_some() {
        block.page_size = Some(Size {
            w: twips_to_pixels(page_width.unwrap_or(DEFAULT_PAGE_WIDTH_TWIPS)),
            h: twips_to_pixels(page_height.unwrap_or(DEFAULT_PAGE_HEIGHT_TWIPS)),
        });
    }

    let top = map_number(sect_pr, "marginTop");
    let bottom = map_number(sect_pr, "marginBottom");
    let left = map_number(sect_pr, "marginLeft");
    let right = map_number(sect_pr, "marginRight");
    if top.is_some() || bottom.is_some() || left.is_some() || right.is_some() {
        *cascade = SectionMarginsTwips {
            top: top.unwrap_or(cascade.top),
            bottom: bottom.unwrap_or(cascade.bottom),
            left: left.unwrap_or(cascade.left),
            right: right.unwrap_or(cascade.right),
        };
        block.margins = Some(PageMargins {
            top: twips_to_pixels(cascade.top),
            right: twips_to_pixels(cascade.right),
            bottom: twips_to_pixels(cascade.bottom),
            left: twips_to_pixels(cascade.left),
            header: None,
            footer: None,
        });
    }

    let column_count = map_number(sect_pr, "columnCount").unwrap_or(1.0);
    if column_count > 1.0 {
        block.columns = Some(ColumnLayout {
            count: column_count,
            gap: twips_to_pixels(
                map_number(sect_pr, "columnSpace").unwrap_or(DEFAULT_COLUMN_GAP_TWIPS),
            ),
            equal_width: Some(map_bool(sect_pr, "equalWidth").unwrap_or(true)),
            separator: map_bool(sect_pr, "separator"),
        });
    }

    Some(block)
}

fn coalesce_runs(runs: Vec<RawRun>) -> Vec<RawRun> {
    let mut result: Vec<RawRun> = Vec::with_capacity(runs.len());
    for run in runs {
        let merged = if let Some(previous) = result.last_mut() {
            match (&mut previous.kind, &run.kind) {
                (RawRunKind::Text(previous_text), RawRunKind::Text(text))
                    if previous.story_end == run.story_start
                        && previous.pm_end == run.pm_start
                        && formatting_equal(&previous.formatting, &run.formatting)
                        && previous.inline_sdt_widget == run.inline_sdt_widget =>
                {
                    previous_text.push_str(text);
                    previous.story_end = run.story_end;
                    previous.pm_end = run.pm_end;
                    true
                }
                _ => false,
            }
        } else {
            false
        };
        if !merged {
            result.push(run);
        }
    }
    result
}

fn formatting_equal(left: &RunFormatting, right: &RunFormatting) -> bool {
    serde_json::to_value(left).expect("RunFormatting serializes")
        == serde_json::to_value(right).expect("RunFormatting serializes")
}

fn raw_run_to_layout(raw: RawRun, paragraph_pm_start: u64) -> Run {
    let pm_start = Some((paragraph_pm_start + 1 + u64::from(raw.pm_start)) as f64);
    let pm_end = Some((paragraph_pm_start + 1 + u64::from(raw.pm_end)) as f64);
    match raw.kind {
        RawRunKind::Text(text) => Run::Text(TextRun {
            fmt: raw.formatting,
            text,
            pm_start,
            pm_end,
            inline_sdt_widget: raw.inline_sdt_widget,
        }),
        RawRunKind::Tab => Run::Tab(TabRun {
            fmt: raw.formatting,
            pm_start,
            pm_end,
            width: None,
            leader_glyphs: None,
        }),
        RawRunKind::LineBreak => Run::LineBreak(LineBreakRun { pm_start, pm_end }),
        RawRunKind::Image(mut image) => {
            image.pm_start = pm_start;
            image.pm_end = pm_end;
            Run::Image(image)
        }
        RawRunKind::Field {
            field_type,
            raw_type,
            instruction,
            fallback,
        } => Run::Field(FieldRun {
            fmt: raw.formatting,
            field_type,
            raw_type,
            instruction,
            fallback,
            pm_start,
            pm_end,
        }),
    }
}

fn stamp_logical_order(runs: &mut [Run]) {
    for (index, run) in runs.iter_mut().enumerate() {
        match run {
            Run::Text(run) if run.fmt.logical_order == Some(u64::MAX) => {
                run.fmt.logical_order = None;
            }
            Run::Text(run) => run.fmt.logical_order = Some(index as u64),
            Run::Tab(run) => run.fmt.logical_order = Some(index as u64),
            Run::Field(run) => run.fmt.logical_order = Some(index as u64),
            Run::Image(_) | Run::LineBreak(_) | Run::Unsupported => {}
        }
    }
}

#[derive(Default)]
struct ListState {
    counters: BTreeMap<String, Vec<i64>>,
    seen_num_ids: BTreeSet<String>,
}

fn format_roman(value: i64, uppercase: bool) -> String {
    if value <= 0 {
        return String::new();
    }
    const ONES: [&str; 10] = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"];
    const TENS: [&str; 10] = ["", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"];
    const HUNDREDS: [&str; 10] = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM"];
    let mut result = "M".repeat((value / 1000) as usize);
    result.push_str(HUNDREDS[((value / 100) % 10) as usize]);
    result.push_str(TENS[((value / 10) % 10) as usize]);
    result.push_str(ONES[(value % 10) as usize]);
    if uppercase {
        result
    } else {
        result.to_ascii_lowercase()
    }
}

fn format_alpha(mut value: i64, uppercase: bool) -> String {
    if value <= 0 {
        return String::new();
    }
    let mut chars = Vec::new();
    while value > 0 {
        chars.push((b'A' + ((value - 1) % 26) as u8) as char);
        value = (value - 1) / 26;
    }
    let result: String = chars.into_iter().rev().collect();
    if uppercase {
        result
    } else {
        result.to_ascii_lowercase()
    }
}

fn format_list_counter(value: i64, format: Option<&str>) -> String {
    if value <= 0 {
        return String::new();
    }
    match format {
        Some("upperRoman") => format_roman(value, true),
        Some("lowerRoman") => format_roman(value, false),
        Some("upperLetter") => format_alpha(value, true),
        Some("lowerLetter") => format_alpha(value, false),
        Some("decimalZero") => format!("{value:02}"),
        Some("decimalZero3") => format!("{value:03}"),
        Some("decimalZero4") => format!("{value:04}"),
        Some("decimalZero5") => format!("{value:05}"),
        Some("none") => String::new(),
        _ => value.to_string(),
    }
}

fn list_level_formats(values: &BTreeMap<String, Any>) -> Vec<String> {
    let Some(Any::Array(formats)) = values.get("listLevelNumFmts") else {
        return Vec::new();
    };
    formats
        .iter()
        .filter_map(any_str)
        .map(str::to_owned)
        .collect()
}

fn resolve_list_template(template: &str, counters: &[i64], formats: &[String]) -> String {
    let chars: Vec<char> = template.chars().collect();
    let mut result = String::new();
    let mut index = 0;
    while index < chars.len() {
        if chars[index] == '%' && index + 1 < chars.len() {
            if let Some(digit) = chars[index + 1].to_digit(10) {
                if digit == 0 {
                    index += 2;
                    if chars
                        .get(index)
                        .is_some_and(|ch| matches!(ch, '.' | ')' | ':' | ']'))
                    {
                        index += 1;
                    }
                    continue;
                } else {
                    let counter_index = digit as usize - 1;
                    let value = counters.get(counter_index).copied().unwrap_or(0);
                    let formatted =
                        format_list_counter(value, formats.get(counter_index).map(String::as_str));
                    index += 2;
                    let punctuation = chars
                        .get(index)
                        .copied()
                        .filter(|ch| matches!(ch, '.' | ')' | ':' | ']'));
                    if !formatted.is_empty() {
                        result.push_str(&formatted);
                        if let Some(punctuation) = punctuation {
                            result.push(punctuation);
                        }
                    }
                    if punctuation.is_some() {
                        index += 1;
                    }
                    continue;
                }
            }
        }
        result.push(chars[index]);
        index += 1;
    }
    result
}

fn compute_list_marker(values: &BTreeMap<String, Any>, state: &mut ListState) -> Option<String> {
    let marker = value_string(values.get("listMarker"));
    let Some(Any::Map(num_pr)) = values.get("numPr") else {
        return marker;
    };
    let Some(num_id) = map_number(num_pr, "numId") else {
        return marker;
    };
    if num_id == 0.0 {
        return marker;
    }
    if values.get("listIsBullet").and_then(any_bool) == Some(true) {
        return Some(marker.unwrap_or_default());
    }

    let level = map_number(num_pr, "ilvl").unwrap_or(0.0).max(0.0) as usize;
    let formats = list_level_formats(values);
    let level_format = formats
        .get(level)
        .cloned()
        .or_else(|| value_string(values.get("listNumFmt")));
    if level_format.as_deref() == Some("none") {
        return marker.filter(|value| !value.is_empty());
    }

    let counter_key = value_number(values.get("listAbstractNumId"))
        .unwrap_or(num_id)
        .to_string();
    let counters = state
        .counters
        .entry(counter_key)
        .or_insert_with(|| vec![0; 9]);
    if counters.len() <= level {
        counters.resize(level + 1, 0);
    }
    let seen_key = format!("{num_id}:{level}");
    if state.seen_num_ids.insert(seen_key) {
        if let Some(start) = value_number(values.get("listStartOverride")) {
            counters[level] = start as i64 - 1;
        }
    }
    counters[level] += 1;
    for value in counters.iter_mut().skip(level + 1) {
        *value = 0;
    }

    if let Some(marker) = marker {
        if marker.contains('%') {
            return Some(resolve_list_template(&marker, counters, &formats));
        }
        return Some(marker);
    }
    let mut parts = Vec::new();
    for value in counters.iter().take(level + 1) {
        if *value <= 0 {
            break;
        }
        parts.push(value.to_string());
    }
    Some(if parts.is_empty() {
        "1.".to_owned()
    } else {
        format!("{}.", parts.join("."))
    })
}

fn lower_run_formatting(attributes: Option<&Attrs>, env: &RenderEnv) -> RunFormatting {
    let mut result = RunFormatting {
        bold: mark_bool(attributes, "bold"),
        ..RunFormatting::default()
    };
    if let Some(map) = attribute_map(attributes, "bold")
        && (map_bool(map, "cs") == Some(true) || map_bool(map, "complexScript") == Some(true))
    {
        result.bold_cs = Some(true);
    }
    result.italic = mark_bool(attributes, "italic");
    if let Some(map) = attribute_map(attributes, "italic")
        && (map_bool(map, "cs") == Some(true) || map_bool(map, "complexScript") == Some(true))
    {
        result.italic_cs = Some(true);
    }
    result.underline = attribute(attributes, "underline").and_then(|value| match value {
        Any::Bool(flag) => Some(UnderlineSpec::Flag(*flag)),
        Any::Map(map) => {
            let style = map_string(map, "style");
            let color = map.get("color").and_then(|color| resolve_color(color, env));
            if style.is_some() || color.is_some() {
                Some(UnderlineSpec::Styled { style, color })
            } else {
                Some(UnderlineSpec::Flag(true))
            }
        }
        Any::Null | Any::Undefined => None,
        _ => Some(UnderlineSpec::Flag(true)),
    });
    result.strike = mark_bool(attributes, "strike");

    result.color = attribute(attributes, "textColor")
        .or_else(|| attribute(attributes, "color"))
        .and_then(|value| resolve_color(value, env));
    result.highlight = attribute(attributes, "highlight").and_then(resolve_highlight);

    lower_font_family(attributes, &mut result);
    lower_font_size(attributes, &mut result);
    lower_language(attributes, &mut result);
    lower_character_spacing(attributes, &mut result);

    result.superscript = mark_bool(attributes, "superscript");
    result.subscript = mark_bool(attributes, "subscript");
    result.all_caps = mark_bool(attributes, "allCaps");
    result.small_caps = mark_bool(attributes, "smallCaps");
    result.imprint = mark_bool(attributes, "imprint");
    result.emboss = mark_bool(attributes, "emboss");
    result.text_shadow = mark_bool(attributes, "textShadow");
    result.text_outline = mark_bool(attributes, "textOutline");
    result.hidden = mark_bool(attributes, "hidden");
    result.rtl = mark_bool(attributes, "rtl");

    if let Some(value) = attribute(attributes, "complexScript") {
        match value {
            Any::Map(map) => {
                result.complex_script = Some(map_bool(map, "enabled").unwrap_or(true));
                result.bold_cs = map_bool(map, "bold").or(result.bold_cs);
                result.italic_cs = map_bool(map, "italic").or(result.italic_cs);
            }
            _ => result.complex_script = any_bool(value).or(Some(true)),
        }
    }
    result.bold_cs = mark_bool(attributes, "boldCs").or(result.bold_cs);
    result.italic_cs = mark_bool(attributes, "italicCs").or(result.italic_cs);

    if let Some(value) = attribute(attributes, "emphasisMark") {
        let emphasis = match value {
            Any::Map(map) => map_string(map, "type"),
            _ => value_string(Some(value)),
        };
        result.emphasis_mark = Some(match emphasis.as_deref() {
            Some("dot" | "comma" | "circle" | "underDot") => emphasis.unwrap(),
            _ => "dot".to_owned(),
        });
    }
    if let Some(value) = attribute(attributes, "textEffect") {
        let effect = match value {
            Any::Map(map) => map_string(map, "effect"),
            _ => value_string(Some(value)),
        };
        if effect.as_deref().is_some_and(|value| {
            matches!(
                value,
                "blinkBackground" | "lights" | "antsBlack" | "antsRed" | "shimmer" | "sparkle"
            )
        }) {
            result.text_effect = effect;
        }
    }
    if let Some(value) = attribute(attributes, "modernTextEffects") {
        let effects = match value {
            Any::Map(map) => map.get("effects").unwrap_or(value),
            _ => value,
        };
        if !is_nullish(effects) {
            result.modern_effects = serde_json::to_value(effects).ok();
        }
    }

    result.hyperlink = attribute(attributes, "hyperlink").and_then(lower_hyperlink);
    lower_revisions(attributes, env, &mut result);
    result
}

fn lower_font_family(attributes: Option<&Attrs>, result: &mut RunFormatting) {
    let Some(value) =
        attribute(attributes, "fontFamily").or_else(|| attribute(attributes, "rFonts"))
    else {
        return;
    };
    match value {
        Any::String(value) => {
            let family = value.to_string();
            result.font_family = Some(family.clone());
            result.font_slots = Some(RunFontSlots {
                ascii: Some(family.clone()),
                h_ansi: Some(family),
                ..RunFontSlots::default()
            });
        }
        Any::Map(map) => {
            let slots = RunFontSlots {
                ascii: map_string(map, "ascii"),
                h_ansi: map_string(map, "hAnsi"),
                east_asia: map_string(map, "eastAsia"),
                cs: map_string(map, "cs"),
                ascii_theme: map_string(map, "asciiTheme"),
                h_ansi_theme: map_string(map, "hAnsiTheme"),
                east_asia_theme: map_string(map, "eastAsiaTheme"),
                cs_theme: map_string(map, "csTheme"),
                hint: map_string(map, "hint"),
            };
            let rtl = mark_bool(attributes, "rtl") == Some(true);
            result.font_family = if rtl { slots.cs.clone() } else { None }
                .or_else(|| slots.ascii.clone())
                .or_else(|| slots.h_ansi.clone())
                .or_else(|| slots.east_asia.clone())
                .or_else(|| slots.cs.clone());
            result.font_slots = Some(slots);
        }
        _ => {}
    }
}

fn lower_font_size(attributes: Option<&Attrs>, result: &mut RunFormatting) {
    let Some(value) = attribute(attributes, "fontSize").or_else(|| attribute(attributes, "sz"))
    else {
        return;
    };
    match value {
        // Scalar Wave-0 marks represent authored `w:sz`; the PM-compatible object form preserves
        // independent `w:sz`/`w:szCs`. Both stay in half-points until this bridge.
        Any::Number(half_points) => result.font_size = Some(*half_points / 2.0),
        Any::BigInt(half_points) => result.font_size = Some(*half_points as f64 / 2.0),
        Any::Map(map) => {
            let size = map_number(map, "size").or_else(|| map_number(map, "sz"));
            let size_cs = map_number(map, "sizeCs").or_else(|| map_number(map, "szCs"));
            let rtl = mark_bool(attributes, "rtl") == Some(true);
            result.font_size = if rtl { size_cs.or(size) } else { size }.map(|value| value / 2.0);
            result.font_size_cs = size_cs.map(|value| value / 2.0);
        }
        _ => {}
    }
}

fn lower_language(attributes: Option<&Attrs>, result: &mut RunFormatting) {
    let Some(value) = attribute(attributes, "language").or_else(|| attribute(attributes, "lang"))
    else {
        return;
    };
    if let Any::Map(map) = value {
        result.language = Some(RunLanguageSlots {
            latin: map_string(map, "latin").or_else(|| map_string(map, "val")),
            east_asia: map_string(map, "eastAsia"),
            bidi: map_string(map, "bidi"),
        });
    }
}

fn lower_character_spacing(attributes: Option<&Attrs>, result: &mut RunFormatting) {
    let Some(Any::Map(map)) = attribute(attributes, "characterSpacing") else {
        return;
    };
    if let Some(spacing) = map_number(map, "spacing").filter(|value| *value != 0.0) {
        result.letter_spacing = Some(twips_to_pixels(spacing));
    }
    if let Some(position) = map_number(map, "position").filter(|value| *value != 0.0) {
        result.position_px = Some(half_points_to_pixels(position));
    }
    if let Some(scale) = map_number(map, "scale").filter(|value| *value != 100.0) {
        result.horizontal_scale = Some(scale);
    }
    if let Some(kerning) = map_number(map, "kerning").filter(|value| *value > 0.0) {
        result.kerning_min_pt = Some(kerning / 2.0);
    }
}

fn lower_hyperlink(value: &Any) -> Option<HyperlinkInfo> {
    match value {
        Any::String(href) => Some(HyperlinkInfo {
            href: href.to_string(),
            tooltip: None,
            no_default_style: None,
            target: None,
            history: None,
            doc_location: None,
        }),
        Any::Map(map) => Some(HyperlinkInfo {
            href: map_string(map, "href")?,
            tooltip: map_string(map, "tooltip"),
            no_default_style: map_bool(map, "noDefaultStyle"),
            target: map_string(map, "target"),
            history: map_bool(map, "history"),
            doc_location: map_string(map, "docLocation"),
        }),
        _ => None,
    }
}

#[derive(Clone, Debug)]
struct RevisionMeta {
    id: f64,
    author: Option<String>,
    date: Option<String>,
}

fn revision_meta(value: &Any, env: &RenderEnv) -> Option<RevisionMeta> {
    let Any::Map(map) = value else {
        return None;
    };
    let id_value = map.get("id").or_else(|| map.get("revisionId"))?;
    let id = match id_value {
        Any::String(value) => numeric_id(value, env),
        _ => any_number(id_value)?,
    };
    Some(RevisionMeta {
        id,
        author: map_string(map, "author"),
        date: map_string(map, "date"),
    })
}

fn lower_revisions(attributes: Option<&Attrs>, env: &RenderEnv, result: &mut RunFormatting) {
    let insertion = attribute(attributes, INS).and_then(|value| revision_meta(value, env));
    let deletion = attribute(attributes, DEL).and_then(|value| revision_meta(value, env));
    if insertion.is_some() {
        result.is_insertion = Some(true);
    }
    if deletion.is_some() {
        result.is_deletion = Some(true);
    }
    // The layout contract has one metadata triple even when both flags are present. Preserve the
    // insertion triple in the nested ins+del case; use deletion metadata when it is the only site.
    if let Some(revision) = insertion.or(deletion) {
        result.change_revision_id = Some(revision.id);
        result.change_author = revision.author;
        result.change_date = revision.date;
    }
}

fn lower_paragraph_attrs(
    values: &BTreeMap<String, Any>,
    pilcrow_attributes: Option<&Attrs>,
    env: &RenderEnv,
    list_state: &mut ListState,
) -> ParagraphAttrs {
    let alignment =
        value_string(values.get("alignment")).and_then(|alignment| match alignment.as_str() {
            "both" | "distribute" => Some("justify".to_owned()),
            "left" | "center" | "right" => Some(alignment),
            _ => None,
        });
    let mut result = ParagraphAttrs {
        alignment,
        style_id: paragraph_style_id(values),
        ..ParagraphAttrs::default()
    };
    lower_paragraph_spacing(values, &mut result);
    lower_paragraph_indent(values, &mut result);
    lower_paragraph_tabs(values, &mut result);

    result.keep_next = true_property(values, "keepNext");
    result.keep_lines = true_property(values, "keepLines");
    result.page_break_before = true_property(values, "pageBreakBefore");
    result.contextual_spacing = true_property(values, "contextualSpacing");
    result.bidi = true_property(values, "bidi");
    result.borders = lower_paragraph_borders(values, env);
    result.shading = values
        .get("shading")
        .and_then(any_map)
        .and_then(|shading| shading.get("fill"))
        .and_then(|fill| resolve_color(fill, env));

    if let Some(Any::Map(num_pr)) = values.get("numPr") {
        result.num_pr = Some(ListNumPr {
            num_id: map_number(num_pr, "numId"),
            ilvl: map_number(num_pr, "ilvl"),
        });
    }
    result.list_marker = compute_list_marker(values, list_state);
    result.list_is_bullet = values.get("listIsBullet").and_then(any_bool);
    result.list_marker_hidden = true_property(values, "listMarkerHidden");
    result.list_marker_font_family = value_string(values.get("listMarkerFontFamily"));
    result.list_marker_font_size = value_number(values.get("listMarkerFontSize"));
    result.list_marker_suffix = value_string(values.get("listMarkerSuffix"));
    result.default_tab_stop_twips = env.default_tab_stop_twips;
    lower_paragraph_defaults(values, &mut result);

    result.p_pr_ins = values
        .get("pPrIns")
        .filter(|value| !is_nullish(value))
        .and_then(|value| paragraph_revision_value(value, env))
        .or_else(|| {
            attribute(pilcrow_attributes, INS)
                .and_then(|value| paragraph_revision_value(value, env))
        });
    result.p_pr_del = values
        .get("pPrDel")
        .filter(|value| !is_nullish(value))
        .and_then(|value| paragraph_revision_value(value, env))
        .or_else(|| {
            attribute(pilcrow_attributes, DEL)
                .and_then(|value| paragraph_revision_value(value, env))
        });
    if result.num_pr.is_some() {
        let numbering_added = values.get("pPrChange").is_some_and(|changes| {
            let Any::Array(changes) = changes else {
                return false;
            };
            changes.iter().any(|change| {
                let Some(change) = any_map(change) else {
                    return true;
                };
                let Some(previous) = change.get("previousFormatting").and_then(any_map) else {
                    return true;
                };
                previous.get("numPr").is_none_or(is_nullish)
            })
        });
        result.list_marker_revision = if result.p_pr_del.is_some() {
            Some("del".to_owned())
        } else if result.p_pr_ins.is_some() || numbering_added {
            Some("ins".to_owned())
        } else {
            None
        };
    }

    result
}

fn lower_paragraph_borders(
    values: &BTreeMap<String, Any>,
    env: &RenderEnv,
) -> Option<ParagraphBorders> {
    let borders = values.get("borders").and_then(any_map)?;
    let side = |key: &str| {
        borders
            .get(key)
            .and_then(any_map)
            .and_then(|border| lower_paragraph_border(border, env))
    };
    let result = ParagraphBorders {
        top: side("top"),
        bottom: side("bottom"),
        left: side("left"),
        right: side("right"),
        between: side("between"),
        bar: side("bar"),
    };
    (result.top.is_some()
        || result.bottom.is_some()
        || result.left.is_some()
        || result.right.is_some()
        || result.between.is_some()
        || result.bar.is_some())
    .then_some(result)
}

fn lower_paragraph_border(
    border: &std::collections::HashMap<String, Any>,
    env: &RenderEnv,
) -> Option<BorderStyle> {
    let style = map_string(border, "style")?;
    if matches!(style.as_str(), "none" | "nil") {
        return None;
    }
    let css_style = match style.as_str() {
        "double"
        | "triple"
        | "thinThickSmallGap"
        | "thickThinSmallGap"
        | "thinThickThinSmallGap"
        | "thinThickMediumGap"
        | "thickThinMediumGap"
        | "thinThickThinMediumGap"
        | "thinThickLargeGap"
        | "thickThinLargeGap"
        | "thinThickThinLargeGap"
        | "doubleWave" => "double",
        "dotted" | "dotDotDash" => "dotted",
        "dashed" | "dashSmallGap" | "dotDash" | "dashDotStroked" => "dashed",
        "threeDEmboss" => "ridge",
        "threeDEngrave" => "groove",
        "outset" => "outset",
        "inset" => "inset",
        _ => "solid",
    };
    Some(BorderStyle {
        style: Some(css_style.to_owned()),
        width: Some(
            ((map_number(border, "size").unwrap_or(0.0) / 8.0) * 1.333)
                .round()
                .max(1.0),
        ),
        color: border
            .get("color")
            .and_then(|value| resolve_color(value, env))
            .or_else(|| Some("#000000".to_owned())),
        space: map_number(border, "space").map(|points| points * 4.0 / 3.0),
    })
}

fn lower_paragraph_spacing(values: &BTreeMap<String, Any>, result: &mut ParagraphAttrs) {
    let spacing_map = values.get("spacing").and_then(any_map);
    let original = values.get("_originalFormatting").and_then(any_map);
    let auto_before = original.and_then(|map| map_bool(map, "beforeAutospacing")) == Some(true);
    let auto_after = original.and_then(|map| map_bool(map, "afterAutospacing")) == Some(true);
    let before = value_number(values.get("spaceBefore"))
        .or_else(|| spacing_map.and_then(|map| map_number(map, "before")));
    let after = value_number(values.get("spaceAfter"))
        .or_else(|| spacing_map.and_then(|map| map_number(map, "after")));
    let line = value_number(values.get("lineSpacing"))
        .or_else(|| spacing_map.and_then(|map| map_number(map, "line")));
    let line_rule = value_string(values.get("lineSpacingRule"))
        .or_else(|| spacing_map.and_then(|map| map_string(map, "lineRule")));

    if auto_before || auto_after || before.is_some() || after.is_some() || line.is_some() {
        let mut spacing = ParagraphSpacing {
            before: if auto_before {
                Some(AUTO_PARAGRAPH_SPACING_PX)
            } else {
                before.map(twips_to_pixels)
            },
            after: if auto_after {
                Some(AUTO_PARAGRAPH_SPACING_PX)
            } else {
                after.map(twips_to_pixels)
            },
            ..ParagraphSpacing::default()
        };
        if let Some(line) = line {
            if matches!(line_rule.as_deref(), Some("exact" | "atLeast")) {
                spacing.line = Some(twips_to_pixels(line));
                spacing.line_unit = Some("px".to_owned());
                spacing.line_rule = line_rule;
            } else {
                spacing.line = Some(line / 240.0);
                spacing.line_unit = Some("multiplier".to_owned());
                spacing.line_rule = Some("auto".to_owned());
            }
        }
        result.spacing = Some(spacing);
    }

    if let Some(Any::Map(explicit)) = values.get("spacingExplicit") {
        result.spacing_explicit = Some(SpacingExplicit {
            before: map_bool(explicit, "before"),
            after: map_bool(explicit, "after"),
        });
    }
}

fn lower_paragraph_indent(values: &BTreeMap<String, Any>, result: &mut ParagraphAttrs) {
    let indent_map = values.get("indent").and_then(any_map);
    let mut left = value_number(values.get("indentLeft"))
        .or_else(|| indent_map.and_then(|map| map_number(map, "left")));
    let right = value_number(values.get("indentRight"))
        .or_else(|| indent_map.and_then(|map| map_number(map, "right")));
    let mut first_line = value_number(values.get("indentFirstLine"))
        .or_else(|| indent_map.and_then(|map| map_number(map, "firstLine")));
    let mut hanging = values
        .get("hangingIndent")
        .and_then(any_bool)
        .or_else(|| indent_map.and_then(|map| map_bool(map, "hanging")))
        .unwrap_or(false);

    if let Some(Any::Map(num_pr)) = values.get("numPr") {
        let num_id = map_number(num_pr, "numId").unwrap_or(0.0);
        if num_id != 0.0 && left.is_none() {
            let level = map_number(num_pr, "ilvl").unwrap_or(0.0);
            left = Some((level + 1.0) * 720.0);
            if first_line.is_none() {
                first_line = Some(-360.0);
                hanging = true;
            }
        }
    }

    if left.is_some() || right.is_some() || first_line.is_some() {
        result.indent = Some(ParagraphIndent {
            left: left.map(twips_to_pixels),
            right: right.map(twips_to_pixels),
            first_line: first_line.filter(|_| !hanging).map(twips_to_pixels),
            hanging: first_line
                .filter(|_| hanging)
                .map(|value| twips_to_pixels(value).abs()),
        });
    }
}

fn lower_paragraph_tabs(values: &BTreeMap<String, Any>, result: &mut ParagraphAttrs) {
    let Some(Any::Array(tabs)) = values.get("tabs") else {
        return;
    };
    let tabs: Vec<TabStop> = tabs
        .iter()
        .filter_map(|tab| {
            let Any::Map(tab) = tab else {
                return None;
            };
            let alignment = map_string(tab, "alignment")
                .or_else(|| map_string(tab, "val"))
                .unwrap_or_else(|| "left".to_owned());
            let val = match alignment.as_str() {
                "left" | "num" => "start",
                "right" => "end",
                "center" => "center",
                "decimal" => "decimal",
                "bar" => "bar",
                "clear" => "clear",
                value @ ("start" | "end") => value,
                _ => "start",
            }
            .to_owned();
            Some(TabStop {
                val,
                // Tab stops are the intentional existing-contract exception: layout consumes
                // their authored twip positions directly and converts at measurement time.
                pos: map_number(tab, "position").or_else(|| map_number(tab, "pos"))?,
                leader: map_string(tab, "leader"),
            })
        })
        .collect();
    if !tabs.is_empty() {
        result.tabs = Some(tabs);
    }
}

fn lower_paragraph_defaults(values: &BTreeMap<String, Any>, result: &mut ParagraphAttrs) {
    let Some(Any::Map(defaults)) = values.get("defaultTextFormatting") else {
        return;
    };
    result.default_font_size = map_number(defaults, "fontSize").map(|value| value / 2.0);
    if let Some(Any::Map(fonts)) = defaults.get("fontFamily") {
        result.default_font_family =
            map_string(fonts, "ascii").or_else(|| map_string(fonts, "hAnsi"));
    }
}

fn paragraph_run_defaults(values: &BTreeMap<String, Any>) -> RunFormatting {
    let mut result = RunFormatting::default();
    let Some(Any::Map(defaults)) = values.get("defaultTextFormatting") else {
        return result;
    };
    if let Some(Any::Map(fonts)) = defaults.get("fontFamily") {
        let slots = RunFontSlots {
            ascii: map_string(fonts, "ascii"),
            h_ansi: map_string(fonts, "hAnsi"),
            east_asia: map_string(fonts, "eastAsia"),
            cs: map_string(fonts, "cs"),
            ascii_theme: map_string(fonts, "asciiTheme"),
            h_ansi_theme: map_string(fonts, "hAnsiTheme"),
            east_asia_theme: map_string(fonts, "eastAsiaTheme"),
            cs_theme: map_string(fonts, "csTheme"),
            hint: map_string(fonts, "hint"),
        };
        result.font_family = slots
            .ascii
            .clone()
            .or_else(|| slots.h_ansi.clone())
            .or_else(|| slots.east_asia.clone())
            .or_else(|| slots.cs.clone());
        result.font_slots = Some(slots);
    }
    result.font_size = map_number(defaults, "fontSize").map(|value| value / 2.0);
    result.font_size_cs = map_number(defaults, "fontSizeCs").map(|value| value / 2.0);
    result.bold_cs = map_bool(defaults, "boldCs");
    result.italic_cs = map_bool(defaults, "italicCs");
    result.complex_script = map_bool(defaults, "cs");
    if let Some(Any::Map(language)) = defaults.get("language") {
        result.language = Some(RunLanguageSlots {
            latin: map_string(language, "latin").or_else(|| map_string(language, "val")),
            east_asia: map_string(language, "eastAsia"),
            bidi: map_string(language, "bidi"),
        });
    }
    result
}

fn apply_run_defaults(target: &mut RunFormatting, defaults: &RunFormatting) {
    if target.font_slots.is_none() {
        target.font_slots = defaults.font_slots.clone();
    }
    if target.font_family.is_none() {
        target.font_family = defaults.font_family.clone();
    }
    if target.font_size.is_none() {
        target.font_size = defaults.font_size;
    }
    if target.font_size_cs.is_none() {
        target.font_size_cs = defaults.font_size_cs;
    }
    if target.bold_cs.is_none() {
        target.bold_cs = defaults.bold_cs;
    }
    if target.italic_cs.is_none() {
        target.italic_cs = defaults.italic_cs;
    }
    if target.complex_script.is_none() {
        target.complex_script = defaults.complex_script;
    }
    if target.language.is_none() {
        target.language = defaults.language.clone();
    }
}

fn strip_toc_hyperlink_style(formatting: &mut RunFormatting) {
    if let Some(hyperlink) = &mut formatting.hyperlink {
        hyperlink.no_default_style = Some(true);
        formatting.color = None;
        formatting.underline = None;
    }
}

fn is_toc_style(style_id: &str) -> bool {
    let upper = style_id.to_ascii_uppercase();
    upper
        .strip_prefix("TOC")
        .is_some_and(|suffix| suffix.chars().all(|ch| ch.is_ascii_digit()))
}

fn paragraph_style_id(values: &BTreeMap<String, Any>) -> Option<String> {
    value_string(values.get("styleId")).or_else(|| value_string(values.get("pStyle")))
}

fn paragraph_revision_value(value: &Any, env: &RenderEnv) -> Option<Value> {
    let revision = revision_meta(value, env)?;
    let mut object = JsonMap::new();
    object.insert("revisionId".to_owned(), Value::from(revision.id));
    object.insert(
        "author".to_owned(),
        revision.author.map(Value::from).unwrap_or(Value::Null),
    );
    object.insert(
        "date".to_owned(),
        revision.date.map(Value::from).unwrap_or(Value::Null),
    );
    Some(Value::Object(object))
}

fn resolve_color(value: &Any, env: &RenderEnv) -> Option<String> {
    match value {
        Any::String(value) => Some(css_hex(value)),
        Any::Map(map) => {
            if map_bool(map, "auto") == Some(true) {
                return Some("#000000".to_owned());
            }
            let rgb = map_string(map, "rgb");
            let mut hex = if let Some(slot) = map_string(map, "themeColor") {
                theme_color(&slot, env).or(rgb)?
            } else {
                rgb?
            };
            hex = hex.trim_start_matches('#').to_ascii_uppercase();
            if let Some(tint) = map_string(map, "themeTint").and_then(|value| hex_byte(&value)) {
                hex = apply_tint(&hex, tint as f64 / 255.0);
            } else if let Some(shade) =
                map_string(map, "themeShade").and_then(|value| hex_byte(&value))
            {
                hex = apply_shade(&hex, shade as f64 / 255.0);
            }
            Some(format!("#{hex}"))
        }
        _ => None,
    }
}

fn theme_color(slot: &str, env: &RenderEnv) -> Option<String> {
    let canonical = match slot.to_ascii_lowercase().as_str() {
        "dark1" | "text1" | "tx1" => "dk1",
        "light1" | "background1" | "bg1" => "lt1",
        "dark2" | "text2" | "tx2" => "dk2",
        "light2" | "background2" | "bg2" => "lt2",
        "hyperlink" => "hlink",
        "followedhyperlink" => "folHlink",
        value => match value {
            "dk1" => "dk1",
            "lt1" => "lt1",
            "dk2" => "dk2",
            "lt2" => "lt2",
            "accent1" => "accent1",
            "accent2" => "accent2",
            "accent3" => "accent3",
            "accent4" => "accent4",
            "accent5" => "accent5",
            "accent6" => "accent6",
            "hlink" => "hlink",
            "folhlink" => "folHlink",
            _ => return None,
        },
    };
    env.theme_colors
        .get(slot)
        .or_else(|| env.theme_colors.get(canonical))
        .cloned()
        .or_else(|| default_theme_color(canonical).map(str::to_owned))
}

fn default_theme_color(slot: &str) -> Option<&'static str> {
    Some(match slot {
        "dk1" => "000000",
        "lt1" => "FFFFFF",
        "dk2" => "44546A",
        "lt2" => "E7E6E6",
        "accent1" => "4472C4",
        "accent2" => "ED7D31",
        "accent3" => "A5A5A5",
        "accent4" => "FFC000",
        "accent5" => "5B9BD5",
        "accent6" => "70AD47",
        "hlink" => "0563C1",
        "folHlink" => "954F72",
        _ => return None,
    })
}

fn resolve_highlight(value: &Any) -> Option<String> {
    let color = match value {
        Any::String(value) => value.as_ref(),
        Any::Map(map) => map.get("color").and_then(any_str)?,
        _ => return None,
    };
    let hex = match color {
        "black" => "000000",
        "blue" => "0000FF",
        "cyan" => "00FFFF",
        "darkBlue" => "00008B",
        "darkCyan" => "008B8B",
        "darkGray" => "A9A9A9",
        "darkGreen" => "006400",
        "darkMagenta" => "8B008B",
        "darkRed" => "8B0000",
        "darkYellow" => "808000",
        "green" => "00FF00",
        "lightGray" => "D3D3D3",
        "magenta" => "FF00FF",
        "red" => "FF0000",
        "white" => "FFFFFF",
        "yellow" => "FFFF00",
        "none" => return Some(String::new()),
        value => return Some(css_hex(value)),
    };
    Some(format!("#{hex}"))
}

fn apply_tint(hex: &str, tint: f64) -> String {
    let (r, g, b) = hex_rgb(hex);
    rgb_hex(
        r as f64 * tint + 255.0 * (1.0 - tint),
        g as f64 * tint + 255.0 * (1.0 - tint),
        b as f64 * tint + 255.0 * (1.0 - tint),
    )
}

fn apply_shade(hex: &str, shade: f64) -> String {
    let (r, g, b) = hex_rgb(hex);
    rgb_hex(r as f64 * shade, g as f64 * shade, b as f64 * shade)
}

fn hex_rgb(hex: &str) -> (u8, u8, u8) {
    let normalized = format!("{hex:0>6}");
    let normalized = &normalized[..normalized.len().min(6)];
    (
        u8::from_str_radix(&normalized[0..2], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[2..4], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[4..6], 16).unwrap_or(0),
    )
}

fn rgb_hex(r: f64, g: f64, b: f64) -> String {
    let channel = |value: f64| value.round().clamp(0.0, 255.0) as u8;
    format!("{:02X}{:02X}{:02X}", channel(r), channel(g), channel(b))
}

fn css_hex(value: &str) -> String {
    if value.eq_ignore_ascii_case("auto") {
        return "#000000".to_owned();
    }
    format!("#{}", value.trim_start_matches('#').to_ascii_uppercase())
}

fn hex_byte(value: &str) -> Option<u8> {
    u8::from_str_radix(value, 16).ok()
}

fn numeric_id(id: &str, env: &RenderEnv) -> f64 {
    if let Some(value) = env.numeric_ids.get(id) {
        return *value;
    }
    if let Ok(value) = id.parse::<f64>()
        && value.is_finite()
    {
        return value;
    }
    // Stable JS-safe fallback for standalone native rendering. The coexistence A/B path supplies
    // an explicit map because the numeric contract cannot losslessly carry `{client}:{counter}`.
    let hash = id.bytes().fold(0xcbf29ce484222325_u64, |hash, byte| {
        (hash ^ u64::from(byte)).wrapping_mul(0x100000001b3)
    });
    (hash & ((1_u64 << 53) - 1)) as f64
}

fn true_property(values: &BTreeMap<String, Any>, key: &str) -> Option<bool> {
    (values.get(key).and_then(any_bool) == Some(true)).then_some(true)
}

fn attribute<'a>(attributes: Option<&'a Attrs>, key: &str) -> Option<&'a Any> {
    attributes
        .and_then(|attributes| attributes.get(key))
        .filter(|value| !is_nullish(value))
}

fn attribute_map<'a>(
    attributes: Option<&'a Attrs>,
    key: &str,
) -> Option<&'a std::collections::HashMap<String, Any>> {
    attribute(attributes, key).and_then(any_map)
}

fn mark_bool(attributes: Option<&Attrs>, key: &str) -> Option<bool> {
    let value = attribute(attributes, key)?;
    match value {
        Any::Map(map) => map_bool(map, "enabled").or(Some(true)),
        _ => any_bool(value).or(Some(true)),
    }
}

fn any_map(value: &Any) -> Option<&std::collections::HashMap<String, Any>> {
    match value {
        Any::Map(map) => Some(map),
        _ => None,
    }
}

fn any_str(value: &Any) -> Option<&str> {
    match value {
        Any::String(value) => Some(value),
        _ => None,
    }
}

fn any_bool(value: &Any) -> Option<bool> {
    match value {
        Any::Bool(value) => Some(*value),
        Any::Number(value) => Some(*value != 0.0),
        Any::BigInt(value) => Some(*value != 0),
        _ => None,
    }
}

fn any_number(value: &Any) -> Option<f64> {
    match value {
        Any::Number(value) => Some(*value),
        Any::BigInt(value) => Some(*value as f64),
        Any::String(value) => value.parse().ok(),
        _ => None,
    }
}

fn value_string(value: Option<&Any>) -> Option<String> {
    value.and_then(any_str).map(str::to_owned)
}

fn value_number(value: Option<&Any>) -> Option<f64> {
    value.and_then(any_number)
}

fn map_string(map: &std::collections::HashMap<String, Any>, key: &str) -> Option<String> {
    value_string(map.get(key))
}

fn map_bool(map: &std::collections::HashMap<String, Any>, key: &str) -> Option<bool> {
    map.get(key).and_then(any_bool)
}

fn map_number(map: &std::collections::HashMap<String, Any>, key: &str) -> Option<f64> {
    map.get(key).and_then(any_number)
}

fn is_nullish(value: &Any) -> bool {
    matches!(value, Any::Null | Any::Undefined)
}

fn utf16_len(value: &str) -> u32 {
    value.encode_utf16().count() as u32
}

fn utf16_slice(value: &str, start: u32, end: u32) -> String {
    let mut utf16_offset = 0_u32;
    let mut byte_start = None;
    let mut byte_end = None;
    for (byte_index, ch) in value.char_indices() {
        if utf16_offset == start {
            byte_start = Some(byte_index);
        }
        if utf16_offset == end {
            byte_end = Some(byte_index);
            break;
        }
        utf16_offset += ch.len_utf16() as u32;
    }
    if utf16_offset == start && byte_start.is_none() {
        byte_start = Some(value.len());
    }
    if utf16_offset == end && byte_end.is_none() {
        byte_end = Some(value.len());
    }
    let start = byte_start.expect("yrs/comment offsets never split a UTF-16 surrogate pair");
    let end = byte_end.expect("yrs/comment offsets never split a UTF-16 surrogate pair");
    value[start..end].to_owned()
}

fn twips_to_pixels(twips: f64) -> f64 {
    twips / 1440.0 * 96.0
}

fn half_points_to_pixels(half_points: f64) -> f64 {
    half_points / 144.0 * 96.0
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::sync::Arc;

    use serde_json::json;
    use yrs::{Any, Text, Transact};

    use super::*;
    use crate::{EditCtx, FormatPolicy, Position, RawOp, SimpleFormat, StoryRange};

    const DATE: &str = "2026-07-13T12:00:00Z";

    fn any_map(entries: impl IntoIterator<Item = (&'static str, Any)>) -> Any {
        Any::Map(Arc::new(
            entries
                .into_iter()
                .map(|(key, value)| (key.to_owned(), value))
                .collect::<HashMap<_, _>>(),
        ))
    }

    fn format_range(doc: &EditingDoc, start: u32, end: u32, attrs: Vec<(&'static str, Any)>) {
        let mut txn = doc.doc.transact_mut_with(doc.client_id);
        let story = story_ref(&txn, "body").unwrap();
        let attrs: Attrs = attrs
            .into_iter()
            .map(|(key, value)| (Arc::from(key), value))
            .collect();
        story.format(&mut txn, start, end - start, attrs);
    }

    fn normalize_block_ids(value: &mut Value) {
        match value {
            Value::Array(values) => {
                for value in values {
                    normalize_block_ids(value);
                }
            }
            Value::Object(values) => {
                if values.contains_key("id") {
                    values.insert("id".to_owned(), Value::from("normalized"));
                }
                for value in values.values_mut() {
                    normalize_block_ids(value);
                }
            }
            _ => {}
        }
    }

    fn replace_story_with_paragraph(doc: &EditingDoc, story_id: &str, para_id: &str, text: &str) {
        doc.create_story(story_id, "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            story_id,
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::Insert {
                    index: 0,
                    text: text.to_owned(),
                    attrs: Attrs::new(),
                },
                RawOp::InsertEmbed {
                    index: utf16_len(text),
                    kind: "pilcrow".to_owned(),
                    payload: vec![
                        ("paraId".to_owned(), Any::from(para_id)),
                        ("hangingIndent".to_owned(), Any::Bool(false)),
                    ],
                    attrs: Attrs::new(),
                },
            ],
            &EditCtx::local("", DATE),
        )
        .unwrap();
    }

    #[test]
    fn native_two_by_two_table_matches_pm_layout_contract_json() {
        let doc = EditingDoc::new(41);
        for (story, para, text) in [
            ("body:t0:r0c0", "c00p", "A"),
            ("body:t0:r0c1", "c01p", "B"),
            ("body:t0:r1c0", "c10p", "C"),
            ("body:t0:r1c1", "c11p", "D"),
        ] {
            replace_story_with_paragraph(&doc, story, para, text);
        }
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::InsertEmbed {
                    index: 0,
                    kind: "table".to_owned(),
                    payload: vec![
                        ("tblPr".to_owned(), Any::from_json("{}").unwrap()),
                        ("grid".to_owned(), Any::from_json("[]").unwrap()),
                        (
                            "rows".to_owned(),
                            Any::from_json(
                                r#"[
                                  {"trPr":{"isHeader":false,"trIns":{"id":"41:9","author":"Ada","date":"2025-01-01T00:00:00Z"}},"cells":[
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r0c0"},
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r0c1"}
                                  ]},
                                  {"trPr":{"isHeader":false},"cells":[
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r1c0"},
                                    {"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"body:t0:r1c1"}
                                  ]}
                                ]"#,
                            )
                            .unwrap(),
                        ),
                    ],
                    attrs: Attrs::new(),
                },
            ],
            &EditCtx::local("", DATE),
        )
        .unwrap();

        let mut actual = serde_json::to_value(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap(),
        )
        .unwrap();
        let stable_table_id = "body:table:body:t0:r0c0";
        assert_eq!(actual.pointer("/0/id"), Some(&Value::from(stable_table_id)));
        assert_eq!(
            actual.pointer("/0/rows/0/id"),
            Some(&Value::from(format!("{stable_table_id}:r0")))
        );
        let paragraph = |para_id: &str, text: &str, start: f64| {
            json!({
                "kind": "paragraph",
                "id": "pm-block",
                "paraId": para_id,
                "runs": [{
                    "kind": "text", "text": text, "logicalOrder": 0,
                    "pmStart": start + 1.0, "pmEnd": start + 2.0
                }],
                "attrs": {},
                "pmStart": start, "pmEnd": start + 3.0
            })
        };
        let cell = |para_id: &str, text: &str, start: f64| {
            json!({
                "id": "pm-block",
                "blocks": [paragraph(para_id, text, start)],
                "colSpan": 1.0,
                "rowSpan": 1.0,
                "padding": { "top": 0.0, "right": 0.0, "bottom": 0.0, "left": 0.0 }
            })
        };
        let mut expected = json!([{
            "kind": "table",
            "id": "pm-block",
            "rows": [
                {
                    "id": "pm-block", "isHeader": false,
                    "trackedIns": {
                        "revisionId": numeric_id("41:9", &RenderEnv::default()),
                        "author": "Ada", "date": "2025-01-01T00:00:00Z"
                    },
                    "cells": [cell("c00p", "A", 3.0), cell("c01p", "B", 8.0)]
                },
                {
                    "id": "pm-block", "isHeader": false,
                    "cells": [cell("c10p", "C", 15.0), cell("c11p", "D", 20.0)]
                }
            ],
            "pmStart": 0.0,
            "pmEnd": 26.0
        }]);
        normalize_block_ids(&mut actual);
        normalize_block_ids(&mut expected);
        assert_eq!(actual, expected);

        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Insert {
                    index: 0,
                    text: "prefix".to_owned(),
                    attrs: Attrs::new(),
                },
                RawOp::InsertEmbed {
                    index: 6,
                    kind: "pilcrow".to_owned(),
                    payload: vec![("paraId".to_owned(), Any::from("prefix-p"))],
                    attrs: Attrs::new(),
                },
            ],
            &EditCtx::local("", DATE),
        )
        .unwrap();
        let shifted = serde_json::to_value(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap(),
        )
        .unwrap();
        assert_eq!(
            shifted.pointer("/1/id"),
            Some(&Value::from(stable_table_id))
        );
        assert_eq!(
            shifted.pointer("/1/rows/0/id"),
            Some(&Value::from(format!("{stable_table_id}:r0")))
        );
    }

    #[test]
    fn native_note_ref_lowers_to_a_superscript_footnote_run() {
        let doc = EditingDoc::new(42);
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::Insert {
                    index: 0,
                    text: "See ".to_owned(),
                    attrs: Attrs::new(),
                },
                RawOp::InsertEmbed {
                    index: 4,
                    kind: "noteRef".to_owned(),
                    payload: vec![("footnoteRefId".to_owned(), Any::Number(5.0))],
                    attrs: Attrs::new(),
                },
                RawOp::InsertEmbed {
                    index: 5,
                    kind: "pilcrow".to_owned(),
                    payload: vec![("paraId".to_owned(), Any::from("body-p"))],
                    attrs: Attrs::new(),
                },
            ],
            &EditCtx::local("", DATE),
        )
        .unwrap();
        replace_story_with_paragraph(&doc, "fn:5", "fn-p", "Footnote text");

        let body = serde_json::to_value(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap(),
        )
        .unwrap();
        assert_eq!(body[0]["runs"][1]["text"], json!("5"));
        assert_eq!(body[0]["runs"][1]["footnoteRefId"], json!(5.0));
        assert_eq!(body[0]["runs"][1]["superscript"], json!(true));

        let footnote = yrs_doc_to_layout_blocks(&doc, "fn:5", &RenderEnv::default()).unwrap();
        assert_eq!(footnote.len(), 1);
    }

    #[test]
    fn native_page_and_column_break_embeds_lower_without_fallback() {
        let doc = EditingDoc::new(43);
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::InsertEmbed {
                    index: 0,
                    kind: "pageBreak".to_owned(),
                    payload: Vec::new(),
                    attrs: Attrs::new(),
                },
                RawOp::InsertEmbed {
                    index: 1,
                    kind: "columnBreak".to_owned(),
                    payload: Vec::new(),
                    attrs: Attrs::new(),
                },
            ],
            &EditCtx::local("", DATE),
        )
        .unwrap();

        let value = serde_json::to_value(
            yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap(),
        )
        .unwrap();
        assert_eq!(
            value,
            json!([
                {
                    "kind": "pageBreak", "id": "body:pageBreak:0",
                    "pmStart": 0.0, "pmEnd": 1.0
                },
                {
                    "kind": "columnBreak", "id": "body:columnBreak:1",
                    "pmStart": 1.0, "pmEnd": 2.0
                }
            ])
        );
    }

    #[test]
    fn pm_formula_matches_explicit_paragraph_node_sizes() {
        // Story: "a😀¶¶wxyz¶". PM paragraph node sizes are text UTF-16 length + 2.
        assert_eq!(utf16_len("a😀"), 3);
        let paragraphs = [(0, 3, 0), (4, 4, 1), (5, 9, 2)];
        let expected_blocks = [(0, 5), (5, 7), (7, 13)];
        let actual: Vec<(u64, u64)> = paragraphs
            .into_iter()
            .map(|(start, pilcrow, before)| {
                (
                    pm_position(start, before) - 1,
                    pm_position(pilcrow, before) + 1,
                )
            })
            .collect();
        assert_eq!(actual, expected_blocks);

        // The non-empty runs sit one position after their paragraph opening tags.
        assert_eq!(pm_position(0, 0), 1);
        assert_eq!(pm_position(3, 0), 4);
        assert_eq!(pm_position(5, 2), 8);
        assert_eq!(pm_position(9, 2), 12);
    }

    #[test]
    fn lowers_authored_run_units_and_passive_formats() {
        let attributes: Attrs = [
            (
                Arc::from("fontSize"),
                any_map([("size", Any::Number(22.0)), ("sizeCs", Any::Number(24.0))]),
            ),
            (
                Arc::from("fontFamily"),
                any_map([
                    ("ascii", Any::from("Aptos")),
                    ("hAnsi", Any::from("Aptos")),
                    ("cs", Any::from("Arial")),
                    ("hint", Any::from("cs")),
                ]),
            ),
            (
                Arc::from("characterSpacing"),
                any_map([
                    ("spacing", Any::Number(15.0)),
                    ("position", Any::Number(12.0)),
                    ("scale", Any::Number(90.0)),
                    ("kerning", Any::Number(16.0)),
                ]),
            ),
            (
                Arc::from("language"),
                any_map([
                    ("latin", Any::from("en-US")),
                    ("eastAsia", Any::from("ja-JP")),
                    ("bidi", Any::from("ar-SA")),
                ]),
            ),
            (Arc::from("highlight"), Any::from("yellow")),
            (Arc::from("superscript"), Any::Bool(true)),
            (Arc::from("allCaps"), Any::Bool(true)),
            (Arc::from("smallCaps"), Any::Bool(true)),
            (Arc::from("imprint"), Any::Bool(true)),
            (Arc::from("emboss"), Any::Bool(true)),
            (Arc::from("textShadow"), Any::Bool(true)),
            (Arc::from("textOutline"), Any::Bool(true)),
            (Arc::from("hidden"), Any::Bool(true)),
            (Arc::from("rtl"), Any::Bool(true)),
            (
                Arc::from("complexScript"),
                any_map([
                    ("enabled", Any::Bool(true)),
                    ("bold", Any::Bool(true)),
                    ("italic", Any::Bool(false)),
                ]),
            ),
            (
                Arc::from("emphasisMark"),
                any_map([("type", Any::from("underDot"))]),
            ),
            (
                Arc::from("textEffect"),
                any_map([("effect", Any::from("shimmer"))]),
            ),
            (
                Arc::from("modernTextEffects"),
                any_map([(
                    "effects",
                    any_map([(
                        "glow",
                        any_map([
                            ("color", Any::from("#00FF00")),
                            ("radius", Any::Number(3.0)),
                        ]),
                    )]),
                )]),
            ),
        ]
        .into_iter()
        .collect();

        let actual = serde_json::to_value(lower_run_formatting(
            Some(&attributes),
            &RenderEnv::default(),
        ))
        .unwrap();
        assert_eq!(
            actual,
            json!({
                "highlight": "#FFFF00",
                "fontFamily": "Arial",
                "fontSlots": {
                    "ascii": "Aptos", "hAnsi": "Aptos", "cs": "Arial", "hint": "cs"
                },
                "fontSize": 12.0,
                "fontSizeCs": 12.0,
                "boldCs": true,
                "italicCs": false,
                "complexScript": true,
                "language": { "latin": "en-US", "eastAsia": "ja-JP", "bidi": "ar-SA" },
                "letterSpacing": 1.0,
                "superscript": true,
                "allCaps": true,
                "smallCaps": true,
                "positionPx": 8.0,
                "horizontalScale": 90.0,
                "kerningMinPt": 8.0,
                "imprint": true,
                "emboss": true,
                "textShadow": true,
                "textOutline": true,
                "emphasisMark": "underDot",
                "hidden": true,
                "rtl": true,
                "textEffect": "shimmer",
                "modernEffects": { "glow": { "color": "#00FF00", "radius": 3 } }
            })
        );
    }

    #[test]
    fn representative_s1_story_matches_pm_layout_contract_json() {
        let doc = EditingDoc::new(41);
        doc.create_story("body", "Alpha link Omega", "Normal", "left")
            .unwrap();
        let split = doc
            .split_paragraph(
                &EditCtx::local("Bob", DATE).suggesting(),
                Position::new("body", 10),
                None,
            )
            .unwrap();
        // New S1 split (op-contract R6): the FIRST half keeps the original paraId and the second
        // half is re-minted — the reverse of the retired foundation split, so the second-half
        // paragraph attrs below target `second_para`.
        let first_para = split.first_para_id.clone();
        let second_para = split.second_para_id.clone();

        let direct = EditCtx::local("", DATE);
        doc.toggle_format(&direct, StoryRange::new("body", 0, 5), SimpleFormat::Bold)
            .unwrap();
        doc.toggle_format(
            &direct,
            StoryRange::new("body", 6, 10),
            SimpleFormat::Italic,
        )
        .unwrap();
        format_range(
            &doc,
            0,
            5,
            vec![("textColor", any_map([("themeColor", Any::from("accent1"))]))],
        );
        format_range(
            &doc,
            6,
            10,
            vec![
                (
                    "underline",
                    any_map([
                        ("style", Any::from("single")),
                        ("color", any_map([("rgb", Any::from("00FF00"))])),
                    ]),
                ),
                ("textColor", any_map([("rgb", Any::from("0563C1"))])),
                (
                    "hyperlink",
                    any_map([
                        ("href", Any::from("https://example.test")),
                        ("tooltip", Any::from("Example")),
                    ]),
                ),
            ],
        );

        doc.set_paragraph_attr(&first_para, "pStyle", Any::from("TOC1"))
            .unwrap();
        doc.set_paragraph_attr(&first_para, "alignment", Any::from("both"))
            .unwrap();
        doc.set_paragraph_attr(&first_para, "spaceBefore", Any::Number(120.0))
            .unwrap();
        doc.set_paragraph_attr(&first_para, "spaceAfter", Any::Number(240.0))
            .unwrap();
        doc.set_paragraph_attr(&first_para, "lineSpacing", Any::Number(360.0))
            .unwrap();
        doc.set_paragraph_attr(&first_para, "lineSpacingRule", Any::from("auto"))
            .unwrap();
        doc.set_paragraph_attr(
            &first_para,
            "spacingExplicit",
            any_map([("before", Any::Bool(true)), ("after", Any::Bool(false))]),
        )
        .unwrap();
        doc.set_paragraph_attr(
            &first_para,
            "numPr",
            any_map([("numId", Any::Number(7.0)), ("ilvl", Any::Number(1.0))]),
        )
        .unwrap();
        doc.set_paragraph_attr(
            &first_para,
            "tabs",
            Any::Array(Arc::from([any_map([
                ("alignment", Any::from("left")),
                ("position", Any::Number(720.0)),
                ("leader", Any::from("dot")),
            ])])),
        )
        .unwrap();
        for key in [
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "contextualSpacing",
            "bidi",
        ] {
            doc.set_paragraph_attr(&first_para, key, Any::Bool(true))
                .unwrap();
        }

        doc.set_paragraph_attr(&second_para, "alignment", Any::from("right"))
            .unwrap();
        doc.set_paragraph_attr(
            &second_para,
            "defaultTextFormatting",
            any_map([
                (
                    "fontFamily",
                    any_map([
                        ("ascii", Any::from("Courier New")),
                        ("hAnsi", Any::from("Courier New")),
                    ]),
                ),
                ("fontSize", Any::Number(20.0)),
            ]),
        )
        .unwrap();

        let insertion_id = doc
            .insert_text(
                &EditCtx::local("Alice", DATE).suggesting(),
                Position::new("body", 12),
                "NEW",
                FormatPolicy::Inherit,
            )
            .unwrap()
            .revision_ids
            .into_iter()
            .next()
            .unwrap();
        let comment_id = doc
            .add_comment(
                &[StoryRange::new("body", 6, 10)],
                "Reviewer",
                DATE,
                Any::from("anchor"),
            )
            .unwrap();

        assert_eq!(first_para, "41:0");
        assert_eq!(second_para, "41:1");
        assert_eq!(insertion_id, "41:3");
        assert_eq!(comment_id, "41:4");
        let env = RenderEnv {
            default_tab_stop_twips: Some(720.0),
            numeric_ids: BTreeMap::from([
                ("41:2".to_owned(), 2.0),
                (insertion_id, 3.0),
                (comment_id, 4.0),
            ]),
            ..RenderEnv::default()
        };

        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &env).unwrap();
        let mut actual = serde_json::to_value(blocks).unwrap();
        let mut expected = json!([
            {
                "kind": "paragraph",
                "id": "block-1",
                "paraId": "41:0",
                "runs": [
                    {
                        "kind": "text", "text": "Alpha", "bold": true,
                        "color": "#4472C4", "logicalOrder": 0, "pmStart": 1.0, "pmEnd": 6.0
                    },
                    {
                        "kind": "text", "text": " ", "logicalOrder": 1,
                        "pmStart": 6.0, "pmEnd": 7.0
                    },
                    {
                        "kind": "text", "text": "link", "italic": true,
                        "hyperlink": {
                            "href": "https://example.test", "tooltip": "Example",
                            "noDefaultStyle": true
                        },
                        "commentIds": [4.0], "logicalOrder": 2,
                        "pmStart": 7.0, "pmEnd": 11.0
                    }
                ],
                "attrs": {
                    "alignment": "justify",
                    "spacing": {
                        "before": 8.0, "after": 16.0, "line": 1.5,
                        "lineUnit": "multiplier", "lineRule": "auto"
                    },
                    "spacingExplicit": { "before": true, "after": false },
                    "indent": { "left": 96.0, "hanging": 24.0 },
                    "keepNext": true, "keepLines": true, "pageBreakBefore": true,
                    "styleId": "TOC1", "contextualSpacing": true, "bidi": true,
                    "tabs": [{ "val": "start", "pos": 720.0, "leader": "dot" }],
                    "numPr": { "numId": 7.0, "ilvl": 1.0 },
                    "listMarker": "1.",
                    "listMarkerRevision": "ins",
                    "defaultTabStopTwips": 720.0,
                    "pPrIns": { "revisionId": 2.0, "author": "Bob", "date": DATE }
                },
                "pmStart": 0.0, "pmEnd": 12.0
            },
            {
                "kind": "paragraph",
                "id": "block-2",
                "paraId": "41:1",
                "runs": [
                    {
                        "kind": "text", "text": " ", "fontFamily": "Courier New",
                        "fontSlots": { "ascii": "Courier New", "hAnsi": "Courier New" },
                        "fontSize": 10.0, "logicalOrder": 0, "pmStart": 13.0, "pmEnd": 14.0
                    },
                    {
                        "kind": "text", "text": "NEW", "fontFamily": "Courier New",
                        "fontSlots": { "ascii": "Courier New", "hAnsi": "Courier New" },
                        "fontSize": 10.0, "isInsertion": true, "changeAuthor": "Alice",
                        "changeDate": DATE, "changeRevisionId": 3.0,
                        "logicalOrder": 1, "pmStart": 14.0, "pmEnd": 17.0
                    },
                    {
                        "kind": "text", "text": "Omega", "fontFamily": "Courier New",
                        "fontSlots": { "ascii": "Courier New", "hAnsi": "Courier New" },
                        "fontSize": 10.0, "logicalOrder": 2, "pmStart": 17.0, "pmEnd": 22.0
                    }
                ],
                "attrs": {
                    "alignment": "right", "styleId": "Normal",
                    "defaultTabStopTwips": 720.0, "defaultFontSize": 10.0,
                    "defaultFontFamily": "Courier New"
                },
                "pmStart": 12.0, "pmEnd": 23.0
            }
        ]);
        normalize_block_ids(&mut actual);
        normalize_block_ids(&mut expected);
        assert_eq!(actual, expected);
        assert_eq!(
            serde_json::to_vec(&actual).unwrap(),
            serde_json::to_vec(&expected).unwrap(),
            "normalized LayoutBlock[] JSON must be byte-identical"
        );
    }

    #[test]
    fn section_break_pilcrows_emit_section_break_blocks_with_margin_cascade() {
        let doc = EditingDoc::new(47);
        doc.create_story("body", "onetwothree", "Normal", "left")
            .unwrap();
        let direct = EditCtx::local("", DATE);
        let split = doc
            .split_paragraph(&direct, Position::new("body", 3), None)
            .unwrap();
        let first = split.first_para_id.clone();
        let split_two = doc
            .split_paragraph(&direct, Position::new("body", 7), None)
            .unwrap();
        let second = split_two.first_para_id.clone();
        assert_eq!(first, "47:0");
        assert_eq!(second, "47:1");

        // Section 1: full sectPr — page size, one overridden margin, columns.
        doc.set_paragraph_attr(
            &first,
            "sectPr",
            any_map([
                ("sectionStart", Any::from("nextPage")),
                ("pageWidth", Any::Number(15840.0)),
                ("pageHeight", Any::Number(12240.0)),
                ("orientation", Any::from("landscape")),
                ("marginTop", Any::Number(720.0)),
                ("columnCount", Any::Number(2.0)),
                ("columnSpace", Any::Number(360.0)),
                ("separator", Any::Bool(true)),
            ]),
        )
        .unwrap();
        // Section 2: break type from the `sectionBreakType` attr; its sectPr
        // overrides only the left margin — the other sides must inherit from
        // section 1's cascade (top 720), not reset to the OOXML default.
        doc.set_paragraph_attr(&second, "sectionBreakType", Any::from("continuous"))
            .unwrap();
        doc.set_paragraph_attr(
            &second,
            "sectPr",
            any_map([("marginLeft", Any::Number(2880.0))]),
        )
        .unwrap();

        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let value = serde_json::to_value(&blocks).unwrap();
        let kinds: Vec<&str> = value
            .as_array()
            .unwrap()
            .iter()
            .map(|block| block["kind"].as_str().unwrap())
            .collect();
        assert_eq!(
            kinds,
            [
                "paragraph",
                "sectionBreak",
                "paragraph",
                "sectionBreak",
                "paragraph"
            ]
        );
        assert_eq!(
            value[1],
            json!({
                "kind": "sectionBreak",
                "id": "sect:47:0",
                "type": "nextPage",
                "pageSize": { "w": 1056.0, "h": 816.0 },
                "margins": { "top": 48.0, "right": 96.0, "bottom": 96.0, "left": 96.0 },
                "columns": { "count": 2.0, "gap": 24.0, "equalWidth": true, "separator": true }
            })
        );
        assert_eq!(
            value[3],
            json!({
                "kind": "sectionBreak",
                "id": "sect:47:1",
                "type": "continuous",
                "margins": { "top": 48.0, "right": 96.0, "bottom": 96.0, "left": 192.0 }
            })
        );
    }

    #[test]
    fn bare_section_break_type_emits_a_type_only_section_break_block() {
        let doc = EditingDoc::new(48);
        let para = doc.create_story("body", "end", "Normal", "left").unwrap();
        doc.set_paragraph_attr(&para, "sectionBreakType", Any::from("oddPage"))
            .unwrap();

        let blocks = yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).unwrap();
        let value = serde_json::to_value(&blocks).unwrap();
        assert_eq!(value.as_array().unwrap().len(), 2);
        assert_eq!(
            value[1],
            json!({ "kind": "sectionBreak", "id": format!("sect:{para}"), "type": "oddPage" })
        );
    }
}
