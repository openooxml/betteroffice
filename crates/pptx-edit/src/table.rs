use pptx_parse::{GraphicFrameData, Table, TableRowSource, TableStructure};
use serde::{Deserialize, Serialize};
use yrs::types::{AsPrelim, text::YChange};
use yrs::{Any, In, Map, Out, ReadTxn, Text, TextRef, Transact};

use crate::deck::{map_number, map_string, map_string_array, required_map, shape_ref};
use crate::story::{materialize_paragraph_joins, snapshot_story, story_ref};
use crate::{DeckSession, EditCtx, EditError, EditResult, PARA_ID, STORIES, ShapeSnapshot};

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum TableInsertPosition {
    Before,
    After,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TableReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub row_count: usize,
    pub column_count: usize,
    pub new_cell_refs: Vec<String>,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum TableEdit {
    InsertRow,
    DeleteRow,
    InsertColumn,
    DeleteColumn,
}

enum Action<'a> {
    Insert(TableInsertPosition, &'a [String]),
    Delete,
}

impl DeckSession {
    pub fn insert_table_row(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        index: u32,
        position: TableInsertPosition,
        cells: &[String],
    ) -> EditResult<TableReceipt> {
        self.edit_table(
            context,
            slide_id,
            shape_id,
            index,
            true,
            Action::Insert(position, cells),
        )
    }

    pub fn delete_table_row(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        index: u32,
    ) -> EditResult<TableReceipt> {
        self.edit_table(context, slide_id, shape_id, index, true, Action::Delete)
    }

    pub fn insert_table_column(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        index: u32,
        position: TableInsertPosition,
        cells: &[String],
    ) -> EditResult<TableReceipt> {
        self.edit_table(
            context,
            slide_id,
            shape_id,
            index,
            false,
            Action::Insert(position, cells),
        )
    }

    pub fn delete_table_column(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        index: u32,
    ) -> EditResult<TableReceipt> {
        self.edit_table(context, slide_id, shape_id, index, false, Action::Delete)
    }

    /// Runs every check of a table structure edit without changing the deck.
    pub fn check_table_edit(
        &self,
        edit: TableEdit,
        slide_id: &str,
        shape_id: &str,
        index: u32,
        position: TableInsertPosition,
        cells: &[String],
    ) -> EditResult<()> {
        let (row_axis, action) = match edit {
            TableEdit::InsertRow => (true, Action::Insert(position, cells)),
            TableEdit::DeleteRow => (true, Action::Delete),
            TableEdit::InsertColumn => (false, Action::Insert(position, cells)),
            TableEdit::DeleteColumn => (false, Action::Delete),
        };
        self.plan_table_edit(
            &EditCtx::local("check"),
            slide_id,
            shape_id,
            index,
            row_axis,
            action,
            false,
        )
        .map(|_| ())
    }

    fn edit_table(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        index: u32,
        row_axis: bool,
        action: Action<'_>,
    ) -> EditResult<TableReceipt> {
        self.plan_table_edit(context, slide_id, shape_id, index, row_axis, action, true)?
            .ok_or_else(|| EditError::InvalidState("table edit was not applied".into()))
    }

    #[allow(clippy::too_many_arguments)]
    fn plan_table_edit(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        index: u32,
        row_axis: bool,
        action: Action<'_>,
        apply: bool,
    ) -> EditResult<Option<TableReceipt>> {
        if self.collaborative.get() {
            return Err(EditError::CollaborativeStructureOperation);
        }
        let snapshot = self.snapshot()?;
        let slide = snapshot
            .slides
            .iter()
            .find(|slide| slide.id == slide_id)
            .ok_or_else(|| EditError::SlideNotFound(slide_id.into()))?;
        let shape = find_shape(&slide.shapes, shape_id)
            .ok_or_else(|| EditError::ShapeNotFound(shape_id.into()))?;
        let Some(GraphicFrameData::Table(mut table)) = shape.graphic.clone() else {
            return Err(invalid("shape must identify a table from the outline"));
        };
        validate_table(&table)?;
        validate_structure(shape)?;
        let rows = table.rows.len();
        let columns = table.grid.len();
        if shape.text_stories.len() != rows * columns {
            return Err(invalid("table cell text stories do not match its grid"));
        }
        let length = if row_axis { rows } else { columns };
        let field = if row_axis { "row" } else { "column" };
        let index = index as usize;
        if index >= length {
            return Err(invalid(format!(
                "{field} index {index} is out of range; valid range is 0..{}",
                length - 1
            )));
        }
        let mut sources = shape
            .table_structure
            .clone()
            .unwrap_or_else(|| TableStructure {
                generation: String::new(),
                story_keys: shape
                    .text_stories
                    .iter()
                    .map(|story| story.id.clone())
                    .collect(),
                grid: table.grid.clone(),
                columns: (0..columns).collect(),
                copied_columns: vec![false; columns],
                rows: table
                    .rows
                    .iter()
                    .enumerate()
                    .map(|(r, row)| TableRowSource {
                        source_row: r,
                        copied: false,
                        copied_cells: vec![false; columns],
                        height: row.height,
                        cells: (0..columns).map(|c| [r, c]).collect(),
                    })
                    .collect(),
            });
        let mut donors: Vec<Vec<(usize, Option<&str>)>> = (0..rows)
            .map(|r| (0..columns).map(|c| (r * columns + c, None)).collect())
            .collect();
        let delta = match action {
            Action::Delete => {
                if length == 1 {
                    return Err(invalid(format!("cannot delete the last {field}")));
                }
                let delta = if row_axis {
                    -table.rows[index].height
                } else {
                    -table.grid[index]
                };
                if row_axis {
                    table.rows.remove(index);
                    sources.rows.remove(index);
                    donors.remove(index);
                } else {
                    table.grid.remove(index);
                    sources.columns.remove(index);
                    sources.copied_columns.remove(index);
                    for ((row, source), donor) in table
                        .rows
                        .iter_mut()
                        .zip(&mut sources.rows)
                        .zip(&mut donors)
                    {
                        row.cells.remove(index);
                        source.cells.remove(index);
                        source.copied_cells.remove(index);
                        donor.remove(index);
                    }
                }
                delta
            }
            Action::Insert(position, cells) => {
                let cell_count = if row_axis { columns } else { rows };
                if cells.len() > cell_count {
                    return Err(invalid(format!(
                        "cells has {} entries; supply at most {cell_count}",
                        cells.len()
                    )));
                }
                for text in cells {
                    if text.contains(['\r', '\n']) {
                        return Err(invalid(
                            "cells text must not contain line breaks; give each cell one line of text",
                        ));
                    }
                    crate::model::validate_xml_text(text)?;
                    if text.encode_utf16().count() > 16_000 {
                        return Err(invalid("cells text must fit in 16000 UTF-16 units"));
                    }
                }
                if rows
                    .saturating_add(1)
                    .saturating_mul(columns.saturating_add(1))
                    > 100_000
                {
                    return Err(invalid("table would exceed the 100000-cell edit limit"));
                }
                let at = index + usize::from(matches!(position, TableInsertPosition::After));
                if row_axis {
                    let row = table.rows[index].clone();
                    let mut source = sources.rows[index].clone();
                    source.copied = true;
                    source.copied_cells.fill(true);
                    table.rows.insert(at, row);
                    sources.rows.insert(at, source);
                    donors.insert(
                        at,
                        (0..columns)
                            .map(|c| {
                                (
                                    index * columns + c,
                                    Some(cells.get(c).map_or("", String::as_str)),
                                )
                            })
                            .collect(),
                    );
                } else {
                    let width = table.grid[index];
                    if width < 2 {
                        return Err(invalid("column width must be at least 2 EMU to split it"));
                    }
                    table.grid[index] = width - width / 2;
                    table.grid.insert(at, width / 2);
                    sources.columns.insert(at, sources.columns[index]);
                    sources.copied_columns.insert(at, true);
                    for (r, ((row, source), donor)) in table
                        .rows
                        .iter_mut()
                        .zip(&mut sources.rows)
                        .zip(&mut donors)
                        .enumerate()
                    {
                        row.cells.insert(at, row.cells[index].clone());
                        source.cells.insert(at, source.cells[index]);
                        source.copied_cells.insert(at, true);
                        donor.insert(
                            at,
                            (
                                r * columns + index,
                                Some(cells.get(r).map_or("", String::as_str)),
                            ),
                        );
                    }
                }
                if row_axis {
                    table.rows[index].height
                } else {
                    0
                }
            }
        };
        sources.grid.clone_from(&table.grid);
        let (width, height) = if row_axis {
            (
                shape.width,
                resize(shape.height, delta, table.rows.iter().map(|row| row.height))?,
            )
        } else {
            (
                resize(shape.width, delta, table.grid.iter().copied())?,
                shape.height,
            )
        };
        let theme = pptx_parse::slide_theme(
            self.package(),
            slide.source_part_path.as_deref(),
            slide.layout_part_path.as_deref(),
        );
        let txn = self.doc.transact();
        let mut copies = Vec::new();
        for row in &donors {
            for &(donor, text) in row {
                let source_id = &shape.text_stories[donor].id;
                let key = storage_key(&txn, source_id)?;
                let copy = if text.is_some() {
                    let story = story_ref(&txn, source_id)?;
                    let mut snapshot = snapshot_story(&story, &txn, source_id)?;
                    if let Some(GraphicFrameData::Table(template)) = &shape.graphic {
                        let body = &template.rows[donor / columns].cells[donor % columns].text;
                        if body
                            .paragraphs
                            .iter()
                            .flat_map(|p| &p.runs)
                            .any(|run| run.field_id.is_some() || run.field_type.is_some())
                        {
                            return Err(invalid(
                                "cannot copy a template cell containing text fields",
                            ));
                        }
                        for (index, paragraph) in snapshot.paragraphs.iter_mut().enumerate() {
                            if paragraph.runs.is_empty()
                                && let Some(properties) = body
                                    .paragraphs
                                    .get(index)
                                    .or_else(|| body.paragraphs.last())
                                    .and_then(|p| {
                                        p.runs
                                            .first()
                                            .map(|run| &run.properties)
                                            .or(p.end_properties.as_ref())
                                    })
                            {
                                paragraph.runs.push(crate::TextRunSnapshot {
                                    text: String::new(),
                                    style: crate::story::style_from_run_properties(
                                        properties,
                                        Some(&theme),
                                    ),
                                });
                            }
                        }
                    }
                    Some((In::Text(story.as_prelim(&txn)), snapshot))
                } else {
                    None
                };
                copies.push((source_id.clone(), key, text, copy));
            }
        }
        let stories = required_map(&txn, STORIES)?;
        let target = shape_ref(&txn, shape_id)?;
        drop(txn);
        if !apply {
            return Ok(None);
        }
        let parent = shape
            .table_structure
            .as_ref()
            .map_or("source", |structure| structure.generation.as_str());
        let generation = self.next_id("table");
        sources.generation.clone_from(&generation);
        let ids: Vec<_> = (0..copies.len())
            .map(|i| {
                format!(
                    "story:{shape_id}:table:{}:{}:{generation}",
                    i / table.grid.len(),
                    i % table.grid.len()
                )
            })
            .collect();
        sources.story_keys = copies
            .iter()
            .enumerate()
            .map(|(i, (_, key, text, _))| {
                if text.is_some() {
                    format!("cell:{generation}:{i}")
                } else {
                    key.clone()
                }
            })
            .collect();
        let graphic_json = serde_json::to_string(&GraphicFrameData::Table(table.clone()))
            .map_err(|error| EditError::Json(error.to_string()))?;
        let source_json =
            serde_json::to_string(&sources).map_err(|error| EditError::Json(error.to_string()))?;
        let mut new_cell_refs = Vec::new();
        let client = yrs::ClientID::new(self.client_id);
        let start = self.doc.transact().state_vector().get(&client);
        self.automatic_undo_barrier();
        let mut txn = self.transact_for(context);
        for (i, (source_id, source_key, text, copy)) in copies.into_iter().enumerate() {
            let Some((value, snapshot)) = copy else {
                continue;
            };
            let key = &sources.story_keys[i];
            let story = stories
                .insert(&mut txn, key.as_str(), value)
                .cast::<TextRef>()
                .map_err(|_| EditError::StoryNotFound(key.clone()))?;
            materialize_paragraph_joins(&story, &mut txn, &source_id)?;
            let prefix = format!("para:{source_key}:");
            let mut paragraph_index = 0;
            for diff in story.diff(&txn, YChange::identity) {
                if let Out::YMap(marker) = diff.insert {
                    let old = map_string(&marker, &txn, PARA_ID).unwrap_or_default();
                    let suffix = old.strip_prefix(&prefix).unwrap_or(&old);
                    marker.insert(&mut txn, PARA_ID, format!("para:{key}:{suffix}"));
                    if text.is_some()
                        && let Some(run) = snapshot.paragraphs[paragraph_index].runs.first()
                    {
                        marker.insert(
                            &mut txn,
                            "emptyStyleJson",
                            serde_json::to_string(&run.style)
                                .map_err(|error| EditError::Json(error.to_string()))?,
                        );
                    }
                    paragraph_index += 1;
                }
            }
            if let Some(text) = text {
                let mut offset = 0;
                let mut strings = Vec::new();
                for diff in story.diff(&txn, YChange::identity) {
                    match diff.insert {
                        Out::Any(Any::String(value)) => {
                            let length = value.encode_utf16().count() as u32;
                            strings.push((offset, length));
                            offset += length;
                        }
                        _ => offset += 1,
                    }
                }
                for (start, length) in strings.into_iter().rev() {
                    story.remove_range(&mut txn, start, length);
                }
                let paragraph_count = story.len(&txn);
                if paragraph_count > 1 {
                    story.remove_range(&mut txn, 1, paragraph_count - 1);
                }
                if !text.is_empty() {
                    let style = snapshot
                        .paragraphs
                        .first()
                        .and_then(|paragraph| paragraph.runs.first())
                        .map(|run| run.style.clone())
                        .unwrap_or_default();
                    crate::story::insert_styled_text(&story, &mut txn, 0, text, &style);
                }
                let snapshot = snapshot_story(&story, &txn, key)?;
                new_cell_refs.push(project_paragraph_id(
                    &snapshot.paragraphs[0].id,
                    key,
                    &ids[i],
                ));
            }
        }
        target.insert(&mut txn, format!("tableEdit:{generation}"), parent);
        target.insert(&mut txn, "tableGeneration", generation.as_str());
        target.insert(&mut txn, "tableColumnCount", table.grid.len() as f64);
        target.insert(
            &mut txn,
            "tableStoryKeys",
            Any::Array(
                sources
                    .story_keys
                    .iter()
                    .map(|key| Any::from(key.as_str()))
                    .collect::<Vec<_>>()
                    .into(),
            ),
        );
        target.insert(&mut txn, "graphicJson", graphic_json);
        target.insert(&mut txn, "tableStructureJson", source_json);
        target.insert(
            &mut txn,
            "textStories",
            Any::Array(
                ids.iter()
                    .map(|id| Any::from(id.as_str()))
                    .collect::<Vec<_>>()
                    .into(),
            ),
        );
        target.insert(&mut txn, "width", width as f64);
        target.insert(&mut txn, "height", height as f64);
        drop(txn);
        self.automatic_undo_barrier();
        let end = self.doc.transact().state_vector().get(&client);
        if end > start {
            let mut changes = yrs::IdSet::new();
            changes.insert(yrs::ID::new(client, start), end - start);
            self.undo.borrow_mut().track_table(shape_id, changes);
        }
        self.table_edits
            .borrow_mut()
            .entry(shape_id.to_owned())
            .or_default()
            .push(generation);
        Ok(Some(TableReceipt {
            slide_id: slide_id.into(),
            shape_id: shape_id.into(),
            row_count: table.rows.len(),
            column_count: table.grid.len(),
            new_cell_refs,
        }))
    }
}

fn resize(size: i64, delta: i64, parts: impl Iterator<Item = i64>) -> EditResult<i64> {
    let floor = checked_sum(parts)?;
    Ok(size
        .checked_add(delta)
        .filter(|size| *size <= 1_000_000_000_000_000)
        .map_or(floor, |size| size.max(floor)))
}

fn checked_sum(mut values: impl Iterator<Item = i64>) -> EditResult<i64> {
    values
        .try_fold(0_i64, |sum, value| sum.checked_add(value))
        .filter(|sum| *sum > 0 && *sum <= 1_000_000_000_000_000)
        .ok_or_else(|| invalid("table geometry exceeds the supported range"))
}

pub(crate) fn validate_table(table: &Table) -> EditResult<()> {
    if table.rows.is_empty()
        || table.grid.is_empty()
        || table.grid.iter().any(|w| *w <= 0)
        || table
            .rows
            .iter()
            .any(|row| row.height <= 0 || row.cells.len() != table.grid.len())
    {
        return Err(invalid(
            "table must have a rectangular grid with positive column widths and row heights",
        ));
    }
    if table
        .rows
        .iter()
        .flat_map(|row| &row.cells)
        .any(|cell| cell.merged || cell.grid_span != 1 || cell.row_span != 1)
    {
        return Err(invalid(
            "merged table cells are unsupported for structural edits; unmerge the table first",
        ));
    }
    Ok(())
}

fn find_shape<'a>(shapes: &'a [ShapeSnapshot], id: &str) -> Option<&'a ShapeSnapshot> {
    shapes.iter().find_map(|shape| {
        if shape.id == id {
            Some(shape)
        } else {
            find_shape(&shape.children, id)
        }
    })
}

fn invalid(message: impl Into<String>) -> EditError {
    EditError::InvalidTable(message.into())
}

pub(crate) fn source_cell(shape: &ShapeSnapshot, story_id: &str) -> Option<[usize; 2]> {
    let pptx_parse::TextTarget::TableCell { row, cell } =
        crate::save::text_target(story_id, &shape.id).ok()?
    else {
        return None;
    };
    match &shape.table_structure {
        Some(structure) => structure.rows.get(row)?.cells.get(cell).copied(),
        None => Some([row, cell]),
    }
}

pub(crate) fn validate_structure(shape: &ShapeSnapshot) -> EditResult<()> {
    let Some(structure) = &shape.table_structure else {
        return Ok(());
    };
    let Some(GraphicFrameData::Table(table)) = &shape.graphic else {
        return Err(invalid("table structure requires a table"));
    };
    validate_table(table)?;
    if structure.grid != table.grid
        || structure.columns.len() != table.grid.len()
        || structure.copied_columns.len() != table.grid.len()
        || structure.rows.len() != table.rows.len()
        || structure.story_keys.len() != shape.text_stories.len()
        || structure
            .story_keys
            .iter()
            .collect::<std::collections::HashSet<_>>()
            .len()
            != structure.story_keys.len()
        || shape.text_stories.len() != table.rows.len() * table.grid.len()
        || structure.rows.iter().zip(&table.rows).any(|(source, row)| {
            source.height != row.height
                || source.cells.len() != row.cells.len()
                || source.copied_cells.len() != row.cells.len()
        })
    {
        return Err(invalid("table grid, rows and text stories must agree"));
    }
    for (i, story) in shape.text_stories.iter().enumerate() {
        let expected = format!(
            "story:{}:table:{}:{}:",
            shape.id,
            i / table.grid.len(),
            i % table.grid.len()
        );
        if !story.id.starts_with(&expected) {
            return Err(invalid(
                "table text stories must match current cell positions",
            ));
        }
    }
    Ok(())
}

pub(crate) fn storage_key<T: ReadTxn>(txn: &T, story_id: &str) -> EditResult<String> {
    let Some((shape_id, address)) = story_id
        .strip_prefix("story:")
        .and_then(|id| id.split_once(":table:"))
    else {
        return Ok(story_id.into());
    };
    let shape = shape_ref(txn, shape_id)?;
    let Some(generation) = map_string(&shape, txn, "tableGeneration") else {
        return Ok(story_id.into());
    };
    let missing = || EditError::StoryNotFound(story_id.into());
    let (row, rest) = address.split_once(':').ok_or_else(missing)?;
    let (column, suffix) = rest.split_once(':').ok_or_else(missing)?;
    let row: usize = row.parse().map_err(|_| missing())?;
    let column: usize = column.parse().map_err(|_| missing())?;
    let columns = map_number(&shape, txn, "tableColumnCount")
        .filter(|n| n.fract() == 0.0 && *n > 0.0)
        .ok_or_else(missing)? as usize;
    if suffix != generation
        || column >= columns
        || address != format!("{row}:{column}:{generation}")
    {
        return Err(missing());
    }
    let at = row
        .checked_mul(columns)
        .and_then(|at| at.checked_add(column))
        .ok_or_else(missing)?;
    let Some(Out::Any(Any::Array(keys))) = shape.get(txn, "tableStoryKeys") else {
        return Err(missing());
    };
    match keys.get(at) {
        Some(Any::String(key)) => Ok(key.to_string()),
        _ => Err(missing()),
    }
}

pub(crate) fn project_paragraph_id(id: &str, storage_key: &str, public_id: &str) -> String {
    if storage_key == public_id {
        return id.into();
    }
    match id.strip_prefix(&format!("para:{storage_key}:")) {
        Some(suffix) => format!("para:{public_id}:{suffix}"),
        None => format!("para:{public_id}:id:{id}"),
    }
}

pub(crate) fn validate_history<T: ReadTxn>(txn: &T, shape: &ShapeSnapshot) -> EditResult<()> {
    let Some(structure) = &shape.table_structure else {
        return Ok(());
    };
    let map = shape_ref(txn, &shape.id)?;
    if map_string(&map, txn, "tableGeneration").as_deref() != Some(structure.generation.as_str())
        || map_number(&map, txn, "tableColumnCount") != Some(structure.grid.len() as f64)
        || map_string_array(&map, txn, "tableStoryKeys")? != structure.story_keys
    {
        return Err(invalid("table story storage and geometry must agree"));
    }
    Ok(())
}

#[derive(Debug, Default, PartialEq)]
pub(crate) struct EditLog {
    generation: Option<String>,
    pub(crate) edits: std::collections::BTreeMap<String, String>,
}

impl EditLog {
    /// The current generation and every structural edit it descends from.
    pub(crate) fn lineage(&self) -> std::collections::HashSet<&str> {
        let mut lineage = std::collections::HashSet::new();
        let mut current = self.generation.as_deref();
        while let Some(generation) = current {
            if !lineage.insert(generation) {
                break;
            }
            current = self.edits.get(generation).map(String::as_str);
        }
        lineage
    }
}

pub(crate) fn edit_log(doc: &yrs::Doc, shape_id: &str) -> EditLog {
    let txn = doc.transact();
    let Ok(shape) = shape_ref(&txn, shape_id) else {
        return EditLog::default();
    };
    EditLog {
        generation: map_string(&shape, &txn, "tableGeneration"),
        edits: shape
            .iter(&txn)
            .filter_map(|(key, value)| {
                let id = key.strip_prefix("tableEdit:")?;
                let Out::Any(Any::String(parent)) = value else {
                    return None;
                };
                Some((id.to_owned(), parent.to_string()))
            })
            .collect(),
    }
}
