use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::fs;
use std::path::Path;
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::{Context, Result, bail};
use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellValue, DisplayList, GridMeta, SheetId,
    UpdateOrigin, UpdateSubscription, Viewport, Workbook, viewport_for_used_range_within,
};
use sha2::{Digest, Sha256};

const UNIX_EPOCH_SERIAL: f64 = 25_569.0;
const SECONDS_PER_DAY: f64 = 86_400.0;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CellMove {
    Up,
    Down,
    Left,
    Right,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct CellDraft {
    cell: CellRef,
    value: String,
}

pub struct XlsxEditor {
    _local_update_subscription: Option<UpdateSubscription>,
    local_updates: Arc<Mutex<VecDeque<Vec<u8>>>>,
    workbook: Workbook,
    source: Vec<u8>,
    source_fingerprint: Option<[u8; 32]>,
    sheet: SheetId,
    viewport: Viewport,
    display_list: DisplayList,
    selection: CellRef,
    draft: Option<CellDraft>,
    edit_impacts: Vec<BTreeSet<SheetId>>,
    history_index: usize,
    revision_ids: Vec<u64>,
    next_revision: u64,
    saved_revision: u64,
    remote_changed_sheets: BTreeSet<SheetId>,
    remote_dirty: bool,
    #[cfg(test)]
    relayout_count: usize,
}

impl XlsxEditor {
    pub fn open(
        source: Vec<u8>,
        sheet_index: usize,
        max_texture_dimension_2d: u32,
        recalculate: bool,
    ) -> Result<Self> {
        Self::open_internal(
            source,
            sheet_index,
            max_texture_dimension_2d,
            recalculate,
            None,
        )
    }

    pub fn open_collaborative(
        source: Vec<u8>,
        sheet_index: usize,
        max_texture_dimension_2d: u32,
        client_id: u64,
    ) -> Result<Self> {
        Self::open_internal(
            source,
            sheet_index,
            max_texture_dimension_2d,
            true,
            Some(client_id),
        )
    }

    fn open_internal(
        source: Vec<u8>,
        sheet_index: usize,
        max_texture_dimension_2d: u32,
        recalculate: bool,
        collaboration_client_id: Option<u64>,
    ) -> Result<Self> {
        let mut workbook = match collaboration_client_id {
            Some(client_id) => Workbook::open_collaborative(&source, client_id)?,
            None => Workbook::open_for_read(&source)?,
        };
        let sheet_count = workbook.sheet_count();
        if sheet_index >= sheet_count {
            bail!(
                "sheet {} is out of range for a workbook with {sheet_count} sheets",
                sheet_index + 1
            );
        }
        let sheet = SheetId(u32::try_from(sheet_index).context("sheet index is too large")?);
        let viewport = viewport_for_used_range_within(workbook.sheet(sheet)?, |viewport| {
            viewport.width.is_finite()
                && viewport.height.is_finite()
                && viewport.width.ceil() <= max_texture_dimension_2d as f32
                && viewport.height.ceil() <= max_texture_dimension_2d as f32
        });
        let requested_width = viewport.width.ceil();
        let requested_height = viewport.height.ceil();
        if requested_width > max_texture_dimension_2d as f32
            || requested_height > max_texture_dimension_2d as f32
        {
            bail!(
                "requested sheet size {requested_width:.0}x{requested_height:.0} exceeds GPU texture dimension ceiling {max_texture_dimension_2d}"
            );
        }
        if recalculate {
            workbook.recalculate_all(calculation_options());
        }
        let display_list = workbook.display_list_for(sheet, &viewport)?;
        let source_fingerprint = collaboration_client_id.map(|_| workbook_checksum(&workbook));
        let local_updates = Arc::new(Mutex::new(VecDeque::new()));
        let local_update_subscription = if collaboration_client_id.is_some() {
            let observed = Arc::clone(&local_updates);
            Some(workbook.observe_update_v1(move |event| {
                if event.origin == UpdateOrigin::Local {
                    observed
                        .lock()
                        .unwrap_or_else(|poisoned| poisoned.into_inner())
                        .push_back(event.update);
                }
            })?)
        } else {
            None
        };
        Ok(Self {
            _local_update_subscription: local_update_subscription,
            local_updates,
            workbook,
            source,
            source_fingerprint,
            sheet,
            viewport,
            display_list,
            selection: CellRef::new(0, 0),
            draft: None,
            edit_impacts: Vec::new(),
            history_index: 0,
            revision_ids: vec![0],
            next_revision: 1,
            saved_revision: 0,
            remote_changed_sheets: BTreeSet::new(),
            remote_dirty: false,
            #[cfg(test)]
            relayout_count: 0,
        })
    }

    pub(crate) fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    pub fn display_list(&self) -> &DisplayList {
        &self.display_list
    }

    pub(crate) fn sheet(&self) -> SheetId {
        self.sheet
    }

    pub fn sheet_count(&self) -> usize {
        self.workbook.sheet_count()
    }

    pub fn sheet_name(&self) -> Result<&str> {
        Ok(&self.workbook.sheet(self.sheet)?.name)
    }

    pub fn selection(&self) -> CellRef {
        self.selection
    }

    pub fn is_editing(&self) -> bool {
        self.draft.is_some()
    }

    pub fn is_dirty(&self) -> bool {
        self.revision_ids[self.history_index] != self.saved_revision || self.remote_dirty
    }

    pub fn is_remote_only_dirty(&self) -> bool {
        self.remote_dirty && self.revision_ids[self.history_index] == self.saved_revision
    }

    pub fn can_undo(&self) -> bool {
        self.workbook.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.workbook.can_redo()
    }

    pub fn draft_value(&self) -> Option<&str> {
        self.draft.as_ref().map(|draft| draft.value.as_str())
    }

    pub fn state_vector(&self) -> Result<Vec<u8>> {
        if !self.workbook.is_collaborative() {
            bail!("collaboration is unavailable for this workbook");
        }
        Ok(self.workbook.encode_state_vector_v1())
    }

    pub fn encode_diff(&self, state_vector: &[u8]) -> Result<Vec<u8>> {
        Ok(self.workbook.encode_diff_v1(state_vector)?)
    }

    pub fn drain_local_updates(&self) -> Vec<Vec<u8>> {
        self.local_updates
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .drain(..)
            .collect()
    }

    pub fn source_fingerprint(&self) -> Result<[u8; 32]> {
        self.source_fingerprint
            .context("collaboration is unavailable for this workbook")
    }

    pub fn apply_remote_update(&mut self, update: &[u8]) -> Result<bool> {
        let selection = self.selection;
        let draft_cell = self.draft.as_ref().map(|draft| draft.cell);
        let result = self
            .workbook
            .apply_update_v1(update, calculation_options())?;
        self.selection = selection;
        if let (Some(draft), Some(cell)) = (&mut self.draft, draft_cell) {
            draft.cell = cell;
        }
        if !result.applied {
            return Ok(false);
        }
        if result.changed.is_empty() {
            self.remote_changed_sheets
                .extend((0..self.workbook.sheet_count()).map(|index| SheetId(index as u32)));
        } else {
            self.remote_changed_sheets
                .extend(result.changed.iter().map(|address| address.sheet));
        }
        self.remote_dirty = true;
        self.rebuild_display_list()?;
        Ok(true)
    }

    pub fn canonical_checksum(&self) -> Result<[u8; 32]> {
        let mut checksum = Sha256::new();
        checksum.update(b"betteroffice-native-xlsx-workbook-v1\0");
        checksum.update(self.workbook.save()?);
        Ok(checksum.finalize().into())
    }

    #[cfg(test)]
    pub fn relayout_count(&self) -> usize {
        self.relayout_count
    }

    pub fn cell_at_point(&self, x: f32, y: f32) -> Option<CellRef> {
        let cell = grid_cell_at(&self.display_list.grid, x, y)?;
        Some(self.merge_anchor(cell))
    }

    pub fn selection_rect(&self) -> Option<betteroffice_xlsx::Rect> {
        let range = self
            .workbook
            .sheet(self.sheet)
            .ok()
            .and_then(|sheet| {
                sheet.merges.iter().copied().find(|range| {
                    self.selection.row >= range.start.row
                        && self.selection.row <= range.end.row
                        && self.selection.col >= range.start.col
                        && self.selection.col <= range.end.col
                })
            })
            .unwrap_or_else(|| CellRange::new(self.selection, self.selection));
        grid_range_rect(&self.display_list.grid, range)
    }

    pub fn select(&mut self, cell: CellRef) -> bool {
        let cell = self.merge_anchor(cell);
        if self.selection == cell {
            return false;
        }
        self.selection = cell;
        true
    }

    pub fn move_selection(&mut self, direction: CellMove) -> bool {
        let rows = axis_addresses(
            self.display_list.grid.start_row,
            self.display_list.grid.row_indices.as_deref(),
            &self.display_list.grid.row_offsets,
        );
        let columns = axis_addresses(
            self.display_list.grid.start_col,
            self.display_list.grid.col_indices.as_deref(),
            &self.display_list.grid.col_offsets,
        );
        let next = match direction {
            CellMove::Up => CellRef::new(
                adjacent_address(&rows, self.selection.row, false),
                self.selection.col,
            ),
            CellMove::Down => CellRef::new(
                adjacent_address(&rows, self.selection.row, true),
                self.selection.col,
            ),
            CellMove::Left => CellRef::new(
                self.selection.row,
                adjacent_address(&columns, self.selection.col, false),
            ),
            CellMove::Right => CellRef::new(
                self.selection.row,
                adjacent_address(&columns, self.selection.col, true),
            ),
        };
        self.select(next)
    }

    pub fn begin_edit(&mut self, seed: Option<&str>) -> Result<bool> {
        if self.draft.is_some() {
            return Ok(false);
        }
        let value = match seed {
            Some(seed) => seed.to_owned(),
            None => self.workbook.cell(self.sheet, self.selection)?.input,
        };
        self.draft = Some(CellDraft {
            cell: self.selection,
            value,
        });
        Ok(true)
    }

    pub fn insert_text(&mut self, text: &str) -> bool {
        let Some(draft) = &mut self.draft else {
            return false;
        };
        draft.value.push_str(text);
        true
    }

    pub fn backspace(&mut self) -> bool {
        let Some(draft) = &mut self.draft else {
            return false;
        };
        draft.value.pop().is_some()
    }

    pub fn cancel(&mut self) -> bool {
        self.draft.take().is_some()
    }

    pub fn commit(&mut self, movement: Option<CellMove>) -> Result<bool> {
        let Some(draft) = self.draft.clone() else {
            return Ok(false);
        };
        let result =
            self.workbook
                .edit_cell(self.sheet, draft.cell, &draft.value, calculation_options())?;
        self.draft = None;
        self.selection = draft.cell;
        if result.applied {
            let mut impact = BTreeSet::from([self.sheet]);
            impact.extend(result.changed.iter().map(|address| address.sheet));
            self.edit_impacts.truncate(self.history_index);
            self.edit_impacts.push(impact);
            self.revision_ids.truncate(self.history_index + 1);
            self.revision_ids.push(self.next_revision);
            self.next_revision = self
                .next_revision
                .checked_add(1)
                .context("XLSX edit revision overflow")?;
            self.history_index += 1;
        }
        if let Some(movement) = movement {
            self.move_selection(movement);
        }
        self.rebuild_display_list()?;
        Ok(true)
    }

    pub fn undo(&mut self) -> Result<bool> {
        self.draft = None;
        let result = self.workbook.undo(calculation_options())?;
        if result.applied {
            self.history_index = self.history_index.saturating_sub(1);
            self.rebuild_display_list()?;
        }
        Ok(result.applied)
    }

    pub fn redo(&mut self) -> Result<bool> {
        self.draft = None;
        let result = self.workbook.redo(calculation_options())?;
        if result.applied {
            self.history_index = (self.history_index + 1).min(self.edit_impacts.len());
            self.rebuild_display_list()?;
        }
        Ok(result.applied)
    }

    pub fn status(&self) -> Result<String> {
        let value = match &self.draft {
            Some(draft) => draft.value.clone(),
            None => self
                .workbook
                .sheet(self.sheet)?
                .cell(self.selection)
                .map_or_else(String::new, |cell| cell_value_text(&cell.value)),
        };
        Ok(format!("{} = {value}", self.selection.to_a1()))
    }

    pub fn save_to(&mut self, path: &Path) -> Result<()> {
        let bytes = if self.history_index == 0 && self.remote_changed_sheets.is_empty() {
            self.source.clone()
        } else {
            let mut changed_sheets = self.edit_impacts[..self.history_index]
                .iter()
                .flatten()
                .copied()
                .collect::<BTreeSet<_>>();
            changed_sheets.extend(self.remote_changed_sheets.iter().copied());
            restore_untouched_worksheets(
                &self.source,
                &self.workbook.save()?,
                &changed_sheets,
                self.workbook.sheet_count(),
            )?
        };
        ensure_save_fidelity(&self.source, &bytes)?;
        fs::write(path, bytes).with_context(|| format!("write edited XLSX {}", path.display()))?;
        self.saved_revision = self.revision_ids[self.history_index];
        self.remote_dirty = false;
        Ok(())
    }

    pub fn recover_layout(&mut self) -> Result<()> {
        self.rebuild_display_list()
    }

    fn rebuild_display_list(&mut self) -> Result<()> {
        self.display_list = self.workbook.display_list_for(self.sheet, &self.viewport)?;
        #[cfg(test)]
        {
            self.relayout_count += 1;
        }
        Ok(())
    }

    fn merge_anchor(&self, cell: CellRef) -> CellRef {
        self.workbook
            .sheet(self.sheet)
            .ok()
            .and_then(|sheet| {
                sheet.merges.iter().find(|range| {
                    cell.row >= range.start.row
                        && cell.row <= range.end.row
                        && cell.col >= range.start.col
                        && cell.col <= range.end.col
                })
            })
            .map_or(cell, |range| range.start)
    }
}

fn workbook_checksum(workbook: &Workbook) -> [u8; 32] {
    let mut checksum = Sha256::new();
    checksum.update(b"betteroffice-native-xlsx-state-v1\0");
    checksum.update(workbook.encode_state_as_update_v1());
    checksum.finalize().into()
}

fn calculation_options() -> CalculationOptions {
    let seconds = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs_f64();
    CalculationOptions {
        now_serial: Some(seconds / SECONDS_PER_DAY + UNIX_EPOCH_SERIAL),
    }
}

fn cell_value_text(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Number { value } => value.to_string(),
        CellValue::Text { value } => value.clone(),
        CellValue::Bool { value } => {
            if *value {
                "TRUE".to_owned()
            } else {
                "FALSE".to_owned()
            }
        }
        CellValue::Error { value } => value.as_str().to_owned(),
    }
}

fn track_at(offsets: &[f32], target: f32) -> Option<usize> {
    let tracks = offsets.len().checked_sub(1)?;
    if tracks == 0 || target < offsets[0] || target >= offsets[tracks] {
        return None;
    }
    let index = offsets.partition_point(|offset| *offset <= target);
    Some(index.saturating_sub(1).min(tracks - 1))
}

fn grid_cell_at(grid: &GridMeta, x: f32, y: f32) -> Option<CellRef> {
    if !x.is_finite() || !y.is_finite() {
        return None;
    }
    let row = track_at(&grid.row_offsets, y)?;
    let col = track_at(&grid.col_offsets, x)?;
    Some(CellRef::new(
        grid.row_indices
            .as_ref()
            .and_then(|indices| indices.get(row).copied())
            .unwrap_or(grid.start_row + row as u32),
        grid.col_indices
            .as_ref()
            .and_then(|indices| indices.get(col).copied())
            .unwrap_or(grid.start_col + col as u32),
    ))
}

fn local_index(
    start: u32,
    indices: Option<&[u32]>,
    offsets: &[f32],
    address: u32,
) -> Option<usize> {
    match indices {
        Some(indices) => indices.binary_search(&address).ok(),
        None => {
            let local = usize::try_from(address.checked_sub(start)?).ok()?;
            (local + 1 < offsets.len()).then_some(local)
        }
    }
}

fn grid_range_rect(grid: &GridMeta, range: CellRange) -> Option<betteroffice_xlsx::Rect> {
    let left = local_index(
        grid.start_col,
        grid.col_indices.as_deref(),
        &grid.col_offsets,
        range.start.col,
    )?;
    let right = local_index(
        grid.start_col,
        grid.col_indices.as_deref(),
        &grid.col_offsets,
        range.end.col,
    )?;
    let top = local_index(
        grid.start_row,
        grid.row_indices.as_deref(),
        &grid.row_offsets,
        range.start.row,
    )?;
    let bottom = local_index(
        grid.start_row,
        grid.row_indices.as_deref(),
        &grid.row_offsets,
        range.end.row,
    )?;
    Some(betteroffice_xlsx::Rect {
        x: grid.col_offsets[left],
        y: grid.row_offsets[top],
        w: grid.col_offsets[right + 1] - grid.col_offsets[left],
        h: grid.row_offsets[bottom + 1] - grid.row_offsets[top],
    })
}

fn axis_addresses(start: u32, indices: Option<&[u32]>, offsets: &[f32]) -> Vec<u32> {
    let tracks = offsets.len().saturating_sub(1);
    match indices {
        Some(indices) => indices.iter().copied().take(tracks).collect(),
        None => (0..tracks).map(|index| start + index as u32).collect(),
    }
}

fn adjacent_address(addresses: &[u32], current: u32, forward: bool) -> u32 {
    if addresses.is_empty() {
        return current;
    }
    match addresses.binary_search(&current) {
        Ok(index) if forward => addresses[(index + 1).min(addresses.len() - 1)],
        Ok(index) => addresses[index.saturating_sub(1)],
        Err(index) if forward => addresses[index.min(addresses.len() - 1)],
        Err(index) => addresses[index.saturating_sub(1)],
    }
}

#[derive(Debug)]
struct PackageRelationship {
    target: String,
    kind: Option<String>,
    external: bool,
}

impl PackageRelationship {
    fn is_worksheet(&self) -> bool {
        self.kind
            .as_deref()
            .and_then(|kind| kind.rsplit('/').next())
            .is_none_or(|kind| kind == "worksheet")
    }
}

fn restore_untouched_worksheets(
    source: &[u8],
    saved: &[u8],
    changed_sheets: &BTreeSet<SheetId>,
    sheet_count: usize,
) -> Result<Vec<u8>> {
    let source_parts = ooxml_opc::unzip_parts(source).map_err(anyhow::Error::msg)?;
    let source_by_path = source_parts
        .iter()
        .map(|(path, bytes)| (path.to_ascii_lowercase(), bytes))
        .collect::<BTreeMap<_, _>>();
    let sheet_paths = source_sheet_paths(&source_parts)?;
    if sheet_paths.len() != sheet_count {
        bail!(
            "cannot save XLSX safely: found {} source sheet entries for {sheet_count} workbook sheets",
            sheet_paths.len()
        );
    }
    let mut saved_parts = ooxml_opc::unzip_parts(saved).map_err(anyhow::Error::msg)?;
    for (index, path) in sheet_paths.into_iter().enumerate() {
        if changed_sheets.contains(&SheetId(index as u32)) {
            continue;
        }
        let Some(path) = path else {
            continue;
        };
        let normalized = path.to_ascii_lowercase();
        let original = source_by_path
            .get(&normalized)
            .with_context(|| format!("source XLSX is missing worksheet part {path}"))?;
        let projected = saved_parts
            .iter_mut()
            .find(|(candidate, _)| candidate.eq_ignore_ascii_case(&path))
            .with_context(|| format!("saved XLSX is missing worksheet part {path}"))?;
        projected.1.clone_from(original);
    }
    ooxml_opc::rezip_parts(&saved_parts).map_err(anyhow::Error::msg)
}

fn source_sheet_paths(parts: &[(String, Vec<u8>)]) -> Result<Vec<Option<String>>> {
    let workbook =
        part_bytes(parts, "xl/workbook.xml").context("source XLSX is missing xl/workbook.xml")?;
    let workbook_xml = std::str::from_utf8(workbook).context("workbook XML is not UTF-8")?;
    let relationships = part_bytes(parts, "xl/_rels/workbook.xml.rels")
        .map(parse_relationships)
        .transpose()?
        .unwrap_or_default();
    direct_children(parse_tags(workbook_xml)?, "sheets", "sheet")
        .into_iter()
        .enumerate()
        .map(|(index, tag)| {
            let relationship = tag.attribute("id").and_then(|id| relationships.get(id));
            if relationship.is_some_and(|relationship| !relationship.is_worksheet()) {
                return Ok(None);
            }
            let path = relationship
                .filter(|relationship| !relationship.external)
                .map(|relationship| resolve_part_path("xl", &relationship.target))
                .unwrap_or_else(|| format!("xl/worksheets/sheet{}.xml", index + 1));
            Ok(Some(path))
        })
        .collect()
}

fn direct_children(tags: Vec<XmlTag>, parent_name: &str, child_name: &str) -> Vec<XmlTag> {
    let mut stack = Vec::new();
    let mut children = Vec::new();
    for tag in tags {
        if tag.end {
            stack.pop();
            continue;
        }
        if tag.name == child_name && stack.last().is_some_and(|parent| parent == parent_name) {
            let empty = tag.empty;
            let name = tag.name.clone();
            children.push(tag);
            if !empty {
                stack.push(name);
            }
            continue;
        }
        if !tag.empty {
            stack.push(tag.name);
        }
    }
    children
}

fn parse_relationships(bytes: &[u8]) -> Result<BTreeMap<String, PackageRelationship>> {
    let xml = std::str::from_utf8(bytes).context("relationship XML is not UTF-8")?;
    Ok(parse_tags(xml)?
        .into_iter()
        .filter(|tag| !tag.end && tag.name == "Relationship")
        .filter_map(|tag| {
            let id = tag.attribute("Id")?.to_owned();
            let target = tag.attribute("Target")?.to_owned();
            let kind = tag.attribute("Type").map(str::to_owned);
            let external = tag
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("external"));
            Some((
                id,
                PackageRelationship {
                    target,
                    kind,
                    external,
                },
            ))
        })
        .collect())
}

fn part_bytes<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> Option<&'a [u8]> {
    parts
        .iter()
        .find(|(candidate, _)| candidate.eq_ignore_ascii_case(path))
        .map(|(_, bytes)| bytes.as_slice())
}

fn resolve_part_path(base: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_owned();
    }
    let mut segments = base
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>();
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            segment => segments.push(segment),
        }
    }
    segments.join("/")
}

fn ensure_save_fidelity(source: &[u8], saved: &[u8]) -> Result<()> {
    if source == saved {
        return Ok(());
    }
    let source_entries = ooxml_opc::unzip_parts(source).map_err(anyhow::Error::msg)?;
    let original = package_parts(source)?;
    let projected = package_parts(saved)?;
    if let Some(path) = original
        .keys()
        .find(|path| path.starts_with("_xmlsignatures/"))
    {
        bail!(
            "cannot save XLSX safely: {path} contains a package digital signature that an edit would invalidate"
        );
    }
    let worksheet_paths = source_sheet_paths(&source_entries)?
        .into_iter()
        .flatten()
        .map(|path| path.to_ascii_lowercase())
        .collect::<BTreeSet<_>>();
    let shared_strings = workbook_part_paths(&source_entries, "sharedStrings")?;
    let calculation_chains = workbook_part_paths(&source_entries, "calcChain")?;
    for (path, bytes) in &original {
        if projected.get(path).is_some_and(|saved| saved == bytes) {
            continue;
        }
        if worksheet_paths.contains(path) {
            if let Some(loss) = worksheet_projection_loss(bytes)? {
                bail!(
                    "cannot save XLSX safely: {path} contains {loss}, which the cell projection cannot round-trip"
                );
            }
            continue;
        }
        if matches!(
            path.as_str(),
            "[content_types].xml"
                | "_rels/.rels"
                | "xl/workbook.xml"
                | "xl/_rels/workbook.xml.rels"
        ) || shared_strings.contains(path)
            || calculation_chains.contains(path)
        {
            continue;
        }
        bail!(
            "cannot save XLSX safely: unrelated package part {path} changed during the engine round trip"
        );
    }
    Ok(())
}

fn workbook_part_paths(
    parts: &[(String, Vec<u8>)],
    relationship_kind: &str,
) -> Result<BTreeSet<String>> {
    Ok(part_bytes(parts, "xl/_rels/workbook.xml.rels")
        .map(parse_relationships)
        .transpose()?
        .unwrap_or_default()
        .into_values()
        .filter(|relationship| {
            !relationship.external
                && relationship
                    .kind
                    .as_deref()
                    .and_then(|kind| kind.rsplit('/').next())
                    == Some(relationship_kind)
        })
        .map(|relationship| resolve_part_path("xl", &relationship.target).to_ascii_lowercase())
        .collect())
}

fn package_parts(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    Ok(ooxml_opc::unzip_parts(bytes)
        .map_err(anyhow::Error::msg)?
        .into_iter()
        .map(|(path, bytes)| (path.to_ascii_lowercase(), bytes))
        .collect())
}

fn worksheet_projection_loss(bytes: &[u8]) -> Result<Option<&'static str>> {
    let xml = std::str::from_utf8(bytes).context("worksheet XML is not UTF-8")?;
    let tags = parse_tags(xml)?;
    let mut stack: Vec<String> = Vec::new();
    for tag in tags {
        if tag.end {
            stack.pop();
            continue;
        }
        let parent = stack.last().map(String::as_str);
        let in_sheet_data = stack.iter().any(|name| name == "sheetData");
        let in_inline_string = stack.iter().any(|name| name == "is");
        let loss = match tag.name.as_str() {
            "cols" => unexpected_attribute(&tag, &[]).then_some("column metadata"),
            "sheetData" => unexpected_attribute(&tag, &[]).then_some("sheet-data metadata"),
            "mergeCells" => unexpected_attribute(&tag, &["count"]).then_some("merge metadata"),
            "col" if parent == Some("cols") => {
                unexpected_attribute(&tag, &["min", "max", "width", "customWidth", "hidden"])
                    .then_some("column formatting or metadata")
            }
            "row" if in_sheet_data => {
                unexpected_attribute(&tag, &["r", "ht", "customHeight", "hidden"])
                    .then_some("row formatting or metadata")
            }
            "c" if in_sheet_data => {
                if unexpected_attribute(&tag, &["r", "s", "t"]) {
                    Some("cell metadata")
                } else if tag.attribute("t") == Some("d") {
                    Some("ISO date cells")
                } else {
                    None
                }
            }
            "f" if in_sheet_data => formula_loss(&tag),
            "r" | "rPh" | "phoneticPr" if in_inline_string => Some("rich inline strings"),
            "mergeCell" if parent == Some("mergeCells") => {
                unexpected_attribute(&tag, &["ref"]).then_some("merge metadata")
            }
            _ if in_sheet_data
                && !matches!(
                    (parent, tag.name.as_str()),
                    (Some("sheetData"), "row")
                        | (Some("row"), "c")
                        | (Some("c"), "f" | "v" | "is")
                        | (Some("is"), "t")
                ) =>
            {
                Some("sheet-data extension markup")
            }
            _ if parent == Some("cols") && tag.name != "col" => Some("column extension markup"),
            _ if parent == Some("mergeCells") && tag.name != "mergeCell" => {
                Some("merge extension markup")
            }
            _ => None,
        };
        if loss.is_some() {
            return Ok(loss);
        }
        if !tag.empty {
            stack.push(tag.name);
        }
    }
    Ok(None)
}

fn formula_loss(tag: &XmlTag) -> Option<&'static str> {
    match tag.attribute("t") {
        Some("shared") => Some("shared formulas"),
        Some("array") => Some("array formulas"),
        Some("dataTable") => Some("data-table formulas"),
        _ if !tag.attributes.is_empty() => Some("formula metadata"),
        _ => None,
    }
}

fn unexpected_attribute(tag: &XmlTag, allowed: &[&str]) -> bool {
    tag.attributes
        .iter()
        .any(|(name, _)| !allowed.contains(&name.as_str()))
}

#[derive(Debug)]
struct XmlTag {
    name: String,
    attributes: Vec<(String, String)>,
    end: bool,
    empty: bool,
}

impl XmlTag {
    fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(candidate, _)| candidate == name)
            .map(|(_, value)| value.as_str())
    }
}

fn parse_tags(xml: &str) -> Result<Vec<XmlTag>> {
    let bytes = xml.as_bytes();
    let mut tags = Vec::new();
    let mut offset = 0;
    while let Some(relative) = bytes[offset..].iter().position(|byte| *byte == b'<') {
        let start = offset + relative;
        if bytes[start..].starts_with(b"<!--") {
            offset = find_after(bytes, start + 4, b"-->")?;
            continue;
        }
        if bytes[start..].starts_with(b"<![CDATA[") {
            offset = find_after(bytes, start + 9, b"]]>")?;
            continue;
        }
        if bytes[start..].starts_with(b"<?") {
            offset = find_after(bytes, start + 2, b"?>")?;
            continue;
        }
        let end = tag_end(bytes, start + 1)?;
        let body = xml[start + 1..end].trim();
        offset = end + 1;
        if body.is_empty() || body.starts_with('!') {
            continue;
        }
        tags.push(parse_tag(body)?);
    }
    Ok(tags)
}

fn find_after(bytes: &[u8], start: usize, needle: &[u8]) -> Result<usize> {
    bytes[start..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|position| start + position + needle.len())
        .context("unterminated worksheet XML construct")
}

fn tag_end(bytes: &[u8], start: usize) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in bytes.iter().copied().enumerate().skip(start) {
        match (quote, byte) {
            (Some(active), current) if current == active => quote = None,
            (None, b'\'' | b'"') => quote = Some(byte),
            (None, b'>') => return Ok(index),
            _ => {}
        }
    }
    bail!("unterminated worksheet XML tag")
}

fn parse_tag(body: &str) -> Result<XmlTag> {
    let end = body.starts_with('/');
    let mut body = body.trim_start_matches('/').trim();
    let empty = !end && body.ends_with('/');
    if empty {
        body = body[..body.len() - 1].trim_end();
    }
    let name_end = body.find(char::is_whitespace).unwrap_or(body.len());
    let name = local_name(&body[..name_end]).to_owned();
    if name.is_empty() {
        bail!("worksheet XML tag has no name");
    }
    let attributes = if end {
        Vec::new()
    } else {
        parse_attributes(&body[name_end..])?
    };
    Ok(XmlTag {
        name,
        attributes,
        end,
        empty,
    })
}

fn parse_attributes(mut source: &str) -> Result<Vec<(String, String)>> {
    let mut attributes = Vec::new();
    loop {
        source = source.trim_start();
        if source.is_empty() {
            return Ok(attributes);
        }
        let name_end = source
            .find(|character: char| character.is_whitespace() || character == '=')
            .unwrap_or(source.len());
        let name = &source[..name_end];
        source = source[name_end..].trim_start();
        let Some(rest) = source.strip_prefix('=') else {
            bail!("worksheet XML attribute {name} has no value");
        };
        source = rest.trim_start();
        let quote = source
            .chars()
            .next()
            .filter(|quote| matches!(quote, '\'' | '"'))
            .with_context(|| format!("worksheet XML attribute {name} is not quoted"))?;
        source = &source[quote.len_utf8()..];
        let value_end = source
            .find(quote)
            .with_context(|| format!("worksheet XML attribute {name} is unterminated"))?;
        attributes.push((local_name(name).to_owned(), source[..value_end].to_owned()));
        source = &source[value_end + quote.len_utf8()..];
    }
}

fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}
