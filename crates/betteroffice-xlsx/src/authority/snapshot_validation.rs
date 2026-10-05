use std::io::{self, Write};
use std::ops::Bound::{Excluded, Unbounded};

use crate::snapshot::yrs_split::{SnapshotKeys, SnapshotParent};
use crate::snapshot::{SnapshotBudget, SnapshotError, SnapshotResult};

use super::*;

#[derive(Default)]
pub(crate) struct SnapshotValidation {
    phase: u8,
    sheet: usize,
    field: usize,
    after: Option<(u32, u32)>,
    dimension: Option<u32>,
    maps: BTreeMap<SnapshotParent, (usize, &'static str)>,
    sheet_keys: BTreeMap<usize, Arc<str>>,
    sheet_indices: BTreeMap<Arc<str>, usize>,
    pending_cell: Option<(CellRef, bool)>,
    spill_after: Option<(u32, u32)>,
    version: i64,
    base_section: u8,
    base_index: usize,
    base_child: usize,
    metadata_child: usize,
    metadata_ref: usize,
    metadata_bytes: usize,
    style_index: u32,
    style_keys: BTreeMap<u32, String>,
    format_keys: BTreeSet<String>,
    number_formats: BTreeMap<u16, usize>,
    unit_bytes: usize,
}

fn invalid() -> SnapshotError {
    SnapshotError::new("snapshot authority and model disagree")
}

fn allowance(bytes: usize, budget: SnapshotBudget, meter: &mut usize) -> SnapshotResult<usize> {
    if bytes > budget.max_bytes() {
        return Err(SnapshotError::new(
            "snapshot authority validation exceeds advance byte budget",
        ));
    }
    *meter = bytes;
    #[cfg(test)]
    crate::snapshot::step::record(1, bytes);
    Ok(bytes)
}

fn augment(bytes: usize, budget: SnapshotBudget, meter: &mut usize) -> SnapshotResult<()> {
    if bytes > budget.max_bytes() {
        return Err(SnapshotError::new(
            "snapshot authority validation exceeds advance byte budget",
        ));
    }
    #[cfg(test)]
    crate::snapshot::step::record(0, bytes.saturating_sub(*meter));
    *meter = bytes;
    Ok(())
}

struct JsonComparison<'a> {
    bytes: &'a [u8],
    offset: usize,
    limit: usize,
}

impl Write for JsonComparison<'_> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let end = self.offset.saturating_add(bytes.len());
        if end > self.limit || self.bytes.get(self.offset..end) != Some(bytes) {
            return Err(io::Error::other("snapshot metadata differs"));
        }
        self.offset = end;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn same_json(value: &impl serde::Serialize, actual: &str, limit: usize) -> SnapshotResult<()> {
    let mut writer = JsonComparison {
        bytes: actual.as_bytes(),
        offset: 0,
        limit,
    };
    serde_json::to_writer(&mut writer, value).map_err(|_| invalid())?;
    if writer.offset != actual.len() {
        return Err(invalid());
    }
    Ok(())
}

fn parent(map: &MapRef) -> SnapshotParent {
    match map.as_ref().id() {
        BranchID::Root(name) => SnapshotParent::Root(name),
        BranchID::Nested(id) => SnapshotParent::Nested(id.client.get(), id.clock),
    }
}

fn atomic<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> SnapshotResult<Any> {
    match map.get(txn, key) {
        Some(Out::Any(value)) => Ok(value),
        _ => Err(invalid()),
    }
}

fn string(value: &Any) -> SnapshotResult<&str> {
    match value {
        Any::String(value) => Ok(value.as_ref()),
        _ => Err(invalid()),
    }
}

fn content_matches(value: Option<&Any>, cell: Option<&Cell>) -> SnapshotResult<bool> {
    let Some(value) = value else {
        return Ok(cell
            .is_none_or(|cell| cell.formula.is_none() && matches!(cell.value, CellValue::Empty)));
    };
    let values = any_values(value, "cell content").map_err(SnapshotError::new)?;
    let kind = values.first().ok_or_else(invalid)?;
    if any_i64(kind, "cell content kind").map_err(SnapshotError::new)? == 1 {
        if values.len() != 3 {
            return Err(invalid());
        }
        value_from_any(&values[2]).map_err(SnapshotError::new)?;
        return Ok(cell.and_then(|cell| cell.formula.as_deref()) == Some(string(&values[1])?));
    }
    let authored = content_from_any(value).map_err(SnapshotError::new)?;
    Ok(authored_content_equal(Some(&authored), cell))
}

impl SnapshotValidation {
    pub(crate) fn advance(
        &mut self,
        authority: &WorkbookAuthority,
        model: &WorkbookModel,
        keys: &mut SnapshotKeys,
        budget: SnapshotBudget,
    ) -> SnapshotResult<bool> {
        let mut bytes = 0;
        for unit in 0..budget.max_records() {
            if bytes == budget.max_bytes() {
                break;
            }
            self.unit_bytes = 0;
            let remaining = SnapshotBudget::new(1, budget.max_bytes() - bytes)?;
            match self.advance_unit(authority, model, keys, remaining) {
                Ok(true) => return Ok(true),
                Ok(false) => bytes += self.unit_bytes,
                Err(failure)
                    if unit != 0
                        && failure.to_string()
                            == "snapshot authority validation exceeds advance byte budget" =>
                {
                    return Ok(false);
                }
                Err(failure) => return Err(failure),
            }
        }
        Ok(false)
    }

    fn advance_unit(
        &mut self,
        authority: &WorkbookAuthority,
        model: &WorkbookModel,
        keys: &mut SnapshotKeys,
        budget: SnapshotBudget,
    ) -> SnapshotResult<bool> {
        let txn = authority.doc.transact();
        let base = &authority.base;
        let meta = txn.get_map(META).ok_or_else(invalid)?;
        let order = txn.get_array(SHEET_ORDER).ok_or_else(invalid)?;
        let sheets = txn.get_map(SHEETS).ok_or_else(invalid)?;
        let formats = txn.get_map(CELL_FORMATS).ok_or_else(invalid)?;
        if self.phase == 0 {
            allowance(256, budget, &mut self.unit_bytes)?;
            self.version = meta
                .get(&txn, "schemaVersion")
                .and_then(|value| value.cast::<i64>().ok())
                .ok_or_else(invalid)?;
            validate_schema_version(self.version).map_err(SnapshotError::new)?;
            structure_generation(&meta, &txn).map_err(SnapshotError::new)?;
            atomic(&meta, &txn, BASE_FINGERPRINT)?;
            if order.len(&txn) as usize != model.sheets.len()
                || base.date_system != model.date_system
            {
                return Err(invalid());
            }
            self.phase = 8;
            return Ok(false);
        }
        if self.phase == 8 {
            let fingerprint = atomic(&meta, &txn, BASE_FINGERPRINT)?;
            let fingerprint = string(&fingerprint)?;
            let candidate = base
                .fingerprints
                .get(&self.version)
                .and_then(|values| values.get(self.metadata_child))
                .ok_or_else(invalid)?;
            allowance(
                fingerprint.len().max(candidate.len()).saturating_add(128),
                budget,
                &mut self.unit_bytes,
            )?;
            if fingerprint == candidate {
                self.phase = 6;
                self.metadata_child = 0;
            } else {
                self.metadata_child += 1;
            }
            return Ok(false);
        }
        if self.phase == 6 {
            if let Some((id, _)) = model.styles.num_fmts.get(self.base_index) {
                allowance(128, budget, &mut self.unit_bytes)?;
                self.number_formats.entry(*id).or_insert(self.base_index);
                self.base_index += 1;
                return Ok(false);
            }
            self.base_index = 0;
            self.phase = 7;
        }
        if self.phase == 7 {
            let index = self.style_index;
            let format = if index == 0 {
                allowance(2048, budget, &mut self.unit_bytes)?;
                CellFormat::default()
            } else if let Some(xf) = model.styles.cell_xfs.get((index - 1) as usize) {
                snapshot_format(
                    model,
                    xf,
                    &self.number_formats,
                    budget,
                    &mut self.unit_bytes,
                )?
            } else {
                self.phase = 1;
                allowance(64, budget, &mut self.unit_bytes)?;
                return Ok(false);
            };
            let (key, payload) = cell_format_entry(&format).map_err(SnapshotError::new)?;
            if index == 0 && formats.get(&txn, &key) != Some(Out::Any(Any::from(payload.as_str())))
            {
                return Err(invalid());
            }
            self.format_keys.insert(key.clone());
            if index != 0 {
                self.style_keys.insert(index - 1, key);
            }
            self.style_index += 1;
            return Ok(false);
        }
        if self.phase == 1 {
            if let Some(sheet) = model.sheets.get(self.sheet) {
                let Some(Out::Any(Any::String(key))) = order.get(&txn, self.sheet as u32) else {
                    return Err(invalid());
                };
                if key.len() > 64 {
                    return Err(invalid());
                }
                let map = sheets
                    .get(&txn, &key)
                    .and_then(|value| value.cast::<MapRef>().ok())
                    .ok_or_else(invalid)?;
                let required = sheet_schema_keys(self.version);
                if self.field == 0 {
                    allowance(
                        sheet
                            .name
                            .len()
                            .saturating_add(key.len())
                            .saturating_add(128),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    let authored_name = atomic(&map, &txn, NAME)?;
                    let authored_name = string(&authored_name)?;
                    augment(
                        sheet
                            .name
                            .len()
                            .max(authored_name.len())
                            .saturating_add(key.len())
                            .saturating_add(128),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    if authored_name != sheet.name || self.sheet_indices.contains_key(&key) {
                        return Err(invalid());
                    }
                    self.sheet_keys.insert(self.sheet, key.clone());
                    self.sheet_indices.insert(key, self.sheet);
                    self.maps.insert(parent(&map), (self.sheet, "sheet"));
                    self.field += 1;
                    return Ok(false);
                }
                if let Some(&field) = required.get(self.field - 1) {
                    if matches!(field, CONTENTS | STYLES | COL_WIDTHS | ROW_HEIGHTS) {
                        allowance(128, budget, &mut self.unit_bytes)?;
                        let child = nested_map(&map, &txn, field).map_err(SnapshotError::new)?;
                        self.maps.insert(parent(&child), (self.sheet, field));
                    } else {
                        let value = atomic(&map, &txn, field)?;
                        match field {
                            NAME => {
                                allowance(
                                    sheet.name.len().saturating_add(64),
                                    budget,
                                    &mut self.unit_bytes,
                                )?;
                            }
                            FREEZE_PANE => {
                                allowance(128, budget, &mut self.unit_bytes)?;
                                if freeze_pane_from_any(&value).map_err(SnapshotError::new)?
                                    != sheet.freeze_pane
                                {
                                    return Err(invalid());
                                }
                            }
                            MERGES => {
                                let values =
                                    any_values(&value, "merges").map_err(SnapshotError::new)?;
                                allowance(512, budget, &mut self.unit_bytes)?;
                                if values.len() != sheet.merges.len() {
                                    return Err(invalid());
                                }
                                if let Some(merge) = values.get(self.metadata_child) {
                                    let single = Any::Array(vec![merge.clone()].into());
                                    if merges_from_any(&single)
                                        .map_err(SnapshotError::new)?
                                        .first()
                                        != sheet.merges.get(self.metadata_child)
                                    {
                                        return Err(invalid());
                                    }
                                    self.metadata_child += 1;
                                    return Ok(false);
                                }
                                self.metadata_child = 0;
                            }
                            HYPERLINKS => {
                                let json = string(&value)?;
                                if let Some(link) = sheet.hyperlinks.get(self.metadata_child) {
                                    let bytes = [
                                        &link.external_target,
                                        &link.location,
                                        &link.tooltip,
                                        &link.display,
                                    ]
                                    .into_iter()
                                    .flatten()
                                    .fold(
                                        256usize,
                                        |bytes, value| {
                                            bytes.saturating_add(value.len().saturating_mul(6))
                                        },
                                    );
                                    allowance(bytes, budget, &mut self.unit_bytes)?;
                                    self.metadata_bytes = self.metadata_bytes.saturating_add(bytes);
                                    self.metadata_child += 1;
                                    return Ok(false);
                                }
                                allowance(
                                    json.len()
                                        .saturating_add(64)
                                        .max(self.metadata_bytes.saturating_add(64)),
                                    budget,
                                    &mut self.unit_bytes,
                                )?;
                                same_json(&sheet.hyperlinks, json, budget.max_bytes())?;
                                self.metadata_child = 0;
                                self.metadata_bytes = 0;
                            }
                            _ => return Err(invalid()),
                        }
                    }
                    self.field += 1;
                    return Ok(false);
                }
                if self.field == required.len() + 1 {
                    if let Some(value) = map.get(&txn, CHARTS) {
                        let Out::Any(value) = value else {
                            return Err(invalid());
                        };
                        let json = string(&value)?;
                        if let Some(chart) = sheet.charts.get(self.metadata_child) {
                            if self.metadata_ref == 0 {
                                let bytes = chart
                                    .part
                                    .len()
                                    .saturating_add(chart.drawing.len())
                                    .saturating_mul(6)
                                    .saturating_add(512);
                                allowance(bytes, budget, &mut self.unit_bytes)?;
                                self.metadata_bytes = self.metadata_bytes.saturating_add(bytes);
                                self.metadata_ref = 1;
                            } else if let Some(reference) = chart.refs.get(self.metadata_ref - 1) {
                                let bytes = reference
                                    .formula
                                    .len()
                                    .saturating_mul(6)
                                    .saturating_add(128);
                                allowance(bytes, budget, &mut self.unit_bytes)?;
                                self.metadata_bytes = self.metadata_bytes.saturating_add(bytes);
                                self.metadata_ref += 1;
                            } else {
                                allowance(64, budget, &mut self.unit_bytes)?;
                                self.metadata_child += 1;
                                self.metadata_ref = 0;
                            }
                            return Ok(false);
                        }
                        allowance(
                            json.len()
                                .saturating_add(64)
                                .max(self.metadata_bytes.saturating_add(64)),
                            budget,
                            &mut self.unit_bytes,
                        )?;
                        same_json(&sheet.charts, json, budget.max_bytes())?;
                        self.metadata_child = 0;
                        self.metadata_bytes = 0;
                    } else {
                        let charts = base
                            .charts
                            .get(base_sheet_index(&key).unwrap_or(usize::MAX))
                            .map(Vec::as_slice)
                            .unwrap_or_default();
                        if charts.len() != sheet.charts.len() {
                            return Err(invalid());
                        }
                        if let Some(chart) = charts.get(self.metadata_child) {
                            let other = &sheet.charts[self.metadata_child];
                            if self.metadata_ref == 0 {
                                allowance(
                                    chart
                                        .part
                                        .len()
                                        .max(other.part.len())
                                        .saturating_add(
                                            chart.drawing.len().max(other.drawing.len()),
                                        )
                                        .saturating_add(256),
                                    budget,
                                    &mut self.unit_bytes,
                                )?;
                                if chart.part != other.part
                                    || chart.drawing != other.drawing
                                    || chart.anchor_index != other.anchor_index
                                    || chart.anchor != other.anchor
                                    || chart.refs.len() != other.refs.len()
                                {
                                    return Err(invalid());
                                }
                                self.metadata_ref = 1;
                            } else if let Some(reference) = chart.refs.get(self.metadata_ref - 1) {
                                let other = &other.refs[self.metadata_ref - 1];
                                allowance(
                                    reference
                                        .formula
                                        .len()
                                        .max(other.formula.len())
                                        .saturating_add(64),
                                    budget,
                                    &mut self.unit_bytes,
                                )?;
                                if reference != other {
                                    return Err(invalid());
                                }
                                self.metadata_ref += 1;
                            } else {
                                allowance(64, budget, &mut self.unit_bytes)?;
                                self.metadata_child += 1;
                                self.metadata_ref = 0;
                            }
                            return Ok(false);
                        }
                        self.metadata_child = 0;
                        allowance(64, budget, &mut self.unit_bytes)?;
                    }
                    self.field += 1;
                    return Ok(false);
                }
                if self.field == required.len() + 2 {
                    allowance(128, budget, &mut self.unit_bytes)?;
                    if self.version < FREEZE_PANE_SCHEMA_VERSION
                        && base
                            .freeze_panes
                            .get(base_sheet_index(&key).unwrap_or(usize::MAX))
                            .copied()
                            .flatten()
                            != sheet.freeze_pane
                    {
                        return Err(invalid());
                    }
                    self.field += 1;
                    return Ok(false);
                }
                if self.field == required.len() + 3 {
                    if self.version < HYPERLINK_SCHEMA_VERSION {
                        let links = base
                            .hyperlinks
                            .get(base_sheet_index(&key).unwrap_or(usize::MAX))
                            .map(Vec::as_slice)
                            .unwrap_or_default();
                        if links.len() != sheet.hyperlinks.len() {
                            return Err(invalid());
                        }
                        if let Some(link) = links.get(self.metadata_child) {
                            let other = &sheet.hyperlinks[self.metadata_child];
                            let size = |link: &Hyperlink| {
                                [
                                    &link.external_target,
                                    &link.location,
                                    &link.tooltip,
                                    &link.display,
                                ]
                                .into_iter()
                                .flatten()
                                .fold(128usize, |bytes, value| bytes.saturating_add(value.len()))
                            };
                            equal_record(
                                link,
                                other,
                                size(link).max(size(other)),
                                budget,
                                &mut self.unit_bytes,
                            )?;
                            self.metadata_child += 1;
                            return Ok(false);
                        }
                    }
                    allowance(64, budget, &mut self.unit_bytes)?;
                    self.metadata_child = 0;
                    self.field += 1;
                    return Ok(false);
                }
                if (required.len() + 4..=required.len() + 5).contains(&self.field)
                    && self.version < SCHEMA_VERSION
                {
                    let columns = self.field == required.len() + 4;
                    if let Some(hidden) = base
                        .hidden_dimensions
                        .get(base_sheet_index(&key).unwrap_or(usize::MAX))
                    {
                        let values = if columns {
                            &hidden.col_widths
                        } else {
                            &hidden.row_heights
                        };
                        let next = match self.dimension {
                            Some(after) => values.range((Excluded(after), Unbounded)).next(),
                            None => values.first_key_value(),
                        };
                        if let Some((&index, &value)) = next {
                            allowance(128, budget, &mut self.unit_bytes)?;
                            let child = nested_map(
                                &map,
                                &txn,
                                if columns { COL_WIDTHS } else { ROW_HEIGHTS },
                            )
                            .map_err(SnapshotError::new)?;
                            let dimensions = if columns {
                                &sheet.col_widths
                            } else {
                                &sheet.row_heights
                            };
                            if child.get(&txn, &index.to_string()).is_none()
                                && dimensions.get(&index).copied().map(f64::to_bits)
                                    != Some(value.to_bits())
                            {
                                return Err(invalid());
                            }
                            self.dimension = Some(index);
                            return Ok(false);
                        }
                    }
                    allowance(64, budget, &mut self.unit_bytes)?;
                    self.dimension = None;
                    self.field += 1;
                    return Ok(false);
                }
                self.sheet += 1;
                self.field = 0;
                allowance(64, budget, &mut self.unit_bytes)?;
                return Ok(false);
            }
            self.phase = 2;
            self.sheet = 0;
            self.field = 0;
        }
        if let Some((at, authored)) = &self.pending_cell {
            let sheet = model.sheets.get(self.sheet).ok_or_else(invalid)?;
            allowance(64, budget, &mut self.unit_bytes)?;
            if let Some((anchor, range)) = sheet.array_formulas_after(self.spill_after).next() {
                self.spill_after = Some((anchor.row, anchor.col));
                if range.contains(*at) && anchor != *at && !*authored {
                    self.pending_cell = None;
                    self.spill_after = None;
                }
                return Ok(false);
            }
            return Err(invalid());
        }
        if self.phase == 2 {
            if let Some((owner, key)) = keys.first() {
                let owner_bytes = match owner {
                    SnapshotParent::Root(name) => name.len(),
                    _ => 0,
                };
                allowance(
                    key.len().saturating_add(owner_bytes).saturating_add(96),
                    budget,
                    &mut self.unit_bytes,
                )?;
                let Some(branch) = owner
                    .branch()
                    .get_branch(&txn)
                    .filter(|branch| !branch.is_deleted())
                else {
                    keys.pop_first();
                    return Ok(false);
                };
                let map = MapRef::from(branch);
                let value = map.get(&txn, key.as_ref());
                if let Some(value) = value {
                    match owner {
                        SnapshotParent::Root(name) if name.as_ref() == META => {
                            if ![BASE_FINGERPRINT, "schemaVersion", STRUCTURE_GENERATION]
                                .contains(&key.as_ref())
                            {
                                return Err(invalid());
                            }
                        }
                        SnapshotParent::Root(name) if name.as_ref() == SHEETS => {
                            if !self.sheet_indices.contains_key(key.as_ref())
                                || !matches!(value, Out::YMap(_))
                            {
                                return Err(invalid());
                            }
                        }
                        SnapshotParent::Root(name) if name.as_ref() == CELL_FORMATS => {
                            let Out::Any(Any::String(payload)) = value else {
                                return Err(invalid());
                            };
                            if payload.len() > MAX_CELL_FORMAT_BYTES
                                || !self.format_keys.contains(key.as_ref())
                            {
                                return Err(invalid());
                            }
                            augment(
                                key.len()
                                    .saturating_add(owner_bytes)
                                    .saturating_add(payload.len().saturating_mul(8))
                                    .saturating_add(1024),
                                budget,
                                &mut self.unit_bytes,
                            )?;
                            let format: CellFormat =
                                serde_json::from_str(&payload).map_err(|_| invalid())?;
                            let (expected, canonical) =
                                cell_format_entry(&format).map_err(SnapshotError::new)?;
                            if key.as_ref() != expected || payload.as_ref() != canonical {
                                return Err(invalid());
                            }
                        }
                        _ => {
                            if let Some(&(sheet_index, field)) = self.maps.get(owner) {
                                let sheet = &model.sheets[sheet_index];
                                match field {
                                    "sheet" => {
                                        if !sheet_schema_keys(self.version).contains(&key.as_ref())
                                            && !sheet_schema_optional_keys(self.version)
                                                .contains(&key.as_ref())
                                        {
                                            return Err(invalid());
                                        }
                                    }
                                    CONTENTS | STYLES => {
                                        let at = parse_cell_key(key).map_err(SnapshotError::new)?;
                                        let Out::Any(value) = value else {
                                            return Err(invalid());
                                        };
                                        if field == CONTENTS {
                                            let bytes = content_bytes(&value)?;
                                            augment(
                                                bytes
                                                    .saturating_add(key.len())
                                                    .saturating_add(owner_bytes)
                                                    .saturating_add(96),
                                                budget,
                                                &mut self.unit_bytes,
                                            )?;
                                            if !content_matches(Some(&value), sheet.cell(at))? {
                                                self.sheet = sheet_index;
                                                let values = any_values(&value, "cell content")
                                                    .map_err(SnapshotError::new)?;
                                                let formula =
                                                    any_i64(&values[0], "cell content kind")
                                                        .map_err(SnapshotError::new)?
                                                        == 1;
                                                self.pending_cell = Some((at, formula));
                                            }
                                        } else {
                                            let actual = string(&value)?;
                                            let style = sheet
                                                .cell(at)
                                                .and_then(|cell| cell.style)
                                                .ok_or_else(invalid)?;
                                            if actual.len() > 64
                                                || self.style_keys.get(&style).map(String::as_str)
                                                    != Some(actual)
                                                || formats.get(&txn, actual).is_none()
                                            {
                                                return Err(invalid());
                                            }
                                        }
                                    }
                                    COL_WIDTHS | ROW_HEIGHTS => {
                                        let index = key.parse::<u32>().map_err(|_| invalid())?;
                                        let dimensions = if field == COL_WIDTHS {
                                            &sheet.col_widths
                                        } else {
                                            &sheet.row_heights
                                        };
                                        if key.as_ref() != index.to_string()
                                            || value.cast::<f64>().ok().map(f64::to_bits)
                                                != dimensions.get(&index).copied().map(f64::to_bits)
                                        {
                                            return Err(invalid());
                                        }
                                    }
                                    _ => return Err(invalid()),
                                }
                            }
                        }
                    }
                }
                keys.pop_first();
                return Ok(false);
            }
            self.phase = 3;
            self.sheet = 0;
        }
        if self.phase == 3 {
            if let Some(sheet) = model.sheets.get(self.sheet) {
                let map = sheets
                    .get(
                        &txn,
                        self.sheet_keys
                            .get(&self.sheet)
                            .ok_or_else(invalid)?
                            .as_ref(),
                    )
                    .and_then(|value| value.cast::<MapRef>().ok())
                    .ok_or_else(invalid)?;
                let start = self.after.map_or((0, 0), |(row, col)| (row, col + 1));
                let first = sheet.cells_in_range(CellRange {
                    start: CellRef::new(start.0, start.1),
                    end: CellRef::new(start.0, u32::MAX),
                });
                let later = sheet.cells_in_range(CellRange {
                    start: CellRef::new(start.0.saturating_add(1), 0),
                    end: CellRef::new(u32::MAX, u32::MAX),
                });
                if let Some((at, cell)) = first.chain(later).next() {
                    let bytes = cell.formula.as_ref().map_or(0, String::len)
                        + match &cell.value {
                            CellValue::Text { value } => value.len(),
                            _ => 0,
                        };
                    allowance(bytes.saturating_add(256), budget, &mut self.unit_bytes)?;
                    let contents = nested_map(&map, &txn, CONTENTS).map_err(SnapshotError::new)?;
                    let value = contents.get(&txn, &cell_key(at));
                    let actual = match &value {
                        Some(Out::Any(value)) => Some(value),
                        None => None,
                        _ => return Err(invalid()),
                    };
                    if let Some(actual) = actual {
                        augment(
                            content_bytes(actual)?
                                .saturating_add(bytes)
                                .saturating_add(256),
                            budget,
                            &mut self.unit_bytes,
                        )?;
                    }
                    self.after = Some((at.row, at.col));
                    if !content_matches(actual, Some(cell))? {
                        self.pending_cell = Some((at, false));
                    }
                    let styles = nested_map(&map, &txn, STYLES).map_err(SnapshotError::new)?;
                    let actual_style = styles.get(&txn, &cell_key(at));
                    if cell.style.is_some() != actual_style.is_some() {
                        return Err(invalid());
                    }
                    return Ok(false);
                }
                if self.field < 2 {
                    let dimensions = if self.field == 0 {
                        &sheet.col_widths
                    } else {
                        &sheet.row_heights
                    };
                    let next = match self.dimension {
                        Some(after) => dimensions.range((Excluded(after), Unbounded)).next(),
                        None => dimensions.first_key_value(),
                    };
                    if let Some((&index, &value)) = next {
                        allowance(128, budget, &mut self.unit_bytes)?;
                        let field = if self.field == 0 {
                            COL_WIDTHS
                        } else {
                            ROW_HEIGHTS
                        };
                        let child = nested_map(&map, &txn, field).map_err(SnapshotError::new)?;
                        let actual = child
                            .get(&txn, &index.to_string())
                            .and_then(|value| value.cast::<f64>().ok());
                        let fallback = if self.version < SCHEMA_VERSION {
                            base.hidden_dimensions
                                .get(
                                    base_sheet_index(
                                        self.sheet_keys
                                            .get(&self.sheet)
                                            .ok_or_else(invalid)?
                                            .as_ref(),
                                    )
                                    .unwrap_or(usize::MAX),
                                )
                                .and_then(|hidden| {
                                    if self.field == 0 {
                                        hidden.col_widths.get(&index)
                                    } else {
                                        hidden.row_heights.get(&index)
                                    }
                                })
                                .copied()
                        } else {
                            None
                        };
                        if actual.or(fallback).map(f64::to_bits) != Some(value.to_bits()) {
                            return Err(invalid());
                        }
                        self.dimension = Some(index);
                        return Ok(false);
                    }
                    self.field += 1;
                    self.dimension = None;
                    allowance(64, budget, &mut self.unit_bytes)?;
                    return Ok(false);
                }
                self.sheet += 1;
                self.field = 0;
                self.after = None;
                allowance(64, budget, &mut self.unit_bytes)?;
                return Ok(false);
            }
            self.phase = 4;
        }
        if self.phase == 4 {
            if !self.advance_base(base, model, budget)? {
                return Ok(false);
            }
            self.phase = 5;
            return Ok(false);
        }
        if let Some((owner, _)) = self.maps.first_key_value() {
            let bytes = match owner {
                SnapshotParent::Root(name) => name.len(),
                _ => 0,
            };
            allowance(bytes + 96, budget, &mut self.unit_bytes)?;
            self.maps.pop_first();
            return Ok(false);
        }
        if let Some((_, key)) = self.sheet_keys.first_key_value() {
            allowance(key.len() + 128, budget, &mut self.unit_bytes)?;
            self.sheet_keys.pop_first();
            return Ok(false);
        }
        if let Some((key, _)) = self.sheet_indices.first_key_value() {
            allowance(key.len() + 128, budget, &mut self.unit_bytes)?;
            self.sheet_indices.pop_first();
            return Ok(false);
        }
        if let Some((_, key)) = self.style_keys.first_key_value() {
            allowance(key.len() + 128, budget, &mut self.unit_bytes)?;
            self.style_keys.pop_first();
            return Ok(false);
        }
        if let Some(key) = self.format_keys.first() {
            allowance(key.len() + 128, budget, &mut self.unit_bytes)?;
            self.format_keys.pop_first();
            return Ok(false);
        }
        if !self.number_formats.is_empty() {
            allowance(128, budget, &mut self.unit_bytes)?;
            self.number_formats.pop_first();
            return Ok(false);
        }
        Ok(true)
    }
}

fn snapshot_format(
    model: &WorkbookModel,
    xf: &xlsx_model::styles::Xf,
    formats: &BTreeMap<u16, usize>,
    budget: SnapshotBudget,
    meter: &mut usize,
) -> SnapshotResult<CellFormat> {
    let default = CellFormat::default();
    let font = xf
        .font
        .and_then(|index| model.styles.fonts.get(index as usize))
        .unwrap_or(&default.font);
    let fill = xf
        .fill
        .and_then(|index| model.styles.fills.get(index as usize))
        .unwrap_or(&default.fill);
    let border = xf
        .border
        .and_then(|index| model.styles.borders.get(index as usize))
        .unwrap_or(&default.border);
    let id = xf.num_fmt_id.unwrap_or(0);
    let pattern = if id >= 164 {
        formats
            .get(&id)
            .map(|&index| model.styles.num_fmts[index].1.as_str())
    } else {
        None
    };
    let bytes = font
        .name
        .as_ref()
        .map_or(0, String::len)
        .saturating_add(font.color.as_ref().map_or(0, color_bytes))
        .saturating_add(match fill {
            xlsx_model::styles::Fill::Solid(color) => color_bytes(color),
            _ => 0,
        })
        .saturating_add(
            [&border.left, &border.right, &border.top, &border.bottom]
                .into_iter()
                .flatten()
                .map(|edge| edge.color.as_ref().map_or(0, color_bytes))
                .sum::<usize>(),
        )
        .saturating_add(pattern.map_or(0, str::len));
    allowance(bytes.saturating_mul(8).saturating_add(2048), budget, meter)?;
    Ok(CellFormat {
        font: font.clone(),
        fill: fill.clone(),
        border: border.clone(),
        alignment: xf.alignment.clone().unwrap_or_default(),
        number_format: match pattern {
            Some(pattern) => xlsx_model::NumberFormat::Custom {
                pattern: pattern.to_owned(),
            },
            None => xlsx_model::NumberFormat::Builtin { id },
        },
    })
}

fn content_bytes(value: &Any) -> SnapshotResult<usize> {
    let values = any_values(value, "cell content").map_err(SnapshotError::new)?;
    if values.len() > 3 {
        return Err(invalid());
    }
    let mut bytes = 128usize;
    for value in values {
        match value {
            Any::String(value) => bytes = bytes.saturating_add(value.len()),
            Any::Array(values) if values.len() <= 2 => {
                for value in values.iter() {
                    if let Any::String(value) = value {
                        bytes = bytes.saturating_add(value.len());
                    }
                }
            }
            Any::Array(_) | Any::Map(_) => return Err(invalid()),
            _ => {}
        }
    }
    Ok(bytes)
}

fn equal_record<T: super::snapshot::Codec>(
    left: &T,
    right: &T,
    bytes: usize,
    budget: SnapshotBudget,
    meter: &mut usize,
) -> SnapshotResult<()> {
    allowance(bytes.saturating_mul(2).saturating_add(64), budget, meter)?;
    let mut left_bytes = crate::snapshot::wire::Writer::new();
    let mut right_bytes = crate::snapshot::wire::Writer::new();
    left.write(&mut left_bytes);
    right.write(&mut right_bytes);
    if left_bytes.into_bytes() != right_bytes.into_bytes() {
        return Err(invalid());
    }
    Ok(())
}

fn color_bytes(color: &xlsx_model::styles::Color) -> usize {
    match color {
        xlsx_model::styles::Color::Rgb(value) => value.len().saturating_add(32),
        _ => 32,
    }
}

impl SnapshotValidation {
    fn advance_base(
        &mut self,
        base: &WorkbookBase,
        model: &WorkbookModel,
        budget: SnapshotBudget,
    ) -> SnapshotResult<bool> {
        let index = self.base_index;
        match self.base_section {
            0 => {
                allowance(128, budget, &mut self.unit_bytes)?;
                if base.defined_names.len() != model.defined_names.len()
                    || base.shared_strings.len() != model.shared_strings.len()
                    || base.tables.len() != model.tables.len()
                {
                    return Err(invalid());
                }
                self.base_section = 1;
                return Ok(false);
            }
            1 => {
                if let Some(left) = base.defined_names.get(index) {
                    let right = &model.defined_names[index];
                    let bytes = left
                        .name
                        .len()
                        .saturating_add(left.formula.len())
                        .max(right.name.len().saturating_add(right.formula.len()))
                        .saturating_add(64);
                    equal_record(left, right, bytes, budget, &mut self.unit_bytes)?;
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            2 => {
                if let Some(left) = base.shared_strings.get(index) {
                    let right = &model.shared_strings[index];
                    allowance(
                        left.len().max(right.len()).saturating_add(64),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    if left != right {
                        return Err(invalid());
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            3 => {
                if let Some(left) = base.tables.get(index) {
                    let right = &model.tables[index];
                    if self.base_child == 0 {
                        allowance(
                            left.name.len().max(right.name.len()).saturating_add(128),
                            budget,
                            &mut self.unit_bytes,
                        )?;
                        if left.name != right.name
                            || left.sheet != right.sheet
                            || left.range != right.range
                            || left.header_rows != right.header_rows
                            || left.totals_rows != right.totals_rows
                            || left.columns.len() != right.columns.len()
                        {
                            return Err(invalid());
                        }
                        self.base_child = 1;
                    } else if let Some(column) = left.columns.get(self.base_child - 1) {
                        let other = &right.columns[self.base_child - 1];
                        allowance(
                            column.len().max(other.len()).saturating_add(64),
                            budget,
                            &mut self.unit_bytes,
                        )?;
                        if column != other {
                            return Err(invalid());
                        }
                        self.base_child += 1;
                    } else {
                        allowance(64, budget, &mut self.unit_bytes)?;
                        self.base_index += 1;
                        self.base_child = 0;
                    }
                    return Ok(false);
                }
            }
            4 => {
                if let Some(sheet) = model.sheets.get(index) {
                    let key = self.sheet_keys.get(&index).ok_or_else(invalid)?;
                    let default = xlsx_model::SheetFormat::default();
                    let left = base
                        .formats
                        .get(base_sheet_index(key).unwrap_or(usize::MAX))
                        .unwrap_or(&default);
                    equal_record(left, &sheet.format, 32, budget, &mut self.unit_bytes)?;
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            5 => {
                if let Some(sheet) = model.sheets.get(index) {
                    let key = self.sheet_keys.get(&index).ok_or_else(invalid)?;
                    let left = base
                        .col_styles
                        .get(base_sheet_index(key).unwrap_or(usize::MAX))
                        .map(Vec::as_slice)
                        .unwrap_or_default();
                    let right = &sheet.col_styles;
                    if left.len() != right.len() {
                        return Err(invalid());
                    }
                    if let Some(column) = left.get(self.base_child) {
                        equal_record(
                            column,
                            &right[self.base_child],
                            32,
                            budget,
                            &mut self.unit_bytes,
                        )?;
                        self.base_child += 1;
                    } else {
                        allowance(64, budget, &mut self.unit_bytes)?;
                        self.base_index += 1;
                        self.base_child = 0;
                    }
                    return Ok(false);
                }
            }
            6 => {
                if let Some(left) = base.styles.fonts.get(index) {
                    let right = model.styles.fonts.get(index).ok_or_else(invalid)?;
                    let size = |font: &xlsx_model::styles::Font| {
                        font.name
                            .as_ref()
                            .map_or(0, String::len)
                            .saturating_add(font.color.as_ref().map_or(0, color_bytes))
                            .saturating_add(64)
                    };
                    equal_record(
                        left,
                        right,
                        size(left).max(size(right)),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    self.base_index += 1;
                    return Ok(false);
                }
                if base.styles.fonts.len() != model.styles.fonts.len() {
                    return Err(invalid());
                }
            }
            7 => {
                if let Some(left) = base.styles.fills.get(index) {
                    let right = model.styles.fills.get(index).ok_or_else(invalid)?;
                    let size = |fill: &xlsx_model::styles::Fill| match fill {
                        xlsx_model::styles::Fill::Solid(color) => color_bytes(color) + 32,
                        _ => 32,
                    };
                    equal_record(
                        left,
                        right,
                        size(left).max(size(right)),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    self.base_index += 1;
                    return Ok(false);
                }
                if base.styles.fills.len() != model.styles.fills.len() {
                    return Err(invalid());
                }
            }
            8 => {
                if let Some(left) = base.styles.borders.get(index) {
                    let right = model.styles.borders.get(index).ok_or_else(invalid)?;
                    let size = |border: &xlsx_model::styles::Border| {
                        [&border.left, &border.right, &border.top, &border.bottom]
                            .into_iter()
                            .flatten()
                            .fold(64usize, |bytes, edge| {
                                bytes.saturating_add(edge.color.as_ref().map_or(32, color_bytes))
                            })
                    };
                    equal_record(
                        left,
                        right,
                        size(left).max(size(right)),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    self.base_index += 1;
                    return Ok(false);
                }
                if base.styles.borders.len() != model.styles.borders.len() {
                    return Err(invalid());
                }
            }
            9 => {
                if let Some(left) = base.styles.cell_xfs.get(index) {
                    equal_record(
                        left,
                        model.styles.cell_xfs.get(index).ok_or_else(invalid)?,
                        64,
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    self.base_index += 1;
                    return Ok(false);
                }
                if base.styles.cell_xfs.len() != model.styles.cell_xfs.len() {
                    return Err(invalid());
                }
            }
            10 => {
                if let Some(left) = base.styles.num_fmts.get(index) {
                    let right = model.styles.num_fmts.get(index).ok_or_else(invalid)?;
                    equal_record(
                        left,
                        right,
                        left.1.len().max(right.1.len()).saturating_add(32),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    self.base_index += 1;
                    return Ok(false);
                }
                if base.styles.num_fmts.len() != model.styles.num_fmts.len() {
                    return Err(invalid());
                }
            }
            11 => {
                if let Some(left) = base.styles.indexed_colors.get(index) {
                    let right = model.styles.indexed_colors.get(index).ok_or_else(invalid)?;
                    allowance(
                        left.len().max(right.len()).saturating_add(64),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    if left != right {
                        return Err(invalid());
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
                if base.styles.indexed_colors.len() != model.styles.indexed_colors.len() {
                    return Err(invalid());
                }
            }
            12 => {
                if let Some(left) = base.styles.theme.colors.get(index) {
                    let right = &model.styles.theme.colors[index];
                    allowance(
                        left.len().max(right.len()).saturating_add(64),
                        budget,
                        &mut self.unit_bytes,
                    )?;
                    if left != right {
                        return Err(invalid());
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            _ => return Ok(true),
        }
        allowance(64, budget, &mut self.unit_bytes)?;
        self.base_section += 1;
        self.base_index = 0;
        self.base_child = 0;
        Ok(false)
    }
}
