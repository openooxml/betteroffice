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
    order_keys: BTreeMap<usize, Arc<str>>,
    style_components: BTreeSet<(u8, u32)>,
    pending_cell: Option<(CellRef, bool)>,
    spill_after: Option<(u32, u32)>,
    version: i64,
    base_section: u8,
    base_index: usize,
    base_child: usize,
    metadata_child: usize,
    metadata_ref: usize,
    json: JsonPosition,
    record_offset: usize,
    format_digest: Sha256,
    format_key: Option<String>,
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
    crate::snapshot::step::record(1, bytes);
    Ok(bytes)
}

fn augment(bytes: usize, budget: SnapshotBudget, meter: &mut usize) -> SnapshotResult<()> {
    if bytes > budget.max_bytes() {
        return Err(SnapshotError::new(
            "snapshot authority validation exceeds advance byte budget",
        ));
    }
    crate::snapshot::step::record(0, bytes.saturating_sub(*meter));
    *meter = bytes;
    Ok(())
}

#[derive(Default)]
struct JsonPosition {
    token: usize,
    offset: usize,
    bytes: usize,
}

struct JsonWindow<'a> {
    position: &'a mut JsonPosition,
    token: usize,
    bytes: Vec<u8>,
    limit: usize,
}

#[derive(Debug)]
enum JsonStop {
    Pending,
    Invalid,
}

type JsonResult = Result<(), JsonStop>;

impl JsonWindow<'_> {
    fn raw(&mut self, value: &[u8]) -> JsonResult {
        let token = self.token;
        self.token += 1;
        if token < self.position.token {
            return Ok(());
        }
        let start = self.position.offset;
        let count = (value.len() - start).min(self.limit - self.bytes.len());
        self.bytes.extend_from_slice(&value[start..start + count]);
        self.position.bytes += count;
        if start + count != value.len() {
            self.position.offset += count;
            return Err(JsonStop::Pending);
        }
        self.position.token += 1;
        self.position.offset = 0;
        Ok(())
    }

    fn string(&mut self, value: &str) -> JsonResult {
        let token = self.token;
        self.token += 1;
        if token < self.position.token {
            return Ok(());
        }
        while self.position.offset < value.len() + 2 {
            let at = self.position.offset;
            let mut escape = [0u8; 6];
            let bytes: &[u8] = if at == 0 || at == value.len() + 1 {
                b"\""
            } else {
                let byte = value.as_bytes()[at - 1];
                match byte {
                    b'"' => b"\\\"",
                    b'\\' => b"\\\\",
                    b'\n' => b"\\n",
                    b'\r' => b"\\r",
                    b'\t' => b"\\t",
                    8 => b"\\b",
                    12 => b"\\f",
                    0..=31 => {
                        escape[..4].copy_from_slice(b"\\u00");
                        escape[4] = b"0123456789abcdef"[(byte >> 4) as usize];
                        escape[5] = b"0123456789abcdef"[(byte & 15) as usize];
                        &escape
                    }
                    _ => {
                        escape[0] = byte;
                        &escape[..1]
                    }
                }
            };
            if bytes.len() > self.limit - self.bytes.len() {
                return Err(JsonStop::Pending);
            }
            self.bytes.extend_from_slice(bytes);
            self.position.bytes += bytes.len();
            self.position.offset += 1;
        }
        self.position.token += 1;
        self.position.offset = 0;
        Ok(())
    }

    fn small(&mut self, value: &impl serde::Serialize) -> JsonResult {
        if self.token < self.position.token {
            self.token += 1;
            return Ok(());
        }
        let bytes = serde_json::to_vec(value).map_err(|_| JsonStop::Invalid)?;
        self.raw(&bytes)
    }

    fn optional_string(&mut self, value: Option<&str>) -> JsonResult {
        match value {
            Some(value) => self.string(value),
            None => self.raw(b"null"),
        }
    }

    fn color(&mut self, value: &xlsx_model::styles::Color) -> JsonResult {
        if let xlsx_model::styles::Color::Rgb(value) = value {
            self.raw(b"{\"Rgb\":")?;
            self.string(value)?;
            self.raw(b"}")
        } else {
            self.small(value)
        }
    }

    fn optional_color(&mut self, value: Option<&xlsx_model::styles::Color>) -> JsonResult {
        match value {
            Some(value) => self.color(value),
            None => self.raw(b"null"),
        }
    }

    fn format(
        &mut self,
        model: &WorkbookModel,
        xf: Option<&xlsx_model::styles::Xf>,
        formats: &BTreeMap<u16, usize>,
    ) -> JsonResult {
        let default = CellFormat::default();
        let font = xf
            .and_then(|xf| xf.font)
            .and_then(|index| model.styles.fonts.get(index as usize))
            .unwrap_or(&default.font);
        let fill = xf
            .and_then(|xf| xf.fill)
            .and_then(|index| model.styles.fills.get(index as usize))
            .unwrap_or(&default.fill);
        let border = xf
            .and_then(|xf| xf.border)
            .and_then(|index| model.styles.borders.get(index as usize))
            .unwrap_or(&default.border);
        self.raw(b"{\"font\":{\"name\":")?;
        self.optional_string(font.name.as_deref())?;
        self.raw(b",\"size_pt\":")?;
        self.small(&font.size_pt)?;
        for (key, value) in [
            (b",\"bold\":".as_slice(), font.bold),
            (b",\"italic\":".as_slice(), font.italic),
            (b",\"underline\":".as_slice(), font.underline),
            (b",\"strike\":".as_slice(), font.strike),
        ] {
            self.raw(key)?;
            self.small(&value)?;
        }
        self.raw(b",\"color\":")?;
        self.optional_color(font.color.as_ref())?;
        self.raw(b"},\"fill\":")?;
        match fill {
            xlsx_model::styles::Fill::None => self.raw(b"\"None\"")?,
            xlsx_model::styles::Fill::Solid(color) => {
                self.raw(b"{\"Solid\":")?;
                self.color(color)?;
                self.raw(b"}")?;
            }
        }
        self.raw(b",\"border\":{")?;
        for (key, edge) in [
            (b"\"left\":".as_slice(), &border.left),
            (b",\"right\":".as_slice(), &border.right),
            (b",\"top\":".as_slice(), &border.top),
            (b",\"bottom\":".as_slice(), &border.bottom),
        ] {
            self.raw(key)?;
            if let Some(edge) = edge {
                self.raw(b"{\"style\":")?;
                self.small(&edge.style)?;
                self.raw(b",\"color\":")?;
                self.optional_color(edge.color.as_ref())?;
                self.raw(b"}")?;
            } else {
                self.raw(b"null")?;
            }
        }
        self.raw(b"},\"numberFormat\":")?;
        let id = xf.and_then(|xf| xf.num_fmt_id).unwrap_or(0);
        if let Some(&index) = formats.get(&id).filter(|_| id >= 164) {
            self.raw(b"{\"kind\":\"custom\",\"pattern\":")?;
            self.string(&model.styles.num_fmts[index].1)?;
            self.raw(b"}")?;
        } else {
            self.raw(b"{\"kind\":\"builtin\",\"id\":")?;
            self.small(&id)?;
            self.raw(b"}")?;
        }
        self.raw(b",\"alignment\":")?;
        self.small(
            xf.and_then(|xf| xf.alignment.as_ref())
                .unwrap_or(&default.alignment),
        )?;
        self.raw(b"}")
    }

    fn hyperlink(&mut self, link: &Hyperlink) -> JsonResult {
        self.raw(b"{\"range\":")?;
        self.small(&link.range)?;
        for (key, value) in [
            (
                b",\"external_target\":".as_slice(),
                link.external_target.as_deref(),
            ),
            (b",\"location\":".as_slice(), link.location.as_deref()),
            (b",\"tooltip\":".as_slice(), link.tooltip.as_deref()),
            (b",\"display\":".as_slice(), link.display.as_deref()),
        ] {
            self.raw(key)?;
            self.optional_string(value)?;
        }
        self.raw(b"}")
    }

    fn chart_header(&mut self, chart: &SheetChart) -> JsonResult {
        self.raw(b"{\"part\":")?;
        self.string(&chart.part)?;
        self.raw(b",\"drawing\":")?;
        self.string(&chart.drawing)?;
        self.raw(b",\"anchorIndex\":")?;
        self.small(&chart.anchor_index)?;
        self.raw(b",\"anchor\":")?;
        self.small(&chart.anchor)?;
        self.raw(b",\"refs\":[")
    }

    fn chart_ref(&mut self, reference: &ChartRef) -> JsonResult {
        self.raw(b"{\"kind\":")?;
        self.small(&reference.kind)?;
        self.raw(b",\"formula\":")?;
        self.string(&reference.formula)?;
        self.raw(b"}")
    }
}

fn json_window(
    position: &mut JsonPosition,
    budget: SnapshotBudget,
    meter: &mut usize,
    write: impl FnOnce(&mut JsonWindow<'_>) -> JsonResult,
) -> SnapshotResult<(bool, Vec<u8>)> {
    allowance(640, budget, meter)?;
    let limit = ((budget.max_bytes() - 512) / 4).min(4096);
    let mut window = JsonWindow {
        position,
        token: 0,
        bytes: Vec::with_capacity(limit),
        limit,
    };
    let done = match write(&mut window) {
        Ok(()) => true,
        Err(JsonStop::Pending) => false,
        Err(JsonStop::Invalid) => return Err(invalid()),
    };
    augment(640.max(512 + window.bytes.len() * 4), budget, meter)?;
    #[cfg(test)]
    crate::snapshot::step::allocate(512 + limit);
    Ok((done, window.bytes))
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

fn compare_text(
    left: &str,
    right: &str,
    budget: SnapshotBudget,
    meter: &mut usize,
    offset: &mut usize,
) -> SnapshotResult<Option<bool>> {
    if left.len() != right.len() {
        *offset = 0;
        return Ok(Some(false));
    }
    let length = (budget.max_bytes().saturating_sub(*meter) / 2).min(4096);
    if length == 0 && *offset < left.len() {
        return Err(SnapshotError::new(
            "snapshot authority validation exceeds advance byte budget",
        ));
    }
    let end = left.len().min(offset.saturating_add(length));
    augment(*meter + (end - *offset) * 2, budget, meter)?;
    if left.as_bytes()[*offset..end] != right.as_bytes()[*offset..end] {
        *offset = 0;
        return Ok(Some(false));
    }
    *offset = end;
    if end == left.len() {
        *offset = 0;
        Ok(Some(true))
    } else {
        Ok(None)
    }
}

fn content_matches(
    value: Option<&Any>,
    cell: Option<&Cell>,
    budget: SnapshotBudget,
    meter: &mut usize,
    offset: &mut usize,
) -> SnapshotResult<Option<bool>> {
    let Some(value) = value else {
        return Ok(Some(cell.is_none_or(|cell| {
            cell.formula.is_none() && matches!(cell.value, CellValue::Empty)
        })));
    };
    let values = any_values(value, "cell content").map_err(SnapshotError::new)?;
    let kind = values.first().ok_or_else(invalid)?;
    let kind = any_i64(kind, "cell content kind").map_err(SnapshotError::new)?;
    let payload = match kind {
        0 if values.len() == 2 => &values[1],
        1 if values.len() == 3 => &values[2],
        _ => return Err(invalid()),
    };
    let authored = any_values(payload, "cell value").map_err(SnapshotError::new)?;
    let text = if authored.len() == 2
        && any_i64(&authored[0], "cell value kind").map_err(SnapshotError::new)? == 2
    {
        Some(string(&authored[1])?)
    } else {
        None
    };
    if kind == 1 {
        if text.is_none() {
            value_from_any(payload).map_err(SnapshotError::new)?;
        }
        let formula = string(&values[1])?;
        return match cell.and_then(|cell| cell.formula.as_deref()) {
            Some(other) => compare_text(formula, other, budget, meter, offset),
            None => Ok(Some(false)),
        };
    }
    if let Some(text) = text {
        return match cell {
            Some(Cell {
                value: CellValue::Text { value },
                formula: None,
                ..
            }) => compare_text(text, value, budget, meter, offset),
            _ => Ok(Some(false)),
        };
    }
    let value = value_from_any(payload).map_err(SnapshotError::new)?;
    if matches!(value, CellValue::Empty) {
        return Err(invalid());
    }
    Ok(Some(cell.is_some_and(|cell| {
        cell.formula.is_none() && cell.value == value
    })))
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
            self.phase = 9;
            return Ok(false);
        }
        if self.phase == 9 {
            allowance(128, budget, &mut self.unit_bytes)?;
            let index = self.order_keys.len();
            if index < model.sheets.len() {
                let value = order.get(&txn, index as u32).ok_or_else(invalid)?;
                let Out::Any(Any::String(key)) = value else {
                    return Err(invalid());
                };
                if key.len() > 64 {
                    return Err(invalid());
                }
                self.order_keys.insert(index, key);
                return Ok(false);
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
            allowance(128, budget, &mut self.unit_bytes)?;
            let Some(matches) = compare_text(
                fingerprint,
                candidate,
                budget,
                &mut self.unit_bytes,
                &mut self.record_offset,
            )?
            else {
                return Ok(false);
            };
            if matches {
                self.phase = 6;
                self.metadata_child = 0;
            } else {
                self.metadata_child += 1;
            }
            return Ok(false);
        }
        if self.phase == 6 {
            if model
                .styles
                .snapshot_field_counts()
                .into_iter()
                .zip(base.styles.snapshot_field_counts())
                .any(|(model, base)| model > base)
            {
                allowance(128, budget, &mut self.unit_bytes)?;
                return Err(SnapshotError::new(
                    "snapshot authority and model disagree: style tables extend beyond the authority base",
                ));
            }
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
            let xf = if index == 0 {
                None
            } else if let Some(xf) = model.styles.cell_xfs.get((index - 1) as usize) {
                Some(xf)
            } else {
                allowance(64, budget, &mut self.unit_bytes)?;
                self.phase = 10;
                return Ok(false);
            };
            let start = self.json.bytes;
            let (done, bytes) =
                json_window(&mut self.json, budget, &mut self.unit_bytes, |window| {
                    window.format(model, xf, &self.number_formats)
                })?;
            if let Some(key) = &self.format_key {
                let payload = atomic(&formats, &txn, key)?;
                let payload = string(&payload)?;
                if payload.len() > MAX_CELL_FORMAT_BYTES
                    || payload.as_bytes().get(start..self.json.bytes) != Some(bytes.as_slice())
                    || (done && payload.len() != self.json.bytes)
                {
                    return Err(invalid());
                }
            } else {
                self.format_digest.update(&bytes);
            }
            if !done {
                return Ok(false);
            }
            self.json = JsonPosition::default();
            let Some(key) = self.format_key.take() else {
                self.format_key = Some(format!(
                    "{:x}",
                    std::mem::take(&mut self.format_digest).finalize()
                ));
                return Ok(false);
            };
            self.format_keys.insert(key.clone());
            if index != 0 {
                self.style_keys.insert(index - 1, key);
            }
            self.style_index += 1;
            return Ok(false);
        }
        if self.phase == 10 {
            allowance(768, budget, &mut self.unit_bytes)?;
            if let Some(xf) = model.styles.cell_xfs.get(self.base_index) {
                if self.base_index >= base.styles.cell_xfs.len()
                    && (xf
                        .font
                        .is_some_and(|index| index as usize >= model.styles.fonts.len())
                        || xf
                            .fill
                            .is_some_and(|index| index as usize >= model.styles.fills.len())
                        || xf
                            .border
                            .is_some_and(|index| index as usize >= model.styles.borders.len())
                        || xf == &xlsx_model::styles::Xf::default())
                {
                    return Err(invalid());
                }
                for (field, index) in [(0, xf.font), (1, xf.fill), (2, xf.border)] {
                    if let Some(index) = index {
                        self.style_components.insert((field, index));
                    }
                }
                if let Some(id) = xf.num_fmt_id {
                    self.style_components.insert((3, u32::from(id)));
                }
                self.base_index += 1;
            } else {
                self.base_index = 0;
                self.phase = 1;
            }
            return Ok(false);
        }
        if self.phase == 1 {
            if let Some(sheet) = model.sheets.get(self.sheet) {
                let key = self
                    .order_keys
                    .get(&self.sheet)
                    .ok_or_else(invalid)?
                    .clone();
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
                    if authored_name.len() != sheet.name.len() {
                        return Err(invalid());
                    }
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
                                let start = self.json.bytes;
                                let link = sheet.hyperlinks.get(self.metadata_child);
                                let (done, bytes) = json_window(
                                    &mut self.json,
                                    budget,
                                    &mut self.unit_bytes,
                                    |window| {
                                        if let Some(link) = link {
                                            window.raw(if self.metadata_child == 0 {
                                                b"["
                                            } else {
                                                b","
                                            })?;
                                            window.hyperlink(link)
                                        } else {
                                            window.raw(if self.metadata_child == 0 {
                                                b"[]"
                                            } else {
                                                b"]"
                                            })
                                        }
                                    },
                                )?;
                                if json.as_bytes().get(start..self.json.bytes)
                                    != Some(bytes.as_slice())
                                {
                                    return Err(invalid());
                                }
                                if !done {
                                    return Ok(false);
                                }
                                self.json.token = 0;
                                self.json.offset = 0;
                                if link.is_some() {
                                    self.metadata_child += 1;
                                    return Ok(false);
                                }
                                if json.len() != self.json.bytes {
                                    return Err(invalid());
                                }
                                self.metadata_child = 0;
                                self.json = JsonPosition::default();
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
                        let start = self.json.bytes;
                        let chart = sheet.charts.get(self.metadata_child);
                        let (done, bytes) =
                            json_window(&mut self.json, budget, &mut self.unit_bytes, |window| {
                                if let Some(chart) = chart {
                                    if self.metadata_ref == 0 {
                                        window.raw(if self.metadata_child == 0 {
                                            b"["
                                        } else {
                                            b","
                                        })?;
                                        window.chart_header(chart)
                                    } else if let Some(reference) =
                                        chart.refs.get(self.metadata_ref - 1)
                                    {
                                        if self.metadata_ref > 1 {
                                            window.raw(b",")?;
                                        }
                                        window.chart_ref(reference)
                                    } else {
                                        window.raw(b"]}")
                                    }
                                } else {
                                    window.raw(if self.metadata_child == 0 {
                                        b"[]"
                                    } else {
                                        b"]"
                                    })
                                }
                            })?;
                        if json.as_bytes().get(start..self.json.bytes) != Some(bytes.as_slice()) {
                            return Err(invalid());
                        }
                        if !done {
                            return Ok(false);
                        }
                        self.json.token = 0;
                        self.json.offset = 0;
                        if let Some(chart) = chart {
                            if self.metadata_ref <= chart.refs.len() {
                                self.metadata_ref += 1;
                            } else {
                                self.metadata_child += 1;
                                self.metadata_ref = 0;
                            }
                            return Ok(false);
                        }
                        if json.len() != self.json.bytes {
                            return Err(invalid());
                        }
                        self.metadata_child = 0;
                        self.json = JsonPosition::default();
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
                            if self.metadata_ref < 2 {
                                let (left, right) = if self.metadata_ref == 0 {
                                    (&chart.part, &other.part)
                                } else {
                                    (&chart.drawing, &other.drawing)
                                };
                                if !equal_record(
                                    left,
                                    right,
                                    0,
                                    budget,
                                    &mut self.unit_bytes,
                                    &mut self.record_offset,
                                )? {
                                    return Ok(false);
                                }
                                if chart.anchor_index != other.anchor_index
                                    || chart.anchor != other.anchor
                                    || chart.refs.len() != other.refs.len()
                                {
                                    return Err(invalid());
                                }
                                self.metadata_ref += 1;
                            } else if let Some(reference) = chart.refs.get(self.metadata_ref - 2) {
                                let other = &other.refs[self.metadata_ref - 2];
                                if !equal_record(
                                    reference,
                                    other,
                                    0,
                                    budget,
                                    &mut self.unit_bytes,
                                    &mut self.record_offset,
                                )? {
                                    return Ok(false);
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
                            if !equal_record(
                                link,
                                other,
                                size(link).max(size(other)),
                                budget,
                                &mut self.unit_bytes,
                                &mut self.record_offset,
                            )? {
                                return Ok(false);
                            }
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
                let known = match owner {
                    SnapshotParent::Root(name) => {
                        matches!(name.as_ref(), META | SHEETS | CELL_FORMATS)
                    }
                    SnapshotParent::Nested(_, _) => self.maps.contains_key(owner),
                };
                if !known {
                    allowance(96, budget, &mut self.unit_bytes)?;
                    keys.pop_first();
                    return Ok(false);
                }
                if key.len() > 64 {
                    allowance(96, budget, &mut self.unit_bytes)?;
                    return Err(invalid());
                }
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
                            if !matches!(value, Out::YMap(_)) {
                                return Err(invalid());
                            }
                            if !self.sheet_indices.contains_key(key.as_ref()) {
                                return Err(SnapshotError::new(
                                    "snapshot retained sheet map is outside the live sheet order",
                                ));
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
                                            let Some(matches) = content_matches(
                                                Some(&value),
                                                sheet.cell(at),
                                                budget,
                                                &mut self.unit_bytes,
                                                &mut self.record_offset,
                                            )?
                                            else {
                                                return Ok(false);
                                            };
                                            if !matches {
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
                                            if style as usize >= base.styles.cell_xfs.len() {
                                                return Err(SnapshotError::new(
                                                    "snapshot authority and model disagree: cell style index exceeds the authority base",
                                                ));
                                            }
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
                    allowance(256, budget, &mut self.unit_bytes)?;
                    if cell
                        .style
                        .is_some_and(|style| style as usize >= base.styles.cell_xfs.len())
                    {
                        return Err(SnapshotError::new(
                            "snapshot authority and model disagree: cell style index exceeds the authority base",
                        ));
                    }
                    let contents = nested_map(&map, &txn, CONTENTS).map_err(SnapshotError::new)?;
                    let value = contents.get(&txn, &cell_key(at));
                    let actual = match &value {
                        Some(Out::Any(value)) => Some(value),
                        None => None,
                        _ => return Err(invalid()),
                    };
                    let Some(matches) = content_matches(
                        actual,
                        Some(cell),
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )?
                    else {
                        return Ok(false);
                    };
                    self.after = Some((at.row, at.col));
                    if !matches {
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
        if let Some((_, key)) = self.order_keys.last_key_value() {
            allowance(key.len() + 128, budget, &mut self.unit_bytes)?;
            self.order_keys.pop_last();
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
        if !self.style_components.is_empty() {
            allowance(128, budget, &mut self.unit_bytes)?;
            self.style_components.pop_first();
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

fn equal_record<T: super::snapshot::Codec>(
    left: &T,
    right: &T,
    _bytes: usize,
    budget: SnapshotBudget,
    meter: &mut usize,
    offset: &mut usize,
) -> SnapshotResult<bool> {
    allowance(128, budget, meter)?;
    let length = ((budget.max_bytes() - 64) / 2).min(4096);
    let mut left_bytes = crate::snapshot::wire::Writer::window(*offset, length);
    let mut right_bytes = crate::snapshot::wire::Writer::window(*offset, length);
    left.write(&mut left_bytes);
    right.write(&mut right_bytes);
    let total = left_bytes.len();
    if total != right_bytes.len() {
        return Err(invalid());
    }
    let left_bytes = left_bytes.into_bytes();
    let right_bytes = right_bytes.into_bytes();
    augment(128.max(64 + left_bytes.len() * 2), budget, meter)?;
    #[cfg(test)]
    crate::snapshot::step::allocate(length * 2);
    if left_bytes != right_bytes {
        return Err(invalid());
    }
    *offset += left_bytes.len();
    if *offset == total {
        *offset = 0;
        Ok(true)
    } else {
        Ok(false)
    }
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
                    if !equal_record(
                        left,
                        right,
                        bytes,
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            2 => {
                if let Some(left) = base.shared_strings.get(index) {
                    let right = &model.shared_strings[index];
                    if !equal_record(
                        left,
                        right,
                        0,
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            3 => {
                if let Some(left) = base.tables.get(index) {
                    let right = &model.tables[index];
                    if self.base_child == 0 {
                        if !equal_record(
                            &left.name,
                            &right.name,
                            0,
                            budget,
                            &mut self.unit_bytes,
                            &mut self.record_offset,
                        )? {
                            return Ok(false);
                        }
                        if left.sheet != right.sheet
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
                        if !equal_record(
                            column,
                            other,
                            0,
                            budget,
                            &mut self.unit_bytes,
                            &mut self.record_offset,
                        )? {
                            return Ok(false);
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
                    if !equal_record(
                        left,
                        &sheet.format,
                        32,
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
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
                        if !equal_record(
                            column,
                            &right[self.base_child],
                            32,
                            budget,
                            &mut self.unit_bytes,
                            &mut self.record_offset,
                        )? {
                            return Ok(false);
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
                    if !equal_record(
                        left,
                        right,
                        size(left).max(size(right)),
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
                if let Some(value) = model.styles.fonts.get(index) {
                    allowance(128, budget, &mut self.unit_bytes)?;
                    if !self.style_components.contains(&(0, index as u32))
                        && value != &xlsx_model::styles::Font::default()
                    {
                        return Err(invalid());
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            7 => {
                if let Some(left) = base.styles.fills.get(index) {
                    let right = model.styles.fills.get(index).ok_or_else(invalid)?;
                    let size = |fill: &xlsx_model::styles::Fill| match fill {
                        xlsx_model::styles::Fill::Solid(color) => color_bytes(color) + 32,
                        _ => 32,
                    };
                    if !equal_record(
                        left,
                        right,
                        size(left).max(size(right)),
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
                if let Some(value) = model.styles.fills.get(index) {
                    allowance(128, budget, &mut self.unit_bytes)?;
                    if !self.style_components.contains(&(1, index as u32))
                        && value != &xlsx_model::styles::Fill::default()
                    {
                        return Err(invalid());
                    }
                    self.base_index += 1;
                    return Ok(false);
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
                    if !equal_record(
                        left,
                        right,
                        size(left).max(size(right)),
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
                if let Some(value) = model.styles.borders.get(index) {
                    allowance(128, budget, &mut self.unit_bytes)?;
                    if !self.style_components.contains(&(2, index as u32))
                        && value != &xlsx_model::styles::Border::default()
                    {
                        return Err(invalid());
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            9 => {
                if let Some(left) = base.styles.cell_xfs.get(index) {
                    if !equal_record(
                        left,
                        model.styles.cell_xfs.get(index).ok_or_else(invalid)?,
                        64,
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            10 => {
                if let Some(left) = base.styles.num_fmts.get(index) {
                    let right = model.styles.num_fmts.get(index).ok_or_else(invalid)?;
                    if !equal_record(
                        left,
                        right,
                        left.1.len().max(right.1.len()).saturating_add(32),
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
                if let Some((id, _)) = model.styles.num_fmts.get(index) {
                    allowance(128, budget, &mut self.unit_bytes)?;
                    if !self.style_components.contains(&(3, u32::from(*id)))
                        || self.number_formats.get(id) != Some(&index)
                    {
                        return Err(invalid());
                    }
                    self.base_index += 1;
                    return Ok(false);
                }
            }
            11 => {
                if let Some(left) = base.styles.indexed_colors.get(index) {
                    let right = model.styles.indexed_colors.get(index).ok_or_else(invalid)?;
                    if !equal_record(
                        left,
                        right,
                        0,
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
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
                    if !equal_record(
                        left,
                        right,
                        0,
                        budget,
                        &mut self.unit_bytes,
                        &mut self.record_offset,
                    )? {
                        return Ok(false);
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
