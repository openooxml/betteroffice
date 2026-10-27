use super::merges::{MAX_MERGE_EDITS_PER_SHEET, MergeIndex};
use super::*;

pub(super) const HYPERLINK_PREFIX: &str = "hyperlink:";
pub(super) type HyperlinkEdits = BTreeMap<String, Option<Hyperlink>>;

pub(super) fn insert_hyperlink_edit(
    edits: &mut HyperlinkEdits,
    key: &str,
    value: &Out,
) -> Result<(), String> {
    if edits.len() >= MAX_MERGE_EDITS_PER_SHEET {
        return Err("too many sheet hyperlink edits".to_string());
    }
    let range_key = key
        .strip_prefix(HYPERLINK_PREFIX)
        .ok_or_else(|| "invalid hyperlink edit prefix".to_string())?;
    if key.len() > 64 {
        return Err("hyperlink edit key is too long".to_string());
    }
    let range =
        CellRange::parse_a1(range_key).map_err(|_| format!("invalid hyperlink edit key {key}"))?;
    if range.to_a1() != range_key
        || range_key.contains('$')
        || range.start.row >= MAX_ROWS
        || range.end.row >= MAX_ROWS
        || range.start.col >= MAX_COLS
        || range.end.col >= MAX_COLS
    {
        return Err(format!(
            "noncanonical or out of bounds hyperlink edit {key}"
        ));
    }
    let link = match value {
        Out::Any(Any::Bool(false)) => None,
        Out::Any(Any::String(json)) => {
            if json.len() > crate::sheet_json::MAX_HYPERLINK_VALUE_BYTES {
                return Err("sheet hyperlink value exceeds its size limit".to_string());
            }
            let mut parsed = decode_hyperlinks(&format!("[{json}]"))?;
            if parsed.len() != 1 {
                return Err(format!("hyperlink edit {key} is not one link"));
            }
            let link = parsed.pop().unwrap();
            if [
                &link.external_target,
                &link.location,
                &link.tooltip,
                &link.display,
            ]
            .into_iter()
            .flatten()
            .any(String::is_empty)
            {
                return Err("sheet hyperlink field must not be empty".to_string());
            }
            if link.range != range || link.range.to_a1() != range_key {
                return Err(format!("hyperlink edit {key} has a different range"));
            }
            Some(link)
        }
        _ => return Err(format!("hyperlink edit {key} is not a link or removal")),
    };
    edits.insert(range_key.to_string(), link);
    Ok(())
}

pub(crate) fn validate_hyperlink_edit_baseline(links: &[Hyperlink]) -> Result<(), String> {
    if links.iter().any(|link| {
        let range = link.range;
        range.start.abs_row || range.start.abs_col || range.end.abs_row || range.end.abs_col
    }) {
        return Err("cannot edit anchored imported hyperlinks".to_string());
    }
    Ok(())
}

pub(super) fn clear_hyperlink_edits(sheet: &MapRef, txn: &mut TransactionMut<'_>) {
    clear_sheet_edits(sheet, txn, HYPERLINK_PREFIX);
}

pub(super) fn sync_hyperlinks(
    sheet: &MapRef,
    txn: &mut TransactionMut<'_>,
    model: &Sheet,
    collaborative: bool,
) -> Result<(), String> {
    if !collaborative {
        let json = serde_json::to_string(&model.hyperlinks)
            .map_err(|error| format!("cannot encode sheet hyperlinks: {error}"))?;
        sheet.try_update(txn, HYPERLINKS, json);
        clear_hyperlink_edits(sheet, txn);
        return Ok(());
    }
    let Some(Out::Any(Any::String(json))) = sheet.get(txn, HYPERLINKS) else {
        return Err("sheet is missing hyperlinks".to_string());
    };
    let baseline = decode_hyperlinks(&json)?;
    validate_hyperlink_edit_baseline(&baseline)?;
    let mut occupied = MergeIndex::default();
    for link in &baseline {
        if occupied.overlaps(1, 0, MAX_COLS - 1, &link.range) {
            return Err("cannot edit overlapping imported hyperlinks".to_string());
        }
        occupied.insert(1, 0, MAX_COLS - 1, &link.range);
    }
    let baseline = baseline
        .into_iter()
        .map(|link| (link.range.to_a1(), link))
        .collect::<BTreeMap<_, _>>();
    let current = model
        .hyperlinks
        .iter()
        .map(|link| (link.range.to_a1(), link))
        .collect::<BTreeMap<_, _>>();
    let mut keys = sheet
        .keys(txn)
        .filter_map(|key| key.strip_prefix(HYPERLINK_PREFIX).map(str::to_string))
        .collect::<BTreeSet<_>>();
    keys.extend(baseline.keys().cloned());
    keys.extend(current.keys().cloned());
    for key in keys {
        let edit_key = format!("{HYPERLINK_PREFIX}{key}");
        match (baseline.get(&key), current.get(&key).copied()) {
            (Some(original), Some(link)) if original == link => {
                sheet.remove(txn, &edit_key);
            }
            (_, Some(link)) => {
                let json = serde_json::to_string(link)
                    .map_err(|error| format!("cannot encode hyperlink: {error}"))?;
                sheet.try_update(txn, edit_key, json);
            }
            (Some(_), None) => {
                sheet.try_update(txn, edit_key, false);
            }
            (None, None) => {
                sheet.remove(txn, &edit_key);
            }
        }
    }
    Ok(())
}

pub(super) fn project_hyperlinks<T: ReadTxn>(
    baseline: Vec<Hyperlink>,
    sheet: &MapRef,
    txn: &T,
) -> Result<Vec<Hyperlink>, String> {
    let mut edits = HyperlinkEdits::new();
    for (key, value) in sheet.iter(txn) {
        if key.starts_with(HYPERLINK_PREFIX) {
            insert_hyperlink_edit(&mut edits, key, &value)?;
        }
    }
    let mut projection = HyperlinkProjection::new(edits);
    let mut links = Vec::new();
    for link in baseline {
        links.extend(projection.baseline(link)?);
    }
    while let Some(link) = projection.next_edit() {
        links.extend(link);
    }
    Ok(links)
}

pub(super) struct HyperlinkProjection {
    edits: HyperlinkEdits,
    changed: bool,
    occupied: MergeIndex,
    baseline_occupied: MergeIndex,
}

impl HyperlinkProjection {
    pub(super) fn new(edits: HyperlinkEdits) -> Self {
        Self {
            changed: !edits.is_empty(),
            edits,
            occupied: MergeIndex::default(),
            baseline_occupied: MergeIndex::default(),
        }
    }

    pub(super) fn baseline(&mut self, link: Hyperlink) -> Result<Option<Hyperlink>, String> {
        if !self.changed {
            return Ok(Some(link));
        }
        validate_hyperlink_edit_baseline(std::slice::from_ref(&link))?;
        if self
            .baseline_occupied
            .overlaps(1, 0, MAX_COLS - 1, &link.range)
        {
            return Err("cannot edit overlapping imported hyperlinks".to_string());
        }
        self.baseline_occupied
            .insert(1, 0, MAX_COLS - 1, &link.range);
        let link = self.edits.remove(&link.range.to_a1()).unwrap_or(Some(link));
        Ok(link.and_then(|link| self.choose(link)))
    }

    pub(super) fn next_edit(&mut self) -> Option<Option<Hyperlink>> {
        let (_, link) = self.edits.pop_first()?;
        Some(link.and_then(|link| self.choose(link)))
    }

    fn choose(&mut self, link: Hyperlink) -> Option<Hyperlink> {
        if self.occupied.overlaps(1, 0, MAX_COLS - 1, &link.range) {
            return None;
        }
        self.occupied.insert(1, 0, MAX_COLS - 1, &link.range);
        Some(link)
    }
}

pub(super) fn hyperlink_metadata(range: CellRange) -> Hyperlink {
    Hyperlink {
        range,
        external_target: None,
        location: None,
        tooltip: None,
        display: None,
    }
}

#[derive(Default)]
pub(super) struct HyperlinkJson {
    baseline: bool,
    offset: usize,
    range: Option<CellRange>,
    field: usize,
    in_string: bool,
    field_bytes: usize,
    destination: bool,
}

impl HyperlinkJson {
    pub(super) fn advance(
        &mut self,
        json: &str,
        limit: usize,
    ) -> Result<Option<CellRange>, String> {
        use crate::sheet_json::{MAX_HYPERLINK_FIELD_BYTES, MAX_HYPERLINK_VALUE_BYTES};

        let invalid = || "sheet hyperlinks are not canonical JSON".to_string();
        if self.offset > MAX_HYPERLINK_VALUE_BYTES {
            return Err("sheet hyperlink value exceeds its size limit".to_string());
        }
        if self.range.is_none() {
            let remaining = json.strip_prefix("{\"range\":").ok_or_else(invalid)?;
            let bytes = remaining.as_bytes();
            let end = bytes.len().min(256);
            let end = bytes[..end]
                .windows(b",\"external_target\":".len())
                .position(|bytes| bytes == b",\"external_target\":")
                .ok_or_else(invalid)?;
            let range: CellRange =
                serde_json::from_str(&remaining[..end]).map_err(|_| invalid())?;
            if range.start.row > range.end.row
                || range.start.col > range.end.col
                || range.end.row >= MAX_ROWS
                || range.end.col >= MAX_COLS
                || !self.baseline
                    && (range.start.abs_row
                        || range.start.abs_col
                        || range.end.abs_row
                        || range.end.abs_col)
                || serde_json::to_string(&range).map_err(|_| invalid())? != remaining[..end]
            {
                return Err(invalid());
            }
            self.offset = json.len() - remaining.len() + end;
            self.range = Some(range);
            return Ok(None);
        }
        if self.field == 4 {
            if json.as_bytes().get(self.offset) != Some(&b'}') || !self.destination {
                return Err(invalid());
            }
            self.offset += 1;
            return Ok(self.range);
        }
        if !self.in_string {
            let prefix = [
                ",\"external_target\":",
                ",\"location\":",
                ",\"tooltip\":",
                ",\"display\":",
            ][self.field];
            let remaining = json.get(self.offset..).ok_or_else(invalid)?;
            let remaining = remaining.strip_prefix(prefix).ok_or_else(invalid)?;
            self.offset += prefix.len();
            if remaining.starts_with("null") {
                self.offset += 4;
                self.field += 1;
            } else if remaining.starts_with('"') {
                self.offset += 1;
                self.field_bytes = 0;
                self.in_string = true;
            } else {
                return Err(invalid());
            }
            return Ok(None);
        }
        let bytes = json.as_bytes();
        let end = bytes.len().min(self.offset.saturating_add(limit));
        while self.offset < end {
            let byte = bytes[self.offset];
            if byte == b'"' {
                if !self.baseline && self.field_bytes == 0 {
                    return Err("sheet hyperlink field must not be empty".to_string());
                }
                self.destination |= self.field < 2 && self.field_bytes != 0;
                self.offset += 1;
                self.field += 1;
                self.in_string = false;
                return Ok(None);
            }
            if byte < 32 {
                return Err(invalid());
            }
            let width = if byte == b'\\' {
                if end - self.offset < 6 && end < bytes.len() {
                    break;
                }
                match bytes.get(self.offset + 1) {
                    Some(b'"' | b'\\' | b'b' | b'f' | b'n' | b'r' | b't') => 2,
                    Some(b'u') => {
                        let escape = bytes
                            .get(self.offset + 2..self.offset + 6)
                            .ok_or_else(invalid)?;
                        if escape[..2] != *b"00" {
                            return Err(invalid());
                        }
                        let digit =
                            |byte| b"0123456789abcdef".iter().position(|&value| value == byte);
                        let value = digit(escape[2])
                            .zip(digit(escape[3]))
                            .map(|(high, low)| high * 16 + low)
                            .ok_or_else(invalid)?;
                        if value >= 32 || matches!(value, 8 | 9 | 10 | 12 | 13) {
                            return Err(invalid());
                        }
                        6
                    }
                    _ => return Err(invalid()),
                }
            } else {
                1
            };
            self.offset += width;
            self.field_bytes += 1;
            if self.field_bytes > MAX_HYPERLINK_FIELD_BYTES {
                return Err("sheet hyperlink field exceeds its length limit".to_string());
            }
        }
        if self.offset == bytes.len() {
            return Err(invalid());
        }
        Ok(None)
    }

    pub(super) fn bytes(&self) -> usize {
        self.offset
    }
}

pub(super) enum HyperlinkRead {
    Pending,
    Link(CellRange, usize, usize),
    Done,
}

#[derive(Default)]
pub(super) struct HyperlinkBaseline {
    offset: usize,
    count: usize,
    value: Option<HyperlinkJson>,
    done: bool,
}

impl HyperlinkBaseline {
    pub(super) fn advance(&mut self, json: &str, limit: usize) -> Result<HyperlinkRead, String> {
        use crate::sheet_json::{MAX_HYPERLINK_STATE_BYTES, MAX_HYPERLINKS_PER_SHEET};

        let invalid = || "sheet hyperlinks are not canonical JSON".to_string();
        if self.done {
            return Ok(HyperlinkRead::Done);
        }
        if json.len() > MAX_HYPERLINK_STATE_BYTES {
            return Err("sheet hyperlink state exceeds its size limit".to_string());
        }
        if self.value.is_none() {
            let separator = if self.count == 0 { b'[' } else { b',' };
            if self.offset == 0 || json.as_bytes().get(self.offset) != Some(&b']') {
                if json.as_bytes().get(self.offset) != Some(&separator) {
                    return Err(invalid());
                }
                self.offset += 1;
                if self.count != 0 && json.as_bytes().get(self.offset) == Some(&b']') {
                    return Err(invalid());
                }
            }
            if json.as_bytes().get(self.offset) == Some(&b']') {
                if self.offset + 1 != json.len() {
                    return Err(invalid());
                }
                self.done = true;
                return Ok(HyperlinkRead::Done);
            }
            if self.count >= MAX_HYPERLINKS_PER_SHEET {
                return Err("sheet has too many hyperlinks".to_string());
            }
            self.value = Some(HyperlinkJson {
                baseline: true,
                ..HyperlinkJson::default()
            });
            return Ok(HyperlinkRead::Pending);
        }
        let value = self.value.as_mut().unwrap();
        let Some(range) = value.advance(&json[self.offset..], limit)? else {
            return Ok(HyperlinkRead::Pending);
        };
        let start = self.offset;
        self.offset += value.bytes();
        self.count += 1;
        self.value = None;
        Ok(HyperlinkRead::Link(range, start, self.offset))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn streamed_hyperlink_json_matches_canonical_serde_and_refuses_other_encodings() {
        for value in ["plain", "é😀", "\"\\\n\r\t\u{0000}\u{0008}\u{000c}\u{001f}"] {
            let mut link = hyperlink_metadata(CellRange::parse_a1("A1:B2").unwrap());
            link.location = Some(value.repeat(100));
            let json = serde_json::to_string(&link).unwrap();
            let mut reader = HyperlinkJson::default();
            loop {
                if let Some(range) = reader.advance(&json, 64).unwrap() {
                    assert_eq!(range, link.range);
                    assert_eq!(reader.bytes(), json.len());
                    break;
                }
            }
        }
        let mut link = hyperlink_metadata(CellRange::parse_a1("A1").unwrap());
        link.location = Some("a".into());
        let canonical = serde_json::to_string(&link).unwrap();
        for json in [
            format!(" {canonical}"),
            canonical.replace("\"a\"", "\"\\u0061\""),
            canonical.replace("\"a\"", "\"\""),
            canonical.replace("\"a\"", "null"),
            canonical.replace("\"a\"", "\"\\u000a\""),
            canonical.replace("\"a\"", "\"\\u001F\""),
        ] {
            let mut reader = HyperlinkJson::default();
            loop {
                match reader.advance(&json, 64) {
                    Err(_) => break,
                    Ok(None) => {}
                    Ok(Some(_)) => panic!("accepted {json}"),
                }
            }
        }
    }
}
