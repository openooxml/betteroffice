use std::collections::HashMap;

use yrs::block::{
    BLOCK_GC_REF_NUMBER, BLOCK_ITEM_ANY_REF_NUMBER, BLOCK_ITEM_BINARY_REF_NUMBER,
    BLOCK_ITEM_DELETED_REF_NUMBER, BLOCK_ITEM_DOC_REF_NUMBER, BLOCK_ITEM_EMBED_REF_NUMBER,
    BLOCK_ITEM_FORMAT_REF_NUMBER, BLOCK_ITEM_JSON_REF_NUMBER, BLOCK_ITEM_STRING_REF_NUMBER,
    BLOCK_ITEM_TYPE_REF_NUMBER, BLOCK_SKIP_REF_NUMBER, HAS_ORIGIN, HAS_PARENT_SUB,
    HAS_RIGHT_ORIGIN,
};
use yrs::types::{
    TYPE_REFS_ARRAY, TYPE_REFS_DOC, TYPE_REFS_MAP, TYPE_REFS_TEXT, TYPE_REFS_UNDEFINED,
    TYPE_REFS_XML_ELEMENT, TYPE_REFS_XML_FRAGMENT, TYPE_REFS_XML_HOOK, TYPE_REFS_XML_TEXT,
};

const MAX_CLIENT_ID: u64 = (1_u64 << 53) - 1;
const MAX_CLOCK: u32 = i32::MAX as u32;
const MAX_NESTING: u8 = 64;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SplitError {
    InvalidLimit,
    Malformed,
    UnsupportedContent(u8),
    UnsupportedType(u8),
    JsonLengthMismatch,
    MissingDependency,
    RetainedDeletion,
}

#[derive(Default)]
struct Client {
    clock: u32,
    delete_offset: Option<usize>,
    delete_count: u32,
}

struct Struct {
    len: u32,
    kind: u8,
    dependencies: [Option<(u64, u32)>; 3],
}

pub(crate) fn split_update_v1(
    update: &[u8],
    max_part_bytes: usize,
) -> Result<Vec<Vec<u8>>, SplitError> {
    if max_part_bytes == 0 {
        return Err(SplitError::InvalidLimit);
    }
    let mut scanner = Scanner::new(update);
    let client_count = scanner.count()?;
    let mut clients: HashMap<u64, Client> = HashMap::new();
    let mut previous_client = None;
    let mut parts = Vec::new();
    let mut has_deleted_structs = false;
    for _ in 0..client_count {
        let struct_count = scanner.count()?;
        let client = scanner.client()?;
        let mut clock = scanner.clock()?;
        if struct_count == 0 || previous_client.is_some_and(|previous| previous <= client) {
            return Err(SplitError::Malformed);
        }
        if clock != 0 {
            return Err(SplitError::MissingDependency);
        }
        previous_client = Some(client);
        clients.insert(client, Client::default());
        let mut run_start = scanner.pos;
        let mut run_clock = clock;
        let mut run_count = 0;
        for _ in 0..struct_count {
            let struct_start = scanner.pos;
            let block = scanner.block()?;
            if block.kind == BLOCK_SKIP_REF_NUMBER {
                return Err(SplitError::MissingDependency);
            }
            has_deleted_structs |= matches!(
                block.kind,
                BLOCK_GC_REF_NUMBER | BLOCK_ITEM_DELETED_REF_NUMBER
            );
            for (dependency_client, dependency_clock) in block.dependencies.into_iter().flatten() {
                if !clients
                    .get(&dependency_client)
                    .is_some_and(|known| dependency_clock < known.clock)
                {
                    return Err(SplitError::MissingDependency);
                }
            }
            let next_clock = clock.checked_add(block.len).ok_or(SplitError::Malformed)?;
            if next_clock > MAX_CLOCK {
                return Err(SplitError::Malformed);
            }
            let candidate_count = run_count + 1;
            let candidate_size = run_size(client, run_clock, candidate_count)
                .checked_add(scanner.pos - run_start)
                .ok_or(SplitError::Malformed)?;
            if run_count != 0 && candidate_size > max_part_bytes {
                parts.push(encode_run(
                    client,
                    run_clock,
                    run_count,
                    &update[run_start..struct_start],
                ));
                run_start = struct_start;
                run_clock = clock;
                run_count = 0;
            }
            run_count += 1;
            clock = next_clock;
            clients.get_mut(&client).unwrap().clock = clock;
        }
        parts.push(encode_run(
            client,
            run_clock,
            run_count,
            &update[run_start..scanner.pos],
        ));
    }
    let delete_start = scanner.pos;
    let delete_clients = scanner.count()?;
    for _ in 0..delete_clients {
        let client = scanner.client()?;
        let count = scanner.count()?;
        let known = clients
            .get_mut(&client)
            .ok_or(SplitError::MissingDependency)?;
        if count == 0 || known.delete_offset.is_some() {
            return Err(SplitError::Malformed);
        }
        known.delete_offset = Some(scanner.pos);
        known.delete_count = count;
        let mut previous_end = None;
        for _ in 0..count {
            let (start, end) = scanner.delete_range()?;
            if end > known.clock {
                return Err(SplitError::MissingDependency);
            }
            if previous_end.is_some_and(|previous| start <= previous) {
                return Err(SplitError::Malformed);
            }
            previous_end = Some(end);
        }
    }
    if scanner.pos != update.len() {
        return Err(SplitError::Malformed);
    }
    if delete_clients != 0 || has_deleted_structs {
        validate_deletions(update, &clients)?;
    }
    let delete_set = &update[delete_start..];
    if let Some(last) = parts.last_mut()
        && last
            .len()
            .checked_add(delete_set.len() - 1)
            .is_some_and(|len| len <= max_part_bytes)
    {
        last.pop();
        last.extend_from_slice(delete_set);
    } else if delete_clients != 0 || parts.is_empty() {
        let mut last = Vec::with_capacity(1 + delete_set.len());
        last.push(0);
        last.extend_from_slice(delete_set);
        parts.push(last);
    }
    Ok(parts)
}

fn run_size(client: u64, clock: u32, count: u32) -> usize {
    2 + var_len(u64::from(count)) + var_len(client) + var_len(u64::from(clock))
}

fn encode_run(client: u64, clock: u32, count: u32, structs: &[u8]) -> Vec<u8> {
    let mut part = Vec::with_capacity(run_size(client, clock, count) + structs.len());
    part.push(1);
    write_var(&mut part, u64::from(count));
    write_var(&mut part, client);
    write_var(&mut part, u64::from(clock));
    part.extend_from_slice(structs);
    part.push(0);
    part
}

fn var_len(mut value: u64) -> usize {
    let mut len = 1;
    while value >= 128 {
        len += 1;
        value >>= 7;
    }
    len
}

fn write_var(bytes: &mut Vec<u8>, mut value: u64) {
    while value >= 128 {
        bytes.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    bytes.push(value as u8);
}

fn validate_deletions(update: &[u8], clients: &HashMap<u64, Client>) -> Result<(), SplitError> {
    let mut scanner = Scanner::new(update);
    let count = scanner.count()?;
    for _ in 0..count {
        let count = scanner.count()?;
        let client = scanner.client()?;
        let mut clock = scanner.clock()?;
        let known = &clients[&client];
        let mut deletes = Scanner::new(update);
        deletes.pos = known.delete_offset.unwrap_or(update.len());
        let mut remaining = known.delete_count;
        let mut range = None;
        for _ in 0..count {
            let block = scanner.block()?;
            let end = clock + block.len;
            while range.is_none_or(|(_, end)| end <= clock) && remaining != 0 {
                range = Some(deletes.delete_range()?);
                remaining -= 1;
            }
            let deleted = matches!(
                block.kind,
                BLOCK_GC_REF_NUMBER | BLOCK_ITEM_DELETED_REF_NUMBER
            );
            let overlaps = range.is_some_and(|(start, stop)| start < end && stop > clock);
            let covered = range.is_some_and(|(start, stop)| start <= clock && stop >= end);
            if (deleted && !covered) || (!deleted && overlaps) {
                return Err(SplitError::RetainedDeletion);
            }
            clock = end;
        }
    }
    Ok(())
}

struct Scanner<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> Scanner<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, pos: 0 }
    }

    fn byte(&mut self) -> Result<u8, SplitError> {
        let byte = *self.bytes.get(self.pos).ok_or(SplitError::Malformed)?;
        self.pos += 1;
        Ok(byte)
    }

    fn take(&mut self, len: usize) -> Result<&'a [u8], SplitError> {
        let end = self.pos.checked_add(len).ok_or(SplitError::Malformed)?;
        let bytes = self.bytes.get(self.pos..end).ok_or(SplitError::Malformed)?;
        self.pos = end;
        Ok(bytes)
    }

    fn var(&mut self) -> Result<u64, SplitError> {
        let mut value = 0_u64;
        let mut shift = 0;
        loop {
            let byte = self.byte()?;
            let payload = u64::from(byte & 0x7f);
            if shift >= 64 || payload > (u64::MAX >> shift) {
                return Err(SplitError::Malformed);
            }
            value |= payload << shift;
            if byte & 0x80 == 0 {
                if shift != 0 && payload == 0 {
                    return Err(SplitError::Malformed);
                }
                return Ok(value);
            }
            shift += 7;
        }
    }

    fn var_u32(&mut self) -> Result<u32, SplitError> {
        u32::try_from(self.var()?).map_err(|_| SplitError::Malformed)
    }

    fn count(&mut self) -> Result<u32, SplitError> {
        let count = self.var_u32()?;
        if count as usize > self.bytes.len() - self.pos {
            return Err(SplitError::Malformed);
        }
        Ok(count)
    }

    fn client(&mut self) -> Result<u64, SplitError> {
        let client = self.var()?;
        if client > MAX_CLIENT_ID {
            return Err(SplitError::Malformed);
        }
        Ok(client)
    }

    fn clock(&mut self) -> Result<u32, SplitError> {
        let clock = self.var_u32()?;
        if clock > MAX_CLOCK {
            return Err(SplitError::Malformed);
        }
        Ok(clock)
    }

    fn id(&mut self) -> Result<(u64, u32), SplitError> {
        Ok((self.client()?, self.clock()?))
    }

    fn buffer(&mut self) -> Result<&'a [u8], SplitError> {
        let len = self.var_u32()? as usize;
        self.take(len)
    }

    fn string(&mut self) -> Result<&'a str, SplitError> {
        std::str::from_utf8(self.buffer()?).map_err(|_| SplitError::Malformed)
    }

    fn signed_var(&mut self) -> Result<(), SplitError> {
        let first = self.byte()?;
        let mut value = u64::from(first & 0x3f);
        let mut byte = first;
        let mut shift = 6;
        while byte & 0x80 != 0 {
            byte = self.byte()?;
            let payload = u64::from(byte & 0x7f);
            if shift >= 63 || payload > (i64::MAX as u64 >> shift) {
                return Err(SplitError::Malformed);
            }
            value |= payload << shift;
            if byte & 0x80 == 0 && payload == 0 {
                return Err(SplitError::Malformed);
            }
            shift += 7;
        }
        if first & 0x40 != 0 && value == 0 {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }

    fn any(&mut self, depth: u8) -> Result<(), SplitError> {
        if depth >= MAX_NESTING {
            return Err(SplitError::Malformed);
        }
        match self.byte()? {
            127 | 126 | 121 | 120 => {}
            125 => self.signed_var()?,
            124 => {
                self.take(4)?;
            }
            123 | 122 => {
                self.take(8)?;
            }
            119 => {
                self.string()?;
            }
            tag @ (118 | 117) => {
                let count = usize::try_from(self.var()?).map_err(|_| SplitError::Malformed)?;
                if count > self.bytes.len() - self.pos {
                    return Err(SplitError::Malformed);
                }
                for _ in 0..count {
                    if tag == 118 {
                        self.string()?;
                    }
                    self.any(depth + 1)?;
                }
            }
            116 => {
                self.buffer()?;
            }
            _ => return Err(SplitError::Malformed),
        }
        Ok(())
    }

    fn block(&mut self) -> Result<Struct, SplitError> {
        let info = self.byte()?;
        let kind = info & 0x0f;
        let mut dependencies = [None; 3];
        let len = if info == BLOCK_GC_REF_NUMBER || info == BLOCK_SKIP_REF_NUMBER {
            self.var_u32()?
        } else {
            if info & 0x10 != 0 {
                return Err(SplitError::UnsupportedContent(kind));
            }
            if info & HAS_ORIGIN != 0 {
                dependencies[0] = Some(self.id()?);
            }
            if info & HAS_RIGHT_ORIGIN != 0 {
                dependencies[1] = Some(self.id()?);
            }
            if info & (HAS_ORIGIN | HAS_RIGHT_ORIGIN) == 0 {
                match self.var_u32()? {
                    0 => dependencies[2] = Some(self.id()?),
                    1 => {
                        self.string()?;
                    }
                    _ => return Err(SplitError::Malformed),
                }
                if info & HAS_PARENT_SUB != 0 {
                    self.string()?;
                }
            }
            match kind {
                BLOCK_ITEM_DELETED_REF_NUMBER => self.var_u32()?,
                BLOCK_ITEM_JSON_REF_NUMBER => {
                    let count = self.count()?;
                    for _ in 0..count {
                        self.string()?;
                    }
                    return Err(SplitError::JsonLengthMismatch);
                }
                BLOCK_ITEM_BINARY_REF_NUMBER => {
                    self.buffer()?;
                    1
                }
                BLOCK_ITEM_STRING_REF_NUMBER => {
                    u32::try_from(self.string()?.encode_utf16().count())
                        .map_err(|_| SplitError::Malformed)?
                }
                BLOCK_ITEM_EMBED_REF_NUMBER => {
                    self.json()?;
                    1
                }
                BLOCK_ITEM_FORMAT_REF_NUMBER => {
                    self.string()?;
                    self.json()?;
                    1
                }
                BLOCK_ITEM_TYPE_REF_NUMBER => {
                    match self.byte()? {
                        TYPE_REFS_ARRAY
                        | TYPE_REFS_MAP
                        | TYPE_REFS_TEXT
                        | TYPE_REFS_XML_FRAGMENT
                        | TYPE_REFS_XML_HOOK
                        | TYPE_REFS_XML_TEXT
                        | TYPE_REFS_DOC
                        | TYPE_REFS_UNDEFINED => {}
                        TYPE_REFS_XML_ELEMENT => {
                            self.string()?;
                        }
                        kind => return Err(SplitError::UnsupportedType(kind)),
                    }
                    1
                }
                BLOCK_ITEM_ANY_REF_NUMBER => {
                    let len = self.count()?;
                    for _ in 0..len {
                        self.any(0)?;
                    }
                    len
                }
                BLOCK_ITEM_DOC_REF_NUMBER => {
                    self.string()?;
                    self.any(0)?;
                    1
                }
                _ => return Err(SplitError::UnsupportedContent(kind)),
            }
        };
        if len == 0 {
            return Err(SplitError::Malformed);
        }
        Ok(Struct {
            len,
            kind,
            dependencies,
        })
    }

    fn delete_range(&mut self) -> Result<(u32, u32), SplitError> {
        let clock = self.clock()?;
        let len = self.var_u32()?;
        let end = clock.checked_add(len).ok_or(SplitError::Malformed)?;
        if len == 0 || end > MAX_CLOCK {
            return Err(SplitError::Malformed);
        }
        Ok((clock, end))
    }

    fn json(&mut self) -> Result<(), SplitError> {
        let mut json = Json {
            scanner: Scanner::new(self.string()?.as_bytes()),
        };
        json.value(0)?;
        json.space();
        if json.scanner.pos != json.scanner.bytes.len() {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }
}

struct Json<'a> {
    scanner: Scanner<'a>,
}

impl Json<'_> {
    fn peek(&self) -> Option<u8> {
        self.scanner.bytes.get(self.scanner.pos).copied()
    }

    fn consume(&mut self, byte: u8) -> bool {
        if self.peek() == Some(byte) {
            self.scanner.pos += 1;
            true
        } else {
            false
        }
    }

    fn space(&mut self) {
        while matches!(self.peek(), Some(b' ' | b'\r' | b'\n' | b'\t')) {
            self.scanner.pos += 1;
        }
    }

    fn quoted(&mut self) -> Result<(), SplitError> {
        if !self.consume(b'"') {
            return Err(SplitError::Malformed);
        }
        loop {
            match self.scanner.byte()? {
                b'"' => return Ok(()),
                b'\\' => match self.scanner.byte()? {
                    b'"' | b'\\' | b'/' | b'b' | b'f' | b'n' | b'r' | b't' => {}
                    b'u' => {
                        let code = self.hex()?;
                        if (0xd800..=0xdbff).contains(&code) {
                            if self.scanner.take(2)? != b"\\u"
                                || !(0xdc00..=0xdfff).contains(&self.hex()?)
                            {
                                return Err(SplitError::Malformed);
                            }
                        } else if (0xdc00..=0xdfff).contains(&code) {
                            return Err(SplitError::Malformed);
                        }
                    }
                    _ => return Err(SplitError::Malformed),
                },
                0..=31 => return Err(SplitError::Malformed),
                _ => {}
            }
        }
    }

    fn hex(&mut self) -> Result<u16, SplitError> {
        let bytes = self.scanner.take(4)?;
        if !bytes.iter().all(u8::is_ascii_hexdigit) {
            return Err(SplitError::Malformed);
        }
        let text = std::str::from_utf8(bytes).map_err(|_| SplitError::Malformed)?;
        u16::from_str_radix(text, 16).map_err(|_| SplitError::Malformed)
    }

    fn digits(&mut self) -> Result<(), SplitError> {
        let start = self.scanner.pos;
        while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
            self.scanner.pos += 1;
        }
        if self.scanner.pos == start {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }

    fn number(&mut self) -> Result<(), SplitError> {
        let start = self.scanner.pos;
        self.consume(b'-');
        if !self.consume(b'0') {
            self.digits()?;
        }
        if self.consume(b'.') {
            self.digits()?;
        }
        if self.consume(b'e') || self.consume(b'E') {
            if !self.consume(b'+') {
                self.consume(b'-');
            }
            self.digits()?;
        }
        let text = std::str::from_utf8(&self.scanner.bytes[start..self.scanner.pos])
            .map_err(|_| SplitError::Malformed)?;
        let value = text.parse::<f64>().map_err(|_| SplitError::Malformed)?;
        if !value.is_finite() {
            return Err(SplitError::Malformed);
        }
        Ok(())
    }

    fn value(&mut self, depth: u8) -> Result<(), SplitError> {
        if depth >= MAX_NESTING {
            return Err(SplitError::Malformed);
        }
        self.space();
        match self.peek() {
            Some(b'"') => self.quoted()?,
            Some(b'-' | b'0'..=b'9') => self.number()?,
            Some(b'n' | b't' | b'f') => {
                let expected: &[u8] = match self.peek() {
                    Some(b'n') => b"null",
                    Some(b't') => b"true",
                    _ => b"false",
                };
                if self.scanner.take(expected.len())? != expected {
                    return Err(SplitError::Malformed);
                }
            }
            Some(open @ (b'[' | b'{')) => {
                self.scanner.pos += 1;
                let close = if open == b'[' { b']' } else { b'}' };
                self.space();
                if self.consume(close) {
                    return Ok(());
                }
                loop {
                    if open == b'{' {
                        self.quoted()?;
                        self.space();
                        if !self.consume(b':') {
                            return Err(SplitError::Malformed);
                        }
                    }
                    self.value(depth + 1)?;
                    self.space();
                    if self.consume(close) {
                        break;
                    }
                    if !self.consume(b',') {
                        return Err(SplitError::Malformed);
                    }
                    self.space();
                }
            }
            _ => return Err(SplitError::Malformed),
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::authority::hydrate_snapshot_part;
    use crate::{
        CalculationOptions, Cell, CellRange, CellRef, CellValue, DefinedName, Sheet, SheetId,
        Workbook, WorkbookModel,
    };
    use xlsx_model::{CellFormat, NumberFormat};
    use yrs::encoding::write::Write;
    use yrs::updates::decoder::Decode;
    use yrs::updates::encoder::{Encoder, EncoderV1};
    use yrs::{Any, Doc, Map, MapPrelim, ReadTxn, StateVector, Transact, Update};

    const LIMITS: [usize; 5] = [1, 64, 4096, 65536, usize::MAX];

    fn small_model() -> WorkbookModel {
        let mut sheet = Sheet::new("Small");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Text {
                    value: "hello".into(),
                },
                ..Cell::default()
            },
        );
        WorkbookModel {
            sheets: vec![sheet],
            ..WorkbookModel::default()
        }
    }

    fn rich_model() -> WorkbookModel {
        let mut model = WorkbookModel {
            shared_strings: vec!["shared 🦀".into(), "second".into()],
            ..WorkbookModel::default()
        };
        let mut format = CellFormat::default();
        format.font.bold = true;
        format.alignment.wrap_text = true;
        format.number_format = NumberFormat::Custom {
            pattern: "0.000".into(),
        };
        let style = model.styles.intern_cell_format(&format).unwrap();
        for index in 0..3 {
            let mut sheet = Sheet::new(format!("Data{index}"));
            for (col, value) in [
                CellValue::Text {
                    value: model.shared_strings[0].clone(),
                },
                CellValue::Number {
                    value: 12.25 + f64::from(index),
                },
                CellValue::Bool {
                    value: index % 2 == 0,
                },
                CellValue::Number {
                    value: 24.5 + f64::from(index) * 2.0,
                },
            ]
            .into_iter()
            .enumerate()
            {
                sheet.set_cell(
                    CellRef::new(0, col as u32),
                    Cell {
                        value,
                        formula: (col == 3).then(|| "B1*2".into()),
                        style,
                    },
                );
            }
            sheet
                .merges
                .push(CellRange::new(CellRef::new(2, 0), CellRef::new(2, 2)));
            sheet.col_widths.insert(0, 24.0);
            sheet.row_heights.insert(0, 28.0);
            model.sheets.push(sheet);
        }
        model.defined_names.push(DefinedName {
            name: "Amount".into(),
            formula: "Data0!$B$1".into(),
            local_sheet: None,
            hidden: false,
        });
        model.defined_names.push(DefinedName {
            name: "LocalAmount".into(),
            formula: "Data1!$D$1".into(),
            local_sheet: Some(SheetId(1)),
            hidden: true,
        });
        model
    }

    fn append_structs(
        update: &[u8],
        output: &mut Vec<u8>,
        clocks: &mut HashMap<u64, u32>,
    ) -> (usize, usize) {
        let mut scanner = Scanner::new(update);
        let clients = scanner.count().unwrap();
        let mut structs = 0;
        for _ in 0..clients {
            let count = scanner.count().unwrap();
            let client = scanner.client().unwrap();
            let mut clock = scanner.clock().unwrap();
            assert_eq!(clock, clocks.get(&client).copied().unwrap_or_default());
            for _ in 0..count {
                let start = scanner.pos;
                let block = scanner.block().unwrap();
                output.extend_from_slice(&update[start..scanner.pos]);
                clock += block.len;
                structs += 1;
            }
            clocks.insert(client, clock);
        }
        (scanner.pos, structs)
    }

    fn vector_bytes(vector: &StateVector) -> Vec<u8> {
        let mut entries = vector
            .iter()
            .map(|(client, clock)| (client.get(), *clock))
            .collect::<Vec<_>>();
        entries.sort_unstable();
        let mut encoder = EncoderV1::new();
        encoder.write_var(entries.len());
        for (client, clock) in entries {
            encoder.write_var(client);
            encoder.write_var(clock);
        }
        encoder.to_vec()
    }

    fn assert_parts(update: &[u8], vector: &[u8], client_id: u64, limit: usize, parts: &[Vec<u8>]) {
        assert!(!parts.is_empty());
        let mut original_structs = Vec::new();
        let (delete_start, _) = append_structs(update, &mut original_structs, &mut HashMap::new());
        let mut split_structs = Vec::with_capacity(original_structs.len());
        let mut clocks = HashMap::new();
        let doc = Doc::with_client_id(client_id);
        for (index, part) in parts.iter().enumerate() {
            let (delete_offset, count) = append_structs(part, &mut split_structs, &mut clocks);
            assert!(part.len() <= limit || count <= 1);
            if index + 1 == parts.len() {
                assert_eq!(&part[delete_offset..], &update[delete_start..]);
            } else {
                assert_eq!(&part[delete_offset..], &[0]);
            }
            hydrate_snapshot_part(&doc, part).unwrap();
            let txn = doc.transact();
            assert!(
                txn.store().pending_update().is_none(),
                "part {index} has missing structs"
            );
            assert!(
                txn.store().pending_ds().is_none(),
                "part {index} has missing deletes"
            );
        }
        assert_eq!(split_structs, original_structs);
        let txn = doc.transact();
        assert_eq!(txn.state_vector(), StateVector::decode_v1(vector).unwrap());
        assert_eq!(vector_bytes(&txn.state_vector()), vector);
        assert_eq!(
            txn.encode_state_as_update_v1(&StateVector::default()),
            update
        );
    }

    fn assert_workbook(workbook: &Workbook, allow_refusal: bool) {
        let update = workbook.encode_state_as_update_v1();
        let vector = workbook.encode_state_vector_v1();
        for limit in LIMITS {
            match split_update_v1(&update, limit) {
                Ok(parts) => assert_parts(&update, &vector, workbook.client_id(), limit, &parts),
                Err(error) => {
                    assert!(allow_refusal, "split at {limit} failed: {error:?}");
                    assert!(matches!(
                        error,
                        SplitError::MissingDependency | SplitError::RetainedDeletion
                    ));
                }
            }
        }
    }

    fn assert_doc(doc: &Doc) {
        let txn = doc.transact();
        let update = txn.encode_state_as_update_v1(&StateVector::default());
        let vector = vector_bytes(&txn.state_vector());
        drop(txn);
        for limit in LIMITS {
            let parts = split_update_v1(&update, limit).unwrap();
            assert_parts(&update, &vector, doc.client_id().get(), limit, &parts);
        }
    }

    #[test]
    fn small_sheet_snapshot_is_exact() {
        let workbook = Workbook::from_model(small_model()).unwrap();
        assert_workbook(&workbook, false);
    }

    #[test]
    fn rich_workbook_snapshot_is_exact() {
        let saved = Workbook::from_model(rich_model()).unwrap().save().unwrap();
        let workbook = Workbook::open(&saved).unwrap();
        assert_eq!(workbook.sheet_count(), 3);
        assert!(!workbook.model().shared_strings.is_empty());
        assert_eq!(workbook.model().defined_names.len(), 2);
        assert_workbook(&workbook, false);
    }

    #[test]
    fn twenty_thousand_cell_snapshot_is_exact() {
        let mut sheet = Sheet::new("Large");
        for row in 0..1000 {
            for col in 0..20 {
                sheet.set_cell(
                    CellRef::new(row, col),
                    Cell {
                        value: CellValue::Number {
                            value: f64::from(row * 20 + col),
                        },
                        ..Cell::default()
                    },
                );
            }
        }
        assert_eq!(sheet.iter_cells().count(), 20_000);
        let workbook = Workbook::from_model(WorkbookModel {
            sheets: vec![sheet],
            ..WorkbookModel::default()
        })
        .unwrap();
        assert_workbook(&workbook, false);
    }

    #[test]
    fn edited_and_undone_snapshot_is_exact_or_refused() {
        let mut workbook = Workbook::from_model_collaborative(small_model(), 1).unwrap();
        let options = CalculationOptions::default();
        for (col, input) in [(0, "changed"), (1, "42"), (0, "again")] {
            workbook
                .edit_cell(SheetId(0), CellRef::new(0, col), input, options)
                .unwrap();
        }
        assert!(workbook.can_undo());
        workbook.undo(options).unwrap();
        let update = workbook.encode_state_as_update_v1();
        let (delete_start, _) = append_structs(&update, &mut Vec::new(), &mut HashMap::new());
        assert!(Scanner::new(&update[delete_start..]).count().unwrap() > 0);
        assert_eq!(
            StateVector::decode_v1(&workbook.encode_state_vector_v1())
                .unwrap()
                .len(),
            2
        );
        assert_workbook(&workbook, true);
    }

    #[test]
    fn deleted_structs_keep_the_original_delete_set() {
        let doc = Doc::with_client_id(7);
        let map = doc.get_or_insert_map("map");
        {
            let mut txn = doc.transact_mut();
            map.insert(&mut txn, "first", "old");
            let child = map.insert(&mut txn, "child", MapPrelim::default());
            child.insert(&mut txn, "nested", "garbage collect");
            map.insert(&mut txn, "last", "keep");
        }
        {
            let mut txn = doc.transact_mut();
            map.insert(&mut txn, "first", "new");
            map.remove(&mut txn, "child");
        }
        assert_doc(&doc);
    }

    #[test]
    fn ordered_client_dependencies_integrate_immediately() {
        let source = Doc::with_client_id(9);
        let map = source.get_or_insert_map("map");
        map.insert(&mut source.transact_mut(), "key", "first");
        let update = source
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let peer = Doc::with_client_id(1);
        hydrate_snapshot_part(&peer, &update).unwrap();
        let map = peer.get_or_insert_map("map");
        map.insert(&mut peer.transact_mut(), "key", "second");
        assert_doc(&peer);
    }

    #[test]
    fn forward_client_dependencies_are_refused() {
        let source = Doc::with_client_id(1);
        let map = source.get_or_insert_map("map");
        map.insert(&mut source.transact_mut(), "key", "first");
        let update = source
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let peer = Doc::with_client_id(9);
        hydrate_snapshot_part(&peer, &update).unwrap();
        let map = peer.get_or_insert_map("map");
        map.insert(&mut peer.transact_mut(), "key", "second");
        let update = peer
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        for limit in LIMITS {
            assert_eq!(
                split_update_v1(&update, limit),
                Err(SplitError::MissingDependency)
            );
        }
    }

    fn raw_update(count: u32, content: impl FnOnce(&mut EncoderV1)) -> Vec<u8> {
        let mut encoder = EncoderV1::new();
        encoder.write_var(1_u32);
        encoder.write_var(count);
        encoder.write_var(7_u64);
        encoder.write_var(0_u32);
        content(&mut encoder);
        encoder.write_var(0_u32);
        encoder.to_vec()
    }

    fn item(encoder: &mut EncoderV1, kind: u8) {
        encoder.write_info(kind);
        encoder.write_parent_info(true);
        encoder.write_string("root");
    }

    fn assert_boundaries(update: &[u8], clock: u32) {
        for limit in LIMITS {
            let parts = split_update_v1(update, limit).unwrap();
            let mut original = Vec::new();
            append_structs(update, &mut original, &mut HashMap::new());
            let mut actual = Vec::new();
            let mut clocks = HashMap::new();
            for part in parts {
                Update::decode_v1(&part).unwrap();
                append_structs(&part, &mut actual, &mut clocks);
            }
            assert_eq!(actual, original);
            assert_eq!(clocks[&7], clock);
        }
    }

    #[test]
    fn content_boundaries_and_utf16_clocks_are_preserved() {
        let update = raw_update(14, |encoder| {
            item(encoder, BLOCK_ITEM_BINARY_REF_NUMBER);
            encoder.write_buf([0, 1, 128, 255]);
            item(encoder, BLOCK_ITEM_STRING_REF_NUMBER);
            encoder.write_string("a🦀é");
            item(encoder, BLOCK_ITEM_EMBED_REF_NUMBER);
            encoder.write_string(r#"{"x":[null,true,false,-1.25e2,"\uD83E\uDD80"]}"#);
            item(encoder, BLOCK_ITEM_FORMAT_REF_NUMBER);
            encoder.write_key("bold");
            encoder.write_json(&Any::Bool(true));
            for kind in [0_u8, 1, 2, 3, 4, 5, 6, 9, 15] {
                item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
                encoder.write_type_ref(kind);
                if kind == 3 {
                    encoder.write_key("element");
                }
            }
            item(encoder, BLOCK_ITEM_DOC_REF_NUMBER);
            encoder.write_string("subdoc");
            encoder.write_any(&Any::Map(std::sync::Arc::new(HashMap::new())));
        });
        assert_boundaries(&update, 17);
    }

    #[test]
    fn item_origins_and_parent_subtitles_are_scanned() {
        let update = raw_update(5, |encoder| {
            item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
            encoder.write_type_ref(TYPE_REFS_MAP);
            encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER | HAS_PARENT_SUB);
            encoder.write_parent_info(false);
            encoder.write_var(7_u64);
            encoder.write_var(0_u32);
            encoder.write_string("key");
            encoder.write_len(1);
            encoder.write_any(&Any::from("first"));
            for flags in [HAS_ORIGIN, HAS_RIGHT_ORIGIN, HAS_ORIGIN | HAS_RIGHT_ORIGIN] {
                encoder.write_info(BLOCK_ITEM_ANY_REF_NUMBER | HAS_PARENT_SUB | flags);
                if flags & HAS_ORIGIN != 0 {
                    encoder.write_var(7_u64);
                    encoder.write_var(1_u32);
                }
                if flags & HAS_RIGHT_ORIGIN != 0 {
                    encoder.write_var(7_u64);
                    encoder.write_var(1_u32);
                }
                encoder.write_len(1);
                encoder.write_any(&Any::Bool(true));
            }
        });
        assert_boundaries(&update, 5);
    }

    #[test]
    fn every_lib0_any_tag_is_scanned() {
        let values = vec![
            Any::Undefined,
            Any::Null,
            Any::Number(-12345.0),
            Any::Number(1.5),
            Any::Number(1.0 / 3.0),
            Any::BigInt(i64::MIN),
            Any::Bool(false),
            Any::Bool(true),
            Any::from("🦀"),
            Any::Buffer(std::sync::Arc::from([0_u8, 128, 255])),
            Any::Array(std::sync::Arc::from([Any::Null, Any::from("nested")])),
            Any::Map(std::sync::Arc::new(HashMap::from([(
                "key".into(),
                Any::Array(std::sync::Arc::from([Any::Bool(true)])),
            )]))),
        ];
        let update = raw_update(1, |encoder| {
            item(encoder, BLOCK_ITEM_ANY_REF_NUMBER);
            encoder.write_len(values.len() as u32);
            for value in &values {
                encoder.write_any(value);
            }
        });
        for limit in LIMITS {
            let parts = split_update_v1(&update, limit).unwrap();
            assert_eq!(parts, vec![update.clone()]);
            Update::decode_v1(&parts[0]).unwrap();
        }
    }

    #[test]
    fn unsupported_or_inconsistent_content_is_refused() {
        let json = raw_update(1, |encoder| {
            item(encoder, BLOCK_ITEM_JSON_REF_NUMBER);
            encoder.write_len(1);
            encoder.write_string("null");
        });
        assert_eq!(
            split_update_v1(&json, 64),
            Err(SplitError::JsonLengthMismatch)
        );
        for kind in [11, 12, 13, 14, 15] {
            let update = raw_update(1, |encoder| item(encoder, kind));
            assert_eq!(
                split_update_v1(&update, 64),
                Err(SplitError::UnsupportedContent(kind))
            );
        }
        for kind in [7, 8, 10, 16, 128] {
            let update = raw_update(1, |encoder| {
                item(encoder, BLOCK_ITEM_TYPE_REF_NUMBER);
                encoder.write_type_ref(kind);
            });
            assert_eq!(
                split_update_v1(&update, 64),
                Err(SplitError::UnsupportedType(kind))
            );
        }
        let skip = raw_update(1, |encoder| {
            encoder.write_info(BLOCK_SKIP_REF_NUMBER);
            encoder.write_var(3_u32);
        });
        assert_eq!(
            split_update_v1(&skip, 64),
            Err(SplitError::MissingDependency)
        );
    }

    #[test]
    fn malformed_updates_are_refused() {
        let update = Workbook::from_model(small_model())
            .unwrap()
            .encode_state_as_update_v1();
        for end in 0..update.len() {
            assert!(
                split_update_v1(&update[..end], 64).is_err(),
                "accepted prefix {end}"
            );
        }
        for suffix in [vec![0], vec![255], vec![0, 0]] {
            let mut trailing = update.clone();
            trailing.extend(suffix);
            assert!(split_update_v1(&trailing, 64).is_err());
        }
        for bytes in [
            vec![128, 0, 0],
            vec![1, 0, 7, 0, 0],
            vec![1, 255, 255, 255, 255, 15],
            vec![0, 1, 7, 1, 0, 0],
            vec![255; 20],
        ] {
            assert!(split_update_v1(&bytes, 64).is_err());
        }
        assert_eq!(split_update_v1(&[0, 0], 1).unwrap(), vec![vec![0, 0]]);
        assert_eq!(split_update_v1(&update, 0), Err(SplitError::InvalidLimit));
    }

    #[test]
    fn malformed_values_are_refused() {
        for json in [
            "",
            "[1,]",
            "{\"x\":}",
            "01",
            "1e",
            "1e999",
            "true false",
            "\"\\uD800\"",
            "\"\\uDC00\"",
            "\"\\u+123\"",
            "\"\\q\"",
        ] {
            let update = raw_update(1, |encoder| {
                item(encoder, BLOCK_ITEM_EMBED_REF_NUMBER);
                encoder.write_string(json);
            });
            assert!(split_update_v1(&update, 64).is_err(), "accepted {json}");
        }
        for value in [
            vec![0],
            vec![125, 64],
            vec![125, 128, 0],
            vec![119, 1, 255],
            vec![117, 2, 126],
            vec![118, 1, 0],
            vec![116, 2, 0],
            vec![125, 255, 255, 255, 255, 255, 255, 255, 255, 255, 127],
        ] {
            let update = raw_update(1, |encoder| {
                item(encoder, BLOCK_ITEM_ANY_REF_NUMBER);
                encoder.write_len(1);
                encoder.write_all(&value);
            });
            assert!(split_update_v1(&update, 64).is_err());
        }
        let deeply_nested = raw_update(1, |encoder| {
            item(encoder, BLOCK_ITEM_ANY_REF_NUMBER);
            encoder.write_len(1);
            for _ in 0..64 {
                encoder.write_u8(117);
                encoder.write_var(1_u32);
            }
            encoder.write_u8(126);
        });
        assert!(split_update_v1(&deeply_nested, 64).is_err());
    }
}
