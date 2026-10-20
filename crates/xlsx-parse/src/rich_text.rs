use std::collections::BTreeMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use xlsx_model::addr::MAX_ROWS;
use xlsx_model::{CellRef, CellValue, Color, RichRun, RunBaseline, RunStyle, RunUnderline};

use crate::read::parse_index;
use crate::xml::{attr, local_name, next_event, reader, resolve_entity, xml_err};
use crate::{ParseError, PreservedPackage};

/// Maximum runs indexed across one source package.
pub const MAX_RICH_TEXT_RUNS: usize = 100_000;
/// Maximum runs indexed in one string item.
pub const MAX_RICH_TEXT_RUNS_PER_ITEM: usize = 4_096;
/// Maximum font-name bytes retained across one source package.
pub const MAX_RICH_TEXT_STYLE_BYTES: usize = 8 * 1024 * 1024;

#[derive(Debug, Default)]
pub(crate) struct RichTextIndex {
    shared: BTreeMap<usize, Vec<RichRun>>,
    inline: BTreeMap<(usize, u32, u32), Vec<RichRun>>,
}

impl PreservedPackage {
    /// Source runs over the exact shared string, or `None` for plain/unsupported items.
    pub fn shared_string_text_runs(&self, index: usize) -> Option<&[RichRun]> {
        if !self.shared_string_is_rich(index) {
            return None;
        }
        self.text_run_index().shared.get(&index).map(Vec::as_slice)
    }

    /// Source runs over the exact inline string; sheet and cell indices are zero-based.
    pub fn inline_string_text_runs(&self, sheet: usize, cell: CellRef) -> Option<&[RichRun]> {
        if !self
            .source_cell_facts(sheet)?
            .rich_inline
            .contains_key(&(cell.row, cell.col))
        {
            return None;
        }
        self.text_run_index()
            .inline
            .get(&(sheet, cell.row, cell.col))
            .map(Vec::as_slice)
    }

    fn text_run_index(&self) -> &RichTextIndex {
        self.text_runs.get_or_init(|| RichTextIndex::build(self))
    }
}

struct Budget {
    runs: usize,
    style_bytes: usize,
    exhausted: bool,
}

impl Default for Budget {
    fn default() -> Self {
        Self {
            runs: MAX_RICH_TEXT_RUNS,
            style_bytes: MAX_RICH_TEXT_STYLE_BYTES,
            exhausted: false,
        }
    }
}

impl Budget {
    fn run(&mut self) -> bool {
        if self.exhausted || self.runs == 0 {
            self.exhausted = true;
            return false;
        }
        self.runs -= 1;
        true
    }

    fn font_name(&mut self, bytes: usize) -> bool {
        if let Some(left) = self.style_bytes.checked_sub(bytes) {
            self.style_bytes = left;
            true
        } else {
            self.exhausted = true;
            false
        }
    }
}

impl RichTextIndex {
    fn build(package: &PreservedPackage) -> Self {
        let mut index = Self::default();
        let mut budget = Budget::default();
        if let Some(bytes) = package
            .shared_strings
            .as_ref()
            .and_then(|part| package.part_bytes(&part.path))
        {
            index.read_shared(
                bytes,
                &package.original_workbook.shared_strings,
                &mut budget,
            );
        }
        for (sheet, source) in package.sheets.iter().enumerate() {
            if budget.exhausted {
                break;
            }
            if source.cell_facts.rich_inline.is_empty() {
                continue;
            }
            let Some(bytes) = package.part_bytes(&source.path) else {
                continue;
            };
            if index
                .read_inline(bytes, package, sheet, &mut budget)
                .is_err()
            {
                index.inline.retain(|&(source, _, _), _| source != sheet);
            }
        }
        index
    }

    fn read_shared(&mut self, bytes: &[u8], strings: &[String], budget: &mut Budget) {
        let mut xml = reader(bytes);
        let mut buffer = Vec::new();
        let mut depth = 0;
        let mut item = 0;
        let result = (|| -> Result<(), ParseError> {
            loop {
                if budget.exhausted {
                    break;
                }
                match next_event(&mut xml, &mut buffer, &mut depth)? {
                    Event::Start(start) if local_name(&start) == b"si" => {
                        let Some(text) = strings.get(item) else {
                            break;
                        };
                        if let Some(runs) =
                            read_item(&mut xml, &mut buffer, &mut depth, text, budget)?
                        {
                            self.shared.insert(item, runs);
                        }
                        item += 1;
                    }
                    Event::Eof => break,
                    _ => {}
                }
            }
            Ok(())
        })();
        if result.is_err() {
            self.shared.clear();
        }
    }

    fn read_inline(
        &mut self,
        bytes: &[u8],
        package: &PreservedPackage,
        sheet: usize,
        budget: &mut Budget,
    ) -> Result<(), ParseError> {
        let mut xml = reader(bytes);
        let mut buffer = Vec::new();
        let mut depth = 0;
        let mut cell = None;
        let mut row: Option<u32> = None;
        let mut col = 0;
        loop {
            if budget.exhausted {
                break;
            }
            match next_event(&mut xml, &mut buffer, &mut depth)? {
                Event::Start(start) if local_name(&start) == b"row" => {
                    row = Some(match attr(&start, b"r")? {
                        Some(value) => parse_index(&value, MAX_ROWS)?,
                        None => row.map_or(0, |row| row + 1),
                    });
                    col = 0;
                }
                Event::Start(start) if local_name(&start) == b"c" => {
                    cell = match attr(&start, b"r")? {
                        Some(value) => CellRef::parse_a1(&value).ok(),
                        None => Some(CellRef::new(row.unwrap_or(0), col)),
                    };
                    if let Some(at) = cell {
                        col = at.col;
                        self.inline.remove(&(sheet, at.row, at.col));
                    }
                    if attr(&start, b"t")?.as_deref() != Some("inlineStr") {
                        cell = None;
                    }
                }
                Event::Start(start) if local_name(&start) == b"is" => {
                    let Some(at) = cell else {
                        continue;
                    };
                    let Some(CellValue::Text { value }) = package.original_workbook.sheets[sheet]
                        .cell(at)
                        .map(|cell| &cell.value)
                    else {
                        continue;
                    };
                    if let Some(runs) = read_item(&mut xml, &mut buffer, &mut depth, value, budget)?
                    {
                        self.inline.insert((sheet, at.row, at.col), runs);
                    } else {
                        self.inline.remove(&(sheet, at.row, at.col));
                    }
                }
                Event::End(end) if end.local_name().as_ref() == b"c" => {
                    cell = None;
                    col += 1;
                }
                Event::End(end) if end.local_name().as_ref() == b"row" => row = None,
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(())
    }
}

fn read_item(
    xml: &mut Reader<&[u8]>,
    buffer: &mut Vec<u8>,
    depth: &mut usize,
    expected: &str,
    budget: &mut Budget,
) -> Result<Option<Vec<RichRun>>, ParseError> {
    let target = *depth;
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut runs = Vec::new();
    let mut style = RunStyle::default();
    let mut offset = 0_usize;
    let mut start = 0;
    let mut valid = true;
    let mut item_runs = 0;
    let mut saw_text = false;
    let mut saw_properties = false;
    loop {
        match next_event(xml, buffer, depth)? {
            Event::Start(element) => {
                let name = local_name(&element);
                match (stack.last().map(Vec::as_slice), name.as_slice()) {
                    (None, b"r") => {
                        item_runs += 1;
                        if valid && (item_runs > MAX_RICH_TEXT_RUNS_PER_ITEM || !budget.run()) {
                            valid = false;
                        }
                        start = offset;
                        style = RunStyle::default();
                        saw_text = false;
                        saw_properties = false;
                    }
                    (Some(b"r"), b"rPr") if !saw_properties && !saw_text => saw_properties = true,
                    (Some(b"r"), b"t") if !saw_text => saw_text = true,
                    (Some(b"rPr"), _) if valid => {
                        if !read_property(&element, &mut style, budget)? {
                            valid = false;
                        }
                    }
                    _ => valid = false,
                }
                stack.push(name);
            }
            Event::Text(text) => {
                let text = text.decode().map_err(xml_err)?;
                read_text(&text, expected, &stack, &mut offset, &mut valid);
            }
            Event::CData(text) => {
                let text = text.decode().map_err(xml_err)?;
                read_text(&text, expected, &stack, &mut offset, &mut valid);
            }
            Event::GeneralRef(entity) => {
                let text = resolve_entity(&entity.decode().map_err(xml_err)?)?;
                read_text(&text, expected, &stack, &mut offset, &mut valid);
            }
            Event::End(_) if *depth < target => break,
            Event::End(_) => {
                if stack.pop().as_deref() == Some(b"r") && valid && saw_text {
                    match (u32::try_from(start), u32::try_from(offset)) {
                        (Ok(start), Ok(end)) => runs.push(RichRun::new(start, end, style.clone())),
                        _ => valid = false,
                    }
                } else if stack.is_empty() && !saw_text {
                    valid = false;
                }
            }
            Event::Eof => {
                return Err(ParseError::Malformed(
                    "unexpected eof in string item".into(),
                ));
            }
            Event::Comment(_) => {}
            _ => valid = false,
        }
    }
    Ok(
        (valid && !budget.exhausted && offset == expected.len() && !runs.is_empty())
            .then_some(runs),
    )
}

fn read_text(text: &str, expected: &str, stack: &[Vec<u8>], offset: &mut usize, valid: &mut bool) {
    if !matches!(stack.last().map(Vec::as_slice), Some(b"t")) {
        *valid = false;
    }
    let Some(end) = offset.checked_add(text.len()) else {
        *valid = false;
        return;
    };
    if expected.get(*offset..end) != Some(text) {
        *valid = false;
    }
    *offset = end;
}

fn read_property(
    element: &BytesStart<'_>,
    style: &mut RunStyle,
    budget: &mut Budget,
) -> Result<bool, ParseError> {
    let value = attr(element, b"val")?;
    let name = local_name(element);
    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_err)?;
        let attribute = attribute.key.local_name();
        if name == b"color" {
            if !matches!(
                attribute.as_ref(),
                b"rgb" | b"theme" | b"indexed" | b"auto" | b"tint"
            ) {
                return Ok(false);
            }
        } else if attribute.as_ref() != b"val" {
            return Ok(false);
        }
    }
    let valid = match name.as_slice() {
        b"rFont" if style.name.is_none() => {
            match value.filter(|value| budget.font_name(value.len())) {
                Some(value) => {
                    style.name = Some(value);
                    true
                }
                None => false,
            }
        }
        b"sz" if style.size_pt.is_none() => {
            style.size_pt = value
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| value.is_finite() && *value > 0.0);
            style.size_pt.is_some()
        }
        b"b" | b"i" | b"strike" => {
            let slot = match name.as_slice() {
                b"b" => &mut style.bold,
                b"i" => &mut style.italic,
                _ => &mut style.strike,
            };
            if slot.is_some() {
                false
            } else {
                *slot = match value.as_deref() {
                    None | Some("1" | "true") => Some(true),
                    Some("0" | "false") => Some(false),
                    _ => None,
                };
                slot.is_some()
            }
        }
        b"u" if style.underline.is_none() => {
            style.underline = match value.as_deref() {
                None | Some("single") => Some(RunUnderline::Single),
                Some("none") => Some(RunUnderline::None),
                Some("double") => Some(RunUnderline::Double),
                Some("singleAccounting") => Some(RunUnderline::SingleAccounting),
                Some("doubleAccounting") => Some(RunUnderline::DoubleAccounting),
                _ => None,
            };
            style.underline.is_some()
        }
        b"vertAlign" if style.baseline.is_none() => {
            style.baseline = match value.as_deref() {
                Some("baseline") => Some(RunBaseline::Baseline),
                Some("superscript") => Some(RunBaseline::Superscript),
                Some("subscript") => Some(RunBaseline::Subscript),
                _ => None,
            };
            style.baseline.is_some()
        }
        b"color" if style.color.is_none() => {
            let tint = attr(element, b"tint")?;
            let mut sources = 0;
            for name in [b"rgb".as_slice(), b"theme", b"indexed", b"auto"] {
                sources += usize::from(attr(element, name)?.is_some());
            }
            if sources != 1
                || tint.as_deref().is_some_and(|value| {
                    value
                        .parse::<f64>()
                        .ok()
                        .is_none_or(|value| !value.is_finite() || !(-1.0..=1.0).contains(&value))
                })
            {
                false
            } else {
                style.color = if attr(element, b"auto")?.as_deref() == Some("true") {
                    Some(Color::Auto)
                } else {
                    crate::styles::parse_color(element)?
                };
                style.color.is_some()
                    && (tint.is_none() || matches!(style.color, Some(Color::Theme { .. })))
            }
        }
        _ => false,
    };
    Ok(valid)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn mismatching_text_discards_all_runs_instead_of_clamping() {
        let mut index = RichTextIndex::default();
        index.read_shared(
            b"<sst><si><r><t>first</t></r><r><t>last</t></r></si></sst>",
            &["first".into()],
            &mut Budget::default(),
        );
        assert!(index.shared.is_empty());
    }

    #[test]
    fn style_byte_budget_exhaustion_discards_the_item_and_later_items() {
        let mut index = RichTextIndex::default();
        let mut budget = Budget {
            style_bytes: 3,
            ..Budget::default()
        };
        index.read_shared(
            br#"<sst><si><r><rPr><rFont val="Long"/></rPr><t>a</t></r></si><si><r><t>b</t></r></si></sst>"#,
            &["a".into(), "b".into()],
            &mut budget,
        );
        assert!(index.shared.is_empty());
        assert!(budget.exhausted);
    }

    #[test]
    fn total_budget_discards_overflowing_item_and_later_items() {
        let mut index = RichTextIndex::default();
        let mut budget = Budget {
            runs: 2,
            ..Budget::default()
        };
        index.read_shared(b"<sst><si><r><t>a</t></r></si><si><r><t>b</t></r><r><t>c</t></r></si><si><r><t>d</t></r></si></sst>", &["a".into(), "bc".into(), "d".into()], &mut budget);
        assert_eq!(index.shared.len(), 1);
        assert!(index.shared.contains_key(&0));
        assert!(budget.exhausted);
    }

    #[test]
    fn per_item_overflow_does_not_exhaust_the_global_budget() {
        let item = "<r><t>x</t></r>".repeat(MAX_RICH_TEXT_RUNS_PER_ITEM + 1);
        let mut index = RichTextIndex::default();
        let mut budget = Budget {
            runs: MAX_RICH_TEXT_RUNS_PER_ITEM + 1,
            ..Budget::default()
        };
        index.read_shared(
            format!("<sst><si>{item}</si><si><r><t>later</t></r></si></sst>").as_bytes(),
            &["x".repeat(MAX_RICH_TEXT_RUNS_PER_ITEM + 1), "later".into()],
            &mut budget,
        );
        assert!(!index.shared.contains_key(&0));
        assert_eq!(index.shared[&1].len(), 1);
        assert_eq!(budget.runs, 0);
        assert!(!budget.exhausted);
    }

    #[test]
    fn invalid_items_do_not_charge_subsequent_runs_or_styles() {
        for (prefix, charged) in [
            ("<unsupported/>", 0),
            ("<r><rPr><unsupported/></rPr><t>x</t></r>", 1),
            ("<r><t>mismatch</t></r>", 1),
        ] {
            let mut index = RichTextIndex::default();
            let mut budget = Budget {
                runs: charged + 1,
                style_bytes: 0,
                exhausted: false,
            };
            index.read_shared(
                format!(
                    r#"<sst><si>{prefix}<r><rPr><rFont val="Long"/></rPr><t>x</t></r><r><t>x</t></r></si><si><r><t>later</t></r></si></sst>"#
                )
                .as_bytes(),
                &["xxx".into(), "later".into()],
                &mut budget,
            );
            assert!(!index.shared.contains_key(&0));
            assert_eq!(index.shared[&1].len(), 1);
            assert_eq!(budget.runs, 0);
            assert!(!budget.exhausted);
        }
    }
}
