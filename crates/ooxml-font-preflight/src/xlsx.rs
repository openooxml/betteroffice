use std::collections::{BTreeMap, BTreeSet};

use crate::model::Faces;
use crate::xml::Node;
use crate::{FontPreflight, Parts, PreflightError, xml, zip::Zip};

pub const OPENING_VIEWPORT_ROWS: u32 = 50;
pub const OPENING_VIEWPORT_COLUMNS: u32 = 20;
const MAX_STYLE_ENTRIES: usize = 4096;

#[derive(Clone, Copy)]
enum AlternateMode {
    Styles,
    Worksheet,
    Unsupported,
}

fn resolve_alternate_content(node: &mut Node, mode: AlternateMode, complete: &mut bool) {
    let mut children = Vec::new();
    let mut wrapped = false;
    for mut child in std::mem::take(&mut node.children) {
        if child.name != "AlternateContent" {
            resolve_alternate_content(&mut child, mode, complete);
            children.push(child);
            continue;
        }
        wrapped = true;
        if matches!(mode, AlternateMode::Unsupported)
            || (matches!(mode, AlternateMode::Worksheet)
                && !matches!(node.name, "s:sheetData" | "s:row" | "s:c"))
        {
            *complete = false;
        }
        let mut taken = false;
        for mut branch in child.children {
            if !matches!(branch.name, "Choice" | "Fallback") {
                *complete = false;
                resolve_alternate_content(&mut branch, mode, complete);
                children.push(branch);
                continue;
            }
            if taken && matches!(mode, AlternateMode::Styles) {
                continue;
            }
            taken = true;
            branch.name = node.name;
            resolve_alternate_content(&mut branch, mode, complete);
            children.extend(branch.children);
        }
        if !taken {
            *complete = false;
        }
    }
    if wrapped && matches!(mode, AlternateMode::Worksheet) && node.name == "s:c" {
        let mut seen = BTreeSet::new();
        children.reverse();
        children.retain(|child| {
            !matches!(child.name, "s:v" | "s:f" | "s:is") || seen.insert(child.name)
        });
        children.reverse();
    }
    node.children = children;
}

fn flag(value: Option<&str>) -> bool {
    value.is_some_and(|value| matches!(value, "1" | "true" | "on"))
}

fn number(node: &Node, key: &str, complete: &mut bool) -> Option<usize> {
    let value = node.attr(key)?;
    let parsed = value.parse().ok();
    if parsed.is_none() {
        *complete = false;
    }
    parsed
}

fn address(value: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row = 0u32;
    let mut digits = false;
    for byte in value.bytes() {
        if byte.is_ascii_uppercase() && !digits {
            col = col
                .checked_mul(26)?
                .checked_add(u32::from(byte - b'A') + 1)?;
        } else if byte.is_ascii_digit() {
            digits = true;
            row = row.checked_mul(10)?.checked_add(u32::from(byte - b'0'))?;
        } else {
            return None;
        }
    }
    if !digits || col > 16384 || row > 1048576 {
        return None;
    }
    Some((row.checked_sub(1)?, col.checked_sub(1)?))
}

fn text_pieces<'a>(node: &'a Node, out: &mut Vec<&'a str>) {
    if node.name == "s:t" {
        out.push(&node.text);
    }
    for child in &node.children {
        text_pieces(child, out);
    }
}

fn has_inherited_styles(node: &Node) -> bool {
    let style = match node.name {
        "s:row" if flag(node.attr("customFormat")) => node.attr("s"),
        "s:col" => node.attr("style"),
        _ => None,
    };
    style
        .and_then(|value| value.parse::<u32>().ok())
        .is_some_and(|style| style != 0)
        || node.children.iter().any(has_inherited_styles)
}

fn cell_is_stored(cell: &Node, inherited_styles: bool) -> Option<bool> {
    let value = cell
        .children
        .iter()
        .rev()
        .find(|node| node.name == "s:v")
        .map(|node| node.text.as_str());
    let has_value = match cell.attr("t") {
        Some("s") => {
            if value
                .and_then(|value| value.trim().parse::<usize>().ok())
                .is_some()
            {
                None
            } else {
                Some(false)
            }
        }
        Some("inlineStr") => Some(cell.child("s:is").is_some()),
        Some("str" | "b" | "e") => Some(true),
        _ => match value.filter(|value| !value.trim().is_empty()) {
            Some(value) => {
                value.trim().parse::<f64>().ok()?;
                Some(true)
            }
            None => Some(false),
        },
    };
    if has_value == Some(true)
        || cell
            .attr("s")
            .and_then(|style| style.parse::<u32>().ok())
            .is_some()
        || cell.child("s:f").is_some()
    {
        Some(true)
    } else if inherited_styles {
        None
    } else {
        has_value
    }
}

fn unescape(source: &str) -> String {
    let bytes = source.as_bytes();
    let mut output = String::with_capacity(source.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes.get(index) == Some(&b'_')
            && matches!(bytes.get(index + 1), Some(b'x' | b'X'))
            && bytes.get(index + 6) == Some(&b'_')
            && bytes[index + 2..index + 6]
                .iter()
                .all(u8::is_ascii_hexdigit)
            && let Ok(code) = u32::from_str_radix(&source[index + 2..index + 6], 16)
            && let Some(ch) = char::from_u32(code)
        {
            output.push(ch);
            index += 7;
        } else if let Some(ch) = source[index..].chars().next() {
            output.push(ch);
            index += ch.len_utf8();
        }
    }
    output
}

fn scripts(pieces: &[&str], budget: &mut usize, complete: &mut bool) -> Vec<String> {
    let bytes = pieces.iter().map(|text| text.len()).sum::<usize>();
    let Some(remaining) = budget.checked_sub(bytes) else {
        *complete = false;
        return Vec::new();
    };
    *budget = remaining;
    let texts: Vec<_> = pieces.iter().map(|text| unescape(text)).collect();
    ooxml_text::paragraph_scripts_with_languages(
        texts.iter().map(|text| (text.as_str(), None)),
        false,
    )
}

#[derive(Clone, Default)]
struct Font {
    family: Option<String>,
    bold: bool,
    italic: bool,
    scheme: bool,
}

struct Styles {
    fonts: Vec<Font>,
    xfs: Vec<Option<usize>>,
    custom_formats: BTreeMap<usize, String>,
    format_ids: Vec<usize>,
}

impl Styles {
    fn parse(root: Option<&Node>, complete: &mut bool) -> Self {
        let mut result = Self {
            fonts: Vec::new(),
            xfs: Vec::new(),
            custom_formats: BTreeMap::new(),
            format_ids: Vec::new(),
        };
        let Some(root) = root else { return result };
        for (index, node) in root
            .children
            .iter()
            .filter(|node| node.name == "s:fonts")
            .flat_map(|node| &node.children)
            .filter(|node| node.name == "s:font")
            .enumerate()
        {
            if index == MAX_STYLE_ENTRIES {
                *complete = false;
                break;
            }
            let property = |name| node.children.iter().rev().find(|child| child.name == name);
            result.fonts.push(Font {
                family: property("s:name")
                    .and_then(|node| node.attr("val"))
                    .map(unescape),
                bold: property("s:b")
                    .is_some_and(|node| node.attr("val").is_none_or(|value| flag(Some(value)))),
                italic: property("s:i")
                    .is_some_and(|node| node.attr("val").is_none_or(|value| flag(Some(value)))),
                scheme: property("s:scheme").is_some(),
            });
        }
        for (index, node) in root
            .children
            .iter()
            .filter(|node| node.name == "s:numFmts")
            .flat_map(|node| &node.children)
            .filter(|node| node.name == "s:numFmt")
            .enumerate()
        {
            if index == MAX_STYLE_ENTRIES {
                *complete = false;
                break;
            }
            if let Some(id) = number(node, "numFmtId", complete) {
                result
                    .custom_formats
                    .entry(id)
                    .or_insert_with(|| node.attr("formatCode").unwrap_or_default().to_owned());
            }
        }
        for (index, node) in root
            .children
            .iter()
            .filter(|node| node.name == "s:cellXfs")
            .flat_map(|node| &node.children)
            .filter(|node| node.name == "s:xf")
            .enumerate()
        {
            if index == MAX_STYLE_ENTRIES {
                *complete = false;
                break;
            }
            result.xfs.push(
                if node.attr("applyFont").is_none_or(|value| flag(Some(value))) {
                    number(node, "fontId", complete)
                } else {
                    None
                },
            );
            result.format_ids.push(
                if node
                    .attr("applyNumberFormat")
                    .is_none_or(|value| flag(Some(value)))
                {
                    number(node, "numFmtId", complete).unwrap_or(0)
                } else {
                    0
                },
            );
        }
        result
    }

    fn font(&self, style: usize, complete: &mut bool) -> Font {
        let selected = self.xfs.get(style);
        if selected.is_none() && style != 0 {
            *complete = false;
        }
        let font = selected.and_then(|index| *index).and_then(|index| {
            let font = self.fonts.get(index);
            if font.is_none() {
                *complete = false;
            }
            font
        });
        let normal = self
            .xfs
            .first()
            .and_then(|index| *index)
            .and_then(|index| self.fonts.get(index))
            .or_else(|| self.fonts.first());
        let family = font
            .and_then(|font| font.family.clone())
            .or_else(|| {
                normal
                    .and_then(|font| font.family.clone())
                    .filter(|name| !name.trim().is_empty())
            })
            .unwrap_or_else(|| "Calibri".into());
        if font.is_some_and(|font| font.scheme && font.family.is_none()) {
            *complete = false;
        }
        Font {
            family: Some(family),
            bold: font.is_some_and(|font| font.bold),
            italic: font.is_some_and(|font| font.italic),
            scheme: false,
        }
    }
}

pub fn xlsx_font_preflight(bytes: &[u8]) -> Result<FontPreflight, PreflightError> {
    let mut parts = Parts {
        zip: Zip::new(bytes)?,
        nodes: xml::MAX_NODES,
    };
    let mut complete = true;
    let mut workbook = parts
        .part("xl/workbook.xml", &mut complete)?
        .ok_or(PreflightError::MissingPart)?;
    resolve_alternate_content(&mut workbook, AlternateMode::Unsupported, &mut complete);
    if workbook.name != "s:workbook" {
        return Err(PreflightError::Xml);
    }
    let rels = parts.rels("xl/workbook.xml", &mut complete)?;
    let sheets: Vec<_> = workbook
        .child("s:sheets")
        .into_iter()
        .flat_map(|node| &node.children)
        .filter(|node| node.name == "s:sheet")
        .collect();
    let active = workbook
        .path(&["s:bookViews", "s:workbookView"])
        .and_then(|node| number(node, "activeTab", &mut complete))
        .unwrap_or(0);
    let visible = |node: &&Node| node.attr("state").is_none_or(|state| state == "visible");
    let sheet = sheets
        .get(active)
        .filter(|node| visible(node))
        .copied()
        .or_else(|| {
            complete = false;
            sheets.iter().find(|node| visible(node)).copied()
        })
        .ok_or(PreflightError::MissingPart)?;
    let id = sheet.attr("relId");
    let fallback_path = format!(
        "xl/worksheets/sheet{}.xml",
        sheets
            .iter()
            .position(|node| std::ptr::eq(*node, sheet))
            .unwrap_or(active)
            + 1
    );
    let path = match rels.iter().find(|rel| Some(rel.id.as_str()) == id) {
        Some(rel) if rel.kind == "worksheet" => rel.target.as_str(),
        Some(_) => return Ok(Faces::default().finish(false)),
        None => &fallback_path,
    };
    let mut sheet = parts
        .part(path, &mut complete)?
        .ok_or(PreflightError::MissingPart)?;
    resolve_alternate_content(&mut sheet, AlternateMode::Worksheet, &mut complete);
    if sheet.name != "s:worksheet" {
        return Err(PreflightError::Xml);
    }
    let styles_path = rels
        .iter()
        .find(|rel| rel.kind == "styles")
        .map(|rel| rel.target.as_str())
        .unwrap_or("xl/styles.xml");
    let mut styles = parts.part(styles_path, &mut complete)?;
    if let Some(styles) = &mut styles {
        resolve_alternate_content(styles, AlternateMode::Styles, &mut complete);
    }
    if styles.is_none() && rels.iter().any(|rel| rel.kind == "styles") {
        complete = false;
    }
    let styles = styles.as_ref();
    if styles.is_some_and(|node| node.name != "s:styleSheet") {
        return Err(PreflightError::Xml);
    }
    let styles = Styles::parse(styles, &mut complete);
    if let Some(view) = sheet.path(&["s:sheetViews", "s:sheetView"])
        && (view.attr("topLeftCell").is_some_and(|cell| cell != "A1")
            || view.child("s:pane").is_some())
    {
        complete = false;
    }
    if ["s:conditionalFormatting", "s:hyperlinks", "s:drawing"]
        .into_iter()
        .any(|name| sheet.child(name).is_some())
    {
        complete = false;
    }
    let mut hidden_columns = [false; OPENING_VIEWPORT_COLUMNS as usize];
    if let Some(cols) = sheet.child("s:cols") {
        for col in &cols.children {
            let start = number(col, "min", &mut complete)
                .unwrap_or(1)
                .saturating_sub(1);
            let end = number(col, "max", &mut complete)
                .unwrap_or(0)
                .min(hidden_columns.len());
            let start = start.min(hidden_columns.len());
            for hidden in hidden_columns.iter_mut().take(end).skip(start) {
                *hidden = flag(col.attr("hidden"))
                    || col
                        .attr("width")
                        .and_then(|value| value.parse::<f64>().ok())
                        == Some(0.0);
            }
        }
    }
    let mut merged_anchors =
        [[false; OPENING_VIEWPORT_COLUMNS as usize]; OPENING_VIEWPORT_ROWS as usize];
    let mut covered = merged_anchors;
    let mut merge_count = 0;
    if let Some(root) = sheet.child("s:mergeCells") {
        for node in &root.children {
            if let Some((start, end)) = node
                .attr("ref")
                .and_then(|value| value.split_once(':'))
                .and_then(|(start, end)| Some((address(start)?, address(end)?)))
            {
                if start.0 < OPENING_VIEWPORT_ROWS && start.1 < OPENING_VIEWPORT_COLUMNS {
                    if merge_count == MAX_STYLE_ENTRIES {
                        complete = false;
                        break;
                    }
                    merge_count += 1;
                    merged_anchors[start.0 as usize][start.1 as usize] |= end != start;
                    for (r, row) in covered
                        .iter_mut()
                        .enumerate()
                        .take((end.0 + 1).min(OPENING_VIEWPORT_ROWS) as usize)
                        .skip(start.0 as usize)
                    {
                        for (c, cell) in row
                            .iter_mut()
                            .enumerate()
                            .take((end.1 + 1).min(OPENING_VIEWPORT_COLUMNS) as usize)
                            .skip(start.1 as usize)
                        {
                            if (r as u32, c as u32) != start {
                                *cell = true;
                            }
                        }
                    }
                }
            } else {
                complete = false;
            }
        }
    }
    let inherited_styles = has_inherited_styles(&sheet);
    let mut candidates = BTreeMap::new();
    let mut hidden_rows = [false; OPENING_VIEWPORT_ROWS as usize];
    if let Some(data) = sheet.child("s:sheetData") {
        for row in &data.children {
            if row.name != "s:row" {
                continue;
            }
            let height = row.attr("ht").and_then(|value| value.parse::<f64>().ok());
            let index = number(row, "r", &mut complete).unwrap_or(1).checked_sub(1);
            if let Some(index) = index {
                if let Some(hidden) = hidden_rows.get_mut(index) {
                    if flag(row.attr("hidden")) || height == Some(0.0) {
                        *hidden = true;
                    } else if height.is_some() {
                        *hidden = false;
                    }
                }
            } else {
                complete = false;
            }
            for cell in &row.children {
                if cell.name != "s:c" {
                    continue;
                }
                let Some((r, c)) = cell.attr("r").and_then(address) else {
                    complete = false;
                    continue;
                };
                if r >= OPENING_VIEWPORT_ROWS || c >= OPENING_VIEWPORT_COLUMNS {
                    continue;
                }
                if candidates.contains_key(&(r, c)) {
                    match cell_is_stored(cell, inherited_styles) {
                        Some(true) => {}
                        Some(false) => continue,
                        None => {
                            complete = false;
                            continue;
                        }
                    }
                } else if cell.attr("s").is_none()
                    && cell.child("s:f").is_none()
                    && cell.attr("t").is_none_or(|kind| kind == "n")
                    && cell
                        .child("s:v")
                        .is_none_or(|node| node.text.trim().is_empty())
                {
                    continue;
                }
                candidates.insert((r, c), cell);
            }
        }
    }
    let mut cells = BTreeMap::new();
    let mut string_ids = BTreeSet::new();
    for ((r, c), cell) in candidates {
        if hidden_rows[r as usize] || hidden_columns[c as usize] {
            if merged_anchors[r as usize][c as usize] {
                complete = false;
            }
            continue;
        }
        if covered[r as usize][c as usize] {
            continue;
        }
        if cell.child("s:f").is_some() {
            complete = false;
        }
        let value = cell.child("s:v").map(|node| node.text.as_str());
        let kind = cell.attr("t").unwrap_or("n");
        let inline = cell.child("s:is");
        if !matches!(kind, "n" | "s" | "str" | "inlineStr" | "b" | "e") {
            complete = false;
            continue;
        }
        if kind == "n"
            && value.is_some_and(|value| {
                !value.trim().is_empty() && value.trim().parse::<f64>().is_err()
            })
        {
            complete = false;
            continue;
        }
        if kind == "inlineStr" {
            let mut pieces = Vec::new();
            if let Some(inline) = inline {
                text_pieces(inline, &mut pieces);
            }
            if pieces.iter().all(|piece| piece.is_empty()) {
                continue;
            }
        } else if value.is_none_or(str::is_empty)
            || (kind == "n" && value.is_some_and(|value| value.trim().is_empty()))
        {
            continue;
        }
        let style = number(cell, "s", &mut complete).unwrap_or(0);
        if let Some(code) = styles
            .format_ids
            .get(style)
            .filter(|id| **id >= 164)
            .and_then(|id| styles.custom_formats.get(id))
        {
            if code == ";;;" {
                continue;
            }
            complete = false;
        }
        if kind == "s" {
            if let Some(id) = value.and_then(|value| value.trim().parse::<usize>().ok()) {
                string_ids.insert(id);
            } else {
                complete = false;
            }
        }
        cells.insert((r, c), (cell, style));
    }
    let mut shared = if string_ids.is_empty() {
        None
    } else {
        let path = rels
            .iter()
            .find(|rel| rel.kind == "sharedStrings")
            .map(|rel| rel.target.as_str())
            .unwrap_or("xl/sharedStrings.xml");
        parts.part(path, &mut complete)?
    };
    if let Some(shared) = &mut shared {
        resolve_alternate_content(shared, AlternateMode::Unsupported, &mut complete);
    }
    if shared.as_ref().is_some_and(|node| node.name != "s:sst") {
        return Err(PreflightError::Xml);
    }
    let strings: BTreeMap<_, _> = shared
        .as_ref()
        .into_iter()
        .flat_map(|node| &node.children)
        .filter(|node| node.name == "s:si")
        .enumerate()
        .filter(|(index, _)| string_ids.contains(index))
        .collect();
    let mut script_budget = 4 * 1024 * 1024;
    let mut shared_scripts = BTreeMap::new();
    for (&index, text) in &strings {
        let mut pieces = Vec::new();
        text_pieces(text, &mut pieces);
        let nonempty = pieces.iter().any(|piece| !piece.is_empty());
        shared_scripts.insert(
            index,
            (
                nonempty,
                scripts(&pieces, &mut script_budget, &mut complete),
            ),
        );
    }
    let mut faces = Faces::default();
    for (cell, style) in cells.into_values() {
        let mut pieces = Vec::new();
        let scripts = match cell.attr("t").unwrap_or("n") {
            "s" => {
                let text = cell
                    .child("s:v")
                    .and_then(|node| node.text.trim().parse::<usize>().ok())
                    .and_then(|index| shared_scripts.get(&index));
                if let Some((nonempty, scripts)) = text {
                    if !*nonempty {
                        continue;
                    }
                    scripts.clone()
                } else {
                    complete = false;
                    continue;
                }
            }
            "inlineStr" => {
                if let Some(text) = cell.child("s:is") {
                    text_pieces(text, &mut pieces);
                }
                scripts(&pieces, &mut script_budget, &mut complete)
            }
            "str" | "e" => {
                if let Some(text) = cell.child("s:v") {
                    pieces.push(text.text.as_str());
                }
                scripts(&pieces, &mut script_budget, &mut complete)
            }
            _ => Vec::new(),
        };
        if matches!(cell.attr("t"), Some("inlineStr" | "str"))
            && pieces.iter().all(|piece| piece.is_empty())
        {
            continue;
        }
        let font = styles.font(style, &mut complete);
        faces.record(
            font.family.as_deref().unwrap_or("Calibri"),
            font.bold,
            font.italic,
            &scripts,
        );
    }
    Ok(faces.finish(complete))
}

#[cfg(test)]
mod tests;
