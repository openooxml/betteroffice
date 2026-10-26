use std::collections::HashSet;
use std::ops::Range;

use xlsx_model::Workbook;

use crate::ParseError;
use crate::chart::quoted_sheet;
use crate::tree::{Element, Part, Replacement, names_are_resolvable};

const NS_APP: &str = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const NS_VT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

pub(super) fn patch(
    bytes: &[u8],
    workbook: &Workbook,
    original: &Workbook,
    origins: &[Option<usize>],
) -> Result<Option<Vec<u8>>, ParseError> {
    if workbook.sheets.len() != original.sheets.len()
        || origins
            .iter()
            .flatten()
            .copied()
            .collect::<HashSet<_>>()
            .len()
            != original.sheets.len()
        || origins
            .iter()
            .any(|origin| origin.is_none_or(|index| index >= original.sheets.len()))
        || workbook
            .sheets
            .iter()
            .zip(&original.sheets)
            .all(|(sheet, old)| sheet.name == old.name)
    {
        return Ok(None);
    }
    let part = Part::decode(bytes)?;
    let root = part.tree()?;
    if !root.is(NS_APP, "Properties") || !names_are_resolvable(&root) {
        return Ok(None);
    }
    let Some(headings) =
        unique_child(&root, NS_APP, "HeadingPairs").and_then(|heading| vector(heading, "variant"))
    else {
        return Ok(None);
    };
    let Some(titles) =
        unique_child(&root, NS_APP, "TitlesOfParts").and_then(|titles| vector(titles, "lpstr"))
    else {
        return Ok(None);
    };
    let headings = headings.child_elements().collect::<Vec<_>>();
    let titles = titles.child_elements().collect::<Vec<_>>();
    if !headings.len().is_multiple_of(2) || titles.iter().any(|title| !scalar(title, "lpstr")) {
        return Ok(None);
    }
    let mut worksheets = None;
    let mut named_ranges = Vec::new();
    let mut offset = 0usize;
    for pair in headings.chunks_exact(2) {
        let Some(label) = variant(pair[0], "lpstr") else {
            return Ok(None);
        };
        let Some(count) = variant(pair[1], "i4")
            .and_then(|count| count.text_content().trim().parse::<usize>().ok())
        else {
            return Ok(None);
        };
        let Some(end) = offset.checked_add(count).filter(|end| *end <= titles.len()) else {
            return Ok(None);
        };
        match label.text_content().as_str() {
            "Worksheets" => {
                if worksheets.replace(offset..end).is_some() {
                    return Ok(None);
                }
            }
            "Named Ranges" => named_ranges.push(offset..end),
            _ => {}
        }
        offset = end;
    }
    let Some(worksheets) = worksheets else {
        return Ok(None);
    };
    if offset != titles.len()
        || worksheets.len() != original.sheets.len()
        || titles[worksheets.clone()]
            .iter()
            .zip(&original.sheets)
            .any(|(title, sheet)| title.text_content() != sheet.name)
    {
        return Ok(None);
    }
    let mut edits = Vec::new();
    for (title, sheet) in titles[worksheets].iter().zip(&workbook.sheets) {
        if title.text_content() != sheet.name {
            replace_text(title, sheet.name.clone(), &mut edits)?;
        }
    }
    for range in named_ranges {
        for title in &titles[range] {
            let text = title.text_content();
            for (sheet, origin) in workbook.sheets.iter().zip(origins) {
                let Some(old) = origin.and_then(|index| original.sheets.get(index)) else {
                    continue;
                };
                if old.name == sheet.name {
                    continue;
                }
                let plain = format!("{}!", old.name);
                let quoted = format!("'{}'!", old.name.replace('\'', "''"));
                if let Some(name) = text.strip_prefix(&quoted).or_else(|| text.strip_prefix(&plain))
                    && !name.is_empty()
                {
                    replace_text(
                        title,
                        format!("{}!{name}", quoted_sheet(&sheet.name)),
                        &mut edits,
                    )?;
                    break;
                }
            }
        }
    }
    if edits.is_empty() {
        Ok(None)
    } else {
        edits.sort_by_key(|(range, _)| range.start);
        part.splice(&edits).map(Some)
    }
}

fn unique_child<'a>(element: &'a Element, namespace: &str, local: &str) -> Option<&'a Element> {
    let mut children = element
        .child_elements()
        .filter(|child| child.is(namespace, local));
    let first = children.next()?;
    children.next().is_none().then_some(first)
}

fn vector<'a>(element: &'a Element, base_type: &str) -> Option<&'a Element> {
    let vector = unique_child(element, NS_VT, "vector")?;
    let size = vector.attribute(None, "size")?.parse::<usize>().ok()?;
    (vector.attribute(None, "baseType")? == base_type && vector.child_elements().count() == size)
        .then_some(vector)
}

fn scalar(element: &Element, local: &str) -> bool {
    element.is(NS_VT, local) && element.child_elements().next().is_none()
}

fn variant<'a>(element: &'a Element, local: &str) -> Option<&'a Element> {
    if !element.is(NS_VT, "variant") || element.child_elements().count() != 1 {
        return None;
    }
    unique_child(element, NS_VT, local).filter(|child| scalar(child, local))
}

fn replace_text(
    element: &Element,
    text: String,
    edits: &mut Vec<(Range<usize>, Replacement)>,
) -> Result<(), ParseError> {
    let range = element.splice_target().ok_or_else(|| {
        ParseError::UnsupportedEdit("app title has no replaceable content".into())
    })?;
    edits.push((range, Replacement::Text(text)));
    Ok(())
}
