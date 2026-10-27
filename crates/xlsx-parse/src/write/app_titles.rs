use std::collections::HashSet;
use std::ops::Range;

use xlsx_model::Workbook;

use crate::chart::quoted_sheet;
use crate::package::Relationship;
use crate::tree::{Element, Part, Replacement, names_are_resolvable};
use crate::xml::resolve_part_path;

const NS_APP: &str = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const NS_VT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

/// the extended properties part the package root names, if exactly one.
pub(super) fn part_path(relationships: &[Relationship]) -> Option<String> {
    let mut targets = relationships
        .iter()
        .filter(|relationship| {
            relationship.has_type("extended-properties")
                && relationship.attribute("TargetMode") != Some("External")
        })
        .filter_map(Relationship::target);
    let target = targets.next()?;
    targets
        .next()
        .is_none()
        .then(|| resolve_part_path("", target))
}

/// `None` leaves the part byte for byte as it is, including when it cannot
/// be read or a title carries more than plain text.
pub(super) fn patch(
    bytes: &[u8],
    workbook: &Workbook,
    original: &Workbook,
    origins: &[Option<usize>],
) -> Option<Vec<u8>> {
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
            .zip(origins)
            .enumerate()
            .all(|(index, (sheet, origin))| {
                *origin == Some(index) && sheet.name == original.sheets[index].name
            })
    {
        return None;
    }
    let part = Part::decode(bytes).ok()?;
    let root = part.tree().ok()?;
    if !root.is(NS_APP, "Properties") || !names_are_resolvable(&root) {
        return None;
    }
    let headings = unique_child(&root, NS_APP, "HeadingPairs")
        .and_then(|heading| vector(heading, "variant"))?;
    let titles =
        unique_child(&root, NS_APP, "TitlesOfParts").and_then(|titles| vector(titles, "lpstr"))?;
    let headings = headings.child_elements().collect::<Vec<_>>();
    let titles = titles.child_elements().collect::<Vec<_>>();
    if !headings.len().is_multiple_of(2) || titles.iter().any(|title| !scalar(title, "lpstr")) {
        return None;
    }
    let mut worksheets = None;
    let mut named_ranges = Vec::new();
    let mut offset = 0usize;
    for pair in headings.as_chunks::<2>().0 {
        let label = variant(pair[0], "lpstr")?;
        let count = variant(pair[1], "i4")
            .and_then(|count| count.text_content().trim().parse::<usize>().ok())?;
        let end = offset
            .checked_add(count)
            .filter(|end| *end <= titles.len())?;
        match label.text_content().as_str() {
            "Worksheets" => {
                if worksheets.replace(offset..end).is_some() {
                    return None;
                }
            }
            "Named Ranges" => named_ranges.push(offset..end),
            _ => {}
        }
        offset = end;
    }
    let worksheets = worksheets?;
    if offset != titles.len()
        || worksheets.len() != original.sheets.len()
        || titles[worksheets.clone()]
            .iter()
            .zip(&original.sheets)
            .any(|(title, sheet)| title.text_content() != sheet.name)
    {
        return None;
    }
    let mut edits = Vec::new();
    for (title, sheet) in titles[worksheets].iter().zip(&workbook.sheets) {
        if title.text_content() != sheet.name {
            replace_text(&part, title, sheet.name.clone(), &mut edits)?;
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
                if let Some(name) = text
                    .strip_prefix(&quoted)
                    .or_else(|| text.strip_prefix(&plain))
                    && !name.is_empty()
                {
                    replace_text(
                        &part,
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
        None
    } else {
        edits.sort_by_key(|(range, _)| range.start);
        part.splice(&edits).ok()
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

/// a title whose content holds a comment, instruction or CDATA section is not
/// rewritten, so nothing beyond its text is lost.
fn replace_text(
    part: &Part,
    element: &Element,
    text: String,
    edits: &mut Vec<(Range<usize>, Replacement)>,
) -> Option<()> {
    let range = element
        .splice_target()
        .filter(|range| !part.text()[range.clone()].contains('<'))?;
    edits.push((range, Replacement::Text(text)));
    Some(())
}
