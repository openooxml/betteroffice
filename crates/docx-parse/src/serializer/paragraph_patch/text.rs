use std::collections::{BTreeMap, BTreeSet, btree_map::Entry};

use serde_json::{Value, json};

use crate::paragraph_patch::{
    Attrs, DonorV1, ParagraphPatchV1, PatchErrorCode, TextEditV1, UnitV1,
};
use crate::serializer::xml_writer::escape_xml;

use super::lower::{Cell, Lowered, Node, WORD, Xml};
use super::{
    Failure, apply_bytes, apply_deltas, attr_deltas, flatten, insert_into_root, properties,
    unit_attrs,
};

struct Insert {
    node: usize,
    at: u32,
    units: Vec<UnitV1>,
    properties: String,
}

struct Token {
    properties: String,
    raw: String,
    text: String,
    inserted: bool,
    element: Option<&'static str>,
}

enum TextPart<'a> {
    Text(&'a str),
    Element(&'static str),
}

/// `text` as the model writer splits it: tabs and the two hyphens are their own elements.
fn text_parts(text: &str) -> Result<Vec<TextPart<'_>>, Failure> {
    let mut parts = Vec::new();
    let mut start = 0;
    for (at, ch) in text.char_indices() {
        let element = match ch {
            '\t' => "<w:tab/>",
            '\u{ad}' => "<w:softHyphen/>",
            '\u{2011}' => "<w:noBreakHyphen/>",
            '\n' | '\u{b}' | '\u{c}' | '\u{f000}'..='\u{f8ff}' => {
                return Err(Failure::structure(
                    "ambiguous: inserted text holds a break or symbol scalar",
                ));
            }
            _ => continue,
        };
        if start < at {
            parts.push(TextPart::Text(&text[start..at]));
        }
        parts.push(TextPart::Element(element));
        start = at + ch.len_utf8();
    }
    if start < text.len() || parts.is_empty() {
        parts.push(TextPart::Text(&text[start..]));
    }
    Ok(parts)
}

fn text_xml(
    text: &str,
    mut element: impl FnMut(&str, &str) -> Result<String, Failure>,
) -> Result<String, Failure> {
    let mut xml = String::new();
    for part in text_parts(text)? {
        match part {
            TextPart::Text(text) => xml.push_str(&element(&escape_text(text), text)?),
            TextPart::Element(special) => xml.push_str(special),
        }
    }
    Ok(xml)
}

fn boundary_error(detail: &str) -> Failure {
    Failure::new(PatchErrorCode::InvalidBoundary, detail)
}

fn conflict(detail: &str) -> Failure {
    Failure::new(PatchErrorCode::PatchConflict, detail)
}

fn donor_error(detail: &str) -> Failure {
    Failure::new(PatchErrorCode::AmbiguousDonor, detail)
}

pub(super) fn write(
    source: &str,
    lowered: &Lowered,
    patch: &ParagraphPatchV1,
) -> Result<(String, Vec<UnitV1>), Failure> {
    let boundary = |offset| {
        offset == lowered.len
            || lowered.cells.iter().any(|cell| {
                #[cfg(test)]
                super::work::add(1);
                cell.from == offset
            })
    };
    let mut retained: Vec<_> = lowered
        .cells
        .iter()
        .map(|cell| Some(cell.unit.clone()))
        .collect();
    let mut formatting: Vec<Option<String>> = vec![None; retained.len()];
    let mut symbols: Vec<Option<String>> = vec![None; retained.len()];
    let restyle = patch
        .para_deltas
        .iter()
        .any(|delta| delta.path == ["pStyle"]);
    let empty = Attrs::new();
    let inherited: Vec<_> = patch
        .seed
        .units
        .iter()
        .enumerate()
        .flat_map(|(at, unit)| {
            let count = match unit {
                UnitV1::Text { text, .. } => text.chars().count(),
                UnitV1::Atom { .. } => 1,
            };
            std::iter::repeat_n(patch.seed.inherit.runs.get(at).unwrap_or(&empty), count)
        })
        .collect();
    let mut run_ranges = Vec::new();
    for edit in &patch.run_edits {
        if edit.from >= edit.to || !boundary(edit.from) || !boundary(edit.to) {
            return Err(boundary_error("invalid run edit boundary"));
        }
        if run_ranges
            .iter()
            .any(|&(from, to)| edit.from < to && from < edit.to)
        {
            return Err(conflict("overlapping run edits"));
        }
        run_ranges.push((edit.from, edit.to));
        let mut property_cache = BTreeMap::new();
        for (at, cell) in lowered.cells.iter().enumerate() {
            #[cfg(test)]
            super::work::add(1);
            if cell.from >= edit.from && cell.to <= edit.to {
                if cell.retained {
                    return Err(Failure::structure(
                        "marker: formatting touches retained content",
                    ));
                }
                let mut attrs = unit_attrs(&cell.unit).clone();
                apply_deltas(&mut attrs, &edit.deltas)?;
                if attrs != edit.target {
                    return Err(Failure::new(
                        PatchErrorCode::PropertyConflict,
                        "run target differs from deltas",
                    ));
                }
                let node = &lowered.xml.nodes[cell.node];
                if node.word("sym")
                    && attrs.get("fontFamily") != unit_attrs(&cell.unit).get("fontFamily")
                {
                    let replacement = symbol_bytes(source, node, &attrs)?;
                    if replacement != source[node.span.clone()] {
                        symbols[at] = Some(replacement);
                    }
                }
                let properties = match property_cache.entry(cell.run) {
                    Entry::Occupied(entry) => entry.into_mut(),
                    Entry::Vacant(entry) => {
                        let original = run_properties(source, lowered, cell.run);
                        let written = properties::run_properties(
                            original,
                            &edit.deltas,
                            &edit.inherit,
                            &edit.target,
                        )?;
                        let written = if restyle {
                            let deltas = properties::restyle_run_deltas(
                                original,
                                &edit.deltas,
                                unit_attrs(&cell.unit),
                                inherited.get(at).copied().unwrap_or(&empty),
                                &edit.inherit,
                            )?;
                            properties::run_properties(
                                original,
                                &deltas,
                                &edit.inherit,
                                &edit.target,
                            )?
                        } else {
                            written
                        };
                        entry.insert(written)
                    }
                };
                if properties.as_str() != run_properties(source, lowered, cell.run) {
                    formatting[at] = Some(properties.clone());
                }
                *super::unit_attrs_mut(retained[at].as_mut().unwrap()) = attrs;
            }
        }
    }
    let mut inserts = Vec::new();
    let mut changed_runs = BTreeSet::new();
    let mut text_runs = BTreeSet::new();
    let mut ranges = Vec::new();
    let mut replay: BTreeMap<u32, Vec<UnitV1>> = BTreeMap::new();
    let mut empty_insert = None;
    for edit in &patch.text_edits {
        if edit
            .insert
            .iter()
            .any(|insert| insert.text.contains(['\r', '\n']))
        {
            return Err(Failure::structure("ambiguous: newline in inserted text"));
        }
        if edit
            .insert
            .iter()
            .any(|insert| !super::valid_xml_text(&insert.text))
        {
            return Err(Failure::new(
                PatchErrorCode::InvalidPlan,
                "invalid inserted XML scalar",
            ));
        }
        if edit.from > edit.to || !boundary(edit.from) || !boundary(edit.to) {
            return Err(boundary_error("invalid text edit boundary"));
        }
        if ranges
            .iter()
            .any(|&(from, to)| (edit.from < to && from < edit.to) || edit.from == from)
        {
            return Err(conflict("overlapping text edits"));
        }
        if run_ranges
            .iter()
            .any(|&(from, to)| edit.from < to && from < edit.to)
        {
            return Err(conflict("text and run edits overlap"));
        }
        ranges.push((edit.from, edit.to));
        if lowered
            .markers
            .iter()
            .chain(&lowered.fences)
            .any(|&at| edit.from <= at && at <= edit.to)
        {
            return Err(Failure::structure(
                "marker: text edit touches a marker or fence",
            ));
        }
        let covered: Vec<_> = lowered
            .cells
            .iter()
            .enumerate()
            .filter(|(_, cell)| edit.from <= cell.from && cell.to <= edit.to)
            .collect();
        if covered
            .iter()
            .any(|(_, cell)| cell.retained || !matches!(cell.unit, UnitV1::Text { .. }))
        {
            return Err(Failure::structure("text edit changes an atom"));
        }
        let scopes: BTreeSet<_> = covered
            .iter()
            .map(|(_, cell)| lowered.runs[cell.run].scope)
            .collect();
        if scopes.len() > 1 {
            return Err(Failure::structure("hyperlink: text edit crosses a wrapper"));
        }
        let donor = choose_donor(lowered, edit, &covered)?;
        if donor.is_some_and(|cell| cell.retained) {
            return Err(Failure::structure("field: insertion uses retained content"));
        }
        text_runs.extend(covered.iter().map(|(_, cell)| cell.run));
        if let Some(donor) = donor {
            text_runs.insert(donor.run);
        }
        let units = edit
            .insert
            .iter()
            .flat_map(|insert| {
                flatten(&[UnitV1::Text {
                    text: insert.text.clone(),
                    attrs: insert.attrs.clone(),
                }])
            })
            .collect();
        replay.insert(edit.from, units);
        match (donor, &edit.donor) {
            (Some(donor), DonorV1::At { at }) => {
                let named = lowered
                    .cells
                    .iter()
                    .enumerate()
                    .find(|(_, cell)| cell.from <= *at && *at < cell.to);
                if !named.is_some_and(|(_, cell)| {
                    cell.run == donor.run && matches!(cell.unit, UnitV1::Text { .. })
                }) {
                    return Err(donor_error("plan names a different donor"));
                }
                let mut donor_attrs = unit_attrs(&donor.unit).clone();
                if lowered.xml.nodes[donor.node].word("sym") {
                    donor_attrs.remove("fontFamily");
                }
                let original = run_properties(source, lowered, donor.run);
                for insert in &edit.insert {
                    let deltas = attr_deltas(&donor_attrs, &insert.attrs);
                    let properties = properties::run_properties(
                        original,
                        &deltas,
                        &insert.inherit,
                        &insert.attrs,
                    )?;
                    let properties = if restyle {
                        let deltas = properties::restyle_run_deltas(
                            original,
                            &deltas,
                            &donor_attrs,
                            named
                                .and_then(|(at, _)| inherited.get(at).copied())
                                .unwrap_or(&empty),
                            &insert.inherit,
                        )?;
                        properties::run_properties(
                            original,
                            &deltas,
                            &insert.inherit,
                            &insert.attrs,
                        )?
                    } else {
                        properties
                    };
                    inserts.push(Insert {
                        node: donor.node,
                        at: edit.from,
                        units: flatten(&[UnitV1::Text {
                            text: insert.text.clone(),
                            attrs: insert.attrs.clone(),
                        }]),
                        properties,
                    });
                }
                changed_runs.insert(donor.run);
            }
            (None, DonorV1::Pilcrow) if lowered.cells.is_empty() => {
                if empty_insert.is_some() {
                    return Err(donor_error("multiple empty insertions"));
                }
                empty_insert = Some(edit);
            }
            _ => return Err(donor_error("no unambiguous donor")),
        }
        for (at, cell) in covered {
            retained[at] = None;
            changed_runs.insert(cell.run);
        }
    }
    for (at, cell) in lowered.cells.iter().enumerate() {
        #[cfg(test)]
        super::work::add(1);
        if formatting[at].is_some() || symbols[at].is_some() {
            changed_runs.insert(cell.run);
        }
    }
    let mut result_units = Vec::new();
    for (at, cell) in lowered.cells.iter().enumerate() {
        #[cfg(test)]
        super::work::add(1);
        if let Some(inserted) = replay.remove(&cell.from) {
            result_units.extend(inserted);
        }
        if let Some(unit) = &retained[at] {
            result_units.push(unit.clone());
        }
    }
    if let Some(inserted) = replay.remove(&lowered.len) {
        result_units.extend(inserted);
    }
    if !replay.is_empty() {
        return Err(boundary_error("unreplayed insertion"));
    }
    let mut node_cells: BTreeMap<usize, Vec<usize>> = BTreeMap::new();
    for (at, cell) in lowered.cells.iter().enumerate() {
        #[cfg(test)]
        super::work::add(1);
        node_cells.entry(cell.node).or_default().push(at);
    }
    let cells = CellWrites {
        by_node: &node_cells,
        retained: &retained,
        formatting: &formatting,
        symbols: &symbols,
    };
    let mut writes = Vec::new();
    for run in changed_runs {
        let replacement = write_run(source, lowered, run, &cells, &inserts)?;
        writes.push((run, replacement));
    }
    let mut output = apply_bytes(source, coalesce_edits(source, lowered, &writes, &text_runs))?;
    if let Some(edit) = empty_insert {
        output = write_empty(source, lowered, edit)?;
    }
    Ok((output, result_units))
}

struct Fragment {
    opener: String,
    key: String,
    leading: String,
    properties: String,
    content: String,
    close: String,
}

fn plain_run(source: &str, xml: &Xml, at: usize) -> bool {
    let node = &xml.nodes[at];
    if !node.word("r") || node.empty || has_identifier(node) {
        return false;
    }
    let mut cursor = node.open.end;
    for (index, &at) in node.children.iter().enumerate() {
        let child = &xml.nodes[at];
        if !source[cursor..child.span.start].trim().is_empty()
            || !(child.word("t") || (index == 0 && child.word("rPr")))
            || xml.nodes[at..]
                .iter()
                .take_while(|descendant| descendant.span.start < child.span.end)
                .any(has_identifier)
        {
            return false;
        }
        if child.word("t")
            && (!child.children.is_empty()
                || (!child.empty && source[child.open.end..child.close.start].contains('<')))
        {
            return false;
        }
        cursor = child.span.end;
    }
    source[cursor..node.close.start].trim().is_empty()
}

fn fragments(
    source: &str,
    lowered: &Lowered,
    run: usize,
    replacement: &str,
) -> Option<Vec<Fragment>> {
    if source[lowered.xml.nodes[0].open.end..].contains("xmlns")
        || !plain_run(source, &lowered.xml, lowered.runs[run].node)
    {
        return None;
    }
    let wrapped = format!("<w:p>{replacement}</w:p>");
    let xml = Xml::parse(&wrapped, &lowered.xml.root_namespaces).ok()?;
    xml.nodes[0]
        .children
        .iter()
        .map(|&at| {
            if !plain_run(&wrapped, &xml, at) {
                return None;
            }
            let node = &xml.nodes[at];
            let opener = &wrapped[node.open.clone()];
            let mut key = String::new();
            let mut cursor = node.open.start;
            for attribute in &node.attributes {
                let Some((prefix, local)) = attribute.name.split_once(':') else {
                    continue;
                };
                if prefix != "w"
                    || !local.starts_with("rsid")
                    || xml.root_namespaces.get(prefix).map(String::as_str) != Some(WORD)
                {
                    continue;
                }
                key.push_str(&wrapped[cursor..attribute.token.start]);
                cursor = attribute.token.end;
            }
            key.push_str(&wrapped[cursor..node.open.end]);
            let properties = xml.child(at, "rPr").ok()?;
            let (leading, properties, content) = if let Some(at) = properties {
                let span = &xml.nodes[at].span;
                (
                    &wrapped[node.open.end..span.start],
                    &wrapped[span.clone()],
                    &wrapped[span.end..node.close.start],
                )
            } else {
                ("", "", &wrapped[node.open.end..node.close.start])
            };
            Some(Fragment {
                opener: opener.to_owned(),
                key,
                leading: leading.to_owned(),
                properties: properties.to_owned(),
                content: content.to_owned(),
                close: wrapped[node.close.clone()].to_owned(),
            })
        })
        .collect()
}

fn flush_fragment(output: &mut String, fragment: Fragment) {
    output.push_str(&fragment.opener);
    output.push_str(&fragment.leading);
    output.push_str(&fragment.properties);
    output.push_str(&fragment.content);
    output.push_str(&fragment.close);
}

fn coalesce_edits(
    source: &str,
    lowered: &Lowered,
    writes: &[(usize, String)],
    text_runs: &BTreeSet<usize>,
) -> Vec<(std::ops::Range<usize>, String)> {
    let mut edits = Vec::new();
    let mut start = 0;
    while start < writes.len() {
        let (first_run, replacement) = &writes[start];
        let first_node = &lowered.xml.nodes[lowered.runs[*first_run].node];
        let parsed = text_runs
            .contains(first_run)
            .then(|| fragments(source, lowered, *first_run, replacement))
            .flatten();
        let Some(parsed) = parsed else {
            edits.push((first_node.span.clone(), replacement.clone()));
            start += 1;
            continue;
        };
        let mut groups = vec![("", parsed)];
        let mut end = start;
        while let Some((run, replacement)) = writes.get(end + 1) {
            let previous = &lowered.runs[writes[end].0];
            let next = &lowered.runs[*run];
            let previous_node = &lowered.xml.nodes[previous.node];
            let next_node = &lowered.xml.nodes[next.node];
            if *run != writes[end].0 + 1
                || previous.scope != next.scope
                || previous_node.parent != next_node.parent
                || !text_runs.contains(run)
            {
                break;
            }
            let gap = &source[previous_node.span.end..next_node.span.start];
            if !gap.trim().is_empty() {
                break;
            }
            let Some(parsed) = fragments(source, lowered, *run, replacement) else {
                break;
            };
            groups.push((gap, parsed));
            end += 1;
        }
        let mut output = String::new();
        let mut pending: Option<Fragment> = None;
        for (mut gap, parsed) in groups {
            if parsed.is_empty() {
                if let Some(previous) = pending.as_mut() {
                    previous.content.push_str(gap);
                } else {
                    output.push_str(gap);
                }
                continue;
            }
            for candidate in parsed {
                if let Some(previous) = pending.as_mut()
                    && previous.key == candidate.key
                    && previous.properties == candidate.properties
                {
                    previous.content.push_str(gap);
                    previous.content.push_str(&candidate.leading);
                    previous.content.push_str(&candidate.content);
                } else {
                    if let Some(previous) = pending.take() {
                        flush_fragment(&mut output, previous);
                    }
                    output.push_str(gap);
                    pending = Some(candidate);
                }
                gap = "";
            }
        }
        if let Some(previous) = pending {
            flush_fragment(&mut output, previous);
        }
        let last_node = &lowered.xml.nodes[lowered.runs[writes[end].0].node];
        edits.push((first_node.span.start..last_node.span.end, output));
        start = end + 1;
    }
    edits
}

fn choose_donor<'a>(
    lowered: &'a Lowered,
    edit: &TextEditV1,
    covered: &[(usize, &'a Cell)],
) -> Result<Option<&'a Cell>, Failure> {
    if edit.from < edit.to {
        return Ok(covered.first().map(|(_, cell)| *cell));
    }
    let previous = lowered
        .cells
        .iter()
        .rev()
        .find(|cell| cell.to == edit.from && matches!(cell.unit, UnitV1::Text { .. }));
    let following = lowered
        .cells
        .iter()
        .find(|cell| cell.from == edit.from && matches!(cell.unit, UnitV1::Text { .. }));
    if let (Some(previous), Some(following)) = (previous, following)
        && lowered.runs[previous.run].scope != lowered.runs[following.run].scope
    {
        return Err(donor_error("insertion at hyperlink edge"));
    }
    Ok(previous.or(if edit.from == 0 { following } else { None }))
}

struct CellWrites<'a> {
    by_node: &'a BTreeMap<usize, Vec<usize>>,
    retained: &'a [Option<UnitV1>],
    formatting: &'a [Option<String>],
    symbols: &'a [Option<String>],
}

fn write_run(
    source: &str,
    lowered: &Lowered,
    run: usize,
    writes: &CellWrites,
    inserts: &[Insert],
) -> Result<String, Failure> {
    let CellWrites {
        by_node,
        retained,
        formatting,
        symbols,
    } = writes;
    let source_run = &lowered.runs[run];
    let node = &lowered.xml.nodes[source_run.node];
    let original = run_properties(source, lowered, run);
    if source_run.from < source_run.to
        && node.children.iter().all(|&child| {
            #[cfg(test)]
            super::work::add(1);
            Some(child) == source_run.properties || lowered.xml.nodes[child].word("t")
        })
        && node
            .children
            .iter()
            .flat_map(|child| by_node.get(child).into_iter().flatten())
            .all(|&at| {
                #[cfg(test)]
                super::work::add(1);
                retained[at].is_none()
            })
        && !inserts
            .iter()
            .any(|insert| node.children.contains(&insert.node) && !insert.units.is_empty())
    {
        return Ok(String::new());
    }
    let mut fragments: Vec<(String, String)> = Vec::new();
    let mut cursor = node.open.end;
    let mut leading = String::new();
    for &child in &node.children {
        #[cfg(test)]
        super::work::add(1);
        let child_node = &lowered.xml.nodes[child];
        if Some(child) == source_run.properties {
            leading.push_str(&source[cursor..child_node.span.start]);
            cursor = child_node.span.end;
            continue;
        }
        let gap = &source[cursor..child_node.span.start];
        cursor = child_node.span.end;
        let cells: Vec<_> = by_node
            .get(&child)
            .map_or(&[][..], Vec::as_slice)
            .iter()
            .map(|&at| (at, &lowered.cells[at]))
            .collect();
        let mut tokens = Vec::new();
        for (at, cell) in &cells {
            #[cfg(test)]
            super::work::add(1);
            for insert in inserts
                .iter()
                .filter(|insert| insert.node == child && insert.at == cell.from)
            {
                tokens.extend(insert_tokens(insert)?);
            }
            if let Some(unit) = &retained[*at] {
                let text = match unit {
                    UnitV1::Text { text, .. } => text.clone(),
                    _ => String::new(),
                };
                tokens.push(Token {
                    element: None,
                    properties: formatting[*at].as_deref().unwrap_or(original).to_owned(),
                    raw: symbols[*at]
                        .clone()
                        .unwrap_or_else(|| source[cell.raw.clone()].to_owned()),
                    text,
                    inserted: false,
                });
            }
        }
        if let Some((_, last)) = cells.last() {
            for insert in inserts
                .iter()
                .filter(|insert| insert.node == child && insert.at == last.to)
            {
                tokens.extend(insert_tokens(insert)?);
            }
        }
        if child_node.word("t") {
            if tokens.is_empty() {
                let element = text_element(source, child_node, "", "")?;
                append(&mut fragments, original, &format!("{gap}{element}"));
            } else {
                let mut groups: Vec<(String, String, String, Option<&'static str>)> = Vec::new();
                for token in tokens {
                    if let Some((properties, raw, text, None)) = groups.last_mut()
                        && token.element.is_none()
                        && *properties == token.properties
                    {
                        raw.push_str(&token.raw);
                        text.push_str(&token.text);
                    } else {
                        groups.push((token.properties, token.raw, token.text, token.element));
                    }
                }
                let texts = groups.iter().filter(|group| group.3.is_none()).count();
                if texts > 1 && has_identifier(child_node) {
                    return Err(Failure::structure(
                        "ambiguous: splitting duplicates a text identifier",
                    ));
                }
                for (index, (properties, raw, text, special)) in groups.into_iter().enumerate() {
                    let gap = if index == 0 { gap } else { "" };
                    let element = match special {
                        Some(special) => special.to_owned(),
                        None => text_element(source, child_node, &raw, &text)?,
                    };
                    append(&mut fragments, &properties, &format!("{gap}{element}"));
                }
            }
        } else if cells.is_empty() {
            append(
                &mut fragments,
                original,
                &format!("{gap}{}", &source[child_node.span.clone()]),
            );
        } else {
            let mut groups: Vec<(String, String, String, bool, Option<&'static str>)> = Vec::new();
            for token in tokens {
                if let Some((properties, raw, text, inserted, None)) = groups.last_mut()
                    && token.element.is_none()
                    && *properties == token.properties
                    && *inserted == token.inserted
                {
                    raw.push_str(&token.raw);
                    text.push_str(&token.text);
                } else {
                    groups.push((
                        token.properties,
                        token.raw,
                        token.text,
                        token.inserted,
                        token.element,
                    ));
                }
            }
            if groups.is_empty() {
                append(&mut fragments, original, gap);
            }
            for (index, (properties, raw, text, inserted, special)) in
                groups.into_iter().enumerate()
            {
                let content = match special {
                    Some(special) => special.to_owned(),
                    None if inserted => generated_text(&raw, &text),
                    None => raw,
                };
                let gap = if index == 0 { gap } else { "" };
                append(&mut fragments, &properties, &format!("{gap}{content}"));
            }
        }
    }
    let trailing = if node.empty {
        ""
    } else {
        &source[cursor..node.close.start]
    };
    if fragments.is_empty() {
        fragments.push((original.to_owned(), String::new()));
    }
    fragments.last_mut().unwrap().1.push_str(trailing);
    let properties_identified = || {
        source_run.properties.is_some_and(|at| {
            let span = &lowered.xml.nodes[at].span;
            lowered.xml.nodes.iter().any(|node| {
                node.span.start >= span.start && node.span.end <= span.end && has_identifier(node)
            })
        })
    };
    if fragments.len() > 1 && (has_identifier(node) || properties_identified()) {
        return Err(Failure::structure(
            "splitting would duplicate a run identifier",
        ));
    }
    let mut replacement = String::new();
    for (properties, content) in fragments {
        #[cfg(test)]
        super::work::add(
            node.open.len() + leading.len() + properties.len() + content.len() + node.close.len(),
        );
        let opener = &source[node.open.clone()];
        if node.empty {
            replacement.push_str(opener.trim_end_matches('>').trim_end_matches('/'));
            replacement.push('>');
        } else {
            replacement.push_str(opener);
        }
        replacement.push_str(&leading);
        replacement.push_str(&properties);
        replacement.push_str(&content);
        if node.empty {
            replacement.push_str(&format!("</{}>", node.name));
        } else {
            replacement.push_str(&source[node.close.clone()]);
        }
    }
    Ok(replacement)
}

fn write_empty(source: &str, lowered: &Lowered, edit: &TextEditV1) -> Result<String, Failure> {
    let empty: Vec<_> = lowered
        .runs
        .iter()
        .enumerate()
        .filter(|(_, run)| {
            run.from == run.to
                && lowered.xml.nodes[0].children.contains(&run.node)
                && lowered.xml.nodes[run.node]
                    .children
                    .iter()
                    .all(|at| Some(*at) == run.properties || lowered.xml.nodes[*at].word("t"))
        })
        .collect();
    if empty.len() > 1 {
        return Err(donor_error("multiple empty runs"));
    }
    let ppr_rpr = lowered
        .ppr
        .map(|at| lowered.xml.child(at, "rPr"))
        .transpose()?
        .flatten();
    let original = empty
        .first()
        .map(|(at, _)| run_properties(source, lowered, *at))
        .or_else(|| ppr_rpr.map(|at| &source[lowered.xml.nodes[at].span.clone()]))
        .unwrap_or("");
    let defaults = direct_mark_attrs(original)?;
    let mut content = String::new();
    for insert in &edit.insert {
        let mut effective = insert.inherit.clone();
        for (key, direct) in &defaults {
            overlay_default(effective.entry(key.clone()).or_insert(Value::Null), direct);
        }
        let deltas = attr_deltas(&effective, &insert.attrs);
        let properties =
            properties::run_properties(original, &deltas, &insert.inherit, &insert.attrs)?;
        let run = if let Some((_, run)) = empty.first() {
            let node = &lowered.xml.nodes[run.node];
            let children: Vec<_> = node
                .children
                .iter()
                .copied()
                .filter(|at| Some(*at) != run.properties)
                .collect();
            if children.len() > 1 {
                return Err(donor_error("multiple empty text elements"));
            }
            let text = if let Some(&at) = children.first() {
                text_xml(&insert.text, |raw, text| {
                    text_element(source, &lowered.xml.nodes[at], raw, text)
                })?
            } else {
                text_xml(&insert.text, |raw, text| Ok(generated_text(raw, text)))?
            };
            let opener = source[node.open.clone()]
                .trim_end_matches('>')
                .trim_end_matches('/');
            let close = if node.empty {
                format!("</{}>", node.name)
            } else {
                source[node.close.clone()].to_owned()
            };
            format!("{opener}>{properties}{text}{close}")
        } else {
            let text = text_xml(&insert.text, |raw, text| Ok(generated_text(raw, text)))?;
            format!("<w:r>{properties}{text}</w:r>")
        };
        content.push_str(&without_identifiers(&run, &lowered.xml.root_namespaces)?);
    }
    if let Some((_, run)) = empty.first() {
        apply_bytes(
            source,
            vec![(lowered.xml.nodes[run.node].span.clone(), content)],
        )
    } else {
        insert_into_root(source, &lowered.xml.nodes[0], &content, false)
    }
}

fn overlay_default(target: &mut Value, direct: &Value) {
    if direct.is_null() {
        return;
    }
    if let Some(object) = target.as_object_mut()
        && let Some(direct) = direct.as_object()
    {
        for (key, value) in direct {
            overlay_default(object.entry(key.clone()).or_insert(Value::Null), value);
        }
    } else {
        *target = direct.clone();
    }
}

fn symbol_bytes(source: &str, node: &Node, attrs: &Attrs) -> Result<String, Failure> {
    let font = attrs.get("fontFamily");
    let selected = font
        .and_then(|font| font.get("ascii"))
        .and_then(Value::as_str);
    let original = node.attr(source, "w:font").unwrap_or_default();
    if selected == Some(original.as_str()) || (selected.is_none() && original.is_empty()) {
        return Ok(source[node.span.clone()].to_owned());
    }
    let selected = selected.ok_or_else(|| {
        Failure::new(
            PatchErrorCode::PropertyConflict,
            "symbol font reset is unrepresentable",
        )
    })?;
    if ["hAnsi", "eastAsia", "cs"].iter().any(|slot| {
        font.and_then(|font| font.get(*slot))
            .and_then(Value::as_str)
            != Some(selected)
    }) {
        return Err(Failure::new(
            PatchErrorCode::PropertyConflict,
            "symbol needs one concrete font",
        ));
    }
    let attribute = node
        .attributes
        .iter()
        .find(|attribute| attribute.name == "w:font")
        .ok_or_else(|| Failure::structure("symbol has no font attribute"))?;
    let span = attribute.value.start - node.span.start..attribute.value.end - node.span.start;
    apply_bytes(
        &source[node.span.clone()],
        vec![(span, escape_xml(selected))],
    )
}

fn has_identifier(node: &Node) -> bool {
    node.attributes.iter().any(|attribute| attribute.identifier)
}

fn without_identifiers(
    source: &str,
    namespaces: &BTreeMap<String, String>,
) -> Result<String, Failure> {
    let xml = super::lower::Xml::parse(source, namespaces)?;
    let edits = xml
        .nodes
        .iter()
        .flat_map(|node| &node.attributes)
        .filter(|attribute| attribute.identifier)
        .map(|attribute| (attribute.token.clone(), String::new()))
        .collect();
    apply_bytes(source, edits)
}

fn insert_tokens(insert: &Insert) -> Result<Vec<Token>, Failure> {
    let mut tokens = Vec::new();
    for unit in &insert.units {
        let UnitV1::Text { text, .. } = unit else {
            continue;
        };
        for part in text_parts(text)? {
            let (text, element) = match part {
                TextPart::Text(text) => (text, None),
                TextPart::Element(element) => ("", Some(element)),
            };
            tokens.push(Token {
                properties: insert.properties.clone(),
                raw: escape_text(text),
                text: text.to_owned(),
                inserted: true,
                element,
            });
        }
    }
    Ok(tokens)
}

fn append(fragments: &mut Vec<(String, String)>, properties: &str, content: &str) {
    if let Some((previous, text)) = fragments.last_mut()
        && previous == properties
    {
        text.push_str(content);
    } else {
        fragments.push((properties.to_owned(), content.to_owned()));
    }
}

fn run_properties<'a>(source: &'a str, lowered: &Lowered, run: usize) -> &'a str {
    lowered.runs[run]
        .properties
        .map_or("", |at| &source[lowered.xml.nodes[at].span.clone()])
}

fn escape_text(text: &str) -> String {
    escape_xml(text).replace('\r', "&#13;")
}

fn generated_text(raw: &str, text: &str) -> String {
    let space = if needs_space(text) {
        " xml:space=\"preserve\""
    } else {
        ""
    };
    format!("<w:t{space}>{raw}</w:t>")
}

fn needs_space(text: &str) -> bool {
    let xml_space = |ch: char| matches!(ch, ' ' | '\t' | '\n' | '\r');
    text.starts_with(xml_space) || text.ends_with(xml_space) || text.contains("  ")
}

fn text_element(source: &str, node: &Node, raw: &str, text: &str) -> Result<String, Failure> {
    #[cfg(test)]
    super::work::add(node.open.len() + raw.len() + text.len() + node.close.len());
    if node.empty && raw.is_empty() {
        return Ok(source[node.span.clone()].to_owned());
    }
    let mut opener = source[node.open.clone()].to_owned();
    if needs_space(text) && node.attr(source, "xml:space").as_deref() != Some("preserve") {
        if let Some(attribute) = node
            .attributes
            .iter()
            .find(|attribute| attribute.name == "xml:space")
        {
            let start = attribute.value.start - node.open.start;
            let end = attribute.value.end - node.open.start;
            opener.replace_range(start..end, "preserve");
        } else {
            let at = opener.len() - if node.empty { 2 } else { 1 };
            opener.insert_str(at, " xml:space=\"preserve\"");
        }
    }
    if node.empty {
        opener = opener
            .trim_end_matches('>')
            .trim_end_matches('/')
            .to_owned()
            + ">";
    }
    let close = if node.empty {
        format!("</{}>", node.name)
    } else {
        source[node.close.clone()].to_owned()
    };
    Ok(format!("{opener}{raw}{close}"))
}

fn direct_mark_attrs(source: &str) -> Result<Attrs, Failure> {
    if source.is_empty() {
        return Ok(Attrs::new());
    }
    let limits = crate::xml::ParseLimits::default();
    let mut budget = crate::xml::ParseBudget::new(&limits);
    let document = crate::xml::parse_xml(source.as_bytes(), "", &mut budget)
        .map_err(|_| Failure::structure("invalid paragraph mark properties"))?;
    let formatting = crate::formatting::parse_run_properties(document.root(), None);
    let formatting = serde_json::to_value(formatting)
        .map_err(|_| Failure::structure("invalid paragraph mark properties"))?;
    let mut attrs = Attrs::new();
    for key in ["bold", "italic", "allCaps", "smallCaps", "hidden"] {
        if let Some(value) = formatting.get(key) {
            attrs.insert(key.to_owned(), value.clone());
        }
    }
    for (key, fields) in [
        (
            "fontSize",
            &[("size", "fontSize"), ("sizeCs", "fontSizeCs")][..],
        ),
        ("runStyle", &[("styleId", "styleId")][..]),
    ] {
        let value: BTreeMap<_, _> = fields
            .iter()
            .filter_map(|(field, source)| {
                formatting
                    .get(*source)
                    .map(|value| ((*field).to_owned(), value.clone()))
            })
            .collect();
        if !value.is_empty() {
            attrs.insert(key.to_owned(), json!(value));
        }
    }
    for (key, source) in [("fontFamily", "fontFamily"), ("underline", "underline")] {
        if let Some(value) = formatting.get(source) {
            let mut value = value.clone();
            if key == "fontFamily" {
                if let Some(object) = value.as_object_mut() {
                    object.remove("hint");
                }
            } else if value.get("style").and_then(Value::as_str) == Some("none") {
                continue;
            }
            attrs.insert(key.to_owned(), value);
        }
    }
    if let Some(color) = formatting.get("color")
        && color.get("auto") != Some(&Value::Bool(true))
    {
        attrs.insert("textColor".to_owned(), color.clone());
    }
    let shading_highlight = formatting
        .get("shading")
        .filter(|shading| {
            shading
                .get("pattern")
                .is_none_or(|value| value.as_str() == Some("clear"))
        })
        .and_then(|shading| shading.get("fill"))
        .filter(|fill| fill.get("auto") != Some(&Value::Bool(true)))
        .and_then(|fill| fill.get("rgb"))
        .and_then(Value::as_str)
        .map(|rgb| json!(format!("#{}", rgb.trim_start_matches('#'))));
    if let Some(highlight) = formatting
        .get("highlight")
        .filter(|value| value.as_str() != Some("none"))
        .cloned()
        .or(shading_highlight)
    {
        attrs.insert("highlight".to_owned(), json!({"color": highlight}));
    }
    if let Some(alignment) = formatting.get("vertAlign").and_then(Value::as_str)
        && matches!(alignment, "superscript" | "subscript")
    {
        attrs.insert(alignment.to_owned(), json!(true));
    }
    if formatting.get("strike") == Some(&Value::Bool(true))
        || formatting.get("doubleStrike") == Some(&Value::Bool(true))
    {
        attrs.insert(
            "strike".to_owned(),
            json!({
                "double": formatting.get("doubleStrike") == Some(&Value::Bool(true))
            }),
        );
    }
    Ok(attrs)
}

#[cfg(test)]
mod tests {
    use super::super::PatchContext;
    use super::*;
    use crate::paragraph_patch::{
        AttrDeltaV1, InheritanceV1, InsertedTextV1, RunEditV1, SeedV1, SlotV1, result_sha256,
        unit_sha256,
    };

    fn text(value: &str, marks: Value) -> UnitV1 {
        UnitV1::Text {
            text: value.to_owned(),
            attrs: serde_json::from_value(marks).unwrap(),
        }
    }

    fn run(id: &str, body: &str, properties: &str) -> String {
        format!(
            "<w:r w:rsidR=\"{id}\" w:rsidRPr=\"{id}\">{properties}<w:t xml:space=\"preserve\">{body}</w:t></w:r>"
        )
    }

    fn replace(source: &str, seed: Vec<UnitV1>, edits: &[(u32, u32, &str, Value)]) -> String {
        replace_with_run_edits(source, seed, edits, Vec::new())
    }

    fn replace_with_run_edits(
        source: &str,
        seed: Vec<UnitV1>,
        edits: &[(u32, u32, &str, Value)],
        run_edits: Vec<RunEditV1>,
    ) -> String {
        let patch = ParagraphPatchV1 {
            ordinal: 0,
            session_key: "source:0".to_owned(),
            seed: SeedV1 {
                unit_sha256: unit_sha256(&seed),
                units: seed.clone(),
                properties: Attrs::new(),
                inherit: InheritanceV1::default(),
            },
            text_edits: edits
                .iter()
                .map(|(from, to, value, marks)| TextEditV1 {
                    from: *from,
                    to: *to,
                    insert: vec![InsertedTextV1 {
                        text: (*value).to_owned(),
                        attrs: serde_json::from_value(marks.clone()).unwrap(),
                        inherit: Attrs::new(),
                    }],
                    donor: DonorV1::At { at: *from },
                })
                .collect(),
            run_edits,
            para_deltas: Vec::new(),
            para_inherit: Attrs::new(),
            mark_inherit: Attrs::new(),
            result_sha256: result_sha256(&seed, &Attrs::new()),
        };
        let mut lowered =
            match super::super::lower::lower(source, &PatchContext::word_document(), &seed) {
                Ok(lowered) => lowered,
                Err(_) => panic!("lowering failed"),
            };
        for (cell, unit) in lowered.cells.iter_mut().zip(flatten(&seed)) {
            cell.unit = unit;
        }
        match write(source, &lowered, &patch) {
            Ok(written) => written.0,
            Err(_) => panic!("writing failed"),
        }
    }

    #[test]
    fn retains_language_after_clearing_bold() {
        preserves_source_properties_after_clearing_bold(
            "<w:lang w:val=\"en-US\"/>",
            "<w:lang w:val=\"de-DE\"/>",
        );
    }

    #[test]
    fn retains_unmodeled_properties_after_clearing_bold() {
        for property in ["<w:noProof/>", "<x:kept x:value=\"1\"/>"] {
            preserves_source_properties_after_clearing_bold("", property);
        }
    }

    fn preserves_source_properties_after_clearing_bold(first: &str, second: &str) {
        let source = format!(
            "<w:p xmlns:x=\"urn:example\">{}{}</w:p>",
            run("1", "abc ", &format!("<w:rPr>{first}</w:rPr>")),
            run("2", "def", &format!("<w:rPr><w:b/>{second}</w:rPr>"))
        );
        let output = replace_with_run_edits(
            &source,
            vec![text("abc ", json!({})), text("def", json!({"bold": true}))],
            &[(0, 1, "A", json!({})), (4, 5, "D", json!({}))],
            vec![RunEditV1 {
                from: 5,
                to: 7,
                deltas: vec![AttrDeltaV1 {
                    path: vec!["bold".to_owned()],
                    before: SlotV1::Present { value: json!(true) },
                    after: SlotV1::Missing,
                }],
                target: Attrs::new(),
                inherit: Attrs::new(),
            }],
        );
        assert_eq!(output.matches("<w:r ").count(), 2, "{output}");
        assert!(output.contains("Abc "), "{output}");
        assert!(output.contains("Def"), "{output}");
        assert!(output.contains(&format!("<w:rPr>{first}</w:rPr>")));
        assert!(output.contains(&format!("<w:rPr>{second}</w:rPr>")));
    }

    #[test]
    fn coalesces_changed_runs_and_preserves_first_properties_and_rsids() {
        let properties = "<w:rPr><w:b/><w:lang w:val=\"de-DE\"/></w:rPr>";
        let source = format!(
            "<w:p>{} \n{}</w:p>",
            run("00000001", "Tango uniform ", properties),
            run("00000002", "victor whiskey.", properties)
        );
        let output = replace(
            &source,
            vec![text("Tango uniform victor whiskey.", json!({"bold": true}))],
            &[(6, 20, "UNIFORM VICTOR", json!({"bold": true}))],
        );
        assert_eq!(output.matches("<w:r ").count(), 1);
        assert!(output.contains("w:rsidR=\"00000001\""));
        assert!(!output.contains("00000002"));
        assert!(output.contains(properties));
        assert!(output.contains("Tango UNIFORM VICTOR"));
        assert!(output.contains(" whiskey."));
    }

    #[test]
    fn coalesces_runs_changed_by_separate_text_edits() {
        let source = format!(
            "<w:p>{}{}</w:p>",
            run("1", "alpha ", ""),
            run("2", "beta", "")
        );
        let output = replace(
            &source,
            vec![text("alpha beta", json!({}))],
            &[(0, 5, "ALPHA", json!({})), (6, 10, "BETA", json!({}))],
        );
        assert_eq!(output.matches("<w:r ").count(), 1);
        assert!(output.contains("ALPHA "));
        assert!(output.contains("BETA"));
    }

    #[test]
    fn retains_untouched_source_splits() {
        let second = run("2", "second part.", "<w:rPr><w:b/></w:rPr>");
        let source = format!(
            "<w:p>{}{second}</w:p>",
            run("1", "Guard first ", "<w:rPr><w:b/></w:rPr>")
        );
        let output = replace(
            &source,
            vec![text("Guard first second part.", json!({"bold": true}))],
            &[(6, 11, "FIRST", json!({"bold": true}))],
        );
        assert_eq!(output, source.replace("first", "FIRST"));
        assert!(output.contains(&second));
    }

    #[test]
    fn retains_different_properties_without_splitting_the_donor() {
        let source = format!(
            "<w:p>{}{}</w:p>",
            run("1", "Xray yankee ", "<w:rPr><w:b/></w:rPr>"),
            run("2", "zulu alpha.", "<w:rPr><w:i/></w:rPr>")
        );
        let output = replace(
            &source,
            vec![
                text("Xray yankee ", json!({"bold": true})),
                text("zulu alpha.", json!({"italic": true})),
            ],
            &[(5, 16, "YANKEE ZULU", json!({"bold": true}))],
        );
        assert_eq!(output.matches("<w:r ").count(), 2);
        assert!(output.contains("Xray YANKEE ZULU"));
        assert!(output.contains("<w:rPr><w:i/></w:rPr>"));
    }

    #[test]
    fn split_text_edits_keep_each_runs_complex_script_properties() {
        let bold = "<w:rPr><w:b/><w:bCs/></w:rPr>";
        let italic = "<w:rPr><w:i/><w:iCs/></w:rPr>";
        let source = format!(
            "<w:p>{}{}</w:p>",
            run("1", "Xray yankee ", bold),
            run("2", "zulu alpha.", italic)
        );
        let output = replace(
            &source,
            vec![
                text("Xray yankee ", json!({"bold": true})),
                text("zulu alpha.", json!({"italic": true})),
            ],
            &[
                (5, 11, "YANKEE", json!({"bold": true})),
                (12, 16, "ZULU", json!({"bold": true})),
            ],
        );
        assert_eq!(output.matches("<w:r ").count(), 3);
        assert!(output.contains("<w:rPr><w:b/><w:iCs/></w:rPr>"));
        assert_eq!(output.matches(bold).count(), 1);
        assert_eq!(output.matches(italic).count(), 1);
        assert!(output.contains("Xray YANKEE"));
        assert!(output.contains("ZULU"));
    }

    #[test]
    fn retains_proof_error_boundaries() {
        let marker = "<w:proofErr w:type=\"spellStart\"/>";
        let source = format!(
            "<w:p>{}{marker}{}</w:p>",
            run("1", "alpha ", ""),
            run("2", "beta", "")
        );
        let output = replace(
            &source,
            vec![text("alpha beta", json!({}))],
            &[(0, 5, "ALPHA", json!({})), (6, 10, "BETA", json!({}))],
        );
        assert_eq!(output.matches("<w:r ").count(), 2);
        assert!(output.contains(marker));
    }

    #[test]
    fn retains_wrapper_boundaries_and_non_text_run_children() {
        for (before, middle, after) in [
            ("<w:hyperlink w:anchor=\"kept\">", "</w:hyperlink>", ""),
            ("<w:smartTag>", "</w:smartTag><w:smartTag>", "</w:smartTag>"),
            ("", "<w:r><w:tab/></w:r>", ""),
        ] {
            let source = format!(
                "<w:p>{before}{}{middle}{}{after}</w:p>",
                run("1", "alpha ", ""),
                run("2", "beta", "")
            );
            let offset = if middle.contains("w:tab") { 7 } else { 6 };
            let seed = if offset == 7 {
                "alpha \tbeta"
            } else {
                "alpha beta"
            };
            let output = replace(
                &source,
                vec![text(seed, json!({}))],
                &[
                    (0, 5, "ALPHA", json!({})),
                    (offset, offset + 4, "BETA", json!({})),
                ],
            );
            assert_eq!(output.matches("<w:r ").count(), 2);
            assert!(output.contains(middle));
        }
    }

    #[test]
    fn compares_property_bytes_and_preserves_unknown_properties() {
        let properties = "<w:rPr><w:b/><w:lang w:val=\"de-DE\"/><x:kept x:value=\"1\"/></w:rPr>";
        for second_properties in [
            properties.to_owned(),
            properties.replace("<w:b/>", "<w:b />"),
        ] {
            let source = format!(
                "<w:p xmlns:x=\"urn:example\">{}{}</w:p>",
                run("1", "alpha ", properties),
                run("2", "beta", &second_properties)
            );
            let output = replace(
                &source,
                vec![text("alpha beta", json!({"bold": true}))],
                &[
                    (0, 5, "ALPHA", json!({"bold": true})),
                    (6, 10, "BETA", json!({"bold": true})),
                ],
            );
            let expected = if second_properties == properties {
                1
            } else {
                2
            };
            assert_eq!(output.matches("<w:r ").count(), expected);
            assert!(output.contains(properties));
            assert!(output.contains("<x:kept x:value=\"1\"/>"));
        }
    }

    #[test]
    fn retains_non_rsid_attributes_and_text_identifiers() {
        for attribute in [
            "w:foo=\"second\"",
            "xml:id=\"second\"",
            "x:rsidR=\"second\"",
        ] {
            let second = run("2", "beta", "").replace("<w:r ", &format!("<w:r {attribute} "));
            let source = format!(
                "<w:p xmlns:x=\"urn:example\">{}{second}</w:p>",
                run("1", "alpha ", "")
            );
            let output = replace(
                &source,
                vec![text("alpha beta", json!({}))],
                &[(0, 5, "ALPHA", json!({})), (6, 10, "BETA", json!({}))],
            );
            assert_eq!(output.matches("<w:r ").count(), 2);
            assert!(output.contains(attribute));
        }
    }

    #[test]
    fn runs_rebinding_a_prefix_keep_their_attributes() {
        let rebound = |id: &str, body: &str| {
            run(id, body, "").replace(
                "<w:r ",
                &format!("<w:r xmlns:x=\"urn:custom\" x:rsidR=\"{id}\" "),
            )
        };
        let source = format!(
            "<w:p xmlns:x=\"{WORD}\">{}{}</w:p>",
            rebound("left", "alpha "),
            rebound("right", "beta")
        );
        let output = replace(
            &source,
            vec![text("alpha beta", json!({}))],
            &[(0, 5, "ALPHA", json!({})), (6, 10, "BETA", json!({}))],
        );
        assert_eq!(output.matches("<w:r ").count(), 2, "{output}");
        assert!(output.contains("x:rsidR=\"left\""));
        assert!(output.contains("x:rsidR=\"right\""));
    }

    #[test]
    fn namespace_declarations_inside_a_paragraph_block_merging() {
        for (declaration, expected) in [("", 1), (" xmlns:y=\"urn:custom\"", 2)] {
            let source = format!(
                "<w:p><w:hyperlink w:anchor=\"kept\"{declaration}>{}{}</w:hyperlink></w:p>",
                run("1", "alpha ", ""),
                run("2", "beta", "")
            );
            let output = replace(
                &source,
                vec![text("alpha beta", json!({}))],
                &[(0, 5, "ALPHA", json!({})), (6, 10, "BETA", json!({}))],
            );
            assert_eq!(output.matches("<w:r ").count(), expected, "{output}");
        }
    }

    #[test]
    fn retains_plain_text_bytes_and_multiple_text_elements() {
        let source = "<w:p><w:r w:rsidR=\"1\"><w:rPr /><w:t xml:space=\"preserve\">a &amp; </w:t><w:t xml:id=\"kept\">beta</w:t></w:r><w:r w:rsidR=\"2\"><w:rPr /><w:t>gamma</w:t></w:r></w:p>";
        let output = replace(
            source,
            vec![text("a & betagamma", json!({}))],
            &[(4, 8, "BETA", json!({})), (8, 13, "GAMMA", json!({}))],
        );
        assert_eq!(output.matches("<w:r ").count(), 2);
        assert!(output.contains("a &amp; "));
        assert!(output.contains("xml:id=\"kept\""));
    }
}
