use std::collections::{BTreeMap, BTreeSet, btree_map::Entry};

use serde_json::{Value, json};

use crate::paragraph_patch::{
    Attrs, DonorV1, ParagraphPatchV1, PatchErrorCode, TextEditV1, UnitV1,
};
use crate::serializer::xml_writer::escape_xml;

use super::lower::{Cell, Lowered, Node};
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
    let boundary =
        |offset| offset == lowered.len || lowered.cells.iter().any(|cell| cell.from == offset);
    let mut retained: Vec<_> = lowered
        .cells
        .iter()
        .map(|cell| Some(cell.unit.clone()))
        .collect();
    let mut formatting: Vec<Option<String>> = vec![None; retained.len()];
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
            if cell.from >= edit.from && cell.to <= edit.to {
                let mut attrs = unit_attrs(&cell.unit).clone();
                apply_deltas(&mut attrs, &edit.deltas)?;
                if attrs != edit.target {
                    return Err(Failure::new(
                        PatchErrorCode::PropertyConflict,
                        "run target differs from deltas",
                    ));
                }
                let properties = match property_cache.entry(cell.run) {
                    Entry::Occupied(entry) => entry.into_mut(),
                    Entry::Vacant(entry) => {
                        let original = run_properties(source, lowered, cell.run);
                        entry.insert(properties::run_properties(
                            original,
                            &edit.deltas,
                            &edit.inherit,
                            &edit.target,
                        )?)
                    }
                };
                formatting[at] = Some(properties.clone());
                *super::unit_attrs_mut(retained[at].as_mut().unwrap()) = attrs;
            }
        }
    }
    let mut inserts = Vec::new();
    let mut changed_runs = BTreeSet::new();
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
            .any(|(_, cell)| !matches!(cell.unit, UnitV1::Text { .. }))
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
                    .find(|cell| cell.from <= *at && *at < cell.to);
                if !named.is_some_and(|cell| {
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
        if formatting[at].is_some() {
            changed_runs.insert(cell.run);
        }
    }
    let mut result_units = Vec::new();
    for (at, cell) in lowered.cells.iter().enumerate() {
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
        node_cells.entry(cell.node).or_default().push(at);
    }
    let cells = CellWrites {
        by_node: &node_cells,
        retained: &retained,
        formatting: &formatting,
    };
    let mut edits = Vec::new();
    for run in changed_runs {
        let replacement = write_run(source, lowered, run, &cells, &inserts)?;
        edits.push((
            lowered.xml.nodes[lowered.runs[run].node].span.clone(),
            replacement,
        ));
    }
    let mut output = apply_bytes(source, edits)?;
    if let Some(edit) = empty_insert {
        output = write_empty(source, lowered, edit)?;
    }
    Ok((output, result_units))
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
    } = writes;
    let source_run = &lowered.runs[run];
    let node = &lowered.xml.nodes[source_run.node];
    let original = run_properties(source, lowered, run);
    if source_run.from < source_run.to
        && node.children.iter().all(|&child| {
            Some(child) == source_run.properties || lowered.xml.nodes[child].word("t")
        })
        && lowered
            .cells
            .iter()
            .enumerate()
            .filter(|(_, cell)| cell.run == run)
            .all(|(at, _)| retained[at].is_none())
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
                    raw: if child_node.word("sym") && formatting[*at].is_some() {
                        symbol_bytes(source, child_node, unit_attrs(unit))?
                    } else {
                        source[cell.raw.clone()].to_owned()
                    },
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
