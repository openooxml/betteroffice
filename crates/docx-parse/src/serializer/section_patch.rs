use std::collections::BTreeMap;
use std::ops::Range;

use serde::{Deserialize, Serialize};
use serde_json::Value;

use super::xml_writer::{escape_xml, int_attr};
use crate::paragraph_identity::{Tag, tags, unescaped};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SectionEdits {
    #[serde(default)]
    pub only_sections: bool,
    pub sha256: String,
    pub edits: Vec<SectionEdit>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SectionEdit {
    pub paragraph: Option<usize>,
    pub donor: Option<usize>,
    pub properties: Value,
}

struct Node<'a> {
    tag: Tag<'a>,
    span: Range<usize>,
    close: usize,
    parent: Option<usize>,
    bindings: BTreeMap<String, String>,
}

fn expanded<'a>(
    name: &'a str,
    bindings: &'a BTreeMap<String, String>,
    attribute: bool,
) -> (&'a str, &'a str) {
    let (prefix, local) = name.split_once(':').unwrap_or(("", name));
    let uri = if attribute && prefix.is_empty() {
        ""
    } else {
        bindings.get(prefix).map(String::as_str).unwrap_or("")
    };
    (uri, local)
}

fn nodes<'a>(xml: &'a str, inherited: &BTreeMap<String, String>) -> Option<Vec<Node<'a>>> {
    let mut result: Vec<Node<'a>> = Vec::new();
    let mut stack: Vec<usize> = Vec::new();
    for tag in tags(xml)? {
        if tag.end {
            let node = &mut result[stack.pop()?];
            if node.tag.name != tag.name {
                return None;
            }
            node.close = tag.range.start;
            node.span.end = tag.range.end;
            continue;
        }
        let mut bindings = stack
            .last()
            .map(|index| result[*index].bindings.clone())
            .unwrap_or_else(|| inherited.clone());
        bindings
            .entry("xml".to_owned())
            .or_insert_with(|| "http://www.w3.org/XML/1998/namespace".to_owned());
        for (name, range) in &tag.attributes {
            if *name == "xmlns" {
                bindings.insert(String::new(), unescaped(xml, range.clone())?);
            } else if let Some(prefix) = name.strip_prefix("xmlns:") {
                bindings.insert(prefix.to_owned(), unescaped(xml, range.clone())?);
            }
        }
        let index = result.len();
        let empty = tag.empty;
        result.push(Node {
            span: tag.range.clone(),
            close: tag.range.end,
            tag,
            parent: stack.last().copied(),
            bindings,
        });
        if !empty {
            stack.push(index);
        }
    }
    stack.is_empty().then_some(result)
}

fn is(node: &Node<'_>, local: &str) -> bool {
    expanded(node.tag.name, &node.bindings, false) == (W, local)
}

fn children(nodes: &[Node<'_>], parent: usize, local: &str) -> Option<Option<usize>> {
    let mut found = nodes
        .iter()
        .enumerate()
        .filter(|(_, node)| node.parent == Some(parent) && is(node, local))
        .map(|(index, _)| index);
    let first = found.next();
    found.next().is_none().then_some(first)
}

fn body(nodes: &[Node<'_>]) -> Option<usize> {
    let mut found = nodes
        .iter()
        .enumerate()
        .filter(|(_, node)| {
            is(node, "body")
                && node
                    .parent
                    .is_some_and(|parent| is(&nodes[parent], "document"))
        })
        .map(|(index, _)| index);
    let first = found.next()?;
    found.next().is_none().then_some(first)
}

fn paragraph(nodes: &[Node<'_>], ordinal: usize) -> Option<usize> {
    let index = nodes
        .iter()
        .enumerate()
        .filter(|(_, node)| is(node, "p"))
        .nth(ordinal)?
        .0;
    (nodes[index].parent == Some(body(nodes)?)).then_some(index)
}

fn section(nodes: &[Node<'_>], ordinal: Option<usize>) -> Option<Option<usize>> {
    match ordinal {
        None => children(nodes, body(nodes)?, "sectPr"),
        Some(ordinal) => {
            let p = paragraph(nodes, ordinal)?;
            match children(nodes, p, "pPr")? {
                Some(pr) => children(nodes, pr, "sectPr"),
                None => Some(None),
            }
        }
    }
}

fn section_locations(nodes: &[Node<'_>]) -> Option<Vec<(Option<usize>, usize)>> {
    let body = body(nodes)?;
    let mut locations = Vec::new();
    if let Some(index) = children(nodes, body, "sectPr")? {
        locations.push((None, index));
    }
    for (ordinal, (index, node)) in nodes
        .iter()
        .enumerate()
        .filter(|(_, node)| is(node, "p"))
        .enumerate()
    {
        if node.parent != Some(body) {
            continue;
        }
        if let Some(properties) = children(nodes, index, "pPr")?
            && let Some(section) = children(nodes, properties, "sectPr")?
        {
            locations.push((Some(ordinal), section));
        }
    }
    Some(locations)
}

pub(crate) fn splice_source(source: &str) -> Option<String> {
    let parsed = nodes(source, &BTreeMap::new())?;
    let mut patches = Vec::new();
    for (_, index) in section_locations(&parsed)? {
        let section = &parsed[index];
        if parsed
            .iter()
            .any(|node| is(node, "p") && section.span.contains(&node.span.start))
        {
            return None;
        }
        patches.push((section.span.clone(), format!("<{} />", section.tag.name)));
    }
    let masked = apply(source, patches)?;
    crate::xml::reads_as_written(masked.as_bytes()).then_some(masked)
}

pub(crate) fn restore_sections(
    source: &str,
    current: &str,
    order: &[Option<u32>],
) -> Option<String> {
    let parsed = nodes(source, &BTreeMap::new())?;
    let mut edits = Vec::new();
    for (donor, _) in section_locations(&parsed)? {
        let paragraph = match donor {
            None => None,
            Some(ordinal) => match order
                .iter()
                .position(|value| *value == Some(ordinal as u32))
            {
                Some(at) => Some(at),
                None => continue,
            },
        };
        edits.push(SectionEdit {
            paragraph,
            donor,
            properties: serde_json::json!({}),
        });
    }
    patch_sections(source, current, &edits)
}

fn apply(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> Option<String> {
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output = String::new();
    let mut cursor = 0;
    for (range, replacement) in edits {
        if range.start < cursor || range.end > xml.len() {
            return None;
        }
        output.push_str(&xml[cursor..range.start]);
        output.push_str(&replacement);
        cursor = range.end;
    }
    output.push_str(&xml[cursor..]);
    Some(output)
}

fn qname(node: &Node<'_>, local: &str) -> Option<String> {
    let prefix = node
        .bindings
        .iter()
        .find(|(prefix, uri)| !prefix.is_empty() && uri.as_str() == W)?
        .0;
    Some(format!("{prefix}:{local}"))
}

fn insert_child(xml: &str, node: &Node<'_>, at: usize, child: &str) -> Option<String> {
    if node.tag.empty {
        let slash = xml[node.tag.range.clone()].rfind('/')? + node.tag.range.start;
        apply(
            xml,
            vec![(
                slash..node.tag.range.end,
                format!(">{child}</{}>", node.tag.name),
            )],
        )
    } else {
        apply(xml, vec![(at..at, child.to_owned())])
    }
}

fn rank(local: &str) -> usize {
    [
        "headerReference",
        "footerReference",
        "footnotePr",
        "endnotePr",
        "type",
        "pgSz",
        "pgMar",
        "paperSrc",
        "pgBorders",
        "lnNumType",
        "pgNumType",
        "cols",
        "formProt",
        "vAlign",
        "noEndnote",
        "titlePg",
        "textDirection",
        "bidi",
        "rtlGutter",
        "docGrid",
        "printerSettings",
        "sectPrChange",
    ]
    .iter()
    .position(|name| *name == local)
    .unwrap_or(usize::MAX)
}

fn patch_properties(
    raw: &str,
    inherited: &BTreeMap<String, String>,
    properties: &Value,
) -> Option<String> {
    let initial = nodes(raw, inherited)?;
    if !is(initial.first()?, "sectPr")
        || initial.iter().filter(|node| node.parent.is_none()).count() != 1
    {
        return None;
    }
    let mut output = raw.to_owned();
    for (key, value) in properties.as_object()? {
        let (leaf, attribute) = match key.as_str() {
            "sectionStart" => ("type", "val"),
            "orientation" => ("pgSz", "orient"),
            "pageWidth" => ("pgSz", "w"),
            "pageHeight" => ("pgSz", "h"),
            "marginTop" => ("pgMar", "top"),
            "marginBottom" => ("pgMar", "bottom"),
            "marginLeft" => ("pgMar", "left"),
            "marginRight" => ("pgMar", "right"),
            "headerDistance" => ("pgMar", "header"),
            "footerDistance" => ("pgMar", "footer"),
            "gutter" => ("pgMar", "gutter"),
            _ => return None,
        };
        let value = if let Some(value) = value.as_str() {
            if !["orientation", "sectionStart"].contains(&key.as_str()) {
                return None;
            }
            if (key == "orientation" && !["portrait", "landscape"].contains(&value))
                || (key == "sectionStart"
                    && !["nextPage", "continuous", "evenPage", "oddPage"].contains(&value))
            {
                return None;
            }
            escape_xml(value)
        } else {
            if ["orientation", "sectionStart"].contains(&key.as_str())
                || (["pageWidth", "pageHeight"].contains(&key.as_str()) && value.as_f64()? <= 0.0)
            {
                return None;
            }
            int_attr(Some(
                value
                    .as_f64()
                    .filter(|value| value.is_finite() && *value >= 0.0)?,
            ))
        };
        let parsed = nodes(&output, inherited)?;
        let root = parsed.first()?;
        if !is(root, "sectPr") {
            return None;
        }
        if let Some(index) = children(&parsed, 0, leaf)? {
            let node = &parsed[index];
            let mut attrs = node
                .tag
                .attributes
                .iter()
                .filter(|(name, _)| expanded(name, &node.bindings, true) == (W, attribute));
            let attr = attrs.next();
            if attrs.next().is_some() {
                return None;
            }
            if let Some((_, range)) = attr {
                output = apply(&output, vec![(range.clone(), value)])?;
            } else {
                let at = if node.tag.empty {
                    output[node.tag.range.clone()].rfind('/')? + node.tag.range.start
                } else {
                    node.tag.range.end - 1
                };
                output = apply(
                    &output,
                    vec![(at..at, format!(" {}=\"{value}\"", qname(node, attribute)?))],
                )?;
            }
        } else {
            let name = qname(root, leaf)?;
            let child = format!("<{name} {}=\"{value}\"/>", qname(root, attribute)?);
            let at = parsed
                .iter()
                .filter(|node| node.parent == Some(0))
                .find(|node| {
                    let (uri, local) = expanded(node.tag.name, &node.bindings, false);
                    uri == W && rank(local) > rank(leaf)
                })
                .map_or(root.close, |node| node.span.start);
            output = insert_child(&output, root, at, &child)?;
        }
    }
    Some(output)
}

pub(crate) fn patch_sections(source: &str, current: &str, edits: &[SectionEdit]) -> Option<String> {
    let original = nodes(source, &BTreeMap::new())?;
    let parsed = nodes(current, &BTreeMap::new())?;
    let mut patches = Vec::new();
    let mut targets = std::collections::BTreeSet::new();
    for edit in edits {
        if !targets.insert(edit.paragraph) {
            return None;
        }
        let donor = section(&original, edit.donor)?;
        let target = section(&parsed, edit.paragraph)?;
        let parent = match edit.paragraph {
            None => body(&parsed)?,
            Some(ordinal) => {
                let p = paragraph(&parsed, ordinal)?;
                children(&parsed, p, "pPr")?.unwrap_or(p)
            }
        };
        let template_parent = donor
            .map(|index| &original[index])
            .unwrap_or(&original[body(&original)?]);
        let raw = match donor {
            Some(index) => source[original[index].span.clone()].to_owned(),
            None => format!("<{} />", qname(template_parent, "sectPr")?),
        };
        let mut replacement = patch_properties(&raw, &template_parent.bindings, &edit.properties)?;
        let template_nodes = nodes(&replacement, &template_parent.bindings)?;
        let mut required = template_parent
            .bindings
            .iter()
            .filter(|(prefix, uri)| parsed[parent].bindings.get(*prefix) != Some(*uri))
            .map(|(prefix, uri)| (prefix.clone(), uri.clone()))
            .collect::<BTreeMap<_, _>>();
        let default = template_parent
            .bindings
            .get("")
            .cloned()
            .unwrap_or_default();
        if parsed[parent].bindings.get("").cloned().unwrap_or_default() != default {
            required.insert(String::new(), default);
        }
        if !required.is_empty() {
            let root = template_nodes.first()?;
            let at = root.tag.name.len() + 1;
            let declarations = required
                .iter()
                .filter(|(prefix, _)| {
                    !root.tag.attributes.iter().any(|(name, _)| {
                        *name
                            == if prefix.is_empty() {
                                "xmlns".to_owned()
                            } else {
                                format!("xmlns:{prefix}")
                            }
                    })
                })
                .map(|(prefix, uri)| {
                    if prefix.is_empty() {
                        format!(" xmlns=\"{}\"", escape_xml(uri))
                    } else {
                        format!(" xmlns:{prefix}=\"{}\"", escape_xml(uri))
                    }
                })
                .collect::<String>();
            replacement = apply(&replacement, vec![(at..at, declarations)])?;
        }
        if let Some(index) = target {
            patches.push((parsed[index].span.clone(), replacement));
        } else {
            let node = &parsed[parent];
            let child = if is(node, "p") {
                format!(
                    "<{}>{replacement}</{}>",
                    qname(node, "pPr")?,
                    qname(node, "pPr")?
                )
            } else {
                replacement
            };
            let at = if is(node, "p") {
                node.tag.range.end
            } else {
                children(&parsed, parent, "pPrChange")?
                    .map_or(node.close, |index| parsed[index].span.start)
            };
            if node.tag.empty {
                let slash = current[node.tag.range.clone()].rfind('/')? + node.tag.range.start;
                patches.push((
                    slash..node.tag.range.end,
                    format!(">{child}</{}>", node.tag.name),
                ));
            } else {
                patches.push((at..at, child));
            }
        }
    }
    apply(current, patches)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn document(paragraph: &str) -> String {
        format!(
            r#"<x:document xmlns:x="{W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:z="urn:opaque"><x:body>{paragraph}<x:sectPr z:flag='kept'><x:headerReference r:id='rId9' x:type="default"/><x:footerReference x:type='even' r:id="rId10" z:row='kept'/><z:unknown a='b'/><x:pgSz x:h='15840' x:w="12240" z:keep='yes'/></x:sectPr></x:body></x:document>"#
        )
    }

    #[test]
    fn section_splices_exempt_only_section_xml_from_plain_xml_validation() {
        let source = document("<x:p><x:r><x:t>Kept</x:t></x:r></x:p>")
            .replace("x:", "w:")
            .replace("xmlns:x", "xmlns:w");
        assert!(!crate::xml::reads_as_written(source.as_bytes()));
        assert!(splice_source(&source).is_some());
        let body_extension = source.replace("<w:t>Kept</w:t>", "<z:opaque/><w:t>Kept</w:t>");
        assert!(splice_source(&body_extension).is_none());
    }

    #[test]
    fn section_patch_keeps_exact_spans_and_clones_reference_rows() {
        for properties in [
            "",
            "<x:pPr/>",
            "<x:pPr z:a='b'><x:keepNext/><x:pPrChange/></x:pPr>",
        ] {
            let paragraph = format!(
                "<x:p z:tag='same'>{properties}<x:r z:r='same'><x:t>Text</x:t></x:r></x:p>"
            );
            let source = document(&paragraph);
            let edit = SectionEdit {
                paragraph: Some(0),
                donor: None,
                properties: json!({}),
            };
            let saved = patch_sections(&source, &source, &[edit]).unwrap();
            let raw = "<x:sectPr z:flag='kept'><x:headerReference r:id='rId9' x:type=\"default\"/><x:footerReference x:type='even' r:id=\"rId10\" z:row='kept'/><z:unknown a='b'/><x:pgSz x:h='15840' x:w=\"12240\" z:keep='yes'/></x:sectPr>";
            assert_eq!(saved.matches(raw).count(), 2);
            assert!(saved.contains("<x:r z:r='same'><x:t>Text</x:t></x:r>"));
            assert!(saved.ends_with(&source[source.find("</x:p>").unwrap()..]));
            if !properties.is_empty() && properties != "<x:pPr/>" {
                assert!(saved.contains(&format!("<x:keepNext/>{raw}<x:pPrChange/>")));
            }
        }
    }

    #[test]
    fn cloned_sections_keep_namespace_bindings_when_the_target_shadows_a_prefix() {
        let source = document(
            "<x:p><y:pPr xmlns:y=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x='urn:other'/><x:r><x:t>Text</x:t></x:r></x:p>",
        );
        let saved = patch_sections(
            &source,
            &source,
            &[SectionEdit {
                paragraph: Some(0),
                donor: None,
                properties: json!({}),
            }],
        )
        .unwrap();
        assert!(saved.contains(&format!("<x:sectPr xmlns:x=\"{W}\" z:flag='kept'>")));
        assert!(saved.contains("xmlns:x='urn:other'"));
        assert_eq!(
            saved
                .matches("<x:headerReference r:id='rId9' x:type=\"default\"/>")
                .count(),
            2
        );
    }

    #[test]
    fn final_section_patch_changes_only_requested_attribute_values() {
        let source = document("<x:p><x:r><x:t>Kept</x:t></x:r></x:p>");
        let saved = patch_sections(
            &source,
            &source,
            &[SectionEdit {
                paragraph: None,
                donor: None,
                properties: json!({"pageWidth":15840,"pageHeight":12240,"orientation":"landscape"}),
            }],
        )
        .unwrap();
        assert_eq!(
            saved,
            source.replace(
                "x:h='15840' x:w=\"12240\" z:keep='yes'",
                "x:h='12240' x:w=\"15840\" z:keep='yes' x:orient=\"landscape\""
            )
        );
    }

    #[test]
    fn section_patch_refuses_table_paragraphs_and_ambiguous_properties() {
        let source = document("<x:tbl><x:tr><x:tc><x:p/></x:tc></x:tr></x:tbl>");
        assert!(
            patch_sections(
                &source,
                &source,
                &[SectionEdit {
                    paragraph: Some(0),
                    donor: None,
                    properties: json!({})
                }]
            )
            .is_none()
        );
        let source = document("<x:p><x:pPr/><x:pPr/></x:p>");
        assert!(
            patch_sections(
                &source,
                &source,
                &[SectionEdit {
                    paragraph: Some(0),
                    donor: None,
                    properties: json!({})
                }]
            )
            .is_none()
        );
    }
}
