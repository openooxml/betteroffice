use std::collections::BTreeMap;

use serde_json::{Map, Value};

use super::lower::{Node, WORD, Xml, decode};
use super::{Failure, apply_bytes};
use crate::serializer::xml_writer::escape_xml;

fn scope(xml: &Xml, at: usize, source: &str) -> Result<BTreeMap<String, String>, Failure> {
    let mut scope = xml.root_namespaces.clone();
    let mut ancestors = vec![at];
    while let Some(parent) = xml.nodes[*ancestors.last().expect("ancestor")].parent {
        ancestors.push(parent);
    }
    for at in ancestors.into_iter().rev() {
        for attr in &xml.nodes[at].attributes {
            if let Some(prefix) = attr
                .name
                .strip_prefix("xmlns:")
                .or_else(|| (attr.name == "xmlns").then_some(""))
            {
                scope.insert(prefix.to_owned(), decode(&source[attr.value.clone()])?);
            }
        }
    }
    Ok(scope)
}

fn word_attribute(name: &str, local: &str, scope: &BTreeMap<String, String>) -> bool {
    name.split_once(':').is_some_and(|(prefix, name)| {
        name == local && scope.get(prefix).is_some_and(|ns| ns == WORD)
    })
}

fn prefix(node: &Node, scope: &BTreeMap<String, String>) -> String {
    node.name
        .split_once(':')
        .map(|(prefix, _)| prefix.to_owned())
        .or_else(|| {
            scope
                .iter()
                .find(|(prefix, ns)| !prefix.is_empty() && ns.as_str() == WORD)
                .map(|(prefix, _)| prefix.clone())
        })
        .unwrap_or_else(|| "w".to_owned())
}

fn rank(parent: &str, child: &str) -> usize {
    let order: &[&str] = match parent {
        "rPr" => super::properties::RANK_R,
        "pPr" => super::properties::RANK_P,
        "style" => &[
            "name",
            "aliases",
            "basedOn",
            "next",
            "link",
            "autoRedefine",
            "hidden",
            "uiPriority",
            "semiHidden",
            "unhideWhenUsed",
            "qFormat",
            "locked",
            "personal",
            "personalCompose",
            "personalReply",
            "rsid",
            "pPr",
            "rPr",
            "tblPr",
            "trPr",
            "tcPr",
            "tblStylePr",
        ],
        "styles" => &["docDefaults", "latentStyles", "style"],
        "docDefaults" => &["rPrDefault", "pPrDefault"],
        "rPrDefault" => &["rPr"],
        _ => &[],
    };
    order
        .iter()
        .position(|name| *name == child)
        .unwrap_or(usize::MAX)
}

fn insert_child(source: &str, xml: &Xml, child: &str, fragment: &str) -> Result<String, Failure> {
    let root = &xml.nodes[0];
    if root.empty {
        let end = root.open.end - 2;
        return apply_bytes(
            source,
            vec![(end..root.open.end, format!(">{fragment}</{}>", root.name))],
        );
    }
    let at = root
        .children
        .iter()
        .map(|at| &xml.nodes[*at])
        .find(|node| {
            (node.namespace == WORD && rank(&root.local, &node.local) > rank(&root.local, child))
                || (root.local == "rPr"
                    && node.namespace == crate::paragraph_identity::W14_NAMESPACE)
        })
        .map_or(root.close.start, |node| node.span.start);
    apply_bytes(source, vec![(at..at, fragment.to_owned())])
}

pub(super) fn run_language(
    source: &str,
    namespaces: &BTreeMap<String, String>,
    slots: &Map<String, Value>,
) -> Result<String, Failure> {
    let missing = source.is_empty();
    let created = "<w:rPr></w:rPr>";
    let source = if missing { created } else { source };
    let xml = Xml::parse(source, namespaces)?;
    let at = xml.child(0, "lang")?;
    let node = &xml.nodes[at.unwrap_or(0)];
    let scope = scope(&xml, at.unwrap_or(0), source)?;
    let prefix = prefix(node, &scope);
    let mut edits = Vec::new();
    let mut additions = String::new();
    let mut remaining = node
        .attributes
        .iter()
        .filter(|attr| attr.name != "xmlns" && !attr.name.starts_with("xmlns:"))
        .count();
    for (key, local) in [("lang", "val"), ("eastAsia", "eastAsia"), ("bidi", "bidi")] {
        let Some(value) = slots.get(key) else {
            continue;
        };
        let value = match value {
            Value::Null => None,
            Value::String(value) => Some(value.as_str()),
            _ => return Err(Failure::structure("invalid language slot")),
        };
        let matching: Vec<_> = node
            .attributes
            .iter()
            .filter(|attr| word_attribute(&attr.name, local, &scope))
            .collect();
        if matching.len() > 1 {
            return Err(Failure::structure("duplicate language slot"));
        }
        if let Some(attr) = matching.first().filter(|_| at.is_some()) {
            if let Some(value) = value {
                if decode(&source[attr.value.clone()])? != value {
                    edits.push((attr.value.clone(), escape_xml(value)));
                }
            } else {
                remaining -= 1;
                edits.push((attr.token.clone(), String::new()));
            }
        } else if let Some(value) = value {
            additions.push_str(&format!(" {prefix}:{local}=\"{}\"", escape_xml(value)));
        }
    }
    if at.is_some() {
        if additions.is_empty()
            && remaining == 0
            && node.children.is_empty()
            && (node.empty || source[node.open.end..node.close.start].trim().is_empty())
        {
            return apply_bytes(source, vec![(node.span.clone(), String::new())]);
        }
        if !additions.is_empty() {
            if scope.get(&prefix).is_none_or(|ns| ns != WORD) {
                additions.insert_str(0, &format!(" xmlns:{prefix}=\"{WORD}\""));
            }
            let at = node.open.end - if node.empty { 2 } else { 1 };
            edits.push((at..at, additions));
        }
        return apply_bytes(source, edits);
    }
    if additions.is_empty() {
        return Ok(if missing {
            String::new()
        } else {
            source.to_owned()
        });
    }
    let declaration = if scope.get(&prefix).is_none_or(|ns| ns != WORD) {
        format!(" xmlns:{prefix}=\"{WORD}\"")
    } else {
        String::new()
    };
    insert_child(
        source,
        &xml,
        "lang",
        &format!("<{prefix}:lang{declaration}{additions}/>"),
    )
}

pub(super) fn path_language(
    source: &str,
    namespaces: &BTreeMap<String, String>,
    path: &[&str],
    slots: &Map<String, Value>,
) -> Result<String, Failure> {
    if path.is_empty() {
        return run_language(source, namespaces, slots);
    }
    let xml = Xml::parse(source, namespaces)?;
    let scope = scope(&xml, 0, source)?;
    let child = path[0];
    if let Some(at) = xml.child(0, child)? {
        let replacement = path_language(
            &source[xml.nodes[at].span.clone()],
            &scope,
            &path[1..],
            slots,
        )?;
        return apply_bytes(source, vec![(xml.nodes[at].span.clone(), replacement)]);
    }
    if slots.values().all(Value::is_null) {
        return Ok(source.to_owned());
    }
    let prefix = prefix(&xml.nodes[0], &scope);
    let declaration = if scope.get(&prefix).is_none_or(|ns| ns != WORD) {
        format!(" xmlns:{prefix}=\"{WORD}\"")
    } else {
        String::new()
    };
    let created = format!("<{prefix}:{child}{declaration}></{prefix}:{child}>");
    let fragment = path_language(&created, &scope, &path[1..], slots)?;
    insert_child(source, &xml, child, &fragment)
}

/// Patches direct style and document-default proofing language slots.
pub fn patch_style_languages(source: &str, entries_json: &str) -> Result<String, String> {
    let entries: BTreeMap<String, Map<String, Value>> =
        serde_json::from_str(entries_json).map_err(|error| error.to_string())?;
    let mut output = source.to_owned();
    for (id, slots) in entries {
        if slots.is_empty() {
            continue;
        }
        if slots
            .keys()
            .any(|key| !matches!(key.as_str(), "lang" | "eastAsia" | "bidi"))
        {
            return Err("unknown language slot".to_owned());
        }
        let patch = || -> Result<String, Failure> {
            let xml = Xml::parse(&output, &BTreeMap::new())?;
            if !xml.nodes[0].word("styles") {
                return Err(Failure::structure("expected styles root"));
            }
            if id.is_empty() {
                return path_language(
                    &output,
                    &BTreeMap::new(),
                    &["docDefaults", "rPrDefault", "rPr"],
                    &slots,
                );
            }
            let mut matching = Vec::new();
            for &at in &xml.nodes[0].children {
                let node = &xml.nodes[at];
                if !node.word("style") {
                    continue;
                }
                let scope = scope(&xml, at, &output)?;
                if node.attributes.iter().any(|attr| {
                    word_attribute(&attr.name, "styleId", &scope)
                        && decode(&output[attr.value.clone()]).ok().as_deref() == Some(id.as_str())
                }) {
                    matching.push(at);
                }
            }
            if matching.len() != 1 {
                return Err(Failure::structure("unknown or duplicate style id"));
            }
            let at = matching[0];
            let replacement = path_language(
                &output[xml.nodes[at].span.clone()],
                &scope(&xml, at, &output)?,
                &["rPr"],
                &slots,
            )?;
            apply_bytes(&output, vec![(xml.nodes[at].span.clone(), replacement)])
        };
        output = patch().map_err(|error| error.detail)?;
    }
    Ok(output)
}
