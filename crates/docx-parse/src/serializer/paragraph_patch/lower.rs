use std::collections::BTreeMap;
use std::ops::Range;

use serde_json::json;

use crate::paragraph_identity::{W14_NAMESPACE, tags};
use crate::paragraph_patch::{Attrs, PatchErrorCode, UnitV1};

use super::{Failure, PatchContext};

pub(super) const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML: &str = "http://www.w3.org/XML/1998/namespace";

#[derive(Clone)]
pub(super) struct Attribute {
    pub name: String,
    pub identifier: bool,
    pub value: Range<usize>,
    pub token: Range<usize>,
}

#[derive(Clone)]
pub(super) struct Node {
    pub name: String,
    pub namespace: String,
    pub local: String,
    pub span: Range<usize>,
    pub open: Range<usize>,
    pub close: Range<usize>,
    pub empty: bool,
    pub attributes: Vec<Attribute>,
    pub children: Vec<usize>,
}

impl Node {
    pub fn word(&self, name: &str) -> bool {
        self.namespace == WORD && self.local == name
    }

    pub fn attr(&self, source: &str, name: &str) -> Option<String> {
        self.attributes
            .iter()
            .find(|attribute| attribute.name == name)
            .and_then(|attribute| decode(&source[attribute.value.clone()]).ok())
    }
}

pub(super) struct Xml {
    pub nodes: Vec<Node>,
    pub root_namespaces: BTreeMap<String, String>,
}

impl Xml {
    pub fn parse(source: &str, namespaces: &BTreeMap<String, String>) -> Result<Self, Failure> {
        #[cfg(test)]
        super::work::add(source.len());
        let scanned = tags(source).ok_or_else(|| Failure::structure("malformed XML"))?;
        let mut nodes: Vec<Node> = Vec::new();
        let mut stack: Vec<(usize, BTreeMap<String, String>)> = Vec::new();
        let mut root_namespaces = BTreeMap::new();
        for tag in scanned {
            #[cfg(test)]
            super::work::add(1);
            if tag.end {
                let (at, _) = stack
                    .pop()
                    .ok_or_else(|| Failure::structure("unmatched end tag"))?;
                if nodes[at].name != tag.name {
                    return Err(Failure::structure("mismatched end tag"));
                }
                nodes[at].span.end = tag.range.end;
                nodes[at].close = tag.range;
                continue;
            }
            let mut scope = stack
                .last()
                .map_or_else(|| namespaces.clone(), |(_, scope)| scope.clone());
            for (name, range) in &tag.attributes {
                if let Some(prefix) = name.strip_prefix("xmlns:") {
                    scope.insert(prefix.to_owned(), decode(&source[range.clone()])?);
                } else if *name == "xmlns" {
                    scope.insert(String::new(), decode(&source[range.clone()])?);
                }
            }
            let (prefix, local) = tag.name.split_once(':').unwrap_or(("", tag.name));
            let namespace = scope.get(prefix).cloned().unwrap_or_default();
            let mut attributes = Vec::new();
            for (name, range) in tag.attributes {
                let before = &source[tag.range.start..range.start - 1];
                let offset = before
                    .rfind(name)
                    .ok_or_else(|| Failure::structure("attribute token"))?;
                let mut start = tag.range.start + offset;
                while start > tag.range.start && source.as_bytes()[start - 1].is_ascii_whitespace()
                {
                    start -= 1;
                }
                attributes.push(Attribute {
                    name: name.to_owned(),
                    identifier: is_identifier(name, &scope),
                    token: start..range.end + 1,
                    value: range,
                });
            }
            let at = nodes.len();
            if let Some((parent, _)) = stack.last() {
                nodes[*parent].children.push(at);
            } else if !nodes.is_empty() {
                return Err(Failure::structure("multiple XML roots"));
            } else {
                root_namespaces = scope.clone();
            }
            nodes.push(Node {
                name: tag.name.to_owned(),
                namespace,
                local: local.to_owned(),
                span: tag.range.clone(),
                open: tag.range.clone(),
                close: tag.range.end..tag.range.end,
                empty: tag.empty,
                attributes,
                children: Vec::new(),
            });
            if !tag.empty {
                stack.push((at, scope));
            }
        }
        if !stack.is_empty() || nodes.is_empty() {
            return Err(Failure::structure("incomplete XML"));
        }
        Ok(Self {
            nodes,
            root_namespaces,
        })
    }

    pub fn child(&self, parent: usize, name: &str) -> Result<Option<usize>, Failure> {
        let children: Vec<_> = self.nodes[parent]
            .children
            .iter()
            .copied()
            .filter(|at| {
                #[cfg(test)]
                super::work::add(1);
                self.nodes[*at].word(name)
            })
            .collect();
        if children.len() > 1 {
            return Err(Failure::new(
                PatchErrorCode::PropertyConflict,
                format!("duplicate {name}"),
            ));
        }
        Ok(children.first().copied())
    }
}

fn is_identifier(name: &str, namespaces: &BTreeMap<String, String>) -> bool {
    let Some((prefix, local)) = name.split_once(':') else {
        return false;
    };
    let namespace = namespaces
        .get(prefix)
        .map(String::as_str)
        .unwrap_or(match prefix {
            "xml" => XML,
            "w" => WORD,
            "w14" => W14_NAMESPACE,
            _ => "",
        });
    match namespace {
        XML | WORD => local == "id",
        W14_NAMESPACE => matches!(local, "paraId" | "textId"),
        _ => false,
    }
}

#[derive(Clone)]
pub(super) struct Cell {
    pub from: u32,
    pub to: u32,
    pub run: usize,
    pub node: usize,
    pub raw: Range<usize>,
    pub unit: UnitV1,
}

pub(super) struct SourceRun {
    pub node: usize,
    pub properties: Option<usize>,
    pub scope: usize,
    pub from: u32,
    pub to: u32,
}

pub(super) struct Lowered {
    pub xml: Xml,
    pub runs: Vec<SourceRun>,
    pub cells: Vec<Cell>,
    pub fences: Vec<u32>,
    pub markers: Vec<u32>,
    pub len: u32,
    pub ppr: Option<usize>,
}

pub(super) fn lower(source: &str, context: &PatchContext) -> Result<Lowered, Failure> {
    let xml = Xml::parse(source, &context.namespaces)?;
    if context.in_sdt {
        return Err(Failure::structure(
            "sdt: paragraph is inside a content control",
        ));
    }
    if !xml.nodes[0].word("p") {
        return Err(Failure::structure("paragraph is not an editable w:p"));
    }
    for node in &xml.nodes {
        #[cfg(test)]
        super::work::add(1);
        let unsupported_word = node.namespace == WORD
            && matches!(
                node.local.as_str(),
                "fldSimple"
                    | "fldChar"
                    | "instrText"
                    | "ins"
                    | "del"
                    | "moveFrom"
                    | "moveTo"
                    | "moveFromRangeStart"
                    | "moveFromRangeEnd"
                    | "moveToRangeStart"
                    | "moveToRangeEnd"
                    | "rPrChange"
                    | "pPrChange"
                    | "drawing"
                    | "pict"
                    | "txbxContent"
                    | "sdt"
            );
        if unsupported_word
            || node.namespace == "http://schemas.openxmlformats.org/officeDocument/2006/math"
        {
            let reason = match node.local.as_str() {
                "fldSimple" | "fldChar" | "instrText" => "field",
                "drawing" | "pict" | "txbxContent" => "drawing",
                "sdt" => "sdt",
                _ if node.namespace == WORD => "tc",
                _ => "math",
            };
            return Err(Failure::structure(format!(
                "{reason}: unsupported element {}",
                node.name
            )));
        }
    }
    let ppr = xml.child(0, "pPr")?;
    let mut result = Lowered {
        xml,
        runs: Vec::new(),
        cells: Vec::new(),
        fences: Vec::new(),
        markers: Vec::new(),
        len: 0,
        ppr,
    };
    visit(source, 0, 0, &mut result)?;
    Ok(result)
}

fn visit(source: &str, parent: usize, scope: usize, result: &mut Lowered) -> Result<(), Failure> {
    for at in result.xml.nodes[parent].children.clone() {
        #[cfg(test)]
        super::work::add(1);
        let node = &result.xml.nodes[at];
        if node.word("pPr") {
            continue;
        }
        if node.word("hyperlink") {
            visit(source, at, at, result)?;
        } else if node.word("smartTag") {
            visit(source, at, scope, result)?;
        } else if node.word("proofErr") || node.word("smartTagPr") {
            continue;
        } else if node.word("r") {
            lower_run(source, at, scope, result)?;
        } else if node.namespace == WORD
            && matches!(
                node.local.as_str(),
                "bookmarkStart"
                    | "bookmarkEnd"
                    | "commentRangeStart"
                    | "commentRangeEnd"
                    | "permStart"
                    | "permEnd"
            )
        {
            result.markers.push(result.len);
        } else {
            return Err(Failure::structure(format!(
                "unsupported element {}",
                node.name
            )));
        }
    }
    Ok(())
}

fn lower_run(source: &str, at: usize, scope: usize, result: &mut Lowered) -> Result<(), Failure> {
    let run = result.runs.len();
    let properties = result.xml.child(at, "rPr")?;
    let from = result.len;
    for child in result.xml.nodes[at].children.clone() {
        #[cfg(test)]
        super::work::add(1);
        let node = &result.xml.nodes[child];
        if node.word("rPr") {
            continue;
        }
        if node.namespace != WORD {
            return Err(Failure::structure(format!(
                "unsupported run child {}",
                node.name
            )));
        }
        let mut emitted = Vec::new();
        match node.local.as_str() {
            "t" => {
                if !node.children.is_empty() {
                    return Err(Failure::structure("nested text element"));
                }
                if !node.empty {
                    emitted = decoded_scalars(source, node.open.end..node.close.start)?;
                }
            }
            "tab" => emitted.push(("\t".to_owned(), node.span.clone())),
            "softHyphen" => emitted.push(("\u{ad}".to_owned(), node.span.clone())),
            "noBreakHyphen" => emitted.push(("\u{2011}".to_owned(), node.span.clone())),
            "sym" => {
                let scalar = node
                    .attr(source, "w:char")
                    .and_then(|value| u32::from_str_radix(&value, 16).ok())
                    .and_then(char::from_u32)
                    .ok_or_else(|| Failure::structure("invalid symbol scalar"))?;
                emitted.push((scalar.to_string(), node.span.clone()));
            }
            "br" | "cr" => {
                if node
                    .attr(source, "w:type")
                    .is_some_and(|kind| kind == "page" || kind == "column")
                {
                    result.fences.push(result.len);
                    continue;
                }
                push_atom(result, run, child, "break", Attrs::new());
                continue;
            }
            "footnoteReference" | "endnoteReference" => {
                let id = node
                    .attr(source, "w:id")
                    .and_then(|value| value.parse::<i64>().ok())
                    .ok_or_else(|| Failure::structure("invalid note reference"))?;
                let key = if node.word("footnoteReference") {
                    "footnoteRefId"
                } else {
                    "endnoteRefId"
                };
                let payload = BTreeMap::from([(key.to_owned(), json!(id))]);
                push_atom(result, run, child, "noteRef", payload);
                continue;
            }
            "commentReference" => {
                let mut payload = json!({
                    "fieldType": "COMMENT", "instruction": "", "displayText": "",
                    "fieldKind": "simple", "fldLock": false, "dirty": false,
                    "displayMode": "result", "hasCachedResult": false,
                    "modelKind": "commentReference"
                });
                if let Some(id) = node
                    .attr(source, "w:id")
                    .and_then(|value| value.parse::<i64>().ok())
                {
                    payload["commentId"] = json!(id);
                }
                push_atom(
                    result,
                    run,
                    child,
                    "field",
                    serde_json::from_value(payload).unwrap(),
                );
                continue;
            }
            "footnoteRef" | "endnoteRef" | "lastRenderedPageBreak" | "proofErr" => continue,
            _ => {
                return Err(Failure::structure(format!(
                    "unsupported run child {}",
                    node.name
                )));
            }
        }
        for (text, raw) in emitted {
            let width = text.encode_utf16().count() as u32;
            result.cells.push(Cell {
                from: result.len,
                to: result.len + width,
                run,
                node: child,
                raw,
                unit: UnitV1::Text {
                    text,
                    attrs: Attrs::new(),
                },
            });
            result.len += width;
        }
    }
    result.runs.push(SourceRun {
        node: at,
        properties,
        scope,
        from,
        to: result.len,
    });
    Ok(())
}

fn push_atom(result: &mut Lowered, run: usize, node: usize, kind: &str, payload: Attrs) {
    result.cells.push(Cell {
        from: result.len,
        to: result.len + 1,
        run,
        node,
        raw: result.xml.nodes[node].span.clone(),
        unit: UnitV1::Atom {
            atom_kind: kind.to_owned(),
            payload,
            attrs: Attrs::new(),
        },
    });
    result.len += 1;
}

pub(super) fn decode(text: &str) -> Result<String, Failure> {
    quick_xml::escape::unescape(text)
        .map(|value| value.into_owned())
        .map_err(|_| Failure::structure("invalid XML entity"))
}

fn decoded_scalars(
    source: &str,
    span: Range<usize>,
) -> Result<Vec<(String, Range<usize>)>, Failure> {
    let mut result: Vec<(String, Range<usize>)> = Vec::new();
    let mut cursor = span.start;
    let mut carriage_return = false;
    while cursor < span.end {
        #[cfg(test)]
        super::work::add(1);
        let end = if source.as_bytes()[cursor] == b'&' {
            cursor
                + source[cursor..span.end]
                    .find(';')
                    .ok_or_else(|| Failure::structure("unterminated entity"))?
                + 1
        } else {
            cursor + source[cursor..span.end].chars().next().unwrap().len_utf8()
        };
        let value = decode(&source[cursor..end])?;
        if value.chars().count() != 1 {
            return Err(Failure::structure("entity is not one scalar"));
        }
        if carriage_return && value == "\n" {
            result.last_mut().unwrap().1.end = end;
            carriage_return = false;
        } else {
            carriage_return = value == "\r";
            let normalized = if value == "\r" || value == "\n" {
                " ".to_owned()
            } else {
                value
            };
            result.push((normalized, cursor..end));
        }
        cursor = end;
    }
    Ok(result)
}
