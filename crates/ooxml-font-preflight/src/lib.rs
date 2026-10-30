mod collect;
mod map;
mod properties;
mod xml;
mod zip;

use std::collections::BTreeSet;
use std::fmt;

#[cfg(feature = "wasm")]
use wasm_bindgen::prelude::*;

use collect::{Collector, Scan};
use map::Map;
use properties::{Props, Styles, Theme};
use xml::Node;
use zip::Zip;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PreflightFace {
    pub family: String,
    pub bold: bool,
    pub italic: bool,
    pub scripts: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FontPreflight {
    pub faces: Vec<PreflightFace>,
    pub complete: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PreflightError {
    Zip,
    Xml,
    MissingDocument,
    Limit,
}

impl fmt::Display for PreflightError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Zip => "Invalid or unsupported DOCX ZIP",
            Self::Xml => "Invalid or unsupported DOCX XML",
            Self::MissingDocument => "Missing word/document.xml",
            Self::Limit => "DOCX preflight limit exceeded",
        })
    }
}

impl std::error::Error for PreflightError {}

impl FontPreflight {
    pub fn to_json(&self) -> String {
        let mut output = String::from("{\"faces\":[");
        for (index, face) in self.faces.iter().enumerate() {
            if index > 0 {
                output.push(',');
            }
            output.push_str("{\"family\":");
            json_string(&mut output, &face.family);
            output.push_str(if face.bold {
                ",\"bold\":true"
            } else {
                ",\"bold\":false"
            });
            output.push_str(if face.italic {
                ",\"italic\":true,\"scripts\":["
            } else {
                ",\"italic\":false,\"scripts\":["
            });
            for (index, script) in face.scripts.iter().enumerate() {
                if index > 0 {
                    output.push(',');
                }
                json_string(&mut output, script);
            }
            output.push_str("]}");
        }
        output.push_str(if self.complete {
            "],\"complete\":true}"
        } else {
            "],\"complete\":false}"
        });
        output
    }
}

fn json_string(output: &mut String, value: &str) {
    output.push('"');
    for ch in value.chars() {
        match ch {
            '"' => output.push_str("\\\""),
            '\\' => output.push_str("\\\\"),
            '\u{00}'..='\u{1f}' => {
                let byte = ch as u8;
                let hex = b"0123456789abcdef";
                output.push_str("\\u00");
                output.push(char::from(hex[usize::from(byte >> 4)]));
                output.push(char::from(hex[usize::from(byte & 15)]));
            }
            _ => output.push(ch),
        }
    }
    output.push('"');
}

struct Parts<'a> {
    zip: Zip<'a>,
    nodes: usize,
}

impl Parts<'_> {
    fn part(&mut self, name: &str, complete: &mut bool) -> Result<Option<Node>, PreflightError> {
        let Some(bytes) = self.zip.read(name)? else {
            return Ok(None);
        };
        let node = xml::parse(&bytes, &mut self.nodes)?;
        *complete &= !node.incomplete;
        Ok(Some(node))
    }
}

pub(crate) const MAX_RELATIONSHIPS: usize = 4096;
const MAX_REFERENCES: usize = 4096;

struct Relationship {
    id: String,
    kind: String,
    target: String,
}

fn relationships(root: Option<&Node>, complete: &mut bool) -> Vec<Relationship> {
    let mut result = Vec::new();
    let mut ids = BTreeSet::new();
    for (index, node) in root.into_iter().flat_map(|node| &node.children).enumerate() {
        if index == MAX_RELATIONSHIPS {
            *complete = false;
            break;
        }
        if node.name != "Relationship" || node.attr("TargetMode") == Some("External") {
            continue;
        }
        let (Some(id), Some(kind), Some(target)) =
            (node.attr("Id"), node.attr("Type"), node.attr("Target"))
        else {
            *complete = false;
            continue;
        };
        if !ids.insert(id) {
            *complete = false;
            continue;
        }
        let mut segments = if target.starts_with('/') {
            Vec::new()
        } else {
            vec!["word"]
        };
        let mut valid = true;
        for segment in target.split('/') {
            match segment {
                "" | "." => {}
                ".." => {
                    valid &= segments.pop().is_some();
                }
                _ => segments.push(segment),
            }
        }
        if !valid || target.contains(['\\', ':', '%', '#', '?']) {
            *complete = false;
            continue;
        }
        result.push(Relationship {
            id: id.to_owned(),
            kind: kind.rsplit('/').next().unwrap_or_default().to_owned(),
            target: segments.join("/"),
        });
    }
    result
}

fn reference(set: &mut BTreeSet<String>, key: String, complete: &mut bool) {
    if set.len() < MAX_REFERENCES || set.contains(&key) {
        set.insert(key);
    } else {
        *complete = false;
    }
}

fn references(
    node: &Node,
    header_footer: &mut BTreeSet<String>,
    notes: &mut BTreeSet<String>,
    complete: &mut bool,
) {
    if matches!(node.name, "w:headerReference" | "w:footerReference")
        && let Some(id) = node.attr("id")
    {
        reference(header_footer, id.to_owned(), complete);
    }
    if matches!(node.name, "w:footnoteReference" | "w:endnoteReference")
        && let Some(id) = node.attr("id")
    {
        let mut key = node.name.to_owned();
        key.push(':');
        key.push_str(id);
        reference(notes, key, complete);
    }
    for child in &node.children {
        references(child, header_footer, notes, complete);
    }
}

fn uses_numbering(node: &Node, styles: &Styles, complete: &mut bool) -> bool {
    if node.name == "w:p" {
        let ppr = node.child("w:pPr");
        let style = styles.get(
            ppr.and_then(|node| node.val("w:pStyle")),
            "paragraph",
            complete,
        );
        if properties::num_pr(ppr)
            .or_else(|| style.and_then(|style| style.num.clone()))
            .is_some_and(|(id, _)| id.is_some_and(|id| id != "0"))
        {
            return true;
        }
    }
    node.children
        .iter()
        .any(|node| uses_numbering(node, styles, complete))
}

pub fn docx_font_preflight(bytes: &[u8]) -> Result<FontPreflight, PreflightError> {
    let mut parts = Parts {
        zip: Zip::new(bytes)?,
        nodes: xml::MAX_NODES,
    };
    let mut complete = true;
    let document = parts
        .part("word/document.xml", &mut complete)?
        .ok_or(PreflightError::MissingDocument)?;
    if document.name != "w:document" || document.child("w:body").is_none() {
        return Err(PreflightError::Xml);
    }
    let rels = parts.part("word/_rels/document.xml.rels", &mut complete)?;
    let rels = relationships(rels.as_ref(), &mut complete);
    let target = |kind, default: &str| {
        rels.iter()
            .find(|rel| rel.kind == kind)
            .map(|rel| rel.target.clone())
            .unwrap_or_else(|| default.to_owned())
    };
    let theme_path = target("theme", "word/theme/theme1.xml");
    if theme_path != "word/theme/theme1.xml" {
        complete = false;
    }
    let theme_xml = parts.part(&theme_path, &mut complete)?;
    if theme_xml.is_none() && rels.iter().any(|rel| rel.kind == "theme") {
        complete = false;
    }
    let settings_path = target("settings", "word/settings.xml");
    if settings_path != "word/settings.xml" {
        complete = false;
    }
    let settings = parts.part(&settings_path, &mut complete)?;
    let theme = Theme::parse(theme_xml.as_ref(), settings.as_ref());
    let styles_path = target("styles", "word/styles.xml");
    if styles_path != "word/styles.xml" {
        complete = false;
    }
    let styles_xml = parts.part(&styles_path, &mut complete)?;
    let styles = Styles::parse(styles_xml.as_ref(), &theme, &mut complete);
    let mut related = Vec::new();
    let mut header_footer = BTreeSet::new();
    let mut notes = BTreeSet::new();
    references(&document, &mut header_footer, &mut notes, &mut complete);
    let mut visited = BTreeSet::new();
    let mut loaded = BTreeSet::new();
    for rel in &rels {
        if (matches!(rel.kind.as_str(), "header" | "footer") && header_footer.contains(&rel.id))
            || matches!(rel.kind.as_str(), "footnotes" | "endnotes")
        {
            if !visited.insert(rel.target.clone()) {
                continue;
            }
            if let Some(root) = parts.part(&rel.target, &mut complete)? {
                loaded.insert(rel.target.clone());
                related.push((rel, root));
            } else {
                complete = false;
            }
        }
    }
    let numbering = if uses_numbering(&document, &styles, &mut complete)
        || related
            .iter()
            .any(|(_, node)| uses_numbering(node, &styles, &mut complete))
    {
        let numbering_path = target("numbering", "word/numbering.xml");
        if numbering_path != "word/numbering.xml" {
            complete = false;
        }
        parts.part(&numbering_path, &mut complete)?
    } else {
        None
    };
    let mut collector = Collector::default();
    let mut stories = Map::new();
    stories.insert("body".to_owned(), &document);
    for kind in ["footnotes", "endnotes"] {
        if let Some((_, root)) = related.iter().find(|(rel, _)| rel.kind == kind)
            && root.children.iter().any(normal_note)
            && let Some(separator) = root
                .children
                .iter()
                .find(|node| node.attr("type") == Some("separator"))
        {
            Scan {
                styles: &styles,
                theme: &theme,
                numbering: numbering.as_ref(),
                complete: &mut complete,
                collector: &mut collector,
                note_label: false,
                shape: false,
            }
            .walk(separator, &Props::default());
        }
    }
    let resolved_headers: BTreeSet<_> = rels
        .iter()
        .filter(|rel| {
            matches!(rel.kind.as_str(), "header" | "footer") && loaded.contains(&rel.target)
        })
        .map(|rel| rel.id.clone())
        .collect();
    let mut resolved_notes = BTreeSet::new();
    let mut story_count = 1;
    for (rel, root) in &related {
        let is_notes = matches!(rel.kind.as_str(), "footnotes" | "endnotes");
        if is_notes {
            if rel.target
                != if rel.kind == "footnotes" {
                    "word/footnotes.xml"
                } else {
                    "word/endnotes.xml"
                }
            {
                complete = false;
            }
            let prefix = if rel.kind == "footnotes" {
                "fn:"
            } else {
                "en:"
            };
            for note in &root.children {
                let id = note.attr("id").unwrap_or_default();
                if normal_note(note) {
                    if story_count == MAX_REFERENCES {
                        complete = false;
                        break;
                    }
                    story_count += 1;
                    let reference = if rel.kind == "footnotes" {
                        "w:footnoteReference:"
                    } else {
                        "w:endnoteReference:"
                    };
                    resolved_notes.insert(join(reference, id));
                    stories.insert(join(prefix, id), note);
                }
            }
        } else {
            if story_count == MAX_REFERENCES {
                complete = false;
                break;
            }
            story_count += 1;
            stories.insert(join("hf:", &rel.id), root);
        }
    }
    for id in stories.keys() {
        let Some(story) = stories.get(id) else {
            continue;
        };
        Scan {
            styles: &styles,
            theme: &theme,
            numbering: numbering.as_ref(),
            complete: &mut complete,
            collector: &mut collector,
            note_label: id.starts_with("fn:") || id.starts_with("en:"),
            shape: false,
        }
        .walk(story, &Props::default());
    }
    complete &= !collector.limited
        && header_footer.is_subset(&resolved_headers)
        && notes.is_subset(&resolved_notes);
    Ok(FontPreflight {
        faces: collector.finish(),
        complete,
    })
}

/// wasm wrapper over [`docx_font_preflight`]: DOCX bytes in, font requirements JSON out.
#[cfg(feature = "wasm")]
#[wasm_bindgen]
pub fn font_preflight(bytes: &[u8]) -> Result<String, JsValue> {
    docx_font_preflight(bytes)
        .map(|result| result.to_json())
        .map_err(|error| JsValue::from_str(&error.to_string()))
}

fn join(prefix: &str, value: &str) -> String {
    let mut result = prefix.to_owned();
    result.push_str(value);
    result
}

fn normal_note(node: &Node) -> bool {
    matches!(node.name, "w:footnote" | "w:endnote")
        && node.attr("type").is_none_or(|kind| kind == "normal")
}

#[cfg(test)]
mod tests;
