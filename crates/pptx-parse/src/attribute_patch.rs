use std::collections::{BTreeMap, HashSet};

use quick_xml::{Reader, events::Event};

use crate::xml::{XmlElement, XmlNode};

pub(crate) fn patch_attributes(
    bytes: &[u8],
    before: &XmlElement,
    after: &XmlElement,
) -> Option<Vec<u8>> {
    let mut changes = Vec::new();
    if !collect(before, after, &mut changes) {
        return None;
    }
    let mut reader = Reader::from_reader(bytes);
    let mut index = 0;
    let mut cursor = 0;
    let mut output = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().ok()? {
            Event::Start(_) | Event::Empty(_) => {
                if let Some((before, after)) = changes.get(index)? {
                    let end = reader.buffer_position() as usize;
                    output.extend_from_slice(&bytes[cursor..start]);
                    output.extend(patch_start(&bytes[start..end], before, after)?);
                    cursor = end;
                }
                index += 1;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    output.extend_from_slice(&bytes[cursor..]);
    Some(output)
}

type Attributes = BTreeMap<String, String>;
type Change<'a> = Option<(&'a Attributes, &'a Attributes)>;

fn collect<'a>(
    before: &'a XmlElement,
    after: &'a XmlElement,
    changes: &mut Vec<Change<'a>>,
) -> bool {
    if before.name != after.name || before.children.len() != after.children.len() {
        return false;
    }
    changes.push(
        (before.attributes != after.attributes).then_some((&before.attributes, &after.attributes)),
    );
    for (before, after) in before.children.iter().zip(&after.children) {
        match (before, after) {
            (XmlNode::Element(before), XmlNode::Element(after)) => {
                if !collect(before, after, changes) {
                    return false;
                }
            }
            _ if before == after => {}
            _ => return false,
        }
    }
    true
}

fn escaped(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('\r', "&#xD;")
        .replace('\n', "&#xA;")
        .replace('\t', "&#x9;")
}

fn patch_start(bytes: &[u8], before: &Attributes, after: &Attributes) -> Option<Vec<u8>> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    let mut pos = 1;
    while !bytes.get(pos)?.is_ascii_whitespace() && !b"/>".contains(bytes.get(pos)?) {
        pos += 1;
    }
    output.extend_from_slice(&bytes[..pos]);
    loop {
        let start = pos;
        while bytes.get(pos)?.is_ascii_whitespace() {
            pos += 1;
        }
        if b"/>".contains(bytes.get(pos)?) {
            for (key, value) in after {
                if !seen.contains(key.as_str()) {
                    output.extend_from_slice(format!(" {key}=\"{}\"", escaped(value)).as_bytes());
                }
            }
            output.extend_from_slice(&bytes[start..]);
            return Some(output);
        }
        let name_start = pos;
        while !bytes.get(pos)?.is_ascii_whitespace() && bytes[pos] != b'=' {
            pos += 1;
        }
        let name = std::str::from_utf8(&bytes[name_start..pos]).ok()?;
        seen.insert(name);
        while bytes.get(pos)?.is_ascii_whitespace() {
            pos += 1;
        }
        if *bytes.get(pos)? != b'=' {
            return None;
        }
        pos += 1;
        while bytes.get(pos)?.is_ascii_whitespace() {
            pos += 1;
        }
        let quote = *bytes.get(pos)?;
        if quote != b'\'' && quote != b'"' {
            return None;
        }
        pos += 1;
        let value_start = pos;
        while *bytes.get(pos)? != quote {
            pos += 1;
        }
        let value_end = pos;
        pos += 1;
        if let Some(value) = after.get(name) {
            if before.get(name) == Some(value) {
                output.extend_from_slice(&bytes[start..pos]);
            } else {
                output.extend_from_slice(&bytes[start..value_start]);
                let value = escaped(value);
                let value = if quote == b'\'' {
                    value.replace('\'', "&apos;")
                } else {
                    value
                };
                output.extend_from_slice(value.as_bytes());
                output.extend_from_slice(&bytes[value_end..pos]);
            }
        }
    }
}
