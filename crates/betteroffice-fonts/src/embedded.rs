use std::borrow::Cow;
use std::collections::{HashMap, HashSet};

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::{FontDefinition, FontError};

const MAX_FONT_BYTES: usize = 32 * 1024 * 1024;
const MAX_FONT_DEFINITIONS: usize = 256;
const MAX_DEFINITION_BYTES: usize = 128 * 1024 * 1024;
const MAX_FONT_REFERENCES: usize = 4_096;
const MAX_FAMILY_BYTES: usize = 256;

fn attribute(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        (attribute.key.local_name().as_ref() == name)
            .then(|| {
                std::str::from_utf8(&attribute.value)
                    .ok()
                    .and_then(|value| quick_xml::escape::unescape(value).ok())
                    .map(|value| value.into_owned())
            })
            .flatten()
    })
}

fn target_path(owner: &str, target: &str) -> Option<String> {
    let mut segments: Vec<_> = if target.starts_with('/') {
        Vec::new()
    } else {
        owner
            .rsplit_once('/')
            .map_or("", |(dir, _)| dir)
            .split('/')
            .collect()
    };
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop()?;
            }
            segment => segments.push(segment),
        }
    }
    Some(segments.join("/"))
}

fn font_signature(bytes: &[u8]) -> bool {
    bytes.starts_with(&[0, 1, 0, 0]) || bytes.starts_with(b"OTTO") || bytes.starts_with(b"ttcf")
}

fn decode(bytes: &[u8], key: Option<String>) -> Result<Vec<u8>, FontError> {
    let invalid = || FontError::InvalidFontData("invalid embedded font".into());
    if bytes.len() > MAX_FONT_BYTES {
        return Err(FontError::InvalidFontData(format!(
            "embedded font exceeds {MAX_FONT_BYTES} bytes"
        )));
    }
    let mut bytes = Cow::Borrowed(bytes);
    if let Some(key) = key {
        let hex: String = key.chars().filter(|ch| ch.is_ascii_hexdigit()).collect();
        if hex.len() != 32 || bytes.len() < 32 {
            return Err(invalid());
        }
        let mut mask = [0; 16];
        for (index, slot) in mask.iter_mut().enumerate() {
            *slot =
                u8::from_str_radix(&hex[index * 2..index * 2 + 2], 16).map_err(|_| invalid())?;
        }
        for (index, byte) in bytes.to_mut()[..32].iter_mut().enumerate() {
            *byte ^= mask[15 - index % 16];
        }
    }
    if font_signature(&bytes) {
        return Ok(bytes.into_owned());
    }
    if bytes.len() >= 82 && bytes[34..36] == [0x4c, 0x50] {
        let length = u32::from_le_bytes(bytes[4..8].try_into().map_err(|_| invalid())?) as usize;
        let offset = bytes.len().checked_sub(length).ok_or_else(invalid)?;
        if offset >= 82 && font_signature(&bytes[offset..]) {
            return Ok(bytes[offset..].to_vec());
        }
    }
    Err(invalid())
}

#[doc(hidden)]
pub fn embedded_font_definitions<'a>(
    owner: &str,
    xml: &[u8],
    relationships: &[u8],
    part: impl Fn(&str) -> Option<&'a [u8]>,
) -> Vec<FontDefinition> {
    let mut paths: Vec<String> = Vec::new();
    let mut path_indices = HashMap::new();
    let mut targets = HashMap::new();
    let mut reader = Reader::from_reader(relationships);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"Relationship" =>
            {
                if attribute(&element, b"TargetMode").as_deref() == Some("External") {
                    continue;
                }
                if let (Some(id), Some(target)) =
                    (attribute(&element, b"Id"), attribute(&element, b"Target"))
                    && let Some(path) = target_path(owner, &target)
                {
                    let index = *path_indices.entry(path.clone()).or_insert_with(|| {
                        paths.push(path);
                        paths.len() - 1
                    });
                    targets.insert(id, index);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    let mut fonts: Vec<FontDefinition> = Vec::new();
    let mut decoded: HashMap<(usize, Option<String>), Option<usize>> = HashMap::new();
    let mut definitions = HashSet::new();
    let mut definition_bytes = 0;
    let mut references = 0;
    let mut family = None;
    let mut embedded = false;
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                let name = element.local_name();
                let name = name.as_ref();
                if name == b"embeddedFont" {
                    embedded = true;
                    family = None;
                }
                if name == b"font" {
                    family = attribute(&element, if embedded { b"typeface" } else { b"name" });
                }
                let style = match name {
                    b"embedRegular" | b"regular" => Some((false, false)),
                    b"embedBold" | b"bold" => Some((true, false)),
                    b"embedItalic" | b"italic" => Some((false, true)),
                    b"embedBoldItalic" | b"boldItalic" => Some((true, true)),
                    _ => None,
                };
                if let Some((bold, italic)) = style
                    && let Some(family) = &family
                    && family.len() <= MAX_FAMILY_BYTES
                    && let Some(id) = attribute(&element, b"id")
                    && let Some(&target) = targets.get(&id)
                {
                    references += 1;
                    if fonts.len() >= MAX_FONT_DEFINITIONS || references > MAX_FONT_REFERENCES {
                        break;
                    }
                    let font_key = attribute(&element, b"fontKey").map(|key| {
                        key.chars()
                            .filter(|ch| ch.is_ascii_hexdigit())
                            .take(33)
                            .collect::<String>()
                    });
                    let key = (target, font_key);
                    if !definitions.insert((family.clone(), bold, italic, key.clone())) {
                        continue;
                    }
                    let data = match decoded.get(&key).copied() {
                        Some(Some(index)) => {
                            let data = &fonts[index].data;
                            if data.len() > MAX_DEFINITION_BYTES - definition_bytes {
                                continue;
                            }
                            data.clone()
                        }
                        Some(None) => continue,
                        None => {
                            let data = part(&paths[target])
                                .and_then(|bytes| decode(bytes, key.1.clone()).ok());
                            let Some(data) = data else {
                                decoded.insert(key, None);
                                continue;
                            };
                            if data.len() > MAX_DEFINITION_BYTES - definition_bytes {
                                decoded.insert(key, None);
                                continue;
                            }
                            decoded.insert(key, Some(fonts.len()));
                            data
                        }
                    };
                    definition_bytes += data.len();
                    fonts.push(FontDefinition {
                        family: family.clone(),
                        data,
                        bold,
                        italic,
                    });
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"embeddedFont" => {
                embedded = false;
                family = None;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    fonts
}
