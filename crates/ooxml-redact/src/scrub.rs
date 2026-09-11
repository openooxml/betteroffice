use std::collections::{HashMap, HashSet};
use std::fmt;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};

use crate::RedactError;
use crate::rels::{self, unqualified_value};

/// Drop scrubbed parts together with the relationship parts they own, the
/// internal parts those relationships target (unless a retained part also
/// references them), their content-type declarations, and the internal
/// relationships that target any of them. Returns the scrubbed parts a
/// surviving relationship still points at: those keep their zip entry,
/// declaration and relationship, and the caller empties their bytes.
pub(crate) fn prune_scrubbed_parts(
    parts: &mut Vec<(String, Vec<u8>)>,
    scrubbed: &HashSet<String>,
) -> Result<HashSet<String>, RedactError> {
    let known: HashSet<String> = parts
        .iter()
        .map(|(path, _)| normalize_part_name(path))
        .collect();
    if has_unresolvable_reference(parts, scrubbed, &known)? {
        return Ok(scrubbed.clone());
    }
    let (mut removed, blanked) = cascade_owned_relationships(parts, scrubbed)?;
    removed.retain(|name| !blanked.contains(name));
    parts.retain(|(path, _)| !removed.contains(&normalize_part_name(path)));
    let kept_extensions: HashSet<String> = parts
        .iter()
        .filter_map(|(path, _)| extension(path))
        .collect();
    for (path, bytes) in parts.iter_mut() {
        let lower = normalize_part_name(path);
        if lower == CONTENT_TYPES {
            *bytes = prune_content_types(bytes, &removed, &kept_extensions, path)?;
        } else if lower.ends_with(".rels") {
            *bytes = prune_relationships(bytes, &removed, path)?;
        }
    }
    Ok(blanked)
}

const CONTENT_TYPES: &str = "[content_types].xml";
const ROOT_RELATIONSHIPS: &str = "_rels/.rels";

/// Whether a relationship part that can outlive scrubbing names an internal
/// target that would not survive deletion. No deletion can then be proven safe,
/// so every scrubbed part is blanked in place instead.
fn has_unresolvable_reference(
    parts: &[(String, Vec<u8>)],
    scrubbed: &HashSet<String>,
    known: &HashSet<String>,
) -> Result<bool, RedactError> {
    for (path, bytes) in parts {
        let name = normalize_part_name(path);
        if !name.ends_with(".rels") {
            continue;
        }
        match owner_of_rels_path(&name) {
            Some(owner) if scrubbed.contains(&owner) || !known.contains(&owner) => continue,
            None if name != ROOT_RELATIONSHIPS => continue,
            _ => {}
        }
        if names_an_unresolvable_part(bytes, &name, scrubbed, known)? {
            return Ok(true);
        }
    }
    Ok(false)
}

/// Whether any internal relationship here would lose its target.
fn names_an_unresolvable_part(
    bytes: &[u8],
    relationship_path: &str,
    scrubbed: &HashSet<String>,
    known: &HashSet<String>,
) -> Result<bool, RedactError> {
    let mut reader = Reader::from_reader(bytes);
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(relationship_path, error))?;
        match event {
            Event::Start(start) | Event::Empty(start)
                if start.name().local_name().as_ref() == b"Relationship" =>
            {
                let values = attributes(&reader, &start, relationship_path)?;
                if rels::external_relationship(&values, true) {
                    continue;
                }
                let Some(target) = unqualified_value(&values, "Target") else {
                    continue;
                };
                match resolve_relationship_target(relationship_path, target) {
                    Some(resolved)
                        if known.contains(&resolved)
                            && !owner_of_rels_path(&resolved)
                                .is_some_and(|owner| scrubbed.contains(&owner)) => {}
                    _ => return Ok(true),
                }
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
    }
}

/// The package's own control entries; nothing may cascade them away.
fn control_part(name: &str) -> bool {
    name == CONTENT_TYPES || name == ROOT_RELATIONSHIPS
}

/// Removes scrubbed parts plus their owned relationship parts and targets,
/// sparing targets still referenced by surviving parts. The second set holds
/// the scrubbed parts a surviving relationship still points at.
fn cascade_owned_relationships(
    parts: &[(String, Vec<u8>)],
    scrubbed: &HashSet<String>,
) -> Result<(HashSet<String>, HashSet<String>), RedactError> {
    let by_name: HashMap<String, &Vec<u8>> = parts
        .iter()
        .map(|(path, bytes)| (normalize_part_name(path), bytes))
        .collect();
    let mut spared: HashSet<String> = HashSet::new();
    let mut target_cache: HashMap<String, Vec<String>> = HashMap::new();
    loop {
        let removed = removal_closure(&by_name, scrubbed, &spared, &mut target_cache)?;
        let targets = survivor_targets(&by_name, &removed, &mut target_cache)?;
        let mut changed = false;
        for target in &targets {
            if scrubbed.contains(target)
                || !removed.contains(target)
                || !spared.insert(target.clone())
            {
                continue;
            }
            changed = true;
        }
        if !changed {
            let blanked = targets
                .into_iter()
                .filter(|target| scrubbed.contains(target))
                .collect();
            return Ok((removed, blanked));
        }
    }
}

/// Collects scrubbed parts, their owned `.rels`, and those `.rels` targets;
/// `spared` parts act as barriers.
fn removal_closure(
    by_name: &HashMap<String, &Vec<u8>>,
    scrubbed: &HashSet<String>,
    spared: &HashSet<String>,
    target_cache: &mut HashMap<String, Vec<String>>,
) -> Result<HashSet<String>, RedactError> {
    let mut removed = scrubbed.clone();
    let mut queue: Vec<String> = scrubbed.iter().cloned().collect();
    while let Some(current) = queue.pop() {
        let owned = owned_rels_path(&current);
        if !by_name.contains_key(&owned) || !removed.insert(owned.clone()) {
            continue;
        }
        queue.push(owned.clone());
        if let Some(bytes) = by_name.get(&owned) {
            let parsed = match target_cache.get_mut(&owned) {
                Some(parsed) => parsed,
                None => {
                    let parsed = internal_targets(bytes, &owned)?;
                    target_cache.insert(owned.clone(), parsed);
                    target_cache.get(&owned).unwrap()
                }
            };
            for target in parsed.clone() {
                if spared.contains(&target)
                    || control_part(&target)
                    || !by_name.contains_key(&target)
                {
                    continue;
                }
                if removed.insert(target.clone()) {
                    queue.push(target);
                }
            }
        }
    }
    Ok(removed)
}

/// Targets of surviving, owner-having `.rels` parts; orphans don't count.
fn survivor_targets(
    by_name: &HashMap<String, &Vec<u8>>,
    removed: &HashSet<String>,
    cache: &mut HashMap<String, Vec<String>>,
) -> Result<Vec<String>, RedactError> {
    let mut targets = Vec::new();
    for (name, bytes) in by_name {
        if removed.contains(name) || !name.ends_with(".rels") {
            continue;
        }
        if let Some(owner) = owner_of_rels_path(name) {
            match by_name.get(&owner) {
                Some(_) if !removed.contains(&owner) => {}
                _ => continue,
            }
        } else if name != ROOT_RELATIONSHIPS {
            continue;
        }
        let parsed = match cache.get(name) {
            Some(parsed) => parsed,
            None => {
                let parsed = internal_targets(bytes, name)?;
                cache.insert(name.clone(), parsed);
                cache.get(name).unwrap()
            }
        };
        targets.extend(parsed.iter().cloned());
    }
    Ok(targets)
}

/// The package part a `*_rels/*.rels` path belongs to, if well-formed.
fn owner_of_rels_path(path: &str) -> Option<String> {
    if let Some((directory, file)) = path.rsplit_once("/_rels/") {
        let file = file.strip_suffix(".rels")?;
        if directory.is_empty() || file.is_empty() {
            return None;
        }
        Some(format!("{directory}/{file}"))
    } else {
        path.strip_prefix("_rels/")?
            .strip_suffix(".rels")
            .filter(|file| !file.is_empty() && *file != ".")
            .map(str::to_owned)
    }
}

fn internal_targets(bytes: &[u8], relationship_path: &str) -> Result<Vec<String>, RedactError> {
    let mut reader = Reader::from_reader(bytes);
    let mut targets = Vec::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(relationship_path, error))?;
        match event {
            Event::Start(start) | Event::Empty(start)
                if start.name().local_name().as_ref() == b"Relationship" =>
            {
                let values = attributes(&reader, &start, relationship_path)?;
                if rels::external_relationship(&values, true) {
                    continue;
                }
                if let Some(target) = unqualified_value(&values, "Target")
                    .and_then(|target| resolve_relationship_target(relationship_path, target))
                {
                    targets.push(target);
                }
            }
            Event::Eof => return Ok(targets),
            _ => {}
        }
    }
}

fn owned_rels_path(part_name: &str) -> String {
    match part_name.rsplit_once('/') {
        Some((directory, name)) => format!("{directory}/_rels/{name}.rels"),
        None => format!("_rels/{part_name}.rels"),
    }
}

fn prune_content_types(
    bytes: &[u8],
    removed: &HashSet<String>,
    kept_extensions: &HashSet<String>,
    path: &str,
) -> Result<Vec<u8>, RedactError> {
    let mut reader = Reader::from_reader(bytes);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut skip_depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(path, error))?;
        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => return Err(unexpected_eof(path)),
                _ => {}
            }
            continue;
        }
        match event {
            Event::Start(start) => {
                let values = attributes(&reader, &start, path)?;
                if drops_entry(&start, &values, removed, kept_extensions) {
                    skip_depth = 1;
                } else {
                    write_start(&mut writer, start.into_owned(), false, path)?;
                }
            }
            Event::Empty(start) => {
                let values = attributes(&reader, &start, path)?;
                if !drops_entry(&start, &values, removed, kept_extensions) {
                    write_start(&mut writer, start.into_owned(), true, path)?;
                }
            }
            Event::Eof => return Ok(writer.into_inner()),
            other => write(&mut writer, other, path)?,
        }
    }
}

fn drops_entry(
    start: &BytesStart<'_>,
    values: &[(String, String)],
    removed: &HashSet<String>,
    kept_extensions: &HashSet<String>,
) -> bool {
    match start.name().local_name().as_ref() {
        b"Override" => unqualified_value(values, "PartName")
            .is_some_and(|name| removed.contains(&normalize_part_name(name))),
        b"Default" => unqualified_value(values, "Extension")
            .is_some_and(|ext| !kept_extensions.contains(&ext.to_lowercase())),
        _ => false,
    }
}

fn prune_relationships(
    bytes: &[u8],
    removed: &HashSet<String>,
    path: &str,
) -> Result<Vec<u8>, RedactError> {
    let mut reader = Reader::from_reader(bytes);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut skip_depth = 0_usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| xml_error(path, error))?;
        if skip_depth > 0 {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => return Err(unexpected_eof(path)),
                _ => {}
            }
            continue;
        }
        match event {
            Event::Start(start) => {
                let remove = start.name().local_name().as_ref() == b"Relationship"
                    && removes_target(&attributes(&reader, &start, path)?, removed, path);
                if remove {
                    skip_depth = 1;
                } else {
                    write_start(&mut writer, start.into_owned(), false, path)?;
                }
            }
            Event::Empty(start) => {
                let remove = start.name().local_name().as_ref() == b"Relationship"
                    && removes_target(&attributes(&reader, &start, path)?, removed, path);
                if !remove {
                    write_start(&mut writer, start.into_owned(), true, path)?;
                }
            }
            Event::Eof => return Ok(writer.into_inner()),
            other => write(&mut writer, other, path)?,
        }
    }
}

fn removes_target(
    values: &[(String, String)],
    removed: &HashSet<String>,
    relationship_path: &str,
) -> bool {
    if rels::external_relationship(values, true) {
        return false;
    }
    unqualified_value(values, "Target")
        .and_then(|target| resolve_relationship_target(relationship_path, target))
        .is_some_and(|target| removed.contains(&target))
}

/// Resolves a relationship target against the source part's directory,
/// handling absolute, relative, and `..` targets while preserving literal
/// percent escapes.
fn resolve_relationship_target(relationship_path: &str, target: &str) -> Option<String> {
    resolve_target_path(relationship_path, target)
}

fn resolve_target_path(relationship_path: &str, target: &str) -> Option<String> {
    if target.is_empty() || rels::external_target(target) {
        return None;
    }
    let clean_target = target
        .split(['?', '#'])
        .next()
        .unwrap_or_default()
        .replace('\\', "/");
    let relationship_path = relationship_path.replace('\\', "/");
    let source_directory = if let Some((directory, _)) = relationship_path.rsplit_once("/_rels/") {
        directory.to_owned()
    } else if relationship_path.starts_with("_rels/") {
        String::new()
    } else {
        relationship_path.clone()
    };
    let mut segments: Vec<String> =
        if clean_target.starts_with('/') || relationship_path.eq_ignore_ascii_case("_rels/.rels") {
            Vec::new()
        } else {
            source_directory
                .split('/')
                .filter(|segment| !segment.is_empty() && *segment != ".")
                .map(str::to_owned)
                .collect()
        };
    let clean_target = clean_target.trim_start_matches('/');
    for segment in clean_target
        .split('/')
        .filter(|segment| !segment.is_empty() && *segment != ".")
    {
        if segment == ".." {
            segments.pop()?;
        } else {
            segments.push(segment.to_owned());
        }
    }
    Some(segments.join("/").to_lowercase())
}

/// Canonical package-entry form: separators normalized, empty and `.` segments
/// collapsed, lowercased the way `ooxml-opc`'s `normalized_security_path` does
/// it, so entries that layer treats as one part are one part here too.
pub(crate) fn normalize_part_name(path: &str) -> String {
    path.replace('\\', "/")
        .split('/')
        .filter(|segment| !segment.is_empty() && *segment != ".")
        .collect::<Vec<_>>()
        .join("/")
        .to_lowercase()
}

fn extension(path: &str) -> Option<String> {
    path.rsplit_once('/')
        .map_or(path, |(_, name)| name)
        .rsplit_once('.')
        .map(|(_, ext)| ext.to_lowercase())
}

fn attributes(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    path: &str,
) -> Result<Vec<(String, String)>, RedactError> {
    start
        .attributes()
        .map(|attribute| {
            let attribute = attribute.map_err(|error| xml_error(path, error))?;
            let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| xml_error(path, error))?
                .into_owned();
            Ok((key, value))
        })
        .collect()
}

fn write_start(
    writer: &mut Writer<Vec<u8>>,
    start: BytesStart<'_>,
    empty: bool,
    path: &str,
) -> Result<(), RedactError> {
    if empty {
        write(writer, Event::Empty(start), path)
    } else {
        write(writer, Event::Start(start), path)
    }
}

fn write(writer: &mut Writer<Vec<u8>>, event: Event<'_>, path: &str) -> Result<(), RedactError> {
    writer
        .write_event(event)
        .map_err(|error| xml_error(path, error))
}

fn unexpected_eof(path: &str) -> RedactError {
    xml_error(path, "unexpected EOF")
}

fn xml_error(path: &str, message: impl fmt::Display) -> RedactError {
    RedactError::Xml {
        part: path.to_owned(),
        message: message.to_string(),
    }
}
