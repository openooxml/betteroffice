use std::collections::{HashMap, HashSet};

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Reader, XmlVersion};

use crate::RedactError;
use crate::mask::placeholder;
use crate::rels::attribute_local;
use crate::scrub::normalize_part_name;

const WORD_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

const UNRESOLVED_STYLE: &str = "UnresolvedStyle";

#[derive(Default)]
pub(crate) struct StyleMap {
    replacements: HashMap<String, String>,
    defined: HashMap<String, String>,
    names: HashMap<String, String>,
    latent: HashMap<String, String>,
    field_names: HashMap<String, Option<String>>,
    ambiguous: HashSet<String>,
    xlsx_cell_styles: HashMap<String, String>,
    xlsx_table_styles: HashMap<String, String>,
}

impl StyleMap {
    #[cfg(test)]
    pub(crate) fn collect(path: &str, bytes: &[u8]) -> Result<Self, RedactError> {
        Self::collect_docx_styles(&[(path, bytes)])
    }

    pub(crate) fn collect_docx(
        parts: &[(String, Vec<u8>)],
        targets: &[&str],
    ) -> Result<Self, RedactError> {
        let mut matched = Vec::new();
        for target in targets {
            for (path, bytes) in parts {
                if normalize_part_name(path) == *target {
                    matched.push((*target, bytes.as_slice()));
                }
            }
        }
        Self::collect_docx_styles(&matched)
    }

    fn collect_docx_styles(parts: &[(&str, &[u8])]) -> Result<Self, RedactError> {
        let mut styles: Vec<StyleDefinition> = Vec::new();
        let mut positions: HashMap<String, usize> = HashMap::new();
        let mut alias_membership: HashMap<usize, HashSet<String>> = HashMap::new();
        let mut latent_names: Vec<String> = Vec::new();
        for (path, bytes) in parts {
            let (part_styles, part_latent) = parse_docx_styles_part(path, bytes)?;
            for style in part_styles {
                let Some(id) = style.id.clone() else { continue };
                if let Some(&index) = positions.get(&id) {
                    let existing = &mut styles[index];
                    existing.custom |= style.custom;
                    existing.name = existing.name.take().or(style.name);
                    let aliases = alias_membership
                        .entry(index)
                        .or_insert_with(|| existing.aliases.iter().cloned().collect());
                    for alias in style.aliases {
                        if aliases.insert(alias.clone()) {
                            existing.aliases.push(alias);
                        }
                    }
                } else {
                    positions.insert(id, styles.len());
                    styles.push(style);
                }
            }
            latent_names.extend(part_latent);
        }

        let mut reserved: HashSet<String> = latent_names.iter().map(|name| key(name)).collect();
        for style in &styles {
            reserved.extend(
                style
                    .id
                    .iter()
                    .chain(&style.name)
                    .chain(&style.aliases)
                    .map(|v| key(v)),
            );
        }
        let mut index = 0;
        let mut next = || loop {
            index += 1;
            let candidate = format!("RedactedStyle{index}");
            if reserved.insert(key(&candidate)) {
                return candidate;
            }
        };

        let mut map = Self::default();
        for style in styles {
            let Some(id) = style.id else { continue };
            map.defined.insert(key(&id), id.clone());
            let name = style.name.unwrap_or_else(|| id.clone());
            if !style.custom && is_builtin_name(&name) {
                if !is_builtin_name(&id) {
                    map.replacements.insert(id.clone(), next());
                }
                for alias in &style.aliases {
                    map.note_field_name(alias, (!is_builtin_name(alias)).then_some(name.as_str()));
                }
                map.note_field_name(&name, None);
                map.note_field_name(&id, (!is_builtin_name(&id)).then_some(name.as_str()));
                map.names.insert(id, name);
                continue;
            }
            let replacement = next();
            for alias in &style.aliases {
                map.note_field_name(alias, Some(replacement.as_str()));
            }
            map.note_field_name(&name, Some(replacement.as_str()));
            map.note_field_name(&id, Some(replacement.as_str()));
            map.names.insert(id.clone(), replacement.clone());
            map.replacements.insert(id, replacement);
        }
        for latent in latent_names {
            if is_builtin_name(&latent) || map.latent.contains_key(&key(&latent)) {
                continue;
            }
            let replacement = match map.field_names.get(&key(&latent)) {
                Some(Some(existing)) => existing.clone(),
                _ => next(),
            };
            map.latent.insert(key(&latent), replacement);
        }
        Ok(map)
    }

    fn note_field_name(&mut self, name: &str, replacement: Option<&str>) {
        let name = key(name.trim());
        let replacement = replacement.map(str::to_owned);
        match self.field_names.get(&name) {
            Some(existing) if *existing != replacement => {
                self.ambiguous.insert(name);
            }
            Some(_) => {}
            None => {
                self.field_names.insert(name, replacement);
            }
        }
    }

    pub(crate) fn collect_xlsx(parts: &[(String, Vec<u8>)]) -> Result<Self, RedactError> {
        let mut seen: [Vec<String>; 2] = Default::default();
        let mut custom: [Vec<String>; 2] = Default::default();
        for (path, bytes) in parts {
            let path = normalize_part_name(path);
            let relevant = path == "xl/styles.xml"
                || [
                    "xl/tables/",
                    "xl/pivottables/",
                    "xl/slicers/",
                    "xl/timelines/",
                ]
                .iter()
                .any(|prefix| path.starts_with(prefix));
            if !relevant || !path.ends_with(".xml") {
                continue;
            }
            let mut reader = Reader::from_reader(bytes.as_slice());
            loop {
                let event = reader
                    .read_event()
                    .map_err(|error| xml_error(&path, error))?;
                let start = match &event {
                    Event::Start(start) | Event::Empty(start) => start,
                    Event::Eof => break,
                    _ => continue,
                };
                let element = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
                let attributes = unqualified_attributes(&path, &reader, start)?;
                for (attribute, value) in &attributes {
                    let Some(kind) = xlsx_style_kind(&element, attribute) else {
                        continue;
                    };
                    let builtin = match kind {
                        XlsxStyle::Cell => is_builtin_xlsx_cell_style(value),
                        XlsxStyle::Table => is_builtin_xlsx_table_style(value),
                    };
                    seen[kind as usize].push(value.clone());
                    if !builtin {
                        custom[kind as usize].push(value.clone());
                    }
                }
            }
        }
        let [seen_cells, seen_tables] = seen;
        let [custom_cells, custom_tables] = custom;
        Ok(Self {
            xlsx_cell_styles: assign_xlsx(&seen_cells, &custom_cells, "RedactedCellStyle"),
            xlsx_table_styles: assign_xlsx(&seen_tables, &custom_tables, "RedactedTableStyle"),
            ..Self::default()
        })
    }

    pub(crate) fn replacement(
        &self,
        element: &str,
        attribute: &str,
        value: &str,
        current_style_id: Option<&str>,
    ) -> Option<String> {
        let element = element.rsplit(':').next().unwrap_or(element);
        let attribute = attribute.rsplit(':').next().unwrap_or(attribute);
        if !attribute.eq_ignore_ascii_case("val") {
            if element.eq_ignore_ascii_case("style") && attribute.eq_ignore_ascii_case("styleId") {
                return self.replacements.get(value).cloned();
            }
            if element.eq_ignore_ascii_case("lsdException")
                && attribute.eq_ignore_ascii_case("name")
            {
                return self.latent.get(&key(value)).cloned();
            }
            return None;
        }
        if element.eq_ignore_ascii_case("name") || element.eq_ignore_ascii_case("aliases") {
            return current_style_id.and_then(|id| self.names.get(id).cloned());
        }
        let reference = matches_ignore_case(
            element,
            &[
                "basedOn",
                "next",
                "link",
                "pStyle",
                "rStyle",
                "tblStyle",
                "numStyleLink",
                "styleLink",
            ],
        );
        if !reference {
            return None;
        }
        if let Some(replacement) = self.replacements.get(value) {
            return Some(replacement.clone());
        }
        match self.defined.get(&key(value)) {
            Some(id) => self.replacements.get(id).cloned(),
            None => Some(UNRESOLVED_STYLE.to_owned()),
        }
    }

    pub(crate) fn xlsx_replacement(
        &self,
        element: &str,
        attribute: &str,
        value: &str,
        _builtin_id: bool,
    ) -> Option<String> {
        match xlsx_style_kind(element, attribute)? {
            XlsxStyle::Cell => self.xlsx_cell_styles.get(value).cloned(),
            XlsxStyle::Table => self.xlsx_table_styles.get(value).cloned(),
        }
    }

    /// The instruction with custom style names mapped, or `None` when it is no
    /// STYLEREF/TOC field this can read, which leaves it to the generic redaction.
    pub(crate) fn rewrite_field_instruction(
        &self,
        path: &str,
        instruction: &str,
    ) -> Result<Option<String>, RedactError> {
        let Some(tokens) = tokenize_field(instruction) else {
            return Ok(None);
        };
        let Some(keyword) = tokens.first().filter(|token| !token.quoted) else {
            return Ok(None);
        };
        let toc = keyword.value.eq_ignore_ascii_case("TOC");
        if !toc && !keyword.value.eq_ignore_ascii_case("STYLEREF") {
            return Ok(None);
        }
        let mut edits = Vec::new();
        let mut positional = false;
        let mut index = 1;
        while index < tokens.len() {
            let token = &tokens[index];
            index += 1;
            if token.quoted || !token.value.starts_with('\\') {
                if toc || positional {
                    return Ok(None);
                }
                let first = index - 1;
                let mut end = index;
                while !token.quoted
                    && end < tokens.len()
                    && !tokens[end].quoted
                    && !tokens[end].value.starts_with('\\')
                {
                    end += 1;
                }
                let name = tokens[first..end]
                    .iter()
                    .map(|token| token.value.as_str())
                    .collect::<Vec<_>>()
                    .join(" ");
                if let Some(replacement) = self.field_style(path, &name)? {
                    let replacement = if token.quoted || first + 1 < end {
                        format!("\"{replacement}\"")
                    } else {
                        replacement
                    };
                    edits.push((token.start, tokens[end - 1].end, replacement));
                }
                positional = true;
                index = end;
                continue;
            }
            let switch = token.value.to_ascii_lowercase();
            if let Some((format, argument)) = ["\\*", "\\#", "\\@"].iter().find_map(|format| {
                switch
                    .strip_prefix(*format)
                    .filter(|argument| !argument.is_empty())
                    .map(|argument| (*format, argument))
            }) {
                if format != "\\*" || !is_field_format(argument) {
                    edits.push((token.start + 2, token.end, placeholder(&token.value[2..])));
                }
                continue;
            }
            let takes_argument = match switch.as_str() {
                "\\*" | "\\#" | "\\@" => true,
                "\\o" | "\\n" | "\\l" | "\\p" | "\\d" | "\\t" if toc => {
                    tokens.get(index).is_some_and(|next| {
                        next.quoted || !next.value.starts_with('\\') || switch == "\\t"
                    })
                }
                "\\h" | "\\u" | "\\w" | "\\x" | "\\z" if toc => false,
                "\\l" | "\\n" | "\\p" | "\\r" | "\\t" | "\\w" if !toc => false,
                _ => return Ok(None),
            };
            if !takes_argument {
                continue;
            }
            let Some(argument) = tokens.get(index) else {
                return Ok(None);
            };
            index += 1;
            if toc && switch == "\\t" {
                if !argument.quoted && argument.value.starts_with('\\') {
                    return Ok(None);
                }
                if let Some(list) = self.rewrite_toc_styles(path, &argument.value)? {
                    edits.push((argument.start, argument.end, format!("\"{list}\"")));
                }
            } else if !(switch == "\\*" && is_field_format(&argument.value)
                || toc
                    && matches!(switch.as_str(), "\\o" | "\\n" | "\\l")
                    && is_heading_range(&argument.value))
            {
                edits.push((
                    argument.start,
                    argument.end,
                    format!("\"{}\"", placeholder(&argument.value)),
                ));
            }
        }
        if !toc && !positional {
            return Ok(None);
        }
        let mut output = String::with_capacity(instruction.len());
        let mut cursor = 0;
        for (start, end, replacement) in edits {
            output.push_str(&instruction[cursor..start]);
            output.push_str(&replacement);
            cursor = end;
        }
        output.push_str(&instruction[cursor..]);
        Ok(Some(output))
    }

    fn rewrite_toc_styles(&self, path: &str, list: &str) -> Result<Option<String>, RedactError> {
        let separator = if list.contains(';') && !list.contains(',') {
            ';'
        } else {
            ','
        };
        let mut changed = false;
        let mut items = Vec::new();
        for (index, item) in list.split(separator).enumerate() {
            let trimmed = item.trim();
            if index % 2 == 0 {
                if let Some(replacement) = self.field_style(path, trimmed)? {
                    changed = true;
                    items.push(replacement);
                } else {
                    items.push(item.to_owned());
                }
            } else if is_heading_level(trimmed) {
                items.push(item.to_owned());
            } else {
                return Err(RedactError::StyleReference {
                    part: path.to_owned(),
                    message: "TOC style level is invalid".to_owned(),
                });
            }
        }
        if items.len() % 2 != 0 {
            return Err(RedactError::StyleReference {
                part: path.to_owned(),
                message: "TOC style list is incomplete".to_owned(),
            });
        }
        Ok(changed.then(|| items.join(&separator.to_string())))
    }

    fn field_style(&self, path: &str, name: &str) -> Result<Option<String>, RedactError> {
        let normalized = key(name.trim());
        if self.ambiguous.contains(&normalized) {
            return Err(RedactError::StyleReference {
                part: path.to_owned(),
                message: "field style name matches more than one style".to_owned(),
            });
        }
        if let Some(replacement) = self.field_names.get(&normalized) {
            return Ok(replacement.clone());
        }
        if is_builtin_name(name) || is_heading_level(name.trim()) {
            return Ok(None);
        }
        Err(RedactError::StyleReference {
            part: path.to_owned(),
            message: "field style name does not resolve to a known style".to_owned(),
        })
    }
}

fn is_heading_level(value: &str) -> bool {
    value.len() == 1 && matches!(value.as_bytes()[0], b'1'..=b'9')
}

fn is_heading_range(value: &str) -> bool {
    match value.split_once('-') {
        Some((first, last)) => is_heading_level(first) && is_heading_level(last),
        None => is_heading_level(value),
    }
}

fn is_field_format(value: &str) -> bool {
    matches!(
        value.to_ascii_lowercase().as_str(),
        "mergeformat"
            | "charformat"
            | "arabic"
            | "alphabetic"
            | "roman"
            | "ordinal"
            | "cardtext"
            | "ordtext"
            | "dollartext"
            | "hex"
            | "caps"
            | "firstcap"
            | "upper"
            | "lower"
    )
}

struct StyleDefinition {
    id: Option<String>,
    custom: bool,
    name: Option<String>,
    aliases: Vec<String>,
}

fn parse_docx_styles_part(
    path: &str,
    bytes: &[u8],
) -> Result<(Vec<StyleDefinition>, Vec<String>), RedactError> {
    let mut reader = NsReader::from_reader(bytes);
    let mut stack = Vec::new();
    let mut styles = Vec::new();
    let mut latent = Vec::new();
    let mut current: Option<StyleDefinition> = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| xml_error(path, error))?;
        let (start, empty) = match &event {
            Event::Start(start) => (start, false),
            Event::Empty(start) => (start, true),
            Event::End(_) => {
                if stack.pop().is_some_and(|style| style)
                    && let Some(style) = current.take()
                {
                    styles.push(style);
                }
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        let word = is_word_namespace(&namespace);
        let local = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
        let is_style = word && local.eq_ignore_ascii_case("style");
        if !empty {
            stack.push(is_style);
        }
        if !word {
            continue;
        }
        if is_style {
            let style = parse_style_attributes(&reader, start, path)?;
            if empty {
                styles.push(style);
            } else {
                current = Some(style);
            }
        } else if local.eq_ignore_ascii_case("lsdException") {
            latent.extend(word_attribute(&reader, start, "name", path)?);
        } else if let Some(style) = current.as_mut() {
            if local.eq_ignore_ascii_case("name") {
                style.name = word_attribute(&reader, start, "val", path)?;
            } else if local.eq_ignore_ascii_case("aliases")
                && let Some(value) = word_attribute(&reader, start, "val", path)?
            {
                style.aliases.extend(
                    value
                        .split(',')
                        .map(str::trim)
                        .filter(|alias| !alias.is_empty())
                        .map(str::to_owned),
                );
            }
        }
    }
    Ok((styles, latent))
}

fn parse_style_attributes<R: std::io::BufRead>(
    reader: &NsReader<R>,
    start: &BytesStart<'_>,
    path: &str,
) -> Result<StyleDefinition, RedactError> {
    let mut style = StyleDefinition {
        id: None,
        custom: false,
        name: None,
        aliases: Vec::new(),
    };
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| xml_error(path, error))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !is_word_namespace(&namespace) {
            continue;
        }
        let name = String::from_utf8_lossy(local.as_ref());
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| xml_error(path, error))?;
        if name.eq_ignore_ascii_case("styleId") {
            style.id = Some(value.into_owned());
        } else if name.eq_ignore_ascii_case("customStyle") {
            style.custom = is_true(&value);
        }
    }
    Ok(style)
}

fn word_attribute<R: std::io::BufRead>(
    reader: &NsReader<R>,
    start: &BytesStart<'_>,
    expected: &str,
    path: &str,
) -> Result<Option<String>, RedactError> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| xml_error(path, error))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if is_word_namespace(&namespace)
            && String::from_utf8_lossy(local.as_ref()).eq_ignore_ascii_case(expected)
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| xml_error(path, error));
        }
    }
    Ok(None)
}

fn unqualified_attributes(
    path: &str,
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<Vec<(String, String)>, RedactError> {
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| xml_error(path, error))?;
        let name = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        if name.contains(':') {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| xml_error(path, error))?;
        attributes.push((name, value.into_owned()));
    }
    Ok(attributes)
}

pub(crate) fn has_builtin_id(attributes: &[(String, String)]) -> bool {
    attributes
        .iter()
        .any(|(name, value)| attribute_local(name) == "builtinId" && !value.trim().is_empty())
}

fn is_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(namespace) if namespace.as_ref() == WORD_TRANSITIONAL || namespace.as_ref() == WORD_STRICT)
}

fn is_true(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "1" | "true" | "on"
    )
}

#[derive(Clone, Copy)]
enum XlsxStyle {
    Cell,
    Table,
}

fn xlsx_style_kind(element: &str, attribute: &str) -> Option<XlsxStyle> {
    match (element, attribute) {
        ("cellStyle", "name") => Some(XlsxStyle::Cell),
        ("tableStyles", "defaultTableStyle" | "defaultPivotStyle")
        | ("slicerStyles", "defaultSlicerStyle")
        | ("timelineStyles", "defaultTimelineStyle")
        | (
            "tableStyle"
            | "tableStyleInfo"
            | "pivotTableStyleInfo"
            | "slicerStyle"
            | "timelineStyle",
            "name",
        )
        | ("slicer" | "timeline", "style") => Some(XlsxStyle::Table),
        _ => None,
    }
}

fn assign_xlsx(seen: &[String], custom: &[String], prefix: &str) -> HashMap<String, String> {
    let mut reserved: HashSet<String> = seen.iter().map(|name| key(name)).collect();
    let mut replacements = HashMap::new();
    let mut index = 0;
    for name in custom {
        if replacements.contains_key(name) {
            continue;
        }
        let candidate = loop {
            index += 1;
            let candidate = format!("{prefix}{index}");
            if reserved.insert(key(&candidate)) {
                break candidate;
            }
        };
        replacements.insert(name.clone(), candidate);
    }
    replacements
}

fn is_builtin_xlsx_table_style(name: &str) -> bool {
    let lower = name.trim().to_ascii_lowercase();
    lower.is_empty()
        || lower == "none"
        || [
            ("tablestylelight", 21),
            ("tablestylemedium", 28),
            ("tablestyledark", 11),
            ("pivotstylelight", 28),
            ("pivotstylemedium", 28),
            ("pivotstyledark", 28),
            ("slicerstylelight", 6),
            ("slicerstyledark", 6),
            ("slicerstyleother", 2),
            ("timeslicerstylelight", 6),
            ("timeslicerstyledark", 6),
            ("timeslicerstyleother", 2),
        ]
        .iter()
        .any(|(prefix, max)| {
            lower
                .strip_prefix(prefix)
                .and_then(|number| number.parse::<u32>().ok())
                .is_some_and(|number| (1..=*max).contains(&number))
        })
}

struct FieldToken {
    value: String,
    quoted: bool,
    start: usize,
    end: usize,
}

fn tokenize_field(instruction: &str) -> Option<Vec<FieldToken>> {
    let mut tokens = Vec::new();
    let mut chars = instruction.char_indices().peekable();
    while let Some(&(start, character)) = chars.peek() {
        if character.is_whitespace() {
            chars.next();
            continue;
        }
        let mut value = String::new();
        if character == '"' {
            chars.next();
            loop {
                match chars.next()? {
                    (_, '\\') if chars.peek().is_some_and(|(_, next)| *next == '"') => {
                        value.push('"');
                        chars.next();
                    }
                    (_, '"') => break,
                    (_, other) => value.push(other),
                }
            }
        } else {
            while let Some(&(_, next)) = chars.peek() {
                if next.is_whitespace() || next == '"' || !value.is_empty() && next == '\\' {
                    break;
                }
                value.push(next);
                chars.next();
            }
        }
        let end = chars.peek().map_or(instruction.len(), |(index, _)| *index);
        tokens.push(FieldToken {
            value,
            quoted: character == '"',
            start,
            end,
        });
    }
    Some(tokens)
}

fn key(value: &str) -> String {
    value.to_ascii_lowercase()
}

fn matches_ignore_case(value: &str, candidates: &[&str]) -> bool {
    candidates
        .iter()
        .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

fn xml_error(path: &str, error: impl std::fmt::Display) -> RedactError {
    RedactError::Xml {
        part: path.to_owned(),
        message: error.to_string(),
    }
}

fn is_builtin_name(value: &str) -> bool {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return false;
    }
    if let Some(base) = trimmed
        .strip_suffix(" Char")
        .or_else(|| trimmed.strip_suffix(" char"))
    {
        return is_builtin_base_name(base);
    }
    if let Some(base) = trimmed.strip_suffix("Char").filter(|base| !base.is_empty())
        && is_builtin_base_name(base)
    {
        return true;
    }
    is_builtin_base_name(trimmed)
}

fn is_builtin_base_name(value: &str) -> bool {
    if value.is_empty()
        || !value
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || matches!(c, ' ' | '(' | ')' | '-'))
    {
        return false;
    }
    let normalized: String = value
        .chars()
        .filter(char::is_ascii_alphanumeric)
        .map(|c| c.to_ascii_lowercase())
        .collect();
    if normalized.is_empty() {
        return false;
    }

    let without_accent = [
        "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    ]
    .iter()
    .find_map(|suffix| normalized.strip_suffix(suffix))
    .unwrap_or(normalized.as_str());

    if matches!(
        without_accent,
        "lightshading"
            | "lightlist"
            | "lightgrid"
            | "mediumshading1"
            | "mediumshading2"
            | "mediumlist1"
            | "mediumlist2"
            | "mediumgrid1"
            | "mediumgrid2"
            | "mediumgrid3"
            | "darklist"
            | "colorfulshading"
            | "colorfullist"
            | "colorfulgrid"
            | "gridtable1light"
            | "gridtable2"
            | "gridtable3"
            | "gridtable4"
            | "gridtable5"
            | "gridtable5dark"
            | "gridtable6"
            | "gridtable6colorful"
            | "gridtable7"
            | "gridtable7colorful"
            | "listtable1light"
            | "listtable2"
            | "listtable3"
            | "listtable4"
            | "listtable5"
            | "listtable5dark"
            | "listtable6"
            | "listtable6colorful"
            | "listtable7"
            | "listtable7colorful"
    ) {
        return true;
    }
    if without_accent != normalized {
        return false;
    }

    if ["heading", "toc", "index"].iter().any(|prefix| {
        normalized.strip_prefix(prefix).is_some_and(|suffix| {
            matches!(suffix, "1" | "2" | "3" | "4" | "5" | "6" | "7" | "8" | "9")
        })
    }) {
        return true;
    }
    if [
        "list",
        "listbullet",
        "listnumber",
        "listcontinue",
        "tablecolumns",
        "plaintable",
    ]
    .iter()
    .any(|prefix| {
        normalized
            .strip_prefix(prefix)
            .is_some_and(|suffix| matches!(suffix, "1" | "2" | "3" | "4" | "5"))
    }) {
        return true;
    }
    if [
        "tablesimple",
        "tablecolorful",
        "table3deffects",
        "tableweb",
        "outlinelist",
    ]
    .iter()
    .any(|prefix| {
        normalized
            .strip_prefix(prefix)
            .is_some_and(|suffix| matches!(suffix, "1" | "2" | "3"))
    }) {
        return true;
    }
    if normalized
        .strip_prefix("tableclassic")
        .is_some_and(|suffix| matches!(suffix, "1" | "2" | "3" | "4"))
    {
        return true;
    }
    if ["tablegrid", "tablelist"].iter().any(|prefix| {
        normalized
            .strip_prefix(prefix)
            .is_some_and(|suffix| matches!(suffix, "1" | "2" | "3" | "4" | "5" | "6" | "7" | "8"))
    }) {
        return true;
    }
    if normalized
        .strip_prefix("tablesubtle")
        .is_some_and(|suffix| matches!(suffix, "1" | "2"))
    {
        return true;
    }

    matches!(
        normalized.as_str(),
        "normal"
            | "normalindent"
            | "normalweb"
            | "defaultparagraphfont"
            | "tablenormal"
            | "normaltable"
            | "nolist"
            | "title"
            | "subtitle"
            | "caption"
            | "header"
            | "footer"
            | "footnotetext"
            | "footnotereference"
            | "endnotetext"
            | "endnotereference"
            | "commenttext"
            | "commentreference"
            | "commentsubject"
            | "annotationtext"
            | "annotationreference"
            | "annotationsubject"
            | "balloontext"
            | "bodytext"
            | "bodytext2"
            | "bodytext3"
            | "bodytextindent"
            | "bodytextindent2"
            | "bodytextindent3"
            | "bodytextfirstindent"
            | "bodytextfirstindent2"
            | "list"
            | "listparagraph"
            | "listbullet"
            | "listnumber"
            | "listcontinue"
            | "nospacing"
            | "quote"
            | "intensequote"
            | "emphasis"
            | "intenseemphasis"
            | "strong"
            | "subtleemphasis"
            | "subtlereference"
            | "intensereference"
            | "booktitle"
            | "bibliography"
            | "hyperlink"
            | "followedhyperlink"
            | "smartlink"
            | "smarthyperlink"
            | "mention"
            | "unresolvedmention"
            | "hashtag"
            | "tocheading"
            | "tableoffigures"
            | "tableofauthorities"
            | "toaheading"
            | "indexheading"
            | "linenumber"
            | "pagenumber"
            | "envelopeaddress"
            | "envelopereturn"
            | "macro"
            | "macrotext"
            | "closing"
            | "signature"
            | "messageheader"
            | "salutation"
            | "date"
            | "noteheading"
            | "blocktext"
            | "documentmap"
            | "plaintext"
            | "emailsignature"
            | "placeholdertext"
            | "htmltopofform"
            | "htmlbottomofform"
            | "htmlnormal"
            | "htmlcode"
            | "htmlpreformatted"
            | "htmltypewriter"
            | "htmlkeyboard"
            | "htmlsample"
            | "htmlvariable"
            | "htmlacronym"
            | "htmladdress"
            | "htmlcite"
            | "htmldefinition"
            | "tablegrid"
            | "tablegridlight"
            | "gridtablelight"
            | "tablecontemporary"
            | "tableelegant"
            | "tableprofessional"
            | "tabletheme"
    )
}

fn is_builtin_xlsx_cell_style(name: &str) -> bool {
    matches!(
        name.trim().to_ascii_lowercase().as_str(),
        "normal"
            | "comma"
            | "comma [0]"
            | "currency"
            | "currency [0]"
            | "percent"
            | "good"
            | "bad"
            | "neutral"
            | "calculation"
            | "check cell"
            | "explanatory text"
            | "input"
            | "linked cell"
            | "note"
            | "output"
            | "warning text"
            | "title"
            | "heading 1"
            | "heading 2"
            | "heading 3"
            | "heading 4"
            | "total"
            | "20% - accent1"
            | "40% - accent1"
            | "60% - accent1"
            | "accent1"
            | "20% - accent2"
            | "40% - accent2"
            | "60% - accent2"
            | "accent2"
            | "20% - accent3"
            | "40% - accent3"
            | "60% - accent3"
            | "accent3"
            | "20% - accent4"
            | "40% - accent4"
            | "60% - accent4"
            | "accent4"
            | "20% - accent5"
            | "40% - accent5"
            | "60% - accent5"
            | "accent5"
            | "20% - accent6"
            | "40% - accent6"
            | "60% - accent6"
            | "accent6"
            | "followed hyperlink"
            | "hyperlink"
    )
}

#[cfg(test)]
mod tests {
    use super::StyleMap;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    #[test]
    fn preserves_builtin_and_default_styles() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="Normal" w:default="1"/><w:style w:styleId="Custom" w:customStyle="1"/></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(map.replacement("style", "styleId", "Normal", None), None);
        assert_eq!(
            map.replacement("style", "styleId", "Custom", None),
            Some("RedactedStyle1".into())
        );
    }

    #[test]
    fn avoids_case_insensitive_collisions_and_supports_forward_references() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="redactedstyle1" w:name="taken"/><w:style w:styleId="Custom" w:customStyle="on"><w:name w:val="Project"/></w:style></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(
            map.replacement("pStyle", "val", "custom", None),
            Some("RedactedStyle3".into())
        );
        assert_eq!(
            map.replacement("pStyle", "val", "Custom", None),
            Some("RedactedStyle3".into())
        );
        assert_eq!(
            map.replacement("name", "val", "ignored", Some("Custom")),
            Some("RedactedStyle3".into())
        );
    }

    #[test]
    fn aliases_use_the_canonical_anonymous_name() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="Custom" w:customStyle="true"><w:name w:val="Project"/><w:aliases w:val="Sensitive,Other"/></w:style></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(
            map.replacement("aliases", "val", "Sensitive,Other", Some("Custom")),
            Some("RedactedStyle1".into())
        );
    }

    #[test]
    fn resolves_strict_namespaces_after_rebinding() {
        let strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        let xml = format!(
            r#"<w:styles xmlns:w="{strict}"><w:style w:styleId="Custom" w:customStyle="1"><w:name w:val="Project"/><w:basedOn w:val="Custom"/></w:style><foreign xmlns:w="urn:foreign"><w:style w:styleId="Ignored" w:customStyle="1"/></foreign></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(
            map.replacement("basedOn", "val", "Custom", None),
            Some("RedactedStyle1".into())
        );
        assert_eq!(map.replacement("style", "styleId", "Ignored", None), None);
    }

    #[test]
    fn anonymizes_unrecognized_default_style_names() {
        let xml = format!(
            r#"<w:styles xmlns:w="{W}"><w:style w:styleId="Default" w:customStyle="1" w:default="on"><w:name w:val="Default"/></w:style></w:styles>"#
        );
        let map = StyleMap::collect("styles.xml", xml.as_bytes()).unwrap();
        assert_eq!(
            map.replacement("style", "styleId", "Default", None),
            Some("RedactedStyle1".into())
        );
    }
}
