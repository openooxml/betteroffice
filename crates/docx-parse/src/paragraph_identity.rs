//! Word paragraph identities (`w14:paraId`): value rules, allocation, the
//! package inventory, and identity patches applied in place to source XML.

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::ops::Range;

use quick_xml::Reader;
use quick_xml::events::Event;

pub const W14_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
pub const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Largest paragraph ID Word accepts; the value must stay below `0x80000000`.
pub const MAX_PARAGRAPH_ID: u32 = 0x7FFF_FFFF;
/// Largest paragraph ID [`allocate_paragraph_id`] generates; it never generates zero.
pub const MAX_GENERATED_PARAGRAPH_ID: u32 = 0x7FFF_FFFE;
const PROBES: u32 = 64;

/// Parses an authored paragraph ID: eight hexadecimal digits in either case, at
/// most [`MAX_PARAGRAPH_ID`].
pub fn parse_paragraph_id(value: &str) -> Option<u32> {
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    u32::from_str_radix(value, 16)
        .ok()
        .filter(|id| *id <= MAX_PARAGRAPH_ID)
}

pub fn format_paragraph_id(id: u32) -> String {
    format!("{id:08X}")
}

/// The `attempt`-th allocation candidate for `owner`: FNV-1a over the owner's
/// UTF-8, a `0xFF` separator and the little-endian attempt, mapped into
/// `1..=MAX_GENERATED_PARAGRAPH_ID`.
pub fn paragraph_id_candidate(owner: &str, attempt: u32) -> u32 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in owner.bytes().chain([0xFF]).chain(attempt.to_le_bytes()) {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    1 + (hash % u64::from(MAX_GENERATED_PARAGRAPH_ID)) as u32
}

/// Allocates deterministically for `owner`: the first of 64 hashed candidates
/// `occupied` lacks, else the smallest ID it lacks. `None` only when no ID is
/// free.
pub fn allocate_paragraph_id(owner: &str, occupied: &BTreeSet<u32>) -> Option<u32> {
    allocate_paragraph_id_where(owner, |id| occupied.contains(&id))
}

/// [`allocate_paragraph_id`] over an occupancy test, for IDs held in several sets.
pub fn allocate_paragraph_id_where(owner: &str, occupied: impl Fn(u32) -> bool) -> Option<u32> {
    allocate_below(owner, occupied, MAX_GENERATED_PARAGRAPH_ID)
}

fn allocate_below(owner: &str, occupied: impl Fn(u32) -> bool, max: u32) -> Option<u32> {
    (0..PROBES)
        .map(|attempt| paragraph_id_candidate(owner, attempt))
        .find(|id| *id <= max && !occupied(*id))
        .or_else(|| (1..=max).find(|id| !occupied(*id)))
}

/// Every valid paragraph ID an XML part of the package already uses: the
/// `paraId` and `paraIdParent` attributes of stories, notes, comments, their
/// companion parts and nested XML alike, each part counted once.
pub fn package_paragraph_ids(parts: &[(String, Vec<u8>)]) -> BTreeSet<u32> {
    let mut ids = BTreeSet::new();
    for (path, bytes) in parts {
        if path.to_ascii_lowercase().ends_with(".xml") {
            ids.extend(paragraph_id_attributes(bytes));
        }
    }
    ids
}

/// The valid Word paragraph ID of every `w:p` in each XML part of a DOCX
/// package, by part URI, in document order and in canonical form. Namespaces
/// resolve as in [`paragraph_occurrences`].
pub fn paragraph_ids_by_part(
    data: &[u8],
) -> Result<BTreeMap<String, Vec<String>>, crate::xml::ParseError> {
    let parts = ooxml_opc::unzip_parts(data).map_err(crate::xml::ParseError::Container)?;
    Ok(parts
        .iter()
        .filter(|(path, _)| path.to_ascii_lowercase().ends_with(".xml"))
        .filter_map(|(path, bytes)| {
            let occurrences = paragraph_occurrences(std::str::from_utf8(bytes).ok()?)?;
            let ids: Vec<String> = occurrences
                .into_iter()
                .filter_map(|occurrence| parse_paragraph_id(occurrence.para_id.as_deref()?))
                .map(format_paragraph_id)
                .collect();
            (!ids.is_empty()).then(|| (format!("/{}", path.trim_start_matches('/')), ids))
        })
        .collect())
}

/// The valid values of every `paraId` and `paraIdParent` attribute of one XML part.
pub fn paragraph_id_attributes(xml: &[u8]) -> BTreeSet<u32> {
    let mut ids = BTreeSet::new();
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                for attribute in element.attributes().flatten() {
                    if matches!(
                        attribute.key.local_name().as_ref(),
                        b"paraId" | b"paraIdParent"
                    ) && let Some(id) = std::str::from_utf8(&attribute.value)
                        .ok()
                        .and_then(parse_paragraph_id)
                    {
                        ids.insert(id);
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {}
        }
    }
    ids
}

/// One `w:p` start tag of a part; ordinals count them in document order.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParagraphOccurrence {
    pub ordinal: u32,
    /// Byte range of the start tag.
    pub tag: Range<usize>,
    /// The authored paragraph ID: `w14:paraId`, else `paraId`, else `w:paraId`.
    pub para_id: Option<String>,
    /// The `w:id` of the enclosing `w:comment`, `w:footnote` or `w:endnote`.
    pub item_id: Option<String>,
}

struct Tag<'a> {
    range: Range<usize>,
    name: &'a str,
    end: bool,
    empty: bool,
    attributes: Vec<(&'a str, Range<usize>)>,
}

/// Walks the tags of `xml` in order, skipping comments, CDATA, processing
/// instructions and declarations. `None` on an unterminated construct.
fn tags(xml: &str) -> Option<Vec<Tag<'_>>> {
    let bytes = xml.as_bytes();
    let mut result = Vec::new();
    let mut cursor = 0;
    while let Some(relative) = xml[cursor..].find('<') {
        let start = cursor + relative;
        let rest = &xml[start..];
        let skip = if rest.starts_with("<!--") {
            Some("-->")
        } else if rest.starts_with("<![CDATA[") {
            Some("]]>")
        } else if rest.starts_with("<?") {
            Some("?>")
        } else {
            None
        };
        if let Some(terminator) = skip {
            cursor = start + rest.find(terminator)? + terminator.len();
            continue;
        }
        let mut quote = None;
        let mut close = None;
        for (offset, byte) in bytes[start + 1..].iter().enumerate() {
            match (quote, *byte) {
                (None, b'"' | b'\'') => quote = Some(*byte),
                (Some(current), byte) if current == byte => quote = None,
                (None, b'>') => {
                    close = Some(start + 1 + offset);
                    break;
                }
                _ => {}
            }
        }
        let close = close?;
        cursor = close + 1;
        if rest.starts_with("<!") {
            continue;
        }
        let end = rest.starts_with("</");
        let inner = &xml[start + if end { 2 } else { 1 }..close];
        let name_end = inner
            .find(|character: char| character.is_ascii_whitespace() || character == '/')
            .unwrap_or(inner.len());
        let name = &inner[..name_end];
        let empty = !end && inner.trim_end().ends_with('/');
        let mut attributes = Vec::new();
        if !end {
            let base = start + 1;
            let mut position = name_end;
            let text = inner.as_bytes();
            loop {
                while position < text.len() && text[position].is_ascii_whitespace() {
                    position += 1;
                }
                if position >= text.len() || text[position] == b'/' {
                    break;
                }
                let key_start = position;
                while position < text.len()
                    && !text[position].is_ascii_whitespace()
                    && text[position] != b'='
                {
                    position += 1;
                }
                let key = &inner[key_start..position];
                while position < text.len() && text[position] != b'"' && text[position] != b'\'' {
                    position += 1;
                }
                let quote = *text.get(position)?;
                let value_start = position + 1;
                let value_end = value_start + inner[value_start..].find(quote as char)?;
                attributes.push((key, base + value_start..base + value_end));
                position = value_end + 1;
            }
        }
        result.push(Tag {
            range: start..close + 1,
            name,
            end,
            empty,
            attributes,
        });
    }
    Some(result)
}

fn attribute<'a>(tag: &Tag<'a>, name: &str) -> Option<Range<usize>> {
    tag.attributes
        .iter()
        .find(|(key, _)| *key == name)
        .map(|(_, range)| range.clone())
}

/// Which of a `w:p` element's attribute names is its Word 2010 `paraId`: one
/// whose prefix `resolve` binds to that namespace, else a literal
/// `w14:paraId` when `w14` is unbound.
pub fn word_para_id_name<'a, 'b>(
    names: impl IntoIterator<Item = &'a str>,
    resolve: impl Fn(&str) -> Option<&'b str>,
) -> Option<&'a str> {
    let mut unbound = None;
    for name in names {
        let Some((prefix, "paraId")) = name.split_once(':') else {
            continue;
        };
        match resolve(prefix) {
            Some(W14_NAMESPACE) => return Some(name),
            None if prefix == "w14" => unbound = Some(name),
            _ => {}
        }
    }
    unbound
}

/// The namespace declarations in scope while walking a part's tags.
#[derive(Default)]
struct Scopes<'a>(Vec<Vec<(&'a str, &'a str)>>);

impl<'a> Scopes<'a> {
    fn enter(&mut self, xml: &'a str, tag: &Tag<'a>) {
        self.0.push(
            tag.attributes
                .iter()
                .filter_map(|(key, range)| {
                    Some((key.strip_prefix("xmlns:")?, xml[range.clone()].trim()))
                })
                .collect(),
        );
    }

    fn leave(&mut self) {
        self.0.pop();
    }

    fn resolve(&self, prefix: &str) -> Option<&'a str> {
        self.0.iter().rev().find_map(|scope| {
            scope
                .iter()
                .find(|(bound, _)| *bound == prefix)
                .map(|(_, uri)| *uri)
        })
    }

    /// The Word 2010 `paraId` attribute of `tag`; see [`word_para_id_name`].
    fn word_para_id(&self, tag: &Tag<'_>) -> Option<Range<usize>> {
        let name = word_para_id_name(tag.attributes.iter().map(|(key, _)| *key), |prefix| {
            self.resolve(prefix)
        })?;
        attribute(tag, name)
    }

    /// The value range of the paragraph ID attribute the parser reads, in its order.
    fn para_id(&self, tag: &Tag<'_>) -> Option<Range<usize>> {
        self.word_para_id(tag)
            .or_else(|| attribute(tag, "paraId"))
            .or_else(|| attribute(tag, "w:paraId"))
    }
}

/// Walks the tags of `xml` with the namespace bindings in scope at each start
/// tag; `visit` sees every tag, end tags included.
fn walk<'a>(xml: &'a str, mut visit: impl FnMut(&Tag<'a>, &Scopes<'a>)) -> Option<()> {
    let mut scopes = Scopes::default();
    for tag in tags(xml)? {
        if tag.end {
            visit(&tag, &scopes);
            scopes.leave();
            continue;
        }
        scopes.enter(xml, &tag);
        visit(&tag, &scopes);
        if tag.empty {
            scopes.leave();
        }
    }
    Some(())
}

/// Every `w:p` start tag of a part in document order. `None` when the XML
/// cannot be scanned.
pub fn paragraph_occurrences(xml: &str) -> Option<Vec<ParagraphOccurrence>> {
    let mut items: Vec<Option<String>> = Vec::new();
    let mut occurrences = Vec::new();
    walk(xml, |tag, scopes| match tag.name {
        "w:comment" | "w:footnote" | "w:endnote" => {
            if tag.end {
                items.pop();
            } else if !tag.empty {
                items.push(attribute(tag, "w:id").map(|range| xml[range].to_owned()));
            }
        }
        "w:p" if !tag.end => occurrences.push(ParagraphOccurrence {
            ordinal: occurrences.len() as u32,
            para_id: scopes.para_id(tag).map(|range| xml[range].to_owned()),
            item_id: items.last().cloned().flatten(),
            tag: tag.range.clone(),
        }),
        _ => {}
    })?;
    Some(occurrences)
}

/// Sets the paragraph ID of the `w:p` start tags at the `patches` ordinals and
/// leaves every other byte in place. Namespaces resolve at each patched tag:
/// its Word 2010 `paraId` attribute is replaced, or one is inserted under a
/// prefix bound to that namespace there, declaring one on the tag when its
/// `w14` is bound elsewhere. With `declare`, a root that does not declare
/// `w14` gains the namespace and lists it in `mc:Ignorable`. `None` when an
/// ordinal is absent, the XML cannot be scanned, or the root binds a needed
/// prefix to another namespace.
pub fn patch_paragraph_ids(
    xml: &str,
    patches: &BTreeMap<u32, String>,
    declare: bool,
) -> Option<String> {
    let mut edits: Vec<(Range<usize>, String)> = Vec::new();
    let mut needs_root = false;
    let mut ordinal = 0u32;
    walk(xml, |tag, scopes| {
        if tag.end || tag.name != "w:p" {
            return;
        }
        if let Some(id) = patches.get(&ordinal) {
            let (edit, unbound) = paragraph_id_edit(tag, scopes, id);
            needs_root |= unbound;
            edits.push(edit);
        }
        ordinal += 1;
    })?;
    if patches.keys().any(|patch| *patch >= ordinal) {
        return None;
    }
    let tags = tags(xml)?;
    let root = tags.iter().find(|tag| !tag.end)?;
    if needs_root && declare {
        match attribute(root, "xmlns:w14") {
            Some(range) if xml[range.clone()].trim() != W14_NAMESPACE => return None,
            Some(_) => {}
            None => {
                let mut declarations = format!(" xmlns:w14=\"{W14_NAMESPACE}\"");
                match attribute(root, "mc:Ignorable") {
                    Some(range) => {
                        if !xml[range.clone()]
                            .split_ascii_whitespace()
                            .any(|prefix| prefix == "w14")
                        {
                            let at = range.end;
                            edits.push((at..at, " w14".to_owned()));
                        }
                    }
                    None => {
                        match attribute(root, "xmlns:mc") {
                            Some(range) if xml[range.clone()].trim() != MC_NAMESPACE => {
                                return None;
                            }
                            Some(_) => {}
                            None => declarations.push_str(&format!(" xmlns:mc=\"{MC_NAMESPACE}\"")),
                        }
                        declarations.push_str(" mc:Ignorable=\"w14\"");
                    }
                }
                let at = root.range.end - if root.empty { 2 } else { 1 };
                edits.push((at..at, declarations));
            }
        }
    }
    Some(apply(xml, edits))
}

/// The edit setting a `w:p` tag's Word 2010 paragraph ID under the namespace
/// bindings in scope at it, and whether it relies on the root declaring `w14`.
fn paragraph_id_edit(
    tag: &Tag<'_>,
    scopes: &Scopes<'_>,
    id: &str,
) -> ((Range<usize>, String), bool) {
    if let Some(range) = scopes.word_para_id(tag) {
        return ((range, id.to_owned()), false);
    }
    let at = tag.range.start + "<w:p".len();
    let insert = |text: String| (at..at, text);
    match scopes.resolve("w14") {
        Some(W14_NAMESPACE) => (insert(format!(" w14:paraId=\"{id}\"")), false),
        None => (insert(format!(" w14:paraId=\"{id}\"")), true),
        Some(_) => {
            let bound = scopes.0.iter().rev().flatten().find(|(prefix, uri)| {
                *uri == W14_NAMESPACE && scopes.resolve(prefix) == Some(W14_NAMESPACE)
            });
            match bound {
                Some((prefix, _)) => (insert(format!(" {prefix}:paraId=\"{id}\"")), false),
                None => {
                    let prefix = (0..)
                        .map(|index| format!("w14p{index}"))
                        .find(|candidate| {
                            scopes.resolve(candidate).is_none()
                                && !tag.attributes.iter().any(|(key, _)| {
                                    key.split_once(':')
                                        .is_some_and(|(used, _)| used == candidate)
                                })
                        })
                        .unwrap_or_default();
                    (
                        insert(format!(
                            " xmlns:{prefix}=\"{W14_NAMESPACE}\" {prefix}:paraId=\"{id}\""
                        )),
                        true,
                    )
                }
            }
        }
    }
}

/// Rewrites every `paraId` and `paraIdParent` attribute value, under any
/// prefix, that numerically equals a key of `renames`, leaving every other
/// byte in place. `None` when the XML cannot be scanned.
pub fn patch_paragraph_id_references(xml: &str, renames: &HashMap<u32, String>) -> Option<String> {
    let mut edits = Vec::new();
    for tag in tags(xml)? {
        for (key, range) in &tag.attributes {
            let local = key.rsplit(':').next().unwrap_or(key);
            if matches!(local, "paraId" | "paraIdParent")
                && let Some(renamed) =
                    parse_paragraph_id(&xml[range.clone()]).and_then(|id| renames.get(&id))
            {
                edits.push((range.clone(), renamed.clone()));
            }
        }
    }
    Some(apply(xml, edits))
}

fn apply(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> String {
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    let mut output = xml.to_owned();
    for (range, text) in edits {
        output.replace_range(range, &text);
    }
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_eight_hex_digits_below_the_word_limit_in_either_case() {
        assert_eq!(parse_paragraph_id("0000000A"), Some(10));
        assert_eq!(parse_paragraph_id("0000000a"), Some(10));
        assert_eq!(parse_paragraph_id("00000000"), Some(0));
        assert_eq!(parse_paragraph_id("7FFFFFFF"), Some(MAX_PARAGRAPH_ID));
        for invalid in [
            "80000000",
            "FFFFFFFF",
            "0000000",
            "000000000",
            "0000000G",
            "",
            "100:5",
        ] {
            assert_eq!(parse_paragraph_id(invalid), None, "{invalid:?}");
        }
    }

    #[test]
    fn allocation_is_deterministic_probes_collisions_and_reports_exhaustion() {
        for owner in ["7:0", "body:p3", "1~42", ""] {
            for attempt in 0..PROBES {
                let id = paragraph_id_candidate(owner, attempt);
                assert_eq!(id, paragraph_id_candidate(owner, attempt));
                assert!((1..=MAX_GENERATED_PARAGRAPH_ID).contains(&id));
            }
        }
        let first = paragraph_id_candidate("owner", 0);
        assert_eq!(
            allocate_paragraph_id("owner", &BTreeSet::new()),
            Some(first)
        );
        assert_eq!(
            allocate_paragraph_id("owner", &BTreeSet::from([first])),
            Some(paragraph_id_candidate("owner", 1))
        );
        let probes: BTreeSet<u32> = (0..PROBES)
            .map(|attempt| paragraph_id_candidate("owner", attempt))
            .chain([1, 2])
            .collect();
        assert_eq!(allocate_paragraph_id("owner", &probes), Some(3));
        assert_eq!(allocate_below("owner", |id| id <= 3, 3), None);
        assert_eq!(allocate_below("owner", |id| id != 2, 3), Some(2));
    }

    #[test]
    fn inventory_reads_every_xml_part_once_and_ignores_other_parts() {
        let parts = vec![
            (
                "word/document.xml".to_owned(),
                br#"<w:document><w:p w14:paraId="0000000A"/><w:txbxContent><w:p w14:paraId="0000000b"/></w:txbxContent><w:p w14:paraId="bad"/></w:document>"#.to_vec(),
            ),
            (
                "word/commentsExtended.xml".to_owned(),
                br#"<w15:commentsEx><w15:commentEx w15:paraId="0000000C" w15:paraIdParent="0000000D"/></w15:commentsEx>"#.to_vec(),
            ),
            (
                "word/_rels/document.xml.rels".to_owned(),
                br#"<Relationships paraId="0000000E"/>"#.to_vec(),
            ),
        ];
        assert_eq!(
            package_paragraph_ids(&parts)
                .into_iter()
                .collect::<Vec<_>>(),
            vec![10, 11, 12, 13]
        );
    }

    const PART: &str = concat!(
        r#"<?xml version="1.0"?><!-- <w:p> --><w:comments xmlns:w="w" >"#,
        r#"<w:comment w:id="3"><w:p w14:paraId="0000000A"/><w:p><![CDATA[<w:p>]]></w:p></w:comment>"#,
        r#"<w:pPr/><w:p paraId='0000000B' w:rsidR="00"><w:pStyle/></w:p></w:comments>"#
    );

    #[test]
    fn occurrences_count_start_tags_outside_comments_and_cdata() {
        let occurrences = paragraph_occurrences(PART).unwrap();
        let summary: Vec<_> = occurrences
            .iter()
            .map(|occurrence| {
                (
                    occurrence.ordinal,
                    occurrence.para_id.as_deref(),
                    occurrence.item_id.as_deref(),
                    &PART[occurrence.tag.clone()],
                )
            })
            .collect();
        assert_eq!(
            summary,
            [
                (
                    0,
                    Some("0000000A"),
                    Some("3"),
                    r#"<w:p w14:paraId="0000000A"/>"#
                ),
                (1, None, Some("3"), "<w:p>"),
                (
                    2,
                    Some("0000000B"),
                    None,
                    r#"<w:p paraId='0000000B' w:rsidR="00">"#
                ),
            ]
        );
    }

    #[test]
    fn patches_touch_only_the_start_tags_and_the_root_declarations() {
        let patches = BTreeMap::from([(1, "00000011".to_owned()), (2, "00000012".to_owned())]);
        let patched = patch_paragraph_ids(PART, &patches, true).unwrap();
        let expected = PART
            .replacen(
                "<w:p><![CDATA",
                r#"<w:p w14:paraId="00000011"><![CDATA"#,
                1,
            )
            .replacen(
                "<w:p paraId='0000000B'",
                r#"<w:p w14:paraId="00000012" paraId='0000000B'"#,
                1,
            )
            .replacen(
                r#"xmlns:w="w" >"#,
                &format!(
                    r#"xmlns:w="w"  xmlns:w14="{W14_NAMESPACE}" xmlns:mc="{MC_NAMESPACE}" mc:Ignorable="w14">"#
                ),
                1,
            );
        assert_eq!(patched, expected);
        assert_eq!(
            patch_paragraph_ids(PART, &BTreeMap::from([(2, "00000012".to_owned())]), false)
                .unwrap(),
            PART.replacen(
                "<w:p paraId='0000000B'",
                r#"<w:p w14:paraId="00000012" paraId='0000000B'"#,
                1
            )
        );
        let ignorable =
            format!(r#"<w:hdr xmlns:mc="{MC_NAMESPACE}" mc:Ignorable="w15"><w:p/></w:hdr>"#);
        assert_eq!(
            patch_paragraph_ids(
                &ignorable,
                &BTreeMap::from([(0, "00000001".to_owned())]),
                true
            )
            .unwrap(),
            ignorable.replacen("\"w15\"", "\"w15 w14\"", 1).replacen(
                "><w:p/>",
                &format!(r#" xmlns:w14="{W14_NAMESPACE}"><w:p w14:paraId="00000001"/>"#),
                1
            )
        );
        let declared = format!(r#"<w:hdr xmlns:w14="{W14_NAMESPACE}"><w:p/></w:hdr>"#);
        assert_eq!(
            patch_paragraph_ids(
                &declared,
                &BTreeMap::from([(0, "00000001".to_owned())]),
                true
            )
            .unwrap(),
            declared.replacen("<w:p/>", r#"<w:p w14:paraId="00000001"/>"#, 1)
        );
        assert_eq!(
            patch_paragraph_ids(PART, &BTreeMap::from([(3, "00000001".to_owned())]), true),
            None
        );
        assert_eq!(
            patch_paragraph_ids(
                r#"<w:hdr xmlns:w14="urn:other"><w:p/></w:hdr>"#,
                &BTreeMap::from([(0, "00000001".to_owned())]),
                true
            ),
            None
        );
    }

    #[test]
    fn reference_patches_rename_every_matching_companion_attribute() {
        let companions = r#"<w15:commentsEx><w15:commentEx w15:paraId="0000000a" w15:done="0"/><w15:commentEx w15:paraId="0000000C" w15:paraIdParent="0000000A"/></w15:commentsEx>"#;
        let renames = HashMap::from([(10, "00000099".to_owned())]);
        assert_eq!(
            patch_paragraph_id_references(companions, &renames).unwrap(),
            companions.replacen("0000000a", "00000099", 1).replacen(
                "\"0000000A\"",
                "\"00000099\"",
                1
            )
        );
    }

    /// The value of each `w:p`'s attribute that resolves to the Word 2010 `paraId`.
    fn resolved_para_ids(xml: &str) -> Vec<Option<String>> {
        use quick_xml::NsReader;
        use quick_xml::name::{Namespace, ResolveResult};
        let mut reader = NsReader::from_str(xml);
        let mut ids = Vec::new();
        loop {
            match reader.read_event().unwrap() {
                Event::Start(element) | Event::Empty(element)
                    if element.name().as_ref() == b"w:p" =>
                {
                    let mut id = None;
                    for attribute in element.attributes() {
                        let attribute = attribute.unwrap();
                        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                        if local.as_ref() == b"paraId"
                            && namespace
                                == ResolveResult::Bound(Namespace(W14_NAMESPACE.as_bytes()))
                        {
                            id = Some(String::from_utf8(attribute.value.to_vec()).unwrap());
                        }
                    }
                    ids.push(id);
                }
                Event::Eof => break,
                _ => {}
            }
        }
        ids
    }

    #[test]
    fn patches_resolve_the_word_namespace_at_each_paragraph() {
        let xml = format!(
            r#"<w:document xmlns:w="w" xmlns:w14="{W14_NAMESPACE}" xmlns:mc="{MC_NAMESPACE}" mc:Ignorable="w14"><w:body><w:p xmlns:w14="urn:other" w14:paraId="ABCDEF01"/><x:wrap xmlns:x="x" xmlns:wx="{W14_NAMESPACE}" xmlns:w14="urn:other"><w:p/><w:p wx:paraId="0000000F"/></x:wrap><w:p/></w:body></w:document>"#
        );
        let patches = BTreeMap::from([
            (0, "0000000A".to_owned()),
            (1, "0000000B".to_owned()),
            (2, "0000000C".to_owned()),
            (3, "0000000D".to_owned()),
        ]);
        let patched = patch_paragraph_ids(&xml, &patches, true).unwrap();
        assert_eq!(
            patched,
            xml.replacen(
                r#"<w:p xmlns:w14="urn:other""#,
                &format!(
                    r#"<w:p xmlns:w14p0="{W14_NAMESPACE}" w14p0:paraId="0000000A" xmlns:w14="urn:other""#
                ),
                1
            )
            .replacen("<w:p/><w:p wx", r#"<w:p wx:paraId="0000000B"/><w:p wx"#, 1)
            .replacen(r#"wx:paraId="0000000F""#, r#"wx:paraId="0000000C""#, 1)
            .replacen("<w:p/></w:body>", r#"<w:p w14:paraId="0000000D"/></w:body>"#, 1)
        );
        assert_eq!(
            resolved_para_ids(&patched),
            ["0000000A", "0000000B", "0000000C", "0000000D"].map(|id| Some(id.to_owned()))
        );
    }
}
