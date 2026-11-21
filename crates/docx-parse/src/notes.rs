//! Non-recursive footnote/endnote property leaves used by settings and sections.

use serde::{Deserialize, Serialize};

use crate::block::{BlockContent, StoryParser};
use crate::xml::XmlElement;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Note {
    #[serde(rename = "type")]
    pub story_type: String,
    pub id: f64,
    pub note_type: String,
    pub content: Vec<BlockContent>,
    /// Inherited root bindings required by foreign note content.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_root_bindings: Vec<crate::paragraph::RawAttribute>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub verbatim_xml: Option<String>,
    /// The occurrence of the first `w:p` inside `verbatim_xml` in its part,
    /// when the parse records source ordinals.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source_ordinal: Option<u32>,
}

impl Note {
    pub fn is_separator(&self) -> bool {
        matches!(
            self.note_type.as_str(),
            "separator" | "continuationSeparator" | "continuationNotice"
        )
    }
}

/// Parses a complete footnote or endnote story part.
pub fn parse_notes(
    root: &XmlElement,
    footnotes: bool,
    parser: &mut StoryParser<'_, '_>,
) -> Result<Vec<Note>, crate::xml::ParseError> {
    let root_name = if footnotes { "footnotes" } else { "endnotes" };
    let note_name = if footnotes { "footnote" } else { "endnote" };
    if root.local_name() != root_name {
        return Ok(Vec::new());
    }
    let mut notes = Vec::new();
    for element in root.children_named("w", note_name) {
        parser.budget.charge_note(parser.part)?;
        let Some(id) = element.parse_numeric_attribute(Some("w"), "id", 1.0) else {
            continue;
        };
        let note_type = parse_note_type(element.attribute(Some("w"), "type"));
        let content = parser.parse_blocks(element, 0, false)?;
        let bookmarked = has_block_bookmark(element);
        let mut custom_root_bindings = if has_foreign_content(element) {
            note_bindings(root, element)
        } else {
            Vec::new()
        };
        if bookmarked {
            for binding in bookmark_bindings(root, element) {
                if !custom_root_bindings
                    .iter()
                    .any(|existing| existing.name == binding.name)
                {
                    custom_root_bindings.push(binding);
                }
            }
        }
        let verbatim_xml = needs_verbatim(element).then(|| {
            let mut bound = element.clone();
            for attribute in &custom_root_bindings {
                bound
                    .attributes
                    .entry(attribute.name.clone())
                    .or_insert_with(|| attribute.value.clone());
            }
            if bookmarked {
                bound.to_bookmarked_raw_inline_xml()
            } else {
                bound.to_raw_inline_xml()
            }
        });
        let source_ordinal = verbatim_xml
            .as_ref()
            .filter(|_| parser.budget.records_source_ordinals())
            .and_then(|_| element.first_paragraph_ordinal());
        notes.push(Note {
            story_type: note_name.to_owned(),
            id,
            note_type: note_type.to_owned(),
            content,
            custom_root_bindings,
            verbatim_xml,
            source_ordinal,
        });
    }
    Ok(notes)
}

fn note_bindings(root: &XmlElement, note: &XmlElement) -> Vec<crate::paragraph::RawAttribute> {
    let mut bindings = crate::document::custom_root_bindings(root);
    for (name, value) in &note.attributes {
        if name.starts_with("xmlns:") && !bindings.iter().any(|binding| &binding.name == name) {
            bindings.push(crate::paragraph::RawAttribute {
                name: name.clone(),
                value: value.clone(),
            });
        }
    }
    bindings
}

fn bookmark_bindings(root: &XmlElement, note: &XmlElement) -> Vec<crate::paragraph::RawAttribute> {
    fn required<'a>(
        element: &'a XmlElement,
        scopes: &mut Vec<&'a str>,
        prefixes: &mut std::collections::HashSet<&'a str>,
        root: bool,
    ) {
        let mark = scopes.len();
        if !root {
            scopes.extend(
                element
                    .attributes
                    .keys()
                    .filter_map(|name| name.strip_prefix("xmlns:")),
            );
        }
        for name in std::iter::once(element.name.as_str())
            .chain(element.attributes.keys().map(String::as_str))
        {
            if let Some((prefix, _)) = name.split_once(':')
                && !matches!(prefix, "xml" | "xmlns")
                && !scopes.contains(&prefix)
            {
                prefixes.insert(prefix);
            }
        }
        for (name, value) in &element.attributes {
            if matches!(
                crate::xml::local_name(name),
                "Ignorable"
                    | "MustUnderstand"
                    | "Requires"
                    | "ProcessContent"
                    | "PreserveElements"
                    | "PreserveAttributes"
            ) {
                for token in value.split_whitespace() {
                    let prefix = token.split_once(':').map_or(token, |(prefix, _)| prefix);
                    if prefix != "xml" && !scopes.contains(&prefix) {
                        prefixes.insert(prefix);
                    }
                }
            }
        }
        for child in element.child_elements() {
            required(child, scopes, prefixes, false);
        }
        scopes.truncate(mark);
    }
    let mut prefixes = std::collections::HashSet::new();
    required(note, &mut Vec::new(), &mut prefixes, true);
    root.attributes
        .iter()
        .chain(&note.attributes)
        .filter(|(name, _)| name.starts_with("xmlns:"))
        .collect::<indexmap::IndexMap<_, _>>()
        .into_iter()
        .filter(|(name, value)| {
            prefixes.contains(name.strip_prefix("xmlns:").unwrap())
                && !crate::serializer::parts::note_root_declares(name, value)
        })
        .map(|(name, value)| crate::paragraph::RawAttribute {
            name: name.clone(),
            value: value.clone(),
        })
        .collect()
}

fn has_block_bookmark(element: &XmlElement) -> bool {
    crate::block::has_direct_block_bookmark(element)
        || element.child_elements().any(has_block_bookmark)
}

fn has_foreign_content(element: &XmlElement) -> bool {
    crate::inline::raw_foreign_inline(element).is_some()
        || element.child_elements().any(has_foreign_content)
}

pub fn parse_note_type(value: Option<&str>) -> &'static str {
    match value {
        Some("separator") => "separator",
        Some("continuationSeparator") => "continuationSeparator",
        Some("continuationNotice") => "continuationNotice",
        _ => "normal",
    }
}

fn needs_verbatim(note: &XmlElement) -> bool {
    note.child_elements().any(|child| {
        matches!(
            child.local_name(),
            "bookmarkStart" | "bookmarkEnd" | "customXml"
        )
    })
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteSeparatorReference {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub kind: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_id: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NoteProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub custom_number_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub separators: Option<Vec<NoteSeparatorReference>>,
}

pub fn parse_footnote_properties(element: Option<&XmlElement>) -> NoteProperties {
    parse_note_properties(element, true)
}

pub fn parse_endnote_properties(element: Option<&XmlElement>) -> NoteProperties {
    parse_note_properties(element, false)
}

fn parse_note_properties(element: Option<&XmlElement>, footnote: bool) -> NoteProperties {
    let Some(element) = element else {
        return NoteProperties::default();
    };
    let position = element
        .child("w", "pos")
        .and_then(|value| value.attribute(Some("w"), "val"))
        .filter(|value| {
            if footnote {
                matches!(*value, "pageBottom" | "beneathText" | "sectEnd" | "docEnd")
            } else {
                matches!(*value, "sectEnd" | "docEnd")
            }
        })
        .map(str::to_owned);
    let num_fmt_element = element.child("w", "numFmt");
    let num_fmt = num_fmt_element
        .and_then(|value| value.attribute(Some("w"), "val"))
        .filter(|value| {
            matches!(
                *value,
                "decimal"
                    | "upperRoman"
                    | "lowerRoman"
                    | "upperLetter"
                    | "lowerLetter"
                    | "ordinal"
                    | "cardinalText"
                    | "ordinalText"
                    | "bullet"
                    | "chicago"
                    | "none"
            )
        })
        .map(str::to_owned);
    let custom_number_format = num_fmt_element
        .and_then(|value| value.attribute(Some("w"), "format"))
        .filter(|value| !value.is_empty())
        .map(|value| truncate_utf16_scalars(value, 255));
    let num_start = element
        .child("w", "numStart")
        .and_then(|value| value.parse_numeric_attribute(Some("w"), "val", 1.0));
    let num_restart = element
        .child("w", "numRestart")
        .and_then(|value| value.attribute(Some("w"), "val"))
        .filter(|value| matches!(*value, "continuous" | "eachSect" | "eachPage"))
        .map(str::to_owned);
    let reference_name = if footnote { "footnote" } else { "endnote" };
    let separators: Vec<_> = element
        .children_named("w", reference_name)
        .take(32)
        .filter_map(|reference| reference.parse_numeric_attribute(Some("w"), "id", 1.0))
        .map(|note_id| NoteSeparatorReference {
            kind: Some(
                if note_id == -1.0 {
                    "separator"
                } else if note_id == 0.0 {
                    "continuationSeparator"
                } else {
                    "continuationNotice"
                }
                .to_owned(),
            ),
            note_id: Some(note_id),
        })
        .collect();

    NoteProperties {
        position,
        num_fmt,
        num_start,
        num_restart,
        custom_number_format,
        separators: (!separators.is_empty()).then_some(separators),
    }
}

fn truncate_utf16_scalars(value: &str, max_units: usize) -> String {
    let mut units = 0usize;
    value
        .chars()
        .take_while(|character| {
            let next = units + character.len_utf16();
            if next > max_units {
                false
            } else {
                units = next;
                true
            }
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::chart::ChartPartsMap;
    use crate::media::MediaMap;
    use crate::paragraph::HexIdAllocator;
    use crate::smart_art::SmartArtContext;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(xml.as_bytes(), "notes.xml", &mut ParseBudget::new(&limits))
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }

    #[test]
    fn parses_known_values_truncates_custom_format_and_caps_separators() {
        let references = (0..40)
            .map(|id| format!(r#"<w:footnote w:id="{id}"/>"#))
            .collect::<String>();
        let element = root(&format!(
            r#"<w:footnotePr><w:pos w:val="pageBottom"/><w:numFmt w:val="decimal" w:format="{}"/><w:numStart w:val="2x"/><w:numRestart w:val="eachSect"/>{references}</w:footnotePr>"#,
            "a".repeat(300)
        ));
        let parsed = parse_footnote_properties(Some(&element));
        assert_eq!(parsed.position.as_deref(), Some("pageBottom"));
        assert_eq!(parsed.custom_number_format.unwrap().len(), 255);
        assert_eq!(parsed.num_start, Some(2.0));
        assert_eq!(parsed.separators.unwrap().len(), 32);
    }

    #[test]
    fn ignores_unknown_values_and_handles_huge_numbers() {
        let element = root(&format!(
            r#"<w:endnotePr><w:pos w:val="pageBottom"/><w:numFmt w:val="future"/><w:numStart w:val="{}"/><w:numRestart w:val="always"/></w:endnotePr>"#,
            "9".repeat(10_000)
        ));
        assert_eq!(
            parse_endnote_properties(Some(&element)),
            NoteProperties::default()
        );
    }

    fn parse_story(xml: &str, footnotes: bool, limits: &ParseLimits) -> Vec<Note> {
        let mut budget = ParseBudget::new(limits);
        let document = parse_xml(xml.as_bytes(), "word/footnotes.xml", &mut budget).unwrap();
        let media = MediaMap::new();
        let charts = ChartPartsMap::new();
        let mut smart_art = SmartArtContext::default();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        let mut parser = StoryParser {
            relationships: None,
            theme: None,
            styles: None,
            doc_defaults: None,
            numbering: None,
            media: &media,
            charts: &charts,
            smart_art: &mut smart_art,
            budget: &mut budget,
            ids: &mut ids,
            part: "word/footnotes.xml",
        };
        parse_notes(document.root().unwrap(), footnotes, &mut parser).unwrap()
    }

    #[test]
    fn keeps_note_zero_and_drops_notes_without_an_id() {
        for (footnotes, name) in [(true, "footnote"), (false, "endnote")] {
            let notes = parse_story(
                &format!(
                    r#"<w:{name}s xmlns:w="w"><w:{name}><w:p/></w:{name}><w:{name} w:id="x"><w:p/></w:{name}><w:{name} w:id="-1" w:type="separator"><w:p/></w:{name}><w:{name} w:id="0" w:type="continuationSeparator"><w:p/></w:{name}></w:{name}s>"#,
                ),
                footnotes,
                &ParseLimits::default(),
            );
            assert_eq!(notes.len(), 2);
            assert_eq!(notes[0].id, -1.0);
            assert_eq!(notes[0].note_type, "separator");
            assert_eq!(notes[1].id, 0.0);
            assert_eq!(notes[1].note_type, "continuationSeparator");
            assert_eq!(notes[1].content.len(), 1);
        }
    }

    #[test]
    fn note_owner_preserves_full_block_order_and_special_type_quirks() {
        let notes = parse_story(
            r#"<w:footnotes xmlns:w="w"><w:footnote w:id="-1"><w:p/></w:footnote><w:footnote w:id="0" w:type="continuationSeparator"><w:p/><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:footnote></w:footnotes>"#,
            true,
            &ParseLimits::default(),
        );
        // Special ids do not imply a type; only w:type classifies a note.
        assert_eq!(notes[0].id, -1.0);
        assert_eq!(notes[0].note_type, "normal");
        assert_eq!(notes[1].note_type, "continuationSeparator");
        assert!(notes[1].is_separator());
        assert_eq!(
            notes[1]
                .content
                .iter()
                .map(BlockContent::node_type)
                .collect::<Vec<_>>(),
            ["paragraph", "table", "blockSdt"]
        );
    }

    #[test]
    fn verbatim_gate_is_shallow_and_preserves_raw_attribute_text() {
        let notes = parse_story(
            r#"<w:endnotes xmlns:w="w"><w:endnote w:id="1" label="a&amp;b"><w:p/><w:bookmarkStart w:id="2"/></w:endnote><w:endnote w:id="2"><w:sdt><w:sdtContent><w:customXml><w:p/></w:customXml></w:sdtContent></w:sdt></w:endnote></w:endnotes>"#,
            false,
            &ParseLimits::default(),
        );
        assert!(
            notes[0]
                .verbatim_xml
                .as_deref()
                .unwrap()
                .contains("a&amp;b")
        );
        assert!(notes[1].verbatim_xml.is_none());
    }

    #[test]
    fn note_attributes_do_not_trigger_verbatim_without_direct_blocks() {
        for attributes in [r#"label="a&amp;b""#, r#"w:rsidR="00112233""#] {
            let notes = parse_story(
                &format!(
                    r#"<w:endnotes xmlns:w="w"><w:endnote w:id="1" {attributes}><w:p/></w:endnote><w:endnote w:id="2" {attributes}><w:bookmarkStart w:id="7" w:name="Range"/><w:p/><w:bookmarkEnd w:id="7"/></w:endnote><w:endnote w:id="3"><w:bookmarkStart w:id="8" w:name="Range"/><w:p/><w:bookmarkEnd w:id="8"/></w:endnote><w:endnote w:id="4"><w:customXml><w:p/></w:customXml></w:endnote></w:endnotes>"#
                ),
                false,
                &ParseLimits::default(),
            );
            assert!(notes[0].verbatim_xml.is_none());
            assert!(
                notes[1]
                    .verbatim_xml
                    .as_deref()
                    .unwrap()
                    .contains(attributes)
            );
            assert!(notes[2].verbatim_xml.is_some());
            assert_eq!(
                notes[2]
                    .content
                    .iter()
                    .map(BlockContent::node_type)
                    .collect::<Vec<_>>(),
                ["rawXml", "paragraph", "rawXml"]
            );
            assert!(notes[3].verbatim_xml.is_some());
        }
    }

    #[test]
    fn verbatim_notes_do_not_accumulate_part_root_bindings() {
        use crate::serializer::{
            SerializerContext, SerializerDeterminism, serialize_endnotes_part,
            serialize_footnotes_part,
        };

        for footnotes in [true, false] {
            let name = if footnotes { "footnote" } else { "endnote" };
            for content in ["<w:p/>", "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"] {
                let note = format!(
                    r#"<w:{name} w:id="1"><w:bookmarkStart w:id="7" w:name="Note"/><w:customXml>{content}</w:customXml><w:bookmarkEnd w:id="7"/></w:{name}>"#
                );
                let xml = format!(
                    r#"<w:{name}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:x="urn:unused" mc:Ignorable="cx x">{note}</w:{name}s>"#
                );
                let mut notes = parse_story(&xml, footnotes, &ParseLimits::default());
                let mut context = SerializerContext::new(&SerializerDeterminism {
                    seed: "0".repeat(64),
                    now: "2000-01-01T00:00:00.000Z".to_owned(),
                })
                .unwrap();
                let serialize = if footnotes {
                    serialize_footnotes_part
                } else {
                    serialize_endnotes_part
                };
                let mut previous = None;
                for _ in 0..3 {
                    assert_eq!(notes[0].verbatim_xml.as_deref(), Some(note.as_str()));
                    assert!(notes[0].custom_root_bindings.is_empty());
                    let saved = serialize(&notes, &mut context).unwrap();
                    if let Some(previous) = previous {
                        assert_eq!(saved, previous);
                    }
                    notes = parse_story(&saved, footnotes, &ParseLimits::default());
                    previous = Some(saved);
                }
            }
        }
    }

    #[test]
    fn foreign_notes_keep_their_root_bindings_and_ignorable_prefixes() {
        use crate::serializer::{
            SerializerContext, SerializerDeterminism, serialize_endnotes_part,
            serialize_footnotes_part,
        };

        for footnotes in [true, false] {
            let name = if footnotes { "footnote" } else { "endnote" };
            for custom_xml in [false, true] {
                let content = if custom_xml {
                    "<w:customXml><x:opaque/><w:p/></w:customXml>"
                } else {
                    "<x:opaque/><w:p/>"
                };
                let xml = format!(
                    r#"<w:{name}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:extension" mc:Ignorable="x"><w:{name} w:id="1">{content}</w:{name}></w:{name}s>"#
                );
                let notes = parse_story(&xml, footnotes, &ParseLimits::default());
                assert_eq!(
                    notes[0].custom_root_bindings,
                    vec![
                        crate::paragraph::RawAttribute {
                            name: "xmlns:x".to_owned(),
                            value: "urn:extension".to_owned(),
                        },
                        crate::paragraph::RawAttribute {
                            name: "mc:Ignorable".to_owned(),
                            value: "x".to_owned(),
                        },
                    ]
                );
                let mut context = SerializerContext::new(&SerializerDeterminism {
                    seed: "0".repeat(64),
                    now: "2000-01-01T00:00:00.000Z".to_owned(),
                })
                .unwrap();
                let serialize = if footnotes {
                    serialize_footnotes_part
                } else {
                    serialize_endnotes_part
                };
                let saved = serialize(&notes, &mut context).unwrap();
                if custom_xml {
                    let expected = format!(
                        r#"<w:{name} w:id="1" xmlns:x="urn:extension" mc:Ignorable="x">{content}</w:{name}>"#
                    );
                    assert_eq!(notes[0].verbatim_xml.as_deref(), Some(expected.as_str()));
                    assert!(saved.contains(&expected), "{saved}");
                } else {
                    assert!(notes[0].verbatim_xml.is_none());
                    assert!(saved.contains(&format!(
                        r#"<w:{name} w:id="1" xmlns:x="urn:extension" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14 x">"#
                    )), "{saved}");
                    assert!(saved.contains("<x:opaque/>"), "{saved}");
                }
            }
        }
    }

    #[test]
    fn note_budget_fails_before_allocating_over_the_limit() {
        let mut limits = ParseLimits::default();
        limits.max_notes = 1;
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(
            br#"<w:footnotes xmlns:w="w"><w:footnote/><w:footnote w:id="2"/></w:footnotes>"#,
            "word/footnotes.xml",
            &mut budget,
        )
        .unwrap();
        let media = MediaMap::new();
        let charts = ChartPartsMap::new();
        let mut smart_art = SmartArtContext::default();
        let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        let mut parser = StoryParser {
            relationships: None,
            theme: None,
            styles: None,
            doc_defaults: None,
            numbering: None,
            media: &media,
            charts: &charts,
            smart_art: &mut smart_art,
            budget: &mut budget,
            ids: &mut ids,
            part: "word/footnotes.xml",
        };
        assert_eq!(
            parse_notes(document.root().unwrap(), true, &mut parser),
            Err(crate::xml::ParseError::ResourceLimit {
                kind: "notes",
                part: "word/footnotes.xml".to_owned(),
            })
        );
    }
}
