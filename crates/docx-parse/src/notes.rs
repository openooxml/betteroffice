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
    #[serde(skip_serializing_if = "Option::is_none")]
    pub verbatim_xml: Option<String>,
}

impl Note {
    pub fn is_separator(&self) -> bool {
        matches!(
            self.note_type.as_str(),
            "separator" | "continuationSeparator" | "continuationNotice"
        )
    }
}

/// Parse a complete `w:footnotes` or `w:endnotes` story part. Modeled note
/// content always enters `StoryParser::parse_blocks`; the leaf only owns note
/// identity, special-note classification, and the incumbent verbatim gate.
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
        let note_type = parse_note_type(element.attribute(Some("w"), "type"));
        let content = parser.parse_blocks(element, 0, false)?;
        let verbatim_xml = has_unmodeled_direct_block(element).then(|| element.to_incumbent_xml());
        notes.push(Note {
            story_type: note_name.to_owned(),
            id: element
                .parse_numeric_attribute(Some("w"), "id", 1.0)
                .unwrap_or(0.0),
            note_type: note_type.to_owned(),
            content,
            verbatim_xml,
        });
    }
    Ok(notes)
}

pub fn parse_note_type(value: Option<&str>) -> &'static str {
    match value {
        Some("separator") => "separator",
        Some("continuationSeparator") => "continuationSeparator",
        Some("continuationNotice") => "continuationNotice",
        _ => "normal",
    }
}

fn has_unmodeled_direct_block(note: &XmlElement) -> bool {
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
    fn verbatim_gate_is_shallow_and_uses_incumbent_xml_serialization() {
        let notes = parse_story(
            r#"<w:endnotes xmlns:w="w"><w:endnote w:id="1" label="a&amp;b"><w:p/><w:bookmarkStart w:id="2"/></w:endnote><w:endnote w:id="2"><w:sdt><w:sdtContent><w:customXml><w:p/></w:customXml></w:sdtContent></w:sdt></w:endnote></w:endnotes>"#,
            false,
            &ParseLimits::default(),
        );
        assert!(notes[0].verbatim_xml.as_deref().unwrap().contains("a&b"));
        assert!(notes[1].verbatim_xml.is_none());
    }

    #[test]
    fn note_budget_fails_before_allocating_over_the_limit() {
        let mut limits = ParseLimits::default();
        limits.max_notes = 1;
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(
            br#"<w:footnotes xmlns:w="w"><w:footnote w:id="1"/><w:footnote w:id="2"/></w:footnotes>"#,
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
