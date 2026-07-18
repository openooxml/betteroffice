//! S10 `numbering.xml` serializer.

use crate::formatting::ParagraphFormatting;
use crate::numbering::{AbstractNumbering, ListLevel, NumberingDefinitions, NumberingInstance};

use super::xml_writer::{XmlWriter, int_attr};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// Serialize numbering definitions to a complete `word/numbering.xml` part.
pub fn serialize_numbering_xml(numbering: &NumberingDefinitions) -> String {
    let mut writer = XmlWriter::with_capacity(256 + numbering.abstract_nums.len() * 128);
    writer
        .declaration()
        .text("\n")
        .start_element("w:numbering")
        .attribute("xmlns:w", W_NS);
    for abstract_numbering in &numbering.abstract_nums {
        write_abstract_numbering(&mut writer, abstract_numbering);
    }
    for instance in &numbering.nums {
        write_numbering_instance(&mut writer, instance);
    }
    if numbering.abstract_nums.is_empty() && numbering.nums.is_empty() {
        // The incumbent template emits an explicit empty root pair.
        writer.text("");
    }
    writer.end_element();
    writer.finish()
}

fn write_level(writer: &mut XmlWriter, level: &ListLevel) {
    writer
        .start_element("w:lvl")
        .attribute("w:ilvl", &int_attr(Some(level.ilvl)));
    if let Some(value) = level.start {
        writer
            .start_element("w:start")
            .attribute("w:val", &int_attr(Some(value)))
            .end_element();
    }
    writer
        .start_element("w:numFmt")
        .attribute("w:val", &level.num_fmt)
        .end_element();
    if let Some(value) = level.suffix.as_deref().filter(|value| !value.is_empty()) {
        writer
            .start_element("w:suff")
            .attribute("w:val", value)
            .end_element();
    }
    writer
        .start_element("w:lvlText")
        .attribute("w:val", &level.lvl_text)
        .end_element();
    if let Some(value) = level.lvl_jc.as_deref().filter(|value| !value.is_empty()) {
        writer
            .start_element("w:lvlJc")
            .attribute("w:val", value)
            .end_element();
    }
    if let Some(properties) = level
        .p_pr
        .as_ref()
        .filter(|properties| has_indentation(properties))
    {
        writer.start_element("w:pPr");
        write_indentation(writer, properties);
        writer.end_element();
    }
    writer.end_element();
}

fn write_abstract_numbering(writer: &mut XmlWriter, numbering: &AbstractNumbering) {
    writer.start_element("w:abstractNum").attribute(
        "w:abstractNumId",
        &int_attr(Some(numbering.abstract_num_id)),
    );
    if let Some(value) = numbering
        .multi_level_type
        .as_deref()
        .filter(|value| !value.is_empty())
    {
        writer
            .start_element("w:multiLevelType")
            .attribute("w:val", value)
            .end_element();
    }
    let mut levels: Vec<_> = numbering.levels.iter().collect();
    levels.sort_by(|left, right| left.ilvl.total_cmp(&right.ilvl));
    for level in levels {
        write_level(writer, level);
    }
    // The TypeScript template always emits a start/end pair here.
    if numbering
        .multi_level_type
        .as_deref()
        .unwrap_or_default()
        .is_empty()
        && numbering.levels.is_empty()
    {
        writer.text("");
    }
    writer.end_element();
}

fn write_numbering_instance(writer: &mut XmlWriter, instance: &NumberingInstance) {
    writer
        .start_element("w:num")
        .attribute("w:numId", &int_attr(Some(instance.num_id)))
        .start_element("w:abstractNumId")
        .attribute("w:val", &int_attr(Some(instance.abstract_num_id)))
        .end_element();
    for level_override in instance.level_overrides.as_deref().unwrap_or_default() {
        writer
            .start_element("w:lvlOverride")
            .attribute("w:ilvl", &int_attr(Some(level_override.ilvl)));
        if let Some(value) = level_override.start_override {
            writer
                .start_element("w:startOverride")
                .attribute("w:val", &int_attr(Some(value)))
                .end_element();
        }
        if let Some(level) = level_override.lvl.as_ref() {
            write_level(writer, level);
        }
        if level_override.start_override.is_none() && level_override.lvl.is_none() {
            writer.text("");
        }
        writer.end_element();
    }
    writer.end_element();
}

fn has_indentation(properties: &ParagraphFormatting) -> bool {
    properties.indent_left.is_some()
        || properties.indent_right.is_some()
        || properties
            .indent_first_line
            .is_some_and(|value| properties.hanging_indent == Some(true) || value != 0.0)
}

fn write_indentation(writer: &mut XmlWriter, properties: &ParagraphFormatting) {
    writer.start_element("w:ind");
    if let Some(value) = properties.indent_left {
        writer.attribute("w:left", &int_attr(Some(value)));
    }
    if let Some(value) = properties.indent_right {
        writer.attribute("w:right", &int_attr(Some(value)));
    }
    if let Some(value) = properties.indent_first_line {
        if properties.hanging_indent == Some(true) {
            writer.attribute("w:hanging", &int_attr(Some(value.abs())));
        } else if value != 0.0 {
            writer.attribute("w:firstLine", &int_attr(Some(value)));
        }
    }
    writer.end_element();
}

#[cfg(test)]
mod tests {
    use crate::numbering::{LevelOverride, NumberingMap};
    use crate::xml::{ParseBudget, ParseLimits};

    use super::*;

    fn level(ilvl: f64, num_fmt: &str, level_text: &str) -> ListLevel {
        ListLevel {
            ilvl,
            start: None,
            num_fmt: num_fmt.to_owned(),
            lvl_text: level_text.to_owned(),
            lvl_jc: None,
            suffix: None,
            p_pr: None,
            r_pr: None,
            lvl_restart: None,
            is_lgl: None,
            legacy: None,
        }
    }

    #[test]
    fn matches_incumbent_order_and_indentation_bytes() {
        let mut first = level(1.0, "lowerLetter", "%2)");
        first.start = Some(2.5);
        first.suffix = Some("space".to_owned());
        first.lvl_jc = Some("right".to_owned());
        first.p_pr = Some(ParagraphFormatting {
            indent_left: Some(720.0),
            indent_right: Some(10.5),
            indent_first_line: Some(-360.0),
            hanging_indent: Some(true),
            ..ParagraphFormatting::default()
        });
        let numbering = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 4.0,
                multi_level_type: Some("multilevel".to_owned()),
                num_style_link: None,
                style_link: None,
                levels: vec![first, level(0.0, "decimal", "%1.")],
                name: None,
            }],
            nums: vec![NumberingInstance {
                num_id: 7.0,
                abstract_num_id: 4.0,
                level_overrides: Some(vec![LevelOverride {
                    ilvl: 1.0,
                    start_override: Some(3.5),
                    lvl: None,
                }]),
            }],
        };
        assert_eq!(
            serialize_numbering_xml(&numbering),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"4\"><w:multiLevelType w:val=\"multilevel\"/><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"3\"/><w:numFmt w:val=\"lowerLetter\"/><w:suff w:val=\"space\"/><w:lvlText w:val=\"%2)\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"720\" w:right=\"11\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=\"7\"><w:abstractNumId w:val=\"4\"/><w:lvlOverride w:ilvl=\"1\"><w:startOverride w:val=\"4\"/></w:lvlOverride></w:num></w:numbering>"
        );
    }

    #[test]
    fn preserves_explicit_empty_container_spelling() {
        assert_eq!(
            serialize_numbering_xml(&NumberingDefinitions::default()),
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:numbering>"
        );
        let numbering = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 0.0,
                multi_level_type: None,
                num_style_link: None,
                style_link: None,
                levels: Vec::new(),
                name: None,
            }],
            nums: vec![NumberingInstance {
                num_id: 1.0,
                abstract_num_id: 0.0,
                level_overrides: Some(vec![LevelOverride {
                    ilvl: 2.0,
                    start_override: None,
                    lvl: None,
                }]),
            }],
        };
        let xml = serialize_numbering_xml(&numbering);
        assert!(xml.contains("<w:abstractNum w:abstractNumId=\"0\"></w:abstractNum>"));
        assert!(xml.contains("<w:lvlOverride w:ilvl=\"2\"></w:lvlOverride>"));
    }

    #[test]
    fn escapes_all_numbering_strings() {
        let attack = "\"/><evil attr='&";
        let mut attacked_level = level(0.0, attack, attack);
        attacked_level.suffix = Some(attack.to_owned());
        attacked_level.lvl_jc = Some(attack.to_owned());
        let numbering = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 0.0,
                multi_level_type: Some(attack.to_owned()),
                num_style_link: None,
                style_link: None,
                levels: vec![attacked_level],
                name: None,
            }],
            nums: Vec::new(),
        };
        let xml = serialize_numbering_xml(&numbering);
        assert!(!xml.contains("<evil"));
        assert!(!xml.contains("attr='"));
        assert_eq!(
            xml.matches("&quot;/&gt;&lt;evil attr=&apos;&amp;").count(),
            5
        );
    }

    #[test]
    fn emitted_numbering_parse_backs_to_the_serialized_model() {
        let mut original_level = level(0.0, "decimal", "%1.");
        original_level.start = Some(2.0);
        original_level.p_pr = Some(ParagraphFormatting {
            indent_left: Some(720.0),
            indent_first_line: Some(-360.0),
            hanging_indent: Some(true),
            ..ParagraphFormatting::default()
        });
        let original = NumberingDefinitions {
            abstract_nums: vec![AbstractNumbering {
                abstract_num_id: 0.0,
                multi_level_type: Some("singleLevel".to_owned()),
                num_style_link: None,
                style_link: None,
                levels: vec![original_level],
                name: None,
            }],
            nums: vec![NumberingInstance {
                num_id: 1.0,
                abstract_num_id: 0.0,
                level_overrides: None,
            }],
        };
        let xml = serialize_numbering_xml(&original);
        let limits = ParseLimits::default();
        let NumberingMap { definitions } = crate::numbering::parse_numbering(
            Some(xml.as_bytes()),
            "word/numbering.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap();
        assert_eq!(definitions, original);
    }
}
