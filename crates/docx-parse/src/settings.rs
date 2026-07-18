//! Bounded parser for the document-wide `word/settings.xml` leaf contract.

use serde::{Deserialize, Serialize};

use crate::notes::{NoteProperties, parse_endnote_properties, parse_footnote_properties};
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

pub const DEFAULT_TAB_STOP_TWIPS: f64 = 720.0;
const MAX_TAB_STOP_TWIPS: f64 = 31_680.0;
const MS_WORD_COMPAT_SETTING_URI: &str = "http://schemas.microsoft.com/office/word";

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CompatibilityFlags {
    pub compatibility_mode: u8,
    pub no_leading: bool,
    pub do_not_expand_shift_return: bool,
    pub use_word97_line_break_rules: bool,
    pub balance_single_byte_double_byte_width: bool,
}

impl Default for CompatibilityFlags {
    fn default() -> Self {
        Self {
            compatibility_mode: 12,
            no_leading: false,
            do_not_expand_shift_return: false,
            use_word97_line_break_rules: false,
            balance_single_byte_double_byte_width: false,
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemeFontLanguage {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bidi: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RevisionView {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub markup: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comments: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub insertions_deletions: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formatting: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentSettings {
    pub default_tab_stop: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_table_style: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub theme_font_lang: Option<ThemeFontLanguage>,
    pub compatibility_flags: CompatibilityFlags,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub update_fields: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub track_revisions: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub do_not_track_moves: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub do_not_track_formatting: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision_view: Option<RevisionView>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<NoteProperties>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<NoteProperties>,
}

impl Default for DocumentSettings {
    fn default() -> Self {
        Self {
            default_tab_stop: DEFAULT_TAB_STOP_TWIPS,
            default_table_style: None,
            theme_font_lang: None,
            compatibility_flags: CompatibilityFlags::default(),
            update_fields: None,
            track_revisions: None,
            do_not_track_moves: None,
            do_not_track_formatting: None,
            revision_view: None,
            footnote_pr: None,
            endnote_pr: None,
        }
    }
}

pub fn parse_settings(
    xml: Option<&[u8]>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<DocumentSettings, ParseError> {
    let Some(xml) = xml.filter(|value| !value.is_empty()) else {
        return Ok(DocumentSettings::default());
    };
    if !incumbent_utf8_text_boundary(xml) {
        return Ok(DocumentSettings::default());
    }
    let document = parse_xml(xml, part, budget)?;
    Ok(parse_settings_element(document.root()))
}

pub(crate) fn incumbent_utf8_text_boundary(xml: &[u8]) -> bool {
    !xml.contains(&0) && std::str::from_utf8(xml).is_ok()
}

pub fn parse_settings_element(root: Option<&XmlElement>) -> DocumentSettings {
    let Some(root) = root else {
        return DocumentSettings::default();
    };
    let default_tab_stop = root
        .child("w", "defaultTabStop")
        .and_then(|element| element.parse_numeric_attribute(Some("w"), "val", 1.0))
        .filter(|value| *value > 0.0 && *value <= MAX_TAB_STOP_TWIPS)
        .unwrap_or(DEFAULT_TAB_STOP_TWIPS);
    let default_table_style = nonempty_attribute(root.child("w", "defaultTableStyle"), "val");
    let theme_font_lang = root.child("w", "themeFontLang").and_then(|element| {
        let language = ThemeFontLanguage {
            east_asia: nonempty(element.attribute(Some("w"), "eastAsia")),
            bidi: nonempty(element.attribute(Some("w"), "bidi")),
        };
        (language != ThemeFontLanguage::default()).then_some(language)
    });
    let update_fields = root.child("w", "updateFields").map(boolean_element);
    let track_revisions = root.child("w", "trackRevisions").map(boolean_element);
    let do_not_track_moves = root.child("w", "doNotTrackMoves").map(boolean_element);
    let do_not_track_formatting = root.child("w", "doNotTrackFormatting").map(boolean_element);
    let revision_view = root.child("w", "revisionView").map(|element| RevisionView {
        markup: boolean_attribute(element.attribute(Some("w"), "markup")),
        comments: boolean_attribute(element.attribute(Some("w"), "comments")),
        insertions_deletions: boolean_attribute(element.attribute(Some("w"), "insDel")),
        formatting: boolean_attribute(element.attribute(Some("w"), "formatting")),
    });
    let footnote_pr = root.child("w", "footnotePr").and_then(|element| {
        let properties = parse_footnote_properties(Some(element));
        (properties != NoteProperties::default()).then_some(properties)
    });
    let endnote_pr = root.child("w", "endnotePr").and_then(|element| {
        let properties = parse_endnote_properties(Some(element));
        (properties != NoteProperties::default()).then_some(properties)
    });

    DocumentSettings {
        default_tab_stop,
        default_table_style,
        theme_font_lang,
        compatibility_flags: parse_compatibility_flags(root),
        update_fields,
        track_revisions,
        do_not_track_moves,
        do_not_track_formatting,
        revision_view,
        footnote_pr,
        endnote_pr,
    }
}

fn parse_compatibility_flags(root: &XmlElement) -> CompatibilityFlags {
    let Some(compatibility) = root.child("w", "compat") else {
        return CompatibilityFlags::default();
    };
    let mut flags = CompatibilityFlags {
        no_leading: compatibility
            .child("w", "noLeading")
            .is_some_and(boolean_element),
        do_not_expand_shift_return: compatibility
            .child("w", "doNotExpandShiftReturn")
            .is_some_and(boolean_element),
        use_word97_line_break_rules: compatibility
            .child("w", "useWord97LineBreakRules")
            .is_some_and(boolean_element),
        balance_single_byte_double_byte_width: compatibility
            .child("w", "balanceSingleByteDoubleByteWidth")
            .is_some_and(boolean_element),
        ..CompatibilityFlags::default()
    };
    for setting in compatibility.children_named("w", "compatSetting") {
        if setting.attribute(Some("w"), "name") != Some("compatibilityMode")
            || setting.attribute(Some("w"), "uri") != Some(MS_WORD_COMPAT_SETTING_URI)
        {
            continue;
        }
        flags.compatibility_mode = clamp_compatibility_mode(setting.attribute(Some("w"), "val"));
        break;
    }
    flags
}

fn clamp_compatibility_mode(raw: Option<&str>) -> u8 {
    let Some(value) = raw.and_then(crate::xml::parse_javascript_integer_prefix) else {
        return 12;
    };
    if value.fract() != 0.0 || value < 11.0 {
        12
    } else {
        value.min(15.0) as u8
    }
}

fn boolean_element(element: &XmlElement) -> bool {
    !matches!(
        element.attribute(Some("w"), "val"),
        Some("0" | "false" | "off")
    )
}

fn boolean_attribute(value: Option<&str>) -> Option<bool> {
    value.map(|value| !matches!(value.to_ascii_lowercase().as_str(), "0" | "false" | "off"))
}

fn nonempty_attribute(element: Option<&XmlElement>, name: &str) -> Option<String> {
    nonempty(element.and_then(|element| element.attribute(Some("w"), name)))
}

fn nonempty(value: Option<&str>) -> Option<String> {
    value.filter(|value| !value.is_empty()).map(str::to_owned)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::ParseLimits;

    fn parse(xml: Option<&str>) -> DocumentSettings {
        let limits = ParseLimits::default();
        parse_settings(
            xml.map(str::as_bytes),
            "word/settings.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
    }

    #[test]
    fn defaults_and_clamps_attacker_controlled_numbers() {
        assert_eq!(parse(None), DocumentSettings::default());
        for value in ["0", "-1", "31681", "garbage", &"9".repeat(10_000)] {
            let parsed = parse(Some(&format!(
                r#"<w:settings><w:defaultTabStop w:val="{value}"/></w:settings>"#
            )));
            assert_eq!(parsed.default_tab_stop, DEFAULT_TAB_STOP_TWIPS);
        }
        assert_eq!(
            parse(Some(
                r#"<w:settings><w:defaultTabStop w:val="567twips"/></w:settings>"#
            ))
            .default_tab_stop,
            567.0
        );
    }

    #[test]
    fn parses_settings_and_boolean_compatibility_quirks() {
        let parsed = parse(Some(
            r#"<w:settings><w:defaultTableStyle w:val="Grid"/><w:themeFontLang w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:updateFields w:val="FALSE"/><w:trackRevisions w:val="false"/><w:revisionView w:markup="FALSE"/><w:compat><w:noLeading/><w:doNotExpandShiftReturn w:val="off"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="9999future"/></w:compat></w:settings>"#,
        ));
        assert_eq!(parsed.default_table_style.as_deref(), Some("Grid"));
        assert_eq!(
            parsed.theme_font_lang.unwrap().east_asia.as_deref(),
            Some("ja-JP")
        );
        assert_eq!(parsed.update_fields, Some(true));
        assert_eq!(parsed.track_revisions, Some(false));
        assert_eq!(parsed.revision_view.unwrap().markup, Some(false));
        assert_eq!(parsed.compatibility_flags.compatibility_mode, 15);
        assert!(parsed.compatibility_flags.no_leading);
        assert!(!parsed.compatibility_flags.do_not_expand_shift_return);
    }

    #[test]
    fn hostile_xml_is_rejected_by_the_shared_safe_core() {
        let limits = crate::xml::ParseLimits::default();
        let error = parse_settings(
            Some(b"<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]><w:settings>&e;</w:settings>"),
            "word/settings.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap_err();
        assert!(matches!(error, ParseError::UnsafeXml { .. }));
    }
}
