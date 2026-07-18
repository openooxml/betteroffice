//! Runs and inline structures with incumbent-compatible field and hyperlink behavior.
//!
//! Field instructions are parsed as inert metadata only. This module has no
//! resolver for DDE, INCLUDE*, OLE, macros, files, processes, or networks.
//! Drawing-owned children retain their source position as stable opaque atoms
//! until the S4 drawing contract is merged.

use serde::{Deserialize, Serialize};

use crate::formatting::{
    TextFill, TextFormatting, TextGlow, TextGradientStop, TextModernEffects, TextOutline,
    TextReflection, TextShadow, merge_text_formatting, parse_run_properties,
};
use crate::relationships::RelationshipMap;
use crate::styles::{DocDefaults, StyleMap, get_default_character_style};
use crate::theme::{Theme, get_theme_color};
use crate::xml::{ParseBudget, ParseError, XmlElement, XmlNode};

const MAX_FIELD_NESTING: usize = 32;
const MAX_HYPERLINK_CHILDREN: usize = 10_000;
const MAX_SIMPLE_FIELD_NESTING: usize = 32;
const MAX_FORM_LIST_ENTRIES: usize = 1_000;
const MAX_SDT_LIST_ITEMS: usize = 10_000;
const MAX_GRADIENT_STOPS: usize = 16;
const W14_ANGLE_UNITS_PER_DEGREE: f64 = 60_000.0;
const W14_PERCENT_SCALE: f64 = 100_000.0;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Run {
    #[serde(rename = "type")]
    pub node_type: RunType,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formatting: Option<TextFormatting>,
    #[serde(rename = "propertyChanges", skip_serializing_if = "Option::is_none")]
    pub property_changes: Option<Vec<RunPropertyChange>>,
    pub content: Vec<RunContent>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum RunType {
    #[serde(rename = "run")]
    Run,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum RunContent {
    #[serde(rename = "text")]
    Text {
        text: String,
        #[serde(rename = "preserveSpace", skip_serializing_if = "Option::is_none")]
        preserve_space: Option<bool>,
    },
    #[serde(rename = "tab")]
    Tab,
    #[serde(rename = "break")]
    Break {
        #[serde(rename = "breakType", skip_serializing_if = "Option::is_none")]
        break_type: Option<String>,
        #[serde(skip_serializing_if = "Option::is_none")]
        clear: Option<String>,
    },
    #[serde(rename = "symbol")]
    Symbol { font: String, char: String },
    #[serde(rename = "footnoteRef")]
    FootnoteRef {
        id: f64,
        #[serde(rename = "customMarkFollows", skip_serializing_if = "Option::is_none")]
        custom_mark_follows: Option<bool>,
    },
    #[serde(rename = "endnoteRef")]
    EndnoteRef {
        id: f64,
        #[serde(rename = "customMarkFollows", skip_serializing_if = "Option::is_none")]
        custom_mark_follows: Option<bool>,
    },
    #[serde(rename = "footnoteRefMark")]
    FootnoteRefMark,
    #[serde(rename = "endnoteRefMark")]
    EndnoteRefMark,
    #[serde(rename = "separator")]
    Separator,
    #[serde(rename = "continuationSeparator")]
    ContinuationSeparator,
    #[serde(rename = "commentReference")]
    CommentReference {
        #[serde(skip_serializing_if = "Option::is_none")]
        id: Option<f64>,
    },
    #[serde(rename = "fieldChar")]
    FieldChar {
        #[serde(rename = "charType")]
        char_type: String,
        #[serde(rename = "fldLock", skip_serializing_if = "Option::is_none")]
        fld_lock: Option<bool>,
        #[serde(skip_serializing_if = "Option::is_none")]
        dirty: Option<bool>,
        #[serde(rename = "formData", skip_serializing_if = "Option::is_none")]
        form_data: Option<FieldFormData>,
    },
    #[serde(rename = "instrText")]
    InstrText { text: String },
    #[serde(rename = "softHyphen")]
    SoftHyphen,
    #[serde(rename = "noBreakHyphen")]
    NoBreakHyphen,
    #[serde(rename = "drawing")]
    Drawing { image: Box<crate::image::Image> },
    #[serde(rename = "shape")]
    Shape { shape: Box<crate::shape::Shape> },
    #[serde(rename = "chart")]
    Chart { chart: Box<crate::chart::Chart> },
    #[serde(rename = "opaqueDrawing")]
    OpaqueDrawing { kind: String },
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FieldFormData {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub enabled: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub help_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub status_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub entry_macro: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub exit_macro: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub calculate_on_exit: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub control_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checked: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default_checked: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checkbox_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_entries: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selected_index: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct RunPropertyChange {
    #[serde(rename = "type")]
    pub node_type: RunPropertyChangeType,
    pub info: PropertyChangeInfo,
    #[serde(rename = "previousFormatting", skip_serializing_if = "Option::is_none")]
    pub previous_formatting: Option<TextFormatting>,
    #[serde(rename = "currentFormatting", skip_serializing_if = "Option::is_none")]
    pub current_formatting: Option<TextFormatting>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum RunPropertyChangeType {
    #[serde(rename = "runPropertyChange")]
    RunPropertyChange,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct PropertyChangeInfo {
    pub id: f64,
    pub author: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rsid: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunProjection {
    pub run: Run,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub resolved_formatting: Option<TextFormatting>,
}

pub fn parse_run(
    element: &XmlElement,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
) -> RunProjection {
    let r_pr = element.child("w", "rPr");
    let formatting = parse_inline_run_properties(r_pr, theme);
    let property_changes = parse_run_property_changes(r_pr, theme, formatting.as_ref());
    let run = Run {
        node_type: RunType::Run,
        formatting,
        property_changes,
        content: parse_run_contents(element),
    };
    let resolved_formatting = resolve_run_formatting(&run, styles, doc_defaults);
    RunProjection {
        run,
        resolved_formatting,
    }
}

fn parse_inline_run_properties(
    r_pr: Option<&XmlElement>,
    theme: Option<&Theme>,
) -> Option<TextFormatting> {
    let r_pr = r_pr?;
    let mut formatting = parse_run_properties(Some(r_pr), theme).unwrap_or_default();
    formatting.modern_effects = parse_modern_text_effects(r_pr, theme);
    (formatting != TextFormatting::default() || r_pr.child("w", "shd").is_some())
        .then_some(formatting)
}

fn resolve_run_formatting(
    run: &Run,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
) -> Option<TextFormatting> {
    let mut resolved = doc_defaults.and_then(|defaults| defaults.r_pr.clone());
    if let Some(styles) = styles {
        resolved = merge_text_formatting(
            resolved.as_ref(),
            get_default_character_style(styles).and_then(|style| style.r_pr.as_ref()),
        );
        if let Some(style_id) = run
            .formatting
            .as_ref()
            .and_then(|formatting| formatting.style_id.as_deref())
        {
            resolved = merge_text_formatting(
                resolved.as_ref(),
                styles.get(style_id).and_then(|style| style.r_pr.as_ref()),
            );
        }
    }
    merge_text_formatting(resolved.as_ref(), run.formatting.as_ref())
}

fn parse_run_property_changes(
    r_pr: Option<&XmlElement>,
    theme: Option<&Theme>,
    current: Option<&TextFormatting>,
) -> Option<Vec<RunPropertyChange>> {
    let changes: Vec<_> = r_pr?
        .children_named("w", "rPrChange")
        .filter_map(|change| {
            let previous = parse_inline_run_properties(change.child("w", "rPr"), theme);
            let current = current.cloned();
            (previous.is_some() || current.is_some()).then(|| RunPropertyChange {
                node_type: RunPropertyChangeType::RunPropertyChange,
                info: parse_property_change_info(change),
                previous_formatting: previous,
                current_formatting: current,
            })
        })
        .collect();
    (!changes.is_empty()).then_some(changes)
}

fn parse_property_change_info(element: &XmlElement) -> PropertyChangeInfo {
    let id = element
        .parse_numeric_attribute(Some("w"), "id", 1.0)
        .filter(|value| value.is_finite() && value.fract() == 0.0 && *value >= 0.0)
        .unwrap_or(0.0);
    let author = trimmed_nonempty(element.attribute(Some("w"), "author"))
        .unwrap_or_else(|| "Unknown".to_owned());
    PropertyChangeInfo {
        id,
        author,
        date: trimmed_nonempty(element.attribute(Some("w"), "date")),
        rsid: trimmed_nonempty(element.attribute(Some("w"), "rsid")),
    }
}

fn parse_run_contents(element: &XmlElement) -> Vec<RunContent> {
    let mut output = Vec::new();
    for child in element.child_elements() {
        match child.local_name() {
            "t" => output.push(RunContent::Text {
                text: incumbent_text_content(child),
                preserve_space: (child.attribute(Some("xml"), "space") == Some("preserve"))
                    .then_some(true),
            }),
            "tab" => output.push(RunContent::Tab),
            "br" => output.push(parse_break(child)),
            "sym" => output.push(RunContent::Symbol {
                font: child
                    .attribute(Some("w"), "font")
                    .unwrap_or_default()
                    .to_owned(),
                char: child
                    .attribute(Some("w"), "char")
                    .unwrap_or_default()
                    .to_owned(),
            }),
            "footnoteReference" => output.push(parse_note_reference(child, false)),
            "endnoteReference" => output.push(parse_note_reference(child, true)),
            "fldChar" => output.push(parse_field_char(child)),
            "instrText" => output.push(RunContent::InstrText {
                text: incumbent_text_content(child),
            }),
            "commentReference" => output.push(RunContent::CommentReference {
                id: child.parse_numeric_attribute(Some("w"), "id", 1.0),
            }),
            "softHyphen" => output.push(RunContent::SoftHyphen),
            "noBreakHyphen" => output.push(RunContent::NoBreakHyphen),
            "drawing" | "pict" | "object" => output.push(RunContent::OpaqueDrawing {
                kind: child.local_name().to_owned(),
            }),
            "AlternateContent" if contains_drawing_owned_content(child) => {
                output.push(RunContent::OpaqueDrawing {
                    kind: "alternateContent".to_owned(),
                });
            }
            "cr" => output.push(RunContent::Break {
                break_type: Some("textWrapping".to_owned()),
                clear: None,
            }),
            "footnoteRef" => output.push(RunContent::FootnoteRefMark),
            "endnoteRef" => output.push(RunContent::EndnoteRefMark),
            "separator" => output.push(RunContent::Separator),
            "continuationSeparator" => output.push(RunContent::ContinuationSeparator),
            _ => {}
        }
    }
    output
}

fn parse_break(element: &XmlElement) -> RunContent {
    let break_type = element
        .attribute(Some("w"), "type")
        .filter(|value| matches!(*value, "page" | "column" | "textWrapping"))
        .map(str::to_owned);
    let clear = element
        .attribute(Some("w"), "clear")
        .filter(|value| matches!(*value, "none" | "left" | "right" | "all"))
        .map(str::to_owned);
    RunContent::Break { break_type, clear }
}

fn parse_note_reference(element: &XmlElement, endnote: bool) -> RunContent {
    let id = element
        .parse_numeric_attribute(Some("w"), "id", 1.0)
        .unwrap_or(0.0);
    let custom_mark_follows = element
        .attribute(Some("w"), "customMarkFollows")
        .filter(|raw| !matches_ci(raw, &["0", "false", "off"]))
        .map(|_| true);
    if endnote {
        RunContent::EndnoteRef {
            id,
            custom_mark_follows,
        }
    } else {
        RunContent::FootnoteRef {
            id,
            custom_mark_follows,
        }
    }
}

fn parse_field_char(element: &XmlElement) -> RunContent {
    let char_type = match element.attribute(Some("w"), "fldCharType") {
        Some("separate") => "separate",
        Some("end") => "end",
        _ => "begin",
    }
    .to_owned();
    RunContent::FieldChar {
        char_type,
        fld_lock: matches!(element.attribute(Some("w"), "fldLock"), Some("true" | "1"))
            .then_some(true),
        dirty: matches!(element.attribute(Some("w"), "dirty"), Some("true" | "1")).then_some(true),
        form_data: parse_field_form_data(element),
    }
}

fn parse_field_form_data(element: &XmlElement) -> Option<FieldFormData> {
    let data = element.child("w", "ffData")?;
    let mut value = FieldFormData::default();
    value.name = child_value(data, "name").map(|value| truncate_chars(value, 255));
    value.help_text = child_value(data, "helpText").map(|value| truncate_chars(value, 255));
    value.status_text = child_value(data, "statusText").map(|value| truncate_chars(value, 255));
    value.entry_macro = child_value(data, "entryMacro").map(|value| truncate_chars(value, 255));
    value.exit_macro = child_value(data, "exitMacro").map(|value| truncate_chars(value, 255));
    value.enabled = data
        .child("w", "enabled")
        .map(|child| child.parse_boolean("w"));
    value.calculate_on_exit = data
        .child("w", "calcOnExit")
        .map(|child| child.parse_boolean("w"));

    if let Some(text) = data.child("w", "textInput") {
        value.control_type = Some("text".to_owned());
        value.default_value =
            child_value(text, "default").map(|value| truncate_chars(value, 32_767));
        value.text_type = child_value(text, "type").map(|value| truncate_chars(value, 63));
        value.text_format = child_value(text, "format").map(|value| truncate_chars(value, 255));
    } else if let Some(checkbox) = data.child("w", "checkBox") {
        value.control_type = Some("checkbox".to_owned());
        value.checked = checkbox
            .child("w", "checked")
            .map(|child| child.parse_boolean("w"));
        value.default_checked = checkbox
            .child("w", "default")
            .map(|child| child.parse_boolean("w"));
        value.checkbox_size = child_value(checkbox, "size")
            .and_then(|raw| raw.trim().parse::<f64>().ok())
            .filter(|size| size.is_finite() && *size >= 0.0 && *size <= 3276.0);
    } else if let Some(list) = data.child("w", "ddList") {
        value.control_type = Some("dropDown".to_owned());
        let entries: Vec<_> = list
            .children_named("w", "listEntry")
            .take(MAX_FORM_LIST_ENTRIES)
            .map(|entry| truncate_chars(entry.attribute(Some("w"), "val").unwrap_or_default(), 255))
            .collect();
        value.list_entries = (!entries.is_empty()).then_some(entries);
        value.selected_index = child_value(list, "result")
            .and_then(|raw| raw.trim().parse::<f64>().ok())
            .filter(|index| {
                index.is_finite()
                    && index.fract() == 0.0
                    && *index >= 0.0
                    && *index < MAX_FORM_LIST_ENTRIES as f64
            });
    }

    (value != FieldFormData::default()).then_some(value)
}

fn child_value<'a>(parent: &'a XmlElement, local: &str) -> Option<&'a str> {
    parent.child("w", local)?.attribute(Some("w"), "val")
}

fn contains_drawing_owned_content(element: &XmlElement) -> bool {
    element.child_elements().any(|child| {
        matches!(child.local_name(), "drawing" | "pict" | "object")
            || contains_drawing_owned_content(child)
    })
}

fn incumbent_text_content(element: &XmlElement) -> String {
    let mut output = String::new();
    append_incumbent_text(element, &mut output);
    output
}

fn append_incumbent_text(element: &XmlElement, output: &mut String) {
    for child in &element.children {
        match child {
            XmlNode::Text(text) => output.push_str(text),
            XmlNode::Element(child) => append_incumbent_text(child, output),
            XmlNode::CData(_) => {}
        }
    }
}

fn parse_modern_text_effects(
    r_pr: &XmlElement,
    theme: Option<&Theme>,
) -> Option<TextModernEffects> {
    let mut effects = TextModernEffects::default();
    if let Some(glow) = r_pr.child("w14", "glow") {
        effects.glow = Some(TextGlow {
            color: parse_w14_color(glow, theme),
            radius: numeric(glow, "rad").map(emu_to_pixels),
        });
    }
    if let Some(shadow) = r_pr.child("w14", "shadow") {
        effects.shadow = Some(TextShadow {
            color: parse_w14_color(shadow, theme),
            blur_radius: numeric(shadow, "blurRad").map(emu_to_pixels),
            distance: numeric(shadow, "dist").map(emu_to_pixels),
            direction: numeric(shadow, "dir").map(|value| value / W14_ANGLE_UNITS_PER_DEGREE),
        });
    }
    if let Some(reflection) = r_pr.child("w14", "reflection") {
        effects.reflection = Some(TextReflection {
            blur_radius: numeric(reflection, "blurRad").map(emu_to_pixels),
            start_opacity: numeric(reflection, "stA")
                .map(|value| (value / W14_PERCENT_SCALE).clamp(0.0, 1.0)),
            end_opacity: numeric(reflection, "endA")
                .map(|value| (value / W14_PERCENT_SCALE).clamp(0.0, 1.0)),
            distance: numeric(reflection, "dist").map(emu_to_pixels),
            direction: numeric(reflection, "dir").map(|value| value / W14_ANGLE_UNITS_PER_DEGREE),
        });
    }
    effects.text_fill = r_pr
        .child("w14", "textFill")
        .and_then(|fill| parse_w14_fill(fill, theme));
    if let Some(outline) = r_pr.child("w14", "textOutline") {
        let fill = parse_w14_fill(outline, theme);
        let dash = outline
            .child_by_local_name("prstDash")
            .and_then(|dash| dash.attribute(Some("w14"), "val"))
            .map(|value| truncate_chars(value, 32));
        effects.text_outline = Some(TextOutline {
            color: fill
                .as_ref()
                .filter(|fill| fill.kind.as_deref() == Some("solid"))
                .and_then(|fill| fill.color.clone()),
            width: numeric(outline, "w").map(emu_to_pixels),
            dash,
            no_fill: fill
                .as_ref()
                .filter(|fill| fill.kind.as_deref() == Some("none"))
                .map(|_| true),
        });
    }
    (effects != TextModernEffects::default()).then_some(effects)
}

fn parse_w14_fill(element: &XmlElement, theme: Option<&Theme>) -> Option<TextFill> {
    if element.child_by_local_name("noFill").is_some() {
        return Some(TextFill {
            kind: Some("none".to_owned()),
            ..TextFill::default()
        });
    }
    if let Some(solid) = element.child_by_local_name("solidFill") {
        return Some(TextFill {
            kind: Some("solid".to_owned()),
            color: parse_w14_color(solid, theme),
            ..TextFill::default()
        });
    }
    let gradient = element.child_by_local_name("gradFill")?;
    let stops: Vec<_> = gradient
        .child_by_local_name("gsLst")
        .into_iter()
        .flat_map(XmlElement::child_elements)
        .filter(|stop| stop.local_name() == "gs")
        .take(MAX_GRADIENT_STOPS)
        .map(|stop| TextGradientStop {
            position: numeric(stop, "pos").map(|value| (value / W14_PERCENT_SCALE).clamp(0.0, 1.0)),
            color: parse_w14_color(stop, theme),
        })
        .collect();
    let angle = gradient
        .child_by_local_name("lin")
        .and_then(|line| numeric(line, "ang"))
        .map(|value| value / W14_ANGLE_UNITS_PER_DEGREE);
    Some(TextFill {
        kind: Some("gradient".to_owned()),
        color: None,
        angle,
        stops: (!stops.is_empty()).then_some(stops),
    })
}

fn parse_w14_color(element: &XmlElement, theme: Option<&Theme>) -> Option<String> {
    for child in element.child_elements() {
        let Some(value) = child.attribute(Some("w14"), "val") else {
            continue;
        };
        match child.local_name() {
            "srgbClr" if value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()) => {
                return Some(format!("#{}", value.to_ascii_lowercase()));
            }
            "schemeClr" => {
                let slot = match value {
                    "tx1" => "dk1",
                    "tx2" => "dk2",
                    "bg1" => "lt1",
                    "bg2" => "lt2",
                    "accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6"
                    | "dk1" | "dk2" | "lt1" | "lt2" | "hlink" | "folHlink" => value,
                    _ => continue,
                };
                return Some(format!("#{}", get_theme_color(theme, slot)));
            }
            _ => {}
        }
    }
    None
}

fn numeric(element: &XmlElement, name: &str) -> Option<f64> {
    element.parse_numeric_attribute(Some("w14"), name, 1.0)
}

fn emu_to_pixels(value: f64) -> f64 {
    (value * 96.0 / 914_400.0).round()
}

fn truncate_chars(value: &str, limit: usize) -> String {
    value.chars().take(limit).collect()
}

fn trimmed_nonempty(value: Option<&str>) -> Option<String> {
    let value = value?.trim();
    (!value.is_empty()).then(|| value.to_owned())
}

fn matches_ci(value: &str, candidates: &[&str]) -> bool {
    candidates
        .iter()
        .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct FieldSwitch {
    #[serde(rename = "switch")]
    pub switch_name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct ParsedFieldInstruction {
    #[serde(rename = "type")]
    pub field_type: String,
    pub raw: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub argument: Option<String>,
    pub switches: Vec<FieldSwitch>,
}

const KNOWN_FIELD_TYPES: &[&str] = &[
    "PAGE",
    "NUMPAGES",
    "NUMWORDS",
    "NUMCHARS",
    "DATE",
    "TIME",
    "CREATEDATE",
    "SAVEDATE",
    "PRINTDATE",
    "EDITTIME",
    "AUTHOR",
    "TITLE",
    "SUBJECT",
    "KEYWORDS",
    "COMMENTS",
    "FILENAME",
    "FILESIZE",
    "TEMPLATE",
    "REVNUM",
    "DOCPROPERTY",
    "DOCVARIABLE",
    "REF",
    "PAGEREF",
    "NOTEREF",
    "HYPERLINK",
    "TOC",
    "TOA",
    "INDEX",
    "SEQ",
    "STYLEREF",
    "AUTONUM",
    "AUTONUMLGL",
    "AUTONUMOUT",
    "SECTION",
    "SECTIONPAGES",
    "USERADDRESS",
    "USERNAME",
    "USERINITIALS",
    "FORMTEXT",
    "FORMCHECKBOX",
    "FORMDROPDOWN",
    "CITATION",
    "BIBLIOGRAPHY",
    "IF",
    "MERGEFIELD",
    "NEXT",
    "NEXTIF",
    "ASK",
    "SET",
    "QUOTE",
    "INCLUDETEXT",
    "INCLUDEPICTURE",
    "SYMBOL",
    "ADVANCE",
];

pub fn parse_field_type(instruction: &str) -> String {
    let mut input = instruction.trim_start();
    if input.is_empty() {
        return "UNKNOWN".to_owned();
    }
    if let Some(rest) = input.strip_prefix('\\') {
        input = rest;
    }
    let mut end = 0usize;
    for (index, character) in input.char_indices() {
        let valid = if index == 0 {
            character.is_ascii_alphabetic()
        } else {
            character.is_ascii_alphanumeric()
        };
        if !valid {
            break;
        }
        end = index + character.len_utf8();
    }
    if end == 0 {
        return "UNKNOWN".to_owned();
    }
    let field_type = input[..end].to_ascii_uppercase();
    KNOWN_FIELD_TYPES
        .contains(&field_type.as_str())
        .then_some(field_type)
        .unwrap_or_else(|| "UNKNOWN".to_owned())
}

/// Bounded linear equivalent of the incumbent field-instruction regex.
pub fn parse_field_instruction(instruction: &str) -> ParsedFieldInstruction {
    let field_type = parse_field_type(instruction);
    let trimmed = instruction.trim();
    let field_name_end = field_name_end(trimmed);
    let remaining = trimmed[field_name_end..].trim();
    let mut switches = Vec::new();
    let mut first_switch = None;
    let mut cursor = 0usize;
    while cursor < remaining.len() {
        let Some(relative) = remaining[cursor..].find('\\') else {
            break;
        };
        let start = cursor + relative;
        let after_slash = start + 1;
        let Some(character) = remaining[after_slash..].chars().next() else {
            break;
        };
        if !matches!(character, '*' | '@' | '#' | '!') && !character.is_ascii_alphabetic() {
            cursor = after_slash;
            continue;
        }
        first_switch.get_or_insert(start);
        let mut end = after_slash + character.len_utf8();
        end = consume_whitespace(remaining, end);
        let mut value = None;
        if remaining[end..].starts_with('"') {
            let quoted_start = end + 1;
            if let Some(closing) = remaining[quoted_start..].find('"') {
                let quoted_end = quoted_start + closing;
                if quoted_end > quoted_start {
                    value = Some(remaining[quoted_start..quoted_end].to_owned());
                }
                end = quoted_end + 1;
            } else {
                let token_end = consume_non_whitespace(remaining, end);
                if token_end > end {
                    value = Some(remaining[end..token_end].to_owned());
                }
                end = token_end;
            }
        } else {
            let token_end = consume_non_whitespace(remaining, end);
            if token_end > end {
                value = Some(remaining[end..token_end].to_owned());
            }
            end = token_end;
        }
        switches.push(FieldSwitch {
            switch_name: character.to_string(),
            value,
        });
        cursor = end.max(after_slash);
    }

    let argument_source = &remaining[..first_switch.unwrap_or(remaining.len())];
    let argument_source = argument_source.trim();
    let argument = if argument_source.starts_with('"')
        && argument_source.ends_with('"')
        && argument_source.len() >= 2
    {
        Some(argument_source[1..argument_source.len() - 1].to_owned())
    } else if argument_source.is_empty() {
        None
    } else {
        Some(argument_source.to_owned())
    };
    ParsedFieldInstruction {
        field_type,
        raw: instruction.to_owned(),
        argument,
        switches,
    }
}

fn field_name_end(value: &str) -> usize {
    let base = usize::from(value.starts_with('\\'));
    let mut cursor = base;
    let mut first = true;
    for (offset, character) in value[base..].char_indices() {
        let valid = if first {
            character.is_ascii_alphabetic()
        } else {
            character.is_ascii_alphanumeric()
        };
        if !valid {
            break;
        }
        first = false;
        cursor = base + offset + character.len_utf8();
    }
    if first { 0 } else { cursor }
}

fn consume_whitespace(value: &str, mut cursor: usize) -> usize {
    while let Some(character) = value[cursor..].chars().next() {
        if !character.is_whitespace() {
            break;
        }
        cursor += character.len_utf8();
    }
    cursor
}

fn consume_non_whitespace(value: &str, mut cursor: usize) -> usize {
    while let Some(character) = value[cursor..].chars().next() {
        if character.is_whitespace() {
            break;
        }
        cursor += character.len_utf8();
    }
    cursor
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum InlineNode {
    Run(Run),
    Hyperlink(Box<Hyperlink>),
    BookmarkStart(BookmarkStart),
    BookmarkEnd(BookmarkEnd),
    SimpleField(Box<SimpleField>),
    ComplexField(Box<ComplexField>),
    InlineSdt(Box<InlineSdt>),
    Math(MathEquation),
}

impl InlineNode {
    pub fn node_type(&self) -> &'static str {
        match self {
            Self::Run(_) => "run",
            Self::Hyperlink(_) => "hyperlink",
            Self::BookmarkStart(_) => "bookmarkStart",
            Self::BookmarkEnd(_) => "bookmarkEnd",
            Self::SimpleField(_) => "simpleField",
            Self::ComplexField(_) => "complexField",
            Self::InlineSdt(_) => "inlineSdt",
            Self::Math(_) => "mathEquation",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct SimpleField {
    #[serde(rename = "type")]
    pub node_type: SimpleFieldType,
    pub instruction: String,
    #[serde(rename = "fieldType")]
    pub field_type: String,
    pub content: Vec<Run>,
    #[serde(rename = "fldLock", skip_serializing_if = "Option::is_none")]
    pub fld_lock: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dirty: Option<bool>,
    #[serde(rename = "structuredResult", skip_serializing_if = "Option::is_none")]
    pub structured_result: Option<StructuredFieldContent>,
    #[serde(rename = "fieldTree", skip_serializing_if = "Option::is_none")]
    pub field_tree: Option<StructuredFieldTree>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum SimpleFieldType {
    #[serde(rename = "simpleField")]
    SimpleField,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct StructuredFieldContent {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inline: Option<Vec<InlineNode>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blocks: Option<Vec<crate::block::BlockContent>>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StructuredFieldTree {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub version: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub code: Option<StructuredFieldContent>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub result: Option<StructuredFieldContent>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub children: Option<Vec<StructuredFieldTree>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub display_mode: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ComplexField {
    #[serde(rename = "type")]
    pub node_type: ComplexFieldType,
    pub instruction: String,
    #[serde(rename = "fieldType")]
    pub field_type: String,
    #[serde(rename = "fieldCode")]
    pub field_code: Vec<Run>,
    #[serde(rename = "fieldResult")]
    pub field_result: Vec<Run>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formatting: Option<TextFormatting>,
    #[serde(rename = "fldLock", skip_serializing_if = "Option::is_none")]
    pub fld_lock: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dirty: Option<bool>,
    #[serde(rename = "structuredCode", skip_serializing_if = "Option::is_none")]
    pub structured_code: Option<StructuredFieldContent>,
    #[serde(rename = "structuredResult", skip_serializing_if = "Option::is_none")]
    pub structured_result: Option<StructuredFieldContent>,
    #[serde(rename = "fieldTree", skip_serializing_if = "Option::is_none")]
    pub field_tree: Option<StructuredFieldTree>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum ComplexFieldType {
    #[serde(rename = "complexField")]
    ComplexField,
}

pub fn parse_simple_field(
    element: &XmlElement,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &ParseBudget<'_>,
) -> Result<SimpleField, ParseError> {
    parse_simple_field_at_depth(element, theme, styles, doc_defaults, part, budget, 0)
}

fn parse_simple_field_at_depth(
    element: &XmlElement,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &ParseBudget<'_>,
    depth: usize,
) -> Result<SimpleField, ParseError> {
    if depth > MAX_SIMPLE_FIELD_NESTING {
        return Err(ParseError::ResourceLimit {
            kind: "fieldDepth",
            part: part.to_owned(),
        });
    }
    budget.check_nesting_depth(depth, part)?;
    let instruction = element
        .attribute(Some("w"), "instr")
        .unwrap_or_default()
        .to_owned();
    let content: Vec<_> = element
        .children_named("w", "r")
        .map(|run| parse_run(run, theme, styles, doc_defaults).run)
        .collect();
    let mut structured: Vec<_> = content.iter().cloned().map(InlineNode::Run).collect();
    for nested in element
        .children_named("w", "fldSimple")
        .take(MAX_SIMPLE_FIELD_NESTING)
    {
        structured.push(InlineNode::SimpleField(Box::new(
            parse_simple_field_at_depth(
                nested,
                theme,
                styles,
                doc_defaults,
                part,
                budget,
                depth + 1,
            )?,
        )));
    }
    let structured_result = (!structured.is_empty()).then(|| StructuredFieldContent {
        inline: Some(structured),
        blocks: None,
    });
    let field_tree = structured_result.clone().map(|result| StructuredFieldTree {
        version: Some(1.0),
        code: None,
        result: Some(result),
        children: None,
        display_mode: Some("result".to_owned()),
    });
    Ok(SimpleField {
        node_type: SimpleFieldType::SimpleField,
        field_type: parse_field_type(&instruction),
        instruction,
        content,
        fld_lock: matches!(element.attribute(Some("w"), "fldLock"), Some("1" | "true"))
            .then_some(true),
        dirty: matches!(element.attribute(Some("w"), "dirty"), Some("1" | "true")).then_some(true),
        structured_result,
        field_tree,
    })
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct BookmarkStart {
    #[serde(rename = "type")]
    pub node_type: BookmarkStartType,
    pub id: f64,
    pub name: String,
    #[serde(rename = "colFirst", skip_serializing_if = "Option::is_none")]
    pub col_first: Option<f64>,
    #[serde(rename = "colLast", skip_serializing_if = "Option::is_none")]
    pub col_last: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ContentPosition>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum BookmarkStartType {
    #[serde(rename = "bookmarkStart")]
    BookmarkStart,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct BookmarkEnd {
    #[serde(rename = "type")]
    pub node_type: BookmarkEndType,
    pub id: f64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub position: Option<ContentPosition>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum BookmarkEndType {
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct ContentPosition {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub offset: Option<f64>,
}

pub fn parse_bookmark_start(element: &XmlElement) -> BookmarkStart {
    BookmarkStart {
        node_type: BookmarkStartType::BookmarkStart,
        id: element
            .parse_numeric_attribute(Some("w"), "id", 1.0)
            .unwrap_or(0.0),
        name: element
            .attribute(Some("w"), "name")
            .unwrap_or_default()
            .to_owned(),
        col_first: element.parse_numeric_attribute(Some("w"), "colFirst", 1.0),
        col_last: element.parse_numeric_attribute(Some("w"), "colLast", 1.0),
        position: None,
    }
}

pub fn parse_bookmark_end(element: &XmlElement) -> BookmarkEnd {
    BookmarkEnd {
        node_type: BookmarkEndType::BookmarkEnd,
        id: element
            .parse_numeric_attribute(Some("w"), "id", 1.0)
            .unwrap_or(0.0),
        position: None,
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Hyperlink {
    #[serde(rename = "type")]
    pub node_type: HyperlinkType,
    #[serde(rename = "rId", skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub href: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub anchor: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub history: Option<bool>,
    #[serde(rename = "docLocation", skip_serializing_if = "Option::is_none")]
    pub doc_location: Option<String>,
    pub children: Vec<InlineNode>,
    #[serde(rename = "structuredChildren", skip_serializing_if = "Option::is_none")]
    pub structured_children: Option<Vec<InlineNode>>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum HyperlinkType {
    #[serde(rename = "hyperlink")]
    Hyperlink,
}

pub fn sanitize_href(href: Option<&str>) -> Option<String> {
    let href = href.filter(|value| !value.is_empty())?;
    let mut probe: String = href
        .chars()
        .filter(|character| !matches!(character, '\t' | '\n' | '\r'))
        .collect();
    probe = probe
        .trim_start_matches(|character: char| (character as u32) <= 0x20)
        .to_owned();
    if probe.is_empty() {
        return None;
    }
    let Some(colon) = probe.find(':') else {
        return Some(href.to_owned());
    };
    let scheme = &probe[..colon];
    let has_scheme = scheme.chars().enumerate().all(|(index, character)| {
        if index == 0 {
            character.is_ascii_alphabetic()
        } else {
            character.is_ascii_alphanumeric() || matches!(character, '+' | '.' | '-')
        }
    }) && !scheme.is_empty();
    if !has_scheme {
        return Some(href.to_owned());
    }
    matches!(
        scheme.to_ascii_lowercase().as_str(),
        "http" | "https" | "mailto" | "tel" | "ftp"
    )
    .then(|| href.to_owned())
}

pub fn parse_hyperlink(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &ParseBudget<'_>,
) -> Result<Hyperlink, ParseError> {
    let relationship_id = element.attribute(Some("r"), "id").map(str::to_owned);
    let mut href = relationship_id
        .as_deref()
        .and_then(|id| relationships.and_then(|relationships| relationships.get(id)))
        .and_then(|relationship| sanitize_href(Some(&relationship.target)));
    let anchor = element
        .attribute(Some("w"), "anchor")
        .filter(|value| !value.is_empty())
        .map(|value| truncate_chars(value, 1024));
    if href.is_none() {
        href = anchor.as_ref().map(|anchor| format!("#{anchor}"));
    }
    let tooltip = element
        .attribute(Some("w"), "tooltip")
        .filter(|value| !value.is_empty())
        .map(|value| truncate_chars(value, 2048));
    let target = element
        .attribute(Some("w"), "tgtFrame")
        .filter(|value| !value.is_empty())
        .map(|value| truncate_chars(value, 255));
    let history =
        matches!(element.attribute(Some("w"), "history"), Some("1" | "true")).then_some(true);
    let doc_location = element
        .attribute(Some("w"), "docLocation")
        .filter(|value| !value.is_empty())
        .map(|value| truncate_chars(value, 2048));
    let mut children = Vec::new();
    let mut structured = Vec::new();
    for child in element.child_elements().take(MAX_HYPERLINK_CHILDREN) {
        match child.local_name() {
            "r" => {
                let run = parse_run(child, theme, styles, doc_defaults).run;
                children.push(InlineNode::Run(run.clone()));
                structured.push(InlineNode::Run(run));
            }
            "bookmarkStart" => {
                let bookmark = parse_bookmark_start(child);
                children.push(InlineNode::BookmarkStart(bookmark.clone()));
                structured.push(InlineNode::BookmarkStart(bookmark));
            }
            "bookmarkEnd" => {
                let bookmark = parse_bookmark_end(child);
                children.push(InlineNode::BookmarkEnd(bookmark.clone()));
                structured.push(InlineNode::BookmarkEnd(bookmark));
            }
            "fldSimple" => structured.push(InlineNode::SimpleField(Box::new(parse_simple_field(
                child,
                theme,
                styles,
                doc_defaults,
                part,
                budget,
            )?))),
            "sdt" => structured.push(InlineNode::InlineSdt(Box::new(parse_hyperlink_inline_sdt(
                child,
                relationships,
                theme,
                styles,
                doc_defaults,
                part,
                budget,
            )?))),
            "oMath" | "oMathPara" => structured.push(InlineNode::Math(parse_hyperlink_math(child))),
            _ => {}
        }
    }
    let structured_children = structured
        .iter()
        .any(|child| child.node_type() != "run")
        .then_some(structured);
    Ok(Hyperlink {
        node_type: HyperlinkType::Hyperlink,
        relationship_id,
        href,
        anchor,
        tooltip,
        target,
        history,
        doc_location,
        children,
        structured_children,
    })
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct MathEquation {
    #[serde(rename = "type")]
    pub node_type: MathType,
    pub display: String,
    #[serde(rename = "ommlXml")]
    pub omml_xml: String,
    #[serde(rename = "plainText", skip_serializing_if = "Option::is_none")]
    pub plain_text: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum MathType {
    #[serde(rename = "mathEquation")]
    MathEquation,
}

fn parse_hyperlink_math(element: &XmlElement) -> MathEquation {
    let text = incumbent_text_content(element);
    MathEquation {
        node_type: MathType::MathEquation,
        display: if element.local_name() == "oMathPara" {
            "block"
        } else {
            "inline"
        }
        .to_owned(),
        omml_xml: element.to_incumbent_xml(),
        plain_text: (!text.is_empty()).then_some(text),
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct SdtCheckboxGlyph {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SdtDateState {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub full_date: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub language: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub calendar: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub storage_format: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct SdtGalleryState {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gallery: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub unique: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SdtControlState {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checked: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selected_index: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selected_value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub enabled: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub placeholder: Option<bool>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SdtListItem {
    pub display_text: String,
    pub value: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SdtDataBinding {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xpath: Option<String>,
    #[serde(rename = "storeItemID", skip_serializing_if = "Option::is_none")]
    pub store_item_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub prefix_mappings: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SdtProperties {
    pub sdt_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alias: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tag: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub lock: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub placeholder: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub showing_placeholder: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_items: Option<Vec<SdtListItem>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checked: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub run_properties: Option<TextFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub temporary: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tab_index: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub multi_line: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub date_state: Option<SdtDateState>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_last_value: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub checked_state: Option<SdtCheckboxGlyph>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub unchecked_state: Option<SdtCheckboxGlyph>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub gallery: Option<SdtGalleryState>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub appearance: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<crate::scalars::ColorValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub control_state: Option<SdtControlState>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub repeating_section: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub repeating_section_item: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_binding: Option<SdtDataBinding>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub raw_properties_xml: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub raw_end_properties_xml: Option<String>,
}

pub fn parse_sdt_properties(
    sdt_pr: Option<&XmlElement>,
    sdt_end_pr: Option<&XmlElement>,
    theme: Option<&Theme>,
) -> SdtProperties {
    let mut properties = SdtProperties {
        sdt_type: parse_sdt_control_type(sdt_pr).to_owned(),
        id: None,
        alias: None,
        tag: None,
        lock: None,
        placeholder: None,
        showing_placeholder: None,
        date_format: None,
        list_items: None,
        checked: None,
        run_properties: None,
        temporary: None,
        label: None,
        tab_index: None,
        multi_line: None,
        date_state: None,
        list_last_value: None,
        checked_state: None,
        unchecked_state: None,
        gallery: None,
        appearance: None,
        color: None,
        control_state: None,
        repeating_section: None,
        repeating_section_item: None,
        data_binding: None,
        raw_properties_xml: sdt_pr.map(XmlElement::to_incumbent_xml),
        raw_end_properties_xml: sdt_end_pr.map(XmlElement::to_xml),
    };
    let Some(sdt_pr) = sdt_pr else {
        return properties;
    };
    for element in sdt_pr.child_elements() {
        match element.local_name() {
            "id" => {
                properties.id = element.parse_numeric_attribute(Some("w"), "val", 1.0);
            }
            "alias" => properties.alias = element.attribute(Some("w"), "val").map(str::to_owned),
            "tag" => properties.tag = element.attribute(Some("w"), "val").map(str::to_owned),
            "lock" => {
                properties.lock = Some(
                    element
                        .attribute(Some("w"), "val")
                        .unwrap_or("unlocked")
                        .to_owned(),
                )
            }
            "rPr" => properties.run_properties = parse_inline_run_properties(Some(element), theme),
            "temporary" => properties.temporary = Some(element.parse_boolean("w")),
            "label" => {
                properties.label =
                    parse_js_number(element.attribute(Some("w"), "val")).filter(|value| {
                        value.fract() == 0.0 && *value >= 0.0 && *value <= u32::MAX as f64
                    })
            }
            "tabIndex" => {
                properties.tab_index =
                    parse_js_number(element.attribute(Some("w"), "val")).filter(|value| {
                        value.fract() == 0.0 && *value >= 0.0 && *value <= u32::MAX as f64
                    })
            }
            "placeholder" => {
                properties.placeholder = element
                    .child("w", "docPart")
                    .and_then(|part| part.attribute(Some("w"), "val"))
                    .map(str::to_owned)
            }
            "showingPlcHdr" => {
                properties.showing_placeholder = Some(
                    element
                        .attribute(Some("w"), "val")
                        .is_none_or(|value| !matches_ci(value, &["0", "false", "off"])),
                )
            }
            "date" => parse_sdt_date(element, &mut properties),
            "dropDownList" | "comboBox" => parse_sdt_list(element, &mut properties),
            "text" => properties.multi_line = Some(element.parse_boolean("w")),
            "checkbox" => parse_sdt_checkbox(element, &mut properties),
            "docPartObj" | "docPartList" => parse_sdt_gallery(element, &mut properties),
            "appearance" => {
                properties.appearance = element
                    .attribute(Some("w15"), "val")
                    .or_else(|| element.attribute(Some("w"), "val"))
                    .filter(|value| matches!(*value, "boundingBox" | "tags" | "hidden"))
                    .map(str::to_owned)
            }
            "color" => {
                properties.color = element
                    .attribute(Some("w15"), "val")
                    .or_else(|| element.attribute(Some("w"), "val"))
                    .filter(|value| {
                        value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit())
                    })
                    .map(|value| crate::scalars::ColorValue {
                        rgb: Some(value.to_owned()),
                        ..crate::scalars::ColorValue::default()
                    })
            }
            "repeatingSection" => {
                properties.repeating_section = Some(true);
                properties.sdt_type = "repeatingSection".to_owned();
            }
            "repeatingSectionItem" => {
                properties.repeating_section_item = Some(true);
                properties.sdt_type = "repeatingSectionItem".to_owned();
            }
            "dataBinding" => {
                properties.data_binding = Some(SdtDataBinding {
                    xpath: element.attribute(Some("w"), "xpath").map(str::to_owned),
                    store_item_id: element
                        .attribute(Some("w"), "storeItemID")
                        .map(str::to_owned),
                    prefix_mappings: element
                        .attribute(Some("w"), "prefixMappings")
                        .map(str::to_owned),
                })
            }
            _ => {}
        }
    }
    properties
}

fn parse_sdt_control_type(element: Option<&XmlElement>) -> &'static str {
    let Some(element) = element else {
        return "richText";
    };
    for child in element.child_elements() {
        let mapped = match child.local_name() {
            "richText" => Some("richText"),
            "text" => Some("plainText"),
            "date" => Some("date"),
            "dropDownList" => Some("dropDownList"),
            "comboBox" => Some("comboBox"),
            "picture" => Some("picture"),
            "docPartObj" | "docPartList" => Some("buildingBlockGallery"),
            "group" => Some("group"),
            "equation" => Some("equation"),
            "citation" => Some("citation"),
            "bibliography" => Some("bibliography"),
            "checkbox" => Some("checkbox"),
            "repeatingSection" => Some("repeatingSection"),
            "repeatingSectionItem" => Some("repeatingSectionItem"),
            _ => None,
        };
        if let Some(mapped) = mapped {
            return mapped;
        }
    }
    // Pinned incumbent quirk: unknown markers fall back to richText even
    // though the public type documentation describes an `unknown` variant.
    "richText"
}

fn parse_sdt_date(element: &XmlElement, properties: &mut SdtProperties) {
    properties.date_format = element
        .child("w", "dateFormat")
        .and_then(|format| format.attribute(Some("w"), "val"))
        .map(str::to_owned);
    let full_date = element.attribute(Some("w"), "fullDate").map(str::to_owned);
    let language = element.child("w", "lid").map(|language| {
        truncate_chars(language.attribute(Some("w"), "val").unwrap_or_default(), 63)
    });
    let calendar = element.child("w", "calendar").map(|calendar| {
        truncate_chars(calendar.attribute(Some("w"), "val").unwrap_or_default(), 63)
    });
    let storage_format = element
        .child("w", "storeMappedDataAs")
        .and_then(|storage| storage.attribute(Some("w"), "val"))
        .filter(|value| matches!(*value, "date" | "dateTime" | "text"))
        .map(str::to_owned);
    properties.date_state = Some(SdtDateState {
        full_date: full_date
            .as_deref()
            .filter(|value| !value.is_empty())
            .map(|value| truncate_chars(value, 255)),
        format: properties
            .date_format
            .as_deref()
            .filter(|value| !value.is_empty())
            .map(|value| truncate_chars(value, 255)),
        language,
        calendar,
        storage_format,
    });
    properties.control_state = Some(SdtControlState {
        value: full_date,
        placeholder: properties.showing_placeholder,
        ..SdtControlState::default()
    });
}

fn parse_sdt_list(element: &XmlElement, properties: &mut SdtProperties) {
    let items: Vec<_> = element
        .child_elements()
        .take(MAX_SDT_LIST_ITEMS)
        .filter(|child| child.local_name() == "listItem")
        .map(|child| SdtListItem {
            display_text: child
                .attribute(Some("w"), "displayText")
                .unwrap_or_default()
                .to_owned(),
            value: child
                .attribute(Some("w"), "value")
                .unwrap_or_default()
                .to_owned(),
        })
        .collect();
    properties.list_items = Some(items.clone());
    properties.list_last_value = element.attribute(Some("w"), "lastValue").map(str::to_owned);
    let selected_index = properties.list_last_value.as_ref().and_then(|selected| {
        items
            .iter()
            .position(|item| &item.value == selected)
            .map(|index| index as f64)
    });
    properties.control_state = Some(SdtControlState {
        selected_value: properties
            .list_last_value
            .as_ref()
            .filter(|value| !value.is_empty())
            .cloned(),
        selected_index,
        placeholder: properties.showing_placeholder,
        ..SdtControlState::default()
    });
}

fn parse_sdt_checkbox(element: &XmlElement, properties: &mut SdtProperties) {
    let checked = element
        .child_by_local_name("checked")
        .is_some_and(|checked| {
            checked
                .attribute(Some("w14"), "val")
                .or_else(|| checked.attribute(Some("w"), "val"))
                == Some("1")
        });
    properties.checked = Some(checked);
    properties.checked_state = parse_checkbox_state(element, "checkedState");
    properties.unchecked_state = parse_checkbox_state(element, "uncheckedState");
    properties.control_state = Some(SdtControlState {
        checked: Some(checked),
        placeholder: properties.showing_placeholder,
        ..SdtControlState::default()
    });
}

fn parse_checkbox_state(element: &XmlElement, name: &str) -> Option<SdtCheckboxGlyph> {
    let state = element.child_by_local_name(name)?;
    Some(SdtCheckboxGlyph {
        value: state
            .attribute(Some("w14"), "val")
            .or_else(|| state.attribute(Some("w"), "val"))
            .map(str::to_owned),
        font: state
            .attribute(Some("w14"), "font")
            .or_else(|| state.attribute(Some("w"), "font"))
            .map(str::to_owned),
    })
}

fn parse_sdt_gallery(element: &XmlElement, properties: &mut SdtProperties) {
    properties.gallery = Some(SdtGalleryState {
        gallery: element
            .child("w", "docPartGallery")
            .and_then(|gallery| gallery.attribute(Some("w"), "val"))
            .map(str::to_owned),
        category: element
            .child("w", "docPartCategory")
            .and_then(|category| category.attribute(Some("w"), "val"))
            .map(str::to_owned),
        unique: element
            .child("w", "docPartUnique")
            .map(|unique| unique.parse_boolean("w")),
    });
}

fn parse_js_number(raw: Option<&str>) -> Option<f64> {
    let value = raw.unwrap_or_default().trim();
    if value.is_empty() {
        return Some(0.0);
    }
    let parsed = if let Some(hex) = value
        .strip_prefix("0x")
        .or_else(|| value.strip_prefix("0X"))
    {
        u64::from_str_radix(hex, 16).ok().map(|value| value as f64)
    } else {
        value.parse::<f64>().ok()
    }?;
    parsed.is_finite().then_some(parsed)
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct InlineSdt {
    #[serde(rename = "type")]
    pub node_type: InlineSdtType,
    pub properties: SdtProperties,
    pub content: Vec<InlineNode>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum InlineSdtType {
    #[serde(rename = "inlineSdt")]
    InlineSdt,
}

fn parse_hyperlink_inline_sdt(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &ParseBudget<'_>,
) -> Result<InlineSdt, ParseError> {
    // Pinned incumbent quirk: hyperlinkParser omits the theme when parsing the
    // SDT's own run properties, even though it passes the theme to child runs.
    let properties = parse_sdt_properties(element.child("w", "sdtPr"), None, None);
    let mut content = Vec::new();
    if let Some(container) = element.child("w", "sdtContent") {
        for child in container.child_elements().take(MAX_HYPERLINK_CHILDREN) {
            match child.local_name() {
                "r" => content.push(InlineNode::Run(
                    parse_run(child, theme, styles, doc_defaults).run,
                )),
                "fldSimple" => content.push(InlineNode::SimpleField(Box::new(parse_simple_field(
                    child,
                    theme,
                    styles,
                    doc_defaults,
                    part,
                    budget,
                )?))),
                "oMath" | "oMathPara" => {
                    content.push(InlineNode::Math(parse_hyperlink_math(child)))
                }
                _ => {}
            }
        }
    }
    let _ = relationships;
    Ok(InlineSdt {
        node_type: InlineSdtType::InlineSdt,
        properties,
        content,
    })
}

#[derive(Clone, Debug)]
struct OpenComplexField {
    instruction: String,
    code_runs: Vec<Run>,
    result_runs: Vec<Run>,
    structured_code: Vec<InlineNode>,
    structured_result: Vec<InlineNode>,
    children: Vec<StructuredFieldTree>,
    mode: FieldMode,
    fld_lock: bool,
    dirty: bool,
    formatting: Option<TextFormatting>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum FieldMode {
    Code,
    Result,
}

/// Parse the S5-owned inline grammar at a paragraph/container boundary.
pub fn parse_inline_container(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &ParseBudget<'_>,
    depth: usize,
) -> Result<Vec<InlineNode>, ParseError> {
    budget.check_nesting_depth(depth, part)?;
    let mut output = Vec::new();
    let mut fields: Vec<OpenComplexField> = Vec::new();
    for child in element.child_elements() {
        match child.local_name() {
            "r" => {
                let run = parse_run(child, theme, styles, doc_defaults).run;
                let mut has_begin = false;
                let mut has_separate = false;
                let mut has_end = false;
                let mut instruction = String::new();
                for content in &run.content {
                    match content {
                        RunContent::FieldChar { char_type, .. } if char_type == "begin" => {
                            has_begin = true
                        }
                        RunContent::FieldChar { char_type, .. } if char_type == "separate" => {
                            has_separate = true
                        }
                        RunContent::FieldChar { char_type, .. } if char_type == "end" => {
                            has_end = true
                        }
                        RunContent::InstrText { text } => instruction.push_str(text),
                        _ => {}
                    }
                }
                if has_begin {
                    if fields.len() >= MAX_FIELD_NESTING {
                        return Err(ParseError::ResourceLimit {
                            kind: "fieldDepth",
                            part: part.to_owned(),
                        });
                    }
                    fields.push(create_open_complex_field(&run));
                }
                if let Some(active) = fields.last_mut() {
                    if !instruction.is_empty() {
                        active.instruction.push_str(&instruction);
                    }
                    if active.formatting.is_none() {
                        active.formatting.clone_from(&run.formatting);
                    }
                    if has_separate {
                        active.mode = FieldMode::Result;
                    }
                    if !has_begin && !has_separate && !has_end {
                        match active.mode {
                            FieldMode::Code => {
                                active.code_runs.push(run.clone());
                                active.structured_code.push(InlineNode::Run(run));
                            }
                            FieldMode::Result => {
                                active.result_runs.push(run.clone());
                                active.structured_result.push(InlineNode::Run(run));
                            }
                        }
                    }
                    if has_end {
                        let completed = finalize_open_complex_field(fields.pop().unwrap());
                        if let Some(parent) = fields.last_mut() {
                            append_nested_field(parent, completed);
                        } else {
                            output.push(InlineNode::ComplexField(Box::new(completed)));
                        }
                    }
                } else {
                    output.push(InlineNode::Run(run));
                }
            }
            "hyperlink" => output.push(InlineNode::Hyperlink(Box::new(parse_hyperlink(
                child,
                relationships,
                theme,
                styles,
                doc_defaults,
                part,
                budget,
            )?))),
            "bookmarkStart" => output.push(InlineNode::BookmarkStart(parse_bookmark_start(child))),
            "bookmarkEnd" => output.push(InlineNode::BookmarkEnd(parse_bookmark_end(child))),
            "fldSimple" => output.push(InlineNode::SimpleField(Box::new(parse_rich_simple_field(
                child,
                relationships,
                theme,
                styles,
                doc_defaults,
                part,
                budget,
                depth + 1,
            )?))),
            "sdt" => {
                if let Some(content) = child.child("w", "sdtContent") {
                    let parsed = parse_inline_container(
                        content,
                        relationships,
                        theme,
                        styles,
                        doc_defaults,
                        part,
                        budget,
                        depth + 1,
                    )?;
                    let allowed = parsed
                        .into_iter()
                        .filter(|node| {
                            matches!(
                                node,
                                InlineNode::Run(_)
                                    | InlineNode::Hyperlink(_)
                                    | InlineNode::SimpleField(_)
                                    | InlineNode::ComplexField(_)
                                    | InlineNode::InlineSdt(_)
                                    | InlineNode::Math(_)
                            )
                        })
                        .collect();
                    output.push(InlineNode::InlineSdt(Box::new(InlineSdt {
                        node_type: InlineSdtType::InlineSdt,
                        properties: parse_sdt_properties(child.child("w", "sdtPr"), None, theme),
                        content: allowed,
                    })));
                }
            }
            "oMath" | "oMathPara" => output.push(InlineNode::Math(parse_paragraph_math(child))),
            _ => {}
        }
    }
    // Pinned compatibility: unterminated fields remain inert structured nodes
    // so later block parsing can attach following result blocks.
    while let Some(field) = fields.pop() {
        let completed = finalize_open_complex_field(field);
        if let Some(parent) = fields.last_mut() {
            append_nested_field(parent, completed);
        } else {
            output.push(InlineNode::ComplexField(Box::new(completed)));
        }
    }
    assign_marker_offsets(&mut output);
    Ok(output)
}

fn parse_rich_simple_field(
    element: &XmlElement,
    relationships: Option<&RelationshipMap>,
    theme: Option<&Theme>,
    styles: Option<&StyleMap>,
    doc_defaults: Option<&DocDefaults>,
    part: &str,
    budget: &ParseBudget<'_>,
    depth: usize,
) -> Result<SimpleField, ParseError> {
    let instruction = element
        .attribute(Some("w"), "instr")
        .unwrap_or_default()
        .to_owned();
    let result: Vec<_> = parse_inline_container(
        element,
        relationships,
        theme,
        styles,
        doc_defaults,
        part,
        budget,
        depth,
    )?
    .into_iter()
    .filter(|node| {
        matches!(
            node,
            InlineNode::Run(_)
                | InlineNode::Hyperlink(_)
                | InlineNode::SimpleField(_)
                | InlineNode::ComplexField(_)
                | InlineNode::InlineSdt(_)
                | InlineNode::Math(_)
        )
    })
    .collect();
    let content = result
        .iter()
        .filter_map(|node| match node {
            InlineNode::Run(run) => Some(run.clone()),
            _ => None,
        })
        .collect();
    let structured_result = (!result.is_empty()).then(|| StructuredFieldContent {
        inline: Some(result),
        blocks: None,
    });
    let field_tree = structured_result.clone().map(|result| StructuredFieldTree {
        version: Some(1.0),
        code: None,
        result: Some(result),
        children: None,
        display_mode: None,
    });
    Ok(SimpleField {
        node_type: SimpleFieldType::SimpleField,
        field_type: parse_field_type(&instruction),
        instruction,
        content,
        fld_lock: matches!(element.attribute(Some("w"), "fldLock"), Some("1" | "true"))
            .then_some(true),
        dirty: matches!(element.attribute(Some("w"), "dirty"), Some("1" | "true")).then_some(true),
        structured_result,
        field_tree,
    })
}

fn create_open_complex_field(run: &Run) -> OpenComplexField {
    let flags = run.content.iter().find_map(|content| match content {
        RunContent::FieldChar {
            char_type,
            fld_lock,
            dirty,
            ..
        } if char_type == "begin" => Some((*fld_lock == Some(true), *dirty == Some(true))),
        _ => None,
    });
    OpenComplexField {
        instruction: String::new(),
        code_runs: Vec::new(),
        result_runs: Vec::new(),
        structured_code: Vec::new(),
        structured_result: Vec::new(),
        children: Vec::new(),
        mode: FieldMode::Code,
        fld_lock: flags.is_some_and(|flags| flags.0),
        dirty: flags.is_some_and(|flags| flags.1),
        formatting: run.formatting.clone(),
    }
}

fn finalize_open_complex_field(field: OpenComplexField) -> ComplexField {
    let structured_code = (!field.structured_code.is_empty()).then(|| StructuredFieldContent {
        inline: Some(field.structured_code),
        blocks: None,
    });
    let structured_result = (!field.structured_result.is_empty()).then(|| StructuredFieldContent {
        inline: Some(field.structured_result),
        blocks: None,
    });
    let field_tree = StructuredFieldTree {
        version: Some(1.0),
        code: structured_code.clone(),
        result: structured_result.clone(),
        children: (!field.children.is_empty()).then_some(field.children),
        display_mode: Some("result".to_owned()),
    };
    let instruction = field.instruction.trim().to_owned();
    ComplexField {
        node_type: ComplexFieldType::ComplexField,
        field_type: parse_field_type(&instruction),
        instruction,
        field_code: field.code_runs,
        field_result: field.result_runs,
        formatting: field.formatting,
        fld_lock: field.fld_lock.then_some(true),
        dirty: field.dirty.then_some(true),
        structured_code,
        structured_result,
        field_tree: Some(field_tree),
    }
}

fn append_nested_field(parent: &mut OpenComplexField, field: ComplexField) {
    let tree = field.field_tree.clone();
    match parent.mode {
        FieldMode::Code => parent
            .structured_code
            .push(InlineNode::ComplexField(Box::new(field))),
        FieldMode::Result => parent
            .structured_result
            .push(InlineNode::ComplexField(Box::new(field))),
    }
    if let Some(tree) = tree {
        parent.children.push(tree);
    }
}

fn parse_paragraph_math(element: &XmlElement) -> MathEquation {
    let mut text = String::new();
    append_math_text(element, &mut text);
    MathEquation {
        node_type: MathType::MathEquation,
        display: if element.local_name() == "oMathPara" {
            "block"
        } else {
            "inline"
        }
        .to_owned(),
        omml_xml: element.to_incumbent_xml(),
        plain_text: (!text.is_empty()).then_some(text),
    }
}

fn append_math_text(element: &XmlElement, output: &mut String) {
    for child in element.child_elements() {
        if child.local_name() == "t" {
            for node in &child.children {
                if let XmlNode::Text(text) = node {
                    output.push_str(text);
                }
            }
        } else {
            append_math_text(child, output);
        }
    }
}

fn assign_marker_offsets(content: &mut [InlineNode]) {
    let mut offset = 0usize;
    for node in content {
        match node {
            InlineNode::BookmarkStart(bookmark) => {
                bookmark.position = Some(ContentPosition {
                    offset: Some(offset as f64),
                })
            }
            InlineNode::BookmarkEnd(bookmark) => {
                bookmark.position = Some(ContentPosition {
                    offset: Some(offset as f64),
                })
            }
            _ => offset = offset.saturating_add(inline_content_length(node)),
        }
    }
}

fn inline_content_length(node: &InlineNode) -> usize {
    match node {
        InlineNode::Run(run) => run
            .content
            .iter()
            .map(|content| match content {
                RunContent::Text { text, .. } | RunContent::InstrText { text } => {
                    // TypeScript String#length counts UTF-16 code units. Preserve
                    // that incumbent offset behavior for astral characters.
                    text.encode_utf16().count()
                }
                RunContent::Tab
                | RunContent::SoftHyphen
                | RunContent::NoBreakHyphen
                | RunContent::Symbol { .. } => 1,
                _ => 0,
            })
            .sum(),
        InlineNode::Hyperlink(hyperlink) => hyperlink
            .children
            .iter()
            .filter(|child| matches!(child, InlineNode::Run(_)))
            .map(inline_content_length)
            .sum(),
        InlineNode::SimpleField(field) => field
            .content
            .iter()
            .cloned()
            .map(InlineNode::Run)
            .map(|node| inline_content_length(&node))
            .sum(),
        InlineNode::ComplexField(field) => field
            .field_result
            .iter()
            .cloned()
            .map(InlineNode::Run)
            .map(|node| inline_content_length(&node))
            .sum(),
        InlineNode::InlineSdt(sdt) => sdt.content.iter().map(inline_content_length).sum(),
        InlineNode::Math(math) => math
            .plain_text
            .as_deref()
            .map(|text| text.encode_utf16().count())
            .unwrap_or(0),
        InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_) => 0,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationships::{Relationship, TargetMode};
    use crate::styles::Style;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(
            xml.as_bytes(),
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
        .root()
        .unwrap()
        .clone()
    }

    #[test]
    fn run_projection_preserves_order_and_opaque_drawing_seam() {
        let run = root(
            r#"<w:r><w:rPr><w:b/><w14:glow w14:rad="9525"><w14:srgbClr w14:val="ABCDEF"/></w14:glow></w:rPr><w:t xml:space="preserve"> a </w:t><w:br w:type="page" w:clear="all"/><w:sym w:font="Wingdings" w:char="not-hex-💣"/><w:drawing><a:r><a:t>owned by S4</a:t></a:r></w:drawing><w:cr/></w:r>"#,
        );
        let projection = parse_run(&run, None, None, None);
        assert_eq!(projection.run.formatting.as_ref().unwrap().bold, Some(true));
        assert_eq!(
            projection
                .run
                .formatting
                .as_ref()
                .unwrap()
                .modern_effects
                .as_ref()
                .unwrap()
                .glow
                .as_ref()
                .unwrap()
                .radius,
            Some(1.0)
        );
        assert!(matches!(
            &projection.run.content[0],
            RunContent::Text {
                text,
                preserve_space: Some(true)
            } if text == " a "
        ));
        assert!(matches!(
            &projection.run.content[2],
            RunContent::Symbol { char, .. } if char == "not-hex-💣"
        ));
        assert!(matches!(
            &projection.run.content[3],
            RunContent::OpaqueDrawing { kind } if kind == "drawing"
        ));
        assert!(matches!(
            &projection.run.content[4],
            RunContent::Break { break_type: Some(kind), .. } if kind == "textWrapping"
        ));
    }

    #[test]
    fn field_instruction_tokenizer_pins_loose_regex_quirks_and_stays_linear() {
        let parsed = parse_field_instruction(r#" DATE "created at" \@ "MMMM d" \* MERGEFORMAT"#);
        assert_eq!(parsed.field_type, "DATE");
        assert_eq!(parsed.argument.as_deref(), Some("created at"));
        assert_eq!(parsed.switches[0].switch_name, "@");
        assert_eq!(parsed.switches[0].value.as_deref(), Some("MMMM d"));
        assert_eq!(parsed.switches[1].value.as_deref(), Some("MERGEFORMAT"));

        // `\s*` followed by the unquoted `\S*` branch consumes the next
        // switch token as a value. This surprising incumbent behavior is pinned.
        let loose = parse_field_instruction(r#"PAGE \h \* MERGEFORMAT"#);
        assert_eq!(loose.switches.len(), 1);
        assert_eq!(loose.switches[0].value.as_deref(), Some(r#"\*"#));

        let hostile = format!("DDE {}\\@ \"unterminated", "x".repeat(200_000));
        let parsed = parse_field_instruction(&hostile);
        assert_eq!(parsed.field_type, "UNKNOWN");
        assert_eq!(parsed.raw.len(), hostile.len());
    }

    #[test]
    fn hyperlink_sanitization_uses_probe_but_preserves_original_and_anchor_fallback() {
        for blocked in [
            "javascript:alert(1)",
            "java\tscript:alert(1)",
            "\rdata:text/html,x",
            " file:///etc/passwd",
        ] {
            assert_eq!(sanitize_href(Some(blocked)), None, "{blocked:?}");
        }
        let allowed = "  HTTPS://example.test/a";
        assert_eq!(sanitize_href(Some(allowed)).as_deref(), Some(allowed));

        let mut relationships = RelationshipMap::new();
        relationships.insert(
            "rId1".to_owned(),
            Relationship {
                id: "rId1".to_owned(),
                relationship_type: "hyperlink".to_owned(),
                target: "java\tscript:alert(1)".to_owned(),
                target_mode: Some(TargetMode::External),
            },
        );
        let element = root(
            r#"<w:hyperlink r:id="rId1" w:anchor="safe" w:tooltip="tip"><w:r><w:t>x</w:t></w:r></w:hyperlink>"#,
        );
        let limits = ParseLimits::default();
        let link = parse_hyperlink(
            &element,
            Some(&relationships),
            None,
            None,
            None,
            "word/document.xml",
            &ParseBudget::new(&limits),
        )
        .unwrap();
        assert_eq!(link.relationship_id.as_deref(), Some("rId1"));
        assert_eq!(link.href.as_deref(), Some("#safe"));
        assert_eq!(link.children.len(), 1);
        assert_eq!(link.structured_children, None);
    }

    #[test]
    fn complex_fields_are_nested_inert_and_unterminated_scopes_survive() {
        let paragraph = root(
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin" w:fldLock="1"/></w:r><w:r><w:instrText> IF </w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> INCLUDETEXT "https://example.test" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cached only</w:t></w:r></w:p>"#,
        );
        let limits = ParseLimits::default();
        let content = parse_inline_container(
            &paragraph,
            None,
            None,
            None,
            None,
            "word/document.xml",
            &ParseBudget::new(&limits),
            0,
        )
        .unwrap();
        let InlineNode::ComplexField(field) = &content[0] else {
            panic!("expected complex field")
        };
        assert_eq!(field.field_type, "IF");
        assert_eq!(field.fld_lock, Some(true));
        assert_eq!(field.field_result.len(), 1);
        assert_eq!(
            field
                .field_tree
                .as_ref()
                .and_then(|tree| tree.children.as_ref())
                .map(Vec::len),
            Some(1)
        );
        let nested = field
            .field_tree
            .as_ref()
            .unwrap()
            .children
            .as_ref()
            .unwrap();
        assert_eq!(nested[0].display_mode.as_deref(), Some("result"));
        // Parsing records INCLUDETEXT; it never resolves or fetches it.
        assert!(field.instruction.starts_with("IF"));
    }

    #[test]
    fn field_depth_limit_is_stable_before_stack_growth() {
        let mut xml = String::from("<w:p>");
        for _ in 0..=MAX_FIELD_NESTING {
            xml.push_str(r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#);
        }
        xml.push_str("</w:p>");
        let paragraph = root(&xml);
        let limits = ParseLimits::default();
        let error = parse_inline_container(
            &paragraph,
            None,
            None,
            None,
            None,
            "word/document.xml",
            &ParseBudget::new(&limits),
            0,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "fieldDepth",
                ..
            }
        ));
    }

    #[test]
    fn resolved_projection_composes_landed_s3_cascade_in_precedence_order() {
        let run = root(r#"<w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:b/></w:rPr></w:r>"#);
        let defaults = DocDefaults {
            r_pr: Some(TextFormatting {
                bold: Some(false),
                ..TextFormatting::default()
            }),
            p_pr: None,
        };
        let mut styles = StyleMap::new();
        styles.insert(
            "DefaultChar".to_owned(),
            Style {
                style_id: "DefaultChar".to_owned(),
                style_type: "character".to_owned(),
                default: Some(true),
                r_pr: Some(TextFormatting {
                    italic: Some(true),
                    ..TextFormatting::default()
                }),
                ..Style::default()
            },
        );
        styles.insert(
            "Emphasis".to_owned(),
            Style {
                style_id: "Emphasis".to_owned(),
                style_type: "character".to_owned(),
                r_pr: Some(TextFormatting {
                    color: Some(crate::scalars::ColorValue {
                        rgb: Some("FF0000".to_owned()),
                        ..crate::scalars::ColorValue::default()
                    }),
                    ..TextFormatting::default()
                }),
                ..Style::default()
            },
        );
        let resolved = parse_run(&run, None, Some(&styles), Some(&defaults))
            .resolved_formatting
            .unwrap();
        assert_eq!(resolved.bold, Some(true));
        assert_eq!(resolved.italic, Some(true));
        assert_eq!(resolved.color.unwrap().rgb.as_deref(), Some("FF0000"));
    }

    #[test]
    fn sdt_properties_pin_unknown_default_and_inert_control_state() {
        let sdt = root(
            r#"<w:sdt><w:sdtPr><w:mystery/><w:label/><w:checkbox><w:checked w:val="1"/><w:checkedState w:val="2612" w:font="MS Gothic"/></w:checkbox><w:dataBinding w:xpath="/x"/></w:sdtPr><w:sdtEndPr><w:rPr><w:i/></w:rPr></w:sdtEndPr></w:sdt>"#,
        );
        let properties =
            parse_sdt_properties(sdt.child("w", "sdtPr"), sdt.child("w", "sdtEndPr"), None);
        assert_eq!(properties.sdt_type, "checkbox");
        assert_eq!(properties.label, Some(0.0));
        assert_eq!(properties.checked, Some(true));
        assert_eq!(
            properties.checked_state.unwrap().value.as_deref(),
            Some("2612")
        );
        assert_eq!(
            properties.data_binding.unwrap().xpath.as_deref(),
            Some("/x")
        );
        assert!(
            properties
                .raw_properties_xml
                .unwrap()
                .starts_with("<w:sdtPr>")
        );
        assert!(properties.raw_end_properties_xml.is_some());

        let unknown = root(r#"<w:sdtPr><w:mystery/></w:sdtPr>"#);
        assert_eq!(parse_sdt_control_type(Some(&unknown)), "richText");
    }
}
