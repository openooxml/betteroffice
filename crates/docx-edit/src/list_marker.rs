//! List counters and marker templates shared by the render bridge and the structured export, so
//! both read the same marker for every list paragraph.

use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use docx_parse::NumberingMap;
use yrs::Any;

use crate::bridge::{any_bool, any_str, map_number, value_number, value_string};

/// The numbering levels OOXML defines, `0..=8`.
pub(crate) const LEVELS: usize = 9;

/// The largest level start honoured; larger starts are clamped.
const MAX_START: i64 = i32::MAX as i64;

/// The largest number Roman numerals write without overlines.
const MAX_ROMAN: i64 = 3_999;

/// A list number its level's format cannot write, such as a Roman numeral past 3,999.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct Unrenderable {
    pub value: i64,
    pub format: String,
}

/// List numbering carried across the whole story.
#[derive(Default)]
pub(crate) struct ListState {
    /// The current number of every level of each list, keyed by [`List::key`].
    pub(crate) counters: BTreeMap<String, Vec<i64>>,
    /// The `{numId}:{level}` pairs already numbered, so a `listStartOverride` read without
    /// definitions applies once rather than on every paragraph at that level.
    seen_num_ids: BTreeSet<String>,
    /// The document's numbering definitions, when the session retains them.
    numbering: Option<Arc<NumberingMap>>,
}

impl ListState {
    pub(crate) fn new(numbering: Option<Arc<NumberingMap>>) -> Self {
        Self {
            numbering,
            ..Self::default()
        }
    }
}

/// The level a numbering property names, when it is one of the nine OOXML defines.
pub(crate) fn numbering_level(value: f64) -> Option<usize> {
    (value.fract() == 0.0 && (0.0..LEVELS as f64).contains(&value)).then_some(value as usize)
}

/// One list's counters as the definitions describe them.
///
/// Numbering instances (`w:num`) that reference the same abstract definition and override no
/// start share one list, so a second instance continues the first one's numbers: Word's "Restart
/// numbering" writes a new instance with a `w:startOverride` precisely because a plain second
/// instance would continue. An instance overriding a start begins a list of its own, at that
/// start. Levels begin at `w:start` (1 when absent) and restart as `w:lvlRestart` says.
struct List {
    key: String,
    /// Per level: its start and the `w:lvlRestart` value.
    levels: [(i64, Option<f64>); LEVELS],
}

impl List {
    fn resolve(numbering: &NumberingMap, num_id: f64) -> Option<Self> {
        let instance = numbering.get_instance(num_id)?;
        let overrides = instance.level_overrides.as_deref().unwrap_or_default();
        let own = overrides.iter().any(|value| value.start_override.is_some());
        let key = if own {
            format!("{}#{num_id}", instance.abstract_num_id)
        } else {
            instance.abstract_num_id.to_string()
        };
        let mut levels = [(1, None); LEVELS];
        for (level, entry) in levels.iter_mut().enumerate() {
            let Some(definition) = numbering.get_level(num_id, level as f64) else {
                continue;
            };
            let start = overrides
                .iter()
                .find(|value| value.ilvl == level as f64)
                .and_then(|value| value.start_override)
                .or(definition.start)
                .filter(|start| start.is_finite())
                .map_or(1, |start| start.clamp(0.0, MAX_START as f64) as i64);
            *entry = (start, definition.lvl_restart);
        }
        Some(Self { key, levels })
    }

    fn initial(&self) -> Vec<i64> {
        self.levels.iter().map(|(start, _)| start - 1).collect()
    }

    /// Restarts the levels below `level` that numbering it restarts.
    fn restart_below(&self, counters: &mut [i64], level: usize) {
        for (deeper, (counter, (start, restart))) in counters
            .iter_mut()
            .zip(self.levels)
            .enumerate()
            .skip(level + 1)
        {
            let restarts = match restart {
                Some(0.0) => false,
                Some(value) if value >= 1.0 && value - 1.0 < deeper as f64 => {
                    level as f64 <= value - 1.0
                }
                _ => true,
            };
            if restarts {
                *counter = start - 1;
            }
        }
    }
}

fn format_roman(value: i64, uppercase: bool) -> String {
    if !(1..=MAX_ROMAN).contains(&value) {
        return String::new();
    }
    const ONES: [&str; 10] = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"];
    const TENS: [&str; 10] = ["", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"];
    const HUNDREDS: [&str; 10] = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM"];
    let mut result = "M".repeat((value / 1000) as usize);
    result.push_str(HUNDREDS[((value / 100) % 10) as usize]);
    result.push_str(TENS[((value / 10) % 10) as usize]);
    result.push_str(ONES[(value % 10) as usize]);
    if uppercase {
        result
    } else {
        result.to_ascii_lowercase()
    }
}

fn format_alpha(mut value: i64, uppercase: bool) -> String {
    if value <= 0 {
        return String::new();
    }
    let mut chars = Vec::new();
    while value > 0 {
        chars.push((b'A' + ((value - 1) % 26) as u8) as char);
        value = (value - 1) / 26;
    }
    let result: String = chars.into_iter().rev().collect();
    if uppercase {
        result
    } else {
        result.to_ascii_lowercase()
    }
}

/// Whether [`format_list_counter`] renders `format` itself rather than falling back to decimal.
pub(crate) fn renders_format(format: &str) -> bool {
    matches!(
        format,
        "decimal"
            | "upperRoman"
            | "lowerRoman"
            | "upperLetter"
            | "lowerLetter"
            | "decimalZero"
            | "decimalZero3"
            | "decimalZero4"
            | "decimalZero5"
            | "none"
            | "bullet"
    )
}

/// `value` written in `format`, in space bounded by the value's digits.
fn format_list_counter(value: i64, format: Option<&str>) -> Result<String, Unrenderable> {
    if value <= 0 {
        return Ok(String::new());
    }
    Ok(match format {
        Some(roman @ ("upperRoman" | "lowerRoman")) if value > MAX_ROMAN => {
            return Err(Unrenderable {
                value,
                format: roman.to_owned(),
            });
        }
        Some("upperRoman") => format_roman(value, true),
        Some("lowerRoman") => format_roman(value, false),
        Some("upperLetter") => format_alpha(value, true),
        Some("lowerLetter") => format_alpha(value, false),
        Some("decimalZero") => format!("{value:02}"),
        Some("decimalZero3") => format!("{value:03}"),
        Some("decimalZero4") => format!("{value:04}"),
        Some("decimalZero5") => format!("{value:05}"),
        Some("none") => String::new(),
        _ => value.to_string(),
    })
}

pub(crate) fn list_level_formats(values: &BTreeMap<String, Any>) -> Vec<String> {
    let Some(Any::Array(formats)) = values.get("listLevelNumFmts") else {
        return Vec::new();
    };
    formats
        .iter()
        .filter_map(any_str)
        .map(str::to_owned)
        .collect()
}

fn resolve_list_template(
    template: &str,
    counters: &[i64],
    formats: &[String],
) -> Result<String, Unrenderable> {
    let chars: Vec<char> = template.chars().collect();
    let mut result = String::new();
    let mut index = 0;
    while index < chars.len() {
        if chars[index] == '%' && index + 1 < chars.len() {
            if let Some(digit) = chars[index + 1].to_digit(10) {
                if digit == 0 {
                    index += 2;
                    if chars
                        .get(index)
                        .is_some_and(|ch| matches!(ch, '.' | ')' | ':' | ']'))
                    {
                        index += 1;
                    }
                    continue;
                } else {
                    let counter_index = digit as usize - 1;
                    let value = counters.get(counter_index).copied().unwrap_or(0);
                    let formatted =
                        format_list_counter(value, formats.get(counter_index).map(String::as_str))?;
                    index += 2;
                    let punctuation = chars
                        .get(index)
                        .copied()
                        .filter(|ch| matches!(ch, '.' | ')' | ':' | ']'));
                    if !formatted.is_empty() {
                        result.push_str(&formatted);
                        if let Some(punctuation) = punctuation {
                            result.push(punctuation);
                        }
                    }
                    if punctuation.is_some() {
                        index += 1;
                    }
                    continue;
                }
            }
        }
        result.push(chars[index]);
        index += 1;
    }
    Ok(result)
}

/// [`list_marker`] without the reason a marker could not be written.
pub(crate) fn compute_list_marker(
    values: &BTreeMap<String, Any>,
    state: &mut ListState,
) -> Option<String> {
    list_marker(values, state).and_then(Result::ok)
}

/// The rendered list marker for one paragraph, advancing `state`. Bullets and authored literal
/// markers pass through; a marker containing `%` is a template resolved against the counter stack,
/// and a paragraph with no marker at all gets the dotted counter path (`"1.2."`). With retained
/// definitions, lists are counted as [`List`] describes; without them numbering a level resets
/// every deeper level. A level outside `0..=8` has no marker, and a number the level's format
/// cannot write is reported instead of written.
pub(crate) fn list_marker(
    values: &BTreeMap<String, Any>,
    state: &mut ListState,
) -> Option<Result<String, Unrenderable>> {
    let marker = value_string(values.get("listMarker"));
    let Some(Any::Map(num_pr)) = values.get("numPr") else {
        return marker.map(Ok);
    };
    let Some(num_id) = map_number(num_pr, "numId") else {
        return marker.map(Ok);
    };
    if num_id == 0.0 {
        return marker.map(Ok);
    }
    if values.get("listIsBullet").and_then(any_bool) == Some(true) {
        return Some(Ok(marker.unwrap_or_default()));
    }

    let level = numbering_level(map_number(num_pr, "ilvl").unwrap_or(0.0))?;
    let mut formats = list_level_formats(values);
    formats.truncate(LEVELS);
    let level_format = formats
        .get(level)
        .cloned()
        .or_else(|| value_string(values.get("listNumFmt")));
    let list = state
        .numbering
        .as_deref()
        .and_then(|numbering| List::resolve(numbering, num_id));
    let counter_key = match &list {
        Some(list) => list.key.clone(),
        None => value_number(values.get("listAbstractNumId"))
            .unwrap_or(num_id)
            .to_string(),
    };
    if level_format.as_deref() == Some("none") {
        if formats.len() <= level {
            formats.resize(level + 1, "decimal".to_owned());
            formats[level] = "none".to_owned();
        }
        let counters = state
            .counters
            .get(&counter_key)
            .map(Vec::as_slice)
            .unwrap_or_default();
        return marker
            .map(|template| resolve_list_template(&template, counters, &formats))
            .filter(|value| !matches!(value, Ok(text) if text.is_empty()));
    }

    let counters = state
        .counters
        .entry(counter_key)
        .or_insert_with(|| list.as_ref().map_or_else(|| vec![0; LEVELS], List::initial));
    if counters.len() < LEVELS {
        counters.resize(LEVELS, 0);
    }
    match &list {
        Some(list) => {
            counters[level] = counters[level].saturating_add(1);
            list.restart_below(counters, level);
        }
        None => {
            if state.seen_num_ids.insert(format!("{num_id}:{level}"))
                && let Some(start) =
                    value_number(values.get("listStartOverride")).filter(|start| start.is_finite())
            {
                counters[level] = start.clamp(0.0, MAX_START as f64) as i64 - 1;
            }
            counters[level] = counters[level].saturating_add(1);
            for value in counters.iter_mut().skip(level + 1) {
                *value = 0;
            }
        }
    }

    if let Some(marker) = marker {
        if marker.contains('%') {
            return Some(resolve_list_template(&marker, counters, &formats));
        }
        return Some(Ok(marker));
    }
    let mut parts = Vec::new();
    for value in counters.iter().take(level + 1) {
        if *value <= 0 {
            break;
        }
        parts.push(value.to_string());
    }
    Some(Ok(if parts.is_empty() {
        "1.".to_owned()
    } else {
        format!("{}.", parts.join("."))
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn definitions(levels: &str, instances: &str) -> Arc<NumberingMap> {
        let xml = format!(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="4">{levels}</w:abstractNum>{instances}</w:numbering>"#
        );
        let limits = docx_parse::ParseLimits::default();
        Arc::new(
            docx_parse::numbering::parse_numbering(
                Some(xml.as_bytes()),
                "word/numbering.xml",
                &mut docx_parse::ParseBudget::new(&limits),
            )
            .unwrap(),
        )
    }

    fn paragraph(num_id: f64, level: f64, template: &str) -> BTreeMap<String, Any> {
        BTreeMap::from([
            (
                "numPr".to_owned(),
                Any::from(std::collections::HashMap::from([
                    ("numId".to_owned(), Any::Number(num_id)),
                    ("ilvl".to_owned(), Any::Number(level)),
                ])),
            ),
            ("listMarker".to_owned(), Any::from(template)),
            ("listAbstractNumId".to_owned(), Any::Number(4.0)),
            (
                "listLevelNumFmts".to_owned(),
                Any::from_json(r#"["decimal","decimal","decimal"]"#).unwrap(),
            ),
        ])
    }

    #[test]
    fn a_restart_level_names_the_level_that_restarts_it() {
        let level = |ilvl: u32, extra: &str| {
            format!(
                r#"<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%{}."/>{extra}</w:lvl>"#,
                ilvl + 1
            )
        };
        let numbering = definitions(
            &[
                level(0, ""),
                level(1, ""),
                level(2, r#"<w:lvlRestart w:val="1"/>"#),
            ]
            .concat(),
            r#"<w:num w:numId="9"><w:abstractNumId w:val="4"/></w:num>"#,
        );
        let mut state = ListState::new(Some(numbering));
        let markers: Vec<Option<String>> = [
            (2.0, "%3."),
            (1.0, "%2."),
            (2.0, "%3."),
            (0.0, "%1."),
            (2.0, "%3."),
        ]
        .into_iter()
        .map(|(level, template)| compute_list_marker(&paragraph(9.0, level, template), &mut state))
        .collect();
        assert_eq!(
            markers,
            ["1.", "1.", "2.", "1.", "1."].map(|marker| Some(marker.to_owned())),
            "level 2 keeps counting past level 1 and restarts after level 0"
        );
    }

    fn decimal_level(start: u32) -> String {
        format!(
            r#"<w:lvl w:ilvl="0"><w:start w:val="{start}"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>"#
        )
    }

    fn markers(state: &mut ListState, paragraphs: &[BTreeMap<String, Any>]) -> Vec<String> {
        paragraphs
            .iter()
            .map(|values| compute_list_marker(values, state).unwrap_or_default())
            .collect()
    }

    #[test]
    fn instances_of_one_definition_continue_each_other() {
        let numbering = definitions(
            &decimal_level(1),
            r#"<w:num w:numId="1"><w:abstractNumId w:val="4"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="4"/></w:num>"#,
        );
        let mut state = ListState::new(Some(numbering));
        assert_eq!(
            markers(
                &mut state,
                &[
                    paragraph(1.0, 0.0, "%1."),
                    paragraph(1.0, 0.0, "%1."),
                    paragraph(2.0, 0.0, "%1."),
                    paragraph(1.0, 0.0, "%1."),
                ]
            ),
            ["1.", "2.", "3.", "4."]
        );
    }

    #[test]
    fn a_start_override_restarts_as_a_list_of_its_own() {
        let numbering = definitions(
            &decimal_level(1),
            concat!(
                r#"<w:num w:numId="1"><w:abstractNumId w:val="4"/></w:num>"#,
                r#"<w:num w:numId="2"><w:abstractNumId w:val="4"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>"#,
                r#"<w:num w:numId="3"><w:abstractNumId w:val="4"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="20"/><w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:lvlOverride></w:num>"#
            ),
        );
        let mut state = ListState::new(Some(numbering));
        assert_eq!(
            markers(
                &mut state,
                &[
                    paragraph(1.0, 0.0, "%1."),
                    paragraph(1.0, 0.0, "%1."),
                    paragraph(2.0, 0.0, "%1."),
                    paragraph(2.0, 0.0, "%1."),
                    paragraph(1.0, 0.0, "%1."),
                    paragraph(3.0, 0.0, "%1."),
                ]
            ),
            ["1.", "2.", "1.", "2.", "3.", "20."],
            "a start override also wins over a whole level override"
        );
    }

    #[test]
    fn numbers_a_format_cannot_write_are_reported() {
        let level = r#"<w:lvl w:ilvl="0"><w:start w:val="2000000000"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>"#;
        let numbering = definitions(
            level,
            r#"<w:num w:numId="1"><w:abstractNumId w:val="4"/></w:num>"#,
        );
        let mut state = ListState::new(Some(numbering));
        let mut values = paragraph(1.0, 0.0, "%1.");
        values.insert(
            "listLevelNumFmts".to_owned(),
            Any::from_json(r#"["upperRoman"]"#).unwrap(),
        );
        assert_eq!(
            list_marker(&values, &mut state),
            Some(Err(Unrenderable {
                value: 2_000_000_000,
                format: "upperRoman".to_owned()
            }))
        );
        assert_eq!(compute_list_marker(&values, &mut state), None);
    }

    #[test]
    fn levels_outside_the_schema_allocate_no_counters() {
        let mut state = ListState::default();
        assert_eq!(
            compute_list_marker(&paragraph(9.0, 4_294_967_295.0, "%1."), &mut state),
            None
        );
        assert!(state.counters.is_empty());
    }
}
