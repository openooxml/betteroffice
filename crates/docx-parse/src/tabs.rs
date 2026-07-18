//! Tab-stop leaf parsing plus bounded pure resolution helpers.

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use thiserror::Error;

use crate::xml::XmlElement;

pub const DEFAULT_TAB_INTERVAL_TWIPS: f64 = 720.0;
pub const DEFAULT_TAB_ALIGNMENT: &str = "left";
pub const DEFAULT_TAB_LEADER: &str = "none";
pub const MAX_GENERATED_TAB_STOPS: usize = 100_000;
pub const MAX_LEADER_CHARACTERS: usize = 1_000_000;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TabStop {
    pub position: f64,
    pub alignment: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub leader: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Error)]
pub enum TabError {
    #[error("tab interval and page width must be finite and the interval must be positive")]
    InvalidRange,
    #[error("generated tab stop limit exceeded")]
    TooManyStops,
    #[error("tab leader length limit exceeded")]
    LeaderTooLong,
}

pub fn parse_tab_stop(tab: &XmlElement) -> Option<TabStop> {
    let position = tab.parse_numeric_attribute(Some("w"), "pos", 1.0)?;
    let alignment = tab.attribute(Some("w"), "val")?;
    if alignment.is_empty() {
        return None;
    }
    Some(TabStop {
        position,
        alignment: alignment.to_owned(),
        leader: tab
            .attribute(Some("w"), "leader")
            .filter(|value| !value.is_empty())
            .map(str::to_owned),
    })
}

pub fn parse_tab_stops(tabs: Option<&XmlElement>) -> Vec<TabStop> {
    let Some(tabs) = tabs else { return Vec::new() };
    let mut stops: Vec<_> = tabs
        .children_named("w", "tab")
        .filter_map(parse_tab_stop)
        .collect();
    stops.sort_by(|left, right| left.position.total_cmp(&right.position));
    stops
}

pub fn parse_tab_stops_from_paragraph_properties(
    properties: Option<&XmlElement>,
) -> Option<Vec<TabStop>> {
    let stops = parse_tab_stops(properties?.child("w", "tabs"));
    (!stops.is_empty()).then_some(stops)
}

pub fn merge_tab_stops(style: Option<&[TabStop]>, direct: Option<&[TabStop]>) -> Vec<TabStop> {
    match (style, direct) {
        (None, None) => Vec::new(),
        (None, Some(direct)) => direct.to_vec(),
        (Some(style), None) => style.to_vec(),
        (Some(style), Some(direct)) => {
            let mut by_position: BTreeMap<i64, TabStop> = style
                .iter()
                .filter_map(|tab| js_number_key(tab.position).map(|key| (key, tab.clone())))
                .collect();
            for tab in direct {
                let Some(key) = js_number_key(tab.position) else {
                    continue;
                };
                if tab.alignment == "clear" {
                    by_position.remove(&key);
                } else {
                    by_position.insert(key, tab.clone());
                }
            }
            let mut result: Vec<_> = by_position.into_values().collect();
            result.sort_by(|left, right| left.position.total_cmp(&right.position));
            result
        }
    }
}

pub fn get_next_tab_stop(current: f64, stops: &[TabStop], page_width: f64) -> TabStop {
    if let Some(tab) = stops
        .iter()
        .find(|tab| tab.position > current && tab.alignment != "clear")
    {
        return tab.clone();
    }
    let default_position =
        ((current + 1.0) / DEFAULT_TAB_INTERVAL_TWIPS).ceil() * DEFAULT_TAB_INTERVAL_TWIPS;
    TabStop {
        position: default_position.min(page_width),
        alignment: DEFAULT_TAB_ALIGNMENT.to_owned(),
        leader: None,
    }
}

pub fn calculate_tab_width(current: f64, stops: &[TabStop], page_width: f64) -> f64 {
    (get_next_tab_stop(current, stops, page_width).position - current).max(0.0)
}

pub fn calculate_tab_width_with_alignment(
    current: f64,
    stops: &[TabStop],
    page_width: f64,
    following_width: f64,
) -> (f64, String) {
    let next = get_next_tab_stop(current, stops, page_width);
    let adjustment = match next.alignment.as_str() {
        "right" | "decimal" => following_width,
        "center" => following_width / 2.0,
        _ => 0.0,
    };
    (
        (next.position - current - adjustment).max(0.0),
        next.alignment,
    )
}

pub fn leader_character(leader: Option<&str>) -> char {
    match leader {
        Some("dot") => '.',
        Some("hyphen") => '-',
        Some("underscore" | "heavy") => '_',
        Some("middleDot") => '·',
        _ => ' ',
    }
}

pub fn has_visible_leader(leader: Option<&str>) -> bool {
    !matches!(leader, None | Some("none"))
}

pub fn generate_leader_string(leader: Option<&str>, width: f64) -> Result<String, TabError> {
    if !has_visible_leader(leader) {
        return Ok(String::new());
    }
    let count = if width.is_finite() && width > 0.0 {
        width.floor() as usize
    } else {
        0
    };
    if count > MAX_LEADER_CHARACTERS {
        return Err(TabError::LeaderTooLong);
    }
    Ok(std::iter::repeat_n(leader_character(leader), count).collect())
}

pub fn is_valid_tab_alignment(value: &str) -> bool {
    matches!(
        value,
        "left" | "center" | "right" | "decimal" | "bar" | "clear" | "num"
    )
}

pub fn is_valid_tab_leader(value: &str) -> bool {
    matches!(
        value,
        "none" | "dot" | "hyphen" | "underscore" | "heavy" | "middleDot"
    )
}

pub fn generate_default_tab_stops(
    page_width: f64,
    interval: f64,
) -> Result<Vec<TabStop>, TabError> {
    if !page_width.is_finite() || !interval.is_finite() || interval <= 0.0 {
        return Err(TabError::InvalidRange);
    }
    let count = if page_width <= interval {
        0
    } else {
        ((page_width - f64::EPSILON) / interval).floor() as usize
    };
    if count > MAX_GENERATED_TAB_STOPS {
        return Err(TabError::TooManyStops);
    }
    Ok((1..=count)
        .map(|index| TabStop {
            position: interval * index as f64,
            alignment: DEFAULT_TAB_ALIGNMENT.to_owned(),
            leader: None,
        })
        .collect())
}

fn js_number_key(value: f64) -> Option<i64> {
    value.is_finite().then(|| value.to_bits() as i64)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn root(xml: &str) -> XmlElement {
        let limits = ParseLimits::default();
        parse_xml(xml.as_bytes(), "tabs.xml", &mut ParseBudget::new(&limits))
            .unwrap()
            .root()
            .unwrap()
            .clone()
    }

    #[test]
    fn parses_unknown_values_bug_compatibly_and_sorts_positions() {
        let tabs = root(
            r#"<w:tabs><w:tab w:pos="720px" w:val="future"/><w:tab w:pos="-20" w:val="clear"/><w:tab w:pos="bad" w:val="left"/></w:tabs>"#,
        );
        let parsed = parse_tab_stops(Some(&tabs));
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].position, -20.0);
        assert_eq!(parsed[1].alignment, "future");
    }

    #[test]
    fn generation_and_leaders_are_bounded_for_hostile_values() {
        assert_eq!(generate_default_tab_stops(1441.0, 720.0).unwrap().len(), 2);
        assert_eq!(
            generate_default_tab_stops(100.0, 0.0),
            Err(TabError::InvalidRange)
        );
        assert_eq!(
            generate_default_tab_stops(f64::INFINITY, 720.0),
            Err(TabError::InvalidRange)
        );
        assert_eq!(
            generate_leader_string(Some("dot"), 1e20),
            Err(TabError::LeaderTooLong)
        );
        assert_eq!(generate_leader_string(Some("dot"), f64::NAN).unwrap(), "");
    }
}
