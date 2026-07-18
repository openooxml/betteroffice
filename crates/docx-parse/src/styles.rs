//! `styles.xml` definitions and incumbent-compatible `basedOn` cascade resolution.

use indexmap::{IndexMap, IndexSet};
use serde::{Deserialize, Serialize};

use crate::formatting::{
    ParagraphFormatting, TableCellFormatting, TableFormatting, TableRowFormatting, TextFormatting,
    merge_cell_formatting, merge_paragraph_formatting, merge_row_formatting,
    merge_table_formatting, merge_text_formatting, parse_paragraph_properties,
    parse_run_properties, parse_table_cell_properties, parse_table_properties,
    parse_table_row_properties,
};
use crate::settings::incumbent_utf8_text_boundary;
use crate::theme::Theme;
use crate::xml::{ParseBudget, ParseError, XmlElement, parse_xml};

pub type StyleMap = IndexMap<String, Style>;

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Style {
    pub style_id: String,
    #[serde(rename = "type")]
    pub style_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub based_on: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub next: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub link: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub ui_priority: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub semi_hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub unhide_when_used: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub q_format: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub default: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub personal: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<ParagraphFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<TextFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tbl_pr: Option<TableFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tr_pr: Option<TableRowFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tc_pr: Option<TableCellFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tbl_style_pr: Option<Vec<TableStyleConditional>>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyleConditional {
    #[serde(rename = "type")]
    pub conditional_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<ParagraphFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<TextFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tbl_pr: Option<TableFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tr_pr: Option<TableRowFormatting>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tc_pr: Option<TableCellFormatting>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct DocDefaults {
    #[serde(rename = "rPr", skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<TextFormatting>,
    #[serde(rename = "pPr", skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<ParagraphFormatting>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LatentStyles {
    pub def_locked_state: bool,
    #[serde(rename = "defUIPriority", skip_serializing_if = "Option::is_none")]
    pub def_ui_priority: Option<f64>,
    pub def_semi_hidden: bool,
    pub def_unhide_when_used: bool,
    pub def_q_format: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub count: Option<f64>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StyleDefinitions {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub doc_defaults: Option<DocDefaults>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub latent_styles: Option<LatentStyles>,
    pub styles: Vec<Style>,
}

pub fn parse_styles(
    xml: &[u8],
    theme: Option<&Theme>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<StyleMap, ParseError> {
    if !incumbent_utf8_text_boundary(xml) {
        return Ok(StyleMap::new());
    }
    let document = parse_xml(xml, part, budget)?;
    parse_style_map(document.root(), theme, part, budget)
}

pub fn parse_style_definitions(
    xml: &[u8],
    theme: Option<&Theme>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<StyleDefinitions, ParseError> {
    if !incumbent_utf8_text_boundary(xml) {
        return Ok(StyleDefinitions::default());
    }
    let document = parse_xml(xml, part, budget)?;
    let Some(root) = document.root() else {
        return Ok(StyleDefinitions::default());
    };
    let doc_defaults = parse_doc_defaults(root.child("w", "docDefaults"), theme);
    let latent_styles = root.child("w", "latentStyles").map(|element| LatentStyles {
        def_locked_state: element.attribute(Some("w"), "defLockedState") == Some("1"),
        def_ui_priority: element.parse_numeric_attribute(Some("w"), "defUIPriority", 1.0),
        def_semi_hidden: element.attribute(Some("w"), "defSemiHidden") == Some("1"),
        def_unhide_when_used: element.attribute(Some("w"), "defUnhideWhenUsed") == Some("1"),
        def_q_format: element.attribute(Some("w"), "defQFormat") == Some("1"),
        count: element.parse_numeric_attribute(Some("w"), "count", 1.0),
    });
    let styles = parse_style_map(Some(root), theme, part, budget)?
        .into_values()
        .collect();
    Ok(StyleDefinitions {
        doc_defaults,
        latent_styles,
        styles,
    })
}

fn parse_style_map(
    root: Option<&XmlElement>,
    theme: Option<&Theme>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<StyleMap, ParseError> {
    let mut styles = StyleMap::new();
    let Some(root) = root else { return Ok(styles) };
    for element in root.children_named("w", "style") {
        budget.charge_leaf_value(part)?;
        let style = parse_style(element, theme, part, budget)?;
        if !style.style_id.is_empty() {
            styles.insert(style.style_id.clone(), style);
        }
    }

    // Deliberately resolve against the live map in insertion order. Earlier
    // entries are already resolved when later entries recurse through them.
    // This order dependence is observable incumbent behavior, including cycles.
    let ids: Vec<_> = styles.keys().cloned().collect();
    for style_id in ids {
        let Some(style) = styles.get(&style_id).cloned() else {
            continue;
        };
        let resolved = resolve_style_inheritance(&style, &styles, part, budget)?;
        styles.insert(style_id, resolved);
    }
    Ok(styles)
}

fn parse_style(
    element: &XmlElement,
    theme: Option<&Theme>,
    part: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<Style, ParseError> {
    let default = element
        .attribute(Some("w"), "default")
        .filter(|raw| !raw.is_empty())
        .map(|raw| matches!(raw, "1" | "true"));
    let mut style = Style {
        style_id: element
            .attribute(Some("w"), "styleId")
            .unwrap_or_default()
            .to_owned(),
        style_type: element
            .attribute(Some("w"), "type")
            .unwrap_or("paragraph")
            .to_owned(),
        name: child_attribute(element, "name", "val"),
        based_on: child_attribute(element, "basedOn", "val"),
        next: child_attribute(element, "next", "val"),
        link: child_attribute(element, "link", "val"),
        ui_priority: numeric_child(element, "uiPriority", "val"),
        hidden: boolean_child(element, "hidden"),
        semi_hidden: boolean_child(element, "semiHidden"),
        unhide_when_used: boolean_child(element, "unhideWhenUsed"),
        q_format: boolean_child(element, "qFormat"),
        default,
        personal: boolean_child(element, "personal"),
        p_pr: parse_paragraph_properties(element.child("w", "pPr"), theme),
        r_pr: parse_run_properties(element.child("w", "rPr"), theme),
        tbl_pr: parse_table_properties(element.child("w", "tblPr")),
        tr_pr: parse_table_row_properties(element.child("w", "trPr")),
        tc_pr: parse_table_cell_properties(element.child("w", "tcPr")),
        tbl_style_pr: None,
    };
    let conditional_elements: Vec<_> = element.children_named("w", "tblStylePr").collect();
    if !conditional_elements.is_empty() {
        let mut conditionals = Vec::new();
        for conditional in conditional_elements {
            budget.charge_leaf_value(part)?;
            let Some(conditional_type) = conditional
                .attribute(Some("w"), "type")
                .filter(|raw| !raw.is_empty())
            else {
                continue;
            };
            conditionals.push(TableStyleConditional {
                conditional_type: conditional_type.to_owned(),
                p_pr: parse_paragraph_properties(conditional.child("w", "pPr"), theme),
                r_pr: parse_run_properties(conditional.child("w", "rPr"), theme),
                tbl_pr: parse_table_properties(conditional.child("w", "tblPr")),
                tr_pr: parse_table_row_properties(conditional.child("w", "trPr")),
                tc_pr: parse_table_cell_properties(conditional.child("w", "tcPr")),
            });
        }
        style.tbl_style_pr = Some(conditionals);
    }
    Ok(style)
}

fn parse_doc_defaults(element: Option<&XmlElement>, theme: Option<&Theme>) -> Option<DocDefaults> {
    let element = element?;
    let value = DocDefaults {
        r_pr: parse_run_properties(
            element
                .child("w", "rPrDefault")
                .and_then(|element| element.child("w", "rPr")),
            theme,
        ),
        p_pr: parse_paragraph_properties(
            element
                .child("w", "pPrDefault")
                .and_then(|element| element.child("w", "pPr")),
            theme,
        ),
    };
    (value != DocDefaults::default()).then_some(value)
}

fn resolve_style_inheritance(
    style: &Style,
    styles: &StyleMap,
    part: &str,
    budget: &ParseBudget<'_>,
) -> Result<Style, ParseError> {
    let mut visited = IndexSet::new();
    let mut callers = Vec::new();
    let mut current = style;
    let mut depth = 0usize;
    let mut resolved_parent = loop {
        budget.check_nesting_depth(depth, part)?;
        if visited.contains(&current.style_id) {
            break current.clone();
        }
        visited.insert(current.style_id.clone());
        let Some(parent_id) = current.based_on.as_deref() else {
            break current.clone();
        };
        let Some(parent) = styles.get(parent_id) else {
            break current.clone();
        };
        callers.push(current.clone());
        current = parent;
        depth += 1;
    };

    while let Some(child) = callers.pop() {
        resolved_parent = merge_resolved_parent(&resolved_parent, &child);
    }
    Ok(resolved_parent)
}

fn merge_resolved_parent(parent: &Style, style: &Style) -> Style {
    let mut resolved = style.clone();
    resolved.p_pr = merge_paragraph_formatting(parent.p_pr.as_ref(), style.p_pr.as_ref());
    resolved.r_pr = merge_text_formatting(parent.r_pr.as_ref(), style.r_pr.as_ref());
    if style.style_type == "table" {
        resolved.tbl_pr = merge_table_formatting(parent.tbl_pr.as_ref(), style.tbl_pr.as_ref());
        resolved.tr_pr = merge_row_formatting(parent.tr_pr.as_ref(), style.tr_pr.as_ref());
        resolved.tc_pr = merge_cell_formatting(parent.tc_pr.as_ref(), style.tc_pr.as_ref());
        resolved.tbl_style_pr = merge_table_style_conditionals(
            parent.tbl_style_pr.as_deref(),
            style.tbl_style_pr.as_deref(),
        );
    }
    resolved
}

fn merge_table_style_conditionals(
    parent: Option<&[TableStyleConditional]>,
    child: Option<&[TableStyleConditional]>,
) -> Option<Vec<TableStyleConditional>> {
    match (parent, child) {
        (None, None) => None,
        (Some(parent), None) => Some(parent.to_vec()),
        (None, Some(child)) => Some(child.to_vec()),
        (Some(parent), Some(child)) => {
            let mut merged = parent.to_vec();
            for part in child {
                if let Some(index) = merged
                    .iter()
                    .position(|candidate| candidate.conditional_type == part.conditional_type)
                {
                    let base = merged[index].clone();
                    merged[index] = TableStyleConditional {
                        conditional_type: part.conditional_type.clone(),
                        p_pr: merge_paragraph_formatting(base.p_pr.as_ref(), part.p_pr.as_ref()),
                        r_pr: merge_text_formatting(base.r_pr.as_ref(), part.r_pr.as_ref()),
                        tbl_pr: merge_table_formatting(base.tbl_pr.as_ref(), part.tbl_pr.as_ref()),
                        tr_pr: merge_row_formatting(base.tr_pr.as_ref(), part.tr_pr.as_ref()),
                        tc_pr: merge_cell_formatting(base.tc_pr.as_ref(), part.tc_pr.as_ref()),
                    };
                } else {
                    merged.push(part.clone());
                }
            }
            Some(merged)
        }
    }
}

pub fn get_default_paragraph_style(styles: &StyleMap) -> Option<&Style> {
    styles
        .values()
        .find(|style| style.style_type == "paragraph" && style.default == Some(true))
        .or_else(|| styles.get("Normal"))
}

pub fn get_default_character_style(styles: &StyleMap) -> Option<&Style> {
    styles
        .values()
        .find(|style| style.style_type == "character" && style.default == Some(true))
}

pub fn get_default_table_style(styles: &StyleMap) -> Option<&Style> {
    styles
        .values()
        .find(|style| style.style_type == "table" && style.default == Some(true))
}

pub fn get_styles_by_type<'a>(styles: &'a StyleMap, style_type: &str) -> Vec<&'a Style> {
    styles
        .values()
        .filter(|style| style.style_type == style_type)
        .collect()
}

fn child_attribute(parent: &XmlElement, child: &str, attribute: &str) -> Option<String> {
    parent
        .child("w", child)?
        .attribute(Some("w"), attribute)
        .map(str::to_owned)
}

fn numeric_child(parent: &XmlElement, child: &str, attribute: &str) -> Option<f64> {
    parent
        .child("w", child)?
        .parse_numeric_attribute(Some("w"), attribute, 1.0)
}

fn boolean_child(parent: &XmlElement, child: &str) -> Option<bool> {
    parent
        .child("w", child)
        .map(|element| element.parse_boolean("w"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::ParseLimits;

    fn parse(xml: &str) -> StyleDefinitions {
        let limits = ParseLimits::default();
        parse_style_definitions(
            xml.as_bytes(),
            None,
            "word/styles.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap()
    }

    #[test]
    fn parses_defaults_latent_flags_and_resolves_all_style_property_bags() {
        let definitions = parse(
            r#"<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:latentStyles w:defLockedState="true" w:defUIPriority="99" w:defQFormat="1" w:count="3garbage"/><w:style w:type="table" w:styleId="Base"><w:pPr><w:spacing w:before="100"/></w:pPr><w:rPr><w:b/></w:rPr><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:i/></w:rPr></w:tblStylePr></w:style><w:style w:type="table" w:styleId="Child" w:default="true"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="200"/></w:pPr><w:tblStylePr w:type="firstRow"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:tblStylePr></w:style></w:styles>"#,
        );
        assert_eq!(
            definitions.doc_defaults.unwrap().r_pr.unwrap().font_size,
            Some(22.0)
        );
        let latent = definitions.latent_styles.unwrap();
        assert!(!latent.def_locked_state);
        assert!(latent.def_q_format);
        assert_eq!(latent.count, Some(3.0));
        let latent_json = serde_json::to_value(&latent).unwrap();
        assert_eq!(latent_json["defUIPriority"], 99.0);
        assert!(latent_json.get("defUiPriority").is_none());
        let child = definitions
            .styles
            .iter()
            .find(|style| style.style_id == "Child")
            .unwrap();
        let paragraph = child.p_pr.as_ref().unwrap();
        assert_eq!(paragraph.space_before, Some(100.0));
        assert_eq!(paragraph.space_after, Some(200.0));
        assert_eq!(child.r_pr.as_ref().unwrap().bold, Some(true));
        assert_eq!(
            child.tbl_pr.as_ref().unwrap().width.as_ref().unwrap().value,
            5000.0
        );
        let conditional = &child.tbl_style_pr.as_ref().unwrap()[0];
        assert_eq!(conditional.r_pr.as_ref().unwrap().italic, Some(true));
        assert_eq!(
            conditional
                .r_pr
                .as_ref()
                .unwrap()
                .color
                .as_ref()
                .unwrap()
                .rgb
                .as_deref(),
            Some("FF0000")
        );
    }

    #[test]
    fn cyclic_based_on_chains_terminate_with_incumbent_live_map_semantics() {
        let definitions = parse(
            r#"<w:styles><w:style w:styleId="A"><w:basedOn w:val="B"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style><w:style w:styleId="B"><w:basedOn w:val="A"/><w:pPr><w:spacing w:before="120"/></w:pPr></w:style></w:styles>"#,
        );
        for style in definitions.styles {
            let paragraph = style.p_pr.unwrap();
            assert_eq!(paragraph.alignment.as_deref(), Some("left"));
            assert_eq!(paragraph.space_before, Some(120.0));
        }
    }

    #[test]
    fn overlong_based_on_chain_returns_a_budget_error_without_recursing_unboundedly() {
        let mut xml = String::from("<w:styles>");
        for index in 0..70 {
            xml.push_str(&format!(
                "<w:style w:styleId=\"S{index}\"><w:basedOn w:val=\"S{}\"/></w:style>",
                index + 1
            ));
        }
        xml.push_str("<w:style w:styleId=\"S70\"/></w:styles>");
        let limits = ParseLimits::default();
        let error = parse_style_definitions(
            xml.as_bytes(),
            None,
            "word/styles.xml",
            &mut ParseBudget::new(&limits),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            ParseError::ResourceLimit {
                kind: "nestingDepth",
                ..
            }
        ));
    }
}
