use serde::Serialize;
use serde_json::{Value, json};
use yrs::types::text::YChange;
use yrs::{Any, Map, Out, ReadTxn, Text};

use crate::{EditCtx, EditingDoc, map_string, story_ref};

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SectionInfo {
    pub closing_paragraph_id: String,
    pub first_paragraph: usize,
    pub last_paragraph: usize,
    pub authored: Value,
    pub resolved: Value,
}

pub(crate) struct SectionPlan {
    pub paragraphs: Vec<(String, Value)>,
    pub final_properties: Option<Value>,
    pub append: Option<crate::ops::paragraph::ParagraphRecord>,
}

fn properties(values: &std::collections::BTreeMap<String, Any>) -> Option<Value> {
    values
        .get("sectPr")
        .or_else(|| values.get("_sectionProperties"))
        .and_then(|value| serde_json::to_value(value).ok())
        .filter(Value::is_object)
        .or_else(|| {
            values
                .get("sectionBreakType")
                .and_then(|value| serde_json::to_value(value).ok())
                .filter(Value::is_string)
                .map(|kind| json!({"sectionStart":kind}))
        })
}

impl EditingDoc {
    pub fn list_sections(&self) -> Result<Vec<SectionInfo>, String> {
        let paragraphs = self.paragraphs("body").map_err(|error| error.to_string())?;
        let margins: Value = serde_json::from_str(&self.section_margins_json())
            .map_err(|error| error.to_string())?;
        let mut result = Vec::new();
        let mut first = 1;
        for (index, paragraph) in paragraphs.iter().enumerate() {
            if let Some(mut authored) = properties(&paragraph.properties) {
                if let Some(patch) = margins.get(&paragraph.para_id).and_then(Value::as_object) {
                    authored.as_object_mut().unwrap().extend(patch.clone());
                }
                result.push(SectionInfo {
                    closing_paragraph_id: paragraph.para_id.clone(),
                    first_paragraph: first,
                    last_paragraph: index + 1,
                    resolved: authored.clone(),
                    authored,
                });
                first = index + 2;
            }
        }
        let authored = self.final_section_properties().unwrap_or_else(|| json!({}));
        result.push(SectionInfo {
            closing_paragraph_id: "final".to_owned(),
            first_paragraph: first,
            last_paragraph: paragraphs.len(),
            resolved: authored.clone(),
            authored,
        });
        let mut resolved: Vec<docx_parse::SectionProperties> = result
            .iter()
            .map(|section| {
                let mut value =
                    serde_json::to_value(docx_parse::section::default_section_properties())
                        .map_err(|error| error.to_string())?;
                value
                    .as_object_mut()
                    .unwrap()
                    .extend(section.authored.as_object().unwrap().clone());
                serde_json::from_value(value).map_err(|error| error.to_string())
            })
            .collect::<Result<_, _>>()?;
        docx_parse::section::apply_section_inheritance(&mut resolved);
        for (section, properties) in result.iter_mut().zip(resolved) {
            section.resolved =
                serde_json::to_value(properties).map_err(|error| error.to_string())?;
        }
        Ok(result)
    }

    pub(crate) fn plan_section_break(
        &self,
        para_id: &str,
        kind: &str,
    ) -> Result<SectionPlan, String> {
        if !["nextPage", "continuous", "evenPage", "oddPage"].contains(&kind) {
            return Err("invalid section start kind".to_owned());
        }
        let paragraphs = self.paragraphs("body").map_err(|error| error.to_string())?;
        let index = paragraphs
            .iter()
            .position(|paragraph| paragraph.para_id == para_id)
            .ok_or("section breaks require a body paragraph")?;
        if properties(&paragraphs[index].properties).is_some() {
            return Err("the paragraph already closes a section".to_owned());
        }
        let section = self
            .list_sections()?
            .into_iter()
            .find(|section| section.last_paragraph > index)
            .ok_or("section unavailable")?;
        let mut following = section.authored.clone();
        following["sectionStart"] = json!(kind);
        let mut writes = vec![(para_id.to_owned(), section.authored)];
        let final_properties = if section.closing_paragraph_id == "final" {
            Some(following)
        } else {
            writes.push((section.closing_paragraph_id, following));
            None
        };
        let append = if index + 1 == paragraphs.len() {
            let styled = self
                .source_metadata()
                .ok_or("section source unavailable")?
                .styled_paragraph(None)?;
            Some(crate::ops::paragraph::ParagraphRecord {
                text: String::new(),
                properties: styled.properties,
                run: styled.run,
            })
        } else {
            None
        };
        Ok(SectionPlan {
            paragraphs: writes,
            final_properties,
            append,
        })
    }

    pub(crate) fn plan_section_properties(
        &self,
        section: &str,
        patch: &Value,
    ) -> Result<SectionPlan, String> {
        let current = self
            .list_sections()?
            .into_iter()
            .find(|entry| entry.closing_paragraph_id == section)
            .ok_or("unknown section")?;
        let mut next = current.authored.clone();
        let patch = patch
            .as_object()
            .ok_or("section properties must be an object")?;
        if patch.is_empty() {
            return Err("supply at least one section property".to_owned());
        }
        for (key, value) in patch {
            match key.as_str() {
                "orientation" if matches!(value.as_str(), Some("portrait" | "landscape")) => {
                    if current.authored["orientation"]
                        .as_str()
                        .unwrap_or("portrait")
                        != value.as_str().unwrap()
                    {
                        next[key] = value.clone();
                    }
                }
                "pageWidth" | "pageHeight"
                    if value.as_f64().is_some_and(|value| {
                        value.is_finite() && value > 0.0 && value.fract() == 0.0
                    }) =>
                {
                    next[key] = value.clone();
                }
                "margins" => {
                    let margins = value.as_object().ok_or("margins must be an object")?;
                    for (name, value) in margins {
                        let field = match name.as_str() {
                            "top" => "marginTop",
                            "bottom" => "marginBottom",
                            "left" => "marginLeft",
                            "right" => "marginRight",
                            "header" => "headerDistance",
                            "footer" => "footerDistance",
                            "gutter" => "gutter",
                            _ => return Err("unknown margin".to_owned()),
                        };
                        if !value.as_f64().is_some_and(|value| {
                            value.is_finite() && value >= 0.0 && value.fract() == 0.0
                        }) {
                            return Err("margins must be nonnegative integer twips".to_owned());
                        }
                        next[field] = value.clone();
                    }
                }
                _ => return Err("invalid section property".to_owned()),
            }
        }
        if let Some(orientation) = patch.get("orientation").and_then(Value::as_str) {
            let width = next["pageWidth"].as_f64().unwrap_or(12240.0);
            let height = next["pageHeight"].as_f64().unwrap_or(15840.0);
            if current.authored["orientation"]
                .as_str()
                .unwrap_or("portrait")
                != orientation
            {
                let (width, height) = if (orientation == "landscape") == (width < height) {
                    (height, width)
                } else {
                    (width, height)
                };
                next["pageWidth"] = json!(width);
                next["pageHeight"] = json!(height);
            }
        }
        let number = |key: &str, default: f64| next[key].as_f64().unwrap_or(default);
        if number("marginLeft", 1440.0) + number("marginRight", 1440.0) + number("gutter", 0.0)
            >= number("pageWidth", 12240.0)
            || number("marginTop", 1440.0) + number("marginBottom", 1440.0)
                >= number("pageHeight", 15840.0)
        {
            return Err("section margins leave no usable page area".to_owned());
        }
        if current.authored == next {
            return Ok(SectionPlan {
                paragraphs: Vec::new(),
                final_properties: None,
                append: None,
            });
        }
        Ok(if section == "final" {
            SectionPlan {
                paragraphs: Vec::new(),
                final_properties: Some(next),
                append: None,
            }
        } else {
            SectionPlan {
                paragraphs: vec![(section.to_owned(), next)],
                final_properties: None,
                append: None,
            }
        })
    }

    pub(crate) fn apply_section_plan(&self, plan: &SectionPlan) -> Result<Vec<String>, String> {
        let mut metadata: Value =
            serde_json::from_str(&self.part_metadata_json()).map_err(|error| error.to_string())?;
        if let Some(properties) = &plan.final_properties {
            if !metadata["body"].is_object() {
                metadata["body"] = json!({});
            }
            metadata["body"]["finalSectionProperties"] = properties.clone();
            if let Some(last) = metadata["body"]["sections"]
                .as_array_mut()
                .and_then(|sections| sections.last_mut())
            {
                last["properties"] = properties.clone();
            }
        }
        let mut txn = self.transact_for(&EditCtx::local("", ""));
        let story = story_ref(&txn, "body").map_err(|error| error.to_string())?;
        let margins = txn
            .get_map(crate::section_margins::SECTION_MARGINS)
            .unwrap();
        for diff in story.diff(&txn, YChange::identity) {
            if let Out::YMap(map) = diff.insert {
                if let Some((id, properties)) = plan.paragraphs.iter().find(|(id, _)| {
                    map_string(&map, &txn, crate::PARA_ID).as_deref() == Some(id.as_str())
                }) {
                    crate::identity::promote(self, &mut txn, &map);
                    let value = Any::from_json(&properties.to_string())
                        .map_err(|error| error.to_string())?;
                    map.insert(&mut txn, "sectPr", value);
                    let start = properties
                        .get("sectionStart")
                        .cloned()
                        .unwrap_or(Value::Null);
                    map.insert(
                        &mut txn,
                        "sectionBreakType",
                        Any::from_json(&start.to_string()).map_err(|error| error.to_string())?,
                    );
                    margins.remove(&mut txn, id);
                }
            }
        }
        if plan.final_properties.is_some() {
            txn.get_map(crate::part_metadata::PART_METADATA)
                .unwrap()
                .insert(&mut txn, "metadata", metadata.to_string());
            margins.remove(&mut txn, "final");
        }
        drop(txn);
        if let Some(paragraph) = &plan.append {
            return self
                .insert_paragraph_records(
                    "body",
                    self.story_len("body").map_err(|error| error.to_string())?,
                    std::slice::from_ref(paragraph),
                )
                .map_err(|error| error.to_string());
        }
        Ok(Vec::new())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn section_property_plans_normalize_once_and_preserve_history_and_replication() {
        let doc = EditingDoc::new(7101);
        doc.create_story("body", "Text", "Normal", "left").unwrap();
        doc.set_part_metadata(
            &EditCtx::local("", ""),
            r#"{"body":{"finalSectionProperties":{"pageWidth":12240,"pageHeight":15840}}}"#,
            &[],
        )
        .unwrap();
        let mut undo = doc.undo_manager();
        let plan = doc
            .plan_section_properties(
                "final",
                &json!({"orientation":"landscape","margins":{"left":1800}}),
            )
            .unwrap();
        doc.apply_section_plan(&plan).unwrap();
        let expected = doc.list_sections().unwrap()[0].authored.clone();
        assert_eq!(expected["pageWidth"], 15840.0);
        assert_eq!(expected["pageHeight"], 12240.0);
        let repeated = doc
            .plan_section_properties("final", &json!({"orientation":"landscape"}))
            .unwrap();
        assert!(repeated.final_properties.is_none());
        assert!(undo.undo());
        assert!(
            doc.list_sections().unwrap()[0]
                .authored
                .get("orientation")
                .is_none()
        );
        assert!(undo.redo());
        let peer = EditingDoc::new(7102);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        assert_eq!(peer.list_sections().unwrap()[0].authored, expected);
        assert_eq!(doc.paragraphs("body").unwrap()[0].text, "Text");
    }

    #[test]
    fn section_planning_refuses_invalid_twips_without_mutating() {
        let doc = EditingDoc::new(7103);
        doc.create_story("body", "Text", "Normal", "left").unwrap();
        let before = doc.encode_state_as_update_v1();
        for properties in [
            json!({"pageWidth":-1}),
            json!({"margins":{"left":20000}}),
            json!({"orientation":"sideways"}),
            json!({"columns":2}),
        ] {
            assert!(doc.plan_section_properties("final", &properties).is_err());
            assert_eq!(doc.encode_state_as_update_v1(), before);
        }
    }
}
