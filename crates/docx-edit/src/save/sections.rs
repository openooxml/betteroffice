#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::collections::{HashMap, HashSet};

use docx_parse::serializer::{SectionEdit, SectionEdits};
use docx_parse::{BlockContent, DocumentBody, Paragraph, SectionProperties};
use serde::Serialize;
use serde::de::DeserializeOwned;
use serde_json::{Map, Value};
use sha2::{Digest, Sha256};

use crate::SectionInfo;
use crate::identity::{ParagraphIdentities, ParagraphOrigin, ParagraphRef, ParagraphSavePlan};

pub(crate) struct SectionPlanInput<'a> {
    pub(crate) source: Option<&'a DocumentBody>,
    pub(crate) output: &'a DocumentBody,
    pub(crate) output_session_keys: &'a [Option<String>],
    pub(crate) metadata_unchanged: bool,
    pub(crate) source_part: Option<&'a [u8]>,
}

#[derive(Debug, Serialize)]
pub(crate) struct SectionPlan {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) plan: Option<SectionEdits>,
    pub(crate) lossy: bool,
    pub(crate) strict: bool,
}

#[derive(Debug)]
pub(crate) struct SectionPlanError(pub(crate) &'static str);

pub(crate) fn from_json_any_depth<T: DeserializeOwned>(json: &str) -> serde_json::Result<T> {
    let mut de = serde_json::Deserializer::from_str(json);
    de.disable_recursion_limit();
    let value = T::deserialize(&mut de)?;
    de.end()?;
    Ok(value)
}

fn properties(section: &SectionProperties) -> Value {
    serde_json::to_value(section).expect("section properties serialize to JSON")
}

fn json_equal(left: &Value, right: &Value) -> bool {
    match (left, right) {
        (Value::Number(left), Value::Number(right)) => left.as_f64() == right.as_f64(),
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len() && left.iter().zip(right).all(|(a, b)| json_equal(a, b))
        }
        (Value::Object(left), Value::Object(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .all(|(key, a)| right.get(key).is_some_and(|b| json_equal(a, b)))
        }
        _ => left == right,
    }
}

fn changed(before: &Value, after: &Value) -> Option<Value> {
    const SUPPORTED: &[&str] = &[
        "orientation",
        "pageWidth",
        "pageHeight",
        "marginTop",
        "marginBottom",
        "marginLeft",
        "marginRight",
        "headerDistance",
        "footerDistance",
        "gutter",
        "sectionStart",
    ];
    let before = before.as_object()?;
    let after = after.as_object()?;
    if before.keys().any(|key| !after.contains_key(key)) {
        return None;
    }
    let mut patch = Map::new();
    for (key, value) in after {
        if before.get(key).is_some_and(|old| json_equal(old, value)) {
            continue;
        }
        if !SUPPORTED.contains(&key.as_str()) {
            return None;
        }
        patch.insert(key.clone(), value.clone());
    }
    Some(Value::Object(patch))
}

fn tag_name_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.' | b'-')
}

fn tag_name_end(xml: &[u8], start: usize) -> usize {
    start
        + xml[start..]
            .iter()
            .take_while(|&&byte| tag_name_byte(byte))
            .count()
}

fn tag_end(xml: &[u8], mut at: usize) -> Option<(usize, bool)> {
    while let Some(&byte) = xml.get(at) {
        match byte {
            b'>' => return Some((at + 1, false)),
            b'/' if xml.get(at + 1) == Some(&b'>') => return Some((at + 2, true)),
            b'\'' | b'"' => {
                at += xml[at + 1..].iter().position(|&next| next == byte)? + 1;
            }
            _ => {}
        }
        at += 1;
    }
    None
}

fn sections_modeled(xml: &[u8]) -> bool {
    const MODELED: &[&[u8]] = &[
        b"headerReference",
        b"footerReference",
        b"footnotePr",
        b"endnotePr",
        b"type",
        b"pgSz",
        b"pgMar",
        b"paperSrc",
        b"pgBorders",
        b"lnNumType",
        b"pgNumType",
        b"cols",
        b"vAlign",
        b"titlePg",
        b"bidi",
        b"docGrid",
    ];
    let mut depth = 0;
    let mut prefix = &[][..];
    let mut cursor = 0;
    while let Some(offset) = xml[cursor..].iter().position(|&byte| byte == b'<') {
        cursor += offset + 1;
        let close = xml.get(cursor) == Some(&b'/');
        let start = cursor + usize::from(close);
        let end = tag_name_end(xml, start);
        if start == end {
            continue;
        }
        let (space, local_start, local_end) = if xml.get(end) == Some(&b':') {
            let local_end = tag_name_end(xml, end + 1);
            if local_end > end + 1 {
                (&xml[start..end], end + 1, local_end)
            } else {
                (&[][..], start, end)
            }
        } else {
            (&[][..], start, end)
        };
        let Some((next, empty)) = tag_end(xml, local_end) else {
            continue;
        };
        cursor = next;
        let local = &xml[local_start..local_end];
        if depth == 0 {
            if local == b"sectPr" && !close && !empty {
                depth = 1;
                prefix = space;
            }
        } else if close {
            depth -= 1;
        } else {
            if depth == 1 && (space != prefix || !MODELED.contains(&local)) {
                return false;
            }
            if !empty {
                depth += 1;
            }
        }
    }
    true
}

fn visit<'a>(
    blocks: &'a [BlockContent],
    mut body_index: Option<usize>,
    paragraphs: &mut Vec<(&'a Paragraph, Option<usize>)>,
) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => {
                paragraphs.push((paragraph, body_index));
                body_index = body_index.map(|index| index + 1);
            }
            BlockContent::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        visit(&cell.content, None, paragraphs);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => visit(&sdt.content, None, paragraphs),
            BlockContent::RawXml(_) => {}
        }
    }
}

fn without_slash(part: &str) -> &str {
    part.strip_prefix('/').unwrap_or(part)
}

pub(crate) fn section_save_plan(
    input: SectionPlanInput<'_>,
    sections: &[SectionInfo],
    identities: &ParagraphIdentities,
    plan: &ParagraphSavePlan,
) -> Result<Option<SectionPlan>, SectionPlanError> {
    let Some(source) = input.source else {
        return Ok(None);
    };
    let planned_splices: Vec<_> = plan
        .spliced_parts
        .iter()
        .chain(&plan.section_spliced_parts)
        .collect();
    let structural = planned_splices
        .iter()
        .any(|part| !part.inserted.is_empty() || !part.removed.is_empty());
    let mut source_paragraphs = Vec::new();
    visit(&source.content, Some(0), &mut source_paragraphs);
    let mut by_ordinal = HashMap::new();
    let mut donors = Vec::new();
    let mut contiguous = true;
    for (index, (paragraph, body_index)) in source_paragraphs.into_iter().enumerate() {
        let ordinal = paragraph
            .source_ordinal
            .map_or(index, |ordinal| ordinal as usize);
        by_ordinal.insert(ordinal, paragraph);
        contiguous &= ordinal == index;
        if let (Some(_), Some(section)) = (body_index, &paragraph.section_properties) {
            donors.push((ordinal, section));
        }
    }
    let mut original_by_key = HashMap::new();
    let mut source_path = None;
    for identity in &identities.paragraphs {
        if let (ParagraphRef::Session { story, para_id }, Some(source)) =
            (&identity.paragraph, &identity.source)
            && story == "body"
        {
            original_by_key.insert(para_id.as_str(), source.paragraph_ordinal as usize);
            source_path.get_or_insert_with(|| without_slash(&source.part_uri));
        }
    }
    let part = source_path.unwrap_or("word/document.xml");
    let empty = Value::Object(Map::new());
    let source_final = source
        .final_section_properties
        .as_ref()
        .map_or_else(|| empty.clone(), properties);
    let prior = |ordinal| {
        by_ordinal
            .get(&ordinal)
            .and_then(|paragraph| paragraph.section_properties.as_ref())
    };
    let section_changes = sections.iter().any(|section| {
        if section.closing_paragraph_id == "final" {
            return !json_equal(&source_final, &section.authored);
        }
        original_by_key
            .get(section.closing_paragraph_id.as_str())
            .copied()
            .and_then(prior)
            .is_none_or(|before| !json_equal(&properties(before), &section.authored))
    });
    if !section_changes {
        return Ok(None);
    }

    let mut output_paragraphs = Vec::new();
    visit(&input.output.content, Some(0), &mut output_paragraphs);
    let body_paragraphs = output_paragraphs
        .iter()
        .filter(|(_, body_index)| body_index.is_some())
        .count();
    let original_at = |index: usize| {
        input
            .output_session_keys
            .get(index)
            .and_then(Option::as_deref)
            .and_then(|key| original_by_key.get(key).copied())
    };
    let neighbor_at = |index: usize| {
        let key = input
            .output_session_keys
            .get(index)
            .and_then(Option::as_deref);
        original_by_key.get(key.unwrap_or("")).copied()
    };
    let mut edits = Vec::new();
    let mut source_positions = Vec::new();
    let mut legacy = false;
    let mut added_boundary = false;
    for (at, (paragraph, body_index)) in output_paragraphs.iter().enumerate() {
        let (Some(body_index), Some(after)) = (body_index, &paragraph.section_properties) else {
            continue;
        };
        let after = properties(after);
        let ordinal = original_at(at);
        let before = ordinal.and_then(prior);
        if before.is_some_and(|before| json_equal(&properties(before), &after)) {
            continue;
        }
        added_boundary |= before.is_none();
        let mut donor = ordinal.zip(before).or_else(|| {
            let following = (at..output_paragraphs.len()).find_map(neighbor_at)?;
            donors
                .iter()
                .copied()
                .find(|(ordinal, _)| *ordinal >= following)
        });
        let donor_properties = donor.map(|(_, section)| properties(section));
        let mut patch = changed(donor_properties.as_ref().unwrap_or(&source_final), &after);
        if patch.is_none()
            && ordinal.is_none()
            && let Some(preceding) = (0..at).rev().find_map(neighbor_at)
            && let Some(previous) = donors
                .iter()
                .copied()
                .find(|(ordinal, _)| *ordinal == preceding)
            && let Some(inherited) = changed(&properties(previous.1), &after)
        {
            donor = Some(previous);
            patch = Some(inherited);
        }
        let Some(patch) = patch else {
            legacy = true;
            continue;
        };
        edits.push(SectionEdit {
            paragraph: Some(*body_index),
            body_paragraphs: Some(body_paragraphs),
            donor: donor.map(|(ordinal, _)| ordinal),
            properties: patch,
        });
        source_positions.push((!structural).then_some(ordinal.unwrap_or(at)));
    }
    let output_final = input
        .output
        .final_section_properties
        .as_ref()
        .map_or(empty, properties);
    if !json_equal(&source_final, &output_final) {
        if let Some(patch) = changed(&source_final, &output_final) {
            edits.push(SectionEdit {
                paragraph: None,
                body_paragraphs: None,
                donor: None,
                properties: patch,
            });
        } else {
            legacy = true;
        }
    }
    let lossy = || added_boundary && !input.source_part.is_some_and(sections_modeled);
    let ordinary = |message| {
        if lossy() {
            return Err(SectionPlanError(message));
        }
        Ok(Some(SectionPlan {
            plan: None,
            lossy: false,
            strict: false,
        }))
    };
    if legacy {
        return ordinary("Cannot safely compose section patches with legacy section properties");
    }
    if edits.is_empty() {
        return Ok(None);
    }
    let Some(bytes) = input.source_part else {
        return ordinary("Section source part is unavailable");
    };
    let known_parts: HashSet<_> = std::iter::once(part)
        .chain(identities.paragraphs.iter().filter_map(|identity| {
            identity
                .source
                .as_ref()
                .map(|source| without_slash(&source.part_uri))
        }))
        .chain(planned_splices.iter().map(|splice| splice.part.as_str()))
        .collect();
    let only_sections = input.metadata_unchanged
        && identities.paragraphs.iter().all(|identity| {
            matches!(&identity.paragraph, ParagraphRef::Source(_))
                || identity.source.is_some()
                || identity.origin == ParagraphOrigin::Synthetic
        })
        && plan.assignments.is_empty()
        && plan.patched_parts.iter().all(|(_, ids)| ids.is_empty())
        && known_parts.into_iter().all(|path| {
            plan.patched_parts.iter().any(|(part, _)| part == path)
                || planned_splices
                    .iter()
                    .find(|splice| splice.part == path)
                    .is_some_and(|splice| {
                        path == part
                            && splice.inserted.is_empty()
                            && splice.removed.is_empty()
                            && splice.table_patches.is_empty()
                            && splice
                                .changed
                                .iter()
                                .all(|ordinal| splice.section_only.contains(ordinal))
                    })
        });
    let strict = structural && !contiguous && edits.iter().any(|edit| edit.paragraph.is_some());
    if only_sections {
        for (edit, source_position) in edits.iter_mut().zip(source_positions) {
            if let Some(position) = source_position {
                edit.paragraph = Some(position);
                edit.body_paragraphs = None;
            }
        }
    }
    Ok(Some(SectionPlan {
        plan: Some(SectionEdits {
            only_sections,
            sha256: format!("{:x}", Sha256::digest(bytes)),
            edits,
        }),
        lossy: lossy(),
        strict,
    }))
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use docx_parse::{Table, TableCell, TableRow};
    use serde_json::json;

    use super::*;

    #[test]
    fn deserializes_a_body_with_deeply_nested_tables() {
        std::thread::Builder::new()
            .stack_size(64 << 20)
            .spawn(|| {
                let paragraph = BlockContent::Paragraph(Arc::new(Paragraph {
                    node_type: "paragraph".to_owned(),
                    para_id: None,
                    repeated_para_id: None,
                    para_id_attribute: None,
                    source_ordinal: None,
                    text_id: None,
                    extra_attributes: Vec::new(),
                    formatting: None,
                    property_changes: None,
                    p_pr_ins: None,
                    p_pr_del: None,
                    content: Vec::new(),
                    list_rendering: None,
                    rendered_page_break_before: None,
                    section_properties: None,
                }));
                let mut content = Vec::new();
                for _ in 0..260 {
                    content.push(paragraph.clone());
                    content = vec![BlockContent::Table(Arc::new(Table {
                        rows: vec![TableRow {
                            node_type: "tableRow".to_owned(),
                            extra_attributes: None,
                            formatting: None,
                            property_changes: None,
                            structural_change: None,
                            cells: vec![TableCell {
                                node_type: "tableCell".to_owned(),
                                formatting: None,
                                property_changes: None,
                                structural_change: None,
                                content,
                                leading_bookmarks: None,
                            }],
                            leading_bookmarks: None,
                            trailing_bookmarks: None,
                        }],
                        ..Table::empty()
                    }))];
                }
                let body = DocumentBody {
                    content,
                    ..DocumentBody::default()
                };
                let json = serde_json::to_string(&body).unwrap();
                let error = serde_json::from_str::<DocumentBody>(&json).unwrap_err();
                assert!(error.to_string().contains("recursion limit exceeded"), "{error}");
                let decoded: DocumentBody = from_json_any_depth(&json).unwrap();
                assert_eq!(serde_json::to_string(&decoded).unwrap(), json);
            })
            .unwrap()
            .join()
            .unwrap();
    }

    #[test]
    fn scans_section_tags_with_the_same_lexical_rules_inside_and_outside_xml() {
        for xml in [
            r#"<w:sectPr x="> /"><w:pgSz a='/> <w:unknown>'/><w:cols><u:unknown/></w:cols></w:sectPr>"#,
            "<a.B-_:sectPr><a.B-_:type/></a.B-_:sectPr><sectPr><docGrid/></sectPr>",
            "<w:sectPr/><w:unknown/><w:sectPr />",
            "<w:sectPr><w:pgSz/ ><w:unknown/></w:pgSz></w:sectPr>",
            "<w:sectPr><w:pgSz:extra/></not-the-section><w:unknown/>",
            "<w:sectPr><w:pgSz <w:unknown/></w:sectPr>",
            "<w:sectPr></anything/><w:unknown/>",
        ] {
            assert!(sections_modeled(xml.as_bytes()), "{xml}");
        }
        for xml in [
            "<w:sectPr><other:pgSz/></w:sectPr>",
            "<w:sectPr><pgSz/></w:sectPr>",
            "<w:sectPr><w:pgSz/><w:unknown/></w:sectPr>",
            "<w:sectPr><!-- <w:unknown/> --></w:sectPr>",
            r#"<w:sectPr><w:pgSz foo="unterminated><w:unknown/></w:sectPr>"#,
            "<w:sectPr><é:pgSz/><w:unknown/></w:sectPr>",
        ] {
            assert!(!sections_modeled(xml.as_bytes()), "{xml}");
        }
    }

    #[test]
    fn compares_serialized_section_properties_as_json() {
        let section = SectionProperties {
            margin_left: Some(1440.0),
            columns: Some(vec![
                docx_parse::section::Column {
                    width: Some(2000.0),
                    space: None,
                },
                docx_parse::section::Column {
                    width: None,
                    space: Some(-0.0),
                },
            ]),
            ..SectionProperties::default()
        };
        let section = properties(&section);
        assert!(json_equal(
            &section,
            &json!({"columns": [{"width": 2000}, {"space": 0}], "marginLeft": 1440}),
        ));
        assert!(!json_equal(
            &section,
            &json!({"columns": [{"space": 0}, {"width": 2000}], "marginLeft": 1440}),
        ));
        assert!(!json_equal(
            &section,
            &json!({"columns": [{"width": 2000, "space": null}, {"space": 0}], "marginLeft": 1440}),
        ));
        assert_eq!(
            changed(
                &section,
                &json!({"columns": [{"width": 2000}, {"space": 0}], "marginLeft": 1800}),
            ),
            Some(json!({"marginLeft": 1800})),
        );
        assert_eq!(changed(&section, &json!({"marginLeft":1440})), None);
    }
}

