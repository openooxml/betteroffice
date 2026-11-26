#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::collections::{HashMap, HashSet};

use docx_parse::serializer::{SectionEdit, SectionEdits};
use docx_parse::{BlockContent, DocumentBody, Paragraph, SectionProperties};
use serde::Serialize;
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

#[cfg(any(test, feature = "wasm"))]
pub(crate) mod transport {
    use std::sync::Arc;

    use serde::Deserialize;

    use super::{BlockContent, DocumentBody, SectionProperties};

    #[derive(Deserialize)]
    #[serde(rename_all = "camelCase")]
    pub(crate) struct BodyFragments {
        content: String,
        final_section_properties: Option<SectionProperties>,
        blocks: Vec<BlockFragment>,
    }

    #[derive(Deserialize)]
    struct BlockFragment {
        parent: usize,
        slot: BlockSlot,
        json: String,
    }

    #[derive(Deserialize)]
    struct BlockSlot {
        block: usize,
        cell: Option<[usize; 2]>,
    }

    fn block_slot(blocks: &mut [BlockContent], slot: BlockSlot) -> Option<&mut Vec<BlockContent>> {
        match (blocks.get_mut(slot.block)?, slot.cell) {
            (BlockContent::Table(table), Some([row, cell])) => {
                let row = Arc::get_mut(table)?.rows.get_mut(row)?;
                Some(&mut row.cells.get_mut(cell)?.content)
            }
            (BlockContent::BlockSdt(sdt), None) => Some(&mut Arc::get_mut(sdt)?.content),
            _ => None,
        }
    }

    type Children = Vec<(BlockSlot, Vec<BlockContent>)>;

    fn rebuild_blocks(json: &str, children: Children) -> Result<Vec<BlockContent>, String> {
        let mut blocks: Vec<BlockContent> =
            serde_json::from_str(json).map_err(|error| error.to_string())?;
        for (slot, content) in children {
            *block_slot(&mut blocks, slot).ok_or("Invalid block fragment slot")? = content;
        }
        Ok(blocks)
    }

    impl BodyFragments {
        pub(crate) fn rebuild(self) -> Result<DocumentBody, String> {
            let slots = self.blocks.len() + 1;
            let mut children: Vec<Children> = (0..slots).map(|_| Vec::new()).collect();
            for fragment in self.blocks.into_iter().rev() {
                let nested = children.pop().expect("one slot per fragment");
                let blocks = rebuild_blocks(&fragment.json, nested)?;
                children
                    .get_mut(fragment.parent)
                    .ok_or("Invalid block fragment parent")?
                    .push((fragment.slot, blocks));
            }
            let nested = children.pop().expect("a slot for the body");
            Ok(DocumentBody {
                content: rebuild_blocks(&self.content, nested)?,
                final_section_properties: self.final_section_properties,
                ..DocumentBody::default()
            })
        }
    }

    #[cfg(test)]
    mod tests {
        use serde_json::json;

        use super::*;

        #[test]
        fn rebuilds_a_body_with_deeply_nested_tables() {
            let body = json!({"content": [], "finalSectionProperties": {"marginLeft": 1440.0}});
            let table = json!([{"type": "table", "rows": [{
                "type": "tableRow", "cells": [{"type": "tableCell", "content": []}]
            }]}]);
            let leaf = json!([{"type": "paragraph", "content": [], "sourceOrdinal": 1}]);
            let depth = 260;
            let fragments = BodyFragments {
                content: table.to_string(),
                final_section_properties: Some(SectionProperties {
                    margin_left: Some(1440.0),
                    ..SectionProperties::default()
                }),
                blocks: (0..depth)
                    .map(|parent| {
                        let value = if parent + 1 == depth { &leaf } else { &table };
                        BlockFragment {
                            parent,
                            slot: BlockSlot {
                                block: 0,
                                cell: Some([0, 0]),
                            },
                            json: value.to_string(),
                        }
                    })
                    .collect(),
            };
            let mut rebuilt = fragments.rebuild().unwrap();
            let mut blocks = std::mem::take(&mut rebuilt.content);
            assert_eq!(serde_json::to_value(&rebuilt).unwrap(), body);
            for _ in 0..depth {
                let BlockContent::Table(table_block) = &mut blocks[0] else {
                    panic!("expected a table");
                };
                let cell = &mut Arc::get_mut(table_block).unwrap().rows[0].cells[0];
                let nested = std::mem::take(&mut cell.content);
                assert_eq!(serde_json::to_value(&blocks).unwrap(), table);
                blocks = nested;
            }
            assert_eq!(serde_json::to_value(&blocks).unwrap(), leaf);
        }
    }

    #[cfg(feature = "wasm")]
    pub(crate) use snapshot::plan_json;

    #[cfg(feature = "wasm")]
    mod snapshot {
        use docx_parse::serializer::{S13ParagraphId, S13PatchedPart, S13SpliceAnchor, S13TablePatch};
        use serde::Deserialize;

        use crate::SectionInfo;
        use crate::identity::{
            ParagraphIdentities, ParagraphIdentity, ParagraphOrigin, ParagraphRef,
            ParagraphSavePlan, SourceParagraphRef, SpliceAnchor, SplicedPart,
        };
        use crate::save::sections::{SectionPlanInput, section_save_plan};

        use super::BodyFragments;

        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct Input {
            source: Option<BodyFragments>,
            output: BodyFragments,
            output_session_keys: Vec<Option<String>>,
            metadata_unchanged: bool,
            source_part: Option<Vec<u8>>,
            sections: Vec<SectionInfo>,
            identities: Identities,
            plan: SavePlan,
        }

        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct Identities {
            session_id: String,
            package_sha256: Option<String>,
            paragraphs: Vec<Identity>,
        }

        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct Identity {
            session: Option<SessionParagraph>,
            source: Option<SourceParagraph>,
            origin: Origin,
        }

        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct SessionParagraph {
            story: String,
            para_id: String,
        }

        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct SourceParagraph {
            package_sha256: String,
            part_uri: String,
            paragraph_ordinal: u32,
        }

        #[derive(Deserialize)]
        #[serde(rename_all = "lowercase")]
        enum Origin {
            Source,
            Authored,
            Synthetic,
        }

        impl TryFrom<Identities> for ParagraphIdentities {
            type Error = &'static str;

            fn try_from(input: Identities) -> Result<Self, Self::Error> {
                let paragraphs = input
                    .paragraphs
                    .into_iter()
                    .map(|identity| {
                        let source = identity.source.map(|source| SourceParagraphRef {
                            package_sha256: source.package_sha256,
                            part_uri: source.part_uri,
                            paragraph_ordinal: source.paragraph_ordinal,
                        });
                        let paragraph = match identity.session {
                            Some(session) => ParagraphRef::Session {
                                story: session.story,
                                para_id: session.para_id,
                            },
                            None => ParagraphRef::Source(
                                source.clone().ok_or("Missing paragraph source identity")?,
                            ),
                        };
                        Ok(ParagraphIdentity {
                            paragraph,
                            origin: match identity.origin {
                                Origin::Source => ParagraphOrigin::Source,
                                Origin::Authored => ParagraphOrigin::Authored,
                                Origin::Synthetic => ParagraphOrigin::Synthetic,
                            },
                            source,
                            ooxml_para_id: None,
                            id_origin: None,
                            source_story: None,
                        })
                    })
                    .collect::<Result<_, Self::Error>>()?;
                Ok(Self {
                    session_id: input.session_id,
                    package_sha256: input.package_sha256,
                    paragraphs,
                })
            }
        }

        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct SavePlan {
            assignments: Vec<S13ParagraphId>,
            patched_parts: Vec<S13PatchedPart>,
            #[serde(default)]
            spliced_parts: Vec<Splice>,
            #[serde(default)]
            section_spliced_parts: Vec<Splice>,
        }

        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct Splice {
            part: String,
            changed: Vec<u32>,
            #[serde(default)]
            section_only: Vec<u32>,
            #[serde(default)]
            inserted: Vec<S13SpliceAnchor>,
            #[serde(default)]
            removed: Vec<u32>,
            #[serde(default)]
            table_patches: Vec<S13TablePatch>,
        }

        impl From<Splice> for SplicedPart {
            fn from(splice: Splice) -> Self {
                Self {
                    part: splice.part,
                    changed: splice.changed,
                    section_only: splice.section_only,
                    inserted: splice
                        .inserted
                        .into_iter()
                        .map(|anchor| match anchor {
                            S13SpliceAnchor::Before(ordinal) => SpliceAnchor::Before(ordinal),
                            S13SpliceAnchor::After(ordinal) => SpliceAnchor::After(ordinal),
                        })
                        .collect(),
                    removed: splice.removed,
                    table_patches: splice.table_patches,
                    ..Self::default()
                }
            }
        }

        impl From<SavePlan> for ParagraphSavePlan {
            fn from(plan: SavePlan) -> Self {
                Self {
                    assignments: plan
                        .assignments
                        .into_iter()
                        .map(|id| (id.part, id.ordinal, id.para_id))
                        .collect(),
                    patched_parts: plan
                        .patched_parts
                        .into_iter()
                        .map(|part| (part.part, part.para_ids))
                        .collect(),
                    spliced_parts: plan.spliced_parts.into_iter().map(Into::into).collect(),
                    section_spliced_parts: plan
                        .section_spliced_parts
                        .into_iter()
                        .map(Into::into)
                        .collect(),
                    ..Self::default()
                }
            }
        }

        pub(crate) fn plan_json(json: &str) -> Result<String, String> {
            let input: Input = serde_json::from_str(json).map_err(|error| error.to_string())?;
            let source = input.source.map(BodyFragments::rebuild).transpose()?;
            let output = input.output.rebuild()?;
            let identities = ParagraphIdentities::try_from(input.identities)?;
            let plan = input.plan.into();
            let result = section_save_plan(
                SectionPlanInput {
                    source: source.as_ref(),
                    output: &output,
                    output_session_keys: &input.output_session_keys,
                    metadata_unchanged: input.metadata_unchanged,
                    source_part: input.source_part.as_deref(),
                },
                &input.sections,
                &identities,
                &plan,
            );
            match result {
                Ok(plan) => serde_json::to_string(&plan).map_err(|error| error.to_string()),
                Err(error) => Ok(serde_json::json!({ "error": error.0 }).to_string()),
            }
        }
    }
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

    use docx_parse::section::parse_section_properties;
    use docx_parse::serializer::{S13ParagraphIds, S13SaveRequest, write_docx_s13_parts};
    use docx_parse::xml::{ParseBudget, ParseLimits, parse_xml};
    use serde_json::json;

    use super::*;
    use crate::identity::{ParagraphIdentity, SourceParagraphRef};

    #[test]
    fn empty_page_border_transforms_allow_a_margin_only_patch() {
        assert_margin_only_patch_matches_whole_write(|authored| {
            authored["pageBorders"]["top"]["color"]["transforms"] = json!([]);
        });
    }

    #[test]
    fn null_column_count_allows_a_margin_only_patch() {
        assert_margin_only_patch_matches_whole_write(|authored| {
            authored["columnCount"] = Value::Null;
        });
    }

    fn assert_margin_only_patch_matches_whole_write(edit_authored: impl FnOnce(&mut Value)) {
        let original = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p/>
        <w:sectPr>
            <w:pgMar w:left='1440'/>
            <w:pgBorders>
                <w:top w:color='ABCDEF' w:val='single'/>
            </w:pgBorders>
        </w:sectPr>
    </w:body>
</w:document>"#;
        let part = "word/document.xml";
        let final_properties = |bytes: &[u8]| {
            let limits = ParseLimits::default();
            let mut budget = ParseBudget::new(&limits);
            let xml = parse_xml(bytes, part, &mut budget).unwrap();
            let body = xml.root().unwrap().child("w", "body").unwrap();
            parse_section_properties(Some(body.child("w", "sectPr").unwrap()))
        };
        let mut model = json!({
            "content": [{"type": "paragraph", "content": []}],
            "finalSectionProperties": final_properties(original.as_bytes()),
        });
        let source: DocumentBody = serde_json::from_value(model.clone()).unwrap();
        let authored = &mut model["finalSectionProperties"];
        authored["marginLeft"] = json!(1800);
        edit_authored(authored);
        let sections = [SectionInfo {
            closing_paragraph_id: "final".to_owned(),
            first_paragraph: 1,
            last_paragraph: 1,
            authored: authored.clone(),
            resolved: authored.clone(),
        }];
        let output: DocumentBody = serde_json::from_value(model).unwrap();
        let plan = section_save_plan(
            SectionPlanInput {
                source: Some(&source),
                output: &output,
                output_session_keys: &[None],
                metadata_unchanged: true,
                source_part: Some(original.as_bytes()),
            },
            &sections,
            &ParagraphIdentities {
                session_id: "session".to_owned(),
                package_sha256: None,
                paragraphs: Vec::new(),
            },
            &ParagraphSavePlan {
                patched_parts: vec![(part.to_owned(), Vec::new())],
                ..ParagraphSavePlan::default()
            },
        )
        .unwrap()
        .unwrap();
        let edits = plan.plan.unwrap();
        assert!(edits.only_sections);
        let [edit] = edits.edits.as_slice() else {
            panic!("expected one section edit");
        };
        assert_eq!(edit.paragraph, None);
        assert!(json_equal(&edit.properties, &json!({"marginLeft": 1800})));

        let whole: S13SaveRequest = serde_json::from_value(json!({
            "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
            "document": output,
            "options": {"updateModifiedDate": false},
        }))
        .unwrap();
        let patch = S13SaveRequest {
            paragraph_ids: Some(S13ParagraphIds {
                section_edits: Some(edits),
                ..S13ParagraphIds::default()
            }),
            ..whole.clone()
        };
        let original_parts = [(part.to_owned(), original.as_bytes().to_vec())];
        let write_body = |request| {
            let saved = write_docx_s13_parts(request, &original_parts, None).unwrap();
            ooxml_opc::unzip_parts(&saved)
                .unwrap()
                .into_iter()
                .find(|(path, _)| path == part)
                .unwrap()
                .1
        };
        let patched = write_body(patch);
        let written = write_body(whole);
        assert_eq!(final_properties(&patched), final_properties(&written));
        assert_eq!(
            patched.as_slice(),
            original.replace("w:left='1440'", "w:left='1800'").as_bytes(),
        );
        let patched_bytes = original
            .bytes()
            .zip(patched)
            .filter(|(before, after)| before != after)
            .count();
        assert!(patched_bytes < original.len().abs_diff(written.len()));
    }

    #[test]
    fn shared_paragraph_edits_use_each_body_occurrence() {
        let source: DocumentBody = serde_json::from_value(json!({"content": [
            {"type": "paragraph", "content": []},
            {"type": "paragraph", "content": [], "sectionProperties": {"marginLeft": 1440}},
            {"type": "paragraph", "content": []},
        ]}))
        .unwrap();
        let mut output = source.clone();
        let BlockContent::Paragraph(paragraph) = &mut output.content[1] else {
            panic!("expected a section paragraph");
        };
        let section = Arc::make_mut(paragraph).section_properties.as_mut().unwrap();
        section.margin_left = Some(1800.0);
        let authored = properties(section);
        let repeated = output.content[1].clone();
        output.content.insert(2, repeated);

        let key = "boundary".to_owned();
        let sections = [SectionInfo {
            closing_paragraph_id: key.clone(),
            first_paragraph: 1,
            last_paragraph: 2,
            authored: authored.clone(),
            resolved: authored,
        }];
        let source_part = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:p><w:pPr><w:sectPr><w:pgMar w:left="1440"/></w:sectPr></w:pPr></w:p><w:p/></w:body></w:document>"#;
        let plan = section_save_plan(
            SectionPlanInput {
                source: Some(&source),
                output: &output,
                output_session_keys: &[None, Some(key.clone()), Some(key.clone()), None],
                metadata_unchanged: false,
                source_part: Some(source_part),
            },
            &sections,
            &ParagraphIdentities {
                session_id: "session".to_owned(),
                package_sha256: None,
                paragraphs: vec![ParagraphIdentity {
                    paragraph: ParagraphRef::Session {
                        story: "body".to_owned(),
                        para_id: key,
                    },
                    origin: ParagraphOrigin::Source,
                    ooxml_para_id: None,
                    id_origin: None,
                    source_story: None,
                    source: Some(SourceParagraphRef {
                        package_sha256: "0".repeat(64),
                        part_uri: "/word/document.xml".to_owned(),
                        paragraph_ordinal: 1,
                    }),
                }],
            },
            &ParagraphSavePlan::default(),
        )
        .unwrap()
        .unwrap();
        let edits = plan.plan.unwrap();
        assert!(!edits.only_sections);
        let [first, second] = edits.edits.as_slice() else {
            panic!("expected two section edits");
        };
        assert_eq!((first.paragraph, second.paragraph), (Some(1), Some(2)));
        for edit in [first, second] {
            assert_eq!(edit.body_paragraphs, Some(4));
            assert_eq!(edit.donor, Some(1));
            assert!(json_equal(&edit.properties, &json!({"marginLeft": 1800})));
        }
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

