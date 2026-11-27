pub(super) use transport::snapshot::plan_json;

mod transport {
    use std::sync::Arc;

    use docx_parse::{BlockContent, DocumentBody, SectionProperties};
    use serde::Deserialize;

    #[derive(Deserialize)]
    #[serde(rename_all = "camelCase")]
    struct BodyFragments {
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
        fn rebuild(self) -> Result<DocumentBody, String> {
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

    pub(super) mod snapshot {
        use docx_parse::serializer::{
            S13ParagraphId, S13PatchedPart, S13SpliceAnchor, S13TablePatch,
        };
        use serde::Deserialize;
        use serde_json::Value;

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
            sections: Vec<SectionInfoInput>,
            identities: Identities,
            plan: SavePlan,
        }

        #[derive(Deserialize)]
        #[serde(rename = "SectionInfo", rename_all = "camelCase")]
        struct SectionInfoInput {
            closing_paragraph_id: String,
            first_paragraph: usize,
            last_paragraph: usize,
            authored: Value,
            resolved: Value,
        }

        impl From<SectionInfoInput> for SectionInfo {
            fn from(section: SectionInfoInput) -> Self {
                Self {
                    closing_paragraph_id: section.closing_paragraph_id,
                    first_paragraph: section.first_paragraph,
                    last_paragraph: section.last_paragraph,
                    authored: section.authored,
                    resolved: section.resolved,
                }
            }
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

        pub(in crate::wasm) fn plan_json(json: &str) -> Result<String, String> {
            let input: Input = serde_json::from_str(json).map_err(|error| error.to_string())?;
            let source = input.source.map(BodyFragments::rebuild).transpose()?;
            let output = input.output.rebuild()?;
            let identities = ParagraphIdentities::try_from(input.identities)?;
            let plan = input.plan.into();
            let sections: Vec<_> = input.sections.into_iter().map(SectionInfo::from).collect();
            let result = section_save_plan(
                SectionPlanInput {
                    source: source.as_ref(),
                    output: &output,
                    output_session_keys: &input.output_session_keys,
                    metadata_unchanged: input.metadata_unchanged,
                    source_part: input.source_part.as_deref(),
                },
                &sections,
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
