use std::rc::Rc;

use docx_layout::regions::DocumentRegions;
use docx_layout::types::{LayoutBlock, ParagraphBlock};
use serde::{Serialize, Serializer};

#[derive(Clone, Debug, Default)]
pub(super) struct SharedBlocks {
    blocks: Vec<Rc<LayoutBlock>>,
}

impl From<Vec<LayoutBlock>> for SharedBlocks {
    fn from(blocks: Vec<LayoutBlock>) -> Self {
        Self {
            blocks: blocks.into_iter().map(Rc::new).collect(),
        }
    }
}

impl SharedBlocks {
    pub(super) fn iter(&self) -> impl Iterator<Item = &LayoutBlock> {
        self.blocks.iter().map(Rc::as_ref)
    }

    pub(super) fn shared(&self) -> &[Rc<LayoutBlock>] {
        &self.blocks
    }

    pub(super) fn shared_mut(&mut self) -> &mut Vec<Rc<LayoutBlock>> {
        &mut self.blocks
    }

    pub(super) fn to_vec(&self) -> Vec<LayoutBlock> {
        self.iter().cloned().collect()
    }
}

impl Serialize for SharedBlocks {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.collect_seq(self.iter())
    }
}

pub(super) trait BlockSource {
    fn len(&self) -> usize;
    fn get(&self, index: usize) -> Option<&LayoutBlock>;
    fn identity(&self, index: usize) -> Option<&Rc<LayoutBlock>>;
}

impl BlockSource for [LayoutBlock] {
    fn len(&self) -> usize {
        <[LayoutBlock]>::len(self)
    }

    fn get(&self, index: usize) -> Option<&LayoutBlock> {
        <[LayoutBlock]>::get(self, index)
    }

    fn identity(&self, _: usize) -> Option<&Rc<LayoutBlock>> {
        None
    }
}

impl BlockSource for Vec<LayoutBlock> {
    fn len(&self) -> usize {
        self.as_slice().len()
    }

    fn get(&self, index: usize) -> Option<&LayoutBlock> {
        self.as_slice().get(index)
    }

    fn identity(&self, _: usize) -> Option<&Rc<LayoutBlock>> {
        None
    }
}

impl BlockSource for SharedBlocks {
    fn len(&self) -> usize {
        self.blocks.len()
    }

    fn get(&self, index: usize) -> Option<&LayoutBlock> {
        self.blocks.get(index).map(Rc::as_ref)
    }

    fn identity(&self, index: usize) -> Option<&Rc<LayoutBlock>> {
        self.blocks.get(index)
    }
}

#[derive(Clone, Debug)]
pub(super) struct BlockCertificate {
    source: Rc<LayoutBlock>,
    previous: Option<Rc<LayoutBlock>>,
    next: Option<Rc<LayoutBlock>>,
    context: Rc<Vec<u8>>,
    section: usize,
    pub(super) reflexive: bool,
    pub(super) raw: bool,
}

fn same_identity(left: Option<&Rc<LayoutBlock>>, right: Option<&Rc<LayoutBlock>>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => Rc::ptr_eq(left, right),
        (None, None) => true,
        _ => false,
    }
}

impl BlockCertificate {
    pub(super) fn new(
        blocks: &SharedBlocks,
        index: usize,
        context: Rc<Vec<u8>>,
        section: usize,
        raw: bool,
    ) -> Self {
        let source = Rc::clone(blocks.identity(index).expect("block identity"));
        let reflexive = source.as_ref().eq(source.as_ref());
        Self {
            source,
            previous: index
                .checked_sub(1)
                .and_then(|i| blocks.identity(i))
                .cloned(),
            next: blocks.identity(index + 1).cloned(),
            context,
            section,
            reflexive,
            raw,
        }
    }

    pub(super) fn matches_source(
        &self,
        blocks: &(impl BlockSource + ?Sized),
        index: usize,
    ) -> bool {
        same_identity(Some(&self.source), blocks.identity(index))
    }

    pub(super) fn matches_neighbors(
        &self,
        blocks: &(impl BlockSource + ?Sized),
        index: usize,
    ) -> bool {
        self.matches_source(blocks, index)
            && same_identity(
                self.previous.as_ref(),
                index.checked_sub(1).and_then(|i| blocks.identity(i)),
            )
            && same_identity(self.next.as_ref(), blocks.identity(index + 1))
    }

    pub(super) fn matches(
        &self,
        blocks: &(impl BlockSource + ?Sized),
        index: usize,
        context: &Rc<Vec<u8>>,
    ) -> bool {
        self.matches_neighbors(blocks, index) && Rc::ptr_eq(&self.context, context)
    }
}

pub(super) fn normalization_contexts(
    regions: &DocumentRegions,
    measurement: &docx_layout::measure_blocks::MeasurementConfig,
    retained: &[Option<BlockCertificate>],
) -> Result<Vec<Rc<Vec<u8>>>, String> {
    let (store, fonts) = docx_layout::measure_fonts_generation();
    let mut missing: Vec<_> = measurement
        .font_chains
        .values()
        .flatten()
        .copied()
        .filter(|&id| id as usize >= fonts)
        .collect();
    missing.sort_unstable();
    missing.dedup();
    let mut contexts = (0..regions.sections.len() + 1)
        .map(|index| {
            let section = regions.sections.get(index);
            serde_json::to_vec(&(
                regions.paragraph_spacing_line_px(index),
                regions.doc_grid_snap_pitch_px(index),
                section.and_then(|section| section.page_size.as_ref()),
                section.and_then(|section| section.margins.as_ref()),
                section.and_then(|section| section.columns.as_ref()),
                &measurement.defaults,
                &measurement.compat,
                measurement.authoritative_shaping,
                store,
                fonts == 0,
                &missing,
            ))
            .map(Rc::new)
            .map_err(|error| format!("retain normalization context: {error}"))
        })
        .collect::<Result<Vec<_>, _>>()?;
    let mut checked = vec![false; contexts.len()];
    for certificate in retained.iter().flatten() {
        if let Some(context) = contexts.get_mut(certificate.section)
            && !checked[certificate.section]
        {
            if context.as_ref() == certificate.context.as_ref() {
                *context = Rc::clone(&certificate.context);
            }
            checked[certificate.section] = true;
        }
    }
    Ok(contexts)
}

/// Only identity and start position survive until pagination consumes the old arena.
pub(super) fn take_block(block: &mut LayoutBlock) -> LayoutBlock {
    let metadata = super::fragment_identity(block).map_or(LayoutBlock::Unsupported, |id| {
        LayoutBlock::Paragraph(ParagraphBlock {
            sdt_groups: None,
            id: id.clone(),
            para_id: None,
            runs: Vec::new(),
            attrs: None,
            pm_start: block.pm_start(),
            pm_end: None,
        })
    });
    std::mem::replace(block, metadata)
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_layout::types::Run;
    use serde_json::json;

    fn blocks() -> SharedBlocks {
        serde_json::from_value::<Vec<LayoutBlock>>(json!([
            {"kind": "paragraph", "id": "first", "runs": [{"kind": "text", "text": "First"}], "pmStart": 1, "pmEnd": 8},
            {"kind": "table", "id": "table", "rows": [{"id": "row", "cells": [{"id": "cell", "blocks": [
                {"kind": "paragraph", "id": "nested", "runs": [{"kind": "text", "text": "Cell"}]}
            ]}]}]},
            {"kind": "paragraph", "id": "tail", "runs": [{"kind": "text", "text": "Tail"}]}
        ])).unwrap().into()
    }

    #[test]
    fn shared_splices_preserve_immutable_block_identity_and_snapshots() {
        let original = blocks();
        let before = serde_json::to_vec(&original).unwrap();
        let mut patched = original.clone();
        let replacement = serde_json::from_value(json!({
            "kind": "paragraph", "id": "first", "runs": [{"kind": "text", "text": "Changed"}]
        }))
        .unwrap();
        patched.shared_mut().splice(0..1, [Rc::new(replacement)]);
        assert!(!Rc::ptr_eq(&original.shared()[0], &patched.shared()[0]));
        assert!(Rc::ptr_eq(&original.shared()[1], &patched.shared()[1]));
        assert!(Rc::ptr_eq(&original.shared()[2], &patched.shared()[2]));
        assert_eq!(serde_json::to_vec(&original).unwrap(), before);
        assert!(
            original
                .iter()
                .zip(original.shared())
                .all(|(block, shared)| std::ptr::eq(block, shared.as_ref()))
        );
        assert!(
            patched
                .iter()
                .zip(patched.shared())
                .all(|(block, shared)| std::ptr::eq(block, shared.as_ref()))
        );
    }

    #[test]
    fn reusing_a_certificate_after_mutation_falls_back() {
        for nested in [false, true] {
            let original = blocks();
            let index = usize::from(nested);
            let context = Rc::new(Vec::new());
            let token = BlockCertificate::new(&original, index, Rc::clone(&context), 0, true);
            let mut changed = original.clone();
            let block = Rc::make_mut(&mut changed.shared_mut()[index]);
            let paragraph = match block {
                LayoutBlock::Paragraph(paragraph) => paragraph,
                LayoutBlock::Table(table) => {
                    let LayoutBlock::Paragraph(paragraph) = &mut table.rows[0].cells[0].blocks[0]
                    else {
                        unreachable!();
                    };
                    paragraph
                }
                _ => unreachable!(),
            };
            let Run::Text(run) = &mut paragraph.runs[0] else {
                unreachable!()
            };
            run.text.push('!');
            assert!(!token.matches(&changed, index, &context));
            assert_ne!(original.get(index), changed.get(index));
        }
    }

    #[test]
    fn position_only_shifts_keep_deep_equality_and_invalidate_the_identity() {
        let original = blocks();
        let context = Rc::new(Vec::new());
        let token = BlockCertificate::new(&original, 0, Rc::clone(&context), 0, true);
        let mut shifted = original.clone();
        let LayoutBlock::Paragraph(paragraph) = Rc::make_mut(&mut shifted.shared_mut()[0]) else {
            unreachable!();
        };
        paragraph.pm_start = Some(11.0);
        paragraph.pm_end = Some(18.0);
        assert_eq!(original.get(0), shifted.get(0));
        assert!(!token.matches(&shifted, 0, &context));
    }

    #[test]
    fn normalization_and_neighbor_changes_invalidate_certificates() {
        let original = blocks();
        let context = Rc::new(vec![1]);
        let token = BlockCertificate::new(&original, 0, Rc::clone(&context), 0, true);
        assert!(token.matches(&original, 0, &context));
        let mut changed = original.clone();
        changed.shared_mut()[1] = Rc::new(LayoutBlock::Unsupported);
        assert!(token.matches_source(&changed, 0));
        assert!(!token.matches(&changed, 0, &context));
        assert!(!token.matches(&original, 0, &Rc::new(vec![2])));
    }
}
