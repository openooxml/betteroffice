use std::collections::BTreeMap;

use crate::block::BlockContent;
use crate::inline::{FieldSpanEnd, InlineNode, RunContent};
use crate::paragraph::{Paragraph, ParagraphContent};

fn is_end(content: &ParagraphContent, paragraph: u32) -> bool {
    let node = match content {
        ParagraphContent::Inline(node) => node,
        ParagraphContent::Tracked(change) if change.content.len() == 1 => &change.content[0],
        _ => return false,
    };
    matches!(node, InlineNode::Run(run) if matches!(
        run.content.as_slice(),
        [RunContent::FieldChar { char_type, span_end: Some(id), .. }]
            if char_type == "end" && *id == paragraph
    ))
}

fn paragraph(block: &BlockContent) -> Option<&Paragraph> {
    match block {
        BlockContent::Paragraph(paragraph) => Some(paragraph),
        _ => None,
    }
}

pub(super) fn has_end(
    id: u32,
    remaining: &[ParagraphContent],
    following: &[BlockContent],
) -> bool {
    remaining.iter().any(|content| is_end(content, id))
        || following
            .iter()
            .filter_map(paragraph)
            .any(|p| p.content.iter().any(|content| is_end(content, id)))
}

pub(super) fn has_source_paragraph(id: u32, following: &[BlockContent]) -> bool {
    following
        .iter()
        .filter_map(paragraph)
        .any(|p| p.source_ordinal == Some(id))
}

#[derive(Default)]
pub(super) struct PendingFieldEnds {
    pub(super) block: usize,
    markers: BTreeMap<usize, Vec<ParagraphContent>>,
}

impl PendingFieldEnds {
    pub(super) fn carry(&mut self, end: &FieldSpanEnd, following: &[BlockContent]) -> bool {
        if !is_end(&end.marker, end.paragraph) {
            return false;
        }
        let target = following
            .iter()
            .enumerate()
            .take_while(|(_, block)| {
                paragraph(block)
                    .and_then(|p| p.source_ordinal)
                    .is_none_or(|ordinal| ordinal < end.paragraph)
            })
            .filter_map(|(index, block)| paragraph(block).map(|_| self.block + index + 1))
            .last()
            .unwrap_or(self.block);
        self.markers
            .entry(target)
            .or_default()
            .push((*end.marker).clone());
        true
    }

    pub(super) fn take(&mut self) -> Vec<ParagraphContent> {
        self.markers.remove(&self.block).unwrap_or_default()
    }
}
