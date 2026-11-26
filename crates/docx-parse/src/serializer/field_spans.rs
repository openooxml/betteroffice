use std::collections::HashMap;

use crate::block::BlockContent;
use crate::inline::{ComplexField, InlineNode, RunContent, StructuredFieldContent};
use crate::paragraph::ParagraphContent;

#[derive(Clone, Debug, Default)]
pub(super) struct FieldSpans(HashMap<usize, bool>);

impl FieldSpans {
    pub(super) fn in_story(blocks: &[BlockContent]) -> Self {
        let mut spans = Self::default();
        let mut open = Vec::new();
        for block in blocks {
            let BlockContent::Paragraph(paragraph) = block else {
                continue;
            };
            for content in &paragraph.content {
                match content {
                    ParagraphContent::Inline(node) => spans.visit(node, &mut open),
                    ParagraphContent::Tracked(change) => {
                        for node in &change.content {
                            spans.visit(node, &mut open);
                        }
                    }
                    _ => {}
                }
            }
        }
        spans
    }

    pub(super) fn has_separate_later(&self, field: &ComplexField) -> Option<bool> {
        self.0.get(&field_key(field)).copied()
    }

    pub(super) fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    fn visit(&mut self, node: &InlineNode, open: &mut Vec<Option<(usize, bool)>>) {
        let children = match node {
            InlineNode::ComplexField(field) => {
                if [&field.structured_code, &field.structured_result]
                    .into_iter()
                    .flatten()
                    .any(StructuredFieldContent::spans_paragraphs)
                {
                    open.push(Some((field_key(field), false)));
                }
                return;
            }
            _ if open.is_empty() => return,
            InlineNode::Run(run) => {
                for item in &run.content {
                    let RunContent::FieldChar { char_type, .. } = item else {
                        continue;
                    };
                    match char_type.as_str() {
                        "begin" => open.push(None),
                        "separate" => {
                            if let Some(Some((_, separate))) = open.last_mut() {
                                *separate = true;
                            }
                        }
                        "end" => {
                            if let Some(Some((key, separate))) = open.pop() {
                                self.0.insert(key, separate);
                            }
                        }
                        _ => {}
                    }
                }
                return;
            }
            InlineNode::Hyperlink(link) => &link.children,
            InlineNode::InlineSdt(sdt) => &sdt.content,
            InlineNode::Tracked(change) => &change.content,
            _ => return,
        };
        for child in children {
            self.visit(child, open);
        }
    }
}

fn field_key(field: &ComplexField) -> usize {
    std::ptr::from_ref(field) as usize
}
