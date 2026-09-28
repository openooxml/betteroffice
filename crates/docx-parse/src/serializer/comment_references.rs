//! One comment reference per paragraph that ends a comment's range, in what a save writes.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use crate::block::BlockContent;
use crate::inline::{InlineNode, Run, RunContent};
use crate::paragraph::{Paragraph, ParagraphContent};

/// Where a paragraph sits in the stories given to [`normalize_comment_references`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ParagraphLocation {
    /// The story's index in the slice.
    pub story: usize,
    /// The story's top-level block holding the paragraph.
    pub block: usize,
    /// Whether the paragraph is nested inside that block rather than being it.
    pub nested: bool,
    pub para_id: Option<String>,
}

/// Normalizes the comment references of the paragraphs a save writes from `stories`, those
/// `serialized` holds for; the others keep their source bytes and are left alone. For every
/// comment with a range end, a written paragraph holding one keeps its first reference of the
/// comment, wherever in the paragraph it sits, and loses the others. A written paragraph without
/// a range end loses the comment's references when the saved output holds one anyway: after a
/// written range end, where the serializer writes it, or in a paragraph kept as source bytes.
/// Otherwise the first such reference stays. References of a comment without a range end stay.
pub(crate) fn normalize_comment_references(
    stories: &mut [&mut Vec<BlockContent>],
    serialized: impl Fn(&ParagraphLocation) -> bool,
) {
    let mut paragraphs = Vec::new();
    for (story, blocks) in stories.iter().enumerate() {
        for (block, content) in blocks.iter().enumerate() {
            collect_block(content, story, block, false, &mut paragraphs);
        }
    }
    let written: Vec<bool> = paragraphs
        .iter()
        .map(|paragraph| serialized(&paragraph.location))
        .collect();
    let mut ends: HashMap<u64, Vec<usize>> = HashMap::new();
    for (index, paragraph) in paragraphs.iter().enumerate() {
        for &id in &paragraph.ends {
            ends.entry(id).or_default().push(index);
        }
    }
    let mut held: HashSet<u64> = HashSet::new();
    for (paragraph, &written) in paragraphs.iter().zip(&written) {
        held.extend(if written {
            &paragraph.ends
        } else {
            &paragraph.references
        });
    }
    let mut stray_kept = HashSet::new();
    let mut drops: Vec<Vec<bool>> = Vec::with_capacity(paragraphs.len());
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let mut kept = HashSet::new();
        let flags = paragraph
            .references
            .iter()
            .map(|id| {
                let Some(indices) = ends.get(id).filter(|_| written[index]) else {
                    return false;
                };
                if indices.contains(&index) {
                    !kept.insert(*id)
                } else {
                    held.contains(id) || !stray_kept.insert(*id)
                }
            })
            .collect();
        drops.push(flags);
    }
    let mut cursor = 0;
    for blocks in stories.iter_mut() {
        for content in blocks.iter_mut() {
            remove_in_block(content, &drops, &mut cursor);
        }
    }
}

struct CollectedParagraph {
    location: ParagraphLocation,
    ends: Vec<u64>,
    references: Vec<u64>,
}

fn comment_key(id: f64) -> u64 {
    (id + 0.0).to_bits()
}

fn collect_block(
    block: &BlockContent,
    story: usize,
    index: usize,
    nested: bool,
    out: &mut Vec<CollectedParagraph>,
) {
    match block {
        BlockContent::Paragraph(paragraph) => {
            let mut references = Vec::new();
            each_reference(paragraph, &mut |id| references.push(comment_key(id)));
            out.push(CollectedParagraph {
                location: ParagraphLocation {
                    story,
                    block: index,
                    nested,
                    para_id: paragraph.para_id.clone(),
                },
                ends: paragraph
                    .content
                    .iter()
                    .filter_map(|content| match content {
                        ParagraphContent::CommentRange(marker)
                            if marker.node_type == "commentRangeEnd" =>
                        {
                            Some(comment_key(marker.id))
                        }
                        _ => None,
                    })
                    .collect(),
                references,
            });
        }
        BlockContent::Table(table) => {
            for cell in table.rows.iter().flat_map(|row| &row.cells) {
                for child in &cell.content {
                    collect_block(child, story, index, true, out);
                }
            }
        }
        BlockContent::BlockSdt(sdt) => {
            for child in &sdt.content {
                collect_block(child, story, index, true, out);
            }
        }
        BlockContent::RawXml(_) => {}
    }
}

fn paragraph_count(block: &BlockContent) -> usize {
    match block {
        BlockContent::Paragraph(_) => 1,
        BlockContent::Table(table) => table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .flat_map(|cell| &cell.content)
            .map(paragraph_count)
            .sum(),
        BlockContent::BlockSdt(sdt) => sdt.content.iter().map(paragraph_count).sum(),
        BlockContent::RawXml(_) => 0,
    }
}

fn remove_in_block(block: &mut BlockContent, drops: &[Vec<bool>], cursor: &mut usize) {
    let count = paragraph_count(block);
    let range = *cursor..*cursor + count;
    *cursor += count;
    if !drops[range.clone()]
        .iter()
        .any(|flags| flags.contains(&true))
    {
        return;
    }
    let mut cursor = range.start;
    match block {
        BlockContent::Paragraph(paragraph) => {
            let mut remover = Remover {
                drops: &drops[cursor],
                next: 0,
            };
            remover.paragraph(Arc::make_mut(paragraph));
        }
        BlockContent::Table(table) => {
            let table = Arc::make_mut(table);
            for cell in table.rows.iter_mut().flat_map(|row| &mut row.cells) {
                for child in &mut cell.content {
                    remove_in_block(child, drops, &mut cursor);
                }
            }
        }
        BlockContent::BlockSdt(sdt) => {
            for child in &mut Arc::make_mut(sdt).content {
                remove_in_block(child, drops, &mut cursor);
            }
        }
        BlockContent::RawXml(_) => {}
    }
}

/// Calls `found` for each comment reference the paragraph writes, in the order it writes them.
fn each_reference(paragraph: &Paragraph, found: &mut impl FnMut(f64)) {
    fn run(value: &Run, found: &mut impl FnMut(f64)) {
        for content in &value.content {
            if let RunContent::CommentReference { id: Some(id) } = content {
                found(*id);
            }
        }
    }
    fn node(value: &InlineNode, found: &mut impl FnMut(f64)) {
        match value {
            InlineNode::Run(child) => run(child, found),
            InlineNode::Hyperlink(link) => {
                link.children.iter().for_each(|child| node(child, found))
            }
            InlineNode::SimpleField(field) => {
                field.content.iter().for_each(|child| run(child, found))
            }
            InlineNode::ComplexField(field) => {
                field.field_code.iter().for_each(|child| run(child, found));
                match written_result(field) {
                    Some(nodes) => nodes.iter().for_each(|child| node(child, found)),
                    None => field
                        .field_result
                        .iter()
                        .for_each(|child| run(child, found)),
                }
            }
            InlineNode::InlineSdt(sdt) => sdt.content.iter().for_each(|child| node(child, found)),
            _ => {}
        }
    }
    for content in &paragraph.content {
        match content {
            ParagraphContent::Inline(child) => node(child, found),
            ParagraphContent::Tracked(change) => {
                change.content.iter().for_each(|child| node(child, found))
            }
            _ => {}
        }
    }
}

/// The inline nodes the serializer writes as a complex field's result, when not its runs.
fn written_result(field: &crate::inline::ComplexField) -> Option<&Vec<InlineNode>> {
    field
        .structured_result
        .as_ref()
        .filter(|content| content.blocks.is_none())
        .and_then(|content| content.inline.as_ref())
}

/// Removes the references `drops` flags, visiting them in [`each_reference`] order, and the runs
/// that held nothing else.
struct Remover<'a> {
    drops: &'a [bool],
    next: usize,
}

impl Remover<'_> {
    fn paragraph(&mut self, paragraph: &mut Paragraph) {
        paragraph.content.retain_mut(|content| match content {
            ParagraphContent::Inline(node) => !self.node(node),
            ParagraphContent::Tracked(change) => {
                self.nodes(&mut change.content);
                true
            }
            _ => true,
        });
    }

    fn nodes(&mut self, nodes: &mut Vec<InlineNode>) {
        nodes.retain_mut(|node| !self.node(node));
    }

    fn runs(&mut self, runs: &mut Vec<Run>) {
        runs.retain_mut(|run| !self.run(run));
    }

    /// Whether the node is now an emptied run.
    fn node(&mut self, node: &mut InlineNode) -> bool {
        match node {
            InlineNode::Run(run) => return self.run(run),
            InlineNode::Hyperlink(link) => self.nodes(&mut link.children),
            InlineNode::SimpleField(field) => self.runs(&mut field.content),
            InlineNode::ComplexField(field) => {
                self.runs(&mut field.field_code);
                let written = field
                    .structured_result
                    .as_mut()
                    .filter(|content| content.blocks.is_none())
                    .and_then(|content| content.inline.as_mut());
                match written {
                    Some(nodes) => self.nodes(nodes),
                    None => self.runs(&mut field.field_result),
                }
            }
            InlineNode::InlineSdt(sdt) => self.nodes(&mut sdt.content),
            _ => {}
        }
        false
    }

    /// Whether the run held only references this removed.
    fn run(&mut self, run: &mut Run) -> bool {
        let before = run.content.len();
        run.content.retain(|content| {
            if !matches!(content, RunContent::CommentReference { id: Some(_) }) {
                return true;
            }
            let drop = self.drops[self.next];
            self.next += 1;
            !drop
        });
        run.content.is_empty() && run.content.len() < before
    }
}
