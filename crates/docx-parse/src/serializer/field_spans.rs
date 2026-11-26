use std::collections::{BTreeSet, HashMap, HashSet};

use crate::block::{BlockContent, story_block_elements};
use crate::inline::{ComplexField, InlineNode, RunContent, parse_run};
use crate::paragraph::{Paragraph, ParagraphContent, TrackedInline, parse_tracked_change_info};
use crate::paragraph_identity::parse_paragraph_id;
use crate::xml::{ParseBudget, ParseLimits, XmlElement, parse_xml};

#[derive(Clone, Debug, Default)]
pub(super) struct FieldSpans {
    pub(super) paragraphs: HashMap<usize, FieldParagraph>,
    pub(super) carried: BTreeSet<u32>,
    pub(super) paired: BTreeSet<u32>,
    pub(super) moved: BTreeSet<u32>,
}

#[derive(Clone, Debug, Default)]
pub(super) struct FieldParagraph {
    pub(super) opening: Option<usize>,
    pub(super) closing: Option<Vec<ParagraphContent>>,
}

pub(super) fn paragraph_key(paragraph: &Paragraph) -> usize {
    std::ptr::from_ref(paragraph) as usize
}

pub(super) fn spans_paragraphs(field: &ComplexField) -> bool {
    [
        field.structured_code.as_ref(),
        field.structured_result.as_ref(),
    ]
    .into_iter()
    .flatten()
    .any(|part| part.spans_paragraphs())
}

fn story_has_span(model: &[BlockContent]) -> bool {
    model.iter().any(|block| match block {
        BlockContent::Paragraph(paragraph) => paragraph.content.iter().any(|child| {
            matches!(child, ParagraphContent::Inline(InlineNode::ComplexField(field))
                if spans_paragraphs(field))
        }),
        BlockContent::BlockSdt(sdt) => story_has_span(&sdt.content),
        _ => false,
    })
}

struct Opening {
    block: usize,
    instruction: String,
    direct: bool,
    separated: bool,
}

struct Closing<'a> {
    run: &'a XmlElement,
    revision: Option<&'a XmlElement>,
    revision_offset: usize,
    offset: usize,
}

impl Closing<'_> {
    fn separate_run(&self) -> bool {
        self.run
            .child_elements()
            .all(|child| matches!(child.name.as_str(), "w:rPr" | "w:fldChar"))
    }

    fn separate_revision(&self) -> bool {
        self.revision
            .is_none_or(|revision| revision.child_elements().count() == 1)
    }

    fn standalone(&self) -> bool {
        self.separate_run() && self.separate_revision()
    }

    fn alone_in(&self, paragraph: &XmlElement) -> bool {
        let mut children = paragraph
            .child_elements()
            .filter(|child| child.name != "w:pPr");
        self.standalone()
            && children
                .next()
                .is_some_and(|child| std::ptr::eq(child, self.revision.unwrap_or(self.run)))
            && children.next().is_none()
    }

    fn marker(&self) -> ParagraphContent {
        let mut run = parse_run(self.run, None, None, None).run;
        run.content.retain(is_end);
        let run = InlineNode::Run(run);
        match self.revision {
            None => ParagraphContent::Inline(run),
            Some(revision) => ParagraphContent::Tracked(TrackedInline {
                node_type: match revision.name.as_str() {
                    "w:ins" => "insertion",
                    "w:del" => "deletion",
                    "w:moveFrom" => "moveFrom",
                    _ => "moveTo",
                }
                .to_owned(),
                info: parse_tracked_change_info(revision),
                content: vec![run],
            }),
        }
    }
}

fn is_end(content: &RunContent) -> bool {
    matches!(content, RunContent::FieldChar { char_type, .. } if char_type == "end")
}

fn has_end(paragraph: &Paragraph) -> bool {
    fn run_has_end(node: &InlineNode) -> bool {
        matches!(node, InlineNode::Run(run) if run.content.iter().any(is_end))
    }
    paragraph.content.iter().any(|child| match child {
        ParagraphContent::Inline(node) => run_has_end(node),
        ParagraphContent::Tracked(change) => change.content.iter().any(run_has_end),
        _ => false,
    })
}

#[derive(Default)]
struct SourceParagraph<'a> {
    text: String,
    openings: Vec<Opening>,
    closings: Vec<Option<Closing<'a>>>,
}

impl<'a> SourceParagraph<'a> {
    fn read(block: usize, paragraph: &'a XmlElement) -> Self {
        let mut source = Self::default();
        for child in paragraph.child_elements() {
            match child.name.as_str() {
                "w:r" => source.run(block, child, None, 0),
                "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => {
                    let offset = source.text.len();
                    for item in child.child_elements() {
                        if item.name == "w:r" {
                            source.run(block, item, Some(child), offset);
                        } else {
                            source.nested(item);
                        }
                    }
                }
                _ => source.nested(child),
            }
        }
        source
    }

    fn run(
        &mut self,
        block: usize,
        run: &'a XmlElement,
        revision: Option<&'a XmlElement>,
        revision_offset: usize,
    ) {
        let own_run = run
            .child_elements()
            .filter(|child| child.name != "w:rPr")
            .nth(1)
            .is_none();
        if !own_run
            && run
                .child_elements()
                .filter(|item| item.name == "w:fldChar")
                .count()
                > 1
        {
            self.openings.clear();
            self.closings.push(None);
            return;
        }
        for item in run.child_elements() {
            match item.name.as_str() {
                "w:fldChar" => match item.attribute(Some("w"), "fldCharType") {
                    Some("begin") => self.openings.push(Opening {
                        block,
                        instruction: String::new(),
                        direct: revision.is_none() && own_run,
                        separated: false,
                    }),
                    Some("separate") => {
                        if let Some(opening) = self.openings.last_mut() {
                            opening.separated = revision.is_none() && own_run;
                        } else {
                            self.closings.push(None);
                        }
                    }
                    Some("end") if self.openings.pop().is_none() => {
                        self.closings.push(Some(Closing {
                            run,
                            revision,
                            revision_offset,
                            offset: self.text.len(),
                        }));
                    }
                    _ => {}
                },
                "w:instrText" => {
                    if let Some(opening) = self.openings.last_mut() {
                        opening.instruction.push_str(&item.text_content());
                    }
                }
                _ => self.nested(item),
            }
        }
    }

    fn nested(&mut self, node: &XmlElement) {
        match node.name.as_str() {
            "w:p" | "w:pPr" | "w:rPr" => {}
            "w:t" | "w:delText" => self.text.push_str(&node.text_content()),
            "w:tab" => self.text.push('\t'),
            _ => {
                for child in node.child_elements() {
                    self.nested(child);
                }
            }
        }
    }
}

fn append_text(node: &InlineNode, text: &mut String) {
    match node {
        InlineNode::Run(run) => {
            for item in &run.content {
                match item {
                    RunContent::Text { text: value, .. } => text.push_str(value),
                    RunContent::Tab => text.push('\t'),
                    _ => {}
                }
            }
        }
        InlineNode::Hyperlink(link) => {
            for child in link.structured_children.as_ref().unwrap_or(&link.children) {
                append_text(child, text);
            }
        }
        InlineNode::ComplexField(field) => {
            if let Some(nodes) = field
                .structured_result
                .as_ref()
                .and_then(|part| part.inline.as_ref())
            {
                for child in nodes {
                    append_text(child, text);
                }
            }
        }
        _ => {}
    }
}

fn append_content_text(content: &ParagraphContent, text: &mut String) {
    match content {
        ParagraphContent::Inline(node) => append_text(node, text),
        ParagraphContent::Tracked(change) => {
            for node in &change.content {
                append_text(node, text);
            }
        }
        _ => {}
    }
}

fn visible_text(paragraph: &Paragraph) -> String {
    let mut text = String::new();
    for child in &paragraph.content {
        append_content_text(child, &mut text);
    }
    text
}

fn unique<K: std::hash::Hash + Eq>(
    keys: impl Iterator<Item = Option<K>>,
) -> HashMap<K, Option<usize>> {
    let mut indices = HashMap::new();
    for (index, key) in keys
        .enumerate()
        .filter_map(|(index, key)| Some((index, key?)))
    {
        indices
            .entry(key)
            .and_modify(|entry| *entry = None)
            .or_insert(Some(index));
    }
    indices
}

fn project_paragraphs(
    source: &[(Vec<u32>, &XmlElement)],
    model: &[BlockContent],
    paragraphs: &[Option<SourceParagraph<'_>>],
    removed: &HashSet<u32>,
) -> Vec<Option<usize>> {
    let ordinals = unique(model.iter().map(|block| paragraph(block)?.source_ordinal));
    let ids = unique(
        model
            .iter()
            .map(|block| parse_paragraph_id(paragraph(block)?.para_id.as_deref()?)),
    );
    let source_ids = unique(source.iter().map(|(_, element)| source_id(element)));
    let controls = unique(model.iter().map(|block| match block {
        BlockContent::BlockSdt(sdt) => sdt.properties.id.map(f64::to_bits),
        _ => None,
    }));
    let source_controls = unique(source.iter().map(|(_, element)| control_id(element)));
    let mut matches: Vec<_> = source
        .iter()
        .enumerate()
        .map(|(index, (_, element))| {
            if let Some(id) = control_id(element) {
                return if source_controls.get(&id) == Some(&Some(index)) {
                    controls.get(&id).copied().flatten()
                } else {
                    None
                };
            }
            let ordinal = element.paragraph_ordinal?;
            if removed.contains(&ordinal) {
                return None;
            }
            ordinals.get(&ordinal).copied().flatten().or_else(|| {
                let id = source_id(element)?;
                if source_ids.get(&id) == Some(&Some(index)) {
                    ids.get(&id).copied().flatten()
                } else {
                    None
                }
            })
        })
        .collect();
    let mut used: HashSet<_> = matches.iter().flatten().copied().collect();
    if source.len() != model.len() {
        let source_text = unique(
            paragraphs
                .iter()
                .map(|p| p.as_ref().map(|p| &p.text).filter(|text| !text.is_empty())),
        );
        let texts: Vec<_> = model
            .iter()
            .map(|block| paragraph(block).map(visible_text))
            .collect();
        let live_text = unique(texts.iter().map(Option::as_ref));
        for (text, index) in source_text {
            let (Some(index), Some(live)) = (index, live_text.get(text).copied().flatten()) else {
                continue;
            };
            if matches[index].is_none()
                && !source[index]
                    .1
                    .paragraph_ordinal
                    .is_some_and(|ordinal| removed.contains(&ordinal))
                && used.insert(live)
            {
                matches[index] = Some(live);
            }
        }
    }
    let is_removed = |index: usize| {
        source[index]
            .1
            .paragraph_ordinal
            .is_some_and(|ordinal| removed.contains(&ordinal))
    };
    let mut prior = (0, 0);
    for (next_source, next_model) in matches
        .clone()
        .into_iter()
        .enumerate()
        .filter_map(|(index, projected)| Some((index, projected?)))
        .chain(std::iter::once((source.len(), model.len())))
    {
        if next_source < prior.0 || next_model < prior.1 {
            return matches;
        }
        if next_source - prior.0 == next_model - prior.1 {
            for (index, live) in (prior.0..next_source).zip(prior.1..next_model) {
                if !is_removed(index)
                    && match &model[live] {
                        BlockContent::Paragraph(_) => paragraphs[index].is_some(),
                        BlockContent::BlockSdt(sdt) => {
                            source[index].1.matches_name("w", "sdt")
                                && control_id(source[index].1)
                                    == sdt.properties.id.map(f64::to_bits)
                        }
                        _ => false,
                    }
                    && used.insert(live)
                {
                    matches[index] = Some(live);
                }
            }
        } else if next_source - prior.0 == 1
            && !is_removed(prior.0)
            && paragraphs[prior.0].as_ref().is_some_and(|entry| {
                matches!(entry.closings.as_slice(), [Some(closing)]
                    if closing.alone_in(source[prior.0].1))
            })
        {
            let mut empty = (prior.1..next_model).filter(|&live| {
                paragraph(&model[live]).is_some_and(|paragraph| paragraph.content.is_empty())
            });
            if let Some(live) = empty.next()
                && empty.next().is_none()
                && used.insert(live)
            {
                matches[prior.0] = Some(live);
            }
        }
        prior = (next_source + 1, next_model + 1);
    }
    matches
}

fn source_id(element: &XmlElement) -> Option<u32> {
    parse_paragraph_id(
        element
            .attributes
            .get(element.para_id_attribute.as_ref()?)?,
    )
}

fn control_id(element: &XmlElement) -> Option<u64> {
    if !element.matches_name("w", "sdt") {
        return None;
    }
    let value = element
        .child("w", "sdtPr")?
        .child("w", "id")?
        .attribute(Some("w"), "val")?;
    crate::xml::parse_javascript_integer_prefix(value).map(f64::to_bits)
}

fn paragraph(block: &BlockContent) -> Option<&Paragraph> {
    match block {
        BlockContent::Paragraph(paragraph) => Some(paragraph),
        _ => None,
    }
}

fn anchored_offset(before: &str, after: &str, current: &str) -> Option<usize> {
    if before.is_empty() {
        return Some(0);
    }
    if after.is_empty() {
        return Some(current.len());
    }
    if before.len() + after.len() == current.len()
        && current.starts_with(before)
        && current.ends_with(after)
    {
        return Some(before.len());
    }
    let unique_offset = |needle: &str| {
        let mut matches = current.match_indices(needle);
        let (offset, _) = matches.next()?;
        matches.next().is_none().then_some(offset)
    };
    let left = before
        .char_indices()
        .rev()
        .take(64)
        .filter_map(|(index, _)| {
            unique_offset(&before[index..]).map(|at| at + before.len() - index)
        })
        .last();
    let right = after
        .char_indices()
        .skip(1)
        .take(63)
        .map(|(index, _)| index)
        .chain(std::iter::once(after.len()))
        .filter_map(|end| unique_offset(&after[..end]))
        .last();
    match (left, right) {
        (Some(left), Some(right)) if left <= right => Some(right),
        (Some(left), None) if current.match_indices(after).nth(1).is_none() => Some(left),
        (None, Some(right)) if current.match_indices(before).nth(1).is_none() => Some(right),
        _ => None,
    }
}

fn run_boundary(
    source: &XmlElement,
    closing: &Closing<'_>,
    paragraph: &Paragraph,
) -> Option<usize> {
    if !closing.standalone() || closing.revision.is_some() {
        return None;
    }
    let mut cut = None;
    let mut original = Vec::new();
    for child in source
        .child_elements()
        .filter(|child| child.name != "w:pPr")
    {
        if std::ptr::eq(child, closing.run) {
            cut = Some(original.len());
            continue;
        }
        if child.name != "w:r" {
            return None;
        }
        let mut text = String::new();
        for item in child.child_elements().filter(|item| item.name != "w:rPr") {
            if item.name != "w:t" {
                return None;
            }
            text.push_str(&item.text_content());
        }
        original.push(text);
    }
    let cut = cut?;
    if cut == 0 || cut == original.len() || original.len() != paragraph.content.len() {
        return None;
    }
    let live = paragraph
        .content
        .iter()
        .map(|child| {
            let ParagraphContent::Inline(InlineNode::Run(run)) = child else {
                return None;
            };
            run.content
                .iter()
                .map(|item| match item {
                    RunContent::Text { text, .. } => Some(text.as_str()),
                    _ => None,
                })
                .collect::<Option<String>>()
        })
        .collect::<Option<Vec<_>>>()?;
    (original[..cut] == live[..cut] || original[cut..] == live[cut..]).then_some(cut)
}

fn place(
    mut content: Vec<ParagraphContent>,
    mut offset: usize,
    marker: ParagraphContent,
    shared_run_end: Option<(RunContent, bool)>,
) -> Option<Vec<ParagraphContent>> {
    for index in 0..content.len() {
        if offset == 0 && shared_run_end.is_none() {
            content.insert(index, marker);
            return Some(content);
        }
        let child = &mut content[index];
        let ParagraphContent::Inline(InlineNode::Run(run)) = child else {
            let mut text = String::new();
            append_content_text(child, &mut text);
            offset = offset.checked_sub(text.len())?;
            continue;
        };
        for at in 0..run.content.len() {
            let width = match &run.content[at] {
                RunContent::Text { text, .. } => text.len(),
                RunContent::Tab => 1,
                _ => 0,
            };
            if let Some((end, prefer_previous)) = &shared_run_end
                && width > 0
                && (offset < width || (offset == width && *prefer_previous))
            {
                if offset == 0 || offset == width {
                    run.content
                        .insert(at + usize::from(offset != 0), end.clone());
                } else {
                    let RunContent::Text {
                        text,
                        preserve_space,
                    } = &mut run.content[at]
                    else {
                        return None;
                    };
                    if !text.is_char_boundary(offset) {
                        return None;
                    }
                    let tail = RunContent::Text {
                        text: text.split_off(offset),
                        preserve_space: *preserve_space,
                    };
                    run.content.splice(at + 1..at + 1, [end.clone(), tail]);
                }
                return Some(content);
            }
            if offset < width {
                let RunContent::Text {
                    text,
                    preserve_space,
                } = &mut run.content[at]
                else {
                    return None;
                };
                if !text.is_char_boundary(offset) {
                    return None;
                }
                let suffix = text.split_off(offset);
                let preserve_space = *preserve_space;
                let mut tail = run.clone();
                tail.content.drain(..at);
                tail.content[0] = RunContent::Text {
                    text: suffix,
                    preserve_space,
                };
                let prefix_len = at + usize::from(offset != 0);
                run.content.truncate(prefix_len);
                let tail = ParagraphContent::Inline(InlineNode::Run(tail));
                content.splice(
                    index + usize::from(prefix_len != 0)..index + 1,
                    [marker, tail],
                );
                return Some(content);
            }
            offset -= width;
        }
    }
    if shared_run_end.is_some() || offset != 0 {
        return None;
    }
    content.push(marker);
    Some(content)
}

fn place_closing(
    source: &XmlElement,
    paragraph: &Paragraph,
    closing: &Closing<'_>,
    text: &str,
    removed: bool,
) -> Option<Vec<ParagraphContent>> {
    let mut marker = closing.marker();
    if removed {
        let mut content = paragraph.content.clone();
        content.push(marker);
        return Some(content);
    }
    if let Some(boundary) = run_boundary(source, closing, paragraph) {
        let mut content = paragraph.content.clone();
        content.insert(boundary, marker);
        return Some(content);
    }
    let shared_run_end = if closing.separate_run() {
        None
    } else {
        let node = match &mut marker {
            ParagraphContent::Inline(node) => node,
            ParagraphContent::Tracked(revision) => revision.content.first_mut()?,
            _ => return None,
        };
        let InlineNode::Run(run) = node else {
            return None;
        };
        let prefer_previous = closing
            .run
            .child_elements()
            .take_while(|child| child.name != "w:fldChar")
            .any(|child| matches!(child.name.as_str(), "w:t" | "w:delText" | "w:tab"));
        Some((run.content.pop()?, prefer_previous))
    };
    if closing.revision.is_some() && !closing.standalone() {
        let ParagraphContent::Tracked(marker) = marker else {
            return None;
        };
        let mut content = paragraph.content.clone();
        let revision = content.iter_mut().find_map(|child| match child {
            ParagraphContent::Tracked(change)
                if change.node_type == marker.node_type && change.info == marker.info =>
            {
                Some(change)
            }
            _ => None,
        })?;
        let mut live = String::new();
        for node in &revision.content {
            append_text(node, &mut live);
        }
        let revision_source = SourceParagraph::read(0, closing.revision?);
        let (before, after) = revision_source
            .text
            .split_at_checked(closing.offset - closing.revision_offset)?;
        let offset = anchored_offset(before, after, &live)?;
        let marker = ParagraphContent::Inline(marker.content.into_iter().next()?);
        let nodes = std::mem::take(&mut revision.content);
        revision.content = place(
            nodes.into_iter().map(ParagraphContent::Inline).collect(),
            offset,
            marker,
            shared_run_end,
        )?
        .into_iter()
        .filter_map(|child| match child {
            ParagraphContent::Inline(node) => Some(node),
            _ => None,
        })
        .collect();
        return Some(content);
    }
    let live = visible_text(paragraph);
    let (before, after) = text.split_at(closing.offset);
    let offset = anchored_offset(before, after, &live)?;
    place(paragraph.content.clone(), offset, marker, shared_run_end)
}

impl FieldSpans {
    pub(super) fn from_document(
        source: &str,
        model: &[BlockContent],
        removed: &[u32],
        selected: impl Fn(&XmlElement, &Paragraph) -> bool,
    ) -> Self {
        if !source.contains("fldChar") || !story_has_span(model) {
            return Self::default();
        }
        let limits = ParseLimits::default();
        let Ok(xml) = parse_xml(
            source.as_bytes(),
            "word/document.xml",
            &mut ParseBudget::new(&limits),
        ) else {
            return Self::default();
        };
        let mut plan = Self::default();
        if let Some(body) = xml.root().and_then(|root| root.child("w", "body")) {
            plan.plan_story(
                &story_block_elements(body),
                model,
                &removed.iter().copied().collect(),
                &selected,
            );
        }
        plan
    }

    pub(super) fn from_blocks(
        source: &[(Vec<u32>, &XmlElement)],
        model: &[BlockContent],
        selected: impl Fn(&XmlElement, &Paragraph) -> bool,
    ) -> Self {
        let mut plan = Self::default();
        if story_has_span(model) {
            plan.plan_story(source, model, &HashSet::new(), &selected);
        }
        plan
    }

    fn plan_story(
        &mut self,
        source: &[(Vec<u32>, &XmlElement)],
        model: &[BlockContent],
        removed: &HashSet<u32>,
        selected: &impl Fn(&XmlElement, &Paragraph) -> bool,
    ) {
        let mut paragraphs: Vec<_> = source
            .iter()
            .enumerate()
            .map(|(index, (_, element))| {
                (element.name == "w:p").then(|| SourceParagraph::read(index, element))
            })
            .collect();
        let matches = project_paragraphs(source, model, &paragraphs, removed);
        let mut following = vec![None; matches.len()];
        let mut next = None;
        for (index, &matched) in matches.iter().enumerate().rev() {
            following[index] = next;
            next = matched.or(next);
        }
        let model_groups: Vec<_> = model
            .iter()
            .scan(0usize, |group, block| {
                *group += usize::from(paragraph(block).is_none());
                Some(*group)
            })
            .collect();
        let mut pending = None;
        for (end, entry) in paragraphs.iter_mut().enumerate() {
            let Some(entry) = entry else {
                pending = None;
                if let Some(index) = matches[end]
                    && let BlockContent::BlockSdt(sdt) = &model[index]
                    && story_has_span(&sdt.content)
                    && let Some(content) = source[end].1.child("w", "sdtContent")
                {
                    self.plan_story(
                        &story_block_elements(content),
                        &sdt.content,
                        removed,
                        selected,
                    );
                }
                continue;
            };
            if !entry.closings.is_empty() {
                if let (Some(opening), [Some(closing)]) =
                    (pending.take(), entry.closings.as_slice())
                {
                    let deleted = source[end]
                        .1
                        .paragraph_ordinal
                        .is_some_and(|id| removed.contains(&id));
                    let next = following[end]
                        .or_else(|| (end + 1 == matches.len()).then_some(model.len()));
                    self.pair(
                        opening,
                        closing,
                        end,
                        &entry.text,
                        model,
                        &matches,
                        source,
                        deleted,
                        next,
                        &model_groups,
                        selected,
                    );
                }
            }
            if !entry.openings.is_empty() {
                pending = match entry.openings.as_mut_slice() {
                    [opening] if pending.is_none() && opening.direct && opening.separated => {
                        Some(Opening {
                            instruction: std::mem::take(&mut opening.instruction),
                            ..*opening
                        })
                    }
                    _ => None,
                };
            }
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn pair(
        &mut self,
        opening: Opening,
        closing: &Closing<'_>,
        end: usize,
        text: &str,
        model: &[BlockContent],
        matches: &[Option<usize>],
        source: &[(Vec<u32>, &XmlElement)],
        deleted: bool,
        next: Option<usize>,
        model_groups: &[usize],
        selected: &impl Fn(&XmlElement, &Paragraph) -> bool,
    ) {
        let Some(owner) = matches[opening.block] else {
            return;
        };
        let Some(opener) = paragraph(&model[owner]) else {
            return;
        };
        let Some(field) = opener.content.iter().rposition(|child| {
            matches!(
                child,
                ParagraphContent::Inline(InlineNode::ComplexField(field))
                    if field.instruction.trim() == opening.instruction.trim()
                        && spans_paragraphs(field)
            )
        }) else {
            return;
        };
        let (target_source, target, removed) = match matches[end] {
            Some(target) if target > owner => (end, target, false),
            Some(_) => return,
            None if closing.standalone() => {
                let Some(next) = next else { return };
                let Some((index, target)) = matches[opening.block..end]
                    .iter()
                    .enumerate()
                    .rev()
                    .find_map(|(index, matched)| Some((opening.block + index, (*matched)?)))
                    .filter(|&(_, block)| block >= owner && block < next)
                else {
                    return;
                };
                if (!deleted && target + 1 != next)
                    || matches[opening.block..end]
                        .iter()
                        .flatten()
                        .any(|&block| block > target)
                {
                    return;
                }
                (index, target, true)
            }
            None => return,
        };
        if model_groups[owner] != model_groups[target] {
            return;
        }
        let Some(closer) = paragraph(&model[target]) else {
            return;
        };
        let writes_opener = selected(source[opening.block].1, opener);
        let writes_closer = selected(source[end].1, closer);
        if !removed && !writes_opener && !writes_closer {
            return;
        }
        if removed || (writes_closer && !has_end(closer)) {
            let Some(content) = place_closing(source[end].1, closer, closing, text, removed) else {
                return;
            };
            let entry = self.paragraphs.entry(paragraph_key(closer)).or_default();
            if entry.closing.is_some() {
                return;
            }
            entry.closing = Some(content);
            if removed
                && let Some(ordinal) = closer
                    .source_ordinal
                    .or(source[target_source].1.paragraph_ordinal)
            {
                self.carried.insert(ordinal);
            }
        }
        self.paired.extend(
            [
                source[opening.block].1.paragraph_ordinal,
                source[end].1.paragraph_ordinal,
            ]
            .into_iter()
            .flatten(),
        );
        if removed && let Some(ordinal) = source[end].1.paragraph_ordinal {
            self.moved.insert(ordinal);
        }
        if writes_opener {
            let field = opener.content[..field]
                .iter()
                .filter(|child| {
                    matches!(child, ParagraphContent::Inline(InlineNode::ComplexField(_)))
                })
                .count();
            self.paragraphs
                .entry(paragraph_key(opener))
                .or_default()
                .opening = Some(field);
        }
    }
}
