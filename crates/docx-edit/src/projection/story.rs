use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use docx_parse::BlockContent;
use docx_parse::inline::{InlineNode, Run, RunContent};
use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::s9::S9DocumentBodyWire;
use yrs::Any;

use super::{
    Attrs, HOOKS, Lowered, LoweringInput, ProjectionContext, ProjectionError, ProjectionHooks, StoryPosition, StoryProjection,
    StoryState, attr_string, embed_hook, fields_dispatch, markers_dispatch, paragraph, runs,
};
use crate::{ParagraphProperties, SegmentContent, StorySegment};

struct PlacedSegment {
    offset: usize,
    segment: StorySegment,
    projected: Option<Vec<ParagraphContent>>,
}

pub(super) fn project(
    context: &mut ProjectionContext<'_, '_>,
    id: &str,
    state: StoryState,
) -> Result<StoryProjection, ProjectionError> {
    let mut story = StoryPosition { id, state, ..StoryPosition::default() };
    story.result.story_id = id.to_owned();
    let read_error = |error| ProjectionError::StoryRead { story_id: id.to_owned(), error };
    let segments = context.doc.story_segments_in(context.txn, id).map_err(read_error)?;
    let saved_ids = context.doc.story_paragraph_ids_in(context.txn, id, context.source_index)
        .map_err(read_error)?;
    for hook in HOOKS {
        hook.setup_story(context, &mut story, &segments)?;
    }
    let segments = tabs_as_embeds(segments);
    let mut items = Vec::new();
    let mut raw_blocks = RawBlocks::new(&context.source.blocks);
    let mut deferred_stories = BTreeSet::from([id.to_owned()]);
    for segment in segments {
        match &segment.content {
            SegmentContent::Text(text) => {
                let offset = story.state.offset;
                story.state.offset += text.encode_utf16().count();
                items.push(PlacedSegment { offset, segment, projected: None });
            }
            SegmentContent::Pilcrow(properties) => {
                let saved = saved_ids.get(story.state.paragraph_index).cloned().flatten();
                let source = context.source.paragraphs.get(&properties.para_id).cloned();
                if let Some(source) = &source {
                    raw_blocks.before_paragraph(source, &mut story.result.blocks);
                }
                finish_paragraph(context, &mut story, &items, Some(properties), saved, source.as_deref())?;
                items.clear();
                story.state.paragraph_index += 1;
                story.state.offset += 1;
                story.state.paragraph_start = story.state.offset;
            }
            SegmentContent::OtherEmbed { kind, payload } if container(kind, payload) => {
                let source_id = container_id(&mut story, kind, payload);
                if let Some(id) = &source_id
                    && let Some(block) = context.source.containers.get(id)
                {
                    raw_blocks.before_block(block, &mut story.result.blocks);
                }
                let hook = embed_hook(kind, payload);
                let mut lowered = hook.lower(context, &mut story, LoweringInput { kind, payload, segment: &segment })?;
                if matches!(lowered, Lowered::Deferred) {
                    story.result.defer(hook.family(), kind, id, Some(story.state.paragraph_index));
                    defer_descendant_containers(context, &mut story.result, kind, payload, &mut deferred_stories)?;
                    if let Some(source) = source_id.and_then(|id| context.source.containers.get(&id)) {
                        lowered = Lowered::Blocks(StoryProjection { blocks: vec![source.clone()], ..StoryProjection::default() });
                    }
                }
                for hook in HOOKS {
                    hook.wrap(context, &mut story, &segment, &mut lowered)?;
                }
                match lowered {
                    Lowered::Blocks(blocks) => story.result.append(blocks),
                    Lowered::Inline(inline) => {
                        items.push(PlacedSegment {
                            offset: story.state.offset,
                            segment,
                            projected: Some(inline),
                        });
                    }
                    Lowered::Deferred => {}
                }
                story.state.offset += 1;
            }
            SegmentContent::OtherEmbed { .. } => {
                let offset = story.state.offset;
                story.state.offset += 1;
                items.push(PlacedSegment { offset, segment, projected: None });
            }
        }
    }
    if !items.is_empty() {
        finish_paragraph(context, &mut story, &items, None, None, None)?;
    }
    raw_blocks.finish(&mut story.result.blocks);
    for hook in HOOKS {
        hook.finish_story(context, &mut story)?;
    }
    Ok(story.result)
}

fn tabs_as_embeds(segments: Vec<StorySegment>) -> Vec<StorySegment> {
    let mut expanded = Vec::with_capacity(segments.len());
    for segment in segments {
        if let SegmentContent::Text(text) = &segment.content
            && text.contains('\t')
        {
            for (index, part) in text.split('\t').enumerate() {
                if index != 0 {
                    expanded.push(StorySegment {
                        content: SegmentContent::OtherEmbed { kind: "tab".to_owned(), payload: Attrs::new() },
                        attributes: segment.attributes.clone(),
                    });
                }
                if !part.is_empty() {
                    expanded.push(StorySegment {
                        content: SegmentContent::Text(part.to_owned()),
                        attributes: segment.attributes.clone(),
                    });
                }
            }
        } else {
            expanded.push(segment);
        }
    }
    expanded
}

fn container(kind: &str, payload: &Attrs) -> bool {
    (crate::segments::is_block_embed(kind) && matches!(kind, "table" | "blockSdt"))
        || (kind == "opaque" && matches!(opaque_type(payload), Some("blockSdt" | "pageBreak")))
}

pub(super) fn opaque_type(payload: &Attrs) -> Option<&str> {
    let Any::Map(blob) = payload.get("blob")? else { return None };
    crate::read_types::payload_str(blob.get("type"))
}

fn container_id(story: &mut StoryPosition<'_>, kind: &str, payload: &Attrs) -> Option<String> {
    match kind {
        "table" => {
            let index = story.state.table_index;
            story.state.table_index += 1;
            let cell = first_table_cell(payload);
            if let Some(cell) = cell {
                crate::structured::source::cell_table(cell).map(str::to_owned)
            } else {
                Some(format!("{}:t{index}", story.id))
            }
        }
        "blockSdt" => {
            story.state.sdt_index += 1;
            attr_string(payload, "story").map(crate::read_types::block_control_id)
        }
        "opaque" if opaque_type(payload) == Some("blockSdt") => {
            let index = story.state.sdt_index;
            story.state.sdt_index += 1;
            Some(format!("{}:sdt{index}", story.id))
        }
        _ => None,
    }
}

fn first_table_cell(payload: &Attrs) -> Option<&str> {
    table_cell_stories(payload).next()
}

fn table_cell_stories(payload: &Attrs) -> impl DoubleEndedIterator<Item = &str> {
    fn array(value: Option<&Any>) -> &[Any] {
        match value {
            Some(Any::Array(values)) => values,
            _ => &[],
        }
    }

    array(payload.get("rows")).iter().flat_map(|row| {
        if let Any::Map(row) = row { array(row.get("cells")) } else { &[] }
    }).filter_map(|cell| match cell {
        Any::Map(cell) => crate::read_types::payload_str(cell.get("story")),
        _ => None,
    })
}

fn push_child_stories(pending: &mut Vec<String>, kind: &str, payload: &Attrs) {
    if kind == "table" {
        pending.extend(table_cell_stories(payload).rev().map(str::to_owned));
    } else if kind == "blockSdt" && let Some(id) = attr_string(payload, "story") {
        pending.push(id.to_owned());
    }
}

fn defer_descendant_containers(
    context: &ProjectionContext<'_, '_>,
    result: &mut StoryProjection,
    kind: &str,
    payload: &Attrs,
    visited: &mut BTreeSet<String>,
) -> Result<(), ProjectionError> {
    let mut pending = Vec::new();
    push_child_stories(&mut pending, kind, payload);
    while let Some(id) = pending.pop() {
        if !visited.insert(id.clone()) {
            continue;
        }
        let segments = context.doc.story_segments_in(context.txn, &id)
            .map_err(|error| ProjectionError::StoryRead { story_id: id.clone(), error })?;
        let mut paragraph = 0;
        for segment in segments {
            match segment.content {
                SegmentContent::Pilcrow(_) => paragraph += 1,
                SegmentContent::OtherEmbed { kind, payload } if container(&kind, &payload) => {
                    let family = embed_hook(&kind, &payload).family();
                    result.defer(family, &kind, &id, Some(paragraph));
                    push_child_stories(&mut pending, &kind, &payload);
                }
                _ => {}
            }
        }
    }
    Ok(())
}

fn finish_paragraph(
    context: &mut ProjectionContext<'_, '_>,
    story: &mut StoryPosition<'_>,
    items: &[PlacedSegment],
    properties: Option<&ParagraphProperties>,
    saved_id: Option<String>,
    source: Option<&Paragraph>,
) -> Result<(), ProjectionError> {
    let empty = Attrs::new();
    let attrs = properties.map_or(&empty, |properties| &properties.values);
    story.state.attributes.clone_from(attrs);
    let first_deferral = story.result.deferred.len();
    let mut content = lower_items(context, story, attrs, items, source)?;
    if properties.is_none() && content.is_empty() {
        for deferral in &mut story.result.deferred[first_deferral..] {
            deferral.paragraph = None;
        }
        return Ok(());
    }
    if let Some(source) = source {
        carry_side_markers(&mut content, &source.content);
    }
    let formatting = properties.map(|properties| {
        paragraph::formatting(context, story.id, &story.state, properties, source)
    }).transpose()?.flatten();
    let mut paragraph = paragraph::new(source, saved_id, formatting, content);
    for hook in HOOKS {
        hook.finish_paragraph(context, story, properties, source, &mut paragraph)?;
    }
    story.result.paragraph_session_keys.push(properties.map_or_else(String::new, |props| props.para_id.clone()));
    story.result.blocks.push(BlockContent::Paragraph(Arc::new(paragraph)));
    Ok(())
}

fn lower_items(
    context: &mut ProjectionContext<'_, '_>,
    story: &mut StoryPosition<'_>,
    attrs: &Attrs,
    items: &[PlacedSegment],
    source: Option<&Paragraph>,
) -> Result<Vec<ParagraphContent>, ProjectionError> {
    let plain = |item: &PlacedSegment| {
        item.projected.is_none() && matches!(item.segment.content, SegmentContent::Text(_))
            && !fields_dispatch::MARKS.iter().chain(markers_dispatch::MARKS)
                .any(|kind| has_mark(&item.segment, kind))
    };
    if items.iter().all(plain) {
        return runs::lower_plain(context, &story.state, attrs, &text_items(items), source, true);
    }

    let mut fallback = SourceInlines::new(source);
    let mut content = Vec::new();
    let paragraph_end = story.state.offset;
    let mut index = 0;
    while index < items.len() {
        let placed = &items[index];
        let item = &placed.segment;
        story.state.offset = placed.offset;
        if let Some(projected) = &placed.projected {
            content.extend_from_slice(projected);
            index += 1;
            continue;
        }
        if plain(placed) {
            let end = items[index..].iter().position(|item| !plain(item)).map_or(items.len(), |count| index + count);
            let text = text_items(&items[index..end]);
            content.extend(runs::lower_plain(context, &story.state, attrs, &text, None, false)?);
            index = end;
            continue;
        }
        let mut lowered = match &item.content {
            SegmentContent::Text(text) => {
                let mark = runs::attributes(&item.attributes)?;
                let run = runs::make_run(context, &story.state, attrs, text, &mark, None)?;
                let kind = ["del", "ins", "hyperlink", "rPrChange"].into_iter().find(|kind| has_mark(item, kind));
                let mut lowered = Lowered::Deferred;
                for mark in fields_dispatch::MARKS.iter().copied().filter(|mark| has_mark(item, mark)) {
                    lowered = fields_dispatch::Hooks.lower(context, story, LoweringInput {
                        kind: mark, payload: &item.attributes, segment: item,
                    })?;
                    if !matches!(lowered, Lowered::Deferred) {
                        break;
                    }
                }
                if matches!(lowered, Lowered::Deferred) {
                    lowered = Lowered::Inline(vec![fallback.marked(kind, text).unwrap_or_else(|| runs::inline(run))]);
                }
                lowered
            }
            SegmentContent::OtherEmbed { kind, payload } => {
                let hook = embed_hook(kind, payload);
                match hook.lower(context, story, LoweringInput { kind, payload, segment: item })? {
                    Lowered::Deferred => {
                        story.result.defer(hook.family(), kind, story.id, Some(story.state.paragraph_index));
                        let source_kind = if kind == "field" && attr_string(payload, "modelKind") == Some("commentReference") {
                            "commentReference"
                        } else {
                            kind
                        };
                        Lowered::Inline(fallback.take(source_kind, payload).into_iter().collect())
                    }
                    lowered => lowered,
                }
            }
            SegmentContent::Pilcrow(_) => unreachable!("pilcrows close segments"),
        };
        for hook in HOOKS {
            hook.wrap(context, story, item, &mut lowered)?;
        }
        match lowered {
            Lowered::Inline(mut inline) => content.append(&mut inline),
            Lowered::Blocks(blocks) => story.result.append(blocks),
            Lowered::Deferred => {}
        }
        index += 1;
    }
    story.state.offset = paragraph_end;
    Ok(content)
}

fn text_items(items: &[PlacedSegment]) -> Vec<(&str, &Attrs)> {
    items.iter().map(|placed| &placed.segment).filter_map(|segment| match &segment.content {
        SegmentContent::Text(text) => Some((text.as_str(), &segment.attributes)),
        _ => None,
    }).collect()
}

pub(super) fn has_mark(item: &StorySegment, key: &str) -> bool {
    item.attributes.get(key).is_some_and(|value| !matches!(value, Any::Null | Any::Undefined | Any::Bool(false)))
}

#[derive(Default)]
struct RawBlocks {
    before: Vec<(BlockContent, Vec<BlockContent>)>,
    trailing: Vec<BlockContent>,
}

impl RawBlocks {
    fn new(blocks: &[BlockContent]) -> Self {
        let mut result = Self::default();
        for block in blocks {
            if matches!(block, BlockContent::RawXml(_)) {
                result.trailing.push(block.clone());
            } else if !result.trailing.is_empty() {
                result.before.push((block.clone(), std::mem::take(&mut result.trailing)));
            }
        }
        result
    }

    fn before_paragraph(&mut self, source: &Paragraph, into: &mut Vec<BlockContent>) {
        for (block, raw) in &mut self.before {
            if matches!(block, BlockContent::Paragraph(paragraph) if std::ptr::eq(paragraph.as_ref(), source)) {
                into.append(raw);
                break;
            }
        }
    }

    fn before_block(&mut self, source: &BlockContent, into: &mut Vec<BlockContent>) {
        for (block, raw) in &mut self.before {
            let same = match (&*block, source) {
                (BlockContent::Table(left), BlockContent::Table(right)) => Arc::ptr_eq(left, right),
                (BlockContent::BlockSdt(left), BlockContent::BlockSdt(right)) => Arc::ptr_eq(left, right),
                _ => false,
            };
            if same {
                into.append(raw);
                break;
            }
        }
    }

    fn finish(self, into: &mut Vec<BlockContent>) {
        into.extend(self.before.into_iter().flat_map(|(_, raw)| raw));
        into.extend(self.trailing);
    }
}

struct SourceInlines {
    items: BTreeMap<&'static str, Vec<Option<ParagraphContent>>>,
}

impl SourceInlines {
    fn new(source: Option<&Paragraph>) -> Self {
        let mut items: BTreeMap<&'static str, Vec<Option<ParagraphContent>>> = BTreeMap::new();
        for content in source.into_iter().flat_map(|source| &source.content) {
            if let ParagraphContent::Inline(InlineNode::Run(run)) = content {
                if run.property_changes.is_some() {
                    items.entry("rPrChange").or_default().push(Some(content.clone()));
                }
                for item in &run.content {
                    if let Some(kind) = run_item_kind(item) {
                        items.entry(kind).or_default().push(Some(runs::inline(Run {
                            content: vec![item.clone()],
                            ..run.clone()
                        })));
                    }
                }
            } else {
                let key = match content.node_type() {
                    "hyperlink" => "hyperlink",
                    "insertion" | "moveTo" => "ins",
                    "deletion" | "moveFrom" => "del",
                    "simpleField" | "complexField" => "field",
                    "inlineSdt" => "sdt",
                    "mathEquation" => "math",
                    "rawXml" => "rawInline",
                    _ => continue,
                };
                items.entry(key).or_default().push(Some(content.clone()));
            }
        }
        Self { items }
    }

    fn take(&mut self, kind: &str, payload: &Attrs) -> Option<ParagraphContent> {
        let candidates = self.items.get_mut(kind)?;
        let raw_xml = attr_string(payload, "xml");
        candidates.iter_mut().find(|item| {
            item.as_ref().is_some_and(|item| match (item, raw_xml) {
                (ParagraphContent::Inline(InlineNode::RawXml(raw)), Some(xml)) => raw.xml == xml,
                _ => true,
            })
        })?.take()
    }

    fn marked(&mut self, kind: Option<&str>, text: &str) -> Option<ParagraphContent> {
        let candidates = self.items.get_mut(kind?)?;
        candidates.iter_mut().find(|item| {
            item.as_ref().is_some_and(|item| match item {
                ParagraphContent::Inline(InlineNode::Hyperlink(link)) => inline_text(&link.children) == text,
                ParagraphContent::Tracked(track) => inline_text(&track.content) == text,
                ParagraphContent::Inline(InlineNode::Run(run)) => run_text(run) == text,
                _ => false,
            })
        })?.take()
    }
}

fn inline_text(content: &[InlineNode]) -> String {
    let mut text = String::new();
    let mut pending = vec![content.iter()];
    while let Some(nodes) = pending.last_mut() {
        match nodes.next() {
            Some(InlineNode::Run(run)) => text.push_str(&run_text(run)),
            Some(InlineNode::Hyperlink(link)) => pending.push(link.children.iter()),
            Some(_) => {}
            None => { pending.pop(); }
        }
    }
    text
}

fn run_text(run: &Run) -> String {
    run.content.iter().filter_map(|content| match content {
        RunContent::Text { text, .. } => Some(text.as_str()),
        RunContent::SoftHyphen => Some("\u{ad}"),
        RunContent::NoBreakHyphen => Some("\u{2011}"),
        _ => None,
    }).collect()
}

fn run_item_kind(content: &RunContent) -> Option<&'static str> {
    match content {
        RunContent::Tab => Some("tab"),
        RunContent::Break { break_type, .. } => match break_type.as_deref() {
            Some("page") => Some("pageBreak"),
            Some("column") => Some("columnBreak"),
            _ => Some("break"),
        },
        RunContent::FootnoteRef { .. } | RunContent::EndnoteRef { .. } => Some("noteRef"),
        RunContent::CommentReference { .. } => Some("commentReference"),
        RunContent::Drawing { .. } => Some("image"),
        RunContent::Shape { .. } => Some("shape"),
        RunContent::HorizontalRule { .. } => Some("horizontalRule"),
        RunContent::Chart { .. } => Some("chart"),
        RunContent::OpaqueDrawing { .. } => Some("opaqueDrawing"),
        RunContent::FieldChar { .. } | RunContent::InstrText { .. } => Some("field"),
        _ => None,
    }
}

fn side_marker(inline: &ParagraphContent) -> bool {
    matches!(inline, ParagraphContent::RangeStart(_) | ParagraphContent::RangeEnd(_) | ParagraphContent::CommentRange(_))
        || matches!(inline.node_type(), "bookmarkStart" | "bookmarkEnd" | "rawXml")
}

fn carry_side_markers(content: &mut Vec<ParagraphContent>, source: &[ParagraphContent]) {
    let supported: Vec<_> = source.iter().filter(|inline| !side_marker(inline)).collect();
    if supported.len() != content.len() || !source.iter().any(side_marker) {
        return;
    }
    let same_units = supported.iter().zip(content.iter()).all(|(source, content)| {
        match (source, content) {
            (ParagraphContent::Inline(InlineNode::Run(source)), ParagraphContent::Inline(InlineNode::Run(content))) => {
                run_text(source) == run_text(content) && source.content.len() == content.content.len()
            }
            _ => *source == content,
        }
    });
    if same_units {
        let mut projected = std::mem::take(content).into_iter();
        content.extend(source.iter().filter_map(|inline| {
            if side_marker(inline) {
                (!projected.as_slice().contains(inline)).then(|| inline.clone())
            } else {
                projected.next()
            }
        }));
    }
}

pub(super) fn has_revision_ranges(block: &BlockContent) -> bool {
    matches!(block, BlockContent::Paragraph(paragraph) if paragraph.content.iter().any(|content| {
        matches!(content, ParagraphContent::RangeStart(_) | ParagraphContent::RangeEnd(_))
    }))
}

pub(super) fn has_comment_ranges(block: &BlockContent) -> bool {
    matches!(block, BlockContent::Paragraph(paragraph) if paragraph.content.iter()
        .any(|content| matches!(content, ParagraphContent::CommentRange(_))))
}

pub(super) fn clamp_sections(body: &mut S9DocumentBodyWire) {
    let Some(sections) = &mut body.sections else { return };
    let len = body.content.len();
    let mut closers = body.content.iter().enumerate().filter_map(|(index, block)| {
        matches!(block, BlockContent::Paragraph(paragraph) if paragraph.section_properties.is_some())
            .then_some(index + 1)
    });
    let mut start = 0;
    let count = sections.len();
    for (index, section) in sections.iter_mut().enumerate() {
        let end = if index + 1 == count {
            len
        } else {
            closers.next().unwrap_or(section.content_end).clamp(start, len)
        };
        section.content_start = start;
        section.content_end = end;
        start = end;
    }
}
