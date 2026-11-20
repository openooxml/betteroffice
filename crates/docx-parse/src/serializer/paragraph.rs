//! Paragraph properties and inline-content serialization.

use std::borrow::Borrow;

use crate::borders::Borders;
use crate::formatting::{
    NumberingProperties, ParagraphCnfStyle, ParagraphFormatting, ParagraphFrame,
};
use crate::inline::{
    BookmarkEnd, BookmarkStart, ComplexField, DirectionScope, Hyperlink, InlineNode, InlineSdt,
    MathEquation, Run, RunContent, SdtProperties, SimpleField, field_run_digest, field_runs,
};
use crate::paragraph::{
    CommentRange, Paragraph, ParagraphContent, ParagraphPropertyChange, RangeEnd, RangeStart,
    RawAttribute, TrackedChangeInfo, TrackedInline,
};
use crate::paragraph_identity::W14_NAMESPACE;
use crate::section::SectionProperties;
use crate::xml::ParseError;

use super::context::{DeferredFormData, SerializerContext, shared_direction_prefix};
use super::foundation::{BorderSide, write_border};
use super::raw::{
    validate_form_field_data, validate_math_subtree, validate_raw_subtree,
    validate_replayed_fragment,
};
use super::run::{
    append_generated, nonempty, nonempty_trimmed, normalized_tracked_id, serialize_deleted_run,
    serialize_run, serialize_text_formatting, write_shading,
};
use super::section::serialize_section_properties;
use super::xml_writer::{XmlWriter, int_attr, js_number};

/// A replayed attribute name must be exactly one XML name, or it could
/// smuggle markup into the tag.
pub(crate) fn is_safe_attribute_name(name: &str) -> bool {
    !name.is_empty()
        && name.chars().all(|character| {
            character.is_ascii_alphanumeric() || matches!(character, ':' | '-' | '_' | '.')
        })
}

/// Serialize one complete `w:p` element.
pub fn serialize_paragraph(
    paragraph: &Paragraph,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    context.enter_paragraph(paragraph.rendered_page_break_before == Some(true));
    let enclosing = std::mem::take(&mut context.direction);
    let result = serialize_paragraph_inner(paragraph, context);
    context.direction = enclosing;
    context.leave_paragraph();
    result.map(
        |xml| match context.record_paragraph(paragraph.source_ordinal, &xml) {
            Some(token) => token + &xml,
            None => xml,
        },
    )
}

/// The prefix a paragraph's Word 2010 attributes take: `w14`, unless the
/// paragraph rebinds it, then a prefix it binds to that namespace or a fresh
/// one. Also whether `w14` is rebound.
fn word_2010_prefix(attributes: &[RawAttribute]) -> (String, bool) {
    let rebound = attributes
        .iter()
        .any(|attribute| attribute.name == "xmlns:w14" && attribute.value.trim() != W14_NAMESPACE);
    if !rebound {
        return ("w14".to_owned(), false);
    }
    let declared = attributes.iter().find_map(|attribute| {
        attribute
            .name
            .strip_prefix("xmlns:")
            .filter(|_| attribute.value.trim() == W14_NAMESPACE)
    });
    let prefix = declared.map(str::to_owned).unwrap_or_else(|| {
        (0..)
            .map(|index| format!("w14p{index}"))
            .find(|candidate| {
                !attributes.iter().any(|attribute| {
                    attribute.name.split_once(':').is_some_and(|(used, local)| {
                        used == candidate || (used == "xmlns" && local == candidate)
                    })
                })
            })
            .unwrap_or_default()
    });
    (prefix, true)
}

fn serialize_paragraph_inner(
    paragraph: &Paragraph,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut writer = XmlWriter::with_capacity(512);
    writer.start_element("w:p");
    let para_id = nonempty(paragraph.para_id.as_deref());
    let text_id = nonempty(paragraph.text_id.as_deref());
    let (prefix, rebound) = word_2010_prefix(&paragraph.extra_attributes);
    if rebound
        && (para_id.is_some() || text_id.is_some())
        && !paragraph
            .extra_attributes
            .iter()
            .any(|attribute| attribute.name == format!("xmlns:{prefix}"))
    {
        writer.dynamic_attribute(&format!("xmlns:{prefix}"), W14_NAMESPACE);
    }
    if let Some(value) = para_id {
        writer.dynamic_attribute(&format!("{prefix}:paraId"), value);
    }
    if let Some(value) = text_id {
        writer.dynamic_attribute(&format!("{prefix}:textId"), value);
    }
    for attribute in &paragraph.extra_attributes {
        let replaced = para_id.is_some()
            && (attribute.name == "w:paraId"
                || (!rebound && attribute.name == "w14:paraId")
                || paragraph.para_id_attribute.as_deref() == Some(attribute.name.as_str()));
        if !replaced && is_safe_attribute_name(&attribute.name) {
            writer.dynamic_attribute(&attribute.name, &attribute.value);
        }
    }
    let properties = serialize_paragraph_formatting(
        paragraph.formatting.as_ref(),
        paragraph.property_changes.as_deref(),
        paragraph.p_pr_ins.as_ref(),
        paragraph.p_pr_del.as_ref(),
        false,
        paragraph.section_properties.as_ref(),
    )?;
    append_generated(&mut writer, &properties);
    let mut generated_comment_references = Vec::new();
    let mut upcoming = None;
    for (index, content) in paragraph.content.iter().enumerate() {
        if paragraph_content_zero_width(content) && context.direction.any_open() {
            if upcoming.is_none() {
                upcoming = Some(upcoming_paragraph_scopes(paragraph, Within::of(context)));
            }
            if let Some(scopes) = &upcoming {
                context.direction.look_ahead(scopes[index]);
            }
        }
        append_generated(&mut writer, &serialize_paragraph_content(content, context)?);
        if let ParagraphContent::CommentRange(marker) = content
            && !generated_comment_references.contains(&marker.id)
            && needs_comment_reference(paragraph, content)
        {
            writer
                .start_element("w:r")
                .start_element("w:rPr")
                .start_element("w:rStyle")
                .attribute("w:val", "CommentReference")
                .end_element()
                .end_element()
                .start_element("w:commentReference")
                .attribute("w:id", &js_number(marker.id))
                .end_element()
                .end_element();
            generated_comment_references.push(marker.id);
        }
    }
    let closing = context.direction.close_all();
    if !closing.is_empty() {
        append_generated(&mut writer, &closing);
    }
    writer.end_element();
    Ok(writer.finish())
}

/// Whether `content` is a `commentRangeEnd` whose comment no reference in the paragraph marks,
/// so a reference run is written after it.
fn needs_comment_reference(paragraph: &Paragraph, content: &ParagraphContent) -> bool {
    let ParagraphContent::CommentRange(marker) = content else {
        return false;
    };
    marker.node_type == "commentRangeEnd"
        && !paragraph.content.iter().any(|content| match content {
            ParagraphContent::Inline(node) => has_comment_reference(node, marker.id),
            ParagraphContent::Tracked(change) => change.content.iter().any(|node| {
                (matches!(
                    node,
                    InlineNode::Run(_) | InlineNode::Hyperlink(_) | InlineNode::InlineSdt(_)
                ) || node.has_tracked_control(false, true))
                    && has_comment_reference(node, marker.id)
            }),
            _ => false,
        })
}

/// The range and annotation marks a direction scope may hold, each one empty `w:` element.
const MARKS_WITHOUT_WIDTH: [&str; 19] = [
    "bookmarkStart",
    "bookmarkEnd",
    "commentRangeStart",
    "commentRangeEnd",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
    "permStart",
    "permEnd",
    "proofErr",
];

/// Whether `xml` is one of [`MARKS_WITHOUT_WIDTH`], written where it sits and taking no width.
fn empty_word_mark(xml: &str) -> bool {
    let Some(tag) = xml.trim().strip_prefix("<w:") else {
        return false;
    };
    let name = tag
        .split([' ', '\t', '\n', '\r', '/', '>'])
        .next()
        .unwrap_or_default();
    if !MARKS_WITHOUT_WIDTH.contains(&name) {
        return false;
    }
    let mut quote = None;
    for (index, character) in tag.char_indices() {
        match quote {
            Some(open) if character == open => quote = None,
            Some(_) => {}
            None if matches!(character, '"' | '\'') => quote = Some(character),
            None if character == '>' => {
                return index + 1 == tag.len() && tag[..index].ends_with('/');
            }
            None => {}
        }
    }
    false
}

fn zero_width(node: &InlineNode) -> bool {
    match node {
        InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_) => true,
        InlineNode::RawXml(raw) => empty_word_mark(&raw.xml),
        _ => false,
    }
}

fn paragraph_content_zero_width(content: &ParagraphContent) -> bool {
    match content {
        ParagraphContent::Inline(node) => zero_width(node),
        ParagraphContent::Tracked(_) => false,
        ParagraphContent::RangeStart(_)
        | ParagraphContent::RangeEnd(_)
        | ParagraphContent::CommentRange(_) => true,
    }
}

/// Whether a node is written inside a content control and inside a tracked change, which decides
/// the children a hyperlink there writes.
#[derive(Clone, Copy)]
struct Within {
    control: bool,
    revision: bool,
}

impl Within {
    fn of(context: &SerializerContext) -> Self {
        Self {
            control: context.in_control,
            revision: context.in_revision,
        }
    }
}

/// Whether a hyperlink written `within` writes its child `child`.
fn link_writes(child: &InlineNode, within: Within) -> bool {
    match child {
        _ if within.control && within.revision => true,
        InlineNode::Run(_)
        | InlineNode::BookmarkStart(_)
        | InlineNode::BookmarkEnd(_)
        | InlineNode::RawXml(_) => true,
        InlineNode::InlineSdt(_) => child.has_tracked_control(false, within.revision),
        _ => false,
    }
}

/// The children a hyperlink written `within` writes, in order.
fn written_link_children(hyperlink: &Hyperlink, within: Within) -> Vec<&InlineNode> {
    let expanded = within.control && within.revision;
    let nested = hyperlink.structured_children.as_ref().filter(|children| {
        expanded
            || children
                .iter()
                .any(|child| child.has_tracked_control(false, within.revision))
    });
    nested
        .unwrap_or(&hyperlink.children)
        .iter()
        .filter(|child| link_writes(child, within))
        .collect()
}

/// The scopes `node`, written `within`, keeps open where it starts.
fn leading_scopes(node: &InlineNode, within: Within) -> &[DirectionScope] {
    let mut shared = None;
    merge_direction_scopes(node, &mut shared, within);
    shared.unwrap_or_default()
}

/// The scopes the next node with width keeps open where it starts, from each of a wrapper
/// element's nodes on: those of what is written after the nodes where none of them follows.
/// They are found in one pass from the end, when a node without width first needs them.
struct Upcoming<'a, N> {
    nodes: &'a [N],
    after: Option<&'a [DirectionScope]>,
    within: Within,
    scopes: Option<Vec<Option<&'a [DirectionScope]>>>,
}

impl<'a, N: Borrow<InlineNode>> Upcoming<'a, N> {
    fn new(nodes: &'a [N], after: Option<&'a [DirectionScope]>, within: Within) -> Self {
        Self {
            nodes,
            after,
            within,
            scopes: None,
        }
    }

    /// Notes, before `nodes[index]` when it takes no width inside an open scope, the scopes the
    /// next node with width keeps open where it starts.
    fn look_ahead(&mut self, index: usize, context: &mut SerializerContext) {
        if !zero_width(self.nodes[index].borrow()) || !context.direction.any_open() {
            return;
        }
        if self.scopes.is_none() {
            self.scopes = Some(upcoming_scopes(self.nodes, self.after, self.within));
        }
        if let Some(scopes) = &self.scopes {
            context.direction.look_ahead(scopes[index]);
        }
    }
}

fn upcoming_scopes<'a, N: Borrow<InlineNode>>(
    nodes: &'a [N],
    after: Option<&'a [DirectionScope]>,
    within: Within,
) -> Vec<Option<&'a [DirectionScope]>> {
    let mut next = after;
    let mut upcoming = vec![None; nodes.len()];
    for (index, node) in nodes.iter().enumerate().rev() {
        let node: &InlineNode = node.borrow();
        if !zero_width(node) {
            next = Some(leading_scopes(node, within));
        }
        upcoming[index] = next;
    }
    upcoming
}

/// The scopes the first paragraph content with width from each index on keeps open where it
/// starts, written `within`, found in one pass from the end; `None` where a reference run is
/// written before it or none follows.
fn upcoming_paragraph_scopes(
    paragraph: &Paragraph,
    within: Within,
) -> Vec<Option<&[DirectionScope]>> {
    let revision = Within {
        revision: true,
        ..within
    };
    let mut next = None;
    let mut upcoming = vec![None; paragraph.content.len()];
    for (index, content) in paragraph.content.iter().enumerate().rev() {
        match content {
            ParagraphContent::Inline(node) if !zero_width(node) => {
                next = Some(leading_scopes(node, within));
            }
            ParagraphContent::Tracked(change) => {
                let scopes = shared_direction_scopes(&change.content, revision);
                next = Some(scopes.unwrap_or_default());
            }
            _ if needs_comment_reference(paragraph, content) => next = None,
            _ => {}
        }
        upcoming[index] = next;
    }
    upcoming
}

#[allow(clippy::too_many_arguments)]
pub fn serialize_paragraph_formatting(
    formatting: Option<&ParagraphFormatting>,
    property_changes: Option<&[ParagraphPropertyChange]>,
    p_pr_ins: Option<&TrackedChangeInfo>,
    p_pr_del: Option<&TrackedChangeInfo>,
    base_only: bool,
    section_properties: Option<&SectionProperties>,
) -> Result<String, ParseError> {
    let mut body = XmlWriter::with_capacity(512);
    if let Some(formatting) = formatting {
        if let Some(value) = nonempty(formatting.style_id.as_deref()) {
            empty_attr(&mut body, "w:pStyle", "w:val", value);
        }
        on_off(&mut body, "w:keepNext", formatting.keep_next);
        on_off(&mut body, "w:keepLines", formatting.keep_lines);
        on_off(
            &mut body,
            "w:contextualSpacing",
            formatting.contextual_spacing,
        );
        on_off(&mut body, "w:pageBreakBefore", formatting.page_break_before);
        write_frame(&mut body, formatting.frame.as_ref());
        on_off(&mut body, "w:widowControl", formatting.widow_control);

        if formatting.num_pr != formatting.num_pr_from_style
            || formatting.num_pr_from_style.is_none()
        {
            write_numbering(&mut body, formatting.num_pr.as_ref());
        }
        write_paragraph_borders(&mut body, formatting.borders.as_ref());
        write_shading(&mut body, formatting.shading.as_ref());
        write_tabs(&mut body, formatting.tabs.as_deref());
        on_off(
            &mut body,
            "w:suppressLineNumbers",
            formatting.suppress_line_numbers,
        );
        on_off(
            &mut body,
            "w:suppressAutoHyphens",
            formatting.suppress_auto_hyphens,
        );
        on_off(&mut body, "w:autoSpaceDE", formatting.auto_space_de);
        on_off(&mut body, "w:autoSpaceDN", formatting.auto_space_dn);
        on_off(&mut body, "w:bidi", formatting.bidi);
        on_off(&mut body, "w:adjustRightInd", formatting.adjust_right_ind);
        on_off(&mut body, "w:snapToGrid", formatting.snap_to_grid);
        write_spacing(&mut body, formatting);
        write_indentation(&mut body, formatting);
        if let Some(value) = nonempty(formatting.alignment.as_deref()) {
            empty_attr(&mut body, "w:jc", "w:val", value);
        }
        if let Some(value) = formatting.outline_level {
            empty_attr(&mut body, "w:outlineLvl", "w:val", &js_number(value));
        }
        write_paragraph_cnf_style(&mut body, formatting.cnf_style.as_ref());
    }

    if !base_only {
        write_paragraph_mark_properties(&mut body, formatting, p_pr_ins, p_pr_del);
        if let Some(properties) = section_properties {
            // A paragraph-level sectPr ends a section even when it holds nothing.
            let section = serialize_section_properties(Some(properties));
            append_generated(
                &mut body,
                if section.is_empty() {
                    "<w:sectPr/>"
                } else {
                    &section
                },
            );
        }
        if let Some(change) = property_changes.and_then(|changes| changes.first()) {
            append_generated(&mut body, &serialize_paragraph_property_change(change)?);
        }
    }
    let body = body.finish();
    if body.is_empty() {
        Ok(String::new())
    } else {
        Ok(format!("<w:pPr>{body}</w:pPr>"))
    }
}

fn serialize_paragraph_property_change(
    change: &ParagraphPropertyChange,
) -> Result<String, ParseError> {
    let previous = serialize_paragraph_formatting(
        change.previous_formatting.as_ref(),
        None,
        None,
        None,
        true,
        None,
    )?;
    let previous = if previous.is_empty() {
        "<w:pPr/>".to_owned()
    } else {
        previous
    };
    let mut writer = XmlWriter::with_capacity(previous.len() + 128);
    writer
        .start_element("w:pPrChange")
        .attribute("w:id", &normalized_tracked_id(change.info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&change.info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = change.info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    append_generated(&mut writer, &previous);
    writer.end_element();
    Ok(writer.finish())
}

pub fn serialize_paragraph_content(
    content: &ParagraphContent,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let marker = match content {
        ParagraphContent::Inline(node) => return serialize_inline_node(node, context),
        ParagraphContent::Tracked(change) => return serialize_tracked_change(change, context),
        ParagraphContent::RangeStart(marker) => serialize_range_start(marker)?,
        ParagraphContent::RangeEnd(marker) => serialize_range_end(marker)?,
        ParagraphContent::CommentRange(marker) => serialize_comment_range(marker)?,
    };
    Ok(prefixed(context.direction.unscoped(true), marker))
}

pub(crate) fn serialize_inline_node(
    node: &InlineNode,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let unscoped = match node {
        InlineNode::Run(run) => return serialize_context_run(run, context),
        InlineNode::Hyperlink(hyperlink) => return serialize_hyperlink(hyperlink, context),
        InlineNode::SimpleField(field) => return serialize_simple_field(field, context),
        InlineNode::ComplexField(field) => return serialize_complex_field(field, context),
        InlineNode::InlineSdt(sdt) => return serialize_inline_sdt(sdt, context),
        InlineNode::Tracked(change) => return serialize_tracked_change(change, context),
        InlineNode::BookmarkStart(bookmark) => serialize_bookmark_start(bookmark),
        InlineNode::BookmarkEnd(bookmark) => serialize_bookmark_end(bookmark),
        InlineNode::Math(math) => serialize_math(math)?,
        InlineNode::RawXml(raw) => {
            validate_replayed_fragment(&raw.xml)?;
            raw.xml.clone()
        }
    };
    if unscoped.is_empty() {
        return Ok(unscoped);
    }
    Ok(prefixed(
        context.direction.unscoped(zero_width(node)),
        unscoped,
    ))
}

pub(super) fn serialize_bookmark_start(bookmark: &BookmarkStart) -> String {
    let mut writer = XmlWriter::with_capacity(96);
    writer
        .start_element("w:bookmarkStart")
        .attribute("w:id", &js_number(bookmark.id))
        .attribute("w:name", &bookmark.name);
    if let Some(value) = bookmark.col_first {
        writer.attribute("w:colFirst", &js_number(value));
    }
    if let Some(value) = bookmark.col_last {
        writer.attribute("w:colLast", &js_number(value));
    }
    writer.end_element();
    writer.finish()
}

pub(super) fn serialize_bookmark_end(bookmark: &BookmarkEnd) -> String {
    let mut writer = XmlWriter::with_capacity(40);
    writer
        .start_element("w:bookmarkEnd")
        .attribute("w:id", &js_number(bookmark.id))
        .end_element();
    writer.finish()
}

fn serialize_context_run(run: &Run, context: &mut SerializerContext) -> Result<String, ParseError> {
    serialize_scoped_run(run, context.deletion, context)
}

fn serialize_scoped_run(
    run: &Run,
    deletion: bool,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let opened = context.direction.transition(&run.direction_scopes);
    let xml = if deletion {
        serialize_deleted_run(run, context)?
    } else {
        serialize_run(run, context)?
    };
    Ok(prefixed(opened, xml))
}

fn prefixed(mut before: String, xml: String) -> String {
    if before.is_empty() {
        return xml;
    }
    before.push_str(&xml);
    before
}

fn merge_run_direction_scopes<'a>(run: &'a Run, shared: &mut Option<&'a [DirectionScope]>) {
    let scopes = run.direction_scopes.as_slice();
    *shared = Some(match *shared {
        Some(current) => &current[..shared_direction_prefix(current, scopes)],
        None => scopes,
    });
}

fn merge_direction_scopes<'a>(
    node: &'a InlineNode,
    shared: &mut Option<&'a [DirectionScope]>,
    within: Within,
) {
    match node {
        InlineNode::Run(run) => merge_run_direction_scopes(run, shared),
        InlineNode::Hyperlink(link) => {
            for child in written_link_children(link, within) {
                merge_direction_scopes(child, shared, within);
            }
        }
        InlineNode::SimpleField(field) => {
            for run in &field.content {
                merge_run_direction_scopes(run, shared);
            }
        }
        InlineNode::ComplexField(field) => {
            for run in field.field_code.iter().chain(&field.field_result) {
                merge_run_direction_scopes(run, shared);
            }
        }
        InlineNode::InlineSdt(sdt) => {
            let within = Within {
                control: true,
                ..within
            };
            for child in &sdt.content {
                merge_direction_scopes(child, shared, within);
            }
        }
        InlineNode::Tracked(change) => {
            let within = Within {
                revision: true,
                ..within
            };
            for child in &change.content {
                merge_direction_scopes(child, shared, within);
            }
        }
        InlineNode::BookmarkStart(_)
        | InlineNode::BookmarkEnd(_)
        | InlineNode::Math(_)
        | InlineNode::RawXml(_) => {}
    }
}

/// The `w:dir`/`w:bdo` scopes every run in `nodes`, written `within`, sits in, outermost first;
/// `None` without runs.
fn shared_direction_scopes(nodes: &[InlineNode], within: Within) -> Option<&[DirectionScope]> {
    let mut shared = None;
    for node in nodes {
        merge_direction_scopes(node, &mut shared, within);
    }
    shared
}

fn serialize_hyperlink(
    hyperlink: &Hyperlink,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut children = String::new();
    let within = Within::of(context);
    let written = written_link_children(hyperlink, within);
    let mut shared = None;
    for &child in &written {
        merge_direction_scopes(child, &mut shared, within);
    }
    let (before, base) = context.direction.enter_wrapper(shared);
    let mut upcoming = Upcoming::new(&written, None, within);
    for (index, child) in written.iter().enumerate() {
        upcoming.look_ahead(index, context);
        children.push_str(&serialize_inline_node(child, context)?);
    }
    children.push_str(&context.direction.leave_wrapper(base));
    let xml = serialize_hyperlink_element(hyperlink, children);
    Ok(prefixed(before, xml))
}

fn serialize_hyperlink_element(hyperlink: &Hyperlink, children: String) -> String {
    let has_attributes = nonempty(hyperlink.relationship_id.as_deref()).is_some()
        || nonempty(hyperlink.anchor.as_deref()).is_some()
        || nonempty(hyperlink.tooltip.as_deref()).is_some()
        || nonempty(hyperlink.target.as_deref()).is_some()
        || hyperlink.history.is_some()
        || nonempty(hyperlink.doc_location.as_deref()).is_some();
    if !has_attributes && nonempty(hyperlink.href.as_deref()).is_none() {
        return children;
    }
    let mut writer = XmlWriter::with_capacity(children.len() + 128);
    writer.start_element("w:hyperlink");
    optional_attr(&mut writer, "r:id", hyperlink.relationship_id.as_deref());
    optional_attr(&mut writer, "w:anchor", hyperlink.anchor.as_deref());
    optional_attr(&mut writer, "w:tooltip", hyperlink.tooltip.as_deref());
    optional_attr(&mut writer, "w:tgtFrame", hyperlink.target.as_deref());
    if let Some(history) = hyperlink.history {
        writer.attribute("w:history", if history { "1" } else { "0" });
    }
    optional_attr(
        &mut writer,
        "w:docLocation",
        hyperlink.doc_location.as_deref(),
    );
    append_generated(&mut writer, &children);
    writer.end_element();
    writer.finish()
}

fn serialize_simple_field(
    field: &SimpleField,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut shared = None;
    for run in &field.content {
        merge_run_direction_scopes(run, &mut shared);
    }
    let (mut output, base) = context.direction.enter_wrapper(shared);
    output.push_str("<w:fldSimple w:instr=\"");
    output.push_str(&super::xml_writer::escape_xml(&field.instruction));
    output.push('"');
    if field.fld_lock == Some(true) {
        output.push_str(" w:fldLock=\"true\"");
    }
    if field.dirty == Some(true) {
        output.push_str(" w:dirty=\"true\"");
    }
    output.push('>');
    for run in &field.content {
        output.push_str(&serialize_context_run(run, context)?);
    }
    output.push_str(&context.direction.leave_wrapper(base));
    output.push_str("</w:fldSimple>");
    Ok(output)
}

enum FieldResultItem<'a> {
    Run(&'a Run),
    Bookmark(&'a InlineNode),
}

fn field_result_children<'a>(
    nodes: &'a [InlineNode],
    runs: Vec<&'a Run>,
    projected: bool,
    items: &mut Vec<FieldResultItem<'a>>,
    aligned: &mut bool,
) {
    *aligned &= nodes
        .iter()
        .filter_map(|node| match node {
            InlineNode::Run(run) => Some(run),
            _ => None,
        })
        .eq(runs);
    for node in nodes {
        match node {
            InlineNode::Run(run) if projected => {
                items.push(FieldResultItem::Run(run));
            }
            _ => field_result_items(std::slice::from_ref(node), false, items, aligned),
        }
    }
}

fn field_result_items<'a>(
    nodes: &'a [InlineNode],
    projected: bool,
    items: &mut Vec<FieldResultItem<'a>>,
    aligned: &mut bool,
) {
    for node in nodes {
        match node {
            InlineNode::Run(run) => {
                if projected {
                    items.push(FieldResultItem::Run(run));
                }
            }
            InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_) => {
                items.push(FieldResultItem::Bookmark(node));
            }
            InlineNode::Hyperlink(link) => field_result_children(
                link.structured_children
                    .as_deref()
                    .unwrap_or(&link.children),
                field_runs(node),
                projected,
                items,
                aligned,
            ),
            InlineNode::SimpleField(field) => match field
                .structured_result
                .as_ref()
                .and_then(|result| result.inline.as_deref())
            {
                Some(nodes) => {
                    field_result_children(nodes, field_runs(node), projected, items, aligned);
                }
                None if projected => {
                    items.extend(field_runs(node).into_iter().map(FieldResultItem::Run));
                }
                None => {}
            },
            InlineNode::ComplexField(field) => {
                for (content, runs) in [
                    (&field.structured_code, &field.field_code),
                    (&field.structured_result, &field.field_result),
                ] {
                    if let Some(nodes) = content
                        .as_ref()
                        .and_then(|content| content.inline.as_deref())
                    {
                        *aligned &= nodes.iter().flat_map(field_runs).eq(runs);
                        field_result_items(nodes, false, items, aligned);
                    }
                }
            }
            InlineNode::InlineSdt(sdt) => {
                field_result_items(&sdt.content, false, items, aligned);
            }
            InlineNode::Tracked(change) => {
                field_result_items(&change.content, false, items, aligned);
            }
            InlineNode::Math(_) | InlineNode::RawXml(_) => {}
        }
    }
}

fn field_run_length(run: &Run) -> usize {
    run.content
        .iter()
        .map(|content| match content {
            RunContent::Text { text, .. } => text.encode_utf16().count(),
            RunContent::CommentReference { .. }
            | RunContent::FieldChar { .. }
            | RunContent::InstrText { .. } => 0,
            _ => 1,
        })
        .sum()
}

fn field_result_bookmarks(field: &ComplexField) -> Vec<(usize, &InlineNode)> {
    let Some(nodes) = field
        .structured_result
        .as_ref()
        .and_then(|content| content.inline.as_ref())
    else {
        return Vec::new();
    };
    let mut items = Vec::new();
    let mut aligned = true;
    field_result_items(nodes, true, &mut items, &mut aligned);
    let in_place = aligned
        && items
            .iter()
            .filter_map(|item| match item {
                FieldResultItem::Run(run) => Some(*run),
                FieldResultItem::Bookmark(_) => None,
            })
            .eq(field.field_result.iter());
    let mut bookmarks = Vec::new();
    let mut offset = 0usize;
    for item in items {
        match item {
            FieldResultItem::Bookmark(node) => bookmarks.push((offset, node)),
            FieldResultItem::Run(run) => {
                offset = offset.saturating_add(if in_place { 1 } else { field_run_length(run) });
            }
        }
    }
    if in_place || bookmarks.is_empty() {
        return bookmarks;
    }
    let mut boundaries = Vec::with_capacity(field.field_result.len() + 1);
    let mut total = 0usize;
    boundaries.push(total);
    for run in &field.field_result {
        total = total.saturating_add(field_run_length(run));
        boundaries.push(total);
    }
    for (offset, node) in &mut bookmarks {
        if *offset == 0 {
            continue;
        }
        if *offset >= total {
            *offset = field.field_result.len();
            continue;
        }
        let after = boundaries.partition_point(|boundary| *boundary < *offset);
        *offset = if boundaries[after] == *offset || matches!(node, InlineNode::BookmarkEnd(_)) {
            after
        } else {
            after - 1
        };
    }
    bookmarks.sort_by_key(|(boundary, _)| *boundary);
    bookmarks
}

fn serialize_complex_field(
    field: &ComplexField,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let mut shared = None;
    for run in field.field_code.iter().chain(&field.field_result) {
        merge_run_direction_scopes(run, &mut shared);
    }
    if let Some(source) = field.source_run_xml.as_deref()
        && !context.deletion
        && !context.in_revision
        && field.source_run_digest.as_deref() == Some(field_run_digest(field).as_str())
    {
        validate_replayed_fragment(source)?;
        let mut output = context.direction.transition(shared.unwrap_or_default());
        output.push_str(source);
        return Ok(output);
    }
    let formatting = field
        .field_result
        .first()
        .and_then(|run| run.formatting.as_ref())
        .or(field.formatting.as_ref());
    let properties = serialize_text_formatting(formatting);
    let mut output = context.direction.transition(shared.unwrap_or_default());
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"begin\"");
    if field.fld_lock == Some(true) {
        output.push_str(" w:fldLock=\"true\"");
    }
    if field.dirty == Some(true) {
        output.push_str(" w:dirty=\"true\"");
    }
    if let Some(raw) = field.raw_form_data_xml.as_deref() {
        output.push('>');
        match validate_form_field_data(raw) {
            Ok(()) => {
                context.note_form_data(raw);
                output.push_str(raw);
            }
            Err(error) => {
                let name = form_field_name(field, raw);
                let detail = error.to_string();
                let Some(token) = context.defer_form_data(&name, &detail) else {
                    return Err(form_data_error(&name, &detail));
                };
                output.push_str(&token);
            }
        }
        output.push_str("</w:fldChar></w:r>");
    } else {
        output.push_str("/></w:r>");
    }
    let code_nodes = field
        .structured_code
        .as_ref()
        .filter(|content| !content.spans_paragraphs())
        .and_then(|content| content.inline.as_deref())
        .filter(|nodes| nodes.iter().any(|node| !matches!(node, InlineNode::Run(_))));
    let in_place = code_nodes.filter(|nodes| {
        !field.field_code.is_empty()
            && nodes
                .iter()
                .flat_map(field_runs)
                .eq(field.field_code.iter())
    });
    let mut result_scopes = None;
    for run in &field.field_result {
        merge_run_direction_scopes(run, &mut result_scopes);
    }
    let result_scopes = result_scopes.unwrap_or_default();
    if let Some(nodes) = in_place {
        let mut upcoming = Upcoming::new(nodes, Some(result_scopes), Within::of(context));
        for (index, node) in nodes.iter().enumerate() {
            upcoming.look_ahead(index, context);
            output.push_str(&serialize_inline_node(node, context)?);
        }
    } else if let Some(nodes) = code_nodes {
        let code_scopes = match field.field_code.first() {
            Some(run) => run.direction_scopes.as_slice(),
            None => shared.unwrap_or_default(),
        };
        for node in nodes.iter().filter(|node| {
            matches!(
                node,
                InlineNode::BookmarkStart(_) | InlineNode::BookmarkEnd(_)
            )
        }) {
            context.direction.look_ahead(Some(code_scopes));
            output.push_str(&serialize_inline_node(node, context)?);
        }
    }
    if in_place.is_none() && field.field_code.is_empty() {
        output.push_str("<w:r>");
        output.push_str(&properties);
        output.push_str("<w:instrText");
        if needs_preserve(&field.instruction) {
            output.push_str(" xml:space=\"preserve\"");
        }
        output.push('>');
        output.push_str(&super::xml_writer::escape_xml(&field.instruction));
        output.push_str("</w:instrText></w:r>");
    } else if in_place.is_none() {
        for run in &field.field_code {
            output.push_str(&serialize_context_run(run, context)?);
        }
    }
    output.push_str(&context.direction.transition(result_scopes));
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"separate\"/></w:r>");
    // Run-level fallback: multi-block results and fields rebuilt by the edit path.
    let structured = field
        .structured_result
        .as_ref()
        .filter(|content| !content.spans_paragraphs())
        .and_then(|content| content.inline.as_ref());
    if let Some(nodes) = structured {
        let mut upcoming = Upcoming::new(nodes, Some(result_scopes), Within::of(context));
        for (index, node) in nodes.iter().enumerate() {
            upcoming.look_ahead(index, context);
            output.push_str(&serialize_inline_node(node, context)?);
        }
    } else {
        let mut bookmarks = field_result_bookmarks(field).into_iter().peekable();
        for (index, run) in field.field_result.iter().enumerate() {
            while bookmarks
                .peek()
                .is_some_and(|(boundary, _)| *boundary == index)
            {
                let (_, node) = bookmarks.next().unwrap();
                context
                    .direction
                    .look_ahead(Some(run.direction_scopes.as_slice()));
                output.push_str(&serialize_inline_node(node, context)?);
            }
            output.push_str(&serialize_context_run(run, context)?);
        }
        for (_, node) in bookmarks {
            context.direction.look_ahead(Some(result_scopes));
            output.push_str(&serialize_inline_node(node, context)?);
        }
    }
    output.push_str(&context.direction.transition(result_scopes));
    output.push_str("<w:r>");
    output.push_str(&properties);
    output.push_str("<w:fldChar w:fldCharType=\"end\"/></w:r>");
    Ok(output)
}

pub fn serialize_inline_sdt(
    sdt: &InlineSdt,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let (properties, end_properties) = serialize_sdt_properties(&sdt.properties)?;
    let mut writer = XmlWriter::with_capacity(256);
    writer.start_element("w:sdt");
    append_generated(&mut writer, &properties);
    append_generated(&mut writer, &end_properties);
    writer.start_element("w:sdtContent");
    let previous = std::mem::replace(&mut context.in_control, true);
    let within = Within::of(context);
    let (before, base) = context
        .direction
        .enter_wrapper(shared_direction_scopes(&sdt.content, within));
    let mut upcoming = Upcoming::new(&sdt.content, None, within);
    let result = sdt
        .content
        .iter()
        .enumerate()
        .try_for_each(|(index, item)| {
            upcoming.look_ahead(index, context);
            append_generated(&mut writer, &serialize_inline_node(item, context)?);
            Ok::<_, ParseError>(())
        });
    context.in_control = previous;
    result?;
    let closing = context.direction.leave_wrapper(base);
    if !closing.is_empty() {
        append_generated(&mut writer, &closing);
    }
    writer.end_element().end_element();
    Ok(prefixed(before, writer.finish()))
}

fn form_field_name(field: &ComplexField, raw: &str) -> String {
    raw.split_once(":name ")
        .and_then(|(_, rest)| rest.split_once(":val=\""))
        .and_then(|(_, rest)| rest.split_once('"'))
        .map(|(name, _)| name)
        .filter(|name| !name.is_empty())
        .map_or_else(|| field.instruction.clone(), str::to_owned)
}

pub(crate) fn form_data_error(field: &str, detail: &str) -> ParseError {
    ParseError::Canonical(format!(
        "the form field data of field {field:?} cannot be written: {detail}"
    ))
}

/// Fails when `xml` still holds a token written in place of form field data.
pub(crate) fn reject_form_data_tokens(
    xml: &str,
    deferred: &DeferredFormData,
) -> Result<(), ParseError> {
    if deferred.tokens.is_empty() {
        return Ok(());
    }
    let outside = deferred
        .in_written_data
        .then(|| comments_outside_form_data(xml))
        .flatten();
    let left = deferred.tokens.iter().find(|(token, _, _)| match &outside {
        Some(comments) => comments.contains(token.as_str()),
        None => xml.contains(token.as_str()),
    });
    match left {
        Some((_, field, detail)) => Err(form_data_error(field, detail)),
        None => Ok(()),
    }
}

/// Every comment of `xml` outside all `w:ffData`, where tokens sit and written data never does.
fn comments_outside_form_data(xml: &str) -> Option<std::collections::HashSet<String>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut comments = std::collections::HashSet::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event().ok()? {
            Event::Start(element) if element.name().as_ref() == b"w:ffData" => depth += 1,
            Event::End(element) if depth > 0 && element.name().as_ref() == b"w:ffData" => {
                depth -= 1;
            }
            Event::Comment(text) if depth == 0 => {
                comments.insert(format!("<!--{}-->", std::str::from_utf8(&text).ok()?));
            }
            Event::Eof => return Some(comments),
            _ => {}
        }
    }
}

pub(crate) fn serialize_sdt_properties(
    properties: &SdtProperties,
) -> Result<(String, String), ParseError> {
    let properties_xml = if let Some(raw) = properties.raw_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtPr")?;
        raw.to_owned()
    } else {
        synthesize_sdt_properties(properties)
    };
    let end_properties = if let Some(raw) = properties.raw_end_properties_xml.as_deref() {
        validate_raw_subtree(raw, "w", "sdtEndPr")?;
        raw.to_owned()
    } else {
        String::new()
    };
    Ok((properties_xml, end_properties))
}

pub fn synthesize_sdt_properties(properties: &SdtProperties) -> String {
    let mut body = XmlWriter::with_capacity(256);
    if let Some(value) = nonempty(properties.alias.as_deref()) {
        empty_attr(&mut body, "w:alias", "w:val", value);
    }
    if let Some(value) = nonempty(properties.tag.as_deref()) {
        empty_attr(&mut body, "w:tag", "w:val", value);
    }
    if let Some(value) = properties.id {
        empty_attr(&mut body, "w:id", "w:val", &js_number(value));
    }
    if let Some(value) = nonempty(properties.lock.as_deref()).filter(|value| *value != "unlocked") {
        empty_attr(&mut body, "w:lock", "w:val", value);
    }
    if let Some(value) = nonempty(properties.placeholder.as_deref()) {
        body.start_element("w:placeholder");
        empty_attr(&mut body, "w:docPart", "w:val", value);
        body.end_element();
    }
    if properties.showing_placeholder == Some(true) {
        body.start_element("w:showingPlcHdr").end_element();
    }
    match properties.sdt_type.as_str() {
        "plainText" => {
            body.start_element("w:text");
            if properties.multi_line == Some(true) {
                body.attribute("w:multiLine", "1");
            }
            body.end_element();
        }
        "date" => {
            body.start_element("w:date");
            if let Some(value) = nonempty(properties.date_format.as_deref()) {
                empty_attr(&mut body, "w:dateFormat", "w:val", value);
            }
            body.end_element();
        }
        "dropDownList" | "comboBox" => {
            let name = if properties.sdt_type == "dropDownList" {
                "w:dropDownList"
            } else {
                "w:comboBox"
            };
            body.start_element(name).attribute("w:lastValue", "");
            for item in properties.list_items.as_deref().unwrap_or_default() {
                body.start_element("w:listItem")
                    .attribute("w:displayText", &item.display_text)
                    .attribute("w:value", &item.value)
                    .end_element();
            }
            body.end_element();
        }
        "checkbox" => {
            body.start_element("w14:checkbox");
            body.start_element("w14:checked")
                .attribute(
                    "w14:val",
                    if properties.checked == Some(true) {
                        "1"
                    } else {
                        "0"
                    },
                )
                .end_element()
                .start_element("w14:checkedState")
                .attribute("w14:val", "2612")
                .attribute("w14:font", "MS Gothic")
                .end_element()
                .start_element("w14:uncheckedState")
                .attribute("w14:val", "2610")
                .attribute("w14:font", "MS Gothic")
                .end_element()
                .end_element();
        }
        "picture" => {
            body.start_element("w:picture").end_element();
        }
        _ => {}
    }
    format!("<w:sdtPr>{}</w:sdtPr>", body.finish())
}

fn serialize_tracked_change(
    change: &TrackedInline,
    context: &mut SerializerContext,
) -> Result<String, ParseError> {
    let (element, deletion) = match change.node_type.as_str() {
        "insertion" => ("w:ins", false),
        "deletion" => ("w:del", true),
        "moveFrom" => ("w:moveFrom", true),
        "moveTo" => ("w:moveTo", false),
        _ => return Ok(String::new()),
    };
    let within = Within {
        revision: true,
        ..Within::of(context)
    };
    let (before, base) = context
        .direction
        .enter_wrapper(shared_direction_scopes(&change.content, within));
    let mut writer = XmlWriter::with_capacity(256);
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(change.info.id))
        .attribute(
            "w:author",
            nonempty_trimmed(&change.info.author).unwrap_or("Unknown"),
        );
    if let Some(date) = change.info.date.as_deref().and_then(nonempty_trimmed) {
        writer.attribute("w:date", date);
    }
    // Tracked wrappers use explicit start and end tags.
    writer.text("");
    let previous_revision = std::mem::replace(&mut context.in_revision, true);
    let previous_deletion =
        std::mem::replace(&mut context.deletion, deletion && context.in_control);
    let mut upcoming = Upcoming::new(&change.content, None, within);
    let result = change
        .content
        .iter()
        .enumerate()
        .try_for_each(|(index, item)| {
            let xml = match item {
                _ if context.in_control => {
                    upcoming.look_ahead(index, context);
                    serialize_inline_node(item, context)?
                }
                InlineNode::Run(run) => serialize_scoped_run(run, deletion, context)?,
                InlineNode::Hyperlink(link) => serialize_hyperlink(link, context)?,
                InlineNode::InlineSdt(_) => {
                    context.deletion = deletion;
                    let result = serialize_inline_node(item, context);
                    context.deletion = false;
                    result?
                }
                InlineNode::Tracked(_) if item.has_tracked_control(false, true) => {
                    serialize_inline_node(item, context)?
                }
                _ => return Ok(()),
            };
            append_generated(&mut writer, &xml);
            Ok::<_, ParseError>(())
        });
    context.in_revision = previous_revision;
    context.deletion = previous_deletion;
    result?;
    let closing = context.direction.leave_wrapper(base);
    if !closing.is_empty() {
        append_generated(&mut writer, &closing);
    }
    writer.end_element();
    Ok(prefixed(before, writer.finish()))
}

fn serialize_range_start(marker: &RangeStart) -> Result<String, ParseError> {
    let element = match marker.node_type.as_str() {
        "moveFromRangeStart" => "w:moveFromRangeStart",
        "moveToRangeStart" => "w:moveToRangeStart",
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(80);
    writer
        .start_element(element)
        .attribute("w:id", &js_number(marker.id))
        .attribute("w:name", &marker.name)
        .end_element();
    Ok(writer.finish())
}

fn serialize_range_end(marker: &RangeEnd) -> Result<String, ParseError> {
    let element = match marker.node_type.as_str() {
        "moveFromRangeEnd" => "w:moveFromRangeEnd",
        "moveToRangeEnd" => "w:moveToRangeEnd",
        _ => return Ok(String::new()),
    };
    let mut writer = XmlWriter::with_capacity(48);
    writer
        .start_element(element)
        .attribute("w:id", &js_number(marker.id))
        .end_element();
    Ok(writer.finish())
}

fn run_has_comment_reference(run: &Run, id: f64) -> bool {
    run.content.iter().any(
        |content| matches!(content, RunContent::CommentReference { id: reference } if *reference == Some(id)),
    )
}

fn has_comment_reference(node: &InlineNode, id: f64) -> bool {
    match node {
        InlineNode::Run(run) => run_has_comment_reference(run, id),
        InlineNode::Hyperlink(hyperlink) => hyperlink
            .children
            .iter()
            .any(|node| matches!(node, InlineNode::Run(run) if run_has_comment_reference(run, id))),
        InlineNode::InlineSdt(sdt) => sdt
            .content
            .iter()
            .any(|node| has_comment_reference(node, id)),
        InlineNode::Tracked(change) => change
            .content
            .iter()
            .any(|node| has_comment_reference(node, id)),
        InlineNode::SimpleField(field) => field
            .content
            .iter()
            .any(|run| run_has_comment_reference(run, id)),
        InlineNode::ComplexField(field) => {
            field
                .field_code
                .iter()
                .any(|run| run_has_comment_reference(run, id))
                || field
                    .structured_result
                    .as_ref()
                    .filter(|content| !content.spans_paragraphs())
                    .and_then(|content| content.inline.as_ref())
                    .map_or_else(
                        || {
                            field
                                .field_result
                                .iter()
                                .any(|run| run_has_comment_reference(run, id))
                        },
                        |nodes| nodes.iter().any(|node| has_comment_reference(node, id)),
                    )
        }
        _ => false,
    }
}

fn serialize_comment_range(marker: &CommentRange) -> Result<String, ParseError> {
    let mut writer = XmlWriter::with_capacity(180);
    match marker.node_type.as_str() {
        "commentRangeStart" => {
            writer
                .start_element("w:commentRangeStart")
                .attribute("w:id", &js_number(marker.id))
                .end_element();
        }
        "commentRangeEnd" => {
            writer
                .start_element("w:commentRangeEnd")
                .attribute("w:id", &js_number(marker.id))
                .end_element();
        }
        _ => {}
    }
    Ok(writer.finish())
}

fn serialize_math(math: &MathEquation) -> Result<String, ParseError> {
    if math.omml_xml.is_empty() {
        return Ok(String::new());
    }
    validate_math_subtree(&math.omml_xml)?;
    Ok(math.omml_xml.clone())
}

fn write_paragraph_mark_properties(
    writer: &mut XmlWriter,
    formatting: Option<&ParagraphFormatting>,
    insertion: Option<&TrackedChangeInfo>,
    deletion: Option<&TrackedChangeInfo>,
) {
    let formatting = formatting
        .and_then(|formatting| formatting.run_properties.as_ref())
        .map(|formatting| serialize_text_formatting(Some(formatting)))
        .unwrap_or_default();
    let inner = formatting
        .strip_prefix("<w:rPr>")
        .and_then(|value| value.strip_suffix("</w:rPr>"))
        .unwrap_or_default();
    if insertion.is_none() && deletion.is_none() && inner.is_empty() {
        return;
    }
    writer.start_element("w:rPr");
    if let Some(info) = insertion {
        write_mark_change(writer, "w:ins", info);
    }
    if let Some(info) = deletion {
        write_mark_change(writer, "w:del", info);
    }
    append_generated(writer, inner);
    writer.end_element();
}

fn write_mark_change(writer: &mut XmlWriter, element: &'static str, info: &TrackedChangeInfo) {
    writer
        .start_element(element)
        .attribute("w:id", &normalized_tracked_id(info.id))
        .attribute(
            "w:author",
            if info.author.is_empty() {
                "Unknown"
            } else {
                &info.author
            },
        );
    if let Some(date) = nonempty(info.date.as_deref()) {
        writer.attribute("w:date", date);
    }
    writer.end_element();
}

fn write_paragraph_borders(writer: &mut XmlWriter, borders: Option<&Borders>) {
    let Some(borders) = borders else {
        return;
    };
    let entries = [
        (borders.top.as_ref(), BorderSide::Top),
        (borders.left.as_ref(), BorderSide::Left),
        (borders.bottom.as_ref(), BorderSide::Bottom),
        (borders.right.as_ref(), BorderSide::Right),
        (borders.between.as_ref(), BorderSide::Between),
        (borders.bar.as_ref(), BorderSide::Bar),
    ];
    if !entries.iter().any(|(border, _)| border.is_some()) {
        return;
    }
    writer.start_element("w:pBdr");
    for (border, side) in entries {
        if let Some(border) = border {
            write_border(writer, border, side);
        }
    }
    writer.end_element();
}

fn write_tabs(writer: &mut XmlWriter, tabs: Option<&[crate::tabs::TabStop]>) {
    let Some(tabs) = tabs.filter(|tabs| !tabs.is_empty()) else {
        return;
    };
    writer.start_element("w:tabs");
    for tab in tabs {
        writer
            .start_element("w:tab")
            .attribute("w:val", &tab.alignment)
            .attribute("w:pos", &int_attr(Some(tab.position)));
        if let Some(leader) = nonempty(tab.leader.as_deref()).filter(|value| *value != "none") {
            writer.attribute("w:leader", leader);
        }
        writer.end_element();
    }
    writer.end_element();
}

fn write_spacing(writer: &mut XmlWriter, formatting: &ParagraphFormatting) {
    if formatting.space_before.is_none()
        && formatting.space_after.is_none()
        && formatting.space_before_lines.is_none()
        && formatting.space_after_lines.is_none()
        && formatting.line_spacing.is_none()
        && nonempty(formatting.line_spacing_rule.as_deref()).is_none()
        && formatting.before_autospacing.is_none()
        && formatting.after_autospacing.is_none()
    {
        return;
    }
    writer.start_element("w:spacing");
    optional_int(writer, "w:before", formatting.space_before);
    optional_int(writer, "w:after", formatting.space_after);
    optional_int(writer, "w:beforeLines", formatting.space_before_lines);
    optional_int(writer, "w:afterLines", formatting.space_after_lines);
    optional_int(writer, "w:line", formatting.line_spacing);
    optional_attr(
        writer,
        "w:lineRule",
        formatting.line_spacing_rule.as_deref(),
    );
    if let Some(auto) = formatting.before_autospacing {
        writer.attribute("w:beforeAutospacing", if auto { "1" } else { "0" });
    }
    if let Some(auto) = formatting.after_autospacing {
        writer.attribute("w:afterAutospacing", if auto { "1" } else { "0" });
    }
    writer.end_element();
}

fn write_indentation(writer: &mut XmlWriter, formatting: &ParagraphFormatting) {
    let hanging = formatting.hanging_indent == Some(true);
    if formatting.indent_left.is_none()
        && formatting.indent_right.is_none()
        && formatting.indent_first_line.is_none()
        && formatting.indent_left_chars.is_none()
        && formatting.indent_right_chars.is_none()
        && formatting.indent_first_line_chars.is_none()
    {
        return;
    }
    writer.start_element("w:ind");
    optional_int(writer, "w:left", formatting.indent_left);
    optional_int(writer, "w:leftChars", formatting.indent_left_chars);
    optional_int(writer, "w:right", formatting.indent_right);
    optional_int(writer, "w:rightChars", formatting.indent_right_chars);
    if let Some(first) = formatting.indent_first_line {
        let name = if hanging { "w:hanging" } else { "w:firstLine" };
        writer.attribute(
            name,
            &int_attr(Some(if hanging { first.abs() } else { first })),
        );
    }
    if let Some(first) = formatting.indent_first_line_chars {
        let hanging = formatting
            .hanging_indent_chars
            .unwrap_or_else(|| first.is_sign_negative());
        let name = if hanging {
            "w:hangingChars"
        } else {
            "w:firstLineChars"
        };
        writer.attribute(name, &int_attr(Some(first.abs())));
    }
    writer.end_element();
}

fn write_numbering(writer: &mut XmlWriter, numbering: Option<&NumberingProperties>) {
    let Some(numbering) = numbering else {
        return;
    };
    if numbering.ilvl.is_none() && numbering.num_id.is_none() {
        return;
    }
    writer.start_element("w:numPr");
    if let Some(value) = numbering.ilvl {
        empty_attr(writer, "w:ilvl", "w:val", &int_attr(Some(value)));
    }
    if let Some(value) = numbering.num_id {
        empty_attr(writer, "w:numId", "w:val", &int_attr(Some(value)));
    }
    writer.end_element();
}

fn write_frame(writer: &mut XmlWriter, frame: Option<&ParagraphFrame>) {
    let Some(frame) = frame else {
        return;
    };
    let has_attributes = frame.width.is_some()
        || frame.height.is_some()
        || nonempty(frame.h_anchor.as_deref()).is_some()
        || nonempty(frame.v_anchor.as_deref()).is_some()
        || frame.x.is_some()
        || frame.y.is_some()
        || nonempty(frame.x_align.as_deref()).is_some()
        || nonempty(frame.y_align.as_deref()).is_some()
        || nonempty(frame.wrap.as_deref()).is_some();
    if !has_attributes {
        return;
    }
    writer.start_element("w:framePr");
    optional_int(writer, "w:w", frame.width);
    optional_int(writer, "w:h", frame.height);
    optional_attr(writer, "w:hAnchor", frame.h_anchor.as_deref());
    optional_attr(writer, "w:vAnchor", frame.v_anchor.as_deref());
    optional_int(writer, "w:x", frame.x);
    optional_int(writer, "w:y", frame.y);
    optional_attr(writer, "w:xAlign", frame.x_align.as_deref());
    optional_attr(writer, "w:yAlign", frame.y_align.as_deref());
    optional_attr(writer, "w:wrap", frame.wrap.as_deref());
    writer.end_element();
}

fn needs_preserve(value: &str) -> bool {
    value.starts_with(' ') || value.ends_with(' ') || value.contains("  ")
}

fn optional_attr(writer: &mut XmlWriter, name: &'static str, value: Option<&str>) {
    if let Some(value) = nonempty(value) {
        writer.attribute(name, value);
    }
}

fn optional_int(writer: &mut XmlWriter, name: &'static str, value: Option<f64>) {
    if value.is_some() {
        writer.attribute(name, &int_attr(value));
    }
}

pub(crate) fn empty_attr(
    writer: &mut XmlWriter,
    element: &'static str,
    name: &'static str,
    value: &str,
) {
    writer
        .start_element(element)
        .attribute(name, value)
        .end_element();
}

fn write_paragraph_cnf_style(writer: &mut XmlWriter, style: Option<&ParagraphCnfStyle>) {
    let Some(style) = style else {
        return;
    };
    writer.start_element("w:cnfStyle");
    for (name, value) in [
        ("w:val", &style.val),
        ("w:firstRow", &style.first_row),
        ("w:lastRow", &style.last_row),
        ("w:firstColumn", &style.first_column),
        ("w:lastColumn", &style.last_column),
        ("w:oddVBand", &style.odd_v_band),
        ("w:evenVBand", &style.even_v_band),
        ("w:oddHBand", &style.odd_h_band),
        ("w:evenHBand", &style.even_h_band),
        ("w:firstRowFirstColumn", &style.first_row_first_column),
        ("w:firstRowLastColumn", &style.first_row_last_column),
        ("w:lastRowFirstColumn", &style.last_row_first_column),
        ("w:lastRowLastColumn", &style.last_row_last_column),
    ] {
        if let Some(value) = value {
            writer.attribute(name, value);
        }
    }
    writer.end_element();
}

pub(super) fn on_off(writer: &mut XmlWriter, name: &'static str, value: Option<bool>) {
    match value {
        Some(true) => {
            writer.start_element(name).end_element();
        }
        Some(false) => {
            writer
                .start_element(name)
                .attribute("w:val", "0")
                .end_element();
        }
        None => {}
    }
}

#[cfg(test)]
mod tests {
    use serde_json::{Value, json};

    use crate::inline::{Run, RunContent, RunType};
    use crate::serializer::s10::SerializerDeterminism;

    use super::*;

    fn context() -> SerializerContext {
        SerializerContext::new(&SerializerDeterminism {
            seed: "0".repeat(64),
            now: "2000-01-01T00:00:00.000Z".to_owned(),
        })
        .unwrap()
    }

    #[test]
    fn a_simple_field_saves_back_as_a_simple_field() {
        let field = SimpleField {
            node_type: crate::inline::SimpleFieldType::SimpleField,
            field_type: "PAGE".to_owned(),
            instruction: " PAGE ".to_owned(),
            content: vec![Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: vec![RunContent::Text {
                    text: "3".to_owned(),
                    preserve_space: None,
                }],
                direction_scopes: Vec::new(),
            }],
            fld_lock: Some(true),
            dirty: Some(true),
            structured_result: None,
            field_tree: None,
        };
        let output = serialize_simple_field(&field, &mut context()).unwrap();
        assert_eq!(
            output,
            "<w:fldSimple w:instr=\" PAGE \" w:fldLock=\"true\" w:dirty=\"true\">\
             <w:r><w:t>3</w:t></w:r></w:fldSimple>"
        );
    }

    #[test]
    fn paragraph_bytes_pin_ids_properties_and_rendered_break_injection() {
        let paragraph = Paragraph {
            node_type: "paragraph".to_owned(),
            para_id: Some("AA&BB\"CC".to_owned()),
            repeated_para_id: None,
            para_id_attribute: None,
            source_ordinal: None,
            text_id: None,
            extra_attributes: Vec::new(),
            formatting: Some(ParagraphFormatting {
                keep_next: Some(false),
                alignment: Some("center".to_owned()),
                ..ParagraphFormatting::default()
            }),
            property_changes: None,
            p_pr_ins: None,
            p_pr_del: None,
            content: vec![ParagraphContent::Inline(InlineNode::Run(Run {
                node_type: RunType::Run,
                formatting: None,
                property_changes: None,
                content: vec![RunContent::Text {
                    text: "hello & goodbye".to_owned(),
                    preserve_space: None,
                }],
                direction_scopes: Vec::new(),
            }))],
            list_rendering: None,
            rendered_page_break_before: Some(true),
            section_properties: None,
        };
        assert_eq!(
            serialize_paragraph(&paragraph, &mut context()).unwrap(),
            "<w:p w14:paraId=\"AA&amp;BB&quot;CC\"><w:pPr><w:keepNext w:val=\"0\"/><w:jc w:val=\"center\"/></w:pPr><w:r><w:lastRenderedPageBreak/><w:t>hello &amp; goodbye</w:t></w:r></w:p>"
        );
    }

    /// A paragraph-level sectPr ends a section even when it holds nothing, so
    /// it always saves; the body-level one and other paragraphs are unchanged.
    #[test]
    fn paragraph_section_marks_save_even_when_empty() {
        let save = |section_properties: Option<SectionProperties>| {
            let paragraph = Paragraph {
                node_type: "paragraph".to_owned(),
                para_id: Some("00000001".to_owned()),
                repeated_para_id: None,
                para_id_attribute: None,
                source_ordinal: None,
                text_id: None,
                extra_attributes: Vec::new(),
                formatting: None,
                property_changes: None,
                p_pr_ins: None,
                p_pr_del: None,
                content: vec![ParagraphContent::Inline(InlineNode::Run(Run {
                    node_type: RunType::Run,
                    formatting: None,
                    property_changes: None,
                    content: vec![RunContent::Text {
                        text: "A".to_owned(),
                        preserve_space: None,
                    }],
                    direction_scopes: Vec::new(),
                }))],
                list_rendering: None,
                rendered_page_break_before: None,
                section_properties,
            };
            serialize_paragraph(&paragraph, &mut context()).unwrap()
        };
        assert_eq!(
            save(None),
            "<w:p w14:paraId=\"00000001\"><w:r><w:t>A</w:t></w:r></w:p>"
        );
        assert_eq!(
            save(Some(SectionProperties {
                section_start: Some("continuous".to_owned()),
                ..SectionProperties::default()
            })),
            "<w:p w14:paraId=\"00000001\"><w:pPr><w:sectPr><w:type w:val=\"continuous\"/></w:sectPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>"
        );
        assert_eq!(
            save(Some(SectionProperties::default())),
            "<w:p w14:paraId=\"00000001\"><w:pPr><w:sectPr/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>"
        );
        assert_eq!(
            serialize_section_properties(Some(&SectionProperties::default())),
            ""
        );
    }

    /// A hanging character indent of zero has no sign to carry its kind, so
    /// the flag has to.
    #[test]
    fn a_zero_hanging_character_indent_stays_hanging() {
        let mut writer = XmlWriter::with_capacity(64);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line_chars: Some(0.0),
                hanging_indent_chars: Some(true),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(writer.finish(), "<w:ind w:hangingChars=\"0\"/>");
    }

    #[test]
    fn indentation_keeps_character_units_and_an_explicit_zero_first_line() {
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_left: Some(200.0),
                indent_left_chars: Some(100.0),
                indent_first_line_chars: Some(200.0),
                indent_first_line: Some(420.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:left=\"200\" w:leftChars=\"100\" w:firstLine=\"420\" w:firstLineChars=\"200\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line: Some(0.0),
                indent_first_line_chars: Some(0.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:firstLine=\"0\" w:firstLineChars=\"0\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                hanging_indent: Some(true),
                indent_first_line: Some(-315.0),
                indent_first_line_chars: Some(-150.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:hanging=\"315\" w:hangingChars=\"150\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line: Some(420.0),
                indent_first_line_chars: Some(-200.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(
            writer.finish(),
            "<w:ind w:firstLine=\"420\" w:hangingChars=\"200\"/>"
        );
        let mut writer = XmlWriter::with_capacity(128);
        write_indentation(
            &mut writer,
            &ParagraphFormatting {
                indent_first_line_chars: Some(-0.0),
                ..ParagraphFormatting::default()
            },
        );
        assert_eq!(writer.finish(), "<w:ind w:hangingChars=\"0\"/>");
    }

    #[test]
    fn empty_tracked_wrappers_use_explicit_end_tags() {
        let change = TrackedInline {
            node_type: "deletion".to_owned(),
            info: TrackedChangeInfo {
                id: 7.0,
                author: "Ada".to_owned(),
                date: None,
            },
            content: Vec::new(),
        };
        assert_eq!(
            serialize_tracked_change(&change, &mut context()).unwrap(),
            "<w:del w:id=\"7\" w:author=\"Ada\"></w:del>"
        );
    }

    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const DIR: (&str, &str) = (r#"<w:dir w:val="rtl">"#, "</w:dir>");

    const BDO: (&str, &str) = (r#"<w:bdo w:val="rtl">"#, "</w:bdo>");

    const SCOPE_TAGS: [&str; 4] = [
        r#"<w:dir w:val="rtl">"#,
        r#"<w:dir w:val="ltr">"#,
        r#"<w:bdo w:val="rtl">"#,
        r#"<w:bdo w:val="ltr">"#,
    ];

    const FIELD_OPEN: &str = concat!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText xml:space="preserve"> SUBJECT </w:instrText></w:r>"#,
        r#"<w:dir w:val="rtl"><w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
    );

    const FIELD_CLOSE: &str = r#"<w:r><w:fldChar w:fldCharType="end"/></w:r></w:dir>"#;

    const SDT_OPEN: &str = concat!(
        r#"<w:sdt><w:sdtPr><w:id w:val="4"/></w:sdtPr>"#,
        r#"<w:sdtContent><w:dir w:val="rtl">"#,
    );

    const DIRECTION_SCOPES: [(&str, &str, &str, &str); 10] = [
        ("paragraph", r#"<w:dir w:val="rtl">"#, "", "</w:dir>"),
        ("bdo", r#"<w:bdo w:val="rtl">"#, "", "</w:bdo>"),
        (
            "hyperlink",
            r#"<w:hyperlink w:anchor="target"><w:dir w:val="rtl">"#,
            "",
            "</w:dir></w:hyperlink>",
        ),
        (
            "insertion",
            r#"<w:ins w:id="1" w:author="Ann"><w:dir w:val="rtl">"#,
            "",
            "</w:dir></w:ins>",
        ),
        (
            "deletion",
            r#"<w:del w:id="2" w:author="Ann"><w:dir w:val="rtl">"#,
            "",
            "</w:dir></w:del>",
        ),
        (
            "content control",
            SDT_OPEN,
            "",
            "</w:dir></w:sdtContent></w:sdt>",
        ),
        (
            "simple field",
            r#"<w:fldSimple w:instr=" SUBJECT "><w:dir w:val="rtl">"#,
            "",
            "</w:dir></w:fldSimple>",
        ),
        (
            "nested",
            r#"<w:bdo w:val="ltr"><w:dir w:val="rtl">"#,
            "",
            "</w:dir></w:bdo>",
        ),
        (
            "adjacent",
            r#"<w:dir w:val="rtl">"#,
            r#"</w:dir><w:dir w:val="ltr">"#,
            "</w:dir>",
        ),
        ("complex field", FIELD_OPEN, "", FIELD_CLOSE),
    ];

    fn round_trip(xml: &str) -> String {
        serialize_paragraph(&parsed(xml), &mut context()).unwrap()
    }

    fn parsed(xml: &str) -> Paragraph {
        let limits = crate::xml::ParseLimits::default();
        let mut budget = crate::xml::ParseBudget::new(&limits);
        let document =
            crate::xml::parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
        let mut ids = crate::paragraph::HexIdAllocator::from_sha256(&"0".repeat(64)).unwrap();
        crate::paragraph::parse_paragraph(
            document.root().unwrap(),
            None,
            None,
            None,
            None,
            None,
            "word/document.xml",
            &mut budget,
            &mut ids,
            None,
            false,
            0,
        )
        .unwrap()
    }

    fn shown(xml: &str) -> String {
        let mut text = String::new();
        let mut rest = xml;
        while let Some(start) = rest.find('<') {
            let Some(end) = rest[start..].find('>') else {
                break;
            };
            let tag = &rest[start + 1..start + end];
            rest = &rest[start + end + 1..];
            let name = tag.split([' ', '/']).next().unwrap_or_default();
            if matches!(name, "w:t" | "w:delText") && !tag.ends_with('/') {
                text.push_str(&rest[..rest.find('<').unwrap_or(rest.len())]);
            }
        }
        text
    }

    fn region<'a>(xml: &'a str, open: &str, close: &str) -> &'a str {
        let start = xml.find(open).map_or(xml.len(), |at| at + open.len());
        let rest = &xml[start..];
        rest.find(close).map_or("", |end| &rest[..end])
    }

    fn container(name: &str) -> Option<(&'static str, &'static str)> {
        match name {
            "hyperlink" => Some(("<w:hyperlink", "</w:hyperlink>")),
            "insertion" => Some(("<w:ins ", "</w:ins>")),
            "deletion" => Some(("<w:del ", "</w:del>")),
            "content control" => Some(("<w:sdtContent>", "</w:sdtContent>")),
            "simple field" => Some(("<w:fldSimple", "</w:fldSimple>")),
            _ => None,
        }
    }

    #[test]
    fn direction_scopes_save_back_around_their_text_in_every_container() {
        for (name, open, middle, close) in DIRECTION_SCOPES {
            let text = if name == "deletion" { "delText" } else { "t" };
            let source = format!(
                r#"<w:p xmlns:w="{WORD}"><w:r><w:t xml:space="preserve">Before </w:t></w:r>{open}<w:r><w:{text}>two</w:{text}></w:r>{middle}<w:r><w:{text} xml:space="preserve"> runs</w:{text}></w:r>{close}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>"#
            );
            let output = round_trip(&source);
            assert_eq!(shown(&output), "Before two runs after", "{name}");
            for tag in SCOPE_TAGS {
                assert_eq!(
                    output.matches(tag).count(),
                    source.matches(tag).count(),
                    "{name} {tag}: {output}"
                );
            }
            let (inner_open, inner_close) = if name == "bdo" { BDO } else { DIR };
            let inner = if name == "adjacent" {
                "two"
            } else {
                "two runs"
            };
            assert_eq!(
                shown(region(&output, inner_open, inner_close)),
                inner,
                "{name}: {output}"
            );
            if name == "adjacent" {
                let ltr = region(&output, r#"<w:dir w:val="ltr">"#, "</w:dir>");
                assert_eq!(shown(ltr), " runs", "{output}");
            }
            if name == "nested" {
                let ltr = region(&output, r#"<w:bdo w:val="ltr">"#, "</w:bdo>");
                assert!(ltr.contains(DIR.0), "{output}");
                assert_eq!(shown(ltr), "two runs", "{output}");
            }
            if let Some((outer_open, outer_close)) = container(name) {
                assert!(
                    region(&output, outer_open, outer_close).contains(DIR.0),
                    "{name}: {output}"
                );
            }
        }
    }

    #[test]
    fn complex_field_result_scope_holds_its_separate_character() {
        let source = format!(
            r#"<w:p xmlns:w="{WORD}"><w:r><w:t xml:space="preserve">Before </w:t></w:r>{FIELD_OPEN}<w:r><w:t>two</w:t></w:r>{FIELD_CLOSE}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>"#
        );
        let output = round_trip(&source);
        assert_eq!(shown(&output), "Before two after", "{output}");
        let scoped = region(&output, DIR.0, DIR.1);
        assert!(
            scoped.starts_with(r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#),
            "{output}"
        );
        assert!(
            scoped.ends_with(r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#),
            "{output}"
        );
        assert_eq!(shown(scoped), "two", "{output}");
    }

    #[test]
    fn unmodeled_word_inline_elements_save_back_verbatim() {
        let output = round_trip(&format!(
            r#"<w:p xmlns:w="{WORD}"><w:r><w:t>a</w:t></w:r><w:zzz><w:r><w:t>x</w:t></w:r></w:zzz><w:r><w:t>b</w:t></w:r></w:p>"#
        ));
        assert!(output.contains("<w:zzz"), "{output}");
        assert!(
            output.contains("<w:r><w:t>x</w:t></w:r></w:zzz>"),
            "{output}"
        );
        assert_eq!(shown(&output), "axb", "{output}");
    }

    const RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const MATH: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    const BEGIN: &str = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#;

    const INSTRUCTION: &str = "<w:r><w:instrText>PAGE</w:instrText></w:r>";

    const SEPARATE: &str = r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#;

    const END: &str = r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#;

    const EQUATION: &str = "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>";

    fn text_run(text: &str) -> String {
        format!("<w:r><w:t>{text}</w:t></w:r>")
    }

    fn rtl(inner: &str) -> String {
        format!(r#"<w:dir w:val="rtl">{inner}</w:dir>"#)
    }

    fn paragraph_xml(content: &str) -> String {
        format!(
            r#"<w:p xmlns:w="{WORD}" xmlns:r="{RELATIONSHIPS}" xmlns:m="{MATH}">{content}</w:p>"#
        )
    }

    fn written(content: &str) -> String {
        round_trip(&paragraph_xml(content))
    }

    fn rtl_code_run(text: &str) -> Value {
        json!({
            "type": "run",
            "content": [{"type": "instrText", "text": text}],
            "directionScopes": [{"kind": "dir", "val": "rtl", "id": 0}]
        })
    }

    fn run_holding<'a>(xml: &'a str, needle: &str) -> &'a str {
        let at = xml.find(needle).expect("the unscoped save holds the run");
        let start = xml[..at].rfind("<w:r>").expect("the needle sits in a run");
        let end = at + xml[at..].find("</w:r>").expect("the run closes") + "</w:r>".len();
        &xml[start..end]
    }

    fn wrapped(xml: &str, node: &str, open: &str, close: &str) -> String {
        assert_eq!(xml.matches(node).count(), 1, "{node}: {xml}");
        xml.replacen(node, &format!("{open}{node}{close}"), 1)
    }

    #[test]
    fn a_direction_scope_closes_before_every_node_that_does_not_carry_it() {
        let (a, b, c, d) = (text_run("A"), text_run("B"), text_run("C"), text_run("D"));
        let cases = [
            (
                "equation",
                format!("{a}{}{EQUATION}{c}", rtl(&b)),
                format!("{a}{b}{EQUATION}{c}"),
                "<w:t>B</w:t>",
            ),
            (
                "comment range end",
                format!(
                    r#"{a}<w:commentRangeStart w:id="1"/>{}<w:commentRangeEnd w:id="1"/>{c}"#,
                    rtl(&b)
                ),
                format!(r#"{a}<w:commentRangeStart w:id="1"/>{b}<w:commentRangeEnd w:id="1"/>{c}"#),
                "<w:t>B</w:t>",
            ),
            (
                "comment range end closing the paragraph",
                format!(
                    r#"{a}<w:commentRangeStart w:id="1"/>{}<w:commentRangeEnd w:id="1"/>"#,
                    rtl(&b)
                ),
                format!(r#"{a}<w:commentRangeStart w:id="1"/>{b}<w:commentRangeEnd w:id="1"/>"#),
                "<w:t>B</w:t>",
            ),
            (
                "kept element",
                format!(r#"{a}{}<w:subDoc r:id="rId9"/>{c}"#, rtl(&b)),
                format!(r#"{a}{b}<w:subDoc r:id="rId9"/>{c}"#),
                "<w:t>B</w:t>",
            ),
            (
                "bookmark end",
                format!(
                    r#"{a}<w:bookmarkStart w:id="3" w:name="mark"/>{}<w:bookmarkEnd w:id="3"/>{c}"#,
                    rtl(&b)
                ),
                format!(
                    r#"{a}<w:bookmarkStart w:id="3" w:name="mark"/>{b}<w:bookmarkEnd w:id="3"/>{c}"#
                ),
                "<w:t>B</w:t>",
            ),
            (
                "scoped field code with an unscoped result",
                format!("{a}{BEGIN}{}{SEPARATE}{b}{END}{c}", rtl(INSTRUCTION)),
                format!("{a}{BEGIN}{INSTRUCTION}{SEPARATE}{b}{END}{c}"),
                "<w:instrText",
            ),
            (
                "field result whose last run alone is scoped",
                format!("{a}{BEGIN}{INSTRUCTION}{SEPARATE}{b}{}{END}{d}", rtl(&c)),
                format!("{a}{BEGIN}{INSTRUCTION}{SEPARATE}{b}{c}{END}{d}"),
                "<w:t>C</w:t>",
            ),
        ];
        for (name, scoped, plain, needle) in cases {
            let unscoped = written(&plain);
            let node = run_holding(&unscoped, needle).to_owned();
            let expected = wrapped(&unscoped, &node, DIR.0, DIR.1);
            assert_eq!(written(&scoped), expected, "{name}");
        }
    }

    #[test]
    fn a_direction_element_without_a_value_saves_without_one() {
        let (a, b, c) = (text_run("A"), text_run("B"), text_run("C"));
        let unscoped = written(&format!("{a}{b}{c}"));
        let expected = wrapped(&unscoped, &b, "<w:dir>", "</w:dir>");
        assert_eq!(written(&format!("{a}<w:dir>{b}</w:dir>{c}")), expected);
    }

    fn runs_spanning<'a>(xml: &'a str, first: &str, last: &str) -> &'a str {
        let at = |run: &str| run.as_ptr() as usize - xml.as_ptr() as usize;
        let first = run_holding(xml, first);
        let last = run_holding(xml, last);
        &xml[at(first)..at(last) + last.len()]
    }

    #[test]
    fn a_bookmark_between_two_runs_of_one_scope_is_written_inside_it() {
        let first = "<w:r><w:t>مرحبا</w:t></w:r>";
        let second = r#"<w:r><w:t xml:space="preserve"> 123</w:t></w:r>"#;
        let start = r#"<w:bookmarkStart w:id="0" w:name="b"/>"#;
        let end = r#"<w:bookmarkEnd w:id="0"/>"#;
        let unscoped = written(&format!("{first}{start}{second}{end}"));
        let held = runs_spanning(&unscoped, "مرحبا", " 123");
        assert!(held.contains(start), "{unscoped}");
        let expected = wrapped(&unscoped, held, DIR.0, DIR.1);
        let output = written(&rtl(&format!("{first}{start}{second}{end}")));
        assert_eq!(output, expected);
        assert_eq!(output.matches("<w:dir").count(), 1, "{output}");
        assert!(output.contains(&format!("</w:dir>{end}")), "{output}");
    }

    #[test]
    fn a_node_without_width_between_two_runs_of_one_scope_is_written_inside_it() {
        let (a, b, c, d) = (text_run("A"), text_run("B"), text_run("C"), text_run("D"));
        let close = r#"<w:commentRangeEnd w:id="1"/>"#;
        for middle in [
            r#"<w:commentRangeStart w:id="1"/>"#,
            r#"<w:customXmlInsRangeStart w:id="7" w:author="Grace"/>"#,
            r#"<w:moveFromRangeStart w:id="5" w:name="move"/>"#,
            r#"<w:bookmarkStart w:id="3" w:name="mark"/><w:commentRangeStart w:id="1"/>"#,
        ] {
            let unscoped = written(&format!("{a}{b}{middle}{c}{close}{d}"));
            let held = runs_spanning(&unscoped, "<w:t>B</w:t>", "<w:t>C</w:t>");
            let expected = wrapped(&unscoped, held, DIR.0, DIR.1);
            let output = written(&format!("{a}{}{close}{d}", rtl(&format!("{b}{middle}{c}"))));
            assert_eq!(output, expected, "{middle}");
            assert_eq!(output.matches("<w:dir").count(), 1, "{middle}: {output}");
        }
    }

    #[test]
    fn a_bookmark_inside_one_scope_stays_inside_it_in_a_hyperlink_and_a_hyperlink_inside_a_content_control()
     {
        let (a, b, c, d) = (text_run("A"), text_run("B"), text_run("C"), text_run("D"));
        let bookmark = r#"<w:bookmarkStart w:id="3" w:name="mark"/>"#;
        for (open, close) in [
            (r#"<w:hyperlink w:anchor="target">"#, "</w:hyperlink>"),
            (
                r#"<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:hyperlink w:anchor="target">"#,
                "</w:hyperlink></w:sdtContent></w:sdt>",
            ),
        ] {
            let unscoped = written(&format!("{a}{open}{b}{bookmark}{c}{close}{d}"));
            let held = runs_spanning(&unscoped, "<w:t>B</w:t>", "<w:t>C</w:t>");
            assert!(held.contains(bookmark), "{unscoped}");
            let expected = wrapped(&unscoped, held, DIR.0, DIR.1);
            let output = written(&format!(
                "{a}{open}{}{close}{d}",
                rtl(&format!("{b}{bookmark}{c}"))
            ));
            assert_eq!(output, expected, "{open}");
        }
    }

    #[test]
    fn a_complex_field_kept_as_its_source_run_stays_inside_its_direction_scope() {
        let (a, b) = (text_run("A"), text_run("B"));
        let field = concat!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGE </w:instrText>"#,
            r#"<w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r>"#
        );
        let output = written(&format!("{a}{}{b}", rtl(field)));
        assert_eq!(output.matches("<w:dir ").count(), 1, "{output}");
        let open = output.find("<w:dir ").unwrap();
        let close = output.find("</w:dir>").unwrap();
        let begin = output.find(r#"w:fldCharType="begin""#).unwrap();
        let end = output.find(r#"w:fldCharType="end""#).unwrap();
        assert!(open < begin && end < close, "{output}");
        assert!(output.find("<w:t>A</w:t>").unwrap() < open, "{output}");
        assert!(close < output.find("<w:t>B</w:t>").unwrap(), "{output}");
    }

    const TWO_FIELDS: &str = concat!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGE </w:instrText>"#,
        r#"<w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/>"#,
        r#"<w:t xml:space="preserve"> of </w:t><w:fldChar w:fldCharType="begin"/>"#,
        r#"<w:instrText xml:space="preserve"> NUMPAGES </w:instrText><w:fldChar w:fldCharType="separate"/>"#,
        r#"<w:t>2</w:t><w:fldChar w:fldCharType="end"/></w:r>"#
    );

    #[test]
    fn a_run_of_two_complex_fields_kept_as_its_source_is_written_once_inside_its_direction_scope() {
        let (a, b) = (text_run("A"), text_run("B"));
        let output = written(&format!("{a}{}{b}", rtl(TWO_FIELDS)));
        assert_eq!(output.matches("<w:dir ").count(), 1, "{output}");
        assert_eq!(output.matches(r#"w:fldCharType="begin""#).count(), 2, "{output}");
        let open = output.find("<w:dir ").unwrap();
        let close = output.find("</w:dir>").unwrap();
        assert!(open < output.find(r#"w:fldCharType="begin""#).unwrap(), "{output}");
        assert!(output.rfind(r#"w:fldCharType="end""#).unwrap() < close, "{output}");
    }

    #[test]
    fn a_run_of_two_complex_fields_without_a_scope_is_written_as_its_source() {
        let (a, b) = (text_run("A"), text_run("B"));
        let output = written(&format!("{a}{TWO_FIELDS}{b}"));
        assert!(!output.contains("w:dir"), "{output}");
        assert_eq!(output.matches(TWO_FIELDS).count(), 1, "{output}");
    }

    #[test]
    fn a_complex_field_rebuilt_from_a_changed_instruction_stays_inside_its_direction_scope() {
        let (a, b) = (text_run("A"), text_run("B"));
        let mut paragraph = parsed(&paragraph_xml(&format!("{a}{}{b}", rtl(TWO_FIELDS))));
        let field = paragraph
            .content
            .iter_mut()
            .find_map(|content| match content {
                ParagraphContent::Inline(InlineNode::ComplexField(field)) => Some(field),
                _ => None,
            })
            .expect("the run parses as a complex field");
        assert!(field.source_run_xml.is_some());
        field.instruction = "SECTIONPAGES".to_owned();
        let output = serialize_paragraph(&paragraph, &mut context()).unwrap();
        assert!(!output.contains(" NUMPAGES "), "{output}");
        assert_eq!(output.matches("<w:dir ").count(), 1, "{output}");
        let open = output.find("<w:dir ").unwrap();
        let close = output.find("</w:dir>").unwrap();
        let begin = output.find(r#"w:fldCharType="begin""#).unwrap();
        let instruction = output.find("SECTIONPAGES").unwrap();
        let end = output.rfind(r#"w:fldCharType="end""#).unwrap();
        assert!(open < begin && begin < instruction && end < close, "{output}");
        assert!(output.find("<w:t>A</w:t>").unwrap() < open, "{output}");
        assert!(close < output.find("<w:t>B</w:t>").unwrap(), "{output}");
    }

    #[test]
    fn a_bookmark_directly_inside_an_inline_content_control_is_dropped_on_a_rewrite() {
        let (a, b, c, d) = (text_run("A"), text_run("B"), text_run("C"), text_run("D"));
        let bookmark = r#"<w:bookmarkStart w:id="3" w:name="mark"/>"#;
        let open = r#"<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>"#;
        let close = "</w:sdtContent></w:sdt>";
        let output = written(&format!("{a}{open}{b}{bookmark}{c}{close}{d}"));
        assert!(!output.contains("w:bookmarkStart"), "{output}");
        assert!(
            output.contains("<w:t>B</w:t>") && output.contains("<w:t>C</w:t>"),
            "{output}"
        );
    }

    #[test]
    fn a_node_without_width_between_two_different_scopes_is_written_outside_both() {
        let (a, b, c, d) = (text_run("A"), text_run("B"), text_run("C"), text_run("D"));
        let bookmark = r#"<w:bookmarkStart w:id="3" w:name="mark"/>"#;
        let ltr = |inner: &str| format!(r#"<w:dir w:val="ltr">{inner}</w:dir>"#);
        let bdo = |inner: &str| format!(r#"<w:bdo w:val="rtl">{inner}</w:bdo>"#);
        let unscoped = written(&format!("{a}{b}{bookmark}{c}{d}"));
        let first = run_holding(&unscoped, "<w:t>B</w:t>").to_owned();
        let second = run_holding(&unscoped, "<w:t>C</w:t>").to_owned();
        let around_first = wrapped(&unscoped, &first, DIR.0, DIR.1);
        for (source, expected) in [
            (
                format!("{a}{}{bookmark}{}{d}", rtl(&b), ltr(&c)),
                wrapped(&around_first, &second, r#"<w:dir w:val="ltr">"#, "</w:dir>"),
            ),
            (
                format!("{a}{}{bookmark}{}{d}", rtl(&b), bdo(&c)),
                wrapped(&around_first, &second, BDO.0, BDO.1),
            ),
            (
                format!("{a}{}{bookmark}{}{d}", rtl(&b), rtl(&c)),
                wrapped(&around_first, &second, DIR.0, DIR.1),
            ),
        ] {
            let output = written(&source);
            assert_eq!(output, expected, "{source}");
            assert!(output.contains(&format!("</w:dir>{bookmark}")), "{output}");
        }
    }

    #[test]
    fn a_mark_between_an_inner_scope_run_and_an_outer_scope_run_keeps_the_outer_open() {
        let content = format!(
            r#"<w:bdo w:val="rtl"><w:dir w:val="ltr">{}</w:dir><w:bookmarkStart w:id="0" w:name="b"/>{}</w:bdo>"#,
            text_run("AB"),
            text_run("CD")
        );
        assert_eq!(written(&content), paragraph_xml(&content));
    }

    #[test]
    fn a_kept_element_that_is_no_range_or_annotation_mark_is_written_outside_the_scope_around_it() {
        let (a, b, c, d) = (text_run("A"), text_run("B"), text_run("C"), text_run("D"));
        let sub_doc = r#"<w:subDoc r:id="rId9"/>"#;
        let unscoped = written(&format!("{a}{b}{sub_doc}{c}{d}"));
        let expected = wrapped(&wrapped(&unscoped, &b, DIR.0, DIR.1), &c, DIR.0, DIR.1);
        let output = written(&format!("{a}{}{d}", rtl(&format!("{b}{sub_doc}{c}"))));
        assert_eq!(output, expected);
        assert_eq!(output.matches("<w:dir").count(), 2, "{output}");
    }

    #[test]
    fn a_mark_before_a_hyperlink_that_writes_no_run_of_the_scope_is_written_outside_it() {
        let (a, b) = (text_run("A"), text_run("B"));
        let bookmark = r#"<w:bookmarkStart w:id="3" w:name="mark"/>"#;
        let link = format!(
            r#"<w:hyperlink w:anchor="target"><w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>{b}</w:sdtContent></w:sdt></w:hyperlink>"#
        );
        let unscoped = written(&format!("{a}{bookmark}{link}"));
        let expected = wrapped(&unscoped, &a, DIR.0, DIR.1);
        assert_eq!(written(&rtl(&format!("{a}{bookmark}{link}"))), expected);
    }

    #[test]
    fn a_code_bookmark_of_a_rebuilt_field_stays_inside_the_scope_of_its_code_runs() {
        let field: ComplexField = serde_json::from_value(json!({
            "type": "complexField",
            "instruction": "QUOTE x",
            "fieldType": "QUOTE",
            "fieldCode": [rtl_code_run("QUOTE x")],
            "fieldResult": [],
            "structuredCode": {"inline": [
                {"type": "bookmarkStart", "id": 5, "name": "CodeMark"},
                rtl_code_run("QUOTE y")
            ]}
        }))
        .unwrap();
        let output = serialize_complex_field(&field, &mut context()).unwrap();
        assert_eq!(output.matches("<w:dir").count(), 1, "{output}");
        let order = [
            r#"w:fldCharType="begin""#,
            "CodeMark",
            "QUOTE x",
            r#"w:fldCharType="separate""#,
        ];
        let at: Vec<_> = order.iter().map(|needle| output.find(needle)).collect();
        assert!(
            at.iter().all(Option::is_some) && at.windows(2).all(|pair| pair[0] < pair[1]),
            "{at:?}: {output}"
        );
    }

    fn two_thousand_marks(scoped: bool) -> String {
        let mut content = String::new();
        for id in 0..2000 {
            content.push_str(&text_run("A"));
            content.push_str(&format!(
                r#"<w:bookmarkStart w:id="{id}" w:name="m{id}"/><w:bookmarkEnd w:id="{id}"/>"#
            ));
        }
        content.push_str(&text_run("A"));
        if scoped { rtl(&content) } else { content }
    }

    #[test]
    fn a_paragraph_of_two_thousand_bookmark_pairs_saves_as_it_reads() {
        for scoped in [false, true] {
            let content = two_thousand_marks(scoped);
            assert!(
                written(&content) == paragraph_xml(&content),
                "scoped: {scoped}"
            );
        }
    }

    #[test]
    #[ignore]
    fn measure_two_thousand_marks_paragraph_rewrite() {
        for (name, scoped) in [("no scope", false), ("one scope over all", true)] {
            let paragraph = parsed(&paragraph_xml(&two_thousand_marks(scoped)));
            let mut times: Vec<_> = (0..5)
                .map(|_| {
                    let start = std::time::Instant::now();
                    serialize_paragraph(&paragraph, &mut context()).unwrap();
                    start.elapsed()
                })
                .collect();
            times.sort();
            eprintln!("two thousand bookmark pairs, {name}: median {:?}", times[2]);
        }
    }

    fn two_thousand_comment_ranges(scoped: bool) -> String {
        let mut content = String::new();
        for id in 0..2000 {
            content.push_str(&text_run("A"));
            content.push_str(&format!(r#"<w:commentRangeStart w:id="{id}"/>"#));
            content.push_str(&text_run("B"));
            content.push_str(&format!(
                r#"<w:commentRangeEnd w:id="{id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{id}"/></w:r>"#
            ));
        }
        content.push_str(&text_run("A"));
        if scoped { rtl(&content) } else { content }
    }

    #[test]
    #[ignore]
    fn measure_two_thousand_comment_ranges_paragraph_rewrite() {
        for (name, scoped) in [("no scope", false), ("one scope over all", true)] {
            let paragraph = parsed(&paragraph_xml(&two_thousand_comment_ranges(scoped)));
            let mut times: Vec<_> = (0..5)
                .map(|_| {
                    let start = std::time::Instant::now();
                    serialize_paragraph(&paragraph, &mut context()).unwrap();
                    start.elapsed()
                })
                .collect();
            times.sort();
            eprintln!(
                "two thousand comment ranges with references, {name}: median {:?}",
                times[2]
            );
        }
    }

    #[test]
    fn kept_word_elements_inside_a_complex_field_stay_between_its_begin_and_end() {
        let (a, b, c) = (text_run("A"), text_run("B"), text_run("C"));
        let output = written(&format!(
            r#"{a}{BEGIN}{INSTRUCTION}{SEPARATE}<w:customXmlInsRangeStart w:id="7" w:author="Grace"/>{b}<w:customXmlInsRangeEnd w:id="7"/><w:subDoc r:id="rId9"/>{END}{c}"#
        ));
        let order = [
            r#"w:fldCharType="begin""#,
            "<w:customXmlInsRangeStart",
            "<w:t>B</w:t>",
            "<w:customXmlInsRangeEnd",
            "<w:subDoc",
            r#"w:fldCharType="end""#,
        ];
        let counts: Vec<_> = order
            .iter()
            .map(|needle| output.matches(needle).count())
            .collect();
        assert_eq!(counts, [1; 6], "{output}");
        let at: Vec<_> = order.iter().map(|needle| output.find(needle)).collect();
        assert!(
            at.windows(2).all(|pair| pair[0] < pair[1]),
            "{at:?}: {output}"
        );
    }

    #[test]
    fn rejects_unvalidated_math_and_sdt_raw_xml() {
        let mut properties = SdtProperties {
            sdt_type: "richText".to_owned(),
            id: None,
            alias: None,
            tag: None,
            lock: None,
            placeholder: None,
            showing_placeholder: None,
            date_format: None,
            list_items: None,
            checked: None,
            run_properties: None,
            temporary: None,
            label: None,
            tab_index: None,
            multi_line: None,
            date_state: None,
            list_last_value: None,
            checked_state: None,
            unchecked_state: None,
            gallery: None,
            appearance: None,
            color: None,
            control_state: None,
            repeating_section: None,
            repeating_section_item: None,
            data_binding: None,
            raw_properties_xml: Some("<w:sdtPr/><evil/>".to_owned()),
            raw_end_properties_xml: None,
        };
        assert!(serialize_sdt_properties(&properties).is_err());
        properties.raw_properties_xml = None;
        assert!(serialize_sdt_properties(&properties).is_ok());
    }

    fn run_value(content: Value) -> Value {
        json!({"type": "run", "content": [content]})
    }

    fn instruction_run() -> Value {
        run_value(json!({"type": "instrText", "text": "QUOTE x"}))
    }

    fn spanning_field(code: Value, result: Value, field_result: Value) -> ComplexField {
        serde_json::from_value(json!({
            "type": "complexField",
            "instruction": "QUOTE x",
            "fieldType": "QUOTE",
            "fieldCode": [instruction_run()],
            "fieldResult": field_result,
            "structuredCode": code,
            "structuredResult": result
        }))
        .unwrap()
    }

    #[test]
    fn a_field_whose_code_spans_paragraphs_writes_its_code_runs() {
        let code = json!({"inline": [
            {"type": "bookmarkStart", "id": 5, "name": "CodeMark"},
            instruction_run()
        ], "blockCount": 1});
        let field = spanning_field(code, Value::Null, json!([]));
        let output = serialize_complex_field(&field, &mut context()).unwrap();
        assert!(output.contains("QUOTE x"), "{output}");
        assert!(!output.contains("CodeMark"), "{output}");
    }

    #[test]
    fn a_field_whose_result_spans_paragraphs_writes_its_result_runs() {
        let copy = run_value(json!({"type": "text", "text": "InlineCopy"}));
        let written = run_value(json!({"type": "text", "text": "RunResult"}));
        for count in [0, 1] {
            let result = json!({"inline": [copy.clone()], "blockCount": count});
            let field = spanning_field(Value::Null, result, json!([written.clone()]));
            let output = serialize_complex_field(&field, &mut context()).unwrap();
            assert!(output.contains("RunResult"), "{count}: {output}");
            assert!(!output.contains("InlineCopy"), "{count}: {output}");
        }
    }

    #[test]
    fn a_field_whose_result_spans_paragraphs_holds_the_references_of_its_result_runs() {
        let reference = |id: u32| run_value(json!({"type": "commentReference", "id": id}));
        let result = json!({"inline": [reference(3)], "blockCount": 1});
        let field = spanning_field(Value::Null, result, json!([reference(4)]));
        let node = InlineNode::ComplexField(Box::new(field));
        assert!(has_comment_reference(&node, 4.0));
        assert!(!has_comment_reference(&node, 3.0));
    }
}
