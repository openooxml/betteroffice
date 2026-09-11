use std::collections::HashSet;
use std::ops::Range;

use quick_xml::{Reader, events::Event};

use crate::PptxError;
use crate::comments::{
    CommentFlavor, CommentWrite, CommentsWrite, NS_A, NS_P, NS_P188, NS_PC, authors_xml,
    legacy_comment_element, modern_body_element, modern_comment_element, text_body,
};
use crate::xml::{ParseBudget, XmlElement, parse_xml, serialize_xml_fragment};

struct Span {
    range: Range<usize>,
    opening_end: usize,
    closing_start: usize,
    children: Vec<Span>,
}

struct Patch {
    range: Range<usize>,
    bytes: Vec<u8>,
}

struct Source<'a> {
    bytes: &'a [u8],
    patches: Vec<Patch>,
}

impl Source<'_> {
    fn replace(&mut self, range: Range<usize>, bytes: Vec<u8>) {
        self.patches.push(Patch { range, bytes });
    }

    fn append(&mut self, element: &XmlElement, span: &Span, bytes: Vec<u8>) {
        if bytes.is_empty() {
            return;
        }
        if span.range.end == span.opening_end {
            let mut replacement = b">".to_vec();
            replacement.extend(bytes);
            replacement.extend(format!("</{}>", element.name).as_bytes());
            self.replace(span.opening_end - 2..span.opening_end, replacement);
        } else {
            self.replace(span.closing_start..span.closing_start, bytes);
        }
    }

    fn attribute(&mut self, element: &XmlElement, span: &Span, name: &str, value: &str) {
        if element.attribute(name) == Some(value) {
            return;
        }
        let escaped = quick_xml::escape::escape(value).into_owned().into_bytes();
        let mut position = span.range.start + 1 + element.name.len();
        while position < span.opening_end {
            while self.bytes[position].is_ascii_whitespace() {
                position += 1;
            }
            if matches!(self.bytes[position], b'/' | b'>') {
                break;
            }
            let start = position;
            while !self.bytes[position].is_ascii_whitespace() && self.bytes[position] != b'=' {
                position += 1;
            }
            let matches = &self.bytes[start..position] == name.as_bytes();
            while self.bytes[position] != b'=' {
                position += 1;
            }
            position += 1;
            while self.bytes[position].is_ascii_whitespace() {
                position += 1;
            }
            let quote = self.bytes[position];
            position += 1;
            let start = position;
            while self.bytes[position] != quote {
                position += 1;
            }
            if matches {
                self.replace(start..position, escaped);
                return;
            }
            position += 1;
        }
        let mut attribute = format!(" {name}=\"").into_bytes();
        attribute.extend(escaped);
        attribute.push(b'"');
        self.replace(position..position, attribute);
    }

    fn finish(mut self) -> Vec<u8> {
        self.patches.sort_by_key(|patch| patch.range.start);
        let mut output = Vec::new();
        let mut position = 0;
        for patch in self.patches {
            output.extend_from_slice(&self.bytes[position..patch.range.start]);
            output.extend(patch.bytes);
            position = patch.range.end;
        }
        output.extend_from_slice(&self.bytes[position..]);
        output
    }
}

fn spans(bytes: &[u8], part: &str) -> Result<Span, PptxError> {
    let mut reader = Reader::from_reader(bytes);
    let mut stack: Vec<Span> = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| PptxError::Write {
            part: part.to_owned(),
            message: error.to_string(),
        })?;
        let end = reader.buffer_position() as usize;
        let complete = match event {
            Event::Start(_) => {
                stack.push(Span {
                    range: start..end,
                    opening_end: end,
                    closing_start: end,
                    children: Vec::new(),
                });
                None
            }
            Event::Empty(_) => Some(Span {
                range: start..end,
                opening_end: end,
                closing_start: end,
                children: Vec::new(),
            }),
            Event::End(_) => stack.pop().map(|mut span| {
                span.range.end = end;
                span.closing_start = start;
                span
            }),
            Event::Eof => {
                return Err(PptxError::Write {
                    part: part.to_owned(),
                    message: "missing comment XML root".to_owned(),
                });
            }
            _ => None,
        };
        if let Some(span) = complete {
            if let Some(parent) = stack.last_mut() {
                parent.children.push(span);
            } else {
                return Ok(span);
            }
        }
    }
}

pub(crate) fn patch_comments_xml(
    bytes: &[u8],
    part: &str,
    write: &CommentsWrite,
    comments: &[CommentWrite],
    slide_id: u32,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<u8>, PptxError> {
    let root = parse_xml(bytes, part, budget)?;
    let span = spans(bytes, part)?;
    let mut source = Source {
        bytes,
        patches: Vec::new(),
    };
    patch_list(
        &mut source,
        &root,
        &span,
        write.flavor,
        comments,
        Some(slide_id),
    );
    Ok(source.finish())
}

fn patch_list(
    source: &mut Source<'_>,
    root: &XmlElement,
    span: &Span,
    flavor: CommentFlavor,
    comments: &[CommentWrite],
    slide_id: Option<u32>,
) {
    let mut existing = HashSet::new();
    let name = if slide_id.is_some() { "cm" } else { "reply" };
    for (element, child_span) in root.child_elements().zip(&span.children) {
        if element.local_name() != name {
            continue;
        }
        let id = match flavor {
            CommentFlavor::Legacy => format!(
                "{}:{}",
                element.attribute("authorId").unwrap_or("0"),
                element.attribute("idx").unwrap_or("1")
            ),
            CommentFlavor::Modern => element.attribute("id").unwrap_or_default().to_owned(),
        };
        existing.insert(id.clone());
        let Some(comment) = comments.iter().find(|comment| comment.id == id) else {
            source.replace(child_span.range.clone(), Vec::new());
            continue;
        };
        if let Some(status) = &comment.status {
            source.attribute(element, child_span, "status", status);
        }
        if flavor == CommentFlavor::Modern && slide_id.is_some() {
            if let Some((replies, replies_span)) = element
                .child_elements()
                .zip(&child_span.children)
                .find(|(child, _)| child.local_name() == "replyLst")
            {
                patch_list(
                    source,
                    replies,
                    replies_span,
                    flavor,
                    &comment.replies,
                    None,
                );
            } else if !comment.replies.is_empty() {
                let mut replies =
                    XmlElement::new("p188:replyLst").with_attribute("xmlns:p188", NS_P188);
                for reply in &comment.replies {
                    replies = replies.with_child(new_comment(reply, flavor, None));
                }
                let bytes = serialize_xml_fragment(&replies);
                if let Some((_, body)) = element
                    .child_elements()
                    .zip(&child_span.children)
                    .find(|(child, _)| child.local_name() == "txBody")
                {
                    source.replace(body.range.start..body.range.start, bytes);
                } else {
                    source.append(element, child_span, bytes);
                }
            }
        }
    }
    let mut added = Vec::new();
    for comment in comments
        .iter()
        .filter(|comment| !existing.contains(&comment.id))
    {
        added.extend(serialize_xml_fragment(&new_comment(
            comment, flavor, slide_id,
        )));
    }
    source.append(root, span, added);
}

fn new_comment(comment: &CommentWrite, flavor: CommentFlavor, slide_id: Option<u32>) -> XmlElement {
    match flavor {
        CommentFlavor::Legacy => legacy_comment_element(comment).with_attribute("xmlns:p", NS_P),
        CommentFlavor::Modern => match slide_id {
            Some(slide_id) => modern_comment_element(comment, slide_id),
            None => modern_body_element("p188:reply", comment).with_child(text_body(&comment.text)),
        }
        .with_attribute("xmlns:p188", NS_P188)
        .with_attribute("xmlns:a", NS_A)
        .with_attribute("xmlns:pc", NS_PC),
    }
}

pub(crate) fn patch_authors_xml(
    bytes: &[u8],
    part: &str,
    write: &CommentsWrite,
    budget: &mut ParseBudget<'_>,
) -> Result<Vec<u8>, PptxError> {
    let root = parse_xml(bytes, part, budget)?;
    let span = spans(bytes, part)?;
    let mut source = Source {
        bytes,
        patches: Vec::new(),
    };
    let mut existing = HashSet::new();
    let name = match write.flavor {
        CommentFlavor::Legacy => "cmAuthor",
        CommentFlavor::Modern => "author",
    };
    for (element, child_span) in root.child_elements().zip(&span.children) {
        if element.local_name() != name {
            continue;
        }
        let id = element.attribute("id").unwrap_or_default();
        existing.insert(id);
        if write.flavor == CommentFlavor::Legacy
            && let Some(author) = write.authors.iter().find(|author| author.id == id)
            && author.last_index
                > element
                    .attribute("lastIdx")
                    .and_then(|value| value.parse().ok())
                    .unwrap_or(0)
        {
            source.attribute(
                element,
                child_span,
                "lastIdx",
                &author.last_index.to_string(),
            );
        }
    }
    let generated = parse_xml(&authors_xml(write), part, budget)?;
    let mut added = Vec::new();
    for author in generated
        .child_elements()
        .filter(|author| !existing.contains(author.attribute("id").unwrap_or_default()))
    {
        let mut author = author.clone();
        match write.flavor {
            CommentFlavor::Legacy => author.set_attribute("xmlns:p", NS_P),
            CommentFlavor::Modern => author.set_attribute("xmlns:p188", NS_P188),
        }
        added.extend(serialize_xml_fragment(&author));
    }
    source.append(&root, &span, added);
    Ok(source.finish())
}
