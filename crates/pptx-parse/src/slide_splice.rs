use std::collections::HashSet;

use crate::comment_patch::{Span, spans};
use crate::xml::{
    ParseBudget, ParseLimits, XmlElement, XmlNode, parse_xml, serialize_xml_fragment,
};

pub(crate) fn splice_slide(
    source: &[u8],
    before: &XmlElement,
    after: &XmlElement,
) -> Option<Vec<u8>> {
    if before == after {
        return Some(source.to_vec());
    }
    if before.name != after.name {
        return None;
    }
    let span = spans(source, "slide").ok()?;
    if span.children.len() != before.child_elements().count() {
        return None;
    }
    let mut root_prefixes = HashSet::new();
    for key in before.attributes.keys() {
        if let Some(prefix) = key.strip_prefix("xmlns:") {
            root_prefixes.insert(prefix.to_owned());
        }
    }
    root_prefixes.insert("xml".to_owned());
    let splicer = Splicer {
        source,
        root_prefixes,
    };
    let range_start = span.range.start;
    let range_end = span.range.end;
    let body = splicer.splice_strict(before, after, &span)?;
    let mut output = Vec::with_capacity(source.len());
    output.extend_from_slice(&source[..range_start]);
    output.extend_from_slice(&body);
    output.extend_from_slice(&source[range_end..]);
    let limits = ParseLimits::default();
    let mut budget = ParseBudget::new(&limits);
    let reparsed = parse_xml(&output, "slide", &mut budget).ok()?;
    if !equivalent(&reparsed, after) {
        return None;
    }
    Some(output)
}

struct Splicer<'a> {
    source: &'a [u8],
    root_prefixes: HashSet<String>,
}

enum Op {
    Match(usize, usize),
    Delete(usize),
    Insert(usize),
}

impl Splicer<'_> {
    fn splice_or_fragment(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
        match self.splice_strict(before, after, span) {
            Some(bytes) => Some(bytes),
            None => self.checked_fragment(after),
        }
    }

    fn splice_strict(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
        if before == after {
            return Some(self.source[span.range.clone()].to_vec());
        }
        if before.name != after.name {
            return None;
        }
        if span.children.len() != before.child_elements().count() {
            return None;
        }
        let inner = self.splice_children(before, after, span)?;
        let opening = &self.source[span.range.start..span.opening_end];
        let mut start =
            crate::attribute_patch::patch_start(opening, &before.attributes, &after.attributes)?;
        let self_closing = opening.ends_with(b"/>");
        if after.children.is_empty() {
            if self_closing {
                return Some(start);
            }
            let mut output = start;
            output.extend_from_slice(&inner);
            output.extend_from_slice(&self.source[span.closing_start..span.range.end]);
            return Some(output);
        }
        if self_closing {
            if !start.ends_with(b"/>") {
                return None;
            }
            start.truncate(start.len() - 2);
            start.push(b'>');
            let mut output = start;
            output.extend_from_slice(&inner);
            output.extend_from_slice(format!("</{}>", after.name).as_bytes());
            Some(output)
        } else {
            let mut output = start;
            output.extend_from_slice(&inner);
            output.extend_from_slice(&self.source[span.closing_start..span.range.end]);
            Some(output)
        }
    }

    fn splice_children(
        &self,
        before: &XmlElement,
        after: &XmlElement,
        span: &Span,
    ) -> Option<Vec<u8>> {
        if before.children == after.children {
            return Some(self.source[span.opening_end..span.closing_start].to_vec());
        }
        let before_texts: Vec<&str> = before
            .children
            .iter()
            .filter_map(|node| match node {
                XmlNode::Text(text) => Some(text.as_str()),
                XmlNode::Element(_) => None,
            })
            .collect();
        let after_texts: Vec<&str> = after
            .children
            .iter()
            .filter_map(|node| match node {
                XmlNode::Text(text) => Some(text.as_str()),
                XmlNode::Element(_) => None,
            })
            .collect();
        if before_texts != after_texts {
            return None;
        }
        let meaningful = before
            .children
            .iter()
            .chain(after.children.iter())
            .any(|node| matches!(node, XmlNode::Text(text) if !text.trim().is_empty()));
        if meaningful {
            return None;
        }
        let before_elems: Vec<&XmlElement> = before.child_elements().collect();
        let after_elems: Vec<&XmlElement> = after.child_elements().collect();
        if span.children.len() != before_elems.len() {
            return None;
        }
        let before_hashes: Vec<u64> = before_elems
            .iter()
            .map(|element| structural_hash(element))
            .collect();
        let after_hashes: Vec<u64> = after_elems
            .iter()
            .map(|element| structural_hash(element))
            .collect();
        let ops = align(&before_elems, &after_elems, &before_hashes, &after_hashes)?;
        let deleted: Vec<usize> = ops
            .iter()
            .filter_map(|op| match op {
                Op::Delete(index) => Some(*index),
                _ => None,
            })
            .collect();
        let mut reused: HashSet<usize> = HashSet::new();
        let mut output = Vec::new();
        let mut cursor = span.opening_end;
        for op in &ops {
            match *op {
                Op::Match(before_index, after_index) => {
                    let child_span = &span.children[before_index];
                    output.extend_from_slice(&self.source[cursor..child_span.range.start]);
                    output.extend_from_slice(&self.splice_or_fragment(
                        before_elems[before_index],
                        after_elems[after_index],
                        child_span,
                    )?);
                    cursor = child_span.range.end;
                }
                Op::Delete(before_index) => {
                    let child_span = &span.children[before_index];
                    let gap = &self.source[cursor..child_span.range.start];
                    let gap_text = std::str::from_utf8(gap).unwrap_or("");
                    let whitespace_only = !gap.is_empty() && gap_text.trim().is_empty();
                    if !whitespace_only || after_texts.iter().any(|text| *text == gap_text) {
                        output.extend_from_slice(gap);
                    }
                    cursor = child_span.range.end;
                }
                Op::Insert(after_index) => {
                    let patched = after_elems[after_index];
                    let reused_index = deleted.iter().copied().find(|index| {
                        !reused.contains(index)
                            && before_hashes[*index] == after_hashes[after_index]
                            && before_elems[*index] == patched
                    });
                    if let Some(index) = reused_index {
                        reused.insert(index);
                        let child_span = &span.children[index];
                        output.extend_from_slice(&self.source[child_span.range.clone()]);
                    } else {
                        output.extend_from_slice(&self.checked_fragment(patched)?);
                    }
                }
            }
        }
        output.extend_from_slice(&self.source[cursor..span.closing_start]);
        Some(output)
    }

    fn checked_fragment(&self, element: &XmlElement) -> Option<Vec<u8>> {
        let mut declared = self.root_prefixes.clone();
        collect_declared(element, &mut declared);
        let mut used = HashSet::new();
        collect_used(element, &mut used);
        if used.iter().all(|prefix| declared.contains(prefix)) {
            Some(serialize_xml_fragment(element))
        } else {
            None
        }
    }
}

fn align(
    before: &[&XmlElement],
    after: &[&XmlElement],
    before_hashes: &[u64],
    after_hashes: &[u64],
) -> Option<Vec<Op>> {
    let n = before.len();
    let m = after.len();
    let before_ids: Vec<Option<&str>> = before
        .iter()
        .map(|element| drawing_object_id(element))
        .collect();
    let after_ids: Vec<Option<&str>> = after
        .iter()
        .map(|element| drawing_object_id(element))
        .collect();
    let equal = |i: usize, j: usize| before_hashes[i] == after_hashes[j] && before[i] == after[j];
    let can_match = |i: usize, j: usize| {
        before[i].name == after[j].name
            && match (before_ids[i], after_ids[j]) {
                (Some(left), Some(right)) => left == right,
                (Some(_), None) | (None, Some(_)) => false,
                (None, None) => true,
            }
    };
    let mut ops = Vec::new();
    let mut prefix = 0;
    while prefix < n && prefix < m && equal(prefix, prefix) {
        ops.push(Op::Match(prefix, prefix));
        prefix += 1;
    }
    let mut suffix = 0;
    while suffix < n - prefix
        && suffix < m - prefix
        && equal(n - 1 - suffix, m - 1 - suffix)
    {
        suffix += 1;
    }
    let middle_n = n - prefix - suffix;
    let middle_m = m - prefix - suffix;
    if middle_n
        .checked_mul(middle_m)
        .is_none_or(|size| size > 1_000_000)
    {
        return None;
    }
    let mut dp = vec![vec![0i32; middle_m + 1]; middle_n + 1];
    for (i, row) in dp.iter_mut().enumerate().take(middle_n) {
        row[middle_m] = -((middle_n - i) as i32);
    }
    for (j, cell) in dp[middle_n].iter_mut().enumerate().take(middle_m) {
        *cell = -((middle_m - j) as i32);
    }
    for i in (0..middle_n).rev() {
        for j in (0..middle_m).rev() {
            let mut best = i32::MIN / 2;
            if can_match(prefix + i, prefix + j) {
                let score = if equal(prefix + i, prefix + j) { 100 } else { 10 };
                best = best.max(score + dp[i + 1][j + 1]);
            }
            best = best.max(-1 + dp[i + 1][j]);
            best = best.max(-1 + dp[i][j + 1]);
            dp[i][j] = best;
        }
    }
    let mut i = 0;
    let mut j = 0;
    while i < middle_n || j < middle_m {
        if i < middle_n && j < middle_m && can_match(prefix + i, prefix + j) {
            let score = if equal(prefix + i, prefix + j) { 100 } else { 10 };
            if dp[i][j] == score + dp[i + 1][j + 1] {
                ops.push(Op::Match(prefix + i, prefix + j));
                i += 1;
                j += 1;
                continue;
            }
        }
        if i < middle_n && dp[i][j] == -1 + dp[i + 1][j] {
            ops.push(Op::Delete(prefix + i));
            i += 1;
            continue;
        }
        ops.push(Op::Insert(prefix + j));
        j += 1;
    }
    for offset in 0..suffix {
        ops.push(Op::Match(n - suffix + offset, m - suffix + offset));
    }
    Some(ops)
}

fn drawing_object_id(element: &XmlElement) -> Option<&str> {
    if element.local_name() == "AlternateContent" {
        return element
            .child_elements()
            .flat_map(|branch| branch.child_elements())
            .find_map(drawing_object_id);
    }
    let non_visual = if matches!(
        element.local_name(),
        "nvSpPr" | "nvPicPr" | "nvGrpSpPr" | "nvGraphicFramePr" | "nvCxnSpPr"
    ) {
        Some(element)
    } else {
        element.child_elements().find(|child| {
            matches!(
                child.local_name(),
                "nvSpPr" | "nvPicPr" | "nvGrpSpPr" | "nvGraphicFramePr" | "nvCxnSpPr"
            )
        })
    }?;
    non_visual
        .child_elements()
        .find(|child| child.local_name() == "cNvPr")?
        .attribute("id")
}

fn structural_hash(element: &XmlElement) -> u64 {
    fn mix(hash: &mut u64, bytes: &[u8]) {
        for byte in (bytes.len() as u64).to_le_bytes() {
            *hash ^= u64::from(byte);
            *hash = hash.wrapping_mul(0x100_0000_01b3);
        }
        for byte in bytes {
            *hash ^= u64::from(*byte);
            *hash = hash.wrapping_mul(0x100_0000_01b3);
        }
    }
    fn hash_element(hash: &mut u64, element: &XmlElement) {
        mix(hash, element.name.as_bytes());
        mix(hash, &(element.attributes.len() as u64).to_le_bytes());
        for (key, value) in &element.attributes {
            mix(hash, key.as_bytes());
            mix(hash, value.as_bytes());
        }
        mix(hash, &(element.children.len() as u64).to_le_bytes());
        for child in &element.children {
            match child {
                XmlNode::Text(text) => {
                    mix(hash, b"text");
                    mix(hash, text.as_bytes());
                }
                XmlNode::Element(child) => {
                    mix(hash, b"element");
                    hash_element(hash, child);
                }
            }
        }
    }
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    hash_element(&mut hash, element);
    hash
}

fn collect_declared(element: &XmlElement, declared: &mut HashSet<String>) {
    for key in element.attributes.keys() {
        if let Some(prefix) = key.strip_prefix("xmlns:") {
            declared.insert(prefix.to_owned());
        }
    }
    for child in element.child_elements() {
        collect_declared(child, declared);
    }
}

fn collect_used(element: &XmlElement, used: &mut HashSet<String>) {
    if let Some((prefix, _)) = element.name.split_once(':') {
        used.insert(prefix.to_owned());
    }
    for key in element.attributes.keys() {
        if key == "xmlns" || key.starts_with("xmlns:") {
            continue;
        }
        if let Some((prefix, _)) = key.split_once(':') {
            used.insert(prefix.to_owned());
        }
    }
    for child in element.child_elements() {
        collect_used(child, used);
    }
}

fn equivalent(before: &XmlElement, after: &XmlElement) -> bool {
    if before.name != after.name || before.attributes != after.attributes {
        return false;
    }
    enum Child<'a> {
        Element(&'a XmlElement),
        Text(String),
    }
    let merged = |element: &XmlElement| {
        let mut children: Vec<Child<'_>> = Vec::new();
        for child in &element.children {
            match (children.last_mut(), child) {
                (Some(Child::Text(previous)), XmlNode::Text(text)) => previous.push_str(text),
                (_, XmlNode::Text(text)) => children.push(Child::Text(text.clone())),
                (_, XmlNode::Element(child)) => children.push(Child::Element(child)),
            }
        }
        children
    };
    let left = merged(before);
    let right = merged(after);
    if left.len() != right.len() {
        return false;
    }
    left.iter()
        .zip(right.iter())
        .all(|(left, right)| match (left, right) {
            (Child::Element(left), Child::Element(right)) => equivalent(left, right),
            (Child::Text(left), Child::Text(right)) => left == right,
            _ => false,
        })
}
