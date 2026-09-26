//! The stylesheet subset the sandbox accepts, and what `simplecss` spends on
//! it. Rules are lone type, class and id compounds: `simplecss` backtracks
//! through combinators, so `g g g … .x` over deep nesting is exponential before
//! any budget sees it. Its cost is then charged in style work, about a
//! nanosecond each: it rescans the text from the start at every token it
//! rejects, and tests every rule against every element instance.

use resvg::usvg::roxmltree::{Document, Node};

use super::{
    MAX_SVG_SELECTOR_BYTES, MAX_SVG_SELECTOR_PARTS, MAX_SVG_STYLE_COPIES, MAX_SVG_STYLE_RULES,
    MAX_SVG_STYLE_WORK, SvgRefusal,
};

/// Every rule the document's `<style>` elements declare, one per selector.
#[derive(Default)]
pub(super) struct StyleSheet<'a> {
    rules: Vec<Rule<'a>>,
    blocks: Vec<Block<'a>>,
    /// Simple selectors across every rule; each is tested per element instance.
    parts: u64,
    /// What parsing the sheets costs `simplecss` once, in style work.
    work: u64,
    /// Selectors times the bytes of their block: `simplecss` copies a block's
    /// declarations once per selector of its list.
    copies: u64,
    /// Bytes of the longest dash list a block declares.
    dash: u64,
    strokes: Strokes,
}

/// What CSS or an element's attributes say about strokes, for the bound on
/// what `usvg` spends stroking every shape whole to measure it.
#[derive(Clone, Copy, Default)]
pub(super) struct Strokes {
    /// Whether anything may be stroked.
    pub(super) stroked: bool,
    /// The widest `stroke-width`, in user units.
    pub(super) width: f64,
    /// Whether a transform may rotate or skew, which has `usvg` stroke in
    /// canvas units instead of the shape's own.
    pub(super) turned: bool,
}

impl Strokes {
    /// Reads CSS text, returning the bytes of the longest dash list it
    /// declares. Declarations are read with the `simplecss` tokenizer `usvg`
    /// reads them with, comments, a leading `*` and `!important` included, in
    /// one pass and only where a stroke width or dash list may be declared:
    /// the caller has charged that pass as [`tokenized`].
    pub(super) fn read(&mut self, text: &str) -> Result<u64, SvgRefusal> {
        let has = |word: &[u8]| super::reference::contains_ignore_case(text, word);
        self.stroked |= has(b"stroke");
        self.turned |= has(b"transform");
        let mut dash = 0;
        if tokenized(text) == 0 {
            return Ok(dash);
        }
        for declaration in simplecss::DeclarationTokenizer::from(text) {
            match declaration.name {
                "stroke-width" => {
                    let width = super::geometry::stroke_width(declaration.value)?;
                    self.width = self.width.max(width);
                }
                "stroke-dasharray" => {
                    dash = dash.max(super::geometry::dash_list(declaration.value)?);
                }
                _ => {}
            }
        }
        Ok(dash)
    }

    pub(super) fn join(&mut self, other: Strokes) {
        self.stroked |= other.stroked;
        self.turned |= other.turned;
        self.width = self.width.max(other.width);
    }
}

struct Rule<'a> {
    tag: Option<&'a str>,
    classes: Vec<&'a str>,
    ids: Vec<&'a str>,
    block: usize,
}

impl Rule<'_> {
    fn write(&self, out: &mut String) {
        if self.tag.is_none() && self.classes.is_empty() && self.ids.is_empty() {
            out.push('*');
        }
        out.push_str(self.tag.unwrap_or_default());
        for class in &self.classes {
            out.push('.');
            out.push_str(class);
        }
        for id in &self.ids {
            out.push('#');
            out.push_str(id);
        }
    }
}

struct Block<'a> {
    declarations: &'a str,
    len: u64,
    /// Declarations at most, each an attribute `usvg` adds to the element.
    colons: u64,
    references: Vec<&'a str>,
}

impl<'a> StyleSheet<'a> {
    /// Writes every rule back out, each block's selectors joined as the text
    /// listed them and its declarations re-emitted by `declarations`.
    pub(super) fn write(
        &self,
        out: &mut String,
        mut declarations: impl FnMut(&str, &mut String) -> Result<(), SvgRefusal>,
    ) -> Result<(), SvgRefusal> {
        let mut rules = self.rules.iter().peekable();
        for (index, block) in self.blocks.iter().enumerate() {
            let mut first = true;
            while let Some(rule) = rules.next_if(|rule| rule.block == index) {
                if !first {
                    out.push(',');
                }
                first = false;
                rule.write(out);
            }
            out.push('{');
            declarations(block.declarations, out)?;
            out.push('}');
        }
        Ok(())
    }

    /// Reads every `<style>` element `usvg` would apply, wherever it sits: its
    /// first text node, all `usvg` reads.
    pub(super) fn collect(document: &'a Document<'a>) -> Result<Self, SvgRefusal> {
        let mut sheet = StyleSheet::default();
        let mut sheets = 0;
        for node in document.descendants() {
            if !node.is_element()
                || node.tag_name().name() != "style"
                || node
                    .attribute("type")
                    .is_some_and(|kind| kind != "text/css")
            {
                continue;
            }
            sheets += 1;
            if let Some(text) = node.text() {
                sheet.work = sheet.work.saturating_add(rescans(text.len()));
                if sheet.work > MAX_SVG_STYLE_WORK {
                    return Err(SvgRefusal::ExpansionTooLarge);
                }
                sheet.parse(text)?;
            }
            if sheets + sheet.rules.len() > MAX_SVG_STYLE_RULES {
                return Err(SvgRefusal::UnsupportedStyle);
            }
        }
        Ok(sheet)
    }

    /// Style work parsing the sheets costs, once per document.
    pub(super) fn work(&self) -> u64 {
        self.work
    }

    /// Bytes of the longest dash list any rule declares.
    pub(super) fn dash(&self) -> u64 {
        self.dash
    }

    /// What the rules say about strokes.
    pub(super) fn strokes(&self) -> Strokes {
        self.strokes
    }

    /// Style work testing every rule against one instance of `node`, for
    /// `simplecss` and for the audit's own match: a simple selector looks its
    /// attribute up among `node`'s and scans the value, and the first to fail
    /// ends the rule.
    pub(super) fn tests(&self, node: Node<'_, '_>) -> u64 {
        if self.rules.is_empty() {
            return 0;
        }
        let attributes = node.attributes().len() as u64;
        let values = ["class", "id"]
            .iter()
            .filter_map(|name| node.attribute(*name))
            .map(str::len)
            .sum::<usize>() as u64;
        (self.rules.len() as u64 * 8)
            .saturating_add(self.parts.saturating_mul(4 + 2 * attributes + values / 2))
    }

    /// Calls `matched` with the block length, its declarations at most and its
    /// same-document references, for every rule that may apply to `node`: a
    /// superset of what `simplecss` matches.
    pub(super) fn each_match(
        &self,
        node: Node<'_, '_>,
        mut matched: impl FnMut(u64, u64, &[&'a str]),
    ) {
        if self.rules.is_empty() {
            return;
        }
        let mut classes: Vec<&str> = node
            .attribute("class")
            .unwrap_or_default()
            .split_ascii_whitespace()
            .collect();
        classes.sort_unstable();
        let id = node.attribute("id");
        for rule in &self.rules {
            let applies = rule.tag.is_none_or(|tag| node.tag_name().name() == tag)
                && rule
                    .classes
                    .iter()
                    .all(|name| classes.binary_search(name).is_ok())
                && rule.ids.iter().all(|name| id == Some(*name));
            if applies {
                let block = &self.blocks[rule.block];
                matched(block.len, block.colons, &block.references);
            }
        }
    }

    fn parse(&mut self, text: &'a str) -> Result<(), SvgRefusal> {
        if text.contains("/*") || text.contains('\\') {
            return Err(SvgRefusal::UnsupportedStyle);
        }
        let mut rest = text.trim_start_matches(|c: char| c.is_ascii_whitespace());
        while !rest.is_empty() {
            let open = rest.find('{').ok_or(SvgRefusal::UnsupportedStyle)?;
            let close = rest[open..].find('}').ok_or(SvgRefusal::UnsupportedStyle)? + open;
            let declarations = &rest[open + 1..close];
            if declarations.contains('{') || !closed(declarations) {
                return Err(SvgRefusal::UnsupportedStyle);
            }
            screen(declarations)?;
            self.work = self.work.saturating_add(tokenized(declarations));
            if self.work > MAX_SVG_STYLE_WORK {
                return Err(SvgRefusal::ExpansionTooLarge);
            }
            self.dash = self.dash.max(self.strokes.read(declarations)?);
            let mut references = Vec::new();
            super::reference::css(declarations, &mut references)?;
            let block = self.blocks.len();
            self.blocks.push(Block {
                declarations,
                len: declarations.len() as u64,
                colons: declarations.bytes().filter(|byte| *byte == b':').count() as u64,
                references,
            });
            for selector in rest[..open].split(',') {
                let rule = compound(
                    selector.trim_matches(|c: char| c.is_ascii_whitespace()),
                    block,
                )
                .ok_or(SvgRefusal::UnsupportedStyle)?;
                if self.rules.len() >= MAX_SVG_STYLE_RULES {
                    return Err(SvgRefusal::UnsupportedStyle);
                }
                self.parts += (rule.tag.is_some() as usize + rule.classes.len() + rule.ids.len())
                    .max(1) as u64;
                self.copies = self.copies.saturating_add(declarations.len() as u64);
                self.rules.push(rule);
            }
            if self.copies > MAX_SVG_STYLE_COPIES {
                return Err(SvgRefusal::ExpansionTooLarge);
            }
            rest = rest[close + 1..].trim_start_matches(|c: char| c.is_ascii_whitespace());
        }
        Ok(())
    }
}

/// Style work `simplecss` spends rescanning CSS text of `len` bytes: at
/// worst a rejected token every other byte, each rescanning the text before
/// it, measured at about 0.17 ns per byte squared.
pub(super) fn rescans(len: usize) -> u64 {
    (len as u64).saturating_mul(len as u64) / 4
}

/// Style work the audit spends reading `text` with the same tokenizer, which
/// it does only for text that may declare a stroke width or dash list.
pub(super) fn tokenized(text: &str) -> u64 {
    if super::reference::contains_ignore_case(text, b"stroke-") {
        rescans(text.len())
    } else {
        0
    }
}

/// Refuses CSS text that could apply a filter or inherit a value: a copy's
/// clip path from whatever element it lands under, or a paint over the
/// presentation attribute the audit read as setting it aside.
pub(super) fn screen(text: &str) -> Result<(), SvgRefusal> {
    if super::reference::contains_ignore_case(text, b"filter")
        || super::reference::contains_ignore_case(text, b"inherit")
    {
        return Err(SvgRefusal::UnsupportedStyle);
    }
    Ok(())
}

/// Whether every function and string in a block closes inside it, so
/// `simplecss`, which skips a function to its `)` and a string to its quote,
/// ends the block at the same `}` this parser does.
fn closed(block: &str) -> bool {
    let bytes = block.as_bytes();
    let mut at = 0;
    while let Some(&byte) = bytes.get(at) {
        let close = match byte {
            b'(' => b')',
            b'\'' | b'"' => byte,
            _ => {
                at += 1;
                continue;
            }
        };
        match bytes[at + 1..].iter().position(|next| *next == close) {
            Some(length) => at += length + 2,
            None => return false,
        }
    }
    true
}

/// `*`, or an optional type followed by `.class` and `#id` parts, nothing
/// else, within [`MAX_SVG_SELECTOR_PARTS`] and [`MAX_SVG_SELECTOR_BYTES`].
fn compound(selector: &str, block: usize) -> Option<Rule<'_>> {
    let mut rule = Rule {
        tag: None,
        classes: Vec::new(),
        ids: Vec::new(),
        block,
    };
    if selector.len() > MAX_SVG_SELECTOR_BYTES {
        return None;
    }
    if selector == "*" {
        return Some(rule);
    }
    let (tag, mut rest) = ident(selector);
    rule.tag = (!tag.is_empty()).then_some(tag);
    while let Some(sigil) = rest.chars().next().filter(|c| matches!(c, '.' | '#')) {
        let (name, after) = ident(&rest[1..]);
        if name.is_empty() || rule.classes.len() + rule.ids.len() >= MAX_SVG_SELECTOR_PARTS {
            return None;
        }
        match sigil {
            '.' => rule.classes.push(name),
            _ => rule.ids.push(name),
        }
        rest = after;
    }
    if !rest.is_empty() {
        return None;
    }
    (rule.tag.is_some() || !rule.classes.is_empty() || !rule.ids.is_empty()).then_some(rule)
}

fn ident(text: &str) -> (&str, &str) {
    let end = text
        .find(|c: char| !(c.is_ascii_alphanumeric() || c == '-' || c == '_'))
        .unwrap_or(text.len());
    text.split_at(end)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sheet(css: &str) -> Result<usize, SvgRefusal> {
        let mut sheet = StyleSheet::default();
        sheet.parse(css)?;
        Ok(sheet.rules.len())
    }

    #[test]
    fn office_and_illustrator_rules_parse() {
        let css = "\n.MsftOfcThm_Accent1_Fill_v2 {\n fill:#4472C4; \n}\n.st0,.st1{fill:url(#SVGID_1_);}\nrect#a.b{stroke:none}\n*{opacity:1}\n.st2{font-family:'Myriad Pro';fill:rgb(1,2,3)}";
        assert_eq!(sheet(css), Ok(6));
    }

    #[test]
    fn combinators_comments_at_rules_and_filters_are_refused() {
        for css in [
            "g .x{fill:red}",
            "g>.x{fill:red}",
            "a+b{fill:red}",
            "[class]{fill:red}",
            "a:first-child{fill:red}",
            "@media print{a{fill:red}}",
            "/* c */a{fill:red}",
            ".a{fill:red",
            ".a{filter:blur(4px)}",
            ".a,{fill:red}",
            "*.a{fill:red}",
            "aé{fill:red}",
            ".a\u{e9}{fill:red}",
            ".a{font-family:'}';fill:url(#g)}",
            ".a{fill:f(}.b{fill:red)}",
            ".a{font-family:\"x\\\"}\"}",
        ] {
            assert_eq!(sheet(css), Err(SvgRefusal::UnsupportedStyle), "{css}");
        }
        assert_eq!(
            sheet(".a{fill:URL(https://example.invalid/p.svg#g)}"),
            Err(SvgRefusal::ExternalReference)
        );
    }

    #[test]
    fn a_selector_is_held_to_its_parts_and_bytes() {
        let parts = |count: usize| format!("{}{{fill:red}}", ".a".repeat(count));
        assert_eq!(sheet(&parts(MAX_SVG_SELECTOR_PARTS)), Ok(1));
        assert_eq!(
            sheet(&parts(MAX_SVG_SELECTOR_PARTS + 1)),
            Err(SvgRefusal::UnsupportedStyle)
        );
        let long = format!(".{}{{fill:red}}", "a".repeat(MAX_SVG_SELECTOR_BYTES));
        assert_eq!(sheet(&long), Err(SvgRefusal::UnsupportedStyle));
    }

    #[test]
    fn a_block_copied_per_selector_counts_every_copy() {
        let selectors: Vec<String> = (0..1_000).map(|index| format!(".s{index}")).collect();
        let css = format!("{}{{{}}}", selectors.join(","), "a:b;".repeat(1_000));
        assert_eq!(sheet(&css), Err(SvgRefusal::ExpansionTooLarge));
    }
}
