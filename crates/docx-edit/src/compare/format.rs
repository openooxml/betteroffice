//! Effective run formatting read from source XML, independent of the editing session.

use std::collections::HashMap;

use docx_parse::{
    ParseBudget, ParseLimits, TextFormatting, Theme, XmlElement, merge_text_formatting,
    parse_run_properties,
};

use super::package::parse;
use super::xml::{Namespaces, run_text};

/// Run formatting as runs of equal formatting: `(units, readings)`, where the readings are
/// [`RunStyles::effective`]'s.
pub(crate) type Formats = Vec<(u32, [TextFormatting; 2])>;

/// Whether each unit's own run properties set complex-script bold and italic, as runs of equal
/// values: `(units, [bold, italic])`.
pub(crate) type ComplexScript = Vec<(u32, [bool; 2])>;

/// Which revisions a paragraph read keeps.
#[derive(Clone, Copy, Eq, PartialEq)]
pub(crate) enum View {
    /// Tracked changes are not expected.
    Plain,
    /// Deleted runs kept, inserted runs skipped.
    Original,
    /// Inserted runs kept, deleted runs skipped.
    Accepted,
}

type Toggle = (
    fn(&TextFormatting) -> Option<bool>,
    fn(&mut TextFormatting, Option<bool>),
);

/// The properties whose style values toggle rather than override.
const TOGGLES: [Toggle; 13] = [
    (|f| f.bold, |f, v| f.bold = v),
    (|f| f.bold_cs, |f, v| f.bold_cs = v),
    (|f| f.italic, |f, v| f.italic = v),
    (|f| f.italic_cs, |f, v| f.italic_cs = v),
    (|f| f.all_caps, |f, v| f.all_caps = v),
    (|f| f.small_caps, |f, v| f.small_caps = v),
    (|f| f.strike, |f, v| f.strike = v),
    (|f| f.double_strike, |f, v| f.double_strike = v),
    (|f| f.emboss, |f, v| f.emboss = v),
    (|f| f.imprint, |f, v| f.imprint = v),
    (|f| f.outline, |f, v| f.outline = v),
    (|f| f.shadow, |f, v| f.shadow = v),
    (|f| f.hidden, |f, v| f.hidden = v),
];

/// The styles and theme a package's runs resolve against.
pub(crate) struct RunStyles {
    defaults: Option<TextFormatting>,
    chains: HashMap<String, Chain>,
    default_paragraph: Option<String>,
    theme: Theme,
}

/// A style's run properties along its `basedOn` chain: other properties override root first,
/// toggles combine by exclusive or across every level that sets them.
struct Chain {
    merged: Option<TextFormatting>,
    toggles: [Option<bool>; TOGGLES.len()],
}

/// A style as its own definition states it, before inheritance.
struct Level {
    based_on: Option<String>,
    r_pr: Option<TextFormatting>,
}

const MAX_CHAIN: usize = 64;

fn toggled(values: impl IntoIterator<Item = Option<bool>>) -> Option<bool> {
    values.into_iter().fold(None, |state, value| match value {
        Some(value) => Some(state.unwrap_or(false) ^ value),
        None => state,
    })
}

impl RunStyles {
    pub(crate) fn read(parts: &[(String, Vec<u8>)]) -> Result<Self, String> {
        let part = |name: &str| {
            parts
                .iter()
                .find(|(path, _)| path == name)
                .map(|(_, bytes)| bytes.as_slice())
        };
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let failed = |error: docx_parse::ParseError| error.to_string();
        let settings =
            docx_parse::parse_settings(part("word/settings.xml"), "word/settings.xml", &mut budget)
                .map_err(failed)?;
        let mut theme = docx_parse::parse_theme(
            part("word/theme/theme1.xml"),
            "word/theme/theme1.xml",
            &mut budget,
        )
        .map_err(failed)?;
        docx_parse::apply_theme_font_lang(&mut theme, settings.theme_font_lang.as_ref());
        let mut levels = HashMap::new();
        let mut defaults = None;
        let mut default_paragraph = None;
        if let Some(xml) = part("word/styles.xml").filter(|xml| !xml.is_empty()) {
            let document =
                parse(xml, "word/styles.xml").ok_or("word/styles.xml is not readable")?;
            let children = document
                .root()
                .into_iter()
                .flat_map(XmlElement::child_elements);
            for element in children {
                match element.local_name() {
                    "docDefaults" => {
                        defaults = parse_run_properties(
                            element
                                .child("w", "rPrDefault")
                                .and_then(|default| default.child("w", "rPr")),
                            Some(&theme),
                        );
                    }
                    "style" => {
                        let Some(id) = element.attribute(Some("w"), "styleId") else {
                            continue;
                        };
                        let value = |name| {
                            element
                                .child("w", name)
                                .and_then(|child| child.attribute(Some("w"), "val"))
                        };
                        if default_paragraph.is_none()
                            && element.attribute(Some("w"), "type") == Some("paragraph")
                            && matches!(element.attribute(Some("w"), "default"), Some("1" | "true"))
                        {
                            default_paragraph = Some(id.to_owned());
                        }
                        levels.insert(
                            id.to_owned(),
                            Level {
                                based_on: value("basedOn").map(str::to_owned),
                                r_pr: parse_run_properties(element.child("w", "rPr"), Some(&theme)),
                            },
                        );
                    }
                    _ => {}
                }
            }
        }
        let mut chains = HashMap::with_capacity(levels.len());
        for id in levels.keys() {
            let mut chain: Vec<(&str, &Level)> = Vec::new();
            let mut cursor = Some(id.as_str());
            while let Some((id, level)) = cursor.and_then(|id| levels.get_key_value(id)) {
                if chain.len() == MAX_CHAIN || chain.iter().any(|(seen, _)| *seen == id) {
                    return Err(format!("style {id:?} has an unbounded basedOn chain"));
                }
                chain.push((id, level));
                cursor = level.based_on.as_deref();
            }
            let merged = chain.iter().rev().fold(None, |merged, (_, level)| {
                merge_text_formatting(merged.as_ref(), level.r_pr.as_ref())
            });
            let toggles = TOGGLES.map(|(get, _)| {
                toggled(
                    chain
                        .iter()
                        .rev()
                        .map(|(_, level)| level.r_pr.as_ref().and_then(get)),
                )
            });
            chains.insert(id.clone(), Chain { merged, toggles });
        }
        Ok(Self {
            defaults,
            chains,
            default_paragraph,
            theme,
        })
    }

    /// The effective formatting of a run with `direct` properties in a paragraph of `style`, read
    /// twice: with style toggles combined across every level as the specification defines, and
    /// with every level overriding the last, as flattening readers resolve them.
    pub(crate) fn effective(
        &self,
        style: Option<&str>,
        direct: Option<&TextFormatting>,
    ) -> [TextFormatting; 2] {
        let chain = |id: Option<&str>| id.and_then(|id| self.chains.get(id));
        let paragraph = chain(style).or_else(|| chain(self.default_paragraph.as_deref()));
        let character = chain(direct.and_then(|direct| direct.style_id.as_deref()));
        let styled = merge_text_formatting(
            self.defaults.as_ref(),
            paragraph.and_then(|chain| chain.merged.as_ref()),
        );
        let mut styled = merge_text_formatting(
            styled.as_ref(),
            character.and_then(|chain| chain.merged.as_ref()),
        )
        .unwrap_or_default();
        let flattened = merge_text_formatting(Some(&styled), direct).unwrap_or_default();
        for (index, (get, set)) in TOGGLES.into_iter().enumerate() {
            let level = |chain: Option<&Chain>| chain.and_then(|chain| chain.toggles[index]);
            set(
                &mut styled,
                toggled([
                    self.defaults.as_ref().and_then(get),
                    level(paragraph),
                    level(character),
                ]),
            );
        }
        [
            merge_text_formatting(Some(&styled), direct).unwrap_or_default(),
            flattened,
        ]
    }

    /// The effective formatting of `paragraph`'s text units in `view` and their own complex-script
    /// bold and italic, or `None` when it holds content other than runs of text, tabs, special
    /// hyphens and line breaks.
    pub(crate) fn paragraph(
        &self,
        paragraph: &XmlElement,
        ns: &mut Namespaces,
        view: View,
    ) -> Option<(Formats, ComplexScript)> {
        let mark = ns.enter(paragraph);
        let mut read = Read {
            styles: self,
            style: None,
            formats: Vec::new(),
            complex_script: Vec::new(),
        };
        let complete = paragraph
            .child_elements()
            .all(|child| read.child(child, ns, view));
        ns.leave(mark);
        complete.then_some((read.formats, read.complex_script))
    }
}

struct Read<'a> {
    styles: &'a RunStyles,
    style: Option<String>,
    formats: Formats,
    complex_script: ComplexScript,
}

fn push<T: PartialEq>(runs: &mut Vec<(u32, T)>, units: u32, value: T) {
    match runs.last_mut() {
        Some((count, last)) if *last == value => *count += units,
        _ => runs.push((units, value)),
    }
}

impl Read<'_> {
    fn child(&mut self, child: &XmlElement, ns: &mut Namespaces, view: View) -> bool {
        let mark = ns.enter(child);
        let complete = match ns.w_local(child) {
            Some("pPr") => {
                self.style = child
                    .child_elements()
                    .find(|property| ns.w_local(property) == Some("pStyle"))
                    .and_then(|property| property.attribute(Some("w"), "val"))
                    .map(str::to_owned);
                true
            }
            Some("r") => self.run(child, ns),
            Some("ins") if view == View::Accepted => {
                child.child_elements().all(|run| self.wrapped(run, ns))
            }
            Some("del") if view == View::Original => {
                child.child_elements().all(|run| self.wrapped(run, ns))
            }
            Some("ins") if view == View::Original => true,
            Some("del") if view == View::Accepted => true,
            Some("proofErr") => true,
            _ => false,
        };
        ns.leave(mark);
        complete
    }

    fn wrapped(&mut self, run: &XmlElement, ns: &mut Namespaces) -> bool {
        let mark = ns.enter(run);
        let complete = ns.w_local(run) == Some("r") && self.run(run, ns);
        ns.leave(mark);
        complete
    }

    fn run(&mut self, run: &XmlElement, ns: &mut Namespaces) -> bool {
        let mut direct = None;
        let mut units = 0u32;
        for child in run.child_elements() {
            let mark = ns.enter(child);
            let local = ns.w_local(child);
            ns.leave(mark);
            match local {
                Some("rPr") => direct = parse_run_properties(Some(child), Some(&self.styles.theme)),
                Some("t" | "delText") => {
                    units += run_text(child).encode_utf16().count() as u32;
                }
                Some("tab" | "softHyphen" | "noBreakHyphen" | "br") => units += 1,
                Some("lastRenderedPageBreak") => {}
                _ => return false,
            }
        }
        if units == 0 {
            return true;
        }
        let effective = self
            .styles
            .effective(self.style.as_deref(), direct.as_ref());
        push(&mut self.formats, units, effective);
        let own = direct.as_ref().map_or([false; 2], |direct| {
            [direct.bold_cs == Some(true), direct.italic_cs == Some(true)]
        });
        push(&mut self.complex_script, units, own);
        true
    }
}
