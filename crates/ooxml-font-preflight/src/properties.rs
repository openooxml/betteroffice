use std::rc::Rc;

use crate::map::Map;
use crate::xml::Node;

#[derive(Clone, Default)]
pub(crate) struct Props {
    pub(crate) fonts: [Option<Rc<str>>; 4],
    pub(crate) family_override: Option<Rc<str>>,
    pub(crate) hint: Option<Rc<str>>,
    pub(crate) bold: Option<bool>,
    pub(crate) italic: Option<bool>,
    pub(crate) bold_cs: Option<bool>,
    pub(crate) italic_cs: Option<bool>,
    pub(crate) cs: Option<bool>,
    pub(crate) rtl: Option<bool>,
    pub(crate) hidden: Option<bool>,
    pub(crate) style: Option<Rc<str>>,
}

pub(crate) fn on_off(value: &str) -> Option<bool> {
    let value = value.trim();
    if value == "1" || value.eq_ignore_ascii_case("true") || value.eq_ignore_ascii_case("on") {
        Some(true)
    } else if value == "0"
        || value.eq_ignore_ascii_case("false")
        || value.eq_ignore_ascii_case("off")
    {
        Some(false)
    } else {
        None
    }
}

impl Props {
    pub(crate) fn allocation_cost(&self) -> usize {
        std::mem::size_of::<Self>()
            + self
                .fonts
                .iter()
                .chain([&self.family_override, &self.hint, &self.style])
                .flatten()
                .map(|value| value.len())
                .sum::<usize>()
    }

    pub(crate) fn apply_defaults(&mut self, mark: &Self) {
        if self.fonts.iter().all(Option::is_none) {
            self.fonts.clone_from(&mark.fonts);
        }
        self.bold_cs = self.bold_cs.or(mark.bold_cs);
        self.italic_cs = self.italic_cs.or(mark.italic_cs);
        self.cs = self.cs.or(mark.cs);
    }

    #[inline(never)]
    pub(crate) fn parse(node: Option<&Node>, theme: &Theme) -> Self {
        let Some(node) = node else {
            return Self::default();
        };
        let mut props = Self::default();
        if let Some(fonts) = node.child("w:rFonts") {
            for (index, (explicit, reference)) in [
                ("ascii", "asciiTheme"),
                ("hAnsi", "hAnsiTheme"),
                ("eastAsia", "eastAsiaTheme"),
                ("cs", "cstheme"),
            ]
            .into_iter()
            .enumerate()
            {
                props.fonts[index] = fonts.attr(explicit).map(Rc::from);
                if props.fonts[index].as_deref().is_none_or(str::is_empty)
                    && let Some(reference) = fonts.attr(reference).filter(|value| !value.is_empty())
                {
                    props.fonts[index] = Some(theme.resolve(reference));
                }
            }
            props.hint = fonts
                .attr("hint")
                .filter(|hint| matches!(*hint, "default" | "eastAsia" | "cs"))
                .map(Rc::from);
        }
        let flag = |name| {
            node.child(name)
                .map(|node| on_off(node.attr("val").unwrap_or("true")) != Some(false))
        };
        props.bold = flag("w:b");
        props.italic = flag("w:i");
        props.bold_cs = flag("w:bCs");
        props.italic_cs = flag("w:iCs");
        props.cs = flag("w:cs");
        props.rtl = flag("w:rtl");
        props.hidden = flag("w:vanish");
        props.style = node.val("w:rStyle").map(Rc::from);
        props
    }

    #[inline(never)]
    pub(crate) fn merge(&mut self, source: &Self, toggle: bool) {
        for (target, source) in self.fonts.iter_mut().zip(&source.fonts) {
            if source.is_some() {
                target.clone_from(source);
            }
        }
        for (target, source) in [
            (&mut self.bold, source.bold),
            (&mut self.italic, source.italic),
            (&mut self.bold_cs, source.bold_cs),
            (&mut self.italic_cs, source.italic_cs),
            (&mut self.hidden, source.hidden),
        ] {
            if let Some(value) = source {
                *target = Some(if toggle {
                    target.unwrap_or(false) ^ value
                } else {
                    value
                });
            }
        }
        for (target, source) in [(&mut self.cs, source.cs), (&mut self.rtl, source.rtl)] {
            if source.is_some() {
                *target = source;
            }
        }
        if source.hint.is_some() {
            self.hint.clone_from(&source.hint);
        }
        if source.style.is_some() {
            self.style.clone_from(&source.style);
        }
    }

    pub(crate) fn family(&self) -> Option<&str> {
        if let Some(family) = self.family_override.as_deref() {
            return Some(family);
        }
        if self.rtl == Some(true) && self.fonts[3].is_some() {
            return self.fonts[3].as_deref();
        }
        self.fonts.iter().find_map(Option::as_deref)
    }

    pub(crate) fn latin(&self) -> Option<&str> {
        self.fonts[0].as_deref().or(self.fonts[1].as_deref())
    }
}

pub(crate) struct Theme {
    fonts: [[Rc<str>; 3]; 2],
}

impl Theme {
    pub(crate) fn parse(theme: Option<&Node>, settings: Option<&Node>) -> Self {
        let mut result = Self {
            fonts: [
                ["Calibri Light".into(), "".into(), "".into()],
                ["Calibri".into(), "".into(), "".into()],
            ],
        };
        let language = settings.and_then(|node| node.child("w:themeFontLang"));
        for (index, name) in ["a:majorFont", "a:minorFont"].into_iter().enumerate() {
            let Some(font) =
                theme.and_then(|node| node.path(&["a:themeElements", "a:fontScheme", name]))
            else {
                continue;
            };
            for (slot, name) in ["a:latin", "a:ea", "a:cs"].into_iter().enumerate() {
                result.fonts[index][slot] = Rc::from(
                    font.child(name)
                        .and_then(|node| node.attr("typeface"))
                        .unwrap_or_default(),
                );
            }
            for (slot, key) in [(1, "eastAsia"), (2, "bidi")] {
                let script = language
                    .and_then(|node| node.attr(key))
                    .and_then(language_script);
                if result.fonts[index][slot].is_empty()
                    && let Some(script) = script
                    && let Some(face) = font
                        .children
                        .iter()
                        .find(|node| node.name == "a:font" && node.attr("script") == Some(script))
                {
                    result.fonts[index][slot] = Rc::from(face.attr("typeface").unwrap_or_default());
                }
            }
        }
        result
    }

    pub(crate) fn resolve(&self, reference: &str) -> Rc<str> {
        let lower = reference.trim_start_matches('+').to_ascii_lowercase();
        let (major, script, drawing) = match lower.split_once('-') {
            Some((slot @ ("mj" | "mn"), script)) => (slot == "mj", script, true),
            _ => (
                lower.starts_with("major"),
                lower
                    .strip_prefix("major")
                    .or_else(|| lower.strip_prefix("minor"))
                    .unwrap_or(&lower),
                false,
            ),
        };
        let index = usize::from(!major);
        let slot = match script {
            "ea" | "eastasia" => 1,
            "cs" | "bidi" => 2,
            _ => 0,
        };
        let family = &self.fonts[index][slot];
        if family.is_empty() && (slot == 0 || drawing) {
            if self.fonts[index][0].is_empty() {
                if major { "Calibri Light" } else { "Calibri" }.into()
            } else {
                self.fonts[index][0].clone()
            }
        } else {
            family.clone()
        }
    }
}

fn language_script(language: &str) -> Option<&'static str> {
    let lower = language.to_ascii_lowercase();
    if lower.starts_with("ja") {
        Some("Jpan")
    } else if lower.starts_with("ko") {
        Some("Hang")
    } else if lower.starts_with("zh") {
        Some(
            if lower.contains("hant")
                || lower
                    .split('-')
                    .any(|part| matches!(part, "tw" | "hk" | "mo"))
            {
                "Hant"
            } else {
                "Hans"
            },
        )
    } else if lower.starts_with("ar") {
        Some("Arab")
    } else if lower.starts_with("he") || lower.starts_with("iw") {
        Some("Hebr")
    } else if lower.starts_with("th") {
        Some("Thai")
    } else if ["hi", "mr", "ne"]
        .iter()
        .any(|prefix| lower.starts_with(prefix))
    {
        Some("Deva")
    } else {
        None
    }
}

#[derive(Clone, Default)]
pub(crate) struct Style {
    pub(crate) run: Props,
    pub(crate) num: Option<(Option<String>, Option<String>)>,
    pub(crate) conditionals: Map<Props>,
    pub(crate) look: Option<Look>,
    pub(crate) row_band: Option<usize>,
    pub(crate) col_band: Option<usize>,
    based_on: Option<String>,
}

pub(crate) struct Styles {
    pub(crate) defaults: Props,
    pub(crate) map: Map<Style>,
    defaults_by_type: Map<String>,
}

impl Styles {
    pub(crate) fn parse(root: Option<&Node>, theme: &Theme, complete: &mut bool) -> Self {
        let mut result = Self {
            defaults: Props::default(),
            map: Map::new(),
            defaults_by_type: Map::new(),
        };
        let Some(root) = root else {
            *complete = false;
            return result;
        };
        result.defaults = Props::parse(
            root.path(&["w:docDefaults", "w:rPrDefault", "w:rPr"]),
            theme,
        );
        for (index, node) in root
            .children
            .iter()
            .filter(|node| node.name == "w:style")
            .enumerate()
        {
            if index == 2048 {
                *complete = false;
                break;
            }
            let Some(id) = node.attr("styleId") else {
                continue;
            };
            let kind = node.attr("type").unwrap_or("paragraph");
            if node.attr("default").and_then(on_off) == Some(true)
                && !result.defaults_by_type.contains_key(kind)
            {
                result
                    .defaults_by_type
                    .insert(kind.to_owned(), id.to_owned());
            }
            let mut conditionals = Map::new();
            for conditional in node
                .children
                .iter()
                .filter(|node| node.name == "w:tblStylePr")
            {
                if let Some(kind) = conditional.attr("type").filter(|kind| {
                    matches!(
                        *kind,
                        "wholeTable"
                            | "firstRow"
                            | "lastRow"
                            | "firstCol"
                            | "lastCol"
                            | "band1Horz"
                            | "band2Horz"
                            | "band1Vert"
                            | "band2Vert"
                            | "nwCell"
                            | "neCell"
                            | "swCell"
                            | "seCell"
                    )
                }) {
                    conditionals.insert(
                        kind.to_owned(),
                        Props::parse(conditional.child("w:rPr"), theme),
                    );
                }
            }
            result.map.insert(
                id.to_owned(),
                Style {
                    run: Props::parse(node.child("w:rPr"), theme),
                    num: num_pr(node.child("w:pPr")),
                    conditionals,
                    look: node.path(&["w:tblPr", "w:tblLook"]).map(Look::parse),
                    row_band: node
                        .path(&["w:tblPr"])
                        .and_then(|node| number(node.val("w:tblStyleRowBandSize"))),
                    col_band: node
                        .path(&["w:tblPr"])
                        .and_then(|node| number(node.val("w:tblStyleColBandSize"))),
                    based_on: node.val("w:basedOn").map(str::to_owned),
                },
            );
        }
        if result.map.contains_key("Normal") && !result.defaults_by_type.contains_key("paragraph") {
            result
                .defaults_by_type
                .insert("paragraph".into(), "Normal".into());
        }
        let ids: Vec<_> = result.map.keys().cloned().collect();
        for id in ids {
            let resolved = result.resolve(&id, complete);
            result.map.insert(id, resolved);
        }
        result
    }

    fn resolve(&self, id: &str, complete: &mut bool) -> Style {
        let mut chain = Vec::new();
        let mut current = Some(id);
        while let Some(id) = current {
            if chain.len() == 128 || chain.contains(&id) {
                *complete = false;
                break;
            }
            let Some(style) = self.map.get(id) else {
                *complete = false;
                break;
            };
            chain.push(id);
            current = style.based_on.as_deref();
        }
        let mut result = Style::default();
        for id in chain.into_iter().rev() {
            let Some(style) = self.map.get(id) else {
                continue;
            };
            result.run.merge(&style.run, false);
            if let Some((id, level)) = &style.num {
                let previous = result.num.get_or_insert((None, None));
                if id.is_some() {
                    previous.0.clone_from(id);
                }
                if level.is_some() {
                    previous.1.clone_from(level);
                }
            }
            for (kind, props) in style.conditionals.iter() {
                if let Some(current) = result.conditionals.get_mut(kind) {
                    current.merge(props, false);
                } else {
                    result.conditionals.insert(kind.clone(), props.clone());
                }
            }
            if style.look.is_some() {
                result.look = style.look.clone();
            }
            result.row_band = style.row_band.or(result.row_band);
            result.col_band = style.col_band.or(result.col_band);
        }
        result
    }

    pub(crate) fn get(&self, id: Option<&str>, kind: &str, complete: &mut bool) -> Option<&Style> {
        if let Some(id) = id {
            if let Some(style) = self.map.get(id) {
                return Some(style);
            }
            *complete = false;
            if kind != "paragraph" {
                return None;
            }
        }
        self.defaults_by_type
            .get(kind)
            .and_then(|id| self.map.get(id))
    }
}

pub(crate) fn num_pr(node: Option<&Node>) -> Option<(Option<String>, Option<String>)> {
    let num = node?.child("w:numPr")?;
    Some((
        num.val("w:numId").map(str::to_owned),
        num.val("w:ilvl").map(str::to_owned),
    ))
}

pub(crate) fn number(value: Option<&str>) -> Option<usize> {
    value?.parse().ok()
}

#[derive(Clone)]
pub(crate) struct Look {
    pub(crate) flags: [bool; 6],
}

impl Default for Look {
    fn default() -> Self {
        Self {
            flags: [true, false, true, false, false, true],
        }
    }
}

impl Look {
    pub(crate) fn parse(node: &Node) -> Self {
        let mask = node
            .attr("val")
            .and_then(|value| u16::from_str_radix(value, 16).ok())
            .unwrap_or(0);
        let mut flags = [false; 6];
        for (index, key) in [
            "firstRow",
            "lastRow",
            "firstColumn",
            "lastColumn",
            "noHBand",
            "noVBand",
        ]
        .into_iter()
        .enumerate()
        {
            flags[index] = node
                .attr(key)
                .and_then(on_off)
                .unwrap_or(mask & (0x20 << index) != 0);
        }
        Self { flags }
    }
}
