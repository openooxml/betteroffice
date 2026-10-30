use std::borrow::Cow;
use std::collections::BTreeMap;
use std::rc::Rc;

use crate::PreflightFace;
use crate::properties::{Props, Styles, Theme, num_pr, number};
use crate::xml::Node;

pub(crate) const MAX_FACES: usize = 512;
pub(crate) const MAX_DERIVED_BYTES: usize = 4 * 1024 * 1024;

type Run<'a> = (Rc<Props>, Option<Cow<'a, str>>);

#[derive(Default)]
struct FaceState {
    named: Option<PreflightFace>,
    implicit: Option<PreflightFace>,
    reached: bool,
}

pub(crate) struct Collector {
    faces: BTreeMap<(String, bool, bool), FaceState>,
    remaining: usize,
    pub(crate) limited: bool,
}

impl Default for Collector {
    fn default() -> Self {
        Self {
            faces: BTreeMap::new(),
            remaining: MAX_DERIVED_BYTES,
            limited: false,
        }
    }
}

fn fold(ch: char) -> char {
    if ('Ａ'..='Ｚ').contains(&ch) {
        char::from_u32(ch as u32 + 32).unwrap_or(ch)
    } else {
        ch.to_ascii_lowercase()
    }
}

fn merge_scripts(face: &mut PreflightFace, scripts: &[String]) {
    for script in scripts {
        if !face.scripts.contains(script) {
            face.scripts.push(script.clone());
        }
    }
}

impl Collector {
    fn charge(&mut self, bytes: usize) -> bool {
        if let Some(remaining) = self.remaining.checked_sub(bytes) {
            self.remaining = remaining;
            true
        } else {
            self.limited = true;
            false
        }
    }

    fn face(&mut self, family: &str, bold: bool, italic: bool) -> Option<&mut FaceState> {
        if family.len() > crate::xml::MAX_VALUE_BYTES {
            self.limited = true;
            return None;
        }
        let key = (family.chars().map(fold).collect(), bold, italic);
        if self.faces.len() == MAX_FACES && !self.faces.contains_key(&key) {
            self.limited = true;
            return None;
        }
        Some(self.faces.entry(key).or_default())
    }

    fn name(&mut self, family: &str, bold: bool, italic: bool, scripts: &[String], reached: bool) {
        let Some(state) = self.face(family, bold, italic) else {
            return;
        };
        let face = state.named.get_or_insert_with(|| PreflightFace {
            family: family.to_owned(),
            bold,
            italic,
            scripts: Vec::new(),
        });
        merge_scripts(face, scripts);
        state.reached |= reached;
    }

    fn reach(&mut self, family: &str, bold: bool, italic: bool, scripts: &[String]) {
        let Some(state) = self.face(family, bold, italic) else {
            return;
        };
        let face = state.implicit.get_or_insert_with(|| PreflightFace {
            family: family.to_owned(),
            bold,
            italic,
            scripts: Vec::new(),
        });
        merge_scripts(face, scripts);
    }

    pub(crate) fn finish(self) -> Vec<PreflightFace> {
        self.faces
            .into_values()
            .filter_map(|state| {
                if state.reached || state.implicit.is_some() {
                    state.named.or(state.implicit)
                } else {
                    None
                }
            })
            .collect()
    }

    #[inline(never)]
    fn run(&mut self, props: &Props, text: Option<&str>, default: &str, scripts: &[String]) {
        let bold = props.bold.unwrap_or(false);
        let italic = props.italic.unwrap_or(false);
        let cs_bold = props.bold_cs.unwrap_or(bold);
        let cs_italic = props.italic_cs.unwrap_or(italic);
        let (east_asia, cs) = slot_use(
            text.unwrap_or_default(),
            props.cs == Some(true),
            props.hint.as_deref(),
        );
        let cs = cs || props.cs == Some(true);
        let mut name = |family: &str, bold, italic, reached| {
            self.name(family, bold, italic, scripts, reached);
            if text.is_some() {
                self.name(family, false, false, scripts, reached);
            }
        };
        name(props.family().unwrap_or(default), bold, italic, true);
        for (slot, bold, italic, reached) in [
            (0, bold, italic, true),
            (1, bold, italic, true),
            (2, bold, italic, east_asia),
            (3, cs_bold, cs_italic, cs),
        ] {
            if let Some(family) = &props.fonts[slot] {
                name(family, bold, italic, reached);
            }
        }
        let run_family = props.family().unwrap_or("Calibri");
        let text_unnamed = props.fonts[1]
            .as_deref()
            .or(props.fonts[0].as_deref())
            .unwrap_or(run_family);
        let unnamed = if text.is_some() {
            text_unnamed
        } else {
            run_family
        };
        self.reach(unnamed, bold, italic, scripts);
        self.reach(run_family, bold, italic, scripts);
        if east_asia {
            self.reach(
                props.fonts[2].as_deref().unwrap_or(unnamed),
                bold,
                italic,
                scripts,
            );
        }
        if cs {
            self.reach(
                props.fonts[3].as_deref().unwrap_or(text_unnamed),
                cs_bold,
                cs_italic,
                scripts,
            );
        }
    }
}

fn slot_use(text: &str, complex: bool, hint: Option<&str>) -> (bool, bool) {
    let mut slot = 1;
    let mut east_asia = false;
    let mut cs = false;
    for ch in text.chars() {
        let point = ch as u32;
        slot = if complex
            || matches!(point, 0x0590..=0x08ff | 0x0900..=0x0dff | 0xfb1d..=0xfdff | 0xfe70..=0xfeff)
        {
            3
        } else if matches!(point, 0x1100..=0x11ff | 0x2e80..=0x30ff | 0x3130..=0x318f | 0x31a0..=0x31ff | 0x3400..=0x4dbf | 0x4e00..=0x9fff | 0xa960..=0xa97f | 0xac00..=0xd7ff | 0xf900..=0xfaff | 0xff00..=0xffef)
        {
            2
        } else if matches!(point, 0x0300..=0x036f | 0x1ab0..=0x1aff | 0x1dc0..=0x1dff | 0x20d0..=0x20ff | 0xfe20..=0xfe2f)
        {
            slot
        } else if ch.is_ascii() {
            0
        } else if hint == Some("eastAsia") {
            2
        } else {
            1
        };
        east_asia |= slot == 2;
        cs |= slot == 3;
    }
    (east_asia, cs)
}

fn scripts(runs: &[Run<'_>]) -> Vec<String> {
    let mut flags = [false; 5];
    for (_, text) in runs {
        for ch in text.as_deref().unwrap_or_default().chars() {
            match ch as u32 {
                0x0590..=0x05ff | 0xfb1d..=0xfb4f => flags[4] = true,
                0x0600..=0x06ff
                | 0x0750..=0x077f
                | 0x0870..=0x08ff
                | 0xfb50..=0xfdff
                | 0xfe70..=0xfeff => flags[3] = true,
                0x1100..=0x11ff
                | 0x3130..=0x318f
                | 0xa960..=0xa97f
                | 0xac00..=0xd7ff
                | 0xffa0..=0xffdc => flags[1] = true,
                0x3040..=0x30ff | 0x31f0..=0x31ff | 0xff66..=0xff9f => flags[0] = true,
                0x3000..=0x303f
                | 0x3400..=0x4dbf
                | 0x4e00..=0x9fff
                | 0xf900..=0xfaff
                | 0xfe30..=0xfe4f
                | 0xff00..=0xff65
                | 0x20000..=0x3ffff => flags[2] = true,
                _ => {}
            }
        }
    }
    flags[2] &= !flags[0] && !flags[1];
    ["cjk-jp", "cjk-kr", "cjk-sc", "arabic", "hebrew"]
        .into_iter()
        .zip(flags)
        .filter(|(_, active)| *active)
        .map(|(key, _)| key.to_owned())
        .collect()
}

pub(crate) struct Scan<'a> {
    pub(crate) styles: &'a Styles,
    pub(crate) theme: &'a Theme,
    pub(crate) numbering: Option<&'a Node>,
    pub(crate) complete: &'a mut bool,
    pub(crate) collector: &'a mut Collector,
    pub(crate) note_label: bool,
    pub(crate) shape: bool,
}

impl Scan<'_> {
    pub(crate) fn walk(&mut self, node: &Node, table: &Props) {
        if self.collector.limited {
            return;
        }
        match node.name {
            "w:p" | "m:oMathPara" => self.paragraph(node, table),
            "w:body" | "w:hdr" | "w:ftr" | "w:footnote" | "w:endnote"
                if node.children.iter().all(|child| child.name == "w:sectPr") =>
            {
                self.paragraph(node, table)
            }
            "w:tbl" => {
                self.note_label = false;
                if self.shape {
                    for child in &node.children {
                        self.walk(child, table);
                    }
                } else {
                    self.table(node);
                }
            }
            "w:txbxContent" => {
                let previous = self.shape;
                self.shape = true;
                for child in &node.children {
                    self.walk(child, &Props::default());
                }
                self.shape = previous;
            }
            "w:body" => {
                for child in &node.children {
                    self.walk(child, table);
                }
                if node
                    .children
                    .iter()
                    .rev()
                    .find(|child| matches!(child.name, "w:p" | "w:tbl" | "w:sdt"))
                    .is_some_and(|child| child.name != "w:p")
                {
                    self.collector.name("Calibri", false, false, &[], true);
                }
            }
            "AlternateContent" => {
                *self.complete = false;
                if let Some(choice) = node.children.first() {
                    self.walk(choice, table);
                }
            }
            "w:pPr" | "w:rPr" | "w:sectPr" => {}
            _ => {
                for child in &node.children {
                    self.walk(child, table);
                }
            }
        }
    }

    #[inline(never)]
    fn paragraph(&mut self, node: &Node, table: &Props) {
        let ppr = node.child("w:pPr");
        let style_id = ppr.and_then(|node| node.val("w:pStyle"));
        let style = if self.shape {
            None
        } else {
            self.styles.get(style_id, "paragraph", self.complete)
        };
        let mut base = if self.shape {
            Props::default()
        } else {
            self.styles.defaults.clone()
        };
        base.merge(table, true);
        if let Some(style) = style {
            base.merge(&style.run, true);
        }
        let mut mark = base.clone();
        if !self.shape
            && let Some(character) = self.styles.get(None, "character", self.complete)
        {
            mark.merge(&character.run, true);
        }
        let direct_mark = Props::parse(
            if self.shape {
                None
            } else {
                ppr.and_then(|node| node.child("w:rPr"))
            },
            self.theme,
        );
        if let Some(id) = direct_mark.style.as_deref() {
            let mut resolved = self.styles.defaults.clone();
            if let Some(character) = self.styles.get(None, "character", self.complete) {
                resolved.merge(&character.run, false);
            }
            if let Some(character) = self.styles.get(Some(id), "character", self.complete) {
                resolved.merge(&character.run, false);
            }
            resolved.merge(&direct_mark, false);
            mark.merge(&resolved, false);
        } else {
            mark.merge(&direct_mark, false);
        }
        let default = mark.latin().unwrap_or("Calibri").to_owned();
        let mut runs = Vec::new();
        self.runs(node, &base, &mark, &mut runs);
        if runs.is_empty() && mark.hidden == Some(true) {
            return;
        }
        let mut generated_label = self.note_label;
        if generated_label {
            self.note_label = false;
            let family = runs
                .iter()
                .find_map(|(props, text)| {
                    text.as_ref()
                        .and_then(|_| props.family())
                        .filter(|family| !family.is_empty())
                })
                .map(Rc::from);
            let props = Rc::new(Props {
                family_override: family,
                ..Props::default()
            });
            if self
                .collector
                .charge(props.allocation_cost() + std::mem::size_of::<Run<'_>>() + 3)
            {
                runs.insert(0, (props, Some("1  ".into())));
            } else {
                generated_label = false;
            }
        }
        let scripts = scripts(&runs);
        self.collector.name(&default, false, false, &scripts, true);
        for (index, (props, text)) in runs.iter().enumerate() {
            self.collector
                .run(props, text.as_deref(), &default, &scripts);
            if !(self.shape || generated_label && index == 0)
                && (props.bold_cs != mark.bold_cs
                    || props.italic_cs != mark.italic_cs
                    || props.cs != mark.cs)
            {
                *self.complete = false;
                let mut seeded = (**props).clone();
                seeded.bold_cs = mark.bold_cs;
                seeded.italic_cs = mark.italic_cs;
                seeded.cs = mark.cs;
                self.collector
                    .run(&seeded, text.as_deref(), &default, &scripts);
            }
            if self.collector.limited {
                break;
            }
        }
        let num = if self.shape {
            None
        } else {
            num_pr(ppr).or_else(|| style.and_then(|style| style.num.clone()))
        };
        if let Some((Some(id), level)) = num
            && id != "0"
        {
            let level = level.as_deref().unwrap_or("0");
            if let Some(level) = self.level(&id, level) {
                let props = Props::parse(level.child("w:rPr"), self.theme);
                if props.hidden != Some(true)
                    && level
                        .val("w:lvlText")
                        .is_some_and(|value| !value.is_empty())
                {
                    let family = props
                        .latin()
                        .filter(|family| !family.is_empty())
                        .or_else(|| {
                            runs.iter().find_map(|(props, text)| {
                                text.as_ref().and_then(|_| props.family())
                            })
                        })
                        .unwrap_or(&default);
                    self.collector.name(
                        family,
                        props.bold.or(mark.bold).unwrap_or(false),
                        props.italic.or(mark.italic).unwrap_or(false),
                        &scripts,
                        true,
                    );
                }
            } else {
                *self.complete = false;
            }
        }
        self.nested(node, table);
    }

    fn nested(&mut self, node: &Node, table: &Props) {
        for child in &node.children {
            if self.collector.limited {
                break;
            }
            if matches!(
                child.name,
                "w:p" | "w:tbl" | "w:txbxContent" | "AlternateContent"
            ) {
                self.walk(child, table);
            } else if child.name != "w:pPr" && child.name != "w:rPr" {
                self.nested(child, table);
            }
        }
    }

    #[inline(never)]
    fn push_run<'n>(
        &mut self,
        runs: &mut Vec<Run<'n>>,
        props: Rc<Props>,
        text: Option<Cow<'n, str>>,
    ) {
        let cost = props.allocation_cost()
            + std::mem::size_of::<Run<'_>>()
            + text.as_ref().map_or(0, |text| text.len());
        if self.collector.charge(cost) {
            runs.push((props, text));
        }
    }

    fn runs<'n>(&mut self, node: &'n Node, base: &Props, mark: &Props, runs: &mut Vec<Run<'n>>) {
        for child in &node.children {
            if self.collector.limited {
                break;
            }
            match child.name {
                "m:oMath" => {
                    let mut text = String::new();
                    math_text(child, &mut text);
                    if text.is_empty() {
                        text.push_str("[equation]");
                    }
                    let mut props = Props {
                        family_override: Some("Cambria Math".into()),
                        italic: Some(true),
                        ..Props::default()
                    };
                    props.apply_defaults(mark);
                    self.push_run(runs, Rc::new(props), Some(Cow::Owned(text)));
                }
                "w:r" => {
                    let direct = Props::parse(child.child("w:rPr"), self.theme);
                    let mut props = base.clone();
                    if !self.shape
                        && let Some(id) = direct.style.as_deref()
                        && let Some(style) = self.styles.get(Some(id), "character", self.complete)
                    {
                        props.merge(&style.run, true);
                    }
                    props.merge(&direct, false);
                    if self.shape {
                        props.rtl = None;
                    }
                    if props.hidden == Some(true) {
                        continue;
                    }
                    props.apply_defaults(mark);
                    let props = Rc::new(props);
                    for content in &child.children {
                        if self.collector.limited {
                            break;
                        }
                        match content.name {
                            "w:t" | "w:delText" => {
                                if !content.text.is_empty() {
                                    self.push_run(
                                        runs,
                                        props.clone(),
                                        Some(Cow::Borrowed(&content.text)),
                                    );
                                }
                            }
                            "w:tab" => self.push_run(runs, props.clone(), None),
                            "w:noBreakHyphen" => {
                                self.push_run(runs, props.clone(), Some(Cow::Borrowed("\u{2011}")))
                            }
                            "w:softHyphen" => {
                                self.push_run(runs, props.clone(), Some(Cow::Borrowed("\u{00ad}")))
                            }
                            "w:sym" if !self.shape => {
                                if let Some(ch) = content
                                    .attr("char")
                                    .and_then(|value| u32::from_str_radix(value, 16).ok())
                                    .and_then(char::from_u32)
                                {
                                    let mut symbol = (*props).clone();
                                    if let Some(font) =
                                        content.attr("font").filter(|value| !value.is_empty())
                                    {
                                        let family = Rc::<str>::from(font);
                                        symbol.fonts =
                                            std::array::from_fn(|_| Some(family.clone()));
                                    }
                                    self.push_run(
                                        runs,
                                        Rc::new(symbol),
                                        Some(Cow::Owned(ch.to_string())),
                                    );
                                }
                            }
                            "w:footnoteReference" | "w:endnoteReference" => {
                                self.push_run(runs, props.clone(), Some(Cow::Borrowed("1")))
                            }
                            "w:fldChar" | "w:instrText" | "w:drawing" | "w:pict" => {
                                *self.complete = false;
                            }
                            _ => {}
                        }
                    }
                }
                "w:pPr" | "w:tbl" | "w:p" | "w:rPr" => {}
                "w:sdt" | "w:fldSimple" | "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo" => {
                    *self.complete = false;
                    self.runs(child, base, mark, runs);
                }
                "AlternateContent" => {
                    *self.complete = false;
                    if let Some(choice) = child.children.first() {
                        self.runs(choice, base, mark, runs);
                    }
                }
                _ => self.runs(child, base, mark, runs),
            }
        }
    }

    fn level(&self, num_id: &str, ilvl: &str) -> Option<&Node> {
        if number(Some(ilvl)).is_none_or(|value| value > 8) {
            return None;
        }
        let numbering = self.numbering?;
        let instance = numbering
            .children
            .iter()
            .find(|node| node.name == "w:num" && node.attr("numId") == Some(num_id))?;
        if let Some(level) = instance
            .children
            .iter()
            .find(|node| node.name == "w:lvlOverride" && node.attr("ilvl") == Some(ilvl))
            .and_then(|node| node.child("w:lvl"))
        {
            return Some(level);
        }
        let abstract_id = instance.val("w:abstractNumId")?;
        let mut abstract_num = numbering.children.iter().find(|node| {
            node.name == "w:abstractNum" && node.attr("abstractNumId") == Some(abstract_id)
        })?;
        if !abstract_num
            .children
            .iter()
            .any(|node| node.name == "w:lvl")
            && let Some(link) = abstract_num.val("w:numStyleLink")
            && let Some(linked) = numbering
                .children
                .iter()
                .find(|node| node.name == "w:abstractNum" && node.val("w:styleLink") == Some(link))
        {
            abstract_num = linked;
        }
        abstract_num
            .children
            .iter()
            .find(|node| node.name == "w:lvl" && node.attr("ilvl") == Some(ilvl))
    }

    #[inline(never)]
    fn table(&mut self, node: &Node) {
        let tpr = node.child("w:tblPr");
        let style = self.styles.get(
            tpr.and_then(|node| node.val("w:tblStyle")),
            "table",
            self.complete,
        );
        let look = tpr
            .and_then(|node| node.child("w:tblLook"))
            .map(crate::properties::Look::parse)
            .or_else(|| style.and_then(|style| style.look.clone()))
            .unwrap_or_default();
        let rows = wrapped_children(node, "w:tr");
        let columns = node
            .child("w:tblGrid")
            .map(|grid| {
                grid.children
                    .iter()
                    .filter(|node| node.name == "w:gridCol")
                    .count()
            })
            .filter(|&count| count > 0)
            .unwrap_or_else(|| {
                rows.iter()
                    .map(|row| {
                        wrapped_children(row, "w:tc")
                            .into_iter()
                            .map(|cell| {
                                number(
                                    cell.path(&["w:tcPr"])
                                        .and_then(|node| node.val("w:gridSpan")),
                                )
                                .unwrap_or(1)
                            })
                            .fold(0usize, usize::saturating_add)
                    })
                    .max()
                    .unwrap_or(0)
            });
        for (row_index, row) in rows.iter().enumerate() {
            if self.collector.limited {
                break;
            }
            let mut column = number(
                row.path(&["w:trPr"])
                    .and_then(|node| node.val("w:gridBefore")),
            )
            .unwrap_or(0);
            for cell in wrapped_children(row, "w:tc") {
                if self.collector.limited {
                    break;
                }
                let span = number(
                    cell.path(&["w:tcPr"])
                        .and_then(|node| node.val("w:gridSpan")),
                )
                .unwrap_or(1);
                let end = column.saturating_add(span);
                let mut props = style.map(|style| style.run.clone()).unwrap_or_default();
                if let Some(style) = style {
                    if let Some(whole) = style.conditionals.get("wholeTable") {
                        props.merge(whole, false);
                    }
                    let [first_row, last_row, first_col, last_col, no_h, no_v] = look.flags;
                    let top = first_row && row_index == 0;
                    let bottom = last_row && row_index + 1 == rows.len();
                    let left = first_col && column == 0;
                    let right = last_col && end == columns;
                    let row_band = number(tpr.and_then(|node| node.val("w:tblStyleRowBandSize")))
                        .or(style.row_band)
                        .unwrap_or(1);
                    let col_band = number(tpr.and_then(|node| node.val("w:tblStyleColBandSize")))
                        .or(style.col_band)
                        .unwrap_or(1);
                    let mut regions = Vec::new();
                    if !no_h && !top && !bottom && row_band > 0 {
                        regions.push(
                            if (row_index.saturating_sub(usize::from(first_row)) / row_band)
                                .is_multiple_of(2)
                            {
                                "band1Horz"
                            } else {
                                "band2Horz"
                            },
                        );
                    }
                    if !no_v && !left && !right && col_band > 0 {
                        regions.push(
                            if (column.saturating_sub(usize::from(first_col)) / col_band)
                                .is_multiple_of(2)
                            {
                                "band1Vert"
                            } else {
                                "band2Vert"
                            },
                        );
                    }
                    for (region, active) in [
                        ("firstCol", left),
                        ("lastCol", right),
                        ("firstRow", top),
                        ("lastRow", bottom),
                        ("nwCell", top && left),
                        ("neCell", top && right),
                        ("swCell", bottom && left),
                        ("seCell", bottom && right),
                    ] {
                        if active {
                            regions.push(region);
                        }
                    }
                    for region in regions {
                        if let Some(conditional) = style.conditionals.get(region) {
                            props.merge(conditional, false);
                        }
                    }
                }
                if cell
                    .path(&["w:tcPr", "w:vMerge"])
                    .is_some_and(|node| node.attr("val") != Some("restart"))
                {
                    *self.complete = false;
                } else {
                    self.walk(cell, &props);
                }
                column = end;
            }
        }
    }
}

fn math_text(node: &Node, text: &mut String) {
    if node.name == "m:t" {
        text.push_str(&node.text);
    }
    for child in &node.children {
        math_text(child, text);
    }
}

fn wrapped_children<'a>(node: &'a Node, name: &str) -> Vec<&'a Node> {
    let mut result = Vec::new();
    for child in &node.children {
        if child.name == name {
            result.push(child);
        } else if matches!(
            child.name,
            "w:sdt" | "w:sdtContent" | "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo"
        ) {
            result.extend(wrapped_children(child, name));
        }
    }
    result
}
