use crate::model::{DrawingTheme, Faces};
use crate::xml::Node;
use crate::{FontPreflight, Parts, PreflightError, xml, zip::Zip};

const MAX_WORK: usize = 1_000_000;
const LEVELS: [&str; 9] = [
    "a:lvl1pPr",
    "a:lvl2pPr",
    "a:lvl3pPr",
    "a:lvl4pPr",
    "a:lvl5pPr",
    "a:lvl6pPr",
    "a:lvl7pPr",
    "a:lvl8pPr",
    "a:lvl9pPr",
];

#[derive(Clone, Default)]
struct Run {
    family: Option<String>,
    bold: Option<bool>,
    italic: Option<bool>,
    language: Option<String>,
}

fn flag(value: Option<&str>) -> Option<bool> {
    value.map(|value| matches!(value, "1" | "true" | "on"))
}

impl Run {
    fn parse(node: Option<&Node>) -> Self {
        let Some(node) = node else {
            return Self::default();
        };
        Self {
            family: node
                .child("a:latin")
                .and_then(|node| node.attr("typeface"))
                .map(str::to_owned),
            bold: flag(node.attr("b")),
            italic: flag(node.attr("i")),
            language: node.attr("lang").map(str::to_owned),
        }
    }

    fn merge(&mut self, source: &Self) {
        if source.family.is_some() {
            self.family.clone_from(&source.family);
        }
        if source.bold.is_some() {
            self.bold = source.bold;
        }
        if source.italic.is_some() {
            self.italic = source.italic;
        }
        if source.language.is_some() {
            self.language.clone_from(&source.language);
        }
    }
}

#[derive(Default)]
struct Paragraph {
    run: Run,
    bullet: Option<bool>,
    bullet_family: Option<String>,
}

impl Paragraph {
    fn merge(&mut self, node: Option<&Node>, work: &mut usize, complete: &mut bool) {
        let Some(node) = node else { return };
        let run = node.child("a:defRPr");
        let cost = node.children.len() + run.map_or(0, |node| node.children.len());
        let Some(remaining) = work.checked_sub(cost) else {
            *complete = false;
            *work = 0;
            return;
        };
        *work = remaining;
        self.run.merge(&Run::parse(run));
        if node.child("a:buNone").is_some() {
            self.bullet = Some(false);
        }
        if let Some(marker) = node.child("a:buChar") {
            self.bullet = Some(marker.attr("char").is_some_and(|value| !value.is_empty()));
        }
        if node.child("a:buAutoNum").is_some() {
            self.bullet = Some(true);
        }
        if node.child("a:buFontTx").is_some() {
            self.bullet_family = None;
        }
        if let Some(font) = node.child("a:buFont") {
            self.bullet_family = font.attr("typeface").map(str::to_owned);
        }
    }
}

fn level(style: Option<&Node>, index: usize) -> Option<&Node> {
    let style = style?;
    if index < LEVELS.len() {
        style.child(LEVELS[index])
    } else {
        style.child(LEVELS[0])
    }
}

fn paragraphs(body: &Node) -> impl Iterator<Item = &Node> {
    body.children.iter().filter(|node| node.name == "a:p")
}

fn placeholder(shape: &Node) -> Option<&Node> {
    shape
        .children
        .iter()
        .find_map(|node| node.path(&["p:nvPr", "p:ph"]))
}

fn placeholder_type(node: &Node) -> &str {
    match node.attr("type").unwrap_or("body") {
        "ctrTitle" => "title",
        "obj" => "body",
        kind => kind,
    }
}

fn matches_placeholder(left: &Node, right: &Node) -> bool {
    let a = placeholder_type(left);
    let b = placeholder_type(right);
    if ["title", "sldNum", "dt", "ftr", "hdr"].contains(&a)
        || ["title", "sldNum", "dt", "ftr", "hdr"].contains(&b)
    {
        return a == b;
    }
    match (
        left.attr("idx").and_then(|value| value.parse::<u32>().ok()),
        right
            .attr("idx")
            .and_then(|value| value.parse::<u32>().ok()),
    ) {
        (Some(a), Some(b)) => a == b,
        _ => a == b,
    }
}

fn tree(root: Option<&Node>) -> Option<&Node> {
    root?.path(&["p:cSld", "p:spTree"])
}

struct Scan<'a> {
    presentation: &'a Node,
    layout: Option<&'a Node>,
    master: Option<&'a Node>,
    table_styles: Option<&'a Node>,
    theme: DrawingTheme,
    faces: Faces,
    complete: bool,
    work: usize,
}

impl Scan<'_> {
    fn charge(&mut self) -> bool {
        self.spend(1)
    }

    fn spend(&mut self, count: usize) -> bool {
        if let Some(work) = self.work.checked_sub(count) {
            self.work = work;
            true
        } else {
            self.complete = false;
            false
        }
    }

    fn find<'a>(&mut self, root: &'a Node, target: &Node) -> Option<&'a Node> {
        for shape in &root.children {
            if !self.charge() {
                return None;
            }
            if placeholder(shape).is_some_and(|value| matches_placeholder(value, target)) {
                return Some(shape);
            }
            if shape.name == "p:grpSp"
                && let Some(found) = self.find(shape, target)
            {
                return Some(found);
            }
        }
        None
    }

    fn record(&mut self, run: &Run, scripts: &[String]) {
        let family = self
            .theme
            .resolve(run.family.as_deref().unwrap_or("+mn-lt"));
        self.faces.record(
            &family,
            run.bold.unwrap_or(false),
            run.italic.unwrap_or(false),
            scripts,
        );
    }

    fn body(
        &mut self,
        body: &Node,
        parents: [Option<&Node>; 2],
        ph: Option<&Node>,
        table_run: Option<&Run>,
    ) {
        let source_work = body.children.len()
            + parents
                .into_iter()
                .flatten()
                .map(|body| body.children.len())
                .sum::<usize>();
        if !self.spend(source_work) {
            return;
        }
        let own_paragraphs: Vec<_> = paragraphs(body).collect();
        let own_list = body.child("a:lstStyle");
        let parent_lists = parents.map(|body| body.and_then(|body| body.child("a:lstStyle")));
        let parent_paragraphs = parents.map(|body| {
            body.map(|body| paragraphs(body).collect::<Vec<_>>())
                .unwrap_or_default()
        });
        for (index, own) in own_paragraphs.iter().enumerate() {
            if !self.charge() || self.faces.limited {
                break;
            }
            let ppr = own.child("a:pPr");
            let level_index = ppr
                .and_then(|node| node.attr("lvl"))
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(0)
                .min(8);
            let mut properties = Paragraph::default();
            let style_kind = match ph.map(placeholder_type) {
                Some("title") => "p:titleStyle",
                Some("body" | "subTitle") => "p:bodyStyle",
                _ => "p:otherStyle",
            };
            let master_style = self
                .master
                .and_then(|node| node.path(&["p:txStyles", style_kind]));
            if !self.spend(master_style.map_or(0, |node| node.children.len())) {
                return;
            }
            if table_run.is_some() {
                let authored = master_style.and_then(|node| node.child(LEVELS[level_index]));
                if authored.is_some() {
                    properties.merge(authored, &mut self.work, &mut self.complete);
                } else {
                    let style = self.presentation.child("p:defaultTextStyle");
                    if !self.spend(style.map_or(0, |node| node.children.len())) {
                        return;
                    }
                    properties.merge(
                        style.and_then(|node| node.child("a:defPPr")),
                        &mut self.work,
                        &mut self.complete,
                    );
                    properties.merge(
                        level(style, level_index),
                        &mut self.work,
                        &mut self.complete,
                    );
                }
                if let Some(run) = table_run {
                    properties.run.merge(run);
                }
            } else {
                properties.merge(
                    level(master_style, level_index),
                    &mut self.work,
                    &mut self.complete,
                );
                if ph.is_none() {
                    let style = self.presentation.child("p:defaultTextStyle");
                    if !self.spend(style.map_or(0, |node| node.children.len())) {
                        return;
                    }
                    properties.merge(
                        style.and_then(|node| node.child("a:defPPr")),
                        &mut self.work,
                        &mut self.complete,
                    );
                    properties.merge(
                        level(style, level_index),
                        &mut self.work,
                        &mut self.complete,
                    );
                }
            }
            for (list, source_paragraphs) in parent_lists
                .into_iter()
                .zip(&parent_paragraphs)
                .rev()
                .chain(std::iter::once((own_list, &own_paragraphs)))
            {
                if !self.spend(list.map_or(0, |node| node.children.len())) {
                    return;
                }
                properties.merge(
                    list.and_then(|node| node.child("a:defPPr")),
                    &mut self.work,
                    &mut self.complete,
                );
                properties.merge(
                    list.and_then(|node| node.child(LEVELS[level_index])),
                    &mut self.work,
                    &mut self.complete,
                );
                let source = source_paragraphs
                    .get(index)
                    .or_else(|| source_paragraphs.get(level_index));
                if !self.spend(source.map_or(0, |node| node.children.len())) {
                    return;
                }
                properties.merge(
                    source.and_then(|node| node.child("a:pPr")),
                    &mut self.work,
                    &mut self.complete,
                );
            }
            let end = std::iter::once(&own_paragraphs)
                .chain(parent_paragraphs.iter())
                .find_map(|paragraphs| {
                    paragraphs
                        .get(index)
                        .and_then(|node| node.child("a:endParaRPr"))
                });
            let end_language = own.child("a:endParaRPr").and_then(|node| node.attr("lang"));
            let style_language = properties.run.language.as_deref();
            let runs: Vec<_> = own
                .children
                .iter()
                .filter(|node| matches!(node.name, "a:r" | "a:fld" | "a:br"))
                .collect();
            let text_context = ooxml_text::paragraph_scripts_with_languages(
                runs.iter()
                    .map(|run| (run.child("a:t").map_or("", |node| node.text.as_str()), None)),
                false,
            );
            let mut paragraph_scripts = Vec::new();
            let mut first = None;
            for run in &runs {
                if !self.charge() {
                    break;
                }
                let rpr = run.child("a:rPr");
                if !self.spend(rpr.map_or(0, |node| node.children.len())) {
                    break;
                }
                let direct = Run::parse(rpr);
                let language = [direct.language.as_deref(), end_language, style_language]
                    .into_iter()
                    .flatten()
                    .find(|tag| ooxml_text::han_script_for_language(tag).is_some());
                let text = if run.name == "a:fld" && run.attr("type") == Some("slidenum") {
                    "1"
                } else {
                    run.child("a:t").map_or("", |node| node.text.as_str())
                };
                let mut scripts =
                    ooxml_text::paragraph_scripts_with_languages([(text, language)], false);
                if language.is_none()
                    && scripts.iter().any(|script| script == "cjk-sc")
                    && text_context
                        .iter()
                        .any(|script| matches!(script.as_str(), "cjk-jp" | "cjk-kr"))
                {
                    scripts.retain(|script| script != "cjk-sc");
                    scripts.extend(
                        text_context
                            .iter()
                            .filter(|script| matches!(script.as_str(), "cjk-jp" | "cjk-kr"))
                            .cloned(),
                    );
                }
                for script in &scripts {
                    if !paragraph_scripts.contains(script) {
                        paragraph_scripts.push(script.clone());
                    }
                }
                let mut resolved = properties.run.clone();
                resolved.merge(&direct);
                if first.is_none() {
                    first = Some(resolved.clone());
                }
                self.record(&resolved, &scripts);
            }
            let first = first.unwrap_or_else(|| {
                let mut run = properties.run.clone();
                run.merge(&Run::parse(end));
                run
            });
            self.record(&first, &[]);
            if properties.bullet == Some(true)
                && let Some(family) = properties.bullet_family
            {
                self.record(
                    &Run {
                        family: Some(family),
                        ..first
                    },
                    &paragraph_scripts,
                );
            }
        }
    }

    fn table_run(&self, table: &Node, row: usize, col: usize, rows: usize, cols: usize) -> Run {
        let properties = table.child("a:tblPr");
        let id = properties
            .and_then(|node| node.child("a:tableStyleId"))
            .map(|node| node.text.trim());
        let style = id.and_then(|id| {
            self.table_styles?
                .children
                .iter()
                .find(|node| node.name == "a:tblStyle" && node.attr("styleId") == Some(id))
        });
        let active = |key| flag(properties.and_then(|node| node.attr(key))) == Some(true);
        let first_row = active("firstRow");
        let last_row = active("lastRow");
        let mut run = Run::default();
        if let Some(style) = style {
            let mut regions = vec!["a:wholeTbl"];
            if active("bandRow")
                && row >= usize::from(first_row)
                && row + usize::from(last_row) < rows
            {
                regions.push(if (row - usize::from(first_row)).is_multiple_of(2) {
                    "a:band1H"
                } else {
                    "a:band2H"
                });
            }
            for (name, selected) in [
                ("a:firstCol", active("firstCol") && col == 0),
                ("a:lastCol", active("lastCol") && col + 1 == cols),
                ("a:firstRow", first_row && row == 0),
                ("a:lastRow", last_row && row + 1 == rows),
            ] {
                if selected {
                    regions.push(name);
                }
            }
            for region in regions {
                if let Some(text) = style.path(&[region, "a:tcTxStyle"]) {
                    run.merge(&Run {
                        bold: flag(text.attr("b")),
                        italic: flag(text.attr("i")),
                        ..Run::default()
                    });
                }
            }
        } else if id == Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
            && ((first_row && row == 0)
                || (last_row && row + 1 == rows)
                || (active("firstCol") && col == 0)
                || (active("lastCol") && col + 1 == cols))
        {
            run.bold = Some(true);
        }
        run
    }

    fn table(&mut self, table: &Node) {
        let rows: Vec<_> = table
            .children
            .iter()
            .filter(|node| node.name == "a:tr")
            .collect();
        let cols = table
            .child("a:tblGrid")
            .map_or(0, |node| node.children.len())
            .max(
                rows.iter()
                    .map(|row| {
                        row.children
                            .iter()
                            .filter(|node| node.name == "a:tc")
                            .count()
                    })
                    .max()
                    .unwrap_or(0),
            );
        for (row_index, row) in rows.iter().enumerate() {
            for (col, cell) in row
                .children
                .iter()
                .filter(|node| node.name == "a:tc")
                .enumerate()
            {
                if !self.charge() {
                    return;
                }
                if flag(cell.attr("hMerge")) == Some(true)
                    || flag(cell.attr("vMerge")) == Some(true)
                {
                    continue;
                }
                if let Some(body) = cell.child("a:txBody") {
                    if !self.spend(self.table_styles.map_or(0, |node| node.children.len())) {
                        return;
                    }
                    let run = self.table_run(table, row_index, col, rows.len(), cols);
                    self.body(body, [None, None], None, Some(&run));
                }
            }
        }
    }

    fn shapes(&mut self, root: &Node, inherit: bool, placeholders: bool) {
        for shape in &root.children {
            if !self.charge() || self.faces.limited {
                return;
            }
            if shape
                .children
                .iter()
                .filter_map(|node| node.child("p:cNvPr"))
                .any(|node| flag(node.attr("hidden")) == Some(true))
            {
                continue;
            }
            let ph = placeholder(shape);
            if !placeholders && ph.is_some() {
                continue;
            }
            match shape.name {
                "p:grpSp" => self.shapes(shape, inherit, true),
                "p:sp" | "p:cxnSp" => {
                    if let Some(body) = shape.child("p:txBody") {
                        let parents = if inherit {
                            ph.map(|ph| {
                                let layout = tree(self.layout).and_then(|tree| self.find(tree, ph));
                                let master = tree(self.master).and_then(|tree| self.find(tree, ph));
                                [
                                    layout.and_then(|shape| shape.child("p:txBody")),
                                    master.and_then(|shape| shape.child("p:txBody")),
                                ]
                            })
                            .unwrap_or([None, None])
                        } else {
                            [None, None]
                        };
                        self.body(body, parents, ph, None);
                    }
                }
                "p:graphicFrame" => {
                    if let Some(graphic) = shape.path(&["a:graphic", "a:graphicData"]) {
                        if let Some(table) = graphic.child("a:tbl") {
                            self.table(table);
                        } else if !graphic.children.is_empty() {
                            self.complete = false;
                        }
                    }
                }
                "AlternateContent" => {
                    self.complete = false;
                    if let Some(choice) = shape.children.first() {
                        self.shapes(choice, inherit, placeholders);
                    }
                }
                _ => {}
            }
        }
    }
}

fn has_table(root: &Node) -> bool {
    root.name == "a:tbl" || root.children.iter().any(has_table)
}

pub fn pptx_font_preflight(
    bytes: &[u8],
    active_slide: usize,
) -> Result<FontPreflight, PreflightError> {
    let mut parts = Parts {
        zip: Zip::new(bytes)?,
        nodes: xml::MAX_NODES,
    };
    let mut complete = true;
    let presentation = parts
        .part("ppt/presentation.xml", &mut complete)?
        .ok_or(PreflightError::MissingPart)?;
    if presentation.name != "p:presentation" {
        return Err(PreflightError::Xml);
    }
    let rels = parts.rels("ppt/presentation.xml", &mut complete)?;
    let id = presentation
        .child("p:sldIdLst")
        .into_iter()
        .flat_map(|node| &node.children)
        .filter(|node| node.name == "p:sldId")
        .nth(active_slide)
        .and_then(|node| node.attr("relId"))
        .ok_or(PreflightError::MissingPart)?;
    let path = &rels
        .iter()
        .find(|rel| rel.id == id && rel.kind == "slide")
        .ok_or(PreflightError::MissingPart)?
        .target;
    let slide = parts
        .part(path, &mut complete)?
        .ok_or(PreflightError::MissingPart)?;
    if slide.name != "p:sld" || tree(Some(&slide)).is_none() {
        return Err(PreflightError::Xml);
    }
    let slide_rels = parts.rels(path, &mut complete)?;
    let layout = parts.related(&slide_rels, "slideLayout", &mut complete)?;
    if layout.is_none() {
        complete = false;
    }
    let layout_rels = match &layout {
        Some((path, _)) => parts.rels(path, &mut complete)?,
        None => Vec::new(),
    };
    let master = parts.related(&layout_rels, "slideMaster", &mut complete)?;
    if master.is_none() {
        complete = false;
    }
    let master_rels = match &master {
        Some((path, _)) => parts.rels(path, &mut complete)?,
        None => Vec::new(),
    };
    let theme = parts.related(&master_rels, "theme", &mut complete)?;
    if theme.is_none() {
        complete = false;
    }
    let table_styles = if has_table(&slide)
        || layout.as_ref().is_some_and(|(_, node)| has_table(node))
        || master.as_ref().is_some_and(|(_, node)| has_table(node))
    {
        parts.related(&rels, "tableStyles", &mut complete)?
    } else {
        None
    };
    for (part, name) in [
        (layout.as_ref(), "p:sldLayout"),
        (master.as_ref(), "p:sldMaster"),
        (theme.as_ref(), "a:theme"),
        (table_styles.as_ref(), "a:tblStyleLst"),
    ] {
        if part.is_some_and(|(_, node)| node.name != name) {
            return Err(PreflightError::Xml);
        }
    }
    let layout = layout.as_ref().map(|(_, node)| node);
    let master = master.as_ref().map(|(_, node)| node);
    let mut scan = Scan {
        presentation: &presentation,
        layout,
        master,
        table_styles: table_styles.as_ref().map(|(_, node)| node),
        theme: DrawingTheme::parse(theme.as_ref().map(|(_, node)| node)),
        faces: Faces::default(),
        complete,
        work: MAX_WORK,
    };
    if flag(slide.attr("showMasterSp")) != Some(false) {
        if layout.is_none_or(|node| flag(node.attr("showMasterSp")) != Some(false))
            && let Some(tree) = tree(master)
        {
            scan.shapes(tree, false, false);
        }
        if let Some(tree) = tree(layout) {
            scan.shapes(tree, false, false);
        }
    }
    if let Some(tree) = tree(Some(&slide)) {
        scan.shapes(tree, true, true);
    }
    Ok(scan.faces.finish(scan.complete))
}

#[cfg(test)]
mod tests;
