use crate::model::{DrawingTheme, Faces};
use crate::xml::Node;
use crate::{FontPreflight, Parts, PreflightError, xml, zip::Zip};

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

fn table_flag(value: Option<&str>) -> Option<bool> {
    match value? {
        "on" | "1" | "true" => Some(true),
        "off" | "0" | "false" => Some(false),
        _ => None,
    }
}

impl Run {
    fn parse(node: Option<&Node>, scan: &mut Scan<'_>) -> Self {
        let Some(node) = node else {
            return Self::default();
        };
        Self {
            family: scan
                .child(node, "a:latin")
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

struct TableStyle<'a> {
    node: Option<&'a Node>,
    builtin: bool,
    first_row: bool,
    last_row: bool,
    first_col: bool,
    last_col: bool,
    band_row: bool,
}

impl Paragraph {
    fn merge(&mut self, node: Option<&Node>, scan: &mut Scan<'_>) {
        let Some(node) = node else { return };
        let run = scan.child(node, "a:defRPr");
        self.run.merge(&Run::parse(run, scan));
        if scan.child(node, "a:buNone").is_some() {
            self.bullet = Some(false);
        }
        if let Some(marker) = scan.child(node, "a:buChar") {
            self.bullet = Some(marker.attr("char").is_some_and(|value| !value.is_empty()));
        }
        if scan.child(node, "a:buAutoNum").is_some() {
            self.bullet = Some(true);
        }
        if scan.child(node, "a:buFontTx").is_some() {
            self.bullet_family = None;
        }
        if let Some(font) = scan.child(node, "a:buFont") {
            self.bullet_family = font.attr("typeface").map(str::to_owned);
        }
    }
}

fn paragraphs(body: &Node) -> impl Iterator<Item = &Node> {
    body.children.iter().filter(|node| node.name == "a:p")
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

fn resolve_alternate_content(node: &mut Node, complete: &mut bool) {
    let mut children = Vec::new();
    for mut child in std::mem::take(&mut node.children) {
        if child.name != "AlternateContent" {
            resolve_alternate_content(&mut child, complete);
            children.push(child);
            continue;
        }
        if !matches!(node.name, "p:spTree" | "p:grpSp") {
            *complete = false;
            children.push(child);
            continue;
        }
        let branch = child
            .children
            .iter()
            .position(|branch| branch.name == "Choice" && branch.supported_choice)
            .or_else(|| {
                child
                    .children
                    .iter()
                    .position(|branch| branch.name == "Fallback")
            });
        if let Some(index) = branch {
            let mut branch = child.children.remove(index);
            branch.name = node.name;
            resolve_alternate_content(&mut branch, complete);
            children.extend(branch.children);
        } else {
            *complete = false;
        }
    }
    node.children = children;
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
    limited: bool,
}

impl Scan<'_> {
    fn charge(&mut self) -> bool {
        self.spend(1)
    }

    fn spend(&mut self, count: usize) -> bool {
        if self.limited {
            return false;
        }
        if let Some(work) = self.work.checked_sub(count) {
            self.work = work;
            true
        } else {
            self.complete = false;
            self.work = 0;
            self.limited = true;
            false
        }
    }

    fn child<'a>(&mut self, node: &'a Node, name: &str) -> Option<&'a Node> {
        for child in &node.children {
            if !self.charge() {
                return None;
            }
            if child.name == name {
                return Some(child);
            }
        }
        None
    }

    fn path<'a>(&mut self, node: &'a Node, names: &[&str]) -> Option<&'a Node> {
        names
            .iter()
            .try_fold(node, |node, name| self.child(node, name))
    }

    fn placeholder<'a>(&mut self, shape: &'a Node) -> Option<&'a Node> {
        for node in &shape.children {
            if !self.charge() {
                return None;
            }
            if let Some(ph) = self.path(node, &["p:nvPr", "p:ph"]) {
                return Some(ph);
            }
        }
        None
    }

    fn finish(self) -> Result<FontPreflight, PreflightError> {
        if self.limited {
            Err(PreflightError::Limit)
        } else {
            Ok(self.faces.finish(self.complete))
        }
    }

    fn find<'a>(&mut self, root: &'a Node, target: &Node) -> Option<&'a Node> {
        for shape in &root.children {
            if !self.charge() {
                return None;
            }
            if self
                .placeholder(shape)
                .is_some_and(|value| matches_placeholder(value, target))
            {
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
        snapshot: bool,
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
        let own_list = self.child(body, "a:lstStyle");
        let parent_lists = parents.map(|body| body.and_then(|body| self.child(body, "a:lstStyle")));
        let parent_paragraphs = parents.map(|body| {
            body.map(|body| paragraphs(body).collect::<Vec<_>>())
                .unwrap_or_default()
        });
        for (index, own) in own_paragraphs.iter().enumerate() {
            if !self.charge() || self.faces.limited {
                break;
            }
            let ppr = self.child(own, "a:pPr");
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
                .and_then(|node| self.path(node, &["p:txStyles", style_kind]));
            if table_run.is_some() {
                let authored = master_style.and_then(|node| self.child(node, LEVELS[level_index]));
                if authored.is_some() {
                    properties.merge(authored, self);
                } else {
                    let style = self.child(self.presentation, "p:defaultTextStyle");
                    properties.merge(style.and_then(|node| self.child(node, "a:defPPr")), self);
                    properties.merge(
                        style.and_then(|node| self.child(node, LEVELS[level_index])),
                        self,
                    );
                }
                if let Some(run) = table_run {
                    properties.run.merge(run);
                }
            } else {
                properties.merge(
                    master_style.and_then(|node| self.child(node, LEVELS[level_index])),
                    self,
                );
                if ph.is_none() {
                    let style = self.child(self.presentation, "p:defaultTextStyle");
                    properties.merge(style.and_then(|node| self.child(node, "a:defPPr")), self);
                    properties.merge(
                        style.and_then(|node| self.child(node, LEVELS[level_index])),
                        self,
                    );
                }
            }
            for (list, source_paragraphs) in parent_lists
                .into_iter()
                .zip(&parent_paragraphs)
                .rev()
                .chain(std::iter::once((own_list, &own_paragraphs)))
            {
                properties.merge(list.and_then(|node| self.child(node, "a:defPPr")), self);
                properties.merge(
                    list.and_then(|node| self.child(node, LEVELS[level_index])),
                    self,
                );
                let source = source_paragraphs
                    .get(index)
                    .or_else(|| source_paragraphs.get(level_index));
                properties.merge(source.and_then(|node| self.child(node, "a:pPr")), self);
            }
            let end = std::iter::once(&own_paragraphs)
                .chain(parent_paragraphs.iter())
                .find_map(|paragraphs| {
                    paragraphs
                        .get(index)
                        .and_then(|node| self.child(node, "a:endParaRPr"))
                });
            let end_language = self
                .child(own, "a:endParaRPr")
                .and_then(|node| node.attr("lang"));
            let style_language = properties.run.language.as_deref();
            let runs: Vec<_> = own
                .children
                .iter()
                .filter(|node| matches!(node.name, "a:r" | "a:fld" | "a:br"))
                .collect();
            let text_context = ooxml_text::paragraph_scripts_with_languages(
                runs.iter().map(|run| {
                    (
                        self.child(run, "a:t").map_or("", |node| node.text.as_str()),
                        None,
                    )
                }),
                false,
            );
            let mut paragraph_scripts = Vec::new();
            let mut first = None;
            for run in &runs {
                if !self.charge() {
                    break;
                }
                let text = if run.name == "a:br" {
                    "\n"
                } else if run.name == "a:fld" && run.attr("type") == Some("slidenum") {
                    "1"
                } else {
                    self.child(run, "a:t").map_or("", |node| node.text.as_str())
                };
                if snapshot && text.is_empty() {
                    continue;
                }
                let rpr = self.child(run, "a:rPr");
                let direct = Run::parse(rpr, self);
                let language = [direct.language.as_deref(), end_language, style_language]
                    .into_iter()
                    .flatten()
                    .find(|tag| ooxml_text::han_script_for_language(tag).is_some());
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
                run.merge(&Run::parse(end, self));
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

    fn table_run(
        &mut self,
        style: &TableStyle<'_>,
        row: usize,
        col: usize,
        rows: usize,
        cols: usize,
    ) -> Run {
        let first_row = style.first_row;
        let last_row = style.last_row;
        let mut run = Run::default();
        if let Some(node) = style.node {
            let mut regions = vec!["a:wholeTbl"];
            if style.band_row && row >= usize::from(first_row) && row + usize::from(last_row) < rows
            {
                regions.push(if (row - usize::from(first_row)).is_multiple_of(2) {
                    "a:band1H"
                } else {
                    "a:band2H"
                });
            }
            for (name, selected) in [
                ("a:firstCol", style.first_col && col == 0),
                ("a:lastCol", style.last_col && col + 1 == cols),
                ("a:firstRow", first_row && row == 0),
                ("a:lastRow", last_row && row + 1 == rows),
            ] {
                if selected {
                    regions.push(name);
                }
            }
            for region in regions {
                if let Some(text) = self.path(node, &[region, "a:tcTxStyle"]) {
                    run.merge(&Run {
                        bold: table_flag(text.attr("b")),
                        italic: table_flag(text.attr("i")),
                        ..Run::default()
                    });
                }
            }
        } else if style.builtin
            && ((first_row && row == 0)
                || (last_row && row + 1 == rows)
                || (style.first_col && col == 0)
                || (style.last_col && col + 1 == cols))
        {
            run.bold = Some(true);
        }
        run
    }

    fn table(&mut self, table: &Node, snapshot: bool) {
        if !self.spend(table.children.len()) {
            return;
        }
        let properties = self.child(table, "a:tblPr");
        let id = properties
            .and_then(|node| self.child(node, "a:tableStyleId"))
            .map(|node| node.text.trim());
        let node = id.and_then(|id| {
            let styles = self.table_styles?;
            for node in &styles.children {
                if !self.charge() {
                    return None;
                }
                if node.name == "a:tblStyle" && node.attr("styleId") == Some(id) {
                    return Some(node);
                }
            }
            None
        });
        let active = |key| flag(properties.and_then(|node| node.attr(key))) == Some(true);
        let style = TableStyle {
            node,
            builtin: id == Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"),
            first_row: active("firstRow"),
            last_row: active("lastRow"),
            first_col: active("firstCol"),
            last_col: active("lastCol"),
            band_row: active("bandRow"),
        };
        let rows: Vec<_> = table
            .children
            .iter()
            .filter(|node| node.name == "a:tr")
            .collect();
        let mut cols = self
            .child(table, "a:tblGrid")
            .map_or(0, |node| node.children.len());
        for row in &rows {
            if !self.spend(row.children.len()) {
                return;
            }
            cols = cols.max(
                row.children
                    .iter()
                    .filter(|node| node.name == "a:tc")
                    .count(),
            );
        }
        for (row_index, row) in rows.iter().enumerate() {
            if !self.spend(row.children.len()) {
                return;
            }
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
                if let Some(body) = self.child(cell, "a:txBody") {
                    let run = self.table_run(&style, row_index, col, rows.len(), cols);
                    self.body(body, [None, None], None, Some(&run), snapshot);
                }
            }
        }
    }

    fn shapes(&mut self, root: &Node, inherit: bool, placeholders: bool) {
        for shape in &root.children {
            if !self.charge() || self.faces.limited {
                return;
            }
            let mut hidden = false;
            for node in &shape.children {
                if !self.charge() {
                    return;
                }
                if self
                    .child(node, "p:cNvPr")
                    .is_some_and(|node| flag(node.attr("hidden")) == Some(true))
                {
                    hidden = true;
                    break;
                }
            }
            if hidden {
                continue;
            }
            let ph = self.placeholder(shape);
            if !placeholders && ph.is_some() {
                continue;
            }
            match shape.name {
                "p:grpSp" => self.shapes(shape, inherit, true),
                "p:sp" | "p:cxnSp" => {
                    if let Some(body) = self.child(shape, "p:txBody") {
                        let parents = if inherit {
                            ph.map(|ph| {
                                let layout = self
                                    .layout
                                    .and_then(|node| self.path(node, &["p:cSld", "p:spTree"]))
                                    .and_then(|tree| self.find(tree, ph));
                                let master = self
                                    .master
                                    .and_then(|node| self.path(node, &["p:cSld", "p:spTree"]))
                                    .and_then(|tree| self.find(tree, ph));
                                [
                                    layout.and_then(|shape| self.child(shape, "p:txBody")),
                                    master.and_then(|shape| self.child(shape, "p:txBody")),
                                ]
                            })
                            .unwrap_or([None, None])
                        } else {
                            [None, None]
                        };
                        self.body(body, parents, ph, None, inherit);
                    }
                }
                "p:graphicFrame" => {
                    if let Some(graphic) = self.path(shape, &["a:graphic", "a:graphicData"]) {
                        if let Some(table) = self.child(graphic, "a:tbl") {
                            self.table(table, inherit);
                        } else if !graphic.children.is_empty() {
                            self.complete = false;
                        }
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
    let mut presentation = parts
        .part("ppt/presentation.xml", &mut complete)?
        .ok_or(PreflightError::MissingPart)?;
    resolve_alternate_content(&mut presentation, &mut complete);
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
    let mut slide = parts
        .part(path, &mut complete)?
        .ok_or(PreflightError::MissingPart)?;
    resolve_alternate_content(&mut slide, &mut complete);
    if slide.name != "p:sld" || tree(Some(&slide)).is_none() {
        return Err(PreflightError::Xml);
    }
    let slide_rels = parts.rels(path, &mut complete)?;
    let mut layout = parts.related(&slide_rels, "slideLayout", &mut complete)?;
    if layout.is_none() {
        complete = false;
    }
    let layout_rels = match &layout {
        Some((path, _)) => parts.rels(path, &mut complete)?,
        None => Vec::new(),
    };
    let mut master = parts.related(&layout_rels, "slideMaster", &mut complete)?;
    if master.is_none() {
        complete = false;
    }
    let master_rels = match &master {
        Some((path, _)) => parts.rels(path, &mut complete)?,
        None => Vec::new(),
    };
    let mut theme = parts.related(&master_rels, "theme", &mut complete)?;
    if theme.is_none() {
        complete = false;
    }
    for (_, root) in [&mut layout, &mut master, &mut theme].into_iter().flatten() {
        resolve_alternate_content(root, &mut complete);
    }
    let mut table_styles = if has_table(&slide)
        || layout.as_ref().is_some_and(|(_, node)| has_table(node))
        || master.as_ref().is_some_and(|(_, node)| has_table(node))
    {
        if rels.iter().any(|rel| rel.kind == "tableStyles") {
            parts.related(&rels, "tableStyles", &mut complete)?
        } else {
            parts
                .part("ppt/tableStyles.xml", &mut complete)?
                .map(|root| ("ppt/tableStyles.xml".to_owned(), root))
        }
    } else {
        None
    };
    if let Some((_, root)) = &mut table_styles {
        resolve_alternate_content(root, &mut complete);
    }
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
        work: parts.nodes,
        limited: false,
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
    scan.finish()
}

#[cfg(test)]
mod tests;
