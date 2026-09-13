use std::collections::{BTreeMap, HashSet};

use vsdx_parse::{
    Cell, Row, Section, Shape, ShapeChild, Sheet, SheetChild, TextToken, VsdxPackage,
};

use crate::text::row_cells;
use crate::{
    Lookup, Provenance, ResolveError, ResolvedCell, ResolvedRow, ResolvedSection, ResolvedShape,
    ResolvedTextToken,
};

const MAX_INHERITANCE_DEPTH: usize = 64;

pub struct Resolver<'a> {
    package: &'a VsdxPackage,
    defaults: BTreeMap<String, Cell>,
}
impl<'a> Resolver<'a> {
    pub fn new(package: &'a VsdxPackage) -> Self {
        Self {
            package,
            defaults: documented_display_defaults(),
        }
    }
    pub fn with_defaults(mut self, defaults: impl IntoIterator<Item = Cell>) -> Self {
        self.defaults
            .extend(defaults.into_iter().map(|c| (c.name.clone(), c)));
        self
    }
    pub fn package(&self) -> &'a VsdxPackage {
        self.package
    }
    pub fn resolve_shape(
        &self,
        page_part: &str,
        shape_id: u32,
    ) -> Result<ResolvedShape, ResolveError> {
        let page_contents = self
            .package
            .page_contents
            .get(page_part)
            .ok_or_else(|| ResolveError::MissingPage(page_part.into()))?;
        let shape =
            find_shape(page_contents, shape_id).ok_or(ResolveError::MissingShape(shape_id))?;
        let page = self
            .package
            .page_part_ids
            .get(page_part)
            .and_then(|id| self.package.page_sheets.get(id))
            .unwrap_or(page_contents);
        self.resolve_shape_ref(shape, page)
    }
    pub fn resolve_shape_in_sheet(
        &self,
        shape: &Shape,
        sheet: &Sheet,
    ) -> Result<ResolvedShape, ResolveError> {
        self.resolve_shape_ref(shape, sheet)
    }
    /// Resolves a page or document ShapeSheet with document-level inheritance.
    pub fn resolve_sheet(&self, sheet: &Sheet) -> Result<ResolvedShape, ResolveError> {
        let shape = Shape {
            id: 0,
            name: None,
            name_u: None,
            shape_type: None,
            master: None,
            master_shape: None,
            line_style: style_attribute(sheet, "LineStyle"),
            fill_style: style_attribute(sheet, "FillStyle"),
            text_style: style_attribute(sheet, "TextStyle"),
            children: sheet
                .children
                .iter()
                .filter_map(|child| match child {
                    SheetChild::Cell(cell) => Some(ShapeChild::Cell(cell.clone())),
                    SheetChild::Section(section) => Some(ShapeChild::Section(section.clone())),
                    _ => None,
                })
                .collect(),
            del: false,
            other_attrs: Vec::new(),
        };
        self.resolve_shape_ref(&shape, sheet)
    }
    pub fn resolve_page_shapes(
        &self,
        page_part: &str,
    ) -> Result<BTreeMap<u32, ResolvedShape>, ResolveError> {
        let page_contents = self
            .package
            .page_contents
            .get(page_part)
            .ok_or_else(|| ResolveError::MissingPage(page_part.into()))?;
        let page = self
            .package
            .page_part_ids
            .get(page_part)
            .and_then(|id| self.package.page_sheets.get(id))
            .unwrap_or(page_contents);
        let mut shapes = BTreeMap::new();
        for shape in page_contents.shapes() {
            self.resolve_page_shape_tree(shape, page, &mut shapes)?;
        }
        Ok(shapes)
    }
    fn resolve_page_shape_tree(
        &self,
        shape: &Shape,
        page: &Sheet,
        shapes: &mut BTreeMap<u32, ResolvedShape>,
    ) -> Result<(), ResolveError> {
        shapes.insert(shape.id, self.resolve_shape_ref(shape, page)?);
        for child in shape.shapes() {
            self.resolve_page_shape_tree(child, page, shapes)?;
        }
        Ok(())
    }
    pub fn resolve_text(
        &self,
        shape: &Shape,
        page: &Sheet,
    ) -> Result<Vec<ResolvedTextToken>, ResolveError> {
        let resolved = self.resolve_shape_ref(shape, page)?;
        self.resolve_text_in_context(shape, page, &resolved)
    }
    pub fn resolve_text_in_context(
        &self,
        shape: &Shape,
        page: &Sheet,
        resolved: &ResolvedShape,
    ) -> Result<Vec<ResolvedTextToken>, ResolveError> {
        let masters = self.master_chain(shape, page)?;
        let tokens = shape
            .text()
            .or_else(|| masters.iter().find_map(|(_, master)| master.text()));
        Ok(tokens
            .unwrap_or_default()
            .iter()
            .map(|token| match token {
                TextToken::Literal(value) => ResolvedTextToken::Literal(value.clone()),
                TextToken::CharacterRun(index) => ResolvedTextToken::CharacterRun {
                    index: *index,
                    properties: row_cells(resolved, "Character", *index),
                },
                TextToken::ParagraphRun(index) => ResolvedTextToken::ParagraphRun {
                    index: *index,
                    properties: row_cells(resolved, "Paragraph", *index),
                },
                TextToken::Tab(index) => ResolvedTextToken::Tab {
                    index: *index,
                    properties: row_cells(resolved, "Tabs", *index),
                },
                TextToken::Field(index) => ResolvedTextToken::Field {
                    index: *index,
                    properties: row_cells(resolved, "Field", *index),
                },
            })
            .collect())
    }
    fn resolve_shape_ref(
        &self,
        shape: &Shape,
        page: &Sheet,
    ) -> Result<ResolvedShape, ResolveError> {
        if shape.del {
            return Ok(ResolvedShape {
                deleted: true,
                ..Default::default()
            });
        }
        let masters = self.master_chain(shape, page)?;
        let styles = self.style_chains(shape, &masters)?;
        let mut names = HashSet::new();
        for source in std::iter::once(shape as &dyn HasCells)
            .chain(masters.iter().map(|(_, s)| *s as &dyn HasCells))
            .chain(
                styles
                    .iter()
                    .flat_map(|(_, sheets)| sheets.iter().map(|s| *s as &dyn HasCells)),
            )
            .chain(std::iter::once(page as &dyn HasCells))
            .chain(
                self.package
                    .document_sheet
                    .iter()
                    .map(|s| s as &dyn HasCells),
            )
        {
            names.extend(source.cells().map(|c| c.name.clone()));
        }
        names.extend(self.defaults.keys().cloned());
        let mut out = ResolvedShape::default();
        for name in names {
            out.cells.insert(
                name.clone(),
                self.resolve_cell(shape, &masters, &styles, page, &name),
            );
        }
        let mut sections = HashSet::new();
        for source in std::iter::once(shape as &dyn HasSections)
            .chain(masters.iter().map(|(_, s)| *s as &dyn HasSections))
            .chain(
                styles
                    .iter()
                    .flat_map(|(_, sheets)| sheets.iter().map(|s| *s as &dyn HasSections)),
            )
            .chain(std::iter::once(page as &dyn HasSections))
            .chain(
                self.package
                    .document_sheet
                    .iter()
                    .map(|s| s as &dyn HasSections),
            )
        {
            sections.extend(
                source
                    .sections()
                    .map(|s| (s.name.clone(), s.index.unwrap_or(0))),
            );
        }
        for (section, index) in sections {
            out.sections.insert(
                crate::section_key(&section, Some(index)),
                self.resolve_section(shape, &masters, &styles, page, &section, index),
            );
        }
        Ok(out)
    }
    fn master_chain(
        &self,
        shape: &Shape,
        source_sheet: &'a Sheet,
    ) -> Result<Vec<(Provenance, &'a Shape)>, ResolveError> {
        let mut out = Vec::new();
        let mut current = shape;
        let mut current_sheet = source_sheet;
        let mut seen = HashSet::new();
        for depth in 0..MAX_INHERITANCE_DEPTH {
            let (master_id, master_shape, own_master) = match (current.master, current.master_shape)
            {
                (Some(master_id), master_shape) => (master_id, master_shape, true),
                (None, Some(master_shape)) => {
                    let Some(master_id) = self.enclosing_master(current_sheet, current.id) else {
                        return Ok(out);
                    };
                    (master_id, Some(master_shape), false)
                }
                (None, None) => return Ok(out),
            };
            let Some(path) = self
                .package
                .master_part_ids
                .iter()
                .find_map(|(path, id)| (*id == master_id).then_some(path))
            else {
                return Err(ResolveError::MissingMaster(master_id));
            };
            let Some(sheet) = self.package.master_contents.get(path) else {
                return Err(ResolveError::MissingMaster(master_id));
            };
            let (next, provenance) = if own_master {
                let root = sheet.shapes().next();
                match master_shape.and_then(|id| root.and_then(|root| find_shape_in(root, id))) {
                    Some(shape) => (Some(shape), Provenance::MasterShape),
                    None => (root, Provenance::Master),
                }
            } else {
                (
                    master_shape.and_then(|id| find_shape(sheet, id)),
                    Provenance::MasterShape,
                )
            };
            let Some(next) = next else {
                return Err(ResolveError::MissingMaster(master_id));
            };
            if !seen.insert((master_id, next.id)) {
                return Err(ResolveError::Cycle(format!(
                    "master {master_id}/{}",
                    next.id
                )));
            }
            out.push((provenance, next));
            current = next;
            current_sheet = sheet;
            if depth + 1 == MAX_INHERITANCE_DEPTH {
                return Err(ResolveError::Cycle("maximum inheritance depth".into()));
            }
        }
        unreachable!()
    }
    fn enclosing_master(&self, sheet: &'a Sheet, shape_id: u32) -> Option<u32> {
        let parent = enclosing_shape(sheet, shape_id)?;
        parent
            .master
            .or_else(|| self.enclosing_master(sheet, parent.id))
    }
    fn style_chains(
        &self,
        shape: &Shape,
        masters: &[(Provenance, &'a Shape)],
    ) -> Result<Vec<(Provenance, Vec<&'a Sheet>)>, ResolveError> {
        let inherited = |select: fn(&Shape) -> Option<u32>| {
            select(shape).or_else(|| masters.iter().find_map(|(_, master)| select(master)))
        };
        [
            (inherited(|shape| shape.line_style), Provenance::StyleLine),
            (inherited(|shape| shape.fill_style), Provenance::StyleFill),
            (inherited(|shape| shape.text_style), Provenance::StyleText),
        ]
        .into_iter()
        .map(|(id, provenance)| self.style_chain(id).map(|chain| (provenance, chain)))
        .collect()
    }
    fn style_chain(&self, mut id: Option<u32>) -> Result<Vec<&'a Sheet>, ResolveError> {
        let mut out = Vec::new();
        let mut seen = HashSet::new();
        while let Some(current) = id {
            if !seen.insert(current) {
                return Err(ResolveError::Cycle(format!("style {current}")));
            }
            let Some(sheet) = self
                .package
                .style_sheets
                .iter()
                .find(|s| s.id == Some(current))
            else {
                break;
            };
            out.push(sheet);
            id = based_on(sheet);
        }
        Ok(out)
    }
    fn resolve_cell(
        &self,
        shape: &Shape,
        masters: &[(Provenance, &'a Shape)],
        styles: &[(Provenance, Vec<&'a Sheet>)],
        page: &Sheet,
        name: &str,
    ) -> Lookup {
        let mut sources: Vec<(Provenance, Option<&Cell>)> =
            vec![(Provenance::Local, shape.cells().find(|c| c.name == name))];
        sources.extend(
            masters
                .iter()
                .map(|(p, s)| (*p, s.cells().find(|c| c.name == name))),
        );
        if let Some(owner) = style_owner(name)
            && let Some((p, chain)) = styles.iter().find(|(p, _)| *p == owner)
        {
            sources.extend(
                chain
                    .iter()
                    .map(|s| (*p, s.cells().find(|c| c.name == name))),
            );
        }
        sources.push((Provenance::Page, page.cells().find(|c| c.name == name)));
        sources.push((
            Provenance::Document,
            self.package
                .document_sheet
                .as_ref()
                .and_then(|s| s.cells().find(|c| c.name == name)),
        ));
        let mut inherited = None;
        for (provenance, cell) in sources {
            if let Some(cell) = cell {
                if cell
                    .formula
                    .as_deref()
                    .is_some_and(|formula| formula.eq_ignore_ascii_case("Inh"))
                {
                    inherited.get_or_insert_with(|| ResolvedCell {
                        cell: cell.clone(),
                        provenance,
                    });
                    continue;
                }
                return found(cell, provenance);
            }
        }
        self.defaults
            .get(name)
            .map(|cell| ResolvedCell {
                cell: cell.clone(),
                provenance: Provenance::Default,
            })
            .map(Lookup::Found)
            .or_else(|| inherited.map(Lookup::Found))
            .unwrap_or(Lookup::Absent)
    }
    fn resolve_section(
        &self,
        shape: &Shape,
        masters: &[(Provenance, &'a Shape)],
        styles: &[(Provenance, Vec<&'a Sheet>)],
        page: &Sheet,
        name: &str,
        index: u32,
    ) -> ResolvedSection {
        let mut sources: Vec<(Provenance, Option<&Section>)> = vec![(
            Provenance::Local,
            shape
                .sections()
                .find(|s| s.name == name && s.index.unwrap_or(0) == index),
        )];
        sources.extend(masters.iter().map(|(p, s)| {
            (
                *p,
                s.sections()
                    .find(|v| v.name == name && v.index.unwrap_or(0) == index),
            )
        }));
        if (name == "Character" || name == "Paragraph" || name == "Tabs" || name == "Field")
            && let Some((p, chain)) = styles.iter().find(|(p, _)| *p == Provenance::StyleText)
        {
            sources.extend(chain.iter().map(|s| {
                (
                    *p,
                    s.sections()
                        .find(|v| v.name == name && v.index.unwrap_or(0) == index),
                )
            }));
        }
        sources.push((
            Provenance::Page,
            page.sections()
                .find(|s| s.name == name && s.index.unwrap_or(0) == index),
        ));
        sources.push((
            Provenance::Document,
            self.package.document_sheet.as_ref().and_then(|s| {
                s.sections()
                    .find(|s| s.name == name && s.index.unwrap_or(0) == index)
            }),
        ));
        if let Some(position) = sources
            .iter()
            .position(|(_, section)| section.is_some_and(|section| section.del))
        {
            if position == 0 {
                return ResolvedSection {
                    name: name.into(),
                    index: sources
                        .iter()
                        .find_map(|(_, section)| section.and_then(|section| section.index)),
                    deleted: true,
                    ..Default::default()
                };
            }
            sources.truncate(position);
        }
        let mut keys = Vec::new();
        for (_, section) in &sources {
            let Some(section) = section else { continue };
            for key in row_keys(section) {
                if !keys.contains(&key) {
                    keys.push(key);
                }
            }
        }
        let mut out = ResolvedSection {
            name: name.into(),
            index: sources
                .iter()
                .find_map(|(_, section)| section.and_then(|section| section.index)),
            unsupported_controls: unsupported_geometry_controls(name, &sources),
            ..Default::default()
        };
        for key in keys {
            let rows: Vec<(Provenance, Option<&Row>)> = sources
                .iter()
                .map(|(p, section)| {
                    (
                        *p,
                        section.and_then(|s| {
                            row_keys(s)
                                .into_iter()
                                .zip(s.rows())
                                .find(|(row_key, _)| row_key == &key)
                                .map(|(_, row)| row)
                        }),
                    )
                })
                .collect();
            let row = rows.iter().find_map(|(_, row)| *row);
            let Some(row) = row else { continue };
            let resolved_key = resolved_row_key(&key, &out.rows);
            if row.del {
                out.rows.insert(
                    resolved_key.clone(),
                    ResolvedRow {
                        key: resolved_key.clone(),
                        deleted: true,
                        ..Default::default()
                    },
                );
                out.row_order.push(resolved_key);
                continue;
            }
            let mut names = HashSet::new();
            for (_, row) in rows
                .iter()
                .filter(|(_, row)| row.is_none_or(|row| !row.del))
            {
                if let Some(row) = row {
                    names.extend(row.cells().map(|c| c.name.clone()));
                }
            }
            let mut cells = BTreeMap::new();
            for cell_name in names {
                let mut inherited = None;
                let mut lookup = None;
                for (provenance, row) in rows
                    .iter()
                    .filter(|(_, row)| row.is_none_or(|row| !row.del))
                {
                    let Some(cell) =
                        row.and_then(|row| row.cells().find(|cell| cell.name == cell_name))
                    else {
                        continue;
                    };
                    if cell
                        .formula
                        .as_deref()
                        .is_some_and(|formula| formula.eq_ignore_ascii_case("Inh"))
                    {
                        inherited.get_or_insert_with(|| found(cell, *provenance));
                        continue;
                    }
                    lookup = Some(found(cell, *provenance));
                    break;
                }
                let lookup = lookup.or(inherited);
                cells.insert(cell_name, lookup.unwrap_or(Lookup::Absent));
            }
            out.rows.insert(
                resolved_key.clone(),
                ResolvedRow {
                    key: resolved_key.clone(),
                    deleted: false,
                    row_type: row.row_type.clone(),
                    cells,
                },
            );
            out.row_order.push(resolved_key);
        }
        out
    }
}

fn unsupported_geometry_controls(
    name: &str,
    sources: &[(Provenance, Option<&Section>)],
) -> Vec<String> {
    if name != "Geometry" {
        return Vec::new();
    }
    let mut unsupported = Vec::new();
    for control in ["NoFill", "NoLine", "NoShow"] {
        for (_, section) in sources {
            let Some(section) = section else { continue };
            let Some(cell) = section.children.iter().find_map(|child| match child {
                vsdx_parse::SectionChild::Unknown(cell)
                    if cell.name.rsplit(':').next() == Some("Cell")
                        && cell
                            .attributes
                            .iter()
                            .any(|(key, value)| key == "N" && value == control) =>
                {
                    Some(cell)
                }
                _ => None,
            }) else {
                continue;
            };
            let attribute = |name: &str| {
                cell.attributes
                    .iter()
                    .find_map(|(key, value)| (key == name).then_some(value.as_str()))
            };
            if attribute("F").is_some_and(|formula| formula.eq_ignore_ascii_case("Inh")) {
                continue;
            }
            if attribute("Del") == Some("1") {
                break;
            }
            let formula = attribute("F").or_else(|| attribute("V"));
            let zero = formula.is_some_and(|formula| {
                formula.eq_ignore_ascii_case("FALSE")
                    || vsdx_formula::evaluate_number(
                        formula,
                        vsdx_formula::Limits {
                            max_depth: 256,
                            max_nodes: 8192,
                            max_tokens: 16384,
                        },
                        &mut |_| None,
                    ) == Some(0.0)
            });
            if !zero {
                unsupported.push(control.to_owned());
            }
            break;
        }
    }
    unsupported
}

/// Documented transform defaults: https://learn.microsoft.com/en-us/office/client-developer/visio/cells-visio-shapesheet-reference
fn documented_display_defaults() -> BTreeMap<String, Cell> {
    [
        ("LocPinX", "Width * 0.5"),
        ("LocPinY", "Height * 0.5"),
        ("TxtPinX", "Width * 0.5"),
        ("TxtPinY", "Height * 0.5"),
    ]
    .into_iter()
    .map(|(name, formula)| {
        (
            name.into(),
            Cell {
                name: name.into(),
                formula: Some(formula.into()),
                value: None,
                unit: None,
                del: false,
                other_attrs: Vec::new(),
            },
        )
    })
    .collect()
}

fn found(cell: &Cell, provenance: Provenance) -> Lookup {
    if cell.del {
        Lookup::Deleted
    } else {
        Lookup::Found(ResolvedCell {
            cell: cell.clone(),
            provenance,
        })
    }
}
fn find_shape(sheet: &Sheet, id: u32) -> Option<&Shape> {
    sheet.shapes().find_map(|shape| find_shape_in(shape, id))
}
fn find_shape_in(shape: &Shape, id: u32) -> Option<&Shape> {
    if shape.id == id {
        return Some(shape);
    }
    shape.shapes().find_map(|child| find_shape_in(child, id))
}
fn enclosing_shape(sheet: &Sheet, id: u32) -> Option<&Shape> {
    sheet
        .shapes()
        .find_map(|shape| enclosing_shape_in(shape, id))
}
fn enclosing_shape_in(shape: &Shape, id: u32) -> Option<&Shape> {
    if shape.shapes().any(|child| child.id == id) {
        return Some(shape);
    }
    shape
        .shapes()
        .find_map(|child| enclosing_shape_in(child, id))
}
fn based_on(sheet: &Sheet) -> Option<u32> {
    sheet
        .other_attrs
        .iter()
        .find(|(name, _)| name == "BasedOn")
        .and_then(|(_, value)| value.parse().ok())
}
fn style_attribute(sheet: &Sheet, attribute: &str) -> Option<u32> {
    sheet
        .other_attrs
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(attribute))
        .and_then(|(_, value)| value.parse().ok())
}
/// Classifies style-owned cells per MS-VSDX 2.2.5; Geometry never inherits styles.
fn style_owner(name: &str) -> Option<Provenance> {
    const LINE: &[&str] = &[
        "LineColor",
        "LinePattern",
        "LineWeight",
        "LineCap",
        "BeginArrow",
        "EndArrow",
        "BeginArrowSize",
        "EndArrowSize",
        "LineColorTrans",
        "LinePatternTrans",
        "CompoundType",
        "Rounding",
        "LineGradientDir",
        "LineGradientAngle",
        "LineGradientEnabled",
    ];
    const FILL: &[&str] = &[
        "FillForegnd",
        "FillBkgnd",
        "FillPattern",
        "FillForegndTrans",
        "FillBkgndTrans",
        "FillGradientDir",
        "FillGradientAngle",
        "FillGradientEnabled",
        "FillGradientStops",
        "FillGradientStopCount",
        "ShdwForegnd",
        "ShdwForegndTrans",
        "ShdwBkgnd",
        "ShdwBkgndTrans",
        "ShdwPattern",
        "ShdwOffsetX",
        "ShdwOffsetY",
        "ShdwType",
        "ShdwObliqueAngle",
        "ShdwScaleFactor",
        "ShapeShdwType",
        "ShapeShdwOffsetX",
        "ShapeShdwOffsetY",
    ];
    const TEXT: &[&str] = &[
        "Char",
        "Para",
        "Text",
        "Font",
        "Color",
        "Size",
        "Style",
        "Case",
        "Pos",
        "FontScale",
        "Letterspace",
        "ColorTrans",
        "Locale",
        "HorzAlign",
        "IndFirst",
        "IndLeft",
        "IndRight",
        "SpLine",
        "SpBefore",
        "SpAfter",
        "BulletStr",
        "Bullet",
        "Flags",
        "VerticalAlign",
        "TxtPinX",
        "TxtPinY",
        "TxtWidth",
        "TxtHeight",
        "TxtAngle",
        "TxtLocPinX",
        "TxtLocPinY",
        "LeftMargin",
        "RightMargin",
        "TopMargin",
        "BottomMargin",
        "TextBkgnd",
        "DefaultTabStop",
        "TextDirection",
        "TextBlockVerticalAlign",
    ];
    if LINE.contains(&name) {
        Some(Provenance::StyleLine)
    } else if FILL.contains(&name) {
        Some(Provenance::StyleFill)
    } else if TEXT.contains(&name) {
        Some(Provenance::StyleText)
    } else {
        None
    }
}
trait HasCells {
    fn cells(&self) -> Box<dyn Iterator<Item = &Cell> + '_>;
}
impl HasCells for Shape {
    fn cells(&self) -> Box<dyn Iterator<Item = &Cell> + '_> {
        Box::new(self.cells())
    }
}
impl HasCells for Sheet {
    fn cells(&self) -> Box<dyn Iterator<Item = &Cell> + '_> {
        Box::new(self.cells())
    }
}
trait HasSections {
    fn sections(&self) -> Box<dyn Iterator<Item = &Section> + '_>;
}
impl HasSections for Shape {
    fn sections(&self) -> Box<dyn Iterator<Item = &Section> + '_> {
        Box::new(self.sections())
    }
}
impl HasSections for Sheet {
    fn sections(&self) -> Box<dyn Iterator<Item = &Section> + '_> {
        Box::new(self.sections())
    }
}
fn row_base_key(row: &Row) -> String {
    row.name
        .clone()
        .map(|name| format!("N:{name}"))
        .or_else(|| row.index.map(|index| format!("IX:{index}")))
        .unwrap_or_else(|| "row".into())
}

fn row_keys(section: &Section) -> Vec<String> {
    let mut occurrences = BTreeMap::new();
    section
        .rows()
        .map(|row| {
            let key = row_base_key(row);
            let occurrence = occurrences.entry(key.clone()).or_insert(0usize);
            let identity = format!("{key}\u{1f}{occurrence}");
            *occurrence += 1;
            identity
        })
        .collect()
}
fn resolved_row_key(identity: &str, rows: &BTreeMap<String, ResolvedRow>) -> String {
    let (key, occurrence) = identity.rsplit_once('\u{1f}').unwrap_or((identity, "0"));
    if !rows.contains_key(key) {
        key.into()
    } else {
        format!("{key}\u{1f}{occurrence}")
    }
}
