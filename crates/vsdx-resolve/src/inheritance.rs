use std::collections::{BTreeMap, HashSet};

use vsdx_parse::{Cell, Row, Section, Shape, Sheet, TextToken, VsdxPackage};

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
            defaults: BTreeMap::new(),
        }
    }
    pub fn with_defaults(mut self, defaults: impl IntoIterator<Item = Cell>) -> Self {
        self.defaults = defaults.into_iter().map(|c| (c.name.clone(), c)).collect();
        self
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
    pub fn resolve_text(
        &self,
        shape: &Shape,
        page: &Sheet,
    ) -> Result<Vec<ResolvedTextToken>, ResolveError> {
        let resolved = self.resolve_shape_ref(shape, page)?;
        Ok(shape
            .text()
            .unwrap_or_default()
            .iter()
            .map(|token| match token {
                TextToken::Literal(value) => ResolvedTextToken::Literal(value.clone()),
                TextToken::CharacterRun(index) => ResolvedTextToken::CharacterRun {
                    index: *index,
                    properties: row_cells(&resolved, "Character", *index),
                },
                TextToken::ParagraphRun(index) => ResolvedTextToken::ParagraphRun {
                    index: *index,
                    properties: row_cells(&resolved, "Paragraph", *index),
                },
                TextToken::Tab(index) => ResolvedTextToken::Tab {
                    index: *index,
                    properties: row_cells(&resolved, "Tabs", *index),
                },
                TextToken::Field(index) => ResolvedTextToken::Field {
                    index: *index,
                    properties: row_cells(&resolved, "Field", *index),
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
        let masters = self.master_chain(shape)?;
        let styles = self.style_chains(shape)?;
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
            sections.extend(source.sections().map(|s| s.name.clone()));
        }
        for section in sections {
            out.sections.insert(
                section.clone(),
                self.resolve_section(shape, &masters, &styles, page, &section),
            );
        }
        Ok(out)
    }
    fn master_chain(&self, shape: &Shape) -> Result<Vec<(Provenance, &'a Shape)>, ResolveError> {
        let mut out = Vec::new();
        let mut current = shape;
        let mut seen = HashSet::new();
        for depth in 0..MAX_INHERITANCE_DEPTH {
            let Some(master_id) = current.master else {
                return Ok(out);
            };
            let Some(path) = self
                .package
                .master_part_ids
                .iter()
                .find_map(|(path, id)| (*id == master_id).then_some(path))
            else {
                return Ok(out);
            };
            let Some(sheet) = self.package.master_contents.get(path) else {
                return Ok(out);
            };
            let id = current.master_shape.unwrap_or(master_id);
            if !seen.insert((master_id, id)) {
                return Err(ResolveError::Cycle(format!("master {master_id}/{id}")));
            }
            let Some(next) = find_shape(sheet, id).or_else(|| find_shape(sheet, master_id)) else {
                return Ok(out);
            };
            out.push((
                if current.master_shape.is_some() {
                    Provenance::MasterShape
                } else {
                    Provenance::Master
                },
                next,
            ));
            current = next;
            if depth + 1 == MAX_INHERITANCE_DEPTH {
                return Err(ResolveError::Cycle("maximum inheritance depth".into()));
            }
        }
        unreachable!()
    }
    fn style_chains(
        &self,
        shape: &Shape,
    ) -> Result<Vec<(Provenance, Vec<&'a Sheet>)>, ResolveError> {
        [
            (shape.line_style, Provenance::StyleLine),
            (shape.fill_style, Provenance::StyleFill),
            (shape.text_style, Provenance::StyleText),
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
        for (provenance, cell) in sources {
            if let Some(cell) = cell {
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
            .unwrap_or(Lookup::Absent)
    }
    fn resolve_section(
        &self,
        shape: &Shape,
        masters: &[(Provenance, &'a Shape)],
        styles: &[(Provenance, Vec<&'a Sheet>)],
        page: &Sheet,
        name: &str,
    ) -> ResolvedSection {
        let mut sources: Vec<(Provenance, Option<&Section>)> =
            vec![(Provenance::Local, shape.sections().find(|s| s.name == name))];
        sources.extend(
            masters
                .iter()
                .map(|(p, s)| (*p, s.sections().find(|v| v.name == name))),
        );
        if (name == "Character" || name == "Paragraph" || name == "Tabs" || name == "Field")
            && let Some((p, chain)) = styles.iter().find(|(p, _)| *p == Provenance::StyleText)
        {
            sources.extend(
                chain
                    .iter()
                    .map(|s| (*p, s.sections().find(|v| v.name == name))),
            );
        }
        sources.push((Provenance::Page, page.sections().find(|s| s.name == name)));
        sources.push((
            Provenance::Document,
            self.package
                .document_sheet
                .as_ref()
                .and_then(|s| s.sections().find(|s| s.name == name)),
        ));
        if let Some(position) = sources
            .iter()
            .position(|(_, section)| section.is_some_and(|section| section.del))
        {
            if position == 0 {
                return ResolvedSection {
                    name: name.into(),
                    deleted: true,
                    ..Default::default()
                };
            }
            sources.truncate(position);
        }
        let mut keys = HashSet::new();
        for (_, section) in &sources {
            if let Some(section) = section {
                keys.extend(section.rows().map(row_key));
            }
        }
        let mut out = ResolvedSection {
            name: name.into(),
            ..Default::default()
        };
        for key in keys {
            let rows: Vec<(Provenance, Option<&Row>)> = sources
                .iter()
                .map(|(p, section)| {
                    (
                        *p,
                        section.and_then(|s| s.rows().find(|r| row_key(r) == key)),
                    )
                })
                .collect();
            let row = rows.iter().find_map(|(_, row)| *row);
            let Some(row) = row else { continue };
            if row.del {
                out.rows.insert(
                    key.clone(),
                    ResolvedRow {
                        key,
                        deleted: true,
                        ..Default::default()
                    },
                );
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
                let lookup = rows
                    .iter()
                    .filter(|(_, row)| row.is_none_or(|row| !row.del))
                    .find_map(|(p, row)| {
                        row.and_then(|row| {
                            row.cells()
                                .find(|c| c.name == cell_name)
                                .map(|c| found(c, *p))
                        })
                    });
                cells.insert(cell_name, lookup.unwrap_or(Lookup::Absent));
            }
            out.rows.insert(
                key.clone(),
                ResolvedRow {
                    key,
                    deleted: false,
                    row_type: row.row_type.clone(),
                    cells,
                },
            );
        }
        out
    }
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
    let mut pending: Vec<&Shape> = sheet.shapes().collect();
    let mut index = 0;
    while let Some(shape) = pending.get(index).copied() {
        index += 1;
        if shape.id == id {
            return Some(shape);
        }
        pending.extend(shape.shapes());
    }
    None
}
fn based_on(sheet: &Sheet) -> Option<u32> {
    sheet
        .other_attrs
        .iter()
        .find(|(name, _)| name == "BasedOn")
        .and_then(|(_, value)| value.parse().ok())
}
/// Built-in ownership is from MS-VSDX ShapeSheet style-sheet cells; unlisted cells bypass styles.
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
        "ShdwPattern",
        "ShapeShdwType",
        "ShapeShdwOffsetX",
        "ShapeShdwOffsetY",
    ];
    const TEXT: &[&str] = &[
        "Char",
        "Para",
        "Text",
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
fn row_key(row: &Row) -> String {
    row.name
        .clone()
        .map(|name| format!("N:{name}"))
        .or_else(|| row.index.map(|index| format!("IX:{index}")))
        .unwrap_or_default()
}
