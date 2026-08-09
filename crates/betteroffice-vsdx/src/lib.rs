//! Typed native facade for inspecting VSDX diagrams.

use vsdx_eval::{
    DocumentReferences, Evaluation, MutationContext, MutationOutcome, Value, decide_mutation,
};
pub use vsdx_parse::StructuralEdit;
use vsdx_parse::{Cell, ParseLimits, Shape, VsdxError, VsdxPackage};
pub use vsdx_parse::{CellLocator, CellRow, CellSheet, MutationGesture, SemanticCellEdit};
use vsdx_resolve::{PageConnectivity, ResolveError, ResolvedShape, Resolver};

#[derive(Debug)]
pub enum Error {
    Parse(VsdxError),
    Resolve(ResolveError),
    Policy(String),
}

impl From<VsdxError> for Error {
    fn from(value: VsdxError) -> Self {
        Self::Parse(value)
    }
}
impl From<ResolveError> for Error {
    fn from(value: ResolveError) -> Self {
        Self::Resolve(value)
    }
}

pub type Result<T> = std::result::Result<T, Error>;

pub struct Diagram {
    package: VsdxPackage,
}

impl Diagram {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            package: vsdx_parse::parse_vsdx(bytes)?,
        })
    }
    pub fn open_with_limits(bytes: &[u8], limits: &ParseLimits) -> Result<Self> {
        Ok(Self {
            package: vsdx_parse::parse_vsdx_with_limits(bytes, limits)?,
        })
    }
    pub fn package(&self) -> &VsdxPackage {
        &self.package
    }
    /// Writes formulas to Cell@F and their evaluated numeric cache to Cell@V.
    pub fn save_cell_edits(&self, edits: &[SemanticCellEdit]) -> Result<Vec<u8>> {
        let resolved = self.resolve_cell_edits(edits)?;
        Ok(vsdx_parse::save_semantic_cell_edits(
            &self.package,
            &resolved,
        )?)
    }
    /// Saves the formula changes accumulated by a collaborative session.
    pub fn save_session(&self, session: &vsdx_edit::DiagramSession) -> Result<Vec<u8>> {
        let edits = session
            .semantic_cell_edits()
            .map_err(|error| Error::Policy(error.to_string()))?;
        self.save_cell_edits(&edits)
    }
    /// Applies semantic and structural edits as one all-or-nothing save request.
    pub fn save_edits(
        &self,
        cell_edits: &[SemanticCellEdit],
        structural_edits: &[StructuralEdit],
    ) -> Result<Vec<u8>> {
        let resolved = self.resolve_cell_edits(cell_edits)?;
        self.authorize_structural_edits(structural_edits)?;
        let cell_saved = vsdx_parse::save_semantic_cell_edits(&self.package, &resolved)?;
        let updated = vsdx_parse::parse_vsdx(&cell_saved)?;
        Ok(vsdx_parse::save_structural_edits(
            &updated,
            structural_edits,
        )?)
    }

    fn resolve_cell_edits(&self, edits: &[SemanticCellEdit]) -> Result<Vec<SemanticCellEdit>> {
        let context = PackageMutationContext {
            package: &self.package,
        };
        let mut resolved = Vec::with_capacity(edits.len());
        for edit in edits {
            if edit.value.is_some() {
                return Err(Error::Policy(
                    "semantic edits write formulas; raw Cell@V edits are not allowed".to_owned(),
                ));
            }
            let formula = edit
                .formula
                .clone()
                .ok_or_else(|| Error::Policy("semantic edits require a formula".to_owned()))?;
            match decide_mutation(
                &context,
                edit.locator.clone(),
                edit.gesture,
                formula,
                &ParseLimits::default(),
            ) {
                MutationOutcome::Allowed { target, formula } => {
                    let value = context.evaluate_formula(&target, &formula)?;
                    resolved.push(SemanticCellEdit {
                        locator: target,
                        gesture: edit.gesture,
                        formula: Some(formula),
                        value: Some(value),
                    });
                }
                MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                    return Err(Error::Policy(reason));
                }
            }
        }
        Ok(resolved)
    }
    /// Deletes shapes after enforcing their effective LockDelete cells.
    pub fn save_structural_edits(&self, edits: &[StructuralEdit]) -> Result<Vec<u8>> {
        self.authorize_structural_edits(edits)?;
        Ok(vsdx_parse::save_structural_edits(&self.package, edits)?)
    }

    fn authorize_structural_edits(&self, edits: &[StructuralEdit]) -> Result<()> {
        let context = PackageMutationContext {
            package: &self.package,
        };
        for edit in edits {
            let StructuralEdit::DeleteShape { page_id, shape_id } = edit else {
                continue;
            };
            let locator = CellLocator {
                sheet: CellSheet::Page(*page_id),
                shape_id: Some(*shape_id),
                section: None,
                row: None,
                cell_name: "LockDelete".to_owned(),
            };
            match decide_mutation(
                &context,
                locator,
                MutationGesture::Delete,
                "0".to_owned(),
                &ParseLimits::default(),
            ) {
                MutationOutcome::Allowed { .. } => {}
                MutationOutcome::Refused { reason } | MutationOutcome::Unsupported { reason } => {
                    return Err(Error::Policy(reason));
                }
            }
        }
        Ok(())
    }
    pub fn pages(&self) -> impl Iterator<Item = Page<'_>> {
        self.package.page_contents.keys().map(|part| Page {
            diagram: self,
            part,
        })
    }
}

struct PackageMutationContext<'a> {
    package: &'a VsdxPackage,
}

impl PackageMutationContext<'_> {
    fn resolved_cell(&self, locator: &CellLocator) -> std::result::Result<Option<Cell>, String> {
        let resolver = Resolver::new(self.package);
        let key = locator_key(locator);
        let resolved = match (&locator.sheet, locator.shape_id) {
            (CellSheet::Page(_), Some(shape_id)) => resolver
                .resolve_shape(&sheet_path(self.package, &locator.sheet)?, shape_id)
                .map_err(|error| error.to_string())?,
            (CellSheet::Page(page_id), None) => resolver
                .resolve_sheet(
                    self.package
                        .page_sheets
                        .get(page_id)
                        .ok_or_else(|| "page sheet does not exist".to_owned())?,
                )
                .map_err(|error| error.to_string())?,
            (CellSheet::Document, None) => resolver
                .resolve_sheet(
                    self.package
                        .document_sheet
                        .as_ref()
                        .ok_or_else(|| "document sheet does not exist".to_owned())?,
                )
                .map_err(|error| error.to_string())?,
            (CellSheet::Master(_), _) | (CellSheet::Document, Some(_)) => {
                return Err("mutation policy does not yet support this sheet identity".to_owned());
            }
        };
        Ok(resolved.cell(&key).and_then(|lookup| match lookup {
            vsdx_resolve::Lookup::Found(cell) => Some(cell.cell.clone()),
            vsdx_resolve::Lookup::Deleted | vsdx_resolve::Lookup::Absent => None,
        }))
    }

    fn evaluate_formula(&self, locator: &CellLocator, formula: &str) -> Result<String> {
        let resolver = Resolver::new(self.package);
        let evaluation = match (&locator.sheet, locator.shape_id) {
            (CellSheet::Page(_), Some(shape_id)) => {
                let page = sheet_path(self.package, &locator.sheet).map_err(Error::Policy)?;
                let references = vsdx_eval::PageShapeReferences::new(&resolver, &page)
                    .map_err(|error| Error::Policy(error.to_string()))?;
                vsdx_eval::evaluate_cell(
                    &locator_key(locator),
                    formula.trim_start_matches('='),
                    &references.for_shape(shape_id),
                    &ParseLimits::default(),
                )
            }
            (CellSheet::Page(page_id), None) => {
                let page = resolver.resolve_sheet(
                    self.package
                        .page_sheets
                        .get(page_id)
                        .ok_or_else(|| Error::Policy("page sheet does not exist".to_owned()))?,
                )?;
                let document = self
                    .package
                    .document_sheet
                    .as_ref()
                    .map(|sheet| resolver.resolve_sheet(sheet))
                    .transpose()?;
                vsdx_eval::evaluate_cell(
                    &locator_key(locator),
                    formula.trim_start_matches('='),
                    &DocumentReferences::new(&page, document.as_ref()),
                    &ParseLimits::default(),
                )
            }
            (CellSheet::Document, None) => {
                let document =
                    resolver.resolve_sheet(self.package.document_sheet.as_ref().ok_or_else(
                        || Error::Policy("document sheet does not exist".to_owned()),
                    )?)?;
                vsdx_eval::evaluate_cell(
                    &locator_key(locator),
                    formula.trim_start_matches('='),
                    &DocumentReferences::new(&document, Some(&document)),
                    &ParseLimits::default(),
                )
            }
            (CellSheet::Master(_), _) | (CellSheet::Document, Some(_)) => {
                return Err(Error::Policy(
                    "formula cache updates do not support this sheet identity".to_owned(),
                ));
            }
        };
        match evaluation {
            Evaluation::Evaluated(value) => match value.value {
                Value::Number(number) => Ok(number.number.to_string()),
                Value::Color(_) => Err(Error::Policy(
                    "formula cache update does not support colour values".to_owned(),
                )),
            },
            Evaluation::Unsupported(reason) => Err(Error::Policy(format!(
                "formula cache update is unsupported: {reason}"
            ))),
            Evaluation::Error(error) => Err(Error::Policy(format!(
                "formula cache update failed: {}",
                error.message
            ))),
        }
    }
}

impl MutationContext for PackageMutationContext<'_> {
    fn current_formula(
        &self,
        locator: &CellLocator,
    ) -> std::result::Result<Option<String>, String> {
        Ok(self.resolved_cell(locator)?.and_then(|cell| cell.formula))
    }

    fn resolve_reference(
        &self,
        from: &CellLocator,
        reference: &str,
    ) -> std::result::Result<CellLocator, String> {
        let (sheet, shape_id, cell_name) = if let Some((shape_id, cell_name)) = reference
            .strip_prefix("Sheet.")
            .and_then(|reference| reference.split_once('!'))
        {
            (
                from.sheet.clone(),
                Some(
                    shape_id
                        .parse()
                        .map_err(|_| "invalid Sheet reference".to_owned())?,
                ),
                cell_name,
            )
        } else if let Some(cell_name) = reference.strip_prefix("ThePage!") {
            (from.sheet.clone(), None, cell_name)
        } else if let Some(cell_name) = reference.strip_prefix("TheDoc!") {
            (CellSheet::Document, None, cell_name)
        } else {
            (from.sheet.clone(), from.shape_id, reference)
        };
        let (section, row, cell_name) = split_reference(cell_name);
        let target = CellLocator {
            sheet,
            shape_id,
            section,
            row,
            cell_name,
        };
        self.resolved_cell(&target)?
            .ok_or_else(|| format!("SETATREF target does not exist: {reference}"))?;
        Ok(target)
    }

    fn lock_enabled(&self, locator: &CellLocator, lock: &str) -> std::result::Result<bool, String> {
        if lock.is_empty() {
            return Ok(false);
        }
        let Some(cell) = self.resolved_cell(&CellLocator {
            sheet: locator.sheet.clone(),
            shape_id: locator.shape_id,
            section: None,
            row: None,
            cell_name: lock.to_owned(),
        })?
        else {
            return Ok(false);
        };
        let expression = cell
            .formula
            .as_deref()
            .or(cell.value.as_deref())
            .ok_or_else(|| {
                format!("cannot evaluate {lock}: it has neither a formula nor a value")
            })?;
        let value = self
            .evaluate_formula(
                &CellLocator {
                    sheet: locator.sheet.clone(),
                    shape_id: locator.shape_id,
                    section: None,
                    row: None,
                    cell_name: lock.to_owned(),
                },
                expression,
            )
            .map_err(|error| match error {
                Error::Policy(reason) => format!("cannot evaluate {lock}: {reason}"),
                Error::Parse(error) => error.to_string(),
                Error::Resolve(error) => error.to_string(),
            })?;
        match value.as_str() {
            "0" => Ok(false),
            "1" => Ok(true),
            _ => Err(format!("cannot evaluate {lock} as a boolean")),
        }
    }
}

fn sheet_path(package: &VsdxPackage, sheet: &CellSheet) -> std::result::Result<String, String> {
    match sheet {
        CellSheet::Page(id) => package
            .page_part_ids
            .iter()
            .find_map(|(path, candidate)| (*candidate == *id).then(|| path.clone()))
            .ok_or_else(|| "page does not exist".to_owned()),
        CellSheet::Document => Ok(package.document_part_path.clone()),
        CellSheet::Master(_) => Err("master sheets are unsupported mutation targets".to_owned()),
    }
}

fn locator_key(locator: &CellLocator) -> String {
    match (&locator.section, &locator.row) {
        (Some(section), Some(CellRow::Name(row))) => {
            format!("{section}.{row}.{}", locator.cell_name)
        }
        (Some(section), Some(CellRow::Index(row))) => {
            format!("{section}.{}{}", locator.cell_name, row)
        }
        _ => locator.cell_name.clone(),
    }
}

fn split_reference(reference: &str) -> (Option<String>, Option<CellRow>, String) {
    let Some((section, rest)) = reference.split_once('.') else {
        return (None, None, reference.to_owned());
    };
    if section == "User" || section == "Prop" || section == "Actions" {
        return (
            Some(section.to_owned()),
            Some(CellRow::Name(rest.to_owned())),
            "Value".to_owned(),
        );
    }
    (None, None, reference.to_owned())
}

pub struct Page<'a> {
    diagram: &'a Diagram,
    part: &'a String,
}

impl<'a> Page<'a> {
    pub fn part_path(&self) -> &str {
        self.part
    }
    pub fn shapes(&'a self) -> impl Iterator<Item = ShapeView<'a>> {
        self.diagram.package.page_contents[self.part]
            .shapes()
            .map(move |shape| ShapeView { page: self, shape })
    }
    pub fn connectivity(&self) -> Result<PageConnectivity> {
        Ok(Resolver::new(&self.diagram.package).resolve_page_connectivity(self.part)?)
    }
}

pub struct ShapeView<'a> {
    page: &'a Page<'a>,
    shape: &'a Shape,
}

impl<'a> ShapeView<'a> {
    pub fn model(&self) -> &Shape {
        self.shape
    }
    pub fn resolved(&self) -> Result<ResolvedShape> {
        Ok(Resolver::new(&self.page.diagram.package)
            .resolve_shape(self.page.part, self.shape.id)?)
    }
}

pub type PageRef<'a> = Page<'a>;
pub type ShapeRef<'a> = ShapeView<'a>;
