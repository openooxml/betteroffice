//! The guard that keeps a sheet rename from leaving preserved formulas behind:
//! data validations and extension formulas the workbook keeps as the file
//! wrote them, which the save renames by sheet identity. A rename they cannot
//! follow is refused before anything changes.

use xlsx_parse::{FragmentFailureReason, FragmentKind};

use crate::{SaveWarning, SaveWarningCode};
use xlsx_model::SheetId;

use super::Workbook;
use crate::{Error, Result};

fn described(kind: FragmentKind) -> &'static str {
    match kind {
        FragmentKind::DataValidation => "data validation",
        FragmentKind::X14DataValidation => "extension data validation",
        FragmentKind::X14ConditionalFormat => "extension conditional format",
        FragmentKind::Sparkline => "sparkline",
        FragmentKind::Extension => "extension",
    }
}

impl Workbook {
    /// A warning for every preserved formula fragment the save keeps as the
    /// file wrote it because it cannot follow the sheet renames: on which
    /// sheet, in which part, and what kind. The guard refuses such a rename
    /// up front, so only a state that reached the workbook some other way
    /// gets here.
    pub(super) fn fragment_save_warnings(&self) -> Vec<SaveWarning> {
        let Some(package) = self.source_package.as_ref() else {
            return Vec::new();
        };
        let Ok(package) = package.materialize() else {
            return Vec::new();
        };
        package
            .fragments_not_following_renames(&self.model, &self.preserved.origins)
            .into_iter()
            .map(|failure| {
                let (code, why) = match failure.failure.reason {
                    FragmentFailureReason::UnwritableName => (
                        SaveWarningCode::UnwritableReference,
                        "the new sheet name holds a character XML cannot carry",
                    ),
                    FragmentFailureReason::Unreadable => (
                        SaveWarningCode::UnrewritableFormula,
                        "they do not read as formulas the rename can rewrite",
                    ),
                };
                SaveWarning {
                    code,
                    message: format!(
                        "the {} formulas in {} on sheet {} were kept as the file wrote them: {why}",
                        described(failure.failure.kind),
                        failure.part,
                        failure.sheet
                    ),
                    part: failure.part,
                    element: Some(failure.failure.kind.element().to_owned()),
                    sheet: Some(failure.sheet),
                }
            })
            .collect()
    }

    /// Refuses renaming sheet `index` of `staged` when it holds a name a
    /// source sheet held, that source sheet is gone, and a preserved formula
    /// names it. The formula binds to this sheet by name, as cell formulas
    /// do, but the save renames preserved formulas by source sheet identity,
    /// which this sheet does not have, so it could not follow the rename.
    pub(super) fn ensure_taken_name_rename_follows(
        &self,
        staged: &xlsx_model::Workbook,
        origins: &[Option<usize>],
        index: SheetId,
    ) -> Result<()> {
        let Some(package) = self.source_package.as_ref() else {
            return Ok(());
        };
        let at = index.0 as usize;
        let (Some(sheet), Some(None)) = (staged.sheets.get(at), origins.get(at)) else {
            return Ok(());
        };
        let package = package.materialize()?;
        let deleted_source_name = package
            .source_sheet_names()
            .enumerate()
            .any(|(source, name)| name == sheet.name && !origins.contains(&Some(source)));
        if !deleted_source_name {
            return Ok(());
        }
        let Some((source, kind)) = package.fragment_naming_sheet(&sheet.name) else {
            return Ok(());
        };
        let holder = origins
            .iter()
            .position(|origin| *origin == Some(source))
            .and_then(|at| staged.sheets.get(at))
            .map_or_else(|| sheet.name.clone(), |sheet| sheet.name.clone());
        Err(Error::InvalidOperation(format!(
            "the {} formulas on sheet '{holder}' name a deleted sheet whose name this sheet took, and cannot follow the sheet rename",
            described(kind)
        )))
    }

    /// Refuses `preview` when a sheet it renamed is named in a preserved
    /// formula the save cannot rewrite: the new name holds a character xml
    /// cannot carry, refused as cell formulas are, or the fragment does not
    /// read as formulas the rename can follow.
    pub(super) fn ensure_preserved_fragments_follow(
        &self,
        preview: &xlsx_model::Workbook,
        origins: &[Option<usize>],
    ) -> Result<()> {
        let Some(package) = self.source_package.as_ref() else {
            return Ok(());
        };
        let current = self
            .preserved
            .origins
            .iter()
            .zip(&self.model.sheets)
            .filter_map(|(origin, sheet)| Some(((*origin)?, sheet.name.as_str())))
            .collect::<std::collections::HashMap<_, _>>();
        if preview.sheets.iter().zip(origins).all(|(sheet, origin)| {
            origin.is_none_or(|origin| current.get(&origin) == Some(&sheet.name.as_str()))
        }) {
            return Ok(());
        }
        let package = package.materialize()?;
        let Some(failure) = package
            .fragments_not_following_renames(preview, origins)
            .into_iter()
            .next()
        else {
            return Ok(());
        };
        Err(match failure.failure.reason {
            FragmentFailureReason::UnwritableName => xlsx_ops::OpError::SheetNameNotWritable.into(),
            FragmentFailureReason::Unreadable => Error::InvalidOperation(format!(
                "the {} formulas on sheet '{}' cannot follow the sheet rename",
                described(failure.failure.kind),
                failure.sheet
            )),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::Workbook;
    use crate::SaveWarningCode;

    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const OTHER: &str = "xl/worksheets/sheet2.xml";

    /// A sheet `Data`, and a sheet `Other` whose data validation formula is
    /// `dv`, kept as the file wrote it.
    fn book(dv: &str) -> Vec<u8> {
        let parts = [
            ("[Content_Types].xml", r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned()),
            ("_rels/.rels", format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#)),
            ("xl/workbook.xml", format!(r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>"#)),
            ("xl/_rels/workbook.xml.rels", format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#)),
            ("xl/worksheets/sheet1.xml", format!(r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#)),
            (OTHER, format!(r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData><dataValidations count="1"><dataValidation type="list" sqref="B1"><formula1>{dv}</formula1></dataValidation></dataValidations></worksheet>"#)),
        ];
        let parts = parts
            .into_iter()
            .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
            .collect::<Vec<_>>();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn part(bytes: &[u8]) -> Vec<u8> {
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == OTHER)
            .map(|(_, bytes)| bytes)
            .unwrap()
    }

    /// Whatever door a rename the fragment cannot follow came through (the
    /// guard refuses every local one), the save keeps the fragment as the file
    /// wrote it and says so, naming the sheet, the part and the kind; an
    /// honest workbook carries no such warning.
    #[test]
    fn a_save_keeps_and_reports_a_fragment_that_cannot_follow_a_rename() {
        for (dv, name, code) in [
            ("Data!A1&lt;x/&gt;", "New Data", None),
            ("Data!A1<x/>", "New Data", Some(SaveWarningCode::UnrewritableFormula)),
            ("Data!$A$1:$A$5", "D\u{1}", Some(SaveWarningCode::UnwritableReference)),
        ] {
            let source = book(dv);
            let mut workbook = Workbook::open(&source).unwrap();
            assert_eq!(workbook.save_warnings(), [], "{dv}");
            workbook.model.sheets[0].name = name.to_owned();
            workbook.edited_since_open = true;
            let (saved, warnings) = workbook.save_with_warnings().unwrap();
            let Some(code) = code else {
                assert_eq!(warnings, [], "{dv}");
                continue;
            };
            assert_eq!(part(&saved), part(&source), "{dv}");
            assert_eq!(warnings.len(), 1, "{dv}: {warnings:?}");
            let warning = &warnings[0];
            assert_eq!(warning.code, code, "{dv}");
            assert_eq!(warning.part, OTHER);
            assert_eq!(warning.element.as_deref(), Some("data-validation"));
            assert_eq!(warning.sheet.as_deref(), Some("Other"));
            assert!(warning.message.contains("kept as the file wrote them"), "{}", warning.message);
            assert_eq!(workbook.save_warnings(), warnings);
        }
    }
}
