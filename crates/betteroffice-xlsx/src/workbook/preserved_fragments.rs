//! The guard that keeps a sheet rename from leaving preserved formulas behind:
//! data validations and extension formulas the workbook keeps as the file
//! wrote them, which the save renames by sheet identity. A rename they cannot
//! follow is refused before anything changes.

use xlsx_parse::{FragmentFailureReason, FragmentKind};

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
        if preview
            .sheets
            .iter()
            .map(|sheet| &sheet.name)
            .eq(self.model.sheets.iter().map(|sheet| &sheet.name))
        {
            return Ok(());
        }
        let package = package.materialize()?;
        let Some(failure) = xlsx_parse::preserved_fragment_failures(preview, package, origins)
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
