use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, Workbook};
use serde_json::Value;

use crate::Result;

pub fn open(bytes: &[u8]) -> Result<Workbook> {
    Ok(Workbook::open(bytes)?)
}

fn value(document: &Workbook, probe: &Value) -> Result<f64> {
    let name = probe["sheet"].as_str().ok_or("Missing worksheet")?;
    let (_, sheet) = document
        .model()
        .sheet_by_name(name)
        .ok_or("Missing worksheet")?;
    let at = CellRef::parse_a1(probe["cell"].as_str().ok_or("Missing cell")?)
        .map_err(|error| format!("Invalid cell: {error:?}"))?;
    let cell = sheet.cell(at).ok_or("Missing original cell")?;
    match cell.value {
        CellValue::Number { value } if cell.formula.is_none() => Ok(value),
        _ => Err("The probe requires a numeric literal".into()),
    }
}

pub fn edit(document: &mut Workbook, probe: &Value) -> Result<()> {
    if value(document, probe)?
        != probe["old"]
            .as_str()
            .ok_or("Missing original value")?
            .parse::<f64>()?
    {
        return Err("Original cell differs from the probe".into());
    }
    let (sheet, _) = document
        .model()
        .sheet_by_name(probe["sheet"].as_str().ok_or("Missing worksheet")?)
        .ok_or("Missing worksheet")?;
    let cell = CellRef::parse_a1(probe["cell"].as_str().ok_or("Missing cell")?)
        .map_err(|error| format!("Invalid cell: {error:?}"))?;
    document.edit_cell(
        sheet,
        cell,
        probe["new"].as_str().ok_or("Missing replacement value")?,
        CalculationOptions::default(),
    )?;
    Ok(())
}

pub fn save(document: &Workbook) -> Result<Vec<u8>> {
    Ok(document.save()?)
}

pub fn verify(document: &Workbook, probe: &Value) -> Result<()> {
    if value(document, probe)?
        != probe["new"]
            .as_str()
            .ok_or("Missing replacement value")?
            .parse::<f64>()?
    {
        return Err("Edited cell did not survive reopening".into());
    }
    Ok(())
}
