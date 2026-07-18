use std::fmt;

use crate::{CellAddress, CellRef, SheetId};

#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    Package(String),
    Spreadsheet(xlsx_parse::ParseError),
    Operation(xlsx_ops::OpError),
    NoSheets,
    DuplicatePart(String),
    SheetOutOfRange(SheetId),
    CellOutOfRange(CellRef),
    InvalidOperation(String),
    ProposalNotFound(String),
    StaleProposal(Vec<CellAddress>),
    RangeTooLarge {
        rows: u64,
        cols: u64,
        max: u64,
    },
    InvalidViewport,
    DisplayTooLarge {
        cells: u64,
        max: u64,
    },
    InvalidScale(f32),
    RenderTooLarge {
        width: u32,
        height: u32,
        max: u32,
    },
    RenderAreaTooLarge {
        width: u32,
        height: u32,
        max_pixels: u64,
    },
    Raster(String),
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => f.write_str(error),
            Self::Spreadsheet(error) => error.fmt(f),
            Self::Operation(error) => error.fmt(f),
            Self::NoSheets => f.write_str("workbook has no sheets"),
            Self::DuplicatePart(name) => write!(f, "duplicate package part: {name}"),
            Self::SheetOutOfRange(sheet) => write!(f, "sheet {} out of range", sheet.0),
            Self::CellOutOfRange(cell) => write!(
                f,
                "cell row {}, column {} is out of range",
                u64::from(cell.row) + 1,
                u64::from(cell.col) + 1
            ),
            Self::InvalidOperation(message) => f.write_str(message),
            Self::ProposalNotFound(id) => write!(f, "no proposal {id}"),
            Self::StaleProposal(cells) => {
                let cells = cells
                    .iter()
                    .map(|address| address.cell.to_a1())
                    .collect::<Vec<_>>()
                    .join(", ");
                write!(f, "stale: {cells}")
            }
            Self::RangeTooLarge { rows, cols, max } => {
                write!(f, "range {rows}x{cols} exceeds the {max}-cell copy cap")
            }
            Self::InvalidViewport => f.write_str("viewport must have finite positive dimensions"),
            Self::DisplayTooLarge { cells, max } => write!(
                f,
                "requested viewport spans {cells} cells, exceeds the {max}-cell display-list cap"
            ),
            Self::InvalidScale(scale) => {
                write!(f, "scale must be a positive number, got {scale}")
            }
            Self::RenderTooLarge { width, height, max } => write!(
                f,
                "requested render is {width}x{height}px, exceeds the {max}px per-side cap; narrow the range or lower scale"
            ),
            Self::RenderAreaTooLarge {
                width,
                height,
                max_pixels,
            } => write!(
                f,
                "requested render is {width}x{height}px, exceeds the {max_pixels}-pixel allocation cap; narrow the range or lower scale"
            ),
            Self::Raster(error) => f.write_str(error),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Spreadsheet(error) => Some(error),
            Self::Operation(error) => Some(error),
            _ => None,
        }
    }
}

impl From<xlsx_parse::ParseError> for Error {
    fn from(error: xlsx_parse::ParseError) -> Self {
        Self::Spreadsheet(error)
    }
}

impl From<xlsx_ops::OpError> for Error {
    fn from(error: xlsx_ops::OpError) -> Self {
        Self::Operation(error)
    }
}
