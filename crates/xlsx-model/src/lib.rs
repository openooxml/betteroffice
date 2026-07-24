//! workbook model: address types, cell values, and the cell-access trait the
//! calc engine reads through. pure data — no io, no xml, no dom.

pub mod addr;
pub mod date;
pub mod numfmt;
pub mod styles;
pub mod value;
pub mod workbook;

pub use addr::{CellRange, CellRef, ColId, MAX_COLS, MAX_ROWS, RowId, SheetId};
pub use date::DateSystem;
pub use styles::{
    Alignment, Border, BorderEdge, BorderStyle, CellFormat, Color, Fill, Font, FormatCode, HAlign,
    NumberFormat, Stylesheet, Theme, VAlign, Xf,
};
pub use value::{CellValue, ErrorValue};
pub use workbook::{Cell, CellProvider, DefinedName, FreezePane, Hyperlink, Sheet, Workbook};
