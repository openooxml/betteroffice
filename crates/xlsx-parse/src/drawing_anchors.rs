//! The anchors in a sheet's drawing that no chart claims (pictures, shapes, connectors, groups
//! and frames without a chart relationship). The model does not hold them, so a save moves them
//! through the sheet's [`SheetAxes`], by the rules chart anchors follow, and an edit a save could
//! not place is refused before it applies.

use xlsx_model::addr::{MAX_COLS, MAX_ROWS};
use xlsx_model::{AnchorCell, AnchorEditAs, ChartAnchor};

use crate::ParseError;
use crate::axis::SheetAxes;
use crate::chart::{
    NS_SPREADSHEET_DRAWING, anchor_geometry, holds_related_chart, is_anchor, sheet_drawings,
};
use crate::tree::Part;

/// One drawing part a source sheet relates to.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SheetDrawing {
    pub(crate) part: String,
    /// The anchors no chart claims, in document order; `None` when the part holds markup this
    /// module cannot place, which then refuses every row or column edit on the sheet.
    pub(crate) anchors: Option<Vec<ChartAnchor>>,
}

/// A row or column insert or delete, in current coordinates.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[doc(hidden)]
pub struct GridShift {
    pub rows: bool,
    pub at: u32,
    pub count: u32,
    pub inserting: bool,
}

/// The drawings of the sheet stored at `sheet_path` that hold anything a row or column edit
/// could move. A sheet whose relationships cannot be read yields one unreadable entry.
pub(crate) fn sheet_drawing_anchors(
    parts: &[(String, Vec<u8>)],
    sheet_path: &str,
) -> Vec<SheetDrawing> {
    let Ok(drawings) = sheet_drawings(parts, sheet_path) else {
        return vec![SheetDrawing {
            part: crate::chart::relationship_part_path(sheet_path),
            anchors: None,
        }];
    };
    drawings
        .into_iter()
        .map(|(part, bytes)| SheetDrawing {
            anchors: unclaimed_anchors(bytes).ok(),
            part,
        })
        .filter(|drawing| {
            drawing
                .anchors
                .as_ref()
                .is_none_or(|anchors| !anchors.is_empty())
        })
        .collect()
}

/// Every anchor of a drawing part that no chart claims. Any child of the drawing that is not an
/// anchor (an `mc:AlternateContent` wrapper, say) is markup that may name cells, so it makes the
/// part unreadable rather than being left where it was.
fn unclaimed_anchors(bytes: &[u8]) -> Result<Vec<ChartAnchor>, ParseError> {
    let part = Part::decode(bytes)?;
    let root = part.tree()?;
    if !root.is(NS_SPREADSHEET_DRAWING, "wsDr") {
        return Err(ParseError::Malformed(
            "drawing part has no wsDr root".into(),
        ));
    }
    let mut anchors = Vec::new();
    for child in root.child_elements() {
        if !is_anchor(&child) {
            return Err(ParseError::UnsupportedEdit(format!(
                "drawing holds {} beside its anchors",
                child.local_name()
            )));
        }
        if !holds_related_chart(child) {
            anchors.push(anchor_geometry(child)?);
        }
    }
    Ok(anchors)
}

/// Where `anchor`, authored in source coordinates, sits once the sheet's rows and columns moved
/// as `axes` records: a corner on a deleted row or column moves to the deletion point, a one-cell
/// mode keeps the size of a two-cell anchor, and an anchor that does not move with cells stays.
/// `None` when a corner would leave the grid.
pub(crate) fn followed(anchor: &ChartAnchor, axes: &SheetAxes) -> Option<ChartAnchor> {
    let cell = |cell: AnchorCell| {
        Some(AnchorCell {
            col: axes.cols.current_at_or_after(cell.col)?,
            row: axes.rows.current_at_or_after(cell.row)?,
            ..cell
        })
    };
    Some(match *anchor {
        ChartAnchor::Absolute { .. }
        | ChartAnchor::TwoCell {
            edit_as: AnchorEditAs::Absolute,
            ..
        } => *anchor,
        ChartAnchor::OneCell { from, extent } => ChartAnchor::OneCell {
            from: cell(from)?,
            extent,
        },
        ChartAnchor::TwoCell {
            from,
            to,
            edit_as: AnchorEditAs::TwoCell,
        } => ChartAnchor::TwoCell {
            from: cell(from)?,
            to: cell(to)?,
            edit_as: AnchorEditAs::TwoCell,
        },
        ChartAnchor::TwoCell {
            from,
            to,
            edit_as: AnchorEditAs::OneCell,
        } => {
            let moved = cell(from)?;
            ChartAnchor::TwoCell {
                from: moved,
                to: AnchorCell {
                    col: translate(to.col, from.col, moved.col, MAX_COLS)?,
                    row: translate(to.row, from.row, moved.row, MAX_ROWS)?,
                    ..to
                },
                edit_as: AnchorEditAs::OneCell,
            }
        }
    })
}

fn translate(index: u32, from: u32, moved: u32, limit: u32) -> Option<u32> {
    let target = i64::from(index) + i64::from(moved) - i64::from(from);
    (0..i64::from(limit))
        .contains(&target)
        .then_some(target as u32)
}

/// Whether `shift` would leave one of `drawing`'s anchors where a save cannot place it, on a sheet
/// whose source rows and columns sit where `axes` says. Refused: a part that cannot be read; a
/// sheet without axes; a cell anchor that does not move with cells, at or after the edit; an
/// object wholly inside deleted rows or columns, or a one-cell anchor whose cell is deleted; and a
/// corner the edit would push off the grid.
pub(crate) fn blocks(drawing: &SheetDrawing, axes: Option<&SheetAxes>, shift: GridShift) -> bool {
    let (Some(anchors), Some(axes)) = (&drawing.anchors, axes) else {
        return true;
    };
    let mut after = axes.clone();
    let map = if shift.rows {
        &mut after.rows
    } else {
        &mut after.cols
    };
    if shift.inserting {
        map.insert(shift.at, shift.count);
    } else {
        map.delete(shift.at, shift.count);
    }
    anchors.iter().any(|anchor| {
        let Some(current) = followed(anchor, axes) else {
            return true;
        };
        refused(&current, shift) || followed(anchor, &after).is_none()
    })
}

fn refused(current: &ChartAnchor, shift: GridShift) -> bool {
    let index = |cell: AnchorCell| if shift.rows { cell.row } else { cell.col };
    let deleted = |cell: AnchorCell| {
        !shift.inserting && index(cell) >= shift.at && index(cell) - shift.at < shift.count
    };
    match *current {
        ChartAnchor::Absolute { .. } => false,
        ChartAnchor::OneCell { from, .. } => deleted(from),
        ChartAnchor::TwoCell {
            from,
            to,
            edit_as: AnchorEditAs::Absolute,
        } => index(from) >= shift.at || index(to) >= shift.at,
        ChartAnchor::TwoCell { from, to, .. } => deleted(from) && deleted(to),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell(col: u32, row: u32) -> AnchorCell {
        AnchorCell {
            col,
            col_off: 7,
            row,
            row_off: 9,
        }
    }

    fn two(from: (u32, u32), to: (u32, u32), edit_as: AnchorEditAs) -> ChartAnchor {
        ChartAnchor::TwoCell {
            from: cell(from.0, from.1),
            to: cell(to.0, to.1),
            edit_as,
        }
    }

    fn drawing(anchors: Vec<ChartAnchor>) -> SheetDrawing {
        SheetDrawing {
            part: "xl/drawings/drawing1.xml".into(),
            anchors: Some(anchors),
        }
    }

    fn rows(at: u32, count: u32, inserting: bool) -> GridShift {
        GridShift {
            rows: true,
            at,
            count,
            inserting,
        }
    }

    #[test]
    fn corners_follow_composed_edits_and_keep_their_offsets() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(0, 3);
        axes.rows.delete(5, 2);
        axes.cols.insert(2, 1);
        let anchor = two((1, 4), (3, 8), AnchorEditAs::TwoCell);
        assert_eq!(
            followed(&anchor, &axes),
            Some(two((1, 5), (4, 9), AnchorEditAs::TwoCell))
        );
    }

    #[test]
    fn a_one_cell_mode_keeps_its_span() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(6, 4);
        let anchor = two((1, 4), (3, 8), AnchorEditAs::OneCell);
        assert_eq!(followed(&anchor, &axes), Some(anchor));
        axes.rows.insert(0, 2);
        assert_eq!(
            followed(&anchor, &axes),
            Some(two((1, 6), (3, 10), AnchorEditAs::OneCell))
        );
    }

    #[test]
    fn refusals_follow_the_current_position() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(0, 10);
        let pinned = drawing(vec![two((1, 4), (3, 8), AnchorEditAs::Absolute)]);
        assert!(blocks(&pinned, Some(&axes), rows(9, 1, true)));
        assert!(blocks(&pinned, Some(&axes), rows(18, 1, true)));
        assert!(!blocks(&pinned, Some(&axes), rows(19, 1, true)));
        let moving = drawing(vec![two((1, 4), (3, 6), AnchorEditAs::TwoCell)]);
        assert!(blocks(&moving, Some(&axes), rows(14, 3, false)));
        assert!(!blocks(&moving, Some(&axes), rows(14, 2, false)));
        assert!(!blocks(&moving, Some(&axes), rows(4, 3, false)));
        assert!(blocks(&moving, None, rows(30, 1, true)));
        let unreadable = SheetDrawing {
            part: "xl/drawings/drawing1.xml".into(),
            anchors: None,
        };
        assert!(blocks(&unreadable, Some(&axes), rows(30, 1, true)));
    }

    #[test]
    fn an_insert_that_would_push_a_corner_off_the_grid_blocks() {
        let edge = drawing(vec![two(
            (1, MAX_ROWS - 3),
            (3, MAX_ROWS - 1),
            AnchorEditAs::TwoCell,
        )]);
        assert!(blocks(&edge, Some(&SheetAxes::default()), rows(0, 1, true)));
        assert!(!blocks(
            &edge,
            Some(&SheetAxes::default()),
            rows(0, 1, false)
        ));
    }
}
