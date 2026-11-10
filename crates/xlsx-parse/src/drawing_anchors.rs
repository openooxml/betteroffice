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
    /// Whether another source sheet relates to the same part, so moving it for one sheet would
    /// move it on the other too.
    pub(crate) shared: bool,
}

/// What a row or column edit does to a sheet's drawings.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Verdict {
    /// Nothing the drawings hold moves.
    Still,
    /// Some anchor moves, and a save can place it.
    Moves,
    /// A save could not place some anchor where the edit leaves it.
    Refused,
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
            shared: false,
        }];
    };
    drawings
        .into_iter()
        .map(|(part, bytes)| SheetDrawing {
            anchors: unclaimed_anchors(bytes).ok(),
            part,
            shared: false,
        })
        .filter(|drawing| {
            drawing
                .anchors
                .as_ref()
                .is_none_or(|anchors| !anchors.is_empty())
        })
        .collect()
}

/// The relationship part of the part at `path`, as a save looks it up.
#[doc(hidden)]
pub fn relationship_part_of(path: &str) -> String {
    crate::chart::relationship_part_path(path)
}

/// The drawing parts the sheet at `sheet_path` relates to, from its relationship part's bytes,
/// read the way a save reads them; `None` when the part cannot be read.
#[doc(hidden)]
pub fn sheet_drawing_targets(rels: &[u8], sheet_path: &str) -> Option<Vec<String>> {
    crate::chart::sheet_drawing_paths(rels, sheet_path).ok()
}

/// Whether a row or column edit on the sheet at `sheet_path` could have to move or refuse
/// anything its drawings hold, judged from `parts` alone, which must carry the sheet's
/// relationship part (when it has one) and every drawing part it names. A drawing holding only
/// chart frames needs nothing; one holding pictures or shapes, or markup this module cannot
/// read, or a named part missing from `parts`, does.
#[doc(hidden)]
pub fn sheet_drawings_may_move(parts: &[(String, Vec<u8>)], sheet_path: &str) -> bool {
    let rels_path = crate::chart::relationship_part_path(sheet_path);
    let Some(rels) = crate::xml::find_part(parts, &rels_path) else {
        return false;
    };
    let Some(targets) = sheet_drawing_targets(rels, sheet_path) else {
        return true;
    };
    targets
        .iter()
        .any(|target| crate::xml::find_part(parts, target).is_none())
        || !sheet_drawing_anchors(parts, sheet_path).is_empty()
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

/// The anchor's place after one grid edit, by the rules chart anchors follow one op at a time:
/// an index at or after an insert moves along, an index inside a delete moves to its start, and
/// one after it moves back. `None` when an insert would push a corner off the grid.
fn shifted(anchor: &ChartAnchor, shift: GridShift) -> Option<ChartAnchor> {
    let limit = if shift.rows { MAX_ROWS } else { MAX_COLS };
    let index = |index: u32| {
        if shift.inserting {
            if index >= shift.at {
                index
                    .checked_add(shift.count)
                    .filter(|moved| *moved < limit)
            } else {
                Some(index)
            }
        } else if index >= shift.at.saturating_add(shift.count) {
            Some(index - shift.count)
        } else {
            Some(index.min(shift.at))
        }
    };
    let cell = |cell: AnchorCell| {
        let mut moved = cell;
        if shift.rows {
            moved.row = index(cell.row)?;
        } else {
            moved.col = index(cell.col)?;
        }
        Some(moved)
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

/// What `shift` does to `drawing` on a sheet whose source rows and columns sit where `axes` says.
/// Refused: a part that cannot be read, or that another sheet shares; a sheet without axes; a
/// cell anchor that does not move with cells, at or after the edit; an object wholly inside
/// deleted rows or columns, or a one-cell anchor whose cell is deleted; a corner pushed off the
/// grid; and an edit after which the save, mapping source corners through the composed axes,
/// would place an anchor anywhere but where the op-by-op rules put it. That last check keeps
/// every accepted sequence of edits saving exactly as the op-by-op rules move it.
pub(crate) fn verdict(
    drawing: &SheetDrawing,
    axes: Option<&SheetAxes>,
    shift: GridShift,
) -> Verdict {
    let (Some(anchors), Some(axes), false) = (&drawing.anchors, axes, drawing.shared) else {
        return Verdict::Refused;
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
    let mut verdict = Verdict::Still;
    for anchor in anchors {
        let Some(current) = followed(anchor, axes) else {
            return Verdict::Refused;
        };
        if refused(&current, shift) {
            return Verdict::Refused;
        }
        let Some(moved) = shifted(&current, shift) else {
            return Verdict::Refused;
        };
        if followed(anchor, &after) != Some(moved) {
            return Verdict::Refused;
        }
        if moved != current {
            verdict = Verdict::Moves;
        }
    }
    verdict
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
            shared: false,
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

    fn apply(axes: &mut SheetAxes, shift: GridShift) {
        if shift.inserting {
            axes.rows.insert(shift.at, shift.count);
        } else {
            axes.rows.delete(shift.at, shift.count);
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
    fn an_anchor_that_does_not_move_with_cells_refuses_edits_at_or_before_its_end() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(20, 5);
        let pinned = drawing(vec![two((1, 4), (3, 8), AnchorEditAs::Absolute)]);
        assert_eq!(
            verdict(&pinned, Some(&axes), rows(3, 1, true)),
            Verdict::Refused
        );
        assert_eq!(
            verdict(&pinned, Some(&axes), rows(8, 1, false)),
            Verdict::Refused
        );
        assert_eq!(
            verdict(&pinned, Some(&axes), rows(9, 1, true)),
            Verdict::Still
        );
    }

    #[test]
    fn refusals_follow_the_current_position() {
        let mut axes = SheetAxes::default();
        axes.rows.insert(0, 10);
        let moving = drawing(vec![two((1, 4), (3, 6), AnchorEditAs::TwoCell)]);
        assert_eq!(
            verdict(&moving, Some(&axes), rows(14, 3, false)),
            Verdict::Refused
        );
        assert_eq!(
            verdict(&moving, Some(&axes), rows(14, 2, false)),
            Verdict::Moves
        );
        assert_eq!(
            verdict(&moving, Some(&axes), rows(4, 3, false)),
            Verdict::Moves
        );
        assert_eq!(
            verdict(&moving, Some(&axes), rows(30, 1, true)),
            Verdict::Still
        );
        assert_eq!(verdict(&moving, None, rows(30, 1, true)), Verdict::Refused);
        let unreadable = SheetDrawing {
            anchors: None,
            ..drawing(Vec::new())
        };
        assert_eq!(
            verdict(&unreadable, Some(&axes), rows(30, 1, true)),
            Verdict::Refused
        );
        let shared = SheetDrawing {
            shared: true,
            ..moving
        };
        assert_eq!(
            verdict(&shared, Some(&axes), rows(30, 1, true)),
            Verdict::Refused
        );
    }

    #[test]
    fn an_insert_that_would_push_a_corner_off_the_grid_refuses() {
        let edge = drawing(vec![two(
            (1, MAX_ROWS - 3),
            (3, MAX_ROWS - 1),
            AnchorEditAs::TwoCell,
        )]);
        let axes = SheetAxes::default();
        assert_eq!(
            verdict(&edge, Some(&axes), rows(0, 1, true)),
            Verdict::Refused
        );
        assert_eq!(
            verdict(&edge, Some(&axes), rows(0, 1, false)),
            Verdict::Moves
        );
    }

    /// A delete that takes rows an earlier insert placed after a corner it also deletes: op by op
    /// the corner lands on the deletion point, while the composed axes would put it on the next
    /// source row. The edit is refused rather than saved anywhere else.
    #[test]
    fn a_sequence_the_composed_axes_cannot_reproduce_refuses() {
        let picture = drawing(vec![two((1, 4), (3, 8), AnchorEditAs::TwoCell)]);
        let mut axes = SheetAxes::default();
        let insert = rows(5, 3, true);
        assert_eq!(verdict(&picture, Some(&axes), insert), Verdict::Moves);
        apply(&mut axes, insert);
        assert_eq!(
            verdict(&picture, Some(&axes), rows(3, 3, false)),
            Verdict::Refused
        );
    }

    /// The other order is reproducible: a corner pulled to the deletion point and then pushed by
    /// an insert at that point sits on the next source row both ways.
    #[test]
    fn every_accepted_sequence_saves_where_the_op_by_op_rules_put_it() {
        let anchors = vec![
            two((1, 4), (3, 8), AnchorEditAs::TwoCell),
            two((5, 4), (7, 8), AnchorEditAs::OneCell),
            ChartAnchor::OneCell {
                from: cell(2, 12),
                extent: Default::default(),
            },
        ];
        let picture = drawing(anchors.clone());
        let sequences = [
            vec![rows(3, 3, false), rows(3, 2, true)],
            vec![rows(0, 2, true), rows(5, 2, false), rows(4, 1, true)],
            vec![rows(10, 4, true), rows(2, 1, false), rows(7, 2, false)],
            vec![rows(6, 1, false), rows(6, 1, false), rows(6, 3, true)],
        ];
        for sequence in sequences {
            let mut axes = SheetAxes::default();
            let mut op_by_op = anchors.clone();
            for shift in &sequence {
                if verdict(&picture, Some(&axes), *shift) == Verdict::Refused {
                    break;
                }
                apply(&mut axes, *shift);
                for anchor in &mut op_by_op {
                    *anchor = shifted(anchor, *shift).unwrap();
                }
                let composed = anchors
                    .iter()
                    .map(|anchor| followed(anchor, &axes).unwrap())
                    .collect::<Vec<_>>();
                assert_eq!(composed, op_by_op, "{sequence:?}");
            }
        }
    }
}
