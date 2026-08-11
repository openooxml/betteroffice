//! charts anchored to a worksheet: where the drawing pins them to the grid and
//! which ranges their `c:f` references name.

use serde::{Deserialize, Serialize};

use crate::addr::{ColId, RowId};

/// which slot a `c:f` reference came from. informational — every slot remaps
/// the same way.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ChartRefKind {
    /// `c:ser/c:tx`, the series name.
    SeriesName,
    /// `c:ser/c:cat` or `c:ser/c:xVal`.
    Categories,
    /// `c:ser/c:val` or `c:ser/c:yVal`.
    Values,
    /// `c:ser/c:bubbleSize`.
    BubbleSize,
    /// a chart or axis title.
    Title,
    /// a data-label range.
    DataLabels,
    /// any other `c:f` in the part.
    Other,
}

/// one `c:f` in a chart part, addressed by its document order within it.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartRef {
    pub kind: ChartRefKind,
    pub formula: String,
}

/// a grid-anchored corner: a cell plus an EMU offset into it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AnchorCell {
    pub col: ColId,
    pub col_off: i64,
    pub row: RowId,
    pub row_off: i64,
}

/// an EMU size.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct AnchorExtent {
    pub cx: i64,
    pub cy: i64,
}

/// an EMU position measured from the sheet origin, independent of the grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct AnchorPos {
    pub x: i64,
    pub y: i64,
}

/// `xdr:twoCellAnchor/@editAs` — how the shape follows a grid edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum AnchorEditAs {
    /// move and size with cells; the schema default.
    #[default]
    TwoCell,
    /// move with cells, keep the size.
    OneCell,
    /// neither move nor size.
    Absolute,
}

impl AnchorEditAs {
    pub fn from_sml(value: &str) -> Option<Self> {
        match value {
            "twoCell" => Some(Self::TwoCell),
            "oneCell" => Some(Self::OneCell),
            "absolute" => Some(Self::Absolute),
            _ => None,
        }
    }

    pub fn as_sml(&self) -> &'static str {
        match self {
            Self::TwoCell => "twoCell",
            Self::OneCell => "oneCell",
            Self::Absolute => "absolute",
        }
    }
}

/// how a drawing pins a chart to the sheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum ChartAnchor {
    /// both corners ride the grid; `edit_as` decides whether the second one
    /// actually follows an edit.
    TwoCell {
        from: AnchorCell,
        to: AnchorCell,
        edit_as: AnchorEditAs,
    },
    /// the top-left corner rides the grid, the size is fixed.
    OneCell {
        from: AnchorCell,
        extent: AnchorExtent,
    },
    /// pinned to the sheet, not the grid.
    Absolute {
        pos: AnchorPos,
        extent: AnchorExtent,
    },
}

impl ChartAnchor {
    /// the grid-anchored top-left corner, absent for an absolute anchor.
    pub fn from_cell(&self) -> Option<AnchorCell> {
        match self {
            Self::TwoCell { from, .. } | Self::OneCell { from, .. } => Some(*from),
            Self::Absolute { .. } => None,
        }
    }

    /// how this anchor reacts to a grid edit.
    pub fn edit_as(&self) -> AnchorEditAs {
        match self {
            Self::TwoCell { edit_as, .. } => *edit_as,
            Self::OneCell { .. } => AnchorEditAs::OneCell,
            Self::Absolute { .. } => AnchorEditAs::Absolute,
        }
    }
}

/// a chart part hanging off a worksheet through a drawing.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetChart {
    /// package path of the `c:chartSpace` part.
    pub part: String,
    /// package path of the `xdr:wsDr` part that anchors it.
    pub drawing: String,
    /// document-order index of the anchor within that drawing part.
    pub anchor_index: usize,
    pub anchor: ChartAnchor,
    /// every `c:f` in the chart part, in document order.
    pub refs: Vec<ChartRef>,
}

impl SheetChart {
    /// What addresses this frame: the drawing part and the anchor in it. The
    /// chart part cannot, because two anchors may share one.
    pub fn frame_id(&self) -> String {
        format!("{}#{}", self.drawing, self.anchor_index)
    }

    /// Whether both name the same drawing anchor.
    pub fn is_same_frame(&self, other: &Self) -> bool {
        self.drawing == other.drawing && self.anchor_index == other.anchor_index
    }

    /// Whether this is the frame an op recorded when it saw `part` at `from`.
    /// A frame id is an ordinal, and a drawing that gained, lost or reordered
    /// an anchor renumbers it, so the id alone cannot say.
    pub fn is_recorded_frame(&self, part: &str, from: ChartAnchor) -> bool {
        self.part == part && self.anchor == from
    }
}
