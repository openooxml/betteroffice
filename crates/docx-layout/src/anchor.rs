use crate::types::ImageRunPosition;

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub(crate) struct AnchorFrame {
    pub page_width: f64,
    pub page_height: f64,
    pub margin_left: f64,
    pub margin_right: f64,
    pub margin_top: f64,
    pub margin_bottom: f64,
    pub flow_x: f64,
    pub flow_y: f64,
    pub flow_width: f64,
    pub flow_height: f64,
    pub odd_page: bool,
}

pub(crate) struct CellAnchorFrame {
    pub x: f64,
    pub width: f64,
    pub paragraph_top: f64,
    pub paragraph_height: f64,
    pub character_x: Option<f64>,
}

#[derive(Clone, Copy, Debug, Default)]
pub(crate) struct CellAnchorMetrics {
    pub paragraph_height: f64,
    pub character_x: Option<f64>,
}

pub(crate) fn anchor_needs_cell_text(position: &ImageRunPosition) -> bool {
    relative_to(position.horizontal.as_ref()) == Some("character")
        || position.vertical.as_ref().is_some_and(|axis| {
            matches!(axis.relative_to.as_deref(), None | Some("paragraph" | "line"))
                && matches!(axis.align.as_deref(), Some("center" | "bottom"))
                && axis.pos_offset.is_none()
        })
}

fn relative_to(axis: Option<&crate::types::AxisPosition>) -> Option<&str> {
    axis.and_then(|axis| axis.relative_to.as_deref())
}

pub(crate) fn cell_axes(position: &ImageRunPosition) -> (bool, bool) {
    if position.use_simple_pos == Some(true)
        && position
            .simple_pos
            .as_ref()
            .is_some_and(serde_json::Value::is_object)
    {
        return (false, false);
    }
    (
        matches!(relative_to(position.horizontal.as_ref()), None | Some("column" | "character" | "margin")),
        matches!(relative_to(position.vertical.as_ref()), None | Some("paragraph" | "line")),
    )
}

pub(crate) fn cell_anchor_uses_page(position: &ImageRunPosition) -> bool {
    cell_axes(position) != (true, true)
}

pub(crate) fn resolve_cell_shape_rect(
    position: &ImageRunPosition,
    size: (f64, f64),
    cell: &CellAnchorFrame,
    page: Option<&AnchorFrame>,
    layout_in_cell: bool,
) -> Option<(f64, f64, f64, f64)> {
    let (cell_x, cell_y) = if layout_in_cell {
        cell_axes(position)
    } else {
        (false, false)
    };
    if page.is_none() && (!cell_x || !cell_y) {
        return None;
    }
    let local = AnchorFrame {
        odd_page: true,
        ..Default::default()
    };
    let (w, h) = size;
    let (x, y) = resolve_position_in_frame(
        Some(position),
        w,
        h,
        page.unwrap_or(&local),
        cell_x.then(|| {
            let start = if relative_to(position.horizontal.as_ref()) == Some("character") {
                cell.character_x.unwrap_or(cell.x)
            } else {
                cell.x
            };
            (start, cell.x + cell.width)
        }),
        cell_y.then_some((cell.paragraph_top, cell.paragraph_top + cell.paragraph_height)),
    );
    Some((x, y, w, h))
}

pub(crate) fn resolve_position(
    position: Option<&ImageRunPosition>,
    width: f64,
    height: f64,
    frame: &AnchorFrame,
) -> (f64, f64) {
    resolve_position_in_frame(position, width, height, frame, None, None)
}

fn resolve_position_in_frame(
    position: Option<&ImageRunPosition>,
    width: f64,
    height: f64,
    frame: &AnchorFrame,
    horizontal_band: Option<(f64, f64)>,
    vertical_band: Option<(f64, f64)>,
) -> (f64, f64) {
    if let Some(position) = position
        && position.use_simple_pos.unwrap_or(false)
        && let Some(simple) = position
            .simple_pos
            .as_ref()
            .and_then(|value| value.as_object())
    {
        let x = simple
            .get("x")
            .and_then(|value| value.as_f64())
            .filter(|value| value.is_finite())
            .unwrap_or(frame.flow_x);
        let y = simple
            .get("y")
            .and_then(|value| value.as_f64())
            .filter(|value| value.is_finite())
            .unwrap_or(frame.flow_y);
        return (x, y);
    }

    let coordinate = |spec: Option<&crate::types::AxisPosition>, horizontal: bool| {
        let relative_to = relative_to(spec).unwrap_or(if horizontal { "column" } else { "paragraph" });
        let odd = frame.odd_page;
        let band = if horizontal { horizontal_band } else { vertical_band };
        let (start, end) = band.unwrap_or_else(|| {
            if horizontal {
                match relative_to {
                    "page" => (0.0, frame.page_width),
                    "margin" => (frame.margin_left, frame.page_width - frame.margin_right),
                    "leftMargin" => (0.0, frame.margin_left),
                    "rightMargin" => (frame.page_width - frame.margin_right, frame.page_width),
                    "insideMargin" if odd => (0.0, frame.margin_left),
                    "insideMargin" => (frame.page_width - frame.margin_right, frame.page_width),
                    "outsideMargin" if odd => (frame.page_width - frame.margin_right, frame.page_width),
                    "outsideMargin" => (0.0, frame.margin_left),
                    _ => (frame.flow_x, frame.flow_x + frame.flow_width),
                }
            } else {
                match relative_to {
                    "page" => (0.0, frame.page_height),
                    "margin" => (frame.margin_top, frame.page_height - frame.margin_bottom),
                    "topMargin" => (0.0, frame.margin_top),
                    "bottomMargin" => (frame.page_height - frame.margin_bottom, frame.page_height),
                    _ => (frame.flow_y, frame.flow_y + frame.flow_height),
                }
            }
        });
        let extent = if horizontal { width } else { height };
        if let Some(offset) = spec
            .and_then(|axis| axis.pos_offset)
            .filter(|value| value.is_finite())
        {
            return start + offset;
        }
        match spec.and_then(|axis| axis.align.as_deref()) {
            Some("center") => start + (end - start - extent) / 2.0,
            Some("right" | "bottom" | "outside") => end - extent,
            Some("inside") if !odd => end - extent,
            _ => start,
        }
    };
    (
        coordinate(position.and_then(|value| value.horizontal.as_ref()), true),
        coordinate(position.and_then(|value| value.vertical.as_ref()), false),
    )
}
