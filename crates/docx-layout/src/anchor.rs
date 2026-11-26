use crate::types::ImageRunPosition;

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

pub(crate) fn resolve_position(
    position: Option<&ImageRunPosition>,
    width: f64,
    height: f64,
    frame: &AnchorFrame,
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
        let relative_to = spec
            .and_then(|axis| axis.relative_to.as_deref())
            .unwrap_or(if horizontal { "column" } else { "paragraph" });
        let odd = frame.odd_page;
        let (start, end) = if horizontal {
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
        };
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

pub(crate) fn is_cell_shape_anchor(
    position: Option<&ImageRunPosition>,
    wrap: Option<&str>,
    rotated: bool,
) -> bool {
    if rotated {
        return false;
    }
    match wrap {
        Some("none" | "behind" | "inFront") => {}
        None if position.is_some() => {}
        _ => return false,
    }
    let Some(position) = position else {
        return true;
    };
    !position.use_simple_pos.unwrap_or(false)
        && position.horizontal.as_ref().is_none_or(|axis| {
            matches!(
                axis.relative_to.as_deref(),
                None | Some("column" | "character" | "margin")
            )
        })
        && position.vertical.as_ref().is_none_or(|axis| {
            matches!(
                axis.relative_to.as_deref(),
                None | Some("paragraph" | "line")
            ) && !matches!(axis.align.as_deref(), Some("center" | "bottom"))
        })
}

pub(crate) fn cell_shape_position(
    position: Option<&ImageRunPosition>,
    (x, width): (f64, f64),
    paragraph_top: f64,
    (shape_width, shape_height): (f64, f64),
) -> (f64, f64) {
    let frame = AnchorFrame {
        page_width: x + width,
        page_height: 0.0,
        margin_left: x,
        margin_right: 0.0,
        margin_top: 0.0,
        margin_bottom: 0.0,
        flow_x: x,
        flow_y: paragraph_top,
        flow_width: width,
        flow_height: 0.0,
        odd_page: true,
    };
    resolve_position(position, shape_width, shape_height, &frame)
}
