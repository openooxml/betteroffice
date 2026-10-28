use crate::types::{ImageRunPosition, LayoutBlock, TextBoxBlock};

pub(crate) fn chart_floats(
    position: Option<&ImageRunPosition>,
    wrap_type: Option<&str>,
    behind_doc: Option<bool>,
) -> bool {
    position.is_some()
        && wrap_type == Some("none")
        && behind_doc != Some(true)
        && position.and_then(|position| position.behind_doc) != Some(true)
}

pub(crate) fn is_floating_text_box_block(block: &TextBoxBlock) -> bool {
    block.display_mode.as_deref() == Some("float")
        || matches!(
            block.wrap_type.as_deref(),
            Some("square" | "tight" | "through" | "behind" | "inFront" | "topAndBottom")
        )
}

pub(crate) fn is_floating_drawing(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Shape(shape) => shape.position.is_some(),
        LayoutBlock::TextBox(text_box) => is_floating_text_box_block(text_box),
        LayoutBlock::Table(table) => table.floating.is_some(),
        LayoutBlock::Image(image) => image
            .anchor
            .as_ref()
            .and_then(|anchor| anchor.is_anchored)
            .unwrap_or(false),
        LayoutBlock::Chart(chart) => chart_floats(
            chart.position.as_ref(),
            chart.wrap_type.as_deref(),
            chart.behind_doc,
        ),
        _ => false,
    }
}

fn chart_anchor_owner<'a>(
    index: usize,
    end: usize,
    block: &impl Fn(usize) -> &'a LayoutBlock,
) -> Option<(usize, usize)> {
    let LayoutBlock::Paragraph(paragraph) = block(end) else {
        return None;
    };
    let (Some(mut start), Some(mut stop)) = (paragraph.pm_start, paragraph.pm_end) else {
        return None;
    };
    if let Some((from, to)) = paragraph
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.anchor_paragraph_offsets)
    {
        stop = start + to;
        start += from;
    }
    let mut first = None;
    let mut has_chart = false;
    for index in index..end {
        let block = block(index);
        if block
            .pm_start()
            .is_some_and(|position| position > start && position < stop)
        {
            first.get_or_insert(index);
            has_chart |= matches!(block, LayoutBlock::Chart(_));
        }
    }
    has_chart.then_some((first?, end))
}

#[derive(Default)]
pub(crate) struct FloatingGroupOwner {
    end: usize,
    owner: Option<(usize, usize)>,
}

impl FloatingGroupOwner {
    pub(crate) fn at<'a>(
        &mut self,
        len: usize,
        index: usize,
        block: impl Fn(usize) -> &'a LayoutBlock,
    ) -> Option<usize> {
        if index >= self.end {
            self.end = index + 1;
            self.owner = None;
            if is_floating_drawing(block(index)) {
                self.end = (index + 1..len)
                    .find(|&next| !is_floating_drawing(block(next)))
                    .unwrap_or(len);
                if self.end < len {
                    self.owner = chart_anchor_owner(index, self.end, &block);
                }
            }
        }
        self.owner
            .filter(|&(first, _)| index >= first)
            .map(|(_, owner)| owner)
    }
}

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
