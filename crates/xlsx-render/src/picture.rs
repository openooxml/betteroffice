//! Pictures a sheet's drawings anchor: placed and clipped like charts, and
//! stacked among the charts of their drawing in anchor order.

use xlsx_model::{PictureCrop, SheetChart, SheetPicture};

use crate::Viewport;
use crate::chart::{plot_rect_to_rect, resolve_chart_anchor, viewport_geometry};
use crate::display_list::{DrawCmd, ImageCrop, Rect};
use crate::geometry::GridGeometry;

/// The pictures one frame shows, each painted once, in anchor order.
pub(crate) struct PictureLayer<'a> {
    pending: Vec<Option<(&'a SheetPicture, DrawCmd)>>,
}

impl<'a> PictureLayer<'a> {
    pub(crate) fn new(
        pictures: &'a [SheetPicture],
        geometry: &GridGeometry,
        viewport: &Viewport,
        frozen_rows: u32,
        frozen_cols: u32,
    ) -> Self {
        let pending = pictures
            .iter()
            .filter_map(|picture| {
                picture_command(picture, geometry, viewport, frozen_rows, frozen_cols)
                    .map(|command| Some((picture, command)))
            })
            .collect();
        Self { pending }
    }

    /// Paints the pictures `chart`'s drawing stacks beneath it.
    pub(crate) fn paint_below(&mut self, chart: &SheetChart, commands: &mut Vec<DrawCmd>) {
        for slot in &mut self.pending {
            if slot.as_ref().is_some_and(|(picture, _)| {
                picture.drawing == chart.drawing && picture.anchor_index < chart.anchor_index
            }) && let Some((_, command)) = slot.take()
            {
                commands.push(command);
            }
        }
    }

    /// Paints every picture no chart covered.
    pub(crate) fn paint_rest(self, commands: &mut Vec<DrawCmd>) {
        commands.extend(
            self.pending
                .into_iter()
                .flatten()
                .map(|(_, command)| command),
        );
    }
}

/// `None` when the anchor does not resolve, the picture is outside the
/// viewport, or its crop leaves nothing.
fn picture_command(
    picture: &SheetPicture,
    geometry: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
) -> Option<DrawCmd> {
    let resolved = resolve_chart_anchor(picture.anchor, geometry, frozen_rows, frozen_cols).ok()?;
    let (frame, clip) = viewport_geometry(resolved, geometry, viewport, frozen_rows, frozen_cols)?;
    let (crop, rect) = placement(picture.crop, plot_rect_to_rect(frame)?)?;
    Some(DrawCmd::Image {
        x: rect.x,
        y: rect.y,
        w: rect.w,
        h: rect.h,
        asset_id: picture.media.as_str().into(),
        crop,
        clip: Some(clip),
    })
}

/// The source fractions to discard and the rect the rest fills: a negative
/// inset shrinks the image inside its frame rather than cropping it.
fn placement(crop: PictureCrop, frame: Rect) -> Option<(ImageCrop, Rect)> {
    let fraction = |inset: i64| inset as f32 / 100_000.0;
    let (left, top) = (fraction(crop.left), fraction(crop.top));
    let (right, bottom) = (fraction(crop.right), fraction(crop.bottom));
    let (kept_x, kept_y) = (1.0 - left - right, 1.0 - top - bottom);
    let discarded = ImageCrop {
        left: left.max(0.0),
        top: top.max(0.0),
        right: right.max(0.0),
        bottom: bottom.max(0.0),
    };
    let (shown_x, shown_y) = (
        1.0 - discarded.left - discarded.right,
        1.0 - discarded.top - discarded.bottom,
    );
    if !(kept_x > 0.0 && kept_y > 0.0 && shown_x > 0.0 && shown_y > 0.0) {
        return None;
    }
    let rect = Rect {
        x: frame.x + (-left).max(0.0) * frame.w / kept_x,
        y: frame.y + (-top).max(0.0) * frame.h / kept_y,
        w: frame.w * shown_x / kept_x,
        h: frame.h * shown_y / kept_y,
    };
    [rect.x, rect.y, rect.w, rect.h]
        .iter()
        .all(|value| value.is_finite())
        .then_some((discarded, rect))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::col_chars_to_px;
    use crate::{RenderError, build_display_list_with_geometry};
    use ooxml_drawingml::chart::ChartSpace;
    use xlsx_model::{
        AnchorCell, AnchorEditAs, AnchorExtent, ChartAnchor, Sheet, SheetId, Workbook,
    };

    const DRAWING: &str = "xl/drawings/drawing1.xml";

    fn cell(col: u32, col_off: i64, row: u32, row_off: i64) -> AnchorCell {
        AnchorCell {
            col,
            col_off,
            row,
            row_off,
        }
    }

    fn picture(media: &str, anchor_index: usize, anchor: ChartAnchor) -> SheetPicture {
        SheetPicture {
            media: media.into(),
            drawing: DRAWING.into(),
            anchor_index,
            anchor,
            crop: PictureCrop::default(),
        }
    }

    fn one_cell(col: u32, row: u32) -> ChartAnchor {
        ChartAnchor::OneCell {
            from: cell(col, 0, row, 0),
            extent: AnchorExtent {
                cx: 952_500,
                cy: 476_250,
            },
        }
    }

    fn frame(sheet: Sheet, pictures: &[SheetPicture], viewport: Viewport) -> Vec<DrawCmd> {
        let mut wb = Workbook::default();
        let geometry = GridGeometry::new(&sheet, &wb.styles);
        wb.sheets.push(sheet);
        build_display_list_with_geometry(
            &wb,
            SheetId(0),
            &viewport,
            &[],
            &geometry,
            None,
            pictures,
            |chart| {
                Err::<ChartSpace, _>(RenderError::ChartSourceUnavailable {
                    part: chart.part.clone(),
                })
            },
        )
        .unwrap()
        .commands
    }

    fn view(x: f32, y: f32) -> Viewport {
        Viewport {
            x,
            y,
            width: 800.0,
            height: 600.0,
        }
    }

    fn image_rects(commands: &[DrawCmd]) -> Vec<(&str, Rect)> {
        commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Image {
                    x,
                    y,
                    w,
                    h,
                    asset_id,
                    ..
                } => Some((
                    &**asset_id,
                    Rect {
                        x: *x,
                        y: *y,
                        w: *w,
                        h: *h,
                    },
                )),
                _ => None,
            })
            .collect()
    }

    fn sized_sheet() -> Sheet {
        let mut sheet = Sheet::new("Sheet1");
        sheet.col_widths = [(0, 10.0), (1, 20.0), (2, 5.0)].into();
        sheet.row_heights = [(0, 30.0), (1, 45.0), (2, 15.0)].into();
        sheet
    }

    #[test]
    fn a_two_cell_anchor_spans_the_column_widths_and_row_heights_it_covers() {
        let anchor = ChartAnchor::TwoCell {
            from: cell(1, 9_525, 1, 0),
            to: cell(3, 0, 3, 19_050),
            edit_as: AnchorEditAs::TwoCell,
        };
        let commands = frame(
            sized_sheet(),
            &[picture("xl/media/image1.png", 0, anchor)],
            view(0.0, 0.0),
        );
        let x = col_chars_to_px(10.0) + 1.0;
        let right = col_chars_to_px(10.0) + col_chars_to_px(20.0) + col_chars_to_px(5.0);
        assert_eq!(
            image_rects(&commands),
            [(
                "xl/media/image1.png",
                Rect {
                    x,
                    y: 40.0,
                    w: right - x,
                    h: 82.0,
                }
            )]
        );
    }

    #[test]
    fn a_one_cell_anchor_takes_its_size_from_ext() {
        let commands = frame(
            sized_sheet(),
            &[picture("xl/media/image1.png", 0, one_cell(1, 1))],
            view(0.0, 0.0),
        );
        assert_eq!(
            image_rects(&commands),
            [(
                "xl/media/image1.png",
                Rect {
                    x: col_chars_to_px(10.0),
                    y: 40.0,
                    w: 100.0,
                    h: 50.0,
                }
            )]
        );
    }

    #[test]
    fn a_picture_scrolled_out_of_the_viewport_draws_nothing() {
        let pictures = [picture("xl/media/image1.png", 0, one_cell(1, 1))];
        assert!(image_rects(&frame(sized_sheet(), &pictures, view(0.0, 2_000.0))).is_empty());
        assert!(image_rects(&frame(sized_sheet(), &pictures, view(2_000.0, 0.0))).is_empty());
    }

    #[test]
    fn pictures_stack_with_the_charts_of_their_drawing_in_anchor_order() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.charts.push(SheetChart {
            raw_xml: None,
            part: "xl/charts/chart1.xml".into(),
            drawing: DRAWING.into(),
            anchor_index: 1,
            anchor: one_cell(0, 0),
            refs: Vec::new(),
        });
        let commands = frame(
            sheet,
            &[
                picture("xl/media/above.png", 2, one_cell(0, 0)),
                picture("xl/media/below.png", 0, one_cell(0, 0)),
            ],
            view(0.0, 0.0),
        );
        let position = |media: &str| {
            commands.iter().position(
                |command| matches!(command, DrawCmd::Image { asset_id, .. } if &**asset_id == media),
            )
        };
        let chart = commands
            .iter()
            .position(|command| matches!(command, DrawCmd::FillRect { clip: Some(_), .. }))
            .unwrap();
        assert!(position("xl/media/below.png").unwrap() < chart);
        assert!(chart < position("xl/media/above.png").unwrap());
    }

    const FRAME: Rect = Rect {
        x: 10.0,
        y: 20.0,
        w: 100.0,
        h: 50.0,
    };

    #[test]
    fn a_positive_inset_crops_and_a_negative_one_pads() {
        let (crop, rect) = placement(
            PictureCrop {
                left: 25_000,
                right: -25_000,
                ..PictureCrop::default()
            },
            FRAME,
        )
        .unwrap();
        assert_eq!(
            crop,
            ImageCrop {
                left: 0.25,
                ..ImageCrop::default()
            }
        );
        assert_eq!(rect, Rect { w: 75.0, ..FRAME });
        let (crop, rect) = placement(PictureCrop::default(), FRAME).unwrap();
        assert!(crop.is_whole());
        assert_eq!(rect, FRAME);
    }

    #[test]
    fn a_crop_that_leaves_nothing_draws_nothing() {
        for crop in [
            PictureCrop {
                left: 60_000,
                right: 40_000,
                ..PictureCrop::default()
            },
            PictureCrop {
                top: 150_000,
                bottom: -60_000,
                ..PictureCrop::default()
            },
        ] {
            assert_eq!(placement(crop, FRAME), None);
        }
    }
}
