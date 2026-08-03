//! display-list hit testing: which object sits under a viewport-local point.
//! a pure function of the regions a frame publishes, so chrome never re-derives
//! geometry the renderer already owns.

use crate::display_list::ChartRegion;

/// The topmost chart whose visible region contains the point. Regions are in
/// paint order, so the last match wins; the clipped region is used, and a chart
/// scrolled under a frozen pane is therefore not hit where it is hidden.
///
/// A chart that degraded to a placeholder still answers: failing to draw it
/// does not stop it being an object on the sheet that occupies that space, and
/// moving it out of the way is exactly what a reader is likely to want.
pub fn chart_at_point(charts: &[ChartRegion], x: f32, y: f32) -> Option<&ChartRegion> {
    if !x.is_finite() || !y.is_finite() {
        return None;
    }
    charts.iter().rev().find(|chart| {
        let clip = chart.clip;
        x >= clip.x && x < clip.x + clip.w && y >= clip.y && y < clip.y + clip.h
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::display_list::Rect;

    fn region(id: &str, rect: Rect, clip: Rect) -> ChartRegion {
        ChartRegion {
            id: id.into(),
            label: String::new(),
            placeholder: false,
            rect,
            clip,
            movable: true,
        }
    }

    fn square(x: f32, y: f32, w: f32, h: f32) -> Rect {
        Rect { x, y, w, h }
    }

    #[test]
    fn point_inside_a_region_names_its_chart() {
        let charts = [region(
            "a",
            square(10.0, 20.0, 100.0, 50.0),
            square(10.0, 20.0, 100.0, 50.0),
        )];
        assert_eq!(chart_at_point(&charts, 10.0, 20.0).unwrap().id, "a");
        assert_eq!(chart_at_point(&charts, 109.9, 69.9).unwrap().id, "a");
        assert!(chart_at_point(&charts, 110.0, 40.0).is_none());
        assert!(chart_at_point(&charts, 9.9, 40.0).is_none());
        assert!(chart_at_point(&charts, f32::NAN, 40.0).is_none());
    }

    #[test]
    fn overlapping_regions_resolve_to_the_last_painted() {
        let charts = [
            region(
                "under",
                square(0.0, 0.0, 100.0, 100.0),
                square(0.0, 0.0, 100.0, 100.0),
            ),
            region(
                "over",
                square(50.0, 50.0, 100.0, 100.0),
                square(50.0, 50.0, 100.0, 100.0),
            ),
        ];
        assert_eq!(chart_at_point(&charts, 60.0, 60.0).unwrap().id, "over");
        assert_eq!(chart_at_point(&charts, 10.0, 10.0).unwrap().id, "under");
    }

    /// A chart the renderer could not draw is painted as a placeholder but is
    /// still an object on the sheet, so it stays selectable and movable.
    #[test]
    fn a_placeholdered_chart_is_still_addressable() {
        let box_ = square(0.0, 0.0, 50.0, 40.0);
        let charts = [ChartRegion {
            placeholder: true,
            ..region("undrawable", box_, box_)
        }];
        let hit = chart_at_point(&charts, 25.0, 20.0).expect("a placeholder still answers");
        assert_eq!(hit.id, "undrawable");
        assert!(hit.movable);
    }

    #[test]
    fn the_clipped_region_bounds_the_hit_not_the_full_rect() {
        let charts = [region(
            "clipped",
            square(0.0, 0.0, 200.0, 100.0),
            square(80.0, 0.0, 120.0, 100.0),
        )];
        assert!(chart_at_point(&charts, 40.0, 50.0).is_none());
        assert_eq!(chart_at_point(&charts, 100.0, 50.0).unwrap().id, "clipped");
    }
}
