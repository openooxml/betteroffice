//! ECMA-376 preset guide and path formulas.

use super::{Frame, Outline, PresetGeometryPath, PresetPathFill, ooxml_angle};
use std::collections::HashMap;

pub(super) fn paths(
    kind: &str,
    adjustments: &HashMap<String, f64>,
    aspect: f64,
) -> Option<Vec<PresetGeometryPath>> {
    let frame = Frame::new(aspect);
    let frame = Frame {
        w: frame.w * 100000.0,
        h: frame.h * 100000.0,
    };
    let layers = match kind {
        "arc" => arc(frame, adjustments),
        "cube" => cube(frame, adjustments),
        "rightBrace" => right_brace(frame, adjustments),
        "leftBrace" => left_brace(frame, adjustments),
        "wedgeRectCallout" => wedge_rect_callout(frame, adjustments),
        "ribbon2" => ribbon2(frame, adjustments),
        "swooshArrow" => swoosh_arrow(frame, adjustments),
        "circularArrow" => circular_arrow(frame, adjustments),
        "foldedCorner" => folded_corner(frame, adjustments),
        "cloudCallout" => cloud_callout(frame, adjustments),
        _ => return None,
    };
    let mut paths: Vec<PresetGeometryPath> = Vec::new();
    for layer in layers {
        if let Some(previous) = paths.last_mut()
            && previous.fill == layer.fill
            && previous.stroke == layer.stroke
        {
            previous.commands.extend(layer.commands);
        } else {
            paths.push(layer);
        }
    }
    Some(paths)
}

pub(super) fn defaults(kind: &str) -> &'static [(&'static str, f64)] {
    match kind {
        "arc" => &[("adj1", 162.0), ("adj2", 0.0)],
        "cube" => &[("adj", 0.25)],
        "rightBrace" => &[("adj1", 0.08333), ("adj2", 0.5)],
        "leftBrace" => &[("adj1", 0.08333), ("adj2", 0.5)],
        "wedgeRectCallout" => &[("adj1", -0.20833), ("adj2", 0.625)],
        "ribbon2" => &[("adj1", 0.16667), ("adj2", 0.5)],
        "swooshArrow" => &[("adj1", 0.25), ("adj2", 0.16667)],
        "circularArrow" => &[
            ("adj1", 0.125),
            ("adj2", 11.42319),
            ("adj3", 204.57681),
            ("adj4", 108.0),
            ("adj5", 0.125),
        ],
        "foldedCorner" => &[("adj", 0.16667)],
        "cloudCallout" => &[("adj1", -0.20833), ("adj2", 0.625)],
        _ => &[],
    }
}

fn adjustment(values: &HashMap<String, f64>, name: &str, default: f64) -> f64 {
    values
        .get(name)
        .copied()
        .filter(|value| value.is_finite())
        .unwrap_or(default)
        .clamp(-1e12, 1e12)
        * 100000.0
}

fn divide(numerator: f64, denominator: f64) -> f64 {
    if denominator == 0.0 {
        0.0
    } else {
        numerator / denominator
    }
}

fn arc(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 162.0);
    let adj2 = adjustment(adjustments, "adj2", 0.0);
    let st_ang: f64 = f64::clamp(adj1, 0.0, f64::max(21599999.0, 0.0));
    let en_ang: f64 = f64::clamp(adj2, 0.0, f64::max(21599999.0, 0.0));
    let sw11: f64 = en_ang + 0.0 - st_ang;
    let sw12: f64 = sw11 + 21600000.0 - 0.0;
    let sw_ang: f64 = if sw11 > 0.0 { sw11 } else { sw12 };
    let wt1: f64 = (w / 2.0) * ooxml_angle(st_ang).sin();
    let ht1: f64 = (h / 2.0) * ooxml_angle(st_ang).cos();
    let dx1: f64 = (w / 2.0) * f64::atan2(wt1, ht1).cos();
    let dy1: f64 = (h / 2.0) * f64::atan2(wt1, ht1).sin();
    let x1: f64 = (w / 2.0) + dx1 - 0.0;
    let y1: f64 = (h / 2.0) + dy1 - 0.0;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(x1, y1)
                .arc_to(w / 2.0, h / 2.0, ooxml_angle(st_ang), ooxml_angle(sw_ang))
                .line_to(w / 2.0, h / 2.0)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(x1, y1)
                .arc_to(w / 2.0, h / 2.0, ooxml_angle(st_ang), ooxml_angle(sw_ang))
                .finish(),
        },
    ]
}

fn cube(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj = adjustment(adjustments, "adj", 0.25);
    let a: f64 = f64::clamp(adj, 0.0, f64::max(100000.0, 0.0));
    let y1: f64 = divide(100000.0 * a, 100000.0);
    let y4: f64 = h + 0.0 - y1;
    let x4: f64 = w + 0.0 - y1;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, y1)
                .line_to(x4, y1)
                .line_to(x4, h)
                .line_to(0.0, h)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(x4, y1)
                .line_to(w, 0.0)
                .line_to(w, y4)
                .line_to(x4, h)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::LightenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, y1)
                .line_to(y1, 0.0)
                .line_to(w, 0.0)
                .line_to(x4, y1)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(0.0, y1)
                .line_to(y1, 0.0)
                .line_to(w, 0.0)
                .line_to(w, y4)
                .line_to(x4, h)
                .line_to(0.0, h)
                .close()
                .move_to(0.0, y1)
                .line_to(x4, y1)
                .line_to(w, 0.0)
                .move_to(x4, y1)
                .line_to(x4, h)
                .finish(),
        },
    ]
}

fn right_brace(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.08333);
    let adj2 = adjustment(adjustments, "adj2", 0.5);
    let a2: f64 = f64::clamp(adj2, 0.0, f64::max(100000.0, 0.0));
    let q1: f64 = 100000.0 + 0.0 - a2;
    let q2: f64 = f64::min(q1, a2);
    let q3: f64 = divide(q2 * 1.0, 2.0);
    let max_adj1: f64 = divide(q3 * h, 100000.0);
    let a1: f64 = f64::clamp(adj1, 0.0, f64::max(max_adj1, 0.0));
    let y1: f64 = divide(100000.0 * a1, 100000.0);
    let y3: f64 = divide(h * a2, 100000.0);
    let y2: f64 = y3 + 0.0 - y1;
    let y4: f64 = h + 0.0 - y1;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, 0.0)
                .arc_to(w / 2.0, y1, ooxml_angle(16200000.0), ooxml_angle(5400000.0))
                .line_to(w / 2.0, y2)
                .arc_to(
                    w / 2.0,
                    y1,
                    ooxml_angle(10800000.0),
                    ooxml_angle(-5400000.0),
                )
                .arc_to(
                    w / 2.0,
                    y1,
                    ooxml_angle(16200000.0),
                    ooxml_angle(-5400000.0),
                )
                .line_to(w / 2.0, y4)
                .arc_to(w / 2.0, y1, ooxml_angle(0.0), ooxml_angle(5400000.0))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(0.0, 0.0)
                .arc_to(w / 2.0, y1, ooxml_angle(16200000.0), ooxml_angle(5400000.0))
                .line_to(w / 2.0, y2)
                .arc_to(
                    w / 2.0,
                    y1,
                    ooxml_angle(10800000.0),
                    ooxml_angle(-5400000.0),
                )
                .arc_to(
                    w / 2.0,
                    y1,
                    ooxml_angle(16200000.0),
                    ooxml_angle(-5400000.0),
                )
                .line_to(w / 2.0, y4)
                .arc_to(w / 2.0, y1, ooxml_angle(0.0), ooxml_angle(5400000.0))
                .finish(),
        },
    ]
}

fn left_brace(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.08333);
    let adj2 = adjustment(adjustments, "adj2", 0.5);
    let a2: f64 = f64::clamp(adj2, 0.0, f64::max(100000.0, 0.0));
    let q1: f64 = 100000.0 + 0.0 - a2;
    let q2: f64 = f64::min(q1, a2);
    let q3: f64 = divide(q2 * 1.0, 2.0);
    let max_adj1: f64 = divide(q3 * h, 100000.0);
    let a1: f64 = f64::clamp(adj1, 0.0, f64::max(max_adj1, 0.0));
    let y1: f64 = divide(100000.0 * a1, 100000.0);
    let y3: f64 = divide(h * a2, 100000.0);
    let y4: f64 = y3 + y1 - 0.0;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(w, h)
                .arc_to(w / 2.0, y1, ooxml_angle(5400000.0), ooxml_angle(5400000.0))
                .line_to(w / 2.0, y4)
                .arc_to(w / 2.0, y1, ooxml_angle(0.0), ooxml_angle(-5400000.0))
                .arc_to(w / 2.0, y1, ooxml_angle(5400000.0), ooxml_angle(-5400000.0))
                .line_to(w / 2.0, y1)
                .arc_to(w / 2.0, y1, ooxml_angle(10800000.0), ooxml_angle(5400000.0))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(w, h)
                .arc_to(w / 2.0, y1, ooxml_angle(5400000.0), ooxml_angle(5400000.0))
                .line_to(w / 2.0, y4)
                .arc_to(w / 2.0, y1, ooxml_angle(0.0), ooxml_angle(-5400000.0))
                .arc_to(w / 2.0, y1, ooxml_angle(5400000.0), ooxml_angle(-5400000.0))
                .line_to(w / 2.0, y1)
                .arc_to(w / 2.0, y1, ooxml_angle(10800000.0), ooxml_angle(5400000.0))
                .finish(),
        },
    ]
}

fn wedge_rect_callout(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", -0.20833);
    let adj2 = adjustment(adjustments, "adj2", 0.625);
    let dx_pos: f64 = divide(w * adj1, 100000.0);
    let dy_pos: f64 = divide(h * adj2, 100000.0);
    let x_pos: f64 = (w / 2.0) + dx_pos - 0.0;
    let y_pos: f64 = (h / 2.0) + dy_pos - 0.0;
    let dq: f64 = divide(dx_pos * h, w);
    let ady: f64 = f64::abs(dy_pos);
    let adq: f64 = f64::abs(dq);
    let dz: f64 = ady + 0.0 - adq;
    let xg1: f64 = if dx_pos > 0.0 { 7.0 } else { 2.0 };
    let xg2: f64 = if dx_pos > 0.0 { 10.0 } else { 5.0 };
    let x1: f64 = divide(w * xg1, 12.0);
    let x2: f64 = divide(w * xg2, 12.0);
    let yg1: f64 = if dy_pos > 0.0 { 7.0 } else { 2.0 };
    let yg2: f64 = if dy_pos > 0.0 { 10.0 } else { 5.0 };
    let y1: f64 = divide(h * yg1, 12.0);
    let y2: f64 = divide(h * yg2, 12.0);
    let t1: f64 = if dx_pos > 0.0 { 0.0 } else { x_pos };
    let xl: f64 = if dz > 0.0 { 0.0 } else { t1 };
    let t2: f64 = if dy_pos > 0.0 { x1 } else { x_pos };
    let xt: f64 = if dz > 0.0 { t2 } else { x1 };
    let t3: f64 = if dx_pos > 0.0 { x_pos } else { w };
    let xr: f64 = if dz > 0.0 { w } else { t3 };
    let t4: f64 = if dy_pos > 0.0 { x_pos } else { x1 };
    let xb: f64 = if dz > 0.0 { t4 } else { x1 };
    let t5: f64 = if dx_pos > 0.0 { y1 } else { y_pos };
    let yl: f64 = if dz > 0.0 { y1 } else { t5 };
    let t6: f64 = if dy_pos > 0.0 { 0.0 } else { y_pos };
    let yt: f64 = if dz > 0.0 { t6 } else { 0.0 };
    let t7: f64 = if dx_pos > 0.0 { y_pos } else { y1 };
    let yr: f64 = if dz > 0.0 { y1 } else { t7 };
    let t8: f64 = if dy_pos > 0.0 { y_pos } else { h };
    let yb: f64 = if dz > 0.0 { t8 } else { h };
    vec![PresetGeometryPath {
        fill: PresetPathFill::Normal,
        stroke: true,
        commands: Outline::new(frame)
            .move_to(0.0, 0.0)
            .line_to(x1, 0.0)
            .line_to(xt, yt)
            .line_to(x2, 0.0)
            .line_to(w, 0.0)
            .line_to(w, y1)
            .line_to(xr, yr)
            .line_to(w, y2)
            .line_to(w, h)
            .line_to(x2, h)
            .line_to(xb, yb)
            .line_to(x1, h)
            .line_to(0.0, h)
            .line_to(0.0, y2)
            .line_to(xl, yl)
            .line_to(0.0, y1)
            .close()
            .finish(),
    }]
}

fn ribbon2(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.16667);
    let adj2 = adjustment(adjustments, "adj2", 0.5);
    let a1: f64 = f64::clamp(adj1, 0.0, f64::max(33333.0, 0.0));
    let a2: f64 = f64::clamp(adj2, 25000.0, f64::max(75000.0, 25000.0));
    let x10: f64 = w + 0.0 - (w / 8.0);
    let dx2: f64 = divide(w * a2, 200000.0);
    let x2: f64 = (w / 2.0) + 0.0 - dx2;
    let x9: f64 = (w / 2.0) + dx2 - 0.0;
    let x3: f64 = x2 + (w / 32.0) - 0.0;
    let x8: f64 = x9 + 0.0 - (w / 32.0);
    let x5: f64 = x2 + (w / 8.0) - 0.0;
    let x6: f64 = x9 + 0.0 - (w / 8.0);
    let x4: f64 = x5 + 0.0 - (w / 32.0);
    let x7: f64 = x6 + (w / 32.0) - 0.0;
    let dy1: f64 = divide(h * a1, 200000.0);
    let y1: f64 = h + 0.0 - dy1;
    let dy2: f64 = divide(h * a1, 100000.0);
    let y2: f64 = h + 0.0 - dy2;
    let y4: f64 = 0.0 + dy2 - 0.0;
    let y3: f64 = divide(y4 + h, 2.0);
    let h_r: f64 = divide(h * a1, 400000.0);
    let y6: f64 = h + 0.0 - h_r;
    let y7: f64 = y1 + 0.0 - h_r;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, h)
                .line_to(x4, h)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(5400000.0),
                    ooxml_angle(-10800000.0),
                )
                .line_to(x3, y1)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(5400000.0),
                    ooxml_angle(10800000.0),
                )
                .line_to(x8, y2)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(16200000.0),
                    ooxml_angle(10800000.0),
                )
                .line_to(x7, y1)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(16200000.0),
                    ooxml_angle(-10800000.0),
                )
                .line_to(w, h)
                .line_to(x10, y3)
                .line_to(w, y4)
                .line_to(x9, y4)
                .line_to(x9, h_r)
                .arc_to(w / 32.0, h_r, ooxml_angle(0.0), ooxml_angle(-5400000.0))
                .line_to(x3, 0.0)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(16200000.0),
                    ooxml_angle(-5400000.0),
                )
                .line_to(x2, y4)
                .line_to(0.0, y4)
                .line_to(w / 8.0, y3)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(x5, y6)
                .arc_to(w / 32.0, h_r, ooxml_angle(0.0), ooxml_angle(-5400000.0))
                .line_to(x3, y1)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(5400000.0),
                    ooxml_angle(10800000.0),
                )
                .line_to(x5, y2)
                .close()
                .move_to(x6, y6)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(10800000.0),
                    ooxml_angle(5400000.0),
                )
                .line_to(x8, y1)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(5400000.0),
                    ooxml_angle(-10800000.0),
                )
                .line_to(x6, y2)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(0.0, h)
                .line_to(w / 8.0, y3)
                .line_to(0.0, y4)
                .line_to(x2, y4)
                .line_to(x2, h_r)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(10800000.0),
                    ooxml_angle(5400000.0),
                )
                .line_to(x8, 0.0)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(16200000.0),
                    ooxml_angle(5400000.0),
                )
                .line_to(x9, y4)
                .line_to(x9, y4)
                .line_to(w, y4)
                .line_to(x10, y3)
                .line_to(w, h)
                .line_to(x7, h)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(5400000.0),
                    ooxml_angle(10800000.0),
                )
                .line_to(x8, y1)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(5400000.0),
                    ooxml_angle(-10800000.0),
                )
                .line_to(x3, y2)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(16200000.0),
                    ooxml_angle(-10800000.0),
                )
                .line_to(x4, y1)
                .arc_to(
                    w / 32.0,
                    h_r,
                    ooxml_angle(16200000.0),
                    ooxml_angle(10800000.0),
                )
                .close()
                .move_to(x5, y2)
                .line_to(x5, y6)
                .move_to(x6, y6)
                .line_to(x6, y2)
                .move_to(x2, y7)
                .line_to(x2, y4)
                .move_to(x9, y4)
                .line_to(x9, y7)
                .finish(),
        },
    ]
}

fn swoosh_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.16667);
    let a1: f64 = f64::clamp(adj1, 1.0, f64::max(75000.0, 1.0));
    let max_adj2: f64 = divide(70000.0 * w, 100000.0);
    let a2: f64 = f64::clamp(adj2, 0.0, f64::max(max_adj2, 0.0));
    let ad1: f64 = divide(h * a1, 100000.0);
    let ad2: f64 = divide(100000.0 * a2, 100000.0);
    let x_b: f64 = w + 0.0 - ad2;
    let y_b: f64 = 0.0 + (100000.0 / 8.0) - 0.0;
    let alfa: f64 = divide(5400000.0 * 1.0, 14.0);
    let dx0: f64 = (100000.0 / 8.0) * ooxml_angle(alfa).tan();
    let x_c: f64 = x_b + 0.0 - dx0;
    let dx1: f64 = ad1 * ooxml_angle(alfa).tan();
    let y_f: f64 = y_b + ad1 - 0.0;
    let x_f: f64 = x_b + dx1 - 0.0;
    let x_e: f64 = x_f + dx0 - 0.0;
    let y_e: f64 = y_f + (100000.0 / 8.0) - 0.0;
    let dy2: f64 = y_e + 0.0 - 0.0;
    let dy22: f64 = divide(dy2 * 1.0, 2.0);
    let dy3: f64 = divide(h * 1.0, 20.0);
    let y_d: f64 = 0.0 + dy22 - dy3;
    let dy4: f64 = divide((h / 6.0) * 1.0, 1.0);
    let y_p1: f64 = (h / 6.0) + dy4 - 0.0;
    let x_p1: f64 = w / 6.0;
    let dy5: f64 = divide((h / 6.0) * 1.0, 2.0);
    let y_p2: f64 = y_f + dy5 - 0.0;
    let x_p2: f64 = w / 4.0;
    vec![PresetGeometryPath {
        fill: PresetPathFill::Normal,
        stroke: true,
        commands: Outline::new(frame)
            .move_to(0.0, h)
            .quad_to(x_p1, y_p1, x_b, y_b)
            .line_to(x_c, 0.0)
            .line_to(w, y_d)
            .line_to(x_e, y_e)
            .line_to(x_f, y_f)
            .quad_to(x_p2, y_p2, 0.0, h)
            .close()
            .finish(),
    }]
}

fn circular_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.125);
    let adj2 = adjustment(adjustments, "adj2", 11.42319);
    let adj3 = adjustment(adjustments, "adj3", 204.57681);
    let adj4 = adjustment(adjustments, "adj4", 108.0);
    let adj5 = adjustment(adjustments, "adj5", 0.125);
    let a5: f64 = f64::clamp(adj5, 0.0, f64::max(25000.0, 0.0));
    let max_adj1: f64 = divide(a5 * 2.0, 1.0);
    let a1: f64 = f64::clamp(adj1, 0.0, f64::max(max_adj1, 0.0));
    let en_ang: f64 = f64::clamp(adj3, 1.0, f64::max(21599999.0, 1.0));
    let st_ang: f64 = f64::clamp(adj4, 0.0, f64::max(21599999.0, 0.0));
    let th: f64 = divide(100000.0 * a1, 100000.0);
    let thh: f64 = divide(100000.0 * a5, 100000.0);
    let th2: f64 = divide(th * 1.0, 2.0);
    let rw1: f64 = (w / 2.0) + th2 - thh;
    let rh1: f64 = (h / 2.0) + th2 - thh;
    let rw2: f64 = rw1 + 0.0 - th;
    let rh2: f64 = rh1 + 0.0 - th;
    let rw3: f64 = rw2 + th2 - 0.0;
    let rh3: f64 = rh2 + th2 - 0.0;
    let wt_h: f64 = rw3 * ooxml_angle(en_ang).sin();
    let ht_h: f64 = rh3 * ooxml_angle(en_ang).cos();
    let dx_h: f64 = rw3 * f64::atan2(wt_h, ht_h).cos();
    let dy_h: f64 = rh3 * f64::atan2(wt_h, ht_h).sin();
    let x_h: f64 = (w / 2.0) + dx_h - 0.0;
    let y_h: f64 = (h / 2.0) + dy_h - 0.0;
    let r_i: f64 = f64::min(rw2, rh2);
    let u1: f64 = divide(dx_h * dx_h, 1.0);
    let u2: f64 = divide(dy_h * dy_h, 1.0);
    let u3: f64 = divide(r_i * r_i, 1.0);
    let u4: f64 = u1 + 0.0 - u3;
    let u5: f64 = u2 + 0.0 - u3;
    let u6: f64 = divide(u4 * u5, u1);
    let u7: f64 = divide(u6 * 1.0, u2);
    let u8: f64 = 1.0 + 0.0 - u7;
    let u9: f64 = f64::sqrt(f64::max(u8, 0.0));
    let u10: f64 = divide(u4 * 1.0, dx_h);
    let u11: f64 = divide(u10 * 1.0, dy_h);
    let u12: f64 = divide(1.0 + u9, u11);
    let u13: f64 = f64::atan2(u12, 1.0).to_degrees() * 60000.0;
    let u14: f64 = u13 + 21600000.0 - 0.0;
    let u15: f64 = if u13 > 0.0 { u13 } else { u14 };
    let u16: f64 = u15 + 0.0 - en_ang;
    let u17: f64 = u16 + 21600000.0 - 0.0;
    let u18: f64 = if u16 > 0.0 { u16 } else { u17 };
    let u19: f64 = u18 + 0.0 - 10800000.0;
    let u20: f64 = u18 + 0.0 - 21600000.0;
    let u21: f64 = if u19 > 0.0 { u20 } else { u18 };
    let max_ang: f64 = f64::abs(u21);
    let a_ang: f64 = f64::clamp(adj2, 0.0, f64::max(max_ang, 0.0));
    let pt_ang: f64 = en_ang + a_ang - 0.0;
    let wt_a: f64 = rw3 * ooxml_angle(pt_ang).sin();
    let ht_a: f64 = rh3 * ooxml_angle(pt_ang).cos();
    let dx_a: f64 = rw3 * f64::atan2(wt_a, ht_a).cos();
    let dy_a: f64 = rh3 * f64::atan2(wt_a, ht_a).sin();
    let x_a: f64 = (w / 2.0) + dx_a - 0.0;
    let y_a: f64 = (h / 2.0) + dy_a - 0.0;
    let wt_e: f64 = rw1 * ooxml_angle(st_ang).sin();
    let ht_e: f64 = rh1 * ooxml_angle(st_ang).cos();
    let dx_e: f64 = rw1 * f64::atan2(wt_e, ht_e).cos();
    let dy_e: f64 = rh1 * f64::atan2(wt_e, ht_e).sin();
    let x_e: f64 = (w / 2.0) + dx_e - 0.0;
    let y_e: f64 = (h / 2.0) + dy_e - 0.0;
    let dx_g: f64 = thh * ooxml_angle(pt_ang).cos();
    let dy_g: f64 = thh * ooxml_angle(pt_ang).sin();
    let x_g: f64 = x_h + dx_g - 0.0;
    let y_g: f64 = y_h + dy_g - 0.0;
    let dx_b: f64 = thh * ooxml_angle(pt_ang).cos();
    let dy_b: f64 = thh * ooxml_angle(pt_ang).sin();
    let x_b: f64 = x_h + 0.0 - dx_b;
    let y_b: f64 = y_h + 0.0 - dy_b;
    let sx1: f64 = x_b + 0.0 - (w / 2.0);
    let sy1: f64 = y_b + 0.0 - (h / 2.0);
    let sx2: f64 = x_g + 0.0 - (w / 2.0);
    let sy2: f64 = y_g + 0.0 - (h / 2.0);
    let r_o: f64 = f64::min(rw1, rh1);
    let x1_o: f64 = divide(sx1 * r_o, rw1);
    let y1_o: f64 = divide(sy1 * r_o, rh1);
    let x2_o: f64 = divide(sx2 * r_o, rw1);
    let y2_o: f64 = divide(sy2 * r_o, rh1);
    let dx_o: f64 = x2_o + 0.0 - x1_o;
    let dy_o: f64 = y2_o + 0.0 - y1_o;
    let d_o: f64 = f64::hypot(dx_o, dy_o).hypot(0.0);
    let q1: f64 = divide(x1_o * y2_o, 1.0);
    let q2: f64 = divide(x2_o * y1_o, 1.0);
    let det_o: f64 = q1 + 0.0 - q2;
    let q3: f64 = divide(r_o * r_o, 1.0);
    let q4: f64 = divide(d_o * d_o, 1.0);
    let q5: f64 = divide(q3 * q4, 1.0);
    let q6: f64 = divide(det_o * det_o, 1.0);
    let q7: f64 = q5 + 0.0 - q6;
    let q8: f64 = f64::max(q7, 0.0);
    let sdel_o: f64 = f64::sqrt(f64::max(q8, 0.0));
    let ndy_o: f64 = -dy_o;
    let sdy_o: f64 = if ndy_o > 0.0 { -1.0 } else { 1.0 };
    let q9: f64 = divide(sdy_o * dx_o, 1.0);
    let q10: f64 = divide(q9 * sdel_o, 1.0);
    let q11: f64 = divide(det_o * dy_o, 1.0);
    let dx_f1: f64 = divide(q11 + q10, q4);
    let q12: f64 = q11 + 0.0 - q10;
    let dx_f2: f64 = divide(q12 * 1.0, q4);
    let ady_o: f64 = f64::abs(dy_o);
    let q13: f64 = divide(ady_o * sdel_o, 1.0);
    let q14: f64 = divide(det_o * dx_o, -1.0);
    let dy_f1: f64 = divide(q14 + q13, q4);
    let q15: f64 = q14 + 0.0 - q13;
    let dy_f2: f64 = divide(q15 * 1.0, q4);
    let q16: f64 = x2_o + 0.0 - dx_f1;
    let q17: f64 = x2_o + 0.0 - dx_f2;
    let q18: f64 = y2_o + 0.0 - dy_f1;
    let q19: f64 = y2_o + 0.0 - dy_f2;
    let q20: f64 = f64::hypot(q16, q18).hypot(0.0);
    let q21: f64 = f64::hypot(q17, q19).hypot(0.0);
    let q22: f64 = q21 + 0.0 - q20;
    let dx_f: f64 = if q22 > 0.0 { dx_f1 } else { dx_f2 };
    let dy_f: f64 = if q22 > 0.0 { dy_f1 } else { dy_f2 };
    let sdx_f: f64 = divide(dx_f * rw1, r_o);
    let sdy_f: f64 = divide(dy_f * rh1, r_o);
    let x_f: f64 = (w / 2.0) + sdx_f - 0.0;
    let y_f: f64 = (h / 2.0) + sdy_f - 0.0;
    let x1_i: f64 = divide(sx1 * r_i, rw2);
    let y1_i: f64 = divide(sy1 * r_i, rh2);
    let x2_i: f64 = divide(sx2 * r_i, rw2);
    let y2_i: f64 = divide(sy2 * r_i, rh2);
    let dx_i: f64 = x2_i + 0.0 - x1_i;
    let dy_i: f64 = y2_i + 0.0 - y1_i;
    let d_i: f64 = f64::hypot(dx_i, dy_i).hypot(0.0);
    let v1: f64 = divide(x1_i * y2_i, 1.0);
    let v2: f64 = divide(x2_i * y1_i, 1.0);
    let det_i: f64 = v1 + 0.0 - v2;
    let v3: f64 = divide(r_i * r_i, 1.0);
    let v4: f64 = divide(d_i * d_i, 1.0);
    let v5: f64 = divide(v3 * v4, 1.0);
    let v6: f64 = divide(det_i * det_i, 1.0);
    let v7: f64 = v5 + 0.0 - v6;
    let v8: f64 = f64::max(v7, 0.0);
    let sdel_i: f64 = f64::sqrt(f64::max(v8, 0.0));
    let v9: f64 = divide(sdy_o * dx_i, 1.0);
    let v10: f64 = divide(v9 * sdel_i, 1.0);
    let v11: f64 = divide(det_i * dy_i, 1.0);
    let dx_c1: f64 = divide(v11 + v10, v4);
    let v12: f64 = v11 + 0.0 - v10;
    let dx_c2: f64 = divide(v12 * 1.0, v4);
    let ady_i: f64 = f64::abs(dy_i);
    let v13: f64 = divide(ady_i * sdel_i, 1.0);
    let v14: f64 = divide(det_i * dx_i, -1.0);
    let dy_c1: f64 = divide(v14 + v13, v4);
    let v15: f64 = v14 + 0.0 - v13;
    let dy_c2: f64 = divide(v15 * 1.0, v4);
    let v16: f64 = x1_i + 0.0 - dx_c1;
    let v17: f64 = x1_i + 0.0 - dx_c2;
    let v18: f64 = y1_i + 0.0 - dy_c1;
    let v19: f64 = y1_i + 0.0 - dy_c2;
    let v20: f64 = f64::hypot(v16, v18).hypot(0.0);
    let v21: f64 = f64::hypot(v17, v19).hypot(0.0);
    let v22: f64 = v21 + 0.0 - v20;
    let dx_c: f64 = if v22 > 0.0 { dx_c1 } else { dx_c2 };
    let dy_c: f64 = if v22 > 0.0 { dy_c1 } else { dy_c2 };
    let sdx_c: f64 = divide(dx_c * rw2, r_i);
    let sdy_c: f64 = divide(dy_c * rh2, r_i);
    let x_c: f64 = (w / 2.0) + sdx_c - 0.0;
    let y_c: f64 = (h / 2.0) + sdy_c - 0.0;
    let ist0: f64 = f64::atan2(sdy_c, sdx_c).to_degrees() * 60000.0;
    let ist1: f64 = ist0 + 21600000.0 - 0.0;
    let ist_ang: f64 = if ist0 > 0.0 { ist0 } else { ist1 };
    let isw1: f64 = st_ang + 0.0 - ist_ang;
    let isw2: f64 = isw1 + 0.0 - 21600000.0;
    let isw_ang: f64 = if isw1 > 0.0 { isw2 } else { isw1 };
    let p1: f64 = x_f + 0.0 - x_c;
    let p2: f64 = y_f + 0.0 - y_c;
    let p3: f64 = f64::hypot(p1, p2).hypot(0.0);
    let p4: f64 = divide(p3 * 1.0, 2.0);
    let p5: f64 = p4 + 0.0 - thh;
    let x_gp: f64 = if p5 > 0.0 { x_f } else { x_g };
    let y_gp: f64 = if p5 > 0.0 { y_f } else { y_g };
    let x_bp: f64 = if p5 > 0.0 { x_c } else { x_b };
    let y_bp: f64 = if p5 > 0.0 { y_c } else { y_b };
    let en0: f64 = f64::atan2(sdy_f, sdx_f).to_degrees() * 60000.0;
    let en1: f64 = en0 + 21600000.0 - 0.0;
    let en2: f64 = if en0 > 0.0 { en0 } else { en1 };
    let sw0: f64 = en2 + 0.0 - st_ang;
    let sw1: f64 = sw0 + 21600000.0 - 0.0;
    let sw_ang: f64 = if sw0 > 0.0 { sw0 } else { sw1 };
    vec![PresetGeometryPath {
        fill: PresetPathFill::Normal,
        stroke: true,
        commands: Outline::new(frame)
            .move_to(x_e, y_e)
            .arc_to(rw1, rh1, ooxml_angle(st_ang), ooxml_angle(sw_ang))
            .line_to(x_gp, y_gp)
            .line_to(x_a, y_a)
            .line_to(x_bp, y_bp)
            .line_to(x_c, y_c)
            .arc_to(rw2, rh2, ooxml_angle(ist_ang), ooxml_angle(isw_ang))
            .close()
            .finish(),
    }]
}

fn folded_corner(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj = adjustment(adjustments, "adj", 0.16667);
    let a: f64 = f64::clamp(adj, 0.0, f64::max(50000.0, 0.0));
    let dy2: f64 = divide(100000.0 * a, 100000.0);
    let dy1: f64 = divide(dy2 * 1.0, 5.0);
    let x1: f64 = w + 0.0 - dy2;
    let x2: f64 = x1 + dy1 - 0.0;
    let y2: f64 = h + 0.0 - dy2;
    let y1: f64 = y2 + dy1 - 0.0;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(w, 0.0)
                .line_to(w, y2)
                .line_to(x1, h)
                .line_to(0.0, h)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(x1, h)
                .line_to(x2, y1)
                .line_to(w, y2)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(x1, h)
                .line_to(x2, y1)
                .line_to(w, y2)
                .line_to(x1, h)
                .line_to(0.0, h)
                .line_to(0.0, 0.0)
                .line_to(w, 0.0)
                .line_to(w, y2)
                .finish(),
        },
    ]
}

fn cloud_callout(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", -0.20833);
    let adj2 = adjustment(adjustments, "adj2", 0.625);
    let dx_pos: f64 = divide(w * adj1, 100000.0);
    let dy_pos: f64 = divide(h * adj2, 100000.0);
    let x_pos: f64 = (w / 2.0) + dx_pos - 0.0;
    let y_pos: f64 = (h / 2.0) + dy_pos - 0.0;
    let ht: f64 = (h / 2.0) * f64::atan2(dy_pos, dx_pos).cos();
    let wt: f64 = (w / 2.0) * f64::atan2(dy_pos, dx_pos).sin();
    let g2: f64 = (w / 2.0) * f64::atan2(wt, ht).cos();
    let g3: f64 = (h / 2.0) * f64::atan2(wt, ht).sin();
    let g4: f64 = (w / 2.0) + g2 - 0.0;
    let g5: f64 = (h / 2.0) + g3 - 0.0;
    let g6: f64 = g4 + 0.0 - x_pos;
    let g7: f64 = g5 + 0.0 - y_pos;
    let g8: f64 = f64::hypot(g6, g7).hypot(0.0);
    let g9: f64 = divide(100000.0 * 6600.0, 21600.0);
    let g10: f64 = g8 + 0.0 - g9;
    let g11: f64 = divide(g10 * 1.0, 3.0);
    let g12: f64 = divide(100000.0 * 1800.0, 21600.0);
    let g13: f64 = g11 + g12 - 0.0;
    let g14: f64 = divide(g13 * g6, g8);
    let g15: f64 = divide(g13 * g7, g8);
    let g16: f64 = g14 + x_pos - 0.0;
    let g17: f64 = g15 + y_pos - 0.0;
    let g18: f64 = divide(100000.0 * 4800.0, 21600.0);
    let g19: f64 = divide(g11 * 2.0, 1.0);
    let g20: f64 = g18 + g19 - 0.0;
    let g21: f64 = divide(g20 * g6, g8);
    let g22: f64 = divide(g20 * g7, g8);
    let g23: f64 = g21 + x_pos - 0.0;
    let g24: f64 = g22 + y_pos - 0.0;
    let g25: f64 = divide(100000.0 * 1200.0, 21600.0);
    let g26: f64 = divide(100000.0 * 600.0, 21600.0);
    let x23: f64 = x_pos + g26 - 0.0;
    let x24: f64 = g16 + g25 - 0.0;
    let x25: f64 = g23 + g12 - 0.0;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: true,
            commands: Outline::new(Frame {
                w: 43200.0,
                h: 43200.0,
            })
            .move_to(3900.0, 14370.0)
            .arc_to(
                6753.0,
                9190.0,
                ooxml_angle(-11429249.0),
                ooxml_angle(7426832.0),
            )
            .arc_to(
                5333.0,
                7267.0,
                ooxml_angle(-8646143.0),
                ooxml_angle(5396714.0),
            )
            .arc_to(
                4365.0,
                5945.0,
                ooxml_angle(-8748475.0),
                ooxml_angle(5983381.0),
            )
            .arc_to(
                4857.0,
                6595.0,
                ooxml_angle(-7859164.0),
                ooxml_angle(7034504.0),
            )
            .arc_to(
                5333.0,
                7273.0,
                ooxml_angle(-4722533.0),
                ooxml_angle(6541615.0),
            )
            .arc_to(
                6775.0,
                9220.0,
                ooxml_angle(-2776035.0),
                ooxml_angle(7816140.0),
            )
            .arc_to(5785.0, 7867.0, ooxml_angle(37501.0), ooxml_angle(6842000.0))
            .arc_to(
                6752.0,
                9215.0,
                ooxml_angle(1347096.0),
                ooxml_angle(6910353.0),
            )
            .arc_to(
                7720.0,
                10543.0,
                ooxml_angle(3974558.0),
                ooxml_angle(4542661.0),
            )
            .arc_to(
                4360.0,
                5918.0,
                ooxml_angle(-16496525.0),
                ooxml_angle(8804134.0),
            )
            .arc_to(
                4345.0,
                5945.0,
                ooxml_angle(-14809710.0),
                ooxml_angle(9151131.0),
            )
            .close()
            .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(x23, y_pos)
                .arc_to(g26, g26, ooxml_angle(0.0), ooxml_angle(21600000.0))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(x24, g17)
                .arc_to(g25, g25, ooxml_angle(0.0), ooxml_angle(21600000.0))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(x25, g24)
                .arc_to(g12, g12, ooxml_angle(0.0), ooxml_angle(21600000.0))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(Frame {
                w: 43200.0,
                h: 43200.0,
            })
            .move_to(4693.0, 26177.0)
            .arc_to(
                4345.0,
                5945.0,
                ooxml_angle(5204520.0),
                ooxml_angle(1585770.0),
            )
            .move_to(6928.0, 34899.0)
            .arc_to(
                4360.0,
                5918.0,
                ooxml_angle(4416628.0),
                ooxml_angle(686848.0),
            )
            .move_to(16478.0, 39090.0)
            .arc_to(
                6752.0,
                9215.0,
                ooxml_angle(8257449.0),
                ooxml_angle(844866.0),
            )
            .move_to(28827.0, 34751.0)
            .arc_to(6752.0, 9215.0, ooxml_angle(387196.0), ooxml_angle(959901.0))
            .move_to(34129.0, 22954.0)
            .arc_to(
                5785.0,
                7867.0,
                ooxml_angle(-4217541.0),
                ooxml_angle(4255042.0),
            )
            .move_to(41798.0, 15354.0)
            .arc_to(
                5333.0,
                7273.0,
                ooxml_angle(1819082.0),
                ooxml_angle(1665090.0),
            )
            .move_to(38324.0, 5426.0)
            .arc_to(
                4857.0,
                6595.0,
                ooxml_angle(-824660.0),
                ooxml_angle(891534.0),
            )
            .move_to(29078.0, 3952.0)
            .arc_to(
                4857.0,
                6595.0,
                ooxml_angle(-8950887.0),
                ooxml_angle(1091722.0),
            )
            .move_to(22141.0, 4720.0)
            .arc_to(
                4365.0,
                5945.0,
                ooxml_angle(-9809656.0),
                ooxml_angle(1061181.0),
            )
            .move_to(14000.0, 5192.0)
            .arc_to(
                6753.0,
                9190.0,
                ooxml_angle(-4002417.0),
                ooxml_angle(739161.0),
            )
            .move_to(4127.0, 15789.0)
            .arc_to(
                6753.0,
                9190.0,
                ooxml_angle(9459261.0),
                ooxml_angle(711490.0),
            )
            .finish(),
        },
    ]
}
