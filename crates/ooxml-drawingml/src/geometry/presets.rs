//! ECMA-376 preset guide and path formulas.

use super::{Frame, Outline, PresetGeometryPath, PresetPathFill, elliptical_point, ooxml_angle};
use std::collections::HashMap;
use std::f64::consts::{FRAC_PI_2, PI};

use crate::GeometryPathCommand;

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
        "actionButtonBackPrevious"
        | "actionButtonBeginning"
        | "actionButtonBlank"
        | "actionButtonDocument"
        | "actionButtonEnd"
        | "actionButtonForwardNext"
        | "actionButtonHelp"
        | "actionButtonHome"
        | "actionButtonInformation"
        | "actionButtonMovie"
        | "actionButtonReturn"
        | "actionButtonSound" => action_button(frame, kind),
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
        "round1Rect" => normal(round1_rect(frame, adjustments)),
        "round2SameRect" => normal(round2_same_rect(frame, adjustments)),
        "round2DiagRect" => normal(round2_diag_rect(frame, adjustments)),
        "snip1Rect" => normal(snip1_rect(frame, adjustments)),
        "snip2SameRect" => normal(snip2_same_rect(frame, adjustments)),
        "snip2DiagRect" => normal(snip2_diag_rect(frame, adjustments)),
        "frame" => normal(rectangular_frame(frame, adjustments)),
        "mathEqual" => normal(math_equal(frame, adjustments)),
        "mathPlus" => normal(math_plus(frame, adjustments)),
        "pie" | "chord" => normal(pie_or_chord(frame, adjustments, kind == "pie")),
        "blockArc" => normal(block_arc(frame, adjustments)),
        "flowChartExtract" => normal(
            Outline::new(Frame { w: 2.0, h: 2.0 })
                .move_to(0.0, 2.0)
                .line_to(1.0, 0.0)
                .line_to(2.0, 2.0)
                .close()
                .finish(),
        ),
        "flowChartMerge" => normal(
            Outline::new(Frame { w: 2.0, h: 2.0 })
                .move_to(0.0, 0.0)
                .line_to(2.0, 0.0)
                .line_to(1.0, 2.0)
                .close()
                .finish(),
        ),
        "flowChartDelay" => normal(
            Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(frame.w / 2.0, 0.0)
                .arc_to(frame.w / 2.0, frame.h / 2.0, 3.0 * FRAC_PI_2, PI)
                .line_to(0.0, frame.h)
                .close()
                .finish(),
        ),
        "bentUpArrow" => normal(bent_up_arrow(frame, adjustments)),
        "notchedRightArrow" => normal(notched_right_arrow(frame, adjustments)),
        "stripedRightArrow" => normal(striped_right_arrow(frame, adjustments)),
        "teardrop" => normal(teardrop(frame, adjustments)),
        "leftBracket" => bracket(frame, adjustments, true),
        "rightBracket" => bracket(frame, adjustments, false),
        "can" => can(frame, adjustments),
        "callout1" => callout1(frame, adjustments, false, false),
        "borderCallout1" => callout1(frame, adjustments, true, false),
        "accentCallout1" => callout1(frame, adjustments, false, true),
        "accentCallout2" => accent_callout2(frame, adjustments),
        "accentCallout3"
        | "accentBorderCallout1"
        | "accentBorderCallout2"
        | "accentBorderCallout3"
        | "borderCallout2"
        | "borderCallout3"
        | "callout2"
        | "callout3" => line_callout(frame, adjustments, kind),
        "leftArrowCallout" => normal(left_arrow_callout(frame, adjustments)),
        "rightArrowCallout" => normal(right_arrow_callout(frame, adjustments)),
        "upDownArrowCallout" => normal(up_down_arrow_callout(frame, adjustments)),
        "leftRightArrowCallout" => normal(left_right_arrow_callout(frame, adjustments)),
        "quadArrowCallout" => normal(quad_arrow_callout(frame, adjustments)),
        "upArrowCallout" => normal(up_arrow_callout(frame, adjustments)),
        "downArrowCallout" => normal(down_arrow_callout(frame, adjustments)),
        "quadArrow" => normal(quad_arrow(frame, adjustments)),
        "cloud" => cloud(),
        "leftCircularArrow" => left_circular_arrow(frame, adjustments),
        "curvedLeftArrow" => curved_left_arrow(frame, adjustments),
        "curvedRightArrow" => curved_right_arrow(frame, adjustments),
        "curvedUpArrow" => curved_up_arrow(frame, adjustments),
        "curvedDownArrow" => curved_down_arrow(frame, adjustments),
        "flowChartOnlineStorage" => normal(
            Outline::new(Frame { w: 6.0, h: 6.0 })
                .move_to(1.0, 0.0)
                .line_to(6.0, 0.0)
                .arc_to(1.0, 3.0, 3.0 * FRAC_PI_2, -PI)
                .line_to(1.0, 6.0)
                .arc_to(1.0, 3.0, FRAC_PI_2, PI)
                .close()
                .finish(),
        ),
        "flowChartDocument" => normal(
            Outline::new(Frame {
                w: 21600.0,
                h: 21600.0,
            })
            .move_to(0.0, 0.0)
            .line_to(21600.0, 0.0)
            .line_to(21600.0, 17322.0)
            .cubic_to(10800.0, 17322.0, 10800.0, 23922.0, 0.0, 20172.0)
            .close()
            .finish(),
        ),
        "gear6" => normal(gear6(frame, adjustments)),
        "gear9" => normal(gear9(frame, adjustments)),
        "uturnArrow" => normal(uturn_arrow(frame, adjustments)),
        "verticalScroll" => vertical_scroll(frame, adjustments),
        "horizontalScroll" => horizontal_scroll(frame, adjustments),
        "wave" => normal(wave(frame, adjustments)),
        "doubleWave" => normal(double_wave(frame, adjustments)),
        "moon" => normal(moon(frame, adjustments)),
        "heart" => normal(heart(frame)),
        "bevel" => bevel(frame, adjustments),
        "bracePair" => brace_pair(frame, adjustments),
        "bracketPair" => bracket_pair(frame, adjustments),
        "chartPlus" => chart_marks(frame, &[(5.0, 0.0, 5.0, 10.0), (0.0, 5.0, 10.0, 5.0)]),
        "chartX" => chart_marks(frame, &[(0.0, 0.0, 10.0, 10.0), (0.0, 10.0, 10.0, 0.0)]),
        "chartStar" => chart_marks(
            frame,
            &[
                (0.0, 0.0, 10.0, 10.0),
                (0.0, 10.0, 10.0, 0.0),
                (5.0, 0.0, 5.0, 10.0),
            ],
        ),
        "cornerTabs" => normal(corner_tabs(frame)),
        "plaqueTabs" => normal(plaque_tabs(frame)),
        "squareTabs" => normal(square_tabs(frame)),
        "diagStripe" => normal(diag_stripe(frame, adjustments)),
        "ellipseRibbon2" => ellipse_ribbon2(frame, adjustments),
        "flowChartCollate" => normal(
            Outline::new(Frame { w: 2.0, h: 2.0 })
                .move_to(0.0, 0.0)
                .line_to(2.0, 0.0)
                .line_to(1.0, 1.0)
                .line_to(2.0, 2.0)
                .line_to(0.0, 2.0)
                .line_to(1.0, 1.0)
                .close()
                .finish(),
        ),
        "flowChartMultidocument" => flow_chart_multidocument(),
        "flowChartOfflineStorage" => flow_chart_offline_storage(),
        "flowChartOr" => flow_chart_or(frame, false),
        "flowChartPunchedCard" => normal(
            Outline::new(Frame { w: 5.0, h: 5.0 })
                .move_to(0.0, 1.0)
                .line_to(1.0, 0.0)
                .line_to(5.0, 0.0)
                .line_to(5.0, 5.0)
                .line_to(0.0, 5.0)
                .close()
                .finish(),
        ),
        "flowChartPunchedTape" => normal(flow_chart_punched_tape()),
        "flowChartSort" => flow_chart_sort(),
        "flowChartSummingJunction" => flow_chart_or(frame, true),
        "funnel" => normal(funnel(frame)),
        "halfFrame" => normal(half_frame(frame, adjustments)),
        "irregularSeal2" => normal(irregular_seal2()),
        "leftRightCircularArrow" => normal(left_right_circular_arrow(frame, adjustments)),
        "leftRightRibbon" => left_right_ribbon(frame, adjustments),
        "leftRightUpArrow" => normal(left_right_up_arrow(frame, adjustments)),
        "leftUpArrow" => normal(left_up_arrow(frame, adjustments)),
        "lightningBolt" => normal(lightning_bolt()),
        "mathDivide" => normal(math_divide(frame, adjustments)),
        "mathMinus" => normal(math_minus(frame, adjustments)),
        "mathNotEqual" => normal(math_not_equal(frame, adjustments)),
        "nonIsoscelesTrapezoid" => normal(non_isosceles_trapezoid(frame, adjustments)),
        "pieWedge" => normal(
            Outline::new(frame)
                .move_to(0.0, frame.h)
                .arc_to(frame.w, frame.h, PI, FRAC_PI_2)
                .line_to(frame.w, frame.h)
                .close()
                .finish(),
        ),
        "plaque" => normal(plaque(frame, adjustments)),
        "smileyFace" => smiley_face(frame, adjustments),
        "snipRoundRect" => normal(snip_round_rect(frame, adjustments)),
        "sun" => normal(sun(frame, adjustments)),
        "irregularSeal1" => normal(
            Outline::new(Frame {
                w: 21600.0,
                h: 21600.0,
            })
            .move_to(10800.0, 5800.0)
            .line_to(14522.0, 0.0)
            .line_to(14155.0, 5325.0)
            .line_to(18380.0, 4457.0)
            .line_to(16702.0, 7315.0)
            .line_to(21097.0, 8137.0)
            .line_to(17607.0, 10475.0)
            .line_to(21600.0, 13290.0)
            .line_to(16837.0, 12942.0)
            .line_to(18145.0, 18095.0)
            .line_to(14020.0, 14457.0)
            .line_to(13247.0, 19737.0)
            .line_to(10532.0, 14935.0)
            .line_to(8485.0, 21600.0)
            .line_to(7715.0, 15627.0)
            .line_to(4762.0, 17617.0)
            .line_to(5667.0, 13937.0)
            .line_to(135.0, 14587.0)
            .line_to(3722.0, 11775.0)
            .line_to(0.0, 8615.0)
            .line_to(4627.0, 7617.0)
            .line_to(370.0, 2295.0)
            .line_to(7312.0, 6320.0)
            .line_to(8352.0, 2295.0)
            .close()
            .finish(),
        ),
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
        "leftCircularArrow" => &[
            ("adj1", 0.125),
            ("adj2", -11.42319),
            ("adj3", 11.42319),
            ("adj4", 108.0),
            ("adj5", 0.125),
        ],
        "foldedCorner" => &[("adj", 0.16667)],
        "cloudCallout" => &[("adj1", -0.20833), ("adj2", 0.625)],
        "round1Rect" | "snip1Rect" => &[("adj", 0.16667)],
        "round2SameRect" | "round2DiagRect" | "snip2SameRect" => {
            &[("adj1", 0.16667), ("adj2", 0.0)]
        }
        "snip2DiagRect" => &[("adj1", 0.0), ("adj2", 0.16667)],
        "frame" => &[("adj1", 0.125)],
        "mathEqual" => &[("adj1", 0.2352), ("adj2", 0.1176)],
        "mathPlus" => &[("adj1", 0.2352)],
        "pie" => &[("adj1", 0.0), ("adj2", 162.0)],
        "chord" => &[("adj1", 27.0), ("adj2", 162.0)],
        "blockArc" => &[("adj1", 108.0), ("adj2", 0.0), ("adj3", 0.25)],
        "bentUpArrow" => &[("adj1", 0.25), ("adj2", 0.25), ("adj3", 0.25)],
        "quadArrow" => &[("adj1", 0.225), ("adj2", 0.225), ("adj3", 0.225)],
        "curvedLeftArrow" | "curvedRightArrow" | "curvedUpArrow" | "curvedDownArrow" => {
            &[("adj1", 0.25), ("adj2", 0.5), ("adj3", 0.25)]
        }
        "notchedRightArrow" | "stripedRightArrow" => &[("adj1", 0.5), ("adj2", 0.5)],
        "teardrop" => &[("adj", 1.0)],
        "leftBracket" | "rightBracket" => &[("adj", 0.08333)],
        "can" => &[("adj", 0.25)],
        "callout1" | "borderCallout1" | "accentCallout1" | "accentBorderCallout1" => &[
            ("adj1", 0.1875),
            ("adj2", -0.08333),
            ("adj3", 1.125),
            ("adj4", -0.38333),
        ],
        "accentCallout2" | "accentBorderCallout2" | "borderCallout2" | "callout2" => &[
            ("adj1", 0.1875),
            ("adj2", -0.08333),
            ("adj3", 0.1875),
            ("adj4", -0.16667),
            ("adj5", 1.125),
            ("adj6", -0.46667),
        ],
        "accentCallout3" | "accentBorderCallout3" | "borderCallout3" | "callout3" => &[
            ("adj1", 0.1875),
            ("adj2", -0.08333),
            ("adj3", 0.1875),
            ("adj4", -0.16667),
            ("adj5", 1.0),
            ("adj6", -0.16667),
            ("adj7", 1.12963),
            ("adj8", -0.08333),
        ],
        "upArrowCallout" | "downArrowCallout" | "leftArrowCallout" | "rightArrowCallout" => &[
            ("adj1", 0.25),
            ("adj2", 0.25),
            ("adj3", 0.25),
            ("adj4", 0.64977),
        ],
        "upDownArrowCallout" | "leftRightArrowCallout" => &[
            ("adj1", 0.25),
            ("adj2", 0.25),
            ("adj3", 0.25),
            ("adj4", 0.48123),
        ],
        "quadArrowCallout" => &[
            ("adj1", 0.18515),
            ("adj2", 0.18515),
            ("adj3", 0.18515),
            ("adj4", 0.48123),
        ],
        "gear6" => &[("adj1", 0.15), ("adj2", 0.03526)],
        "gear9" => &[("adj1", 0.1), ("adj2", 0.01763)],
        "uturnArrow" => &[
            ("adj1", 0.25),
            ("adj2", 0.25),
            ("adj3", 0.25),
            ("adj4", 0.4375),
            ("adj5", 0.75),
        ],
        "verticalScroll" | "horizontalScroll" => &[("adj", 0.125)],
        "wave" => &[("adj1", 0.125), ("adj2", 0.0)],
        "doubleWave" => &[("adj1", 0.0625), ("adj2", 0.0)],
        "moon" => &[("adj", 0.5)],
        "bevel" => &[("adj", 0.125)],
        "bracePair" => &[("adj", 0.08333)],
        "bracketPair" => &[("adj", 0.16667)],
        "diagStripe" => &[("adj", 0.5)],
        "ellipseRibbon2" => &[("adj1", 0.25), ("adj2", 0.5), ("adj3", 0.125)],
        "halfFrame" => &[("adj1", 0.33333), ("adj2", 0.33333)],
        "leftRightCircularArrow" => &[
            ("adj1", 0.125),
            ("adj2", 11.42319),
            ("adj3", 204.57681),
            ("adj4", 119.42319),
            ("adj5", 0.125),
        ],
        "leftRightRibbon" => &[("adj1", 0.5), ("adj2", 0.5), ("adj3", 0.16667)],
        "leftRightUpArrow" | "leftUpArrow" => &[("adj1", 0.25), ("adj2", 0.25), ("adj3", 0.25)],
        "mathDivide" => &[("adj1", 0.2352), ("adj2", 0.0588), ("adj3", 0.1176)],
        "mathMinus" => &[("adj1", 0.2352)],
        "mathNotEqual" => &[("adj1", 0.2352), ("adj2", 66.0), ("adj3", 0.1176)],
        "nonIsoscelesTrapezoid" => &[("adj1", 0.25), ("adj2", 0.25)],
        "plaque" => &[("adj", 0.16667)],
        "smileyFace" => &[("adj", 0.04653)],
        "snipRoundRect" => &[("adj1", 0.16667), ("adj2", 0.16667)],
        "sun" => &[("adj", 0.25)],
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

fn normal(commands: Vec<GeometryPathCommand>) -> Vec<PresetGeometryPath> {
    vec![PresetGeometryPath {
        commands,
        fill: PresetPathFill::Normal,
        stroke: true,
    }]
}

fn pie_or_chord(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
    pie: bool,
) -> Vec<GeometryPathCommand> {
    let (wd2, hd2) = (frame.w / 2.0, frame.h / 2.0);
    let st_ang =
        adjustment(adjustments, "adj1", if pie { 0.0 } else { 27.0 }).clamp(0.0, 21599999.0);
    let en_ang = adjustment(adjustments, "adj2", 162.0).clamp(0.0, 21599999.0);
    let sw1 = en_ang - st_ang;
    let sw_ang = if sw1 > 0.0 { sw1 } else { sw1 + 21600000.0 };
    let (dx1, dy1) = elliptical_point(ooxml_angle(st_ang), wd2, hd2, frame);
    let mut outline = Outline::new(frame).move_to(wd2 + dx1, hd2 + dy1).arc_to(
        wd2,
        hd2,
        ooxml_angle(st_ang),
        ooxml_angle(sw_ang),
    );
    if pie {
        outline = outline.line_to(wd2, hd2);
    }
    outline.close().finish()
}

fn block_arc(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (wd2, hd2) = (frame.w / 2.0, frame.h / 2.0);
    let st_ang = adjustment(adjustments, "adj1", 108.0).clamp(0.0, 21599999.0);
    let ist_ang = adjustment(adjustments, "adj2", 0.0).clamp(0.0, 21599999.0);
    let a3 = adjustment(adjustments, "adj3", 0.25).clamp(0.0, 50000.0);
    let sw11 = ist_ang - st_ang;
    let sw_ang = if sw11 > 0.0 { sw11 } else { sw11 + 21600000.0 };
    let (dx1, dy1) = elliptical_point(ooxml_angle(st_ang), wd2, hd2, frame);
    let iwd2 = wd2 - a3;
    let ihd2 = hd2 - a3;
    let (dx2, dy2) = elliptical_point(ooxml_angle(ist_ang), iwd2, ihd2, frame);
    Outline::new(frame)
        .move_to(wd2 + dx1, hd2 + dy1)
        .arc_to(wd2, hd2, ooxml_angle(st_ang), ooxml_angle(sw_ang))
        .line_to(wd2 + dx2, hd2 + dy2)
        .arc_to(iwd2, ihd2, ooxml_angle(ist_ang), ooxml_angle(-sw_ang))
        .close()
        .finish()
}

fn bent_up_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.25).clamp(0.0, 50000.0);
    let a2 = adjustment(adjustments, "adj2", 0.25).clamp(0.0, 50000.0);
    let a3 = adjustment(adjustments, "adj3", 0.25).clamp(0.0, 50000.0);
    let y1 = a3;
    let x1 = w - a2 * 2.0;
    let x3 = w - a2;
    let dx2 = a1 / 2.0;
    let x2 = x3 - dx2;
    let x4 = x3 + dx2;
    let y2 = h - a1;
    Outline::new(frame)
        .move_to(0.0, y2)
        .line_to(x2, y2)
        .line_to(x2, y1)
        .line_to(x1, y1)
        .line_to(x3, 0.0)
        .line_to(w, y1)
        .line_to(x4, y1)
        .line_to(x4, h)
        .line_to(0.0, h)
        .close()
        .finish()
}

fn notched_right_arrow(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.5).clamp(0.0, 100000.0);
    let a2 = adjustment(adjustments, "adj2", 0.5).clamp(0.0, w);
    let dx2 = a2;
    let x2 = w - dx2;
    let dy1 = h * a1 / 200000.0;
    let y1 = h / 2.0 - dy1;
    let y2 = h / 2.0 + dy1;
    let x1 = dy1 * dx2 / (h / 2.0);
    Outline::new(frame)
        .move_to(0.0, y1)
        .line_to(x2, y1)
        .line_to(x2, 0.0)
        .line_to(w, h / 2.0)
        .line_to(x2, h)
        .line_to(x2, y2)
        .line_to(0.0, y2)
        .line_to(x1, h / 2.0)
        .close()
        .finish()
}

fn striped_right_arrow(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.5).clamp(0.0, 100000.0);
    let a2 = adjustment(adjustments, "adj2", 0.5).clamp(0.0, 0.84375 * w);
    let x4 = 100000.0 * 5.0 / 32.0;
    let x5 = w - a2;
    let dy1 = h * a1 / 200000.0;
    let y1 = h / 2.0 - dy1;
    let y2 = h / 2.0 + dy1;
    Outline::new(frame)
        .move_to(0.0, y1)
        .line_to(100000.0 / 32.0, y1)
        .line_to(100000.0 / 32.0, y2)
        .line_to(0.0, y2)
        .close()
        .move_to(100000.0 / 16.0, y1)
        .line_to(100000.0 / 8.0, y1)
        .line_to(100000.0 / 8.0, y2)
        .line_to(100000.0 / 16.0, y2)
        .close()
        .move_to(x4, y1)
        .line_to(x5, y1)
        .line_to(x5, 0.0)
        .line_to(w, h / 2.0)
        .line_to(x5, h)
        .line_to(x5, y2)
        .line_to(x4, y2)
        .close()
        .finish()
}

fn teardrop(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 1.0).clamp(0.0, 200000.0);
    let tw = w / 2.0 * 2.0f64.sqrt();
    let th = h / 2.0 * 2.0f64.sqrt();
    let sw = tw * a / 100000.0;
    let sh = th * a / 100000.0;
    let dx1 = sw * ooxml_angle(2700000.0).cos();
    let dy1 = sh * ooxml_angle(2700000.0).sin();
    let x1 = w / 2.0 + dx1;
    let y1 = h / 2.0 - dy1;
    let x2 = (w / 2.0 + x1) / 2.0;
    let y2 = (h / 2.0 + y1) / 2.0;
    Outline::new(frame)
        .move_to(0.0, h / 2.0)
        .arc_to(w / 2.0, h / 2.0, PI, FRAC_PI_2)
        .quad_to(x2, 0.0, x1, y1)
        .quad_to(w, y2, w, h / 2.0)
        .arc_to(w / 2.0, h / 2.0, 0.0, FRAC_PI_2)
        .arc_to(w / 2.0, h / 2.0, FRAC_PI_2, FRAC_PI_2)
        .close()
        .finish()
}

fn bracket(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
    left: bool,
) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let y1 = adjustment(adjustments, "adj", 0.08333).clamp(0.0, h / 2.0);
    let y2 = h - y1;
    let outline = if left {
        Outline::new(frame)
            .move_to(w, h)
            .arc_to(w, y1, FRAC_PI_2, FRAC_PI_2)
            .line_to(0.0, y1)
            .arc_to(w, y1, PI, FRAC_PI_2)
    } else {
        Outline::new(frame)
            .move_to(0.0, 0.0)
            .arc_to(w, y1, 3.0 * FRAC_PI_2, FRAC_PI_2)
            .line_to(w, y2)
            .arc_to(w, y1, 0.0, FRAC_PI_2)
    };
    let commands = outline.finish();
    let mut fill = commands.clone();
    fill.push(GeometryPathCommand::Close);
    vec![
        PresetGeometryPath {
            commands: fill,
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn can(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.25).clamp(0.0, h / 2.0);
    let y1 = a / 2.0;
    let y3 = h - y1;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, y1)
                .arc_to(w / 2.0, y1, PI, -PI)
                .line_to(w, y3)
                .arc_to(w / 2.0, y1, 0.0, PI)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::Lighten,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, y1)
                .arc_to(w / 2.0, y1, PI, PI)
                .arc_to(w / 2.0, y1, 0.0, PI)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(w, y1)
                .arc_to(w / 2.0, y1, 0.0, PI)
                .arc_to(w / 2.0, y1, PI, PI)
                .line_to(w, y3)
                .arc_to(w / 2.0, y1, 0.0, PI)
                .line_to(0.0, y1)
                .finish(),
        },
    ]
}

fn round1_rect(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let dx1 = adjustment(adjustments, "adj", 0.16667).clamp(0.0, 50000.0);
    let x1 = w - dx1;
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(x1, 0.0)
        .arc_to(dx1, dx1, 3.0 * FRAC_PI_2, FRAC_PI_2)
        .line_to(w, h)
        .line_to(0.0, h)
        .close()
        .finish()
}

fn round2_same_rect(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let tx1 = adjustment(adjustments, "adj1", 0.16667).clamp(0.0, 50000.0);
    let tx2 = w - tx1;
    let bx1 = adjustment(adjustments, "adj2", 0.0).clamp(0.0, 50000.0);
    let by1 = h - bx1;
    Outline::new(frame)
        .move_to(tx1, 0.0)
        .line_to(tx2, 0.0)
        .arc_to(tx1, tx1, 3.0 * FRAC_PI_2, FRAC_PI_2)
        .line_to(w, by1)
        .arc_to(bx1, bx1, 0.0, FRAC_PI_2)
        .line_to(bx1, h)
        .arc_to(bx1, bx1, FRAC_PI_2, FRAC_PI_2)
        .line_to(0.0, tx1)
        .arc_to(tx1, tx1, PI, FRAC_PI_2)
        .close()
        .finish()
}

fn round2_diag_rect(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let x1 = adjustment(adjustments, "adj1", 0.16667).clamp(0.0, 50000.0);
    let y1 = h - x1;
    let a = adjustment(adjustments, "adj2", 0.0).clamp(0.0, 50000.0);
    let x2 = w - a;
    Outline::new(frame)
        .move_to(x1, 0.0)
        .line_to(x2, 0.0)
        .arc_to(a, a, 3.0 * FRAC_PI_2, FRAC_PI_2)
        .line_to(w, y1)
        .arc_to(x1, x1, 0.0, FRAC_PI_2)
        .line_to(a, h)
        .arc_to(a, a, FRAC_PI_2, FRAC_PI_2)
        .line_to(0.0, x1)
        .arc_to(x1, x1, PI, FRAC_PI_2)
        .close()
        .finish()
}

fn snip1_rect(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let dx1 = adjustment(adjustments, "adj", 0.16667).clamp(0.0, 50000.0);
    let x1 = w - dx1;
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(x1, 0.0)
        .line_to(w, dx1)
        .line_to(w, h)
        .line_to(0.0, h)
        .close()
        .finish()
}

fn snip2_same_rect(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let tx1 = adjustment(adjustments, "adj1", 0.16667).clamp(0.0, 50000.0);
    let tx2 = w - tx1;
    let bx1 = adjustment(adjustments, "adj2", 0.0).clamp(0.0, 50000.0);
    let bx2 = w - bx1;
    let by1 = h - bx1;
    Outline::new(frame)
        .move_to(tx1, 0.0)
        .line_to(tx2, 0.0)
        .line_to(w, tx1)
        .line_to(w, by1)
        .line_to(bx2, h)
        .line_to(bx1, h)
        .line_to(0.0, by1)
        .line_to(0.0, tx1)
        .close()
        .finish()
}

fn snip2_diag_rect(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let lx1 = adjustment(adjustments, "adj1", 0.0).clamp(0.0, 50000.0);
    let lx2 = w - lx1;
    let ly1 = h - lx1;
    let rx1 = adjustment(adjustments, "adj2", 0.16667).clamp(0.0, 50000.0);
    let rx2 = w - rx1;
    let ry1 = h - rx1;
    Outline::new(frame)
        .move_to(lx1, 0.0)
        .line_to(rx2, 0.0)
        .line_to(w, rx1)
        .line_to(w, ly1)
        .line_to(lx2, h)
        .line_to(rx1, h)
        .line_to(0.0, ry1)
        .line_to(0.0, lx1)
        .close()
        .finish()
}

fn rectangular_frame(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let x1 = adjustment(adjustments, "adj1", 0.125).clamp(0.0, 50000.0);
    let x4 = w - x1;
    let y4 = h - x1;
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(w, 0.0)
        .line_to(w, h)
        .line_to(0.0, h)
        .close()
        .move_to(x1, x1)
        .line_to(x1, y4)
        .line_to(x4, y4)
        .line_to(x4, x1)
        .close()
        .finish()
}

fn math_equal(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.2352).clamp(0.0, 36745.0);
    let a2 = adjustment(adjustments, "adj2", 0.1176).clamp(0.0, 100000.0 - 2.0 * a1);
    let dy1 = h * a1 / 100000.0;
    let dy2 = h * a2 / 200000.0;
    let dx1 = w * 73490.0 / 200000.0;
    let y2 = h / 2.0 - dy2;
    let y3 = h / 2.0 + dy2;
    let y1 = y2 - dy1;
    let y4 = y3 + dy1;
    let x1 = w / 2.0 - dx1;
    let x2 = w / 2.0 + dx1;
    Outline::new(frame)
        .move_to(x1, y1)
        .line_to(x2, y1)
        .line_to(x2, y2)
        .line_to(x1, y2)
        .close()
        .move_to(x1, y3)
        .line_to(x2, y3)
        .line_to(x2, y4)
        .line_to(x1, y4)
        .close()
        .finish()
}

fn math_plus(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.2352).clamp(0.0, 73490.0);
    let dx1 = w * 73490.0 / 200000.0;
    let dy1 = h * 73490.0 / 200000.0;
    let dx2 = a1 / 2.0;
    let (x1, x2, x3, x4) = (w / 2.0 - dx1, w / 2.0 - dx2, w / 2.0 + dx2, w / 2.0 + dx1);
    let (y1, y2, y3, y4) = (h / 2.0 - dy1, h / 2.0 - dx2, h / 2.0 + dx2, h / 2.0 + dy1);
    Outline::new(frame)
        .move_to(x1, y2)
        .line_to(x2, y2)
        .line_to(x2, y1)
        .line_to(x3, y1)
        .line_to(x3, y2)
        .line_to(x4, y2)
        .line_to(x4, y3)
        .line_to(x3, y3)
        .line_to(x3, y4)
        .line_to(x2, y4)
        .line_to(x2, y3)
        .line_to(x1, y3)
        .close()
        .finish()
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
    let (dx1, dy1) = elliptical_point(ooxml_angle(st_ang), w / 2.0, h / 2.0, frame);
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
    let (dx_h, dy_h) = elliptical_point(ooxml_angle(en_ang), rw3, rh3, frame);
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
    let (dx_a, dy_a) = elliptical_point(ooxml_angle(pt_ang), rw3, rh3, frame);
    let x_a: f64 = (w / 2.0) + dx_a - 0.0;
    let y_a: f64 = (h / 2.0) + dy_a - 0.0;
    let (dx_e, dy_e) = elliptical_point(ooxml_angle(st_ang), rw1, rh1, frame);
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
        cloud_body(),
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
        cloud_accents(),
    ]
}

fn cloud() -> Vec<PresetGeometryPath> {
    vec![cloud_body(), cloud_accents()]
}

fn cloud_body() -> PresetGeometryPath {
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
    }
}

fn cloud_accents() -> PresetGeometryPath {
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
    }
}

fn callout1(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
    border: bool,
    accent: bool,
) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.1875);
    let adj2 = adjustment(adjustments, "adj2", -0.08333);
    let adj3 = adjustment(adjustments, "adj3", 1.125);
    let adj4 = adjustment(adjustments, "adj4", -0.38333);
    let y1 = divide(h * adj1, 100000.0);
    let x1 = divide(w * adj2, 100000.0);
    let y2 = divide(h * adj3, 100000.0);
    let x2 = divide(w * adj4, 100000.0);
    let mut leader = Outline::new(frame);
    if accent {
        leader = leader.move_to(x1, 0.0).line_to(x1, h);
    }
    leader = leader.move_to(x1, y1).line_to(x2, y2);
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: border,
            commands: Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(w, 0.0)
                .line_to(w, h)
                .line_to(0.0, h)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: leader.finish(),
        },
    ]
}

fn accent_callout2(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.1875);
    let adj2 = adjustment(adjustments, "adj2", -0.08333);
    let adj3 = adjustment(adjustments, "adj3", 0.1875);
    let adj4 = adjustment(adjustments, "adj4", -0.16667);
    let adj5 = adjustment(adjustments, "adj5", 1.125);
    let adj6 = adjustment(adjustments, "adj6", -0.46667);
    let y1 = divide(h * adj1, 100000.0);
    let x1 = divide(w * adj2, 100000.0);
    let y2 = divide(h * adj3, 100000.0);
    let x2 = divide(w * adj4, 100000.0);
    let y3 = divide(h * adj5, 100000.0);
    let x3 = divide(w * adj6, 100000.0);
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(w, 0.0)
                .line_to(w, h)
                .line_to(0.0, h)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(x1, 0.0)
                .line_to(x1, h)
                .move_to(x1, y1)
                .line_to(x2, y2)
                .line_to(x3, y3)
                .finish(),
        },
    ]
}

fn line_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
    kind: &str,
) -> Vec<PresetGeometryPath> {
    let values = defaults(kind);
    let x1 = divide(
        frame.w * adjustment(adjustments, values[1].0, values[1].1),
        100000.0,
    );
    let mut leader = Outline::new(frame);
    if kind.starts_with("accent") {
        leader = leader.move_to(x1, 0.0).close().line_to(x1, frame.h);
    }
    for (i, pair) in values.as_chunks::<2>().0.iter().enumerate() {
        let y = divide(
            frame.h * adjustment(adjustments, pair[0].0, pair[0].1),
            100000.0,
        );
        let x = divide(
            frame.w * adjustment(adjustments, pair[1].0, pair[1].1),
            100000.0,
        );
        leader = if i == 0 {
            leader.move_to(x, y)
        } else {
            leader.line_to(x, y)
        };
    }
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: kind.contains("Border") || kind.starts_with("border"),
            commands: Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(frame.w, 0.0)
                .line_to(frame.w, frame.h)
                .line_to(0.0, frame.h)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: leader.finish(),
        },
    ]
}

fn up_arrow_callout(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.25);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let adj4 = adjustment(adjustments, "adj4", 0.64977);
    let max_adj2 = divide(50000.0 * w, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = divide(100000.0 * h, 100000.0);
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q2 = divide(a3 * 100000.0, h);
    let max_adj4 = 100000.0 - q2;
    let a4 = adj4.clamp(0.0, max_adj4.max(0.0));
    let dx1 = a2;
    let dx2 = a1 / 2.0;
    let x1 = w / 2.0 - dx1;
    let x2 = w / 2.0 - dx2;
    let x3 = w / 2.0 + dx2;
    let x4 = w / 2.0 + dx1;
    let y1 = a3;
    let dy2 = divide(h * a4, 100000.0);
    let y2 = h - dy2;
    Outline::new(frame)
        .move_to(0.0, y2)
        .line_to(x2, y2)
        .line_to(x2, y1)
        .line_to(x1, y1)
        .line_to(w / 2.0, 0.0)
        .line_to(x4, y1)
        .line_to(x3, y1)
        .line_to(x3, y2)
        .line_to(w, y2)
        .line_to(w, h)
        .line_to(0.0, h)
        .close()
        .finish()
}

fn down_arrow_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.25);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let adj4 = adjustment(adjustments, "adj4", 0.64977);
    let max_adj2 = divide(50000.0 * w, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = divide(100000.0 * h, 100000.0);
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q2 = divide(a3 * 100000.0, h);
    let max_adj4 = 100000.0 - q2;
    let a4 = adj4.clamp(0.0, max_adj4.max(0.0));
    let dx1 = a2;
    let dx2 = a1 / 2.0;
    let x1 = w / 2.0 - dx1;
    let x2 = w / 2.0 - dx2;
    let x3 = w / 2.0 + dx2;
    let x4 = w / 2.0 + dx1;
    let y1 = divide(h * a4, 100000.0);
    let y2 = h - a3;
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(w, 0.0)
        .line_to(w, y1)
        .line_to(x3, y1)
        .line_to(x3, y2)
        .line_to(x4, y2)
        .line_to(w / 2.0, h)
        .line_to(x1, y2)
        .line_to(x2, y2)
        .line_to(x2, y1)
        .line_to(0.0, y1)
        .close()
        .finish()
}

fn quad_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.225);
    let adj2 = adjustment(adjustments, "adj2", 0.225);
    let adj3 = adjustment(adjustments, "adj3", 0.225);
    let a2 = adj2.clamp(0.0, 50000.0);
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let q1 = 100000.0 - max_adj1;
    let max_adj3 = q1 / 2.0;
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let x1 = a3;
    let dx2 = a2;
    let x2 = w / 2.0 - dx2;
    let x5 = w / 2.0 + dx2;
    let dx3 = a1 / 2.0;
    let x3 = w / 2.0 - dx3;
    let x4 = w / 2.0 + dx3;
    let x6 = w - x1;
    let y1 = x1;
    let y2 = h / 2.0 - dx2;
    let y5 = h / 2.0 + dx2;
    let y3 = h / 2.0 - dx3;
    let y4 = h / 2.0 + dx3;
    let y6 = h - y1;
    Outline::new(frame)
        .move_to(0.0, h / 2.0)
        .line_to(x1, y2)
        .line_to(x1, y3)
        .line_to(x3, y3)
        .line_to(x3, y1)
        .line_to(x2, y1)
        .line_to(w / 2.0, 0.0)
        .line_to(x5, y1)
        .line_to(x4, y1)
        .line_to(x4, y3)
        .line_to(x6, y3)
        .line_to(x6, y2)
        .line_to(w, h / 2.0)
        .line_to(x6, y5)
        .line_to(x6, y4)
        .line_to(x4, y4)
        .line_to(x4, y6)
        .line_to(x5, y6)
        .line_to(w / 2.0, h)
        .line_to(x2, y6)
        .line_to(x3, y6)
        .line_to(x3, y4)
        .line_to(x1, y4)
        .line_to(x1, y5)
        .close()
        .finish()
}

fn left_circular_arrow(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.125);
    let adj2 = adjustment(adjustments, "adj2", -11.42319);
    let adj3 = adjustment(adjustments, "adj3", 11.42319);
    let adj4 = adjustment(adjustments, "adj4", 108.0);
    let adj5 = adjustment(adjustments, "adj5", 0.125);
    let a5 = adj5.clamp(0.0, 25000.0);
    let max_adj1 = a5 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let en_ang = adj3.clamp(1.0, 21599999.0);
    let st_ang = adj4.clamp(0.0, 21599999.0);
    let th = a1;
    let thh = a5;
    let th2 = th / 2.0;
    let rw1 = w / 2.0 + th2 - thh;
    let rh1 = h / 2.0 + th2 - thh;
    let rw2 = rw1 - th;
    let rh2 = rh1 - th;
    let rw3 = rw2 + th2;
    let rh3 = rh2 + th2;
    let (dx_h, dy_h) = elliptical_point(ooxml_angle(en_ang), rw3, rh3, frame);
    let x_h = w / 2.0 + dx_h;
    let y_h = h / 2.0 + dy_h;
    let r_i = rw2.min(rh2);
    let u1 = dx_h * dx_h;
    let u2 = dy_h * dy_h;
    let u3 = r_i * r_i;
    let u4 = u1 - u3;
    let u5 = u2 - u3;
    let u6 = divide(u4 * u5, u1);
    let u7 = divide(u6, u2);
    let u8 = 1.0 - u7;
    let u9 = u8.max(0.0).sqrt();
    let u10 = divide(u4, dx_h);
    let u11 = divide(u10, dy_h);
    let u12 = divide(1.0 + u9, u11);
    let u13 = f64::atan2(u12, 1.0).to_degrees() * 60000.0;
    let u15 = if u13 > 0.0 { u13 } else { u13 + 21600000.0 };
    let u16 = u15 - en_ang;
    let u18 = if u16 > 0.0 { u16 } else { u16 + 21600000.0 };
    let u21 = if u18 - 10800000.0 > 0.0 {
        u18 - 21600000.0
    } else {
        u18
    };
    let min_ang = -u21.abs();
    let a2 = -adj2.abs();
    let a_ang = a2.clamp(min_ang, 0.0_f64.max(min_ang));
    let pt_ang = en_ang + a_ang;
    let (dx_a, dy_a) = elliptical_point(ooxml_angle(pt_ang), rw3, rh3, frame);
    let x_a = w / 2.0 + dx_a;
    let y_a = h / 2.0 + dy_a;
    let (dx_e, dy_e) = elliptical_point(ooxml_angle(st_ang), rw1, rh1, frame);
    let x_e = w / 2.0 + dx_e;
    let y_e = h / 2.0 + dy_e;
    let (dx_d, dy_d) = elliptical_point(ooxml_angle(st_ang), rw2, rh2, frame);
    let x_d = w / 2.0 + dx_d;
    let y_d = h / 2.0 + dy_d;
    let dx_g = thh * ooxml_angle(pt_ang).cos();
    let dy_g = thh * ooxml_angle(pt_ang).sin();
    let x_g = x_h + dx_g;
    let y_g = y_h + dy_g;
    let x_b = x_h - dx_g;
    let y_b = y_h - dy_g;
    let sx1 = x_b - w / 2.0;
    let sy1 = y_b - h / 2.0;
    let sx2 = x_g - w / 2.0;
    let sy2 = y_g - h / 2.0;
    let r_o = rw1.min(rh1);
    let x1_o = divide(sx1 * r_o, rw1);
    let y1_o = divide(sy1 * r_o, rh1);
    let x2_o = divide(sx2 * r_o, rw1);
    let y2_o = divide(sy2 * r_o, rh1);
    let dx_o = x2_o - x1_o;
    let dy_o = y2_o - y1_o;
    let d_o = f64::hypot(dx_o, dy_o);
    let det_o = x1_o * y2_o - x2_o * y1_o;
    let q4 = d_o * d_o;
    let q7 = r_o * r_o * q4 - det_o * det_o;
    let sdel_o = q7.max(0.0).sqrt();
    let sdy_o = if -dy_o > 0.0 { -1.0 } else { 1.0 };
    let q10 = sdy_o * dx_o * sdel_o;
    let q11 = det_o * dy_o;
    let dx_f1 = divide(q11 + q10, q4);
    let dx_f2 = divide(q11 - q10, q4);
    let q13 = dy_o.abs() * sdel_o;
    let q14 = -det_o * dx_o;
    let dy_f1 = divide(q14 + q13, q4);
    let dy_f2 = divide(q14 - q13, q4);
    let q20 = f64::hypot(x2_o - dx_f1, y2_o - dy_f1);
    let q21 = f64::hypot(x2_o - dx_f2, y2_o - dy_f2);
    let (dx_f, dy_f) = if q21 > q20 {
        (dx_f1, dy_f1)
    } else {
        (dx_f2, dy_f2)
    };
    let sdx_f = divide(dx_f * rw1, r_o);
    let sdy_f = divide(dy_f * rh1, r_o);
    let x_f = w / 2.0 + sdx_f;
    let y_f = h / 2.0 + sdy_f;
    let x1_i = divide(sx1 * r_i, rw2);
    let y1_i = divide(sy1 * r_i, rh2);
    let x2_i = divide(sx2 * r_i, rw2);
    let y2_i = divide(sy2 * r_i, rh2);
    let dx_i = x2_i - x1_i;
    let dy_i = y2_i - y1_i;
    let d_i = f64::hypot(dx_i, dy_i);
    let det_i = x1_i * y2_i - x2_i * y1_i;
    let v4 = d_i * d_i;
    let v7 = r_i * r_i * v4 - det_i * det_i;
    let sdel_i = v7.max(0.0).sqrt();
    let v10 = sdy_o * dx_i * sdel_i;
    let v11 = det_i * dy_i;
    let dx_c1 = divide(v11 + v10, v4);
    let dx_c2 = divide(v11 - v10, v4);
    let v13 = dy_i.abs() * sdel_i;
    let v14 = -det_i * dx_i;
    let dy_c1 = divide(v14 + v13, v4);
    let dy_c2 = divide(v14 - v13, v4);
    let v20 = f64::hypot(x1_i - dx_c1, y1_i - dy_c1);
    let v21 = f64::hypot(x1_i - dx_c2, y1_i - dy_c2);
    let (dx_c, dy_c) = if v21 > v20 {
        (dx_c1, dy_c1)
    } else {
        (dx_c2, dy_c2)
    };
    let sdx_c = divide(dx_c * rw2, r_i);
    let sdy_c = divide(dy_c * rh2, r_i);
    let x_c = w / 2.0 + sdx_c;
    let y_c = h / 2.0 + sdy_c;
    let ist0 = f64::atan2(sdy_c, sdx_c).to_degrees() * 60000.0;
    let ist_ang0 = if ist0 > 0.0 { ist0 } else { ist0 + 21600000.0 };
    let isw1 = st_ang - ist_ang0;
    let isw_ang0 = if isw1 > 0.0 { isw1 } else { isw1 + 21600000.0 };
    let ist_ang = ist_ang0 + isw_ang0;
    let isw_ang = -isw_ang0;
    let p5 = f64::hypot(x_f - x_c, y_f - y_c) / 2.0 - thh;
    let (x_gp, y_gp, x_bp, y_bp) = if p5 > 0.0 {
        (x_f, y_f, x_c, y_c)
    } else {
        (x_g, y_g, x_b, y_b)
    };
    let en0 = f64::atan2(sdy_f, sdx_f).to_degrees() * 60000.0;
    let en2 = if en0 > 0.0 { en0 } else { en0 + 21600000.0 };
    let sw0 = en2 - st_ang;
    let sw_ang = if sw0 > 0.0 { sw0 - 21600000.0 } else { sw0 };
    let st_ang0 = st_ang + sw_ang;
    let sw_ang0 = -sw_ang;
    vec![PresetGeometryPath {
        fill: PresetPathFill::Normal,
        stroke: true,
        commands: Outline::new(frame)
            .move_to(x_e, y_e)
            .line_to(x_d, y_d)
            .arc_to(rw2, rh2, ooxml_angle(ist_ang), ooxml_angle(isw_ang))
            .line_to(x_bp, y_bp)
            .line_to(x_a, y_a)
            .line_to(x_gp, y_gp)
            .line_to(x_f, y_f)
            .arc_to(rw1, rh1, ooxml_angle(st_ang0), ooxml_angle(sw_ang0))
            .close()
            .finish(),
    }]
}

fn curved_left_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.5);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let max_adj2 = divide(50000.0 * h, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let a1 = adj1.clamp(0.0, a2.max(0.0));
    let th = a1;
    let aw = a2;
    let q1 = (th + aw) / 4.0;
    let h_r = h / 2.0 - q1;
    let q7 = h_r * 2.0;
    let q10 = q7 * q7 - th * th;
    let q11 = q10.max(0.0).sqrt();
    let idx = divide(q11 * w, q7);
    let max_adj3 = idx;
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let ah = a3;
    let y3 = h_r + th;
    let q4 = w * w - ah * ah;
    let q5 = q4.max(0.0).sqrt();
    let dy = divide(q5 * h_r, w);
    let y5 = h_r + dy;
    let y7 = y3 + dy;
    let dh = (aw - th) / 2.0;
    let y4 = y5 - dh;
    let y8 = y7 + dh;
    let y6 = h - aw / 2.0;
    let x1 = ah;
    let sw_ang = f64::atan2(dy, ah).to_degrees() * 60000.0;
    let q12 = th / 2.0;
    let dang2 = f64::atan2(q12, idx).to_degrees() * 60000.0;
    let sw_ang2 = dang2 - sw_ang;
    let sw_ang3 = sw_ang + dang2;
    let st_ang3 = -dang2;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, y6)
                .line_to(x1, y4)
                .line_to(x1, y5)
                .arc_to(w, h_r, ooxml_angle(sw_ang), ooxml_angle(sw_ang2))
                .arc_to(w, h_r, ooxml_angle(st_ang3), ooxml_angle(sw_ang3))
                .line_to(x1, y8)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(w, y3)
                .arc_to(w, h_r, ooxml_angle(0.0), ooxml_angle(-5400000.0))
                .line_to(0.0, 0.0)
                .arc_to(w, h_r, ooxml_angle(16200000.0), ooxml_angle(5400000.0))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(w, y3)
                .arc_to(w, h_r, ooxml_angle(0.0), ooxml_angle(-5400000.0))
                .line_to(0.0, 0.0)
                .arc_to(w, h_r, ooxml_angle(16200000.0), ooxml_angle(5400000.0))
                .line_to(w, y3)
                .arc_to(w, h_r, ooxml_angle(0.0), ooxml_angle(sw_ang))
                .line_to(x1, y8)
                .line_to(0.0, y6)
                .line_to(x1, y4)
                .line_to(x1, y5)
                .arc_to(w, h_r, ooxml_angle(sw_ang), ooxml_angle(sw_ang2))
                .finish(),
        },
    ]
}

fn curved_right_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.5);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let max_adj2 = divide(50000.0 * h, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let a1 = adj1.clamp(0.0, a2.max(0.0));
    let th = a1;
    let aw = a2;
    let q1 = (th + aw) / 4.0;
    let h_r = h / 2.0 - q1;
    let q7 = h_r * 2.0;
    let q10 = q7 * q7 - th * th;
    let q11 = q10.max(0.0).sqrt();
    let idx = divide(q11 * w, q7);
    let max_adj3 = idx;
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let ah = a3;
    let y3 = h_r + th;
    let q4 = w * w - ah * ah;
    let q5 = q4.max(0.0).sqrt();
    let dy = divide(q5 * h_r, w);
    let y5 = h_r + dy;
    let y7 = y3 + dy;
    let dh = (aw - th) / 2.0;
    let y4 = y5 - dh;
    let y8 = y7 + dh;
    let y6 = h - aw / 2.0;
    let x1 = w - ah;
    let sw_ang = f64::atan2(dy, ah).to_degrees() * 60000.0;
    let st_ang = 10800000.0 - sw_ang;
    let msw_ang = -sw_ang;
    let q12 = th / 2.0;
    let dang2 = f64::atan2(q12, idx).to_degrees() * 60000.0;
    let sw_ang2 = dang2 - 5400000.0;
    let sw_ang3 = 5400000.0 + dang2;
    let st_ang3 = 10800000.0 - dang2;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(0.0, h_r)
                .arc_to(w, h_r, ooxml_angle(10800000.0), ooxml_angle(msw_ang))
                .line_to(x1, y4)
                .line_to(w, y6)
                .line_to(x1, y8)
                .line_to(x1, y7)
                .arc_to(w, h_r, ooxml_angle(st_ang), ooxml_angle(sw_ang))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(w, th)
                .arc_to(w, h_r, ooxml_angle(16200000.0), ooxml_angle(sw_ang2))
                .arc_to(w, h_r, ooxml_angle(st_ang3), ooxml_angle(sw_ang3))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(0.0, h_r)
                .arc_to(w, h_r, ooxml_angle(10800000.0), ooxml_angle(msw_ang))
                .line_to(x1, y4)
                .line_to(w, y6)
                .line_to(x1, y8)
                .line_to(x1, y7)
                .arc_to(w, h_r, ooxml_angle(st_ang), ooxml_angle(sw_ang))
                .line_to(0.0, h_r)
                .arc_to(w, h_r, ooxml_angle(10800000.0), ooxml_angle(5400000.0))
                .line_to(w, th)
                .arc_to(w, h_r, ooxml_angle(16200000.0), ooxml_angle(sw_ang2))
                .finish(),
        },
    ]
}

fn curved_up_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.5);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let max_adj2 = divide(50000.0 * w, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let a1 = adj1.clamp(0.0, a2.max(0.0));
    let th = a1;
    let aw = a2;
    let q1 = (th + aw) / 4.0;
    let w_r = w / 2.0 - q1;
    let q7 = w_r * 2.0;
    let q10 = q7 * q7 - th * th;
    let q11 = q10.max(0.0).sqrt();
    let idy = divide(q11 * h, q7);
    let max_adj3 = idy;
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let ah = a3;
    let x3 = w_r + th;
    let q4 = h * h - ah * ah;
    let q5 = q4.max(0.0).sqrt();
    let dx = divide(q5 * w_r, h);
    let x5 = w_r + dx;
    let x7 = x3 + dx;
    let dh = (aw - th) / 2.0;
    let x4 = x5 - dh;
    let x8 = x7 + dh;
    let x6 = w - aw / 2.0;
    let y1 = ah;
    let sw_ang = f64::atan2(dx, ah).to_degrees() * 60000.0;
    let iy = idy;
    let ix = (w_r + x3) / 2.0;
    let q12 = th / 2.0;
    let dang2 = f64::atan2(q12, idy).to_degrees() * 60000.0;
    let sw_ang2 = dang2 - sw_ang;
    let st_ang3 = 5400000.0 - sw_ang;
    let sw_ang3 = sw_ang + dang2;
    let st_ang2 = 5400000.0 - dang2;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(x6, 0.0)
                .line_to(x8, y1)
                .line_to(x7, y1)
                .arc_to(w_r, h, ooxml_angle(st_ang3), ooxml_angle(sw_ang3))
                .arc_to(w_r, h, ooxml_angle(st_ang2), ooxml_angle(sw_ang2))
                .line_to(x4, y1)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(w_r, h)
                .arc_to(w_r, h, ooxml_angle(5400000.0), ooxml_angle(5400000.0))
                .line_to(th, 0.0)
                .arc_to(w_r, h, ooxml_angle(10800000.0), ooxml_angle(-5400000.0))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(ix, iy)
                .arc_to(w_r, h, ooxml_angle(st_ang2), ooxml_angle(sw_ang2))
                .line_to(x4, y1)
                .line_to(x6, 0.0)
                .line_to(x8, y1)
                .line_to(x7, y1)
                .arc_to(w_r, h, ooxml_angle(st_ang3), ooxml_angle(sw_ang))
                .line_to(w_r, h)
                .arc_to(w_r, h, ooxml_angle(5400000.0), ooxml_angle(5400000.0))
                .line_to(th, 0.0)
                .arc_to(w_r, h, ooxml_angle(10800000.0), ooxml_angle(-5400000.0))
                .finish(),
        },
    ]
}

fn curved_down_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.5);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let max_adj2 = divide(50000.0 * w, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let a1 = adj1.clamp(0.0, 100000.0);
    let th = a1;
    let aw = a2;
    let q1 = (th + aw) / 4.0;
    let w_r = w / 2.0 - q1;
    let q7 = w_r * 2.0;
    let q10 = q7 * q7 - th * th;
    let q11 = q10.max(0.0).sqrt();
    let idy = divide(q11 * h, q7);
    let ah = adj3;
    let x3 = w_r + th;
    let q4 = h * h - ah * ah;
    let q5 = q4.max(0.0).sqrt();
    let dx = divide(q5 * w_r, h);
    let x5 = w_r + dx;
    let x7 = x3 + dx;
    let dh = (aw - th) / 2.0;
    let x4 = x5 - dh;
    let x8 = x7 + dh;
    let x6 = w - aw / 2.0;
    let y1 = h - ah;
    let sw_ang = f64::atan2(dx, ah).to_degrees() * 60000.0;
    let st_ang = 16200000.0 + sw_ang;
    let q12 = th / 2.0;
    let dang2 = f64::atan2(q12, idy).to_degrees() * 60000.0;
    let msw_ang = -sw_ang;
    let iy = h - idy;
    let ix = (w_r + x3) / 2.0;
    let st_ang2 = 16200000.0 - dang2;
    let sw_ang2 = dang2 - 5400000.0;
    let sw_ang3 = 5400000.0 + dang2;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(x6, h)
                .line_to(x4, y1)
                .line_to(x5, y1)
                .arc_to(w_r, h, ooxml_angle(st_ang), ooxml_angle(msw_ang))
                .line_to(x3, 0.0)
                .arc_to(w_r, h, ooxml_angle(16200000.0), ooxml_angle(sw_ang))
                .line_to(x8, y1)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(ix, iy)
                .arc_to(w_r, h, ooxml_angle(st_ang2), ooxml_angle(sw_ang2))
                .line_to(0.0, h)
                .arc_to(w_r, h, ooxml_angle(10800000.0), ooxml_angle(sw_ang3))
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(ix, iy)
                .arc_to(w_r, h, ooxml_angle(st_ang2), ooxml_angle(sw_ang2))
                .line_to(0.0, h)
                .arc_to(w_r, h, ooxml_angle(10800000.0), ooxml_angle(5400000.0))
                .line_to(x3, 0.0)
                .arc_to(w_r, h, ooxml_angle(16200000.0), ooxml_angle(sw_ang))
                .line_to(x8, y1)
                .line_to(x6, h)
                .line_to(x4, y1)
                .line_to(x5, y1)
                .arc_to(w_r, h, ooxml_angle(st_ang), ooxml_angle(msw_ang))
                .finish(),
        },
    ]
}

struct GearTooth {
    b_a: f64,
    b_d: f64,
    x_a: f64,
    y_a: f64,
    x_b: f64,
    y_b: f64,
    x_c: f64,
    y_c: f64,
    x_d: f64,
    y_d: f64,
}

fn gear_angle(rw: f64, rh: f64, angle: f64) -> f64 {
    let rad = ooxml_angle(angle);
    f64::atan2(rh * rad.sin(), rw * rad.cos()).to_degrees() * 60000.0
}

fn gear_point(frame: Frame, rw: f64, rh: f64, b_ang: f64) -> (f64, f64) {
    let (sin, cos) = ooxml_angle(b_ang).sin_cos();
    let m = f64::hypot(rh * cos, rw * sin);
    let n = divide(rw * rh, m);
    (frame.w / 2.0 + n * cos, frame.h / 2.0 + n * sin)
}

fn gear_tooth(
    frame: Frame,
    rw: f64,
    rh: f64,
    th: f64,
    l_fd: f64,
    center_ang: f64,
    ha: f64,
) -> GearTooth {
    let b_a = gear_angle(rw, rh, center_ang - ha);
    let b_d = gear_angle(rw, rh, center_ang + ha);
    let (x_a, y_a) = gear_point(frame, rw, rh, b_a);
    let (x_d, y_d) = gear_point(frame, rw, rh, b_d);
    let a = f64::atan2(x_a - x_d, y_a - y_d);
    let (sin_a, cos_a) = a.sin_cos();
    let (dx_f, dy_f) = (l_fd * sin_a, l_fd * cos_a);
    let (x_f, y_f) = (x_d + dx_f, y_d + dy_f);
    let (x_e, y_e) = (x_a - dx_f, y_a - dy_f);
    let (y_ct, x_ct) = (th * sin_a, th * cos_a);
    GearTooth {
        b_a,
        b_d,
        x_a,
        y_a,
        x_b: x_e - x_ct,
        y_b: y_e + y_ct,
        x_c: x_f - x_ct,
        y_c: y_f + y_ct,
        x_d,
        y_d,
    }
}

fn gear6(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let th = adjustment(adjustments, "adj1", 0.15).clamp(0.0, 20000.0);
    let l_fd = adjustment(adjustments, "adj2", 0.03526).clamp(0.0, 5358.0);
    let l3 = th / 2.0 + l_fd / 2.0;
    let rh = h / 2.0 - th;
    let rw = w / 2.0 - th;
    let maxr = if rw - rh > 0.0 { rh } else { rw };
    let ha = f64::atan2(l3, maxr).to_degrees() * 60000.0;
    let t1 = gear_tooth(frame, rw, rh, th, l_fd, 19800000.0, ha);
    let b_d6 = gear_angle(rw, rh, 16200000.0 + ha);
    let (x_d6, y_d6) = gear_point(frame, rw, rh, b_d6);
    let x_a6 = w - x_d6;
    let x_f6 = x_d6 - l_fd;
    let x_e6 = x_a6 + l_fd;
    let y_c6 = y_d6 - th;
    let sw_ang1 = t1.b_a - b_d6;
    let b_a2 = gear_angle(rw, rh, 1800000.0 - ha);
    let b_d2 = gear_angle(rw, rh, 1800000.0 + ha);
    let y_a2 = h - t1.y_d;
    let y_d2 = h - t1.y_a;
    let y_c2 = h - t1.y_b;
    let y_b2 = h - t1.y_c;
    let sw_ang2 = b_a2 - t1.b_d;
    let b_d3 = gear_angle(rw, rh, 5400000.0 + ha);
    let y_d3 = h - y_d6;
    let y_b3 = h - y_c6;
    let b_d4 = gear_angle(rw, rh, 9000000.0 + ha);
    let x_d4 = w - t1.x_d;
    let x_c4 = w - t1.x_c;
    let x_b4 = w - t1.x_b;
    let b_d5 = gear_angle(rw, rh, 12600000.0 + ha);
    let x_d5 = w - t1.x_a;
    let x_c5 = w - t1.x_b;
    let x_b5 = w - t1.x_c;
    Outline::new(frame)
        .move_to(t1.x_a, t1.y_a)
        .line_to(t1.x_b, t1.y_b)
        .line_to(t1.x_c, t1.y_c)
        .line_to(t1.x_d, t1.y_d)
        .arc_to(rw, rh, ooxml_angle(t1.b_d), ooxml_angle(sw_ang2))
        .line_to(t1.x_c, y_b2)
        .line_to(t1.x_b, y_c2)
        .line_to(t1.x_a, y_d2)
        .arc_to(rw, rh, ooxml_angle(b_d2), ooxml_angle(sw_ang1))
        .line_to(x_f6, y_b3)
        .line_to(x_e6, y_b3)
        .line_to(x_a6, y_d3)
        .arc_to(rw, rh, ooxml_angle(b_d3), ooxml_angle(sw_ang1))
        .line_to(x_b4, y_c2)
        .line_to(x_c4, y_b2)
        .line_to(x_d4, y_a2)
        .arc_to(rw, rh, ooxml_angle(b_d4), ooxml_angle(sw_ang2))
        .line_to(x_b5, t1.y_c)
        .line_to(x_c5, t1.y_b)
        .line_to(x_d5, t1.y_a)
        .arc_to(rw, rh, ooxml_angle(b_d5), ooxml_angle(sw_ang1))
        .line_to(x_e6, y_c6)
        .line_to(x_f6, y_c6)
        .line_to(x_d6, y_d6)
        .arc_to(rw, rh, ooxml_angle(b_d6), ooxml_angle(sw_ang1))
        .close()
        .finish()
}

fn gear9(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let th = adjustment(adjustments, "adj1", 0.1).clamp(0.0, 20000.0);
    let l_fd = adjustment(adjustments, "adj2", 0.01763).clamp(0.0, 2679.0);
    let l3 = th / 2.0 + l_fd / 2.0;
    let rh = h / 2.0 - th;
    let rw = w / 2.0 - th;
    let maxr = if rw - rh > 0.0 { rh } else { rw };
    let ha = f64::atan2(l3, maxr).to_degrees() * 60000.0;
    let t1 = gear_tooth(frame, rw, rh, th, l_fd, 18600000.0, ha);
    let t2 = gear_tooth(frame, rw, rh, th, l_fd, 21000000.0, ha);
    let t3 = gear_tooth(frame, rw, rh, th, l_fd, 1800000.0, ha);
    let t4 = gear_tooth(frame, rw, rh, th, l_fd, 4200000.0, ha);
    let sw_ang1 = t2.b_a - t1.b_d;
    let sw_ang2 = t3.b_a - t2.b_d;
    let sw_ang3 = t4.b_a - t3.b_d;
    let b_a5 = gear_angle(rw, rh, 6600000.0 - ha);
    let b_d5 = gear_angle(rw, rh, 6600000.0 + ha);
    let sw_ang4 = b_a5 - t4.b_d;
    let b_d6 = gear_angle(rw, rh, 9000000.0 + ha);
    let b_d7 = gear_angle(rw, rh, 11400000.0 + ha);
    let b_d8 = gear_angle(rw, rh, 13800000.0 + ha);
    let b_a9 = gear_angle(rw, rh, 16200000.0 - ha);
    let b_d9 = gear_angle(rw, rh, 16200000.0 + ha);
    let (x_d9, y_d9) = gear_point(frame, rw, rh, b_d9);
    let x_a9 = w - x_d9;
    let x_f9 = x_d9 - l_fd;
    let x_e9 = x_a9 + l_fd;
    let y_c9 = y_d9 - th;
    let sw_ang5 = b_a9 - b_d8;
    Outline::new(frame)
        .move_to(t1.x_a, t1.y_a)
        .line_to(t1.x_b, t1.y_b)
        .line_to(t1.x_c, t1.y_c)
        .line_to(t1.x_d, t1.y_d)
        .arc_to(rw, rh, ooxml_angle(t1.b_d), ooxml_angle(sw_ang1))
        .line_to(t2.x_b, t2.y_b)
        .line_to(t2.x_c, t2.y_c)
        .line_to(t2.x_d, t2.y_d)
        .arc_to(rw, rh, ooxml_angle(t2.b_d), ooxml_angle(sw_ang2))
        .line_to(t3.x_b, t3.y_b)
        .line_to(t3.x_c, t3.y_c)
        .line_to(t3.x_d, t3.y_d)
        .arc_to(rw, rh, ooxml_angle(t3.b_d), ooxml_angle(sw_ang3))
        .line_to(t4.x_b, t4.y_b)
        .line_to(t4.x_c, t4.y_c)
        .line_to(t4.x_d, t4.y_d)
        .arc_to(rw, rh, ooxml_angle(t4.b_d), ooxml_angle(sw_ang4))
        .line_to(w - t4.x_c, t4.y_c)
        .line_to(w - t4.x_b, t4.y_b)
        .line_to(w - t4.x_a, t4.y_a)
        .arc_to(rw, rh, ooxml_angle(b_d5), ooxml_angle(sw_ang3))
        .line_to(w - t3.x_c, t3.y_c)
        .line_to(w - t3.x_b, t3.y_b)
        .line_to(w - t3.x_a, t3.y_a)
        .arc_to(rw, rh, ooxml_angle(b_d6), ooxml_angle(sw_ang2))
        .line_to(w - t2.x_c, t2.y_c)
        .line_to(w - t2.x_b, t2.y_b)
        .line_to(w - t2.x_a, t2.y_a)
        .arc_to(rw, rh, ooxml_angle(b_d7), ooxml_angle(sw_ang1))
        .line_to(w - t1.x_c, t1.y_c)
        .line_to(w - t1.x_b, t1.y_b)
        .line_to(w - t1.x_a, t1.y_a)
        .arc_to(rw, rh, ooxml_angle(b_d8), ooxml_angle(sw_ang5))
        .line_to(x_e9, y_c9)
        .line_to(x_f9, y_c9)
        .line_to(x_d9, y_d9)
        .arc_to(rw, rh, ooxml_angle(b_d9), ooxml_angle(sw_ang5))
        .close()
        .finish()
}

fn uturn_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.25);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let adj4 = adjustment(adjustments, "adj4", 0.4375);
    let adj5 = adjustment(adjustments, "adj5", 0.75);
    let a2 = adj2.clamp(0.0, 25000.0);
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let q2 = divide(a1 * 100000.0, h);
    let q3 = 100000.0 - q2;
    let max_adj3 = divide(q3 * h, 100000.0);
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q1 = a3 + a1;
    let min_adj5 = divide(q1 * 100000.0, h);
    let a5 = adj5.clamp(min_adj5, 100000.0_f64.max(min_adj5));
    let th = a1;
    let aw2 = a2;
    let th2 = th / 2.0;
    let dh2 = aw2 - th2;
    let y5 = divide(h * a5, 100000.0);
    let ah = a3;
    let y4 = y5 - ah;
    let x9 = w - dh2;
    let bw = x9 / 2.0;
    let bs = bw.min(y4);
    let max_adj4 = bs;
    let a4 = adj4.clamp(0.0, max_adj4.max(0.0));
    let bd = a4;
    let bd3 = bd - th;
    let bd2 = bd3.max(0.0);
    let x3 = th + bd2;
    let x8 = w - aw2;
    let x6 = x8 - aw2;
    let x7 = x6 + dh2;
    let x4 = x9 - bd;
    Outline::new(frame)
        .move_to(0.0, h)
        .line_to(0.0, bd)
        .arc_to(bd, bd, PI, FRAC_PI_2)
        .line_to(x4, 0.0)
        .arc_to(bd, bd, 3.0 * FRAC_PI_2, FRAC_PI_2)
        .line_to(x9, y4)
        .line_to(w, y4)
        .line_to(x8, y5)
        .line_to(x6, y4)
        .line_to(x7, y4)
        .line_to(x7, x3)
        .arc_to(bd2, bd2, 0.0, -FRAC_PI_2)
        .line_to(x3, th)
        .arc_to(bd2, bd2, 3.0 * FRAC_PI_2, -FRAC_PI_2)
        .line_to(th, h)
        .close()
        .finish()
}

fn moon(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let adj = adjustment(adjustments, "adj", 0.5);
    let g0 = adj.clamp(0.0, 87500.0);
    let g0w = divide(g0 * w, 100000.0);
    let g1 = 100000.0 - g0;
    let g2 = divide(g0 * g0, g1);
    let g3 = divide(100000.0 * 100000.0, g1);
    let g5 = g3 * 2.0 - g2;
    let g6 = g5 - g0;
    let g6w = divide(g6 * w, 100000.0);
    let g8 = g5 / 2.0 - g0;
    let dy1 = divide(g8 * (h / 2.0), 100000.0);
    let g18w = (g6w - g0w) / 2.0;
    let dx2 = -(g0w + g18w - w);
    let dy2 = -(h / 2.0);
    let st_ang1 = f64::atan2(dy2, dx2).to_degrees() * 60000.0;
    let en_angp1 = f64::atan2(h / 2.0, dx2).to_degrees() * 60000.0;
    let en_ang1 = en_angp1 - 21600000.0;
    let sw_ang1 = en_ang1 - st_ang1;
    Outline::new(frame)
        .move_to(w, h)
        .arc_to(w, h / 2.0, FRAC_PI_2, PI)
        .arc_to(g18w, dy1, ooxml_angle(st_ang1), ooxml_angle(sw_ang1))
        .close()
        .finish()
}

fn heart(frame: Frame) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let dx1 = w * 49.0 / 48.0;
    let dx2 = w * 10.0 / 48.0;
    let x1 = w / 2.0 - dx1;
    let x2 = w / 2.0 - dx2;
    let x3 = w / 2.0 + dx2;
    let x4 = w / 2.0 + dx1;
    let y1 = -h / 3.0;
    Outline::new(frame)
        .move_to(w / 2.0, h / 4.0)
        .cubic_to(x3, y1, x4, h / 4.0, w / 2.0, h)
        .cubic_to(x1, h / 4.0, x2, y1, w / 2.0, h / 4.0)
        .close()
        .finish()
}

fn vertical_scroll(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let ch = adjustment(adjustments, "adj", 0.125).clamp(0.0, 25000.0);
    let ch2 = ch / 2.0;
    let ch4 = ch / 4.0;
    let x3 = ch + ch2;
    let x4 = ch + ch;
    let x6 = w - ch;
    let x7 = w - ch2;
    let y3 = h - ch;
    let y4 = h - ch2;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(ch2, h)
                .arc_to(ch2, ch2, FRAC_PI_2, -FRAC_PI_2)
                .line_to(ch2, y4)
                .arc_to(ch4, ch4, FRAC_PI_2, -PI)
                .line_to(ch, y3)
                .line_to(ch, ch2)
                .arc_to(ch2, ch2, PI, FRAC_PI_2)
                .line_to(x7, 0.0)
                .arc_to(ch2, ch2, 3.0 * FRAC_PI_2, PI)
                .line_to(x6, ch)
                .line_to(x6, y4)
                .arc_to(ch2, ch2, 0.0, FRAC_PI_2)
                .close()
                .move_to(x4, ch2)
                .arc_to(ch2, ch2, 0.0, FRAC_PI_2)
                .arc_to(ch4, ch4, FRAC_PI_2, PI)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(x4, ch2)
                .arc_to(ch2, ch2, 0.0, FRAC_PI_2)
                .arc_to(ch4, ch4, FRAC_PI_2, PI)
                .close()
                .move_to(ch, y4)
                .arc_to(ch2, ch2, 0.0, 3.0 * FRAC_PI_2)
                .arc_to(ch4, ch4, 3.0 * FRAC_PI_2, PI)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(ch, y3)
                .line_to(ch, ch2)
                .arc_to(ch2, ch2, PI, FRAC_PI_2)
                .line_to(x7, 0.0)
                .arc_to(ch2, ch2, 3.0 * FRAC_PI_2, PI)
                .line_to(x6, ch)
                .line_to(x6, y4)
                .arc_to(ch2, ch2, 0.0, FRAC_PI_2)
                .line_to(ch2, h)
                .arc_to(ch2, ch2, FRAC_PI_2, PI)
                .close()
                .move_to(x3, 0.0)
                .arc_to(ch2, ch2, 3.0 * FRAC_PI_2, PI)
                .arc_to(ch4, ch4, FRAC_PI_2, PI)
                .line_to(x4, ch2)
                .move_to(x6, ch)
                .line_to(x3, ch)
                .move_to(ch2, y3)
                .arc_to(ch4, ch4, 3.0 * FRAC_PI_2, PI)
                .line_to(ch, y4)
                .move_to(ch2, h)
                .arc_to(ch2, ch2, FRAC_PI_2, -FRAC_PI_2)
                .line_to(ch, y3)
                .finish(),
        },
    ]
}

fn horizontal_scroll(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let ch = adjustment(adjustments, "adj", 0.125).clamp(0.0, 25000.0);
    let ch2 = ch / 2.0;
    let ch4 = ch / 4.0;
    let y3 = ch + ch2;
    let y4 = ch + ch;
    let y6 = h - ch;
    let y7 = h - ch2;
    let y5 = y6 - ch2;
    let x3 = w - ch;
    let x4 = w - ch2;
    vec![
        PresetGeometryPath {
            fill: PresetPathFill::Normal,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(w, ch2)
                .arc_to(ch2, ch2, 0.0, FRAC_PI_2)
                .line_to(x4, ch2)
                .arc_to(ch4, ch4, 0.0, PI)
                .line_to(x3, ch)
                .line_to(ch2, ch)
                .arc_to(ch2, ch2, 3.0 * FRAC_PI_2, -FRAC_PI_2)
                .line_to(0.0, y7)
                .arc_to(ch2, ch2, PI, -PI)
                .line_to(ch, y6)
                .line_to(x4, y6)
                .arc_to(ch2, ch2, FRAC_PI_2, -FRAC_PI_2)
                .close()
                .move_to(ch2, y4)
                .arc_to(ch2, ch2, FRAC_PI_2, -FRAC_PI_2)
                .arc_to(ch4, ch4, 0.0, -PI)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::DarkenLess,
            stroke: false,
            commands: Outline::new(frame)
                .move_to(ch2, y4)
                .arc_to(ch2, ch2, FRAC_PI_2, -FRAC_PI_2)
                .arc_to(ch4, ch4, 0.0, -PI)
                .close()
                .move_to(x4, ch)
                .arc_to(ch2, ch2, FRAC_PI_2, -3.0 * FRAC_PI_2)
                .arc_to(ch4, ch4, PI, -PI)
                .close()
                .finish(),
        },
        PresetGeometryPath {
            fill: PresetPathFill::None,
            stroke: true,
            commands: Outline::new(frame)
                .move_to(0.0, y3)
                .arc_to(ch2, ch2, PI, FRAC_PI_2)
                .line_to(x3, ch)
                .line_to(x3, ch2)
                .arc_to(ch2, ch2, PI, PI)
                .line_to(w, y5)
                .arc_to(ch2, ch2, 0.0, FRAC_PI_2)
                .line_to(ch, y6)
                .line_to(ch, y7)
                .arc_to(ch2, ch2, 0.0, PI)
                .close()
                .move_to(x3, ch)
                .line_to(x4, ch)
                .arc_to(ch2, ch2, FRAC_PI_2, -FRAC_PI_2)
                .move_to(x4, ch)
                .line_to(x4, ch2)
                .arc_to(ch4, ch4, 0.0, PI)
                .move_to(ch2, y4)
                .line_to(ch2, y3)
                .arc_to(ch4, ch4, PI, PI)
                .arc_to(ch2, ch2, 0.0, PI)
                .move_to(ch, y3)
                .line_to(ch, y6)
                .finish(),
        },
    ]
}

fn wave(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.125).clamp(0.0, 20000.0);
    let a2 = adjustment(adjustments, "adj2", 0.0).clamp(-10000.0, 10000.0);
    let y1 = divide(h * a1, 100000.0);
    let dy2 = y1 * 10.0 / 3.0;
    let y2 = y1 - dy2;
    let y3 = y1 + dy2;
    let y4 = h - y1;
    let y5 = y4 - dy2;
    let y6 = y4 + dy2;
    let of2 = divide(w * a2, 50000.0);
    let dx2 = if of2 > 0.0 { 0.0 } else { of2 };
    let x2 = -dx2;
    let dx5 = if of2 > 0.0 { of2 } else { 0.0 };
    let x5 = w - dx5;
    let dx3 = (dx2 + x5) / 3.0;
    let x3 = x2 + dx3;
    let x4 = (x3 + x5) / 2.0;
    let x6 = dx5;
    let x10 = w + dx2;
    let x7 = x6 + dx3;
    let x8 = (x7 + x10) / 2.0;
    Outline::new(frame)
        .move_to(x2, y1)
        .cubic_to(x3, y2, x4, y3, x5, y1)
        .line_to(x10, y4)
        .cubic_to(x8, y6, x7, y5, x6, y4)
        .close()
        .finish()
}

fn layer(
    commands: Vec<GeometryPathCommand>,
    fill: PresetPathFill,
    stroke: bool,
) -> PresetGeometryPath {
    PresetGeometryPath {
        commands,
        fill,
        stroke,
    }
}

fn chart_marks(_frame: Frame, segments: &[(f64, f64, f64, f64)]) -> Vec<PresetGeometryPath> {
    let small = Frame { w: 10.0, h: 10.0 };
    let mut marks = Outline::new(small);
    for &(x1, y1, x2, y2) in segments {
        marks = marks.move_to(x1, y1).line_to(x2, y2);
    }
    vec![
        layer(marks.finish(), PresetPathFill::None, true),
        layer(
            Outline::new(small)
                .move_to(0.0, 0.0)
                .line_to(0.0, 10.0)
                .line_to(10.0, 10.0)
                .line_to(10.0, 0.0)
                .close()
                .finish(),
            PresetPathFill::Normal,
            false,
        ),
    ]
}

fn tab_dx(frame: Frame) -> f64 {
    frame.w.hypot(frame.h) / 20.0
}

fn corner_tabs(frame: Frame) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let dx = tab_dx(frame);
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(dx, 0.0)
        .line_to(0.0, dx)
        .close()
        .move_to(0.0, h - dx)
        .line_to(dx, h)
        .line_to(0.0, h)
        .close()
        .move_to(w - dx, 0.0)
        .line_to(w, 0.0)
        .line_to(w, dx)
        .close()
        .move_to(w, h - dx)
        .line_to(w, h)
        .line_to(w - dx, h)
        .close()
        .finish()
}

fn square_tabs(frame: Frame) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let dx = tab_dx(frame);
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(dx, 0.0)
        .line_to(dx, dx)
        .line_to(0.0, dx)
        .close()
        .move_to(0.0, h - dx)
        .line_to(dx, h - dx)
        .line_to(dx, h)
        .line_to(0.0, h)
        .close()
        .move_to(w - dx, 0.0)
        .line_to(w, 0.0)
        .line_to(w, dx)
        .line_to(w - dx, dx)
        .close()
        .move_to(w - dx, h - dx)
        .line_to(w, h - dx)
        .line_to(w, h)
        .line_to(w - dx, h)
        .close()
        .finish()
}

fn plaque_tabs(frame: Frame) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let dx = tab_dx(frame);
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(dx, 0.0)
        .arc_to(dx, dx, 0.0, FRAC_PI_2)
        .close()
        .move_to(0.0, h - dx)
        .arc_to(dx, dx, 3.0 * FRAC_PI_2, FRAC_PI_2)
        .line_to(0.0, h)
        .close()
        .move_to(w, 0.0)
        .line_to(w, dx)
        .arc_to(dx, dx, FRAC_PI_2, FRAC_PI_2)
        .close()
        .move_to(w - dx, h)
        .arc_to(dx, dx, PI, FRAC_PI_2)
        .line_to(w, h)
        .close()
        .finish()
}

fn diag_stripe(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.5).clamp(0.0, 100000.0);
    let x2 = w * a / 100000.0;
    let y2 = h * a / 100000.0;
    Outline::new(frame)
        .move_to(0.0, y2)
        .line_to(x2, 0.0)
        .line_to(w, 0.0)
        .line_to(0.0, h)
        .close()
        .finish()
}

fn plaque(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.16667).clamp(0.0, 50000.0);
    let x1 = 100000.0 * a / 100000.0;
    let x2 = w - x1;
    let y2 = h - x1;
    Outline::new(frame)
        .move_to(0.0, x1)
        .arc_to(x1, x1, FRAC_PI_2, -FRAC_PI_2)
        .line_to(x2, 0.0)
        .arc_to(x1, x1, PI, -FRAC_PI_2)
        .line_to(w, y2)
        .arc_to(x1, x1, 3.0 * FRAC_PI_2, -FRAC_PI_2)
        .line_to(x1, h)
        .arc_to(x1, x1, 0.0, -FRAC_PI_2)
        .close()
        .finish()
}

fn snip_round_rect(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.16667).clamp(0.0, 50000.0);
    let a2 = adjustment(adjustments, "adj2", 0.16667).clamp(0.0, 50000.0);
    let x1 = a1;
    let dx2 = a2;
    let x2 = w - dx2;
    Outline::new(frame)
        .move_to(x1, 0.0)
        .line_to(x2, 0.0)
        .line_to(w, dx2)
        .line_to(w, h)
        .line_to(0.0, h)
        .line_to(0.0, x1)
        .arc_to(x1, x1, PI, FRAC_PI_2)
        .close()
        .finish()
}

fn half_frame(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let max_adj2 = 100000.0 * w / 100000.0;
    let a2 = adjustment(adjustments, "adj2", 0.33333).clamp(0.0, max_adj2.max(0.0));
    let x1 = a2;
    let g1 = h * x1 / w.max(f64::MIN_POSITIVE);
    let g2 = h - g1;
    let max_adj1 = 100000.0 * g2 / 100000.0;
    let a1 = adjustment(adjustments, "adj1", 0.33333).clamp(0.0, max_adj1.max(0.0));
    let y1 = a1;
    let dx2 = y1 * w / h.max(f64::MIN_POSITIVE);
    let x2 = w - dx2;
    let dy2 = x1 * h / w.max(f64::MIN_POSITIVE);
    let y2 = h - dy2;
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(w, 0.0)
        .line_to(x2, y1)
        .line_to(x1, y1)
        .line_to(x1, y2)
        .line_to(0.0, h)
        .close()
        .finish()
}

fn non_isosceles_trapezoid(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let w = frame.w;
    let max_adj = 50000.0 * w / 100000.0;
    let a1 = adjustment(adjustments, "adj1", 0.25).clamp(0.0, max_adj.max(0.0));
    let a2 = adjustment(adjustments, "adj2", 0.25).clamp(0.0, max_adj.max(0.0));
    let x2 = a1;
    let x3 = w - a2;
    Outline::new(frame)
        .move_to(0.0, frame.h)
        .line_to(x2, 0.0)
        .line_to(x3, 0.0)
        .line_to(w, frame.h)
        .close()
        .finish()
}

fn double_wave(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.0625).clamp(0.0, 12500.0);
    let a2 = adjustment(adjustments, "adj2", 0.0).clamp(-10000.0, 10000.0);
    let y1 = divide(h * a1, 100000.0);
    let dy2 = y1 * 10.0 / 3.0;
    let y2 = y1 - dy2;
    let y3 = y1 + dy2;
    let y4 = h - y1;
    let y5 = y4 - dy2;
    let y6 = y4 + dy2;
    let of2 = divide(w * a2, 50000.0);
    let dx2 = if of2 > 0.0 { 0.0 } else { of2 };
    let x2 = -dx2;
    let dx8 = if of2 > 0.0 { of2 } else { 0.0 };
    let x8 = w - dx8;
    let dx3 = (dx2 + x8) / 6.0;
    let x3 = x2 + dx3;
    let dx4 = (dx2 + x8) / 3.0;
    let x4 = x2 + dx4;
    let x5 = (x2 + x8) / 2.0;
    let x6 = x5 + dx3;
    let x7 = (x6 + x8) / 2.0;
    let x9 = dx8;
    let x15 = w + dx2;
    let x10 = x9 + dx3;
    let x11 = x9 + dx4;
    let x12 = (x9 + x15) / 2.0;
    let x13 = x12 + dx3;
    let x14 = (x13 + x15) / 2.0;
    Outline::new(frame)
        .move_to(x2, y1)
        .cubic_to(x3, y2, x4, y3, x5, y1)
        .cubic_to(x6, y2, x7, y3, x8, y1)
        .line_to(x15, y4)
        .cubic_to(x14, y6, x13, y5, x12, y4)
        .cubic_to(x11, y6, x10, y5, x9, y4)
        .close()
        .finish()
}

fn left_arrow_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let vc = frame.h / 2.0;
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.25);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let adj4 = adjustment(adjustments, "adj4", 0.64977);
    let max_adj2 = divide(50000.0 * frame.h, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = divide(100000.0 * frame.w, 100000.0);
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q2 = divide(a3 * 100000.0, frame.w);
    let max_adj4 = 100000.0 - q2;
    let a4 = adj4.clamp(0.0, max_adj4.max(0.0));
    let dy1 = a2;
    let dy2 = a1 / 2.0;
    let y1 = vc - dy1;
    let y2 = vc - dy2;
    let y3 = vc + dy2;
    let y4 = vc + dy1;
    let x1 = a3;
    let dx2 = divide(frame.w * a4, 100000.0);
    let x2 = frame.w - dx2;
    Outline::new(frame)
        .move_to(0.0, vc)
        .line_to(x1, y1)
        .line_to(x1, y2)
        .line_to(x2, y2)
        .line_to(x2, 0.0)
        .line_to(frame.w, 0.0)
        .line_to(frame.w, frame.h)
        .line_to(x2, frame.h)
        .line_to(x2, y3)
        .line_to(x1, y3)
        .line_to(x1, y4)
        .close()
        .finish()
}

fn right_arrow_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let vc = frame.h / 2.0;
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.25);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let adj4 = adjustment(adjustments, "adj4", 0.64977);
    let max_adj2 = divide(50000.0 * frame.h, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = divide(100000.0 * frame.w, 100000.0);
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q2 = divide(a3 * 100000.0, frame.w);
    let max_adj4 = 100000.0 - q2;
    let a4 = adj4.clamp(0.0, max_adj4.max(0.0));
    let dy1 = a2;
    let dy2 = a1 / 2.0;
    let y1 = vc - dy1;
    let y2 = vc - dy2;
    let y3 = vc + dy2;
    let y4 = vc + dy1;
    let dx3 = a3;
    let x3 = frame.w - dx3;
    let x2 = divide(frame.w * a4, 100000.0);
    Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(x2, 0.0)
        .line_to(x2, y2)
        .line_to(x3, y2)
        .line_to(x3, y1)
        .line_to(frame.w, vc)
        .line_to(x3, y4)
        .line_to(x3, y3)
        .line_to(x2, y3)
        .line_to(x2, frame.h)
        .line_to(0.0, frame.h)
        .close()
        .finish()
}

fn up_down_arrow_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.25);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let adj4 = adjustment(adjustments, "adj4", 0.48123);
    let max_adj2 = divide(50000.0 * frame.w, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = divide(50000.0 * frame.h, 100000.0);
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q2 = divide(a3 * 100000.0, vc);
    let max_adj4 = 100000.0 - q2;
    let a4 = adj4.clamp(0.0, max_adj4.max(0.0));
    let dx1 = a2;
    let dx2 = a1 / 2.0;
    let x1 = hc - dx1;
    let x2 = hc - dx2;
    let x3 = hc + dx2;
    let x4 = hc + dx1;
    let y1 = a3;
    let y4 = frame.h - y1;
    let dy2 = divide(frame.h * a4, 200000.0);
    let y2 = vc - dy2;
    let y3 = vc + dy2;
    Outline::new(frame)
        .move_to(0.0, y2)
        .line_to(x2, y2)
        .line_to(x2, y1)
        .line_to(x1, y1)
        .line_to(hc, 0.0)
        .line_to(x4, y1)
        .line_to(x3, y1)
        .line_to(x3, y2)
        .line_to(frame.w, y2)
        .line_to(frame.w, y3)
        .line_to(x3, y3)
        .line_to(x3, y4)
        .line_to(x4, y4)
        .line_to(hc, frame.h)
        .line_to(x1, y4)
        .line_to(x2, y4)
        .line_to(x2, y3)
        .line_to(0.0, y3)
        .close()
        .finish()
}

fn left_right_arrow_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let adj1 = adjustment(adjustments, "adj1", 0.25);
    let adj2 = adjustment(adjustments, "adj2", 0.25);
    let adj3 = adjustment(adjustments, "adj3", 0.25);
    let adj4 = adjustment(adjustments, "adj4", 0.48123);
    let max_adj2 = divide(50000.0 * frame.h, 100000.0);
    let a2 = adj2.clamp(0.0, max_adj2.max(0.0));
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = divide(50000.0 * frame.w, 100000.0);
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q2 = divide(a3 * 100000.0, hc);
    let max_adj4 = 100000.0 - q2;
    let a4 = adj4.clamp(0.0, max_adj4.max(0.0));
    let dy1 = a2;
    let dy2 = a1 / 2.0;
    let y1 = vc - dy1;
    let y2 = vc - dy2;
    let y3 = vc + dy2;
    let y4 = vc + dy1;
    let x1 = a3;
    let x4 = frame.w - x1;
    let dx2 = divide(frame.w * a4, 200000.0);
    let x2 = hc - dx2;
    let x3 = hc + dx2;
    Outline::new(frame)
        .move_to(0.0, vc)
        .line_to(x1, y1)
        .line_to(x1, y2)
        .line_to(x2, y2)
        .line_to(x2, 0.0)
        .line_to(x3, 0.0)
        .line_to(x3, y2)
        .line_to(x4, y2)
        .line_to(x4, y1)
        .line_to(frame.w, vc)
        .line_to(x4, y4)
        .line_to(x4, y3)
        .line_to(x3, y3)
        .line_to(x3, frame.h)
        .line_to(x2, frame.h)
        .line_to(x2, y3)
        .line_to(x1, y3)
        .line_to(x1, y4)
        .close()
        .finish()
}

fn quad_arrow_callout(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let adj1 = adjustment(adjustments, "adj1", 0.18515);
    let adj2 = adjustment(adjustments, "adj2", 0.18515);
    let adj3 = adjustment(adjustments, "adj3", 0.18515);
    let adj4 = adjustment(adjustments, "adj4", 0.48123);
    let a2 = adj2.clamp(0.0, 50000.0);
    let max_adj1 = a2 * 2.0;
    let a1 = adj1.clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = 50000.0 - a2;
    let a3 = adj3.clamp(0.0, max_adj3.max(0.0));
    let q2 = a3 * 2.0;
    let max_adj4 = 100000.0 - q2;
    let a4 = adj4.clamp(a1, max_adj4.max(a1));
    let dx2 = a2;
    let dx3 = a1 / 2.0;
    let ah = a3;
    let dx1 = divide(frame.w * a4, 200000.0);
    let dy1 = divide(frame.h * a4, 200000.0);
    let x8 = frame.w - ah;
    let x2 = hc - dx1;
    let x7 = hc + dx1;
    let x3 = hc - dx2;
    let x6 = hc + dx2;
    let x4 = hc - dx3;
    let x5 = hc + dx3;
    let y8 = frame.h - ah;
    let y2 = vc - dy1;
    let y7 = vc + dy1;
    let y3 = vc - dx2;
    let y6 = vc + dx2;
    let y4 = vc - dx3;
    let y5 = vc + dx3;
    Outline::new(frame)
        .move_to(0.0, vc)
        .line_to(ah, y3)
        .line_to(ah, y4)
        .line_to(x2, y4)
        .line_to(x2, y2)
        .line_to(x4, y2)
        .line_to(x4, ah)
        .line_to(x3, ah)
        .line_to(hc, 0.0)
        .line_to(x6, ah)
        .line_to(x5, ah)
        .line_to(x5, y2)
        .line_to(x7, y2)
        .line_to(x7, y4)
        .line_to(x8, y4)
        .line_to(x8, y3)
        .line_to(frame.w, vc)
        .line_to(x8, y6)
        .line_to(x8, y5)
        .line_to(x7, y5)
        .line_to(x7, y7)
        .line_to(x5, y7)
        .line_to(x5, y8)
        .line_to(x6, y8)
        .line_to(hc, frame.h)
        .line_to(x3, y8)
        .line_to(x4, y8)
        .line_to(x4, y7)
        .line_to(x2, y7)
        .line_to(x2, y5)
        .line_to(ah, y5)
        .line_to(ah, y6)
        .close()
        .finish()
}

fn action_button(frame: Frame, kind: &str) -> Vec<PresetGeometryPath> {
    let border = Outline::new(frame)
        .move_to(0.0, 0.0)
        .line_to(frame.w, 0.0)
        .line_to(frame.w, frame.h)
        .line_to(0.0, frame.h)
        .close()
        .finish();
    if kind == "actionButtonBlank" {
        return normal(border);
    }
    let mut layers = match kind {
        "actionButtonBackPrevious" => action_button_back_previous(frame),
        "actionButtonBeginning" => action_button_beginning(frame),
        "actionButtonDocument" => action_button_document(frame),
        "actionButtonEnd" => action_button_end(frame),
        "actionButtonForwardNext" => action_button_forward_next(frame),
        "actionButtonHelp" => action_button_help(frame),
        "actionButtonHome" => action_button_home(frame),
        "actionButtonInformation" => action_button_information(frame),
        "actionButtonMovie" => action_button_movie(frame),
        "actionButtonReturn" => action_button_return(frame),
        "actionButtonSound" => action_button_sound(frame),
        _ => unreachable!(),
    };
    let mut fill = border.clone();
    fill.append(&mut layers[0].commands);
    layers[0].commands = fill;
    layers.push(PresetGeometryPath {
        commands: border,
        fill: PresetPathFill::None,
        stroke: true,
    });
    layers
}

fn action_button_back_previous(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let path0 = Outline::new(frame)
        .move_to(g11, vc)
        .line_to(g12, g9)
        .line_to(g12, g10)
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_beginning(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13, 8.0);
    let g15 = divide(g13, 4.0);
    let g16 = g11 + g14;
    let g17 = g11 + g15;
    let path0 = Outline::new(frame)
        .move_to(g17, vc)
        .line_to(g12, g9)
        .line_to(g12, g10)
        .close()
        .move_to(g16, g9)
        .line_to(g11, g9)
        .line_to(g11, g10)
        .line_to(g16, g10)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(g17, vc)
        .line_to(g12, g9)
        .line_to(g12, g10)
        .close()
        .move_to(g16, g9)
        .line_to(g16, g10)
        .line_to(g11, g10)
        .line_to(g11, g9)
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_document(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let dx1 = divide(100000.0 * 9.0, 32.0);
    let g11 = hc - dx1;
    let g12 = hc + dx1;
    let g13 = divide(100000.0 * 3.0, 16.0);
    let g14 = g12 - g13;
    let g15 = g9 + g13;
    let path0 = Outline::new(frame)
        .move_to(g11, g9)
        .line_to(g14, g9)
        .line_to(g12, g15)
        .line_to(g12, g10)
        .line_to(g11, g10)
        .close()
        .finish();
    let path1 = Outline::new(frame)
        .move_to(g11, g9)
        .line_to(g14, g9)
        .line_to(g14, g15)
        .line_to(g12, g15)
        .line_to(g12, g10)
        .line_to(g11, g10)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(g14, g9)
        .line_to(g14, g15)
        .line_to(g12, g15)
        .close()
        .finish();
    let path3 = Outline::new(frame)
        .move_to(g11, g9)
        .line_to(g14, g9)
        .line_to(g12, g15)
        .line_to(g12, g10)
        .line_to(g11, g10)
        .close()
        .move_to(g12, g15)
        .line_to(g14, g15)
        .line_to(g14, g9)
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path1,
            fill: PresetPathFill::DarkenLess,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path3,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_end(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13 * 3.0, 4.0);
    let g15 = divide(g13 * 7.0, 8.0);
    let g16 = g11 + g14;
    let g17 = g11 + g15;
    let path0 = Outline::new(frame)
        .move_to(g16, vc)
        .line_to(g11, g9)
        .line_to(g11, g10)
        .close()
        .move_to(g17, g9)
        .line_to(g12, g9)
        .line_to(g12, g10)
        .line_to(g17, g10)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(g16, vc)
        .line_to(g11, g10)
        .line_to(g11, g9)
        .close()
        .move_to(g17, g9)
        .line_to(g12, g9)
        .line_to(g12, g10)
        .line_to(g17, g10)
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_forward_next(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let path0 = Outline::new(frame)
        .move_to(g12, vc)
        .line_to(g11, g9)
        .line_to(g11, g10)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(g12, vc)
        .line_to(g11, g10)
        .line_to(g11, g9)
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_help(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g11 = hc - dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13, 7.0);
    let g15 = divide(g13 * 3.0, 14.0);
    let g16 = divide(g13 * 2.0, 7.0);
    let g19 = divide(g13 * 3.0, 7.0);
    let g20 = divide(g13 * 4.0, 7.0);
    let g21 = divide(g13 * 17.0, 28.0);
    let g23 = divide(g13 * 21.0, 28.0);
    let g24 = divide(g13 * 11.0, 14.0);
    let g27 = g9 + g16;
    let g29 = g9 + g21;
    let g30 = g9 + g23;
    let g31 = g9 + g24;
    let g33 = g11 + g15;
    let g36 = g11 + g19;
    let g37 = g11 + g20;
    let g41 = divide(g13, 14.0);
    let g42 = divide(g13 * 3.0, 28.0);
    let path0 = Outline::new(frame)
        .move_to(g33, g27)
        .arc_to(g16, g16, ooxml_angle(10800000.0), ooxml_angle(10800000.0))
        .arc_to(g14, g15, ooxml_angle(0.0), ooxml_angle(5400000.0))
        .arc_to(g41, g42, ooxml_angle(16200000.0), ooxml_angle(-5400000.0))
        .line_to(g37, g30)
        .line_to(g36, g30)
        .line_to(g36, g29)
        .arc_to(g14, g15, ooxml_angle(10800000.0), ooxml_angle(5400000.0))
        .arc_to(g41, g42, ooxml_angle(5400000.0), ooxml_angle(-5400000.0))
        .arc_to(g14, g14, ooxml_angle(0.0), ooxml_angle(-10800000.0))
        .close()
        .move_to(hc, g31)
        .arc_to(g42, g42, ooxml_angle(16200000.0), ooxml_angle(21600000.0))
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_home(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13, 16.0);
    let g15 = divide(g13, 8.0);
    let g16 = divide(g13 * 3.0, 16.0);
    let g17 = divide(g13 * 5.0, 16.0);
    let g18 = divide(g13 * 7.0, 16.0);
    let g19 = divide(g13 * 9.0, 16.0);
    let g20 = divide(g13 * 11.0, 16.0);
    let g21 = divide(g13 * 3.0, 4.0);
    let g22 = divide(g13 * 13.0, 16.0);
    let g23 = divide(g13 * 7.0, 8.0);
    let g24 = g9 + g14;
    let g25 = g9 + g16;
    let g26 = g9 + g17;
    let g27 = g9 + g21;
    let g28 = g11 + g15;
    let g29 = g11 + g18;
    let g30 = g11 + g19;
    let g31 = g11 + g20;
    let g32 = g11 + g22;
    let g33 = g11 + g23;
    let path0 = Outline::new(frame)
        .move_to(hc, g9)
        .line_to(g11, vc)
        .line_to(g28, vc)
        .line_to(g28, g10)
        .line_to(g33, g10)
        .line_to(g33, vc)
        .line_to(g12, vc)
        .line_to(g32, g26)
        .line_to(g32, g24)
        .line_to(g31, g24)
        .line_to(g31, g25)
        .close()
        .finish();
    let path1 = Outline::new(frame)
        .move_to(g32, g26)
        .line_to(g32, g24)
        .line_to(g31, g24)
        .line_to(g31, g25)
        .close()
        .move_to(g28, vc)
        .line_to(g28, g10)
        .line_to(g29, g10)
        .line_to(g29, g27)
        .line_to(g30, g27)
        .line_to(g30, g10)
        .line_to(g33, g10)
        .line_to(g33, vc)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(hc, g9)
        .line_to(g11, vc)
        .line_to(g12, vc)
        .close()
        .move_to(g29, g27)
        .line_to(g30, g27)
        .line_to(g30, g10)
        .line_to(g29, g10)
        .close()
        .finish();
    let path3 = Outline::new(frame)
        .move_to(hc, g9)
        .line_to(g31, g25)
        .line_to(g31, g24)
        .line_to(g32, g24)
        .line_to(g32, g26)
        .line_to(g12, vc)
        .line_to(g33, vc)
        .line_to(g33, g10)
        .line_to(g28, g10)
        .line_to(g28, vc)
        .line_to(g11, vc)
        .close()
        .move_to(g31, g25)
        .line_to(g32, g26)
        .move_to(g33, vc)
        .line_to(g28, vc)
        .move_to(g29, g10)
        .line_to(g29, g27)
        .line_to(g30, g27)
        .line_to(g30, g10)
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path1,
            fill: PresetPathFill::DarkenLess,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path3,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_information(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g11 = hc - dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13, 32.0);
    let g17 = divide(g13 * 5.0, 16.0);
    let g18 = divide(g13 * 3.0, 8.0);
    let g19 = divide(g13 * 13.0, 32.0);
    let g20 = divide(g13 * 19.0, 32.0);
    let g22 = divide(g13 * 11.0, 16.0);
    let g23 = divide(g13 * 13.0, 16.0);
    let g24 = divide(g13 * 7.0, 8.0);
    let g25 = g9 + g14;
    let g28 = g9 + g17;
    let g29 = g9 + g18;
    let g30 = g9 + g23;
    let g31 = g9 + g24;
    let g32 = g11 + g17;
    let g34 = g11 + g19;
    let g35 = g11 + g20;
    let g37 = g11 + g22;
    let g38 = divide(g13 * 3.0, 32.0);
    let path0 = Outline::new(frame)
        .move_to(hc, g9)
        .arc_to(dx2, dx2, ooxml_angle(16200000.0), ooxml_angle(21600000.0))
        .close()
        .finish();
    let path1 = Outline::new(frame)
        .move_to(hc, g9)
        .arc_to(dx2, dx2, ooxml_angle(16200000.0), ooxml_angle(21600000.0))
        .close()
        .move_to(hc, g25)
        .arc_to(g38, g38, ooxml_angle(16200000.0), ooxml_angle(21600000.0))
        .move_to(g32, g28)
        .line_to(g32, g29)
        .line_to(g34, g29)
        .line_to(g34, g30)
        .line_to(g32, g30)
        .line_to(g32, g31)
        .line_to(g37, g31)
        .line_to(g37, g30)
        .line_to(g35, g30)
        .line_to(g35, g28)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(hc, g25)
        .arc_to(g38, g38, ooxml_angle(16200000.0), ooxml_angle(21600000.0))
        .move_to(g32, g28)
        .line_to(g35, g28)
        .line_to(g35, g30)
        .line_to(g37, g30)
        .line_to(g37, g31)
        .line_to(g32, g31)
        .line_to(g32, g30)
        .line_to(g34, g30)
        .line_to(g34, g29)
        .line_to(g32, g29)
        .close()
        .finish();
    let path3 = Outline::new(frame)
        .move_to(hc, g9)
        .arc_to(dx2, dx2, ooxml_angle(16200000.0), ooxml_angle(21600000.0))
        .close()
        .move_to(hc, g25)
        .arc_to(g38, g38, ooxml_angle(16200000.0), ooxml_angle(21600000.0))
        .move_to(g32, g28)
        .line_to(g35, g28)
        .line_to(g35, g30)
        .line_to(g37, g30)
        .line_to(g37, g31)
        .line_to(g32, g31)
        .line_to(g32, g30)
        .line_to(g34, g30)
        .line_to(g34, g29)
        .line_to(g32, g29)
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path1,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::Lighten,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path3,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_movie(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13 * 1455.0, 21600.0);
    let g15 = divide(g13 * 1905.0, 21600.0);
    let g16 = divide(g13 * 2325.0, 21600.0);
    let g17 = divide(g13 * 16155.0, 21600.0);
    let g18 = divide(g13 * 17010.0, 21600.0);
    let g19 = divide(g13 * 19335.0, 21600.0);
    let g20 = divide(g13 * 19725.0, 21600.0);
    let g21 = divide(g13 * 20595.0, 21600.0);
    let g22 = divide(g13 * 5280.0, 21600.0);
    let g23 = divide(g13 * 5730.0, 21600.0);
    let g24 = divide(g13 * 6630.0, 21600.0);
    let g25 = divide(g13 * 7492.0, 21600.0);
    let g26 = divide(g13 * 9067.0, 21600.0);
    let g27 = divide(g13 * 9555.0, 21600.0);
    let g28 = divide(g13 * 13342.0, 21600.0);
    let g29 = divide(g13 * 14580.0, 21600.0);
    let g30 = divide(g13 * 15592.0, 21600.0);
    let g31 = g11 + g14;
    let g32 = g11 + g15;
    let g33 = g11 + g16;
    let g34 = g11 + g17;
    let g35 = g11 + g18;
    let g36 = g11 + g19;
    let g37 = g11 + g20;
    let g38 = g11 + g21;
    let g39 = g9 + g22;
    let g40 = g9 + g23;
    let g41 = g9 + g24;
    let g42 = g9 + g25;
    let g43 = g9 + g26;
    let g44 = g9 + g27;
    let g45 = g9 + g28;
    let g46 = g9 + g29;
    let g47 = g9 + g30;
    let path0 = Outline::new(frame)
        .move_to(g11, g39)
        .line_to(g11, g44)
        .line_to(g31, g44)
        .line_to(g32, g43)
        .line_to(g33, g43)
        .line_to(g33, g47)
        .line_to(g35, g47)
        .line_to(g35, g45)
        .line_to(g36, g45)
        .line_to(g38, g46)
        .line_to(g12, g46)
        .line_to(g12, g41)
        .line_to(g38, g41)
        .line_to(g37, g42)
        .line_to(g35, g42)
        .line_to(g35, g41)
        .line_to(g34, g40)
        .line_to(g32, g40)
        .line_to(g31, g39)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(g11, g39)
        .line_to(g31, g39)
        .line_to(g32, g40)
        .line_to(g34, g40)
        .line_to(g35, g41)
        .line_to(g35, g42)
        .line_to(g37, g42)
        .line_to(g38, g41)
        .line_to(g12, g41)
        .line_to(g12, g46)
        .line_to(g38, g46)
        .line_to(g36, g45)
        .line_to(g35, g45)
        .line_to(g35, g47)
        .line_to(g33, g47)
        .line_to(g33, g43)
        .line_to(g32, g43)
        .line_to(g31, g44)
        .line_to(g11, g44)
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_return(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13 * 7.0, 8.0);
    let g15 = divide(g13 * 3.0, 4.0);
    let g16 = divide(g13 * 5.0, 8.0);
    let g17 = divide(g13 * 3.0, 8.0);
    let g18 = divide(g13, 4.0);
    let g19 = g9 + g15;
    let g20 = g9 + g16;
    let g21 = g9 + g18;
    let g22 = g11 + g14;
    let g23 = g11 + g15;
    let g24 = g11 + g16;
    let g25 = g11 + g17;
    let g26 = g11 + g18;
    let g27 = divide(g13, 8.0);
    let path0 = Outline::new(frame)
        .move_to(g12, g21)
        .line_to(g23, g9)
        .line_to(hc, g21)
        .line_to(g24, g21)
        .line_to(g24, g20)
        .arc_to(g27, g27, ooxml_angle(0.0), ooxml_angle(5400000.0))
        .line_to(g25, g19)
        .arc_to(g27, g27, ooxml_angle(5400000.0), ooxml_angle(5400000.0))
        .line_to(g26, g21)
        .line_to(g11, g21)
        .line_to(g11, g20)
        .arc_to(g17, g17, ooxml_angle(10800000.0), ooxml_angle(-5400000.0))
        .line_to(hc, g10)
        .arc_to(g17, g17, ooxml_angle(5400000.0), ooxml_angle(-5400000.0))
        .line_to(g22, g21)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(g12, g21)
        .line_to(g22, g21)
        .line_to(g22, g20)
        .arc_to(g17, g17, ooxml_angle(0.0), ooxml_angle(5400000.0))
        .line_to(g25, g10)
        .arc_to(g17, g17, ooxml_angle(5400000.0), ooxml_angle(5400000.0))
        .line_to(g11, g21)
        .line_to(g26, g21)
        .line_to(g26, g20)
        .arc_to(g27, g27, ooxml_angle(10800000.0), ooxml_angle(-5400000.0))
        .line_to(hc, g19)
        .arc_to(g27, g27, ooxml_angle(5400000.0), ooxml_angle(-5400000.0))
        .line_to(g24, g21)
        .line_to(hc, g21)
        .line_to(g23, g9)
        .close()
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn action_button_sound(frame: Frame) -> Vec<PresetGeometryPath> {
    let hc = frame.w / 2.0;
    let vc = frame.h / 2.0;
    let dx2 = divide(100000.0 * 3.0, 8.0);
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let g13 = divide(100000.0 * 3.0, 4.0);
    let g14 = divide(g13, 8.0);
    let g15 = divide(g13 * 5.0, 16.0);
    let g16 = divide(g13 * 5.0, 8.0);
    let g17 = divide(g13 * 11.0, 16.0);
    let g18 = divide(g13 * 3.0, 4.0);
    let g19 = divide(g13 * 7.0, 8.0);
    let g20 = g9 + g14;
    let g21 = g9 + g15;
    let g22 = g9 + g17;
    let g23 = g9 + g19;
    let g24 = g11 + g15;
    let g25 = g11 + g16;
    let g26 = g11 + g18;
    let path0 = Outline::new(frame)
        .move_to(g11, g21)
        .line_to(g11, g22)
        .line_to(g24, g22)
        .line_to(g25, g10)
        .line_to(g25, g9)
        .line_to(g24, g21)
        .close()
        .finish();
    let path2 = Outline::new(frame)
        .move_to(g11, g21)
        .line_to(g24, g21)
        .line_to(g25, g9)
        .line_to(g25, g10)
        .line_to(g24, g22)
        .line_to(g11, g22)
        .close()
        .move_to(g26, g21)
        .line_to(g12, g20)
        .move_to(g26, vc)
        .line_to(g12, vc)
        .move_to(g26, g22)
        .line_to(g12, g23)
        .finish();
    vec![
        PresetGeometryPath {
            commands: path0.clone(),
            fill: PresetPathFill::Normal,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path0,
            fill: PresetPathFill::Darken,
            stroke: false,
        },
        PresetGeometryPath {
            commands: path2,
            fill: PresetPathFill::None,
            stroke: true,
        },
    ]
}

fn flow_chart_multidocument() -> Vec<PresetGeometryPath> {
    let f = Frame {
        w: 21600.0,
        h: 21600.0,
    };
    vec![
        layer(
            Outline::new(f)
                .move_to(0.0, 20782.0)
                .cubic_to(9298.0, 23542.0, 9298.0, 18022.0, 18595.0, 18022.0)
                .line_to(18595.0, 3675.0)
                .line_to(0.0, 3675.0)
                .close()
                .move_to(1532.0, 3675.0)
                .line_to(1532.0, 1815.0)
                .line_to(20000.0, 1815.0)
                .line_to(20000.0, 16252.0)
                .cubic_to(19298.0, 16252.0, 18595.0, 16352.0, 18595.0, 16352.0)
                .line_to(18595.0, 3675.0)
                .close()
                .move_to(2972.0, 1815.0)
                .line_to(2972.0, 0.0)
                .line_to(21600.0, 0.0)
                .line_to(21600.0, 14392.0)
                .cubic_to(20800.0, 14392.0, 20000.0, 14467.0, 20000.0, 14467.0)
                .line_to(20000.0, 1815.0)
                .close()
                .finish(),
            PresetPathFill::Normal,
            false,
        ),
        layer(
            Outline::new(f)
                .move_to(0.0, 3675.0)
                .line_to(18595.0, 3675.0)
                .line_to(18595.0, 18022.0)
                .cubic_to(9298.0, 18022.0, 9298.0, 23542.0, 0.0, 20782.0)
                .close()
                .move_to(1532.0, 3675.0)
                .line_to(1532.0, 1815.0)
                .line_to(20000.0, 1815.0)
                .line_to(20000.0, 16252.0)
                .cubic_to(19298.0, 16252.0, 18595.0, 16352.0, 18595.0, 16352.0)
                .move_to(2972.0, 1815.0)
                .line_to(2972.0, 0.0)
                .line_to(21600.0, 0.0)
                .line_to(21600.0, 14392.0)
                .cubic_to(20800.0, 14392.0, 20000.0, 14467.0, 20000.0, 14467.0)
                .finish(),
            PresetPathFill::None,
            true,
        ),
        layer(
            Outline::new(f)
                .move_to(0.0, 20782.0)
                .cubic_to(9298.0, 23542.0, 9298.0, 18022.0, 18595.0, 18022.0)
                .line_to(18595.0, 16352.0)
                .cubic_to(18595.0, 16352.0, 19298.0, 16252.0, 20000.0, 16252.0)
                .line_to(20000.0, 14467.0)
                .cubic_to(20000.0, 14467.0, 20800.0, 14392.0, 21600.0, 14392.0)
                .line_to(21600.0, 0.0)
                .line_to(2972.0, 0.0)
                .line_to(2972.0, 1815.0)
                .line_to(1532.0, 1815.0)
                .line_to(1532.0, 3675.0)
                .line_to(0.0, 3675.0)
                .close()
                .finish(),
            PresetPathFill::None,
            false,
        ),
    ]
}

fn flow_chart_offline_storage() -> Vec<PresetGeometryPath> {
    let tri = || {
        Outline::new(Frame { w: 2.0, h: 2.0 })
            .move_to(0.0, 0.0)
            .line_to(2.0, 0.0)
            .line_to(1.0, 2.0)
            .close()
            .finish()
    };
    vec![
        layer(tri(), PresetPathFill::Normal, false),
        layer(
            Outline::new(Frame { w: 5.0, h: 5.0 })
                .move_to(2.0, 4.0)
                .line_to(3.0, 4.0)
                .finish(),
            PresetPathFill::None,
            true,
        ),
        layer(tri(), PresetPathFill::None, true),
    ]
}

fn flow_chart_or(frame: Frame, summing: bool) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let ellipse = || {
        Outline::new(frame)
            .move_to(0.0, h / 2.0)
            .arc_to(w / 2.0, h / 2.0, PI, FRAC_PI_2)
            .arc_to(w / 2.0, h / 2.0, 3.0 * FRAC_PI_2, FRAC_PI_2)
            .arc_to(w / 2.0, h / 2.0, 0.0, FRAC_PI_2)
            .arc_to(w / 2.0, h / 2.0, FRAC_PI_2, FRAC_PI_2)
            .close()
            .finish()
    };
    let cross = if summing {
        let c = std::f64::consts::FRAC_1_SQRT_2;
        let (il, ir) = (w / 2.0 - (w / 2.0) * c, w / 2.0 + (w / 2.0) * c);
        let (it, ib) = (h / 2.0 - (h / 2.0) * c, h / 2.0 + (h / 2.0) * c);
        Outline::new(frame)
            .move_to(il, it)
            .line_to(ir, ib)
            .move_to(ir, it)
            .line_to(il, ib)
            .finish()
    } else {
        Outline::new(frame)
            .move_to(w / 2.0, 0.0)
            .line_to(w / 2.0, h)
            .move_to(0.0, h / 2.0)
            .line_to(w, h / 2.0)
            .finish()
    };
    vec![
        layer(ellipse(), PresetPathFill::Normal, false),
        layer(cross, PresetPathFill::None, true),
        layer(ellipse(), PresetPathFill::None, true),
    ]
}

fn flow_chart_punched_tape() -> Vec<GeometryPathCommand> {
    Outline::new(Frame { w: 20.0, h: 20.0 })
        .move_to(0.0, 2.0)
        .arc_to(5.0, 2.0, PI, -PI)
        .arc_to(5.0, 2.0, PI, PI)
        .line_to(20.0, 18.0)
        .arc_to(5.0, 2.0, 0.0, -PI)
        .arc_to(5.0, 2.0, 0.0, PI)
        .close()
        .finish()
}

fn flow_chart_sort() -> Vec<PresetGeometryPath> {
    let diamond = || {
        Outline::new(Frame { w: 2.0, h: 2.0 })
            .move_to(0.0, 1.0)
            .line_to(1.0, 0.0)
            .line_to(2.0, 1.0)
            .line_to(1.0, 2.0)
            .close()
            .finish()
    };
    vec![
        layer(diamond(), PresetPathFill::Normal, false),
        layer(
            Outline::new(Frame { w: 2.0, h: 2.0 })
                .move_to(0.0, 1.0)
                .line_to(2.0, 1.0)
                .finish(),
            PresetPathFill::None,
            true,
        ),
        layer(diamond(), PresetPathFill::None, true),
    ]
}

fn funnel(frame: Frame) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let (wd2, hd4) = (w / 2.0, h / 4.0);
    let d = 100000.0 / 20.0;
    let rw2 = wd2 - d;
    let rh2 = hd4 - d;
    let ang = ooxml_angle(480000.0);
    let t1 = wd2 * ang.cos();
    let t2 = hd4 * ang.sin();
    let da = t2.atan2(t1);
    let st1 = PI - da;
    let (dx1, dy1) = elliptical_point(st1, wd2, hd4, frame);
    let (x1, y1) = (w / 2.0 + dx1, hd4 + dy1);
    let (rw3, rh3) = (wd2 / 4.0, hd4 / 4.0);
    let (dx3, dy3) = elliptical_point(da, rw3, rh3, frame);
    let (x3, y2) = (w / 2.0 + dx3, h - rh3 + dy3);
    let x2 = wd2 - rw2;
    Outline::new(frame)
        .move_to(x1, y1)
        .arc_to(wd2, hd4, st1, PI + 2.0 * da)
        .line_to(x3, y2)
        .arc_to(rw3, rh3, da, PI - 2.0 * da)
        .close()
        .move_to(x2, hd4)
        .arc_to(rw2, rh2, PI, -2.0 * PI)
        .close()
        .finish()
}

fn irregular_seal2() -> Vec<GeometryPathCommand> {
    let pts: [(f64, f64); 28] = [
        (11462.0, 4342.0),
        (14790.0, 0.0),
        (14525.0, 5777.0),
        (18007.0, 3172.0),
        (16380.0, 6532.0),
        (21600.0, 6645.0),
        (16985.0, 9402.0),
        (18270.0, 11290.0),
        (16380.0, 12310.0),
        (18877.0, 15632.0),
        (14640.0, 14350.0),
        (14942.0, 17370.0),
        (12180.0, 15935.0),
        (11612.0, 18842.0),
        (9872.0, 17370.0),
        (8700.0, 19712.0),
        (7527.0, 18125.0),
        (4917.0, 21600.0),
        (4805.0, 18240.0),
        (1285.0, 17825.0),
        (3330.0, 15370.0),
        (0.0, 12877.0),
        (3935.0, 11592.0),
        (1172.0, 8270.0),
        (5372.0, 7817.0),
        (4502.0, 3625.0),
        (8550.0, 6382.0),
        (9722.0, 1887.0),
    ];
    let mut outline = Outline::new(Frame {
        w: 21600.0,
        h: 21600.0,
    })
    .move_to(pts[0].0, pts[0].1);
    for &(x, y) in &pts[1..] {
        outline = outline.line_to(x, y);
    }
    outline.close().finish()
}

fn lightning_bolt() -> Vec<GeometryPathCommand> {
    let pts: [(f64, f64); 11] = [
        (8472.0, 0.0),
        (12860.0, 6080.0),
        (11050.0, 6797.0),
        (16577.0, 12007.0),
        (14767.0, 12877.0),
        (21600.0, 21600.0),
        (10012.0, 14915.0),
        (12222.0, 13987.0),
        (5022.0, 9705.0),
        (7602.0, 8382.0),
        (0.0, 3890.0),
    ];
    let mut outline = Outline::new(Frame {
        w: 21600.0,
        h: 21600.0,
    })
    .move_to(pts[0].0, pts[0].1);
    for &(x, y) in &pts[1..] {
        outline = outline.line_to(x, y);
    }
    outline.close().finish()
}

fn left_up_arrow(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a2 = adjustment(adjustments, "adj2", 0.25).clamp(0.0, 50000.0);
    let max_adj1 = a2 * 2.0;
    let a1 = adjustment(adjustments, "adj1", 0.25).clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = 100000.0 - max_adj1;
    let a3 = adjustment(adjustments, "adj3", 0.25).clamp(0.0, max_adj3.max(0.0));
    let x1 = a3;
    let dx2 = a2 * 2.0;
    let x2 = w - dx2;
    let y2 = h - dx2;
    let dx4 = a2;
    let x4 = w - dx4;
    let y4 = h - dx4;
    let dx3 = a1 / 2.0;
    let x3 = x4 - dx3;
    let x5 = x4 + dx3;
    let y3 = y4 - dx3;
    let y5 = y4 + dx3;
    Outline::new(frame)
        .move_to(0.0, y4)
        .line_to(x1, y2)
        .line_to(x1, y3)
        .line_to(x3, y3)
        .line_to(x3, x1)
        .line_to(x2, x1)
        .line_to(x4, 0.0)
        .line_to(w, x1)
        .line_to(x5, x1)
        .line_to(x5, y5)
        .line_to(x1, y5)
        .line_to(x1, h)
        .close()
        .finish()
}

fn left_right_up_arrow(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a2 = adjustment(adjustments, "adj2", 0.25).clamp(0.0, 50000.0);
    let max_adj1 = a2 * 2.0;
    let a1 = adjustment(adjustments, "adj1", 0.25).clamp(0.0, max_adj1.max(0.0));
    let max_adj3 = (100000.0 - max_adj1) / 2.0;
    let a3 = adjustment(adjustments, "adj3", 0.25).clamp(0.0, max_adj3.max(0.0));
    let x1 = a3;
    let dx2 = a2;
    let x2 = w / 2.0 - dx2;
    let x5 = w / 2.0 + dx2;
    let dx3 = a1 / 2.0;
    let x3 = w / 2.0 - dx3;
    let x4 = w / 2.0 + dx3;
    let x6 = w - x1;
    let dy2 = a2 * 2.0;
    let y2 = h - dy2;
    let y4 = h - dx2;
    let y3 = y4 - dx3;
    let y5 = y4 + dx3;
    Outline::new(frame)
        .move_to(0.0, y4)
        .line_to(x1, y2)
        .line_to(x1, y3)
        .line_to(x3, y3)
        .line_to(x3, x1)
        .line_to(x2, x1)
        .line_to(w / 2.0, 0.0)
        .line_to(x5, x1)
        .line_to(x4, x1)
        .line_to(x4, y3)
        .line_to(x6, y3)
        .line_to(x6, y2)
        .line_to(w, y4)
        .line_to(x6, h)
        .line_to(x6, y5)
        .line_to(x1, y5)
        .line_to(x1, h)
        .close()
        .finish()
}

fn math_minus(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.2352).clamp(0.0, 100000.0);
    let dy1 = h * a1 / 200000.0;
    let dx1 = w * 73490.0 / 200000.0;
    let (y1, y2) = (h / 2.0 - dy1, h / 2.0 + dy1);
    let (x1, x2) = (w / 2.0 - dx1, w / 2.0 + dx1);
    Outline::new(frame)
        .move_to(x1, y1)
        .line_to(x2, y1)
        .line_to(x2, y2)
        .line_to(x1, y2)
        .close()
        .finish()
}

fn math_divide(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.2352).clamp(1000.0, 36745.0);
    let max_adj3 = ((73490.0 - a1) / 4.0).min(36745.0 * w / h.max(f64::MIN_POSITIVE));
    let a3 = adjustment(adjustments, "adj3", 0.1176).clamp(1000.0, max_adj3.max(1000.0));
    let max_adj2 = 73490.0 - 4.0 * a3 - a1;
    let a2 = adjustment(adjustments, "adj2", 0.0588).clamp(0.0, max_adj2.max(0.0));
    let dy1 = h * a1 / 200000.0;
    let yg = h * a2 / 100000.0;
    let rad = h * a3 / 100000.0;
    let dx1 = w * 73490.0 / 200000.0;
    let y3 = h / 2.0 - dy1;
    let y4 = h / 2.0 + dy1;
    let y2 = y3 - (yg + rad);
    let y1 = y2 - rad;
    let y5 = h - y1;
    let (x1, x3) = (w / 2.0 - dx1, w / 2.0 + dx1);
    Outline::new(frame)
        .move_to(w / 2.0, y1)
        .arc_to(rad, rad, 3.0 * FRAC_PI_2, 2.0 * PI)
        .close()
        .move_to(w / 2.0, y5)
        .arc_to(rad, rad, FRAC_PI_2, 2.0 * PI)
        .close()
        .move_to(x1, y3)
        .line_to(x3, y3)
        .line_to(x3, y4)
        .line_to(x1, y4)
        .close()
        .finish()
}

fn math_not_equal(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.2352).clamp(0.0, 50000.0);
    let cr_ang = adjustment(adjustments, "adj2", 66.0).clamp(4200000.0, 6600000.0);
    let max_adj3 = 100000.0 - 2.0 * a1;
    let a3 = adjustment(adjustments, "adj3", 0.1176).clamp(0.0, max_adj3.max(0.0));
    let dy1 = h * a1 / 100000.0;
    let dy2 = h * a3 / 200000.0;
    let dx1 = w * 73490.0 / 200000.0;
    let (hc, vc, hd2) = (w / 2.0, h / 2.0, h / 2.0);
    let x1 = hc - dx1;
    let x8 = hc + dx1;
    let y2 = vc - dy2;
    let y3 = vc + dy2;
    let y1 = y2 - dy1;
    let y4 = y3 + dy1;
    let cadj2 = cr_ang - 5400000.0;
    let xadj2 = hd2 * ooxml_angle(cadj2).tan();
    let len = xadj2.hypot(hd2);
    let bhw = len * dy1 / hd2;
    let x7 = hc + xadj2 - bhw / 2.0;
    let at = |y: f64| x7 - xadj2 * y / hd2;
    let (x6, x5, x4, x3) = (at(y1), at(y2), at(y3), at(y4));
    let (rx6, rx5, rx4, rx3) = (x6 + bhw, x5 + bhw, x4 + bhw, x3 + bhw);
    let dx7 = dy1 * hd2 / len.max(f64::MIN_POSITIVE);
    let (rx, lx) = if cadj2 > 0.0 {
        (x7 + dx7, x7)
    } else {
        (x7 + bhw, x7 + bhw - dx7)
    };
    let dy3 = dy1 * xadj2 / len.max(f64::MIN_POSITIVE);
    let (ry, ly) = if cadj2 > 0.0 { (dy3, 0.0) } else { (0.0, -dy3) };
    let (dlx, drx, dly, dry) = (w - rx, w - lx, h - ry, h - ly);
    Outline::new(frame)
        .move_to(x1, y1)
        .line_to(x6, y1)
        .line_to(lx, ly)
        .line_to(rx, ry)
        .line_to(rx6, y1)
        .line_to(x8, y1)
        .line_to(x8, y2)
        .line_to(rx5, y2)
        .line_to(rx4, y3)
        .line_to(x8, y3)
        .line_to(x8, y4)
        .line_to(rx3, y4)
        .line_to(drx, dry)
        .line_to(dlx, dly)
        .line_to(x3, y4)
        .line_to(x1, y4)
        .line_to(x1, y3)
        .line_to(x4, y3)
        .line_to(x5, y2)
        .line_to(x1, y2)
        .close()
        .finish()
}

fn sun(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.25).clamp(12500.0, 46875.0);
    let g0 = 50000.0 - a;
    let g1 = g0 * 30274.0 / 32768.0;
    let g2 = g0 * 12540.0 / 32768.0;
    let g5 = 50000.0 - g1;
    let g6 = 50000.0 - g2;
    let g10 = g5 * 3.0 / 4.0;
    let g11 = g6 * 3.0 / 4.0;
    let g12 = g10 + 3662.0;
    let g13 = g11 + 3662.0;
    let g14 = g11 + 12500.0;
    let g15 = 100000.0 - g10;
    let g16 = 100000.0 - g12;
    let g17 = 100000.0 - g13;
    let g18 = 100000.0 - g14;
    let (ox1, oy1) = (w * 18436.0 / 21600.0, h * 3163.0 / 21600.0);
    let (ox2, oy2) = (w * 3163.0 / 21600.0, h * 18436.0 / 21600.0);
    let x = |g: f64| w * g / 100000.0;
    let y = |g: f64| h * g / 100000.0;
    Outline::new(frame)
        .move_to(w, h / 2.0)
        .line_to(x(g15), y(g18))
        .line_to(x(g15), y(g14))
        .close()
        .move_to(ox1, oy1)
        .line_to(x(g16), y(g13))
        .line_to(x(g17), y(g12))
        .close()
        .move_to(w / 2.0, 0.0)
        .line_to(x(g18), y(g10))
        .line_to(x(g14), y(g10))
        .close()
        .move_to(ox2, oy1)
        .line_to(x(g13), y(g12))
        .line_to(x(g12), y(g13))
        .close()
        .move_to(0.0, h / 2.0)
        .line_to(x(g10), y(g14))
        .line_to(x(g10), y(g18))
        .close()
        .move_to(ox2, oy2)
        .line_to(x(g12), y(g17))
        .line_to(x(g13), y(g16))
        .close()
        .move_to(w / 2.0, h)
        .line_to(x(g14), y(g15))
        .line_to(x(g18), y(g15))
        .close()
        .move_to(ox1, oy2)
        .line_to(x(g17), y(g16))
        .line_to(x(g16), y(g17))
        .close()
        .move_to(x(a), h / 2.0)
        .arc_to(x(g0), y(g0), PI, 2.0 * PI)
        .close()
        .finish()
}

fn smiley_face(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.04653).clamp(-4653.0, 4653.0);
    let x1 = w * 4969.0 / 21699.0;
    let x2 = w * 6215.0 / 21600.0;
    let x3 = w * 13135.0 / 21600.0;
    let x4 = w * 16640.0 / 21600.0;
    let y1 = h * 7570.0 / 21600.0;
    let y3 = h * 16515.0 / 21600.0;
    let dy2 = h * a / 100000.0;
    let y2 = y3 - dy2;
    let y4 = y3 + dy2;
    let y5 = y4 + h * a / 50000.0;
    let (wr, hr) = (w * 1125.0 / 21600.0, h * 1125.0 / 21600.0);
    let face = || {
        Outline::new(frame)
            .move_to(0.0, h / 2.0)
            .arc_to(w / 2.0, h / 2.0, PI, 2.0 * PI)
            .close()
            .finish()
    };
    vec![
        layer(face(), PresetPathFill::Normal, false),
        layer(
            Outline::new(frame)
                .move_to(x2, y1)
                .arc_to(wr, hr, PI, 2.0 * PI)
                .move_to(x3, y1)
                .arc_to(wr, hr, PI, 2.0 * PI)
                .finish(),
            PresetPathFill::DarkenLess,
            true,
        ),
        layer(
            Outline::new(frame)
                .move_to(x1, y2)
                .quad_to(w / 2.0, y5, x4, y2)
                .finish(),
            PresetPathFill::None,
            true,
        ),
        layer(face(), PresetPathFill::None, true),
    ]
}

fn bevel(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.125).clamp(0.0, 50000.0);
    let x1 = a;
    let x2 = w - x1;
    let y2 = h - x1;
    let rect = |x1: f64, y1: f64, x2: f64, y2: f64| {
        Outline::new(frame)
            .move_to(x1, y1)
            .line_to(x2, y1)
            .line_to(x2, y2)
            .line_to(x1, y2)
            .close()
            .finish()
    };
    vec![
        layer(rect(x1, x1, x2, y2), PresetPathFill::Normal, false),
        layer(
            Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(w, 0.0)
                .line_to(x2, x1)
                .line_to(x1, x1)
                .close()
                .finish(),
            PresetPathFill::LightenLess,
            false,
        ),
        layer(
            Outline::new(frame)
                .move_to(0.0, h)
                .line_to(x1, y2)
                .line_to(x2, y2)
                .line_to(w, h)
                .close()
                .finish(),
            PresetPathFill::DarkenLess,
            false,
        ),
        layer(
            Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(x1, x1)
                .line_to(x1, y2)
                .line_to(0.0, h)
                .close()
                .finish(),
            PresetPathFill::Lighten,
            false,
        ),
        layer(
            Outline::new(frame)
                .move_to(w, 0.0)
                .line_to(w, h)
                .line_to(x2, y2)
                .line_to(x2, x1)
                .close()
                .finish(),
            PresetPathFill::Darken,
            false,
        ),
        layer(
            Outline::new(frame)
                .move_to(0.0, 0.0)
                .line_to(w, 0.0)
                .line_to(w, h)
                .line_to(0.0, h)
                .close()
                .move_to(x1, x1)
                .line_to(x2, x1)
                .line_to(x2, y2)
                .line_to(x1, y2)
                .close()
                .move_to(0.0, 0.0)
                .line_to(x1, x1)
                .move_to(0.0, h)
                .line_to(x1, y2)
                .move_to(w, 0.0)
                .line_to(x2, x1)
                .move_to(w, h)
                .line_to(x2, y2)
                .finish(),
            PresetPathFill::None,
            true,
        ),
    ]
}

fn brace_pair(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.08333).clamp(0.0, 25000.0);
    let x1 = a;
    let x2 = a * 2.0;
    let x3 = w - x2;
    let x4 = w - x1;
    let y2 = h / 2.0 - x1;
    let y3 = h / 2.0 + x1;
    let y4 = h - x1;
    let body = |close: bool| {
        let outline = Outline::new(frame)
            .move_to(x2, h)
            .arc_to(x1, x1, FRAC_PI_2, FRAC_PI_2)
            .line_to(x1, y3)
            .arc_to(x1, x1, 0.0, -FRAC_PI_2)
            .arc_to(x1, x1, FRAC_PI_2, -FRAC_PI_2)
            .line_to(x1, x1)
            .arc_to(x1, x1, PI, FRAC_PI_2);
        let outline = if close {
            outline
                .line_to(x3, 0.0)
                .arc_to(x1, x1, 3.0 * FRAC_PI_2, FRAC_PI_2)
                .line_to(x4, y2)
                .arc_to(x1, x1, PI, -FRAC_PI_2)
                .arc_to(x1, x1, 3.0 * FRAC_PI_2, -FRAC_PI_2)
                .line_to(x4, y4)
                .arc_to(x1, x1, 0.0, FRAC_PI_2)
                .close()
        } else {
            outline
                .move_to(x3, 0.0)
                .arc_to(x1, x1, 3.0 * FRAC_PI_2, FRAC_PI_2)
                .line_to(x4, y2)
                .arc_to(x1, x1, PI, -FRAC_PI_2)
                .arc_to(x1, x1, 3.0 * FRAC_PI_2, -FRAC_PI_2)
                .line_to(x4, y4)
                .arc_to(x1, x1, 0.0, FRAC_PI_2)
        };
        outline.finish()
    };
    vec![
        layer(body(true), PresetPathFill::Normal, false),
        layer(body(false), PresetPathFill::None, true),
    ]
}

fn bracket_pair(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let a = adjustment(adjustments, "adj", 0.16667).clamp(0.0, 50000.0);
    let x1 = a;
    let x2 = w - x1;
    let y2 = h - x1;
    vec![
        layer(
            Outline::new(frame)
                .move_to(0.0, x1)
                .arc_to(x1, x1, PI, FRAC_PI_2)
                .line_to(x2, 0.0)
                .arc_to(x1, x1, 3.0 * FRAC_PI_2, FRAC_PI_2)
                .line_to(w, y2)
                .arc_to(x1, x1, 0.0, FRAC_PI_2)
                .line_to(x1, h)
                .arc_to(x1, x1, FRAC_PI_2, FRAC_PI_2)
                .close()
                .finish(),
            PresetPathFill::Normal,
            false,
        ),
        layer(
            Outline::new(frame)
                .move_to(x1, h)
                .arc_to(x1, x1, FRAC_PI_2, FRAC_PI_2)
                .line_to(0.0, x1)
                .arc_to(x1, x1, PI, FRAC_PI_2)
                .move_to(x2, 0.0)
                .arc_to(x1, x1, 3.0 * FRAC_PI_2, FRAC_PI_2)
                .line_to(w, y2)
                .arc_to(x1, x1, 0.0, FRAC_PI_2)
                .finish(),
            PresetPathFill::None,
            true,
        ),
    ]
}

fn ellipse_ribbon2(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let a1 = adjustment(adjustments, "adj1", 0.25).clamp(0.0, 100000.0);
    let a2 = adjustment(adjustments, "adj2", 0.5).clamp(25000.0, 75000.0);
    let min_adj3 = (a1 - (100000.0 - a1) / 2.0).max(0.0);
    let a3 = adjustment(adjustments, "adj3", 0.125).clamp(min_adj3, a1.max(min_adj3));
    let hc = w / 2.0;
    let dx2 = w * a2 / 200000.0;
    let x2 = hc - dx2;
    let x3 = x2 + w / 8.0;
    let x4 = w - x3;
    let x5 = w - x2;
    let x6 = w - w / 8.0;
    let dy1 = h * a3 / 100000.0;
    let f1 = 4.0 * dy1 / w.max(f64::MIN_POSITIVE);
    let u = |x: f64| f1 * (x - x * x / w.max(f64::MIN_POSITIVE));
    let y1 = h - u(x3);
    let cx1 = x3 / 2.0;
    let cy1 = h - f1 * cx1;
    let cx2 = w - cx1;
    let q1 = h * a1 / 100000.0;
    let dy3 = q1 - dy1;
    let q5v = u(x2);
    let u3 = q5v + dy3;
    let y3 = h - u3;
    let cu3 = dy1 + dy3 - u3 + dy1 + dy3;
    let cy3 = h - cu3;
    let rh = h - q1;
    let y2 = h - (dy1 * 14.0 / 16.0 + rh) / 2.0;
    let y5 = h - (q5v + rh);
    let y6 = h - (u3 + rh);
    let cx4 = x2 / 2.0;
    let cy4 = h - (f1 * cx4 + rh);
    let cx5 = w - cx4;
    let cy6 = h - (cu3 + rh);
    let u7 = u(x3) + dy3;
    let y7 = h - u7;
    let cy7 = h - (2.0 * q1 - u7);
    vec![
        layer(
            Outline::new(frame)
                .move_to(0.0, h)
                .quad_to(cx1, cy1, x3, y1)
                .line_to(x2, y3)
                .quad_to(hc, cy3, x5, y3)
                .line_to(x4, y1)
                .quad_to(cx2, cy1, w, h)
                .line_to(x6, y2)
                .line_to(w, q1)
                .quad_to(cx5, cy4, x5, y5)
                .line_to(x5, y6)
                .quad_to(hc, cy6, x2, y6)
                .line_to(x2, y5)
                .quad_to(cx4, cy4, 0.0, q1)
                .line_to(w / 8.0, y2)
                .close()
                .finish(),
            PresetPathFill::Normal,
            false,
        ),
        layer(
            Outline::new(frame)
                .move_to(x3, y7)
                .line_to(x3, y1)
                .line_to(x2, y3)
                .quad_to(hc, cy3, x5, y3)
                .line_to(x4, y1)
                .line_to(x4, y7)
                .quad_to(hc, cy7, x3, y7)
                .close()
                .finish(),
            PresetPathFill::DarkenLess,
            false,
        ),
        layer(
            Outline::new(frame)
                .move_to(0.0, h)
                .line_to(w / 8.0, y2)
                .line_to(0.0, q1)
                .quad_to(cx4, cy4, x2, y5)
                .line_to(x2, y6)
                .quad_to(hc, cy6, x5, y6)
                .line_to(x5, y5)
                .quad_to(cx5, cy4, w, q1)
                .line_to(x6, y2)
                .line_to(w, h)
                .quad_to(cx2, cy1, x4, y1)
                .line_to(x5, y3)
                .quad_to(hc, cy3, x2, y3)
                .line_to(x3, y1)
                .quad_to(cx1, cy1, 0.0, h)
                .close()
                .move_to(x2, y3)
                .line_to(x2, y5)
                .move_to(x5, y5)
                .line_to(x5, y3)
                .move_to(x3, y7)
                .line_to(x3, y1)
                .move_to(x4, y1)
                .line_to(x4, y7)
                .finish(),
            PresetPathFill::None,
            true,
        ),
    ]
}

fn left_right_ribbon(frame: Frame, adjustments: &HashMap<String, f64>) -> Vec<PresetGeometryPath> {
    let (w, h) = (frame.w, frame.h);
    let a3 = adjustment(adjustments, "adj3", 0.16667).clamp(0.0, 33333.0);
    let max_adj1 = 100000.0 - a3;
    let a1 = adjustment(adjustments, "adj1", 0.5).clamp(0.0, max_adj1.max(0.0));
    let wd32 = w / 32.0;
    let max_adj2 = w / 2.0 - wd32;
    let a2 = adjustment(adjustments, "adj2", 0.5).clamp(0.0, max_adj2.max(0.0));
    let x1 = a2;
    let x4 = w - x1;
    let dy1 = h * a1 / 200000.0;
    let dy2 = -h * a3 / 200000.0;
    let vc = h / 2.0;
    let ly1 = vc + dy2 - dy1;
    let ry4 = vc + dy1 - dy2;
    let ly2 = ly1 + dy1;
    let ry3 = h - ly2;
    let ly4 = ly2 * 2.0;
    let ry1 = h - ly4;
    let ly3 = ly4 - ly1;
    let ry2 = h - ly3;
    let hr = a3 / 4.0;
    let x2 = w / 2.0 - wd32;
    let x3 = w / 2.0 + wd32;
    let y1 = ly1 + hr;
    let y2 = ry2 - hr;
    let body = |close: bool| {
        let outline = Outline::new(frame)
            .move_to(0.0, ly2)
            .line_to(x1, 0.0)
            .line_to(x1, ly1)
            .line_to(w / 2.0, ly1)
            .arc_to(wd32, hr, 3.0 * FRAC_PI_2, PI)
            .arc_to(wd32, hr, 3.0 * FRAC_PI_2, -PI)
            .line_to(x4, ry2)
            .line_to(x4, ry1)
            .line_to(w, ry3)
            .line_to(x4, h)
            .line_to(x4, ry4)
            .line_to(w / 2.0, ry4)
            .arc_to(wd32, hr, FRAC_PI_2, FRAC_PI_2)
            .line_to(x2, ly3)
            .line_to(x1, ly3)
            .line_to(x1, ly4);
        let outline = if close { outline.close() } else { outline };
        outline.finish()
    };
    vec![
        layer(body(true), PresetPathFill::Normal, false),
        layer(
            Outline::new(frame)
                .move_to(x3, y1)
                .arc_to(wd32, hr, 0.0, FRAC_PI_2)
                .arc_to(wd32, hr, 3.0 * FRAC_PI_2, -PI)
                .line_to(x3, ry2)
                .close()
                .finish(),
            PresetPathFill::DarkenLess,
            false,
        ),
        layer(
            {
                let mut commands = body(true);
                commands.extend(
                    Outline::new(frame)
                        .move_to(x3, y1)
                        .line_to(x3, ry2)
                        .move_to(x2, y2)
                        .line_to(x2, ly3)
                        .finish(),
                );
                commands
            },
            PresetPathFill::None,
            true,
        ),
    ]
}

fn left_right_circular_arrow(
    frame: Frame,
    adjustments: &HashMap<String, f64>,
) -> Vec<GeometryPathCommand> {
    let (w, h) = (frame.w, frame.h);
    let (hc, vc) = (w / 2.0, h / 2.0);
    let a5 = adjustment(adjustments, "adj5", 0.125).clamp(0.0, 25000.0);
    let max_adj1 = a5 * 2.0;
    let a1 = adjustment(adjustments, "adj1", 0.125).clamp(0.0, max_adj1.max(0.0));
    let en_ang = adjustment(adjustments, "adj3", 204.57681).clamp(1.0, 21599999.0);
    let st_ang = adjustment(adjustments, "adj4", 119.42319).clamp(0.0, 21599999.0);
    let th = a1;
    let thh = a5;
    let th2 = th / 2.0;
    let rw1 = hc + th2 - thh;
    let rh1 = vc + th2 - thh;
    let rw2 = rw1 - th;
    let rh2 = rh1 - th;
    let rw3 = rw2 + th2;
    let rh3 = rh2 + th2;
    let (dx_h, dy_h) = elliptical_point(ooxml_angle(en_ang), rw3, rh3, frame);
    let (x_h, y_h) = (hc + dx_h, vc + dy_h);
    let r_i = rw2.min(rh2);
    let u1 = dx_h * dx_h;
    let u2 = dy_h * dy_h;
    let u3 = r_i * r_i;
    let u4 = u1 - u3;
    let u5 = u2 - u3;
    let u6 = divide(u4 * u5, u1);
    let u7 = divide(u6, u2);
    let u9 = (1.0 - u7).max(0.0).sqrt();
    let u10 = divide(u4, dx_h);
    let u11 = divide(u10, dy_h);
    let u12 = divide(1.0 + u9, u11);
    let u13 = u12.atan2(1.0).to_degrees() * 60000.0;
    let u15 = if u13 > 0.0 { u13 } else { u13 + 21600000.0 };
    let u16 = u15 - en_ang;
    let u18 = if u16 > 0.0 { u16 } else { u16 + 21600000.0 };
    let u21 = if u18 - 10800000.0 > 0.0 {
        u18 - 21600000.0
    } else {
        u18
    };
    let max_ang = u21.abs();
    let a_ang = adjustment(adjustments, "adj2", 11.42319).clamp(0.0, max_ang.max(0.0));
    let pt_ang = en_ang + a_ang;
    let (dx_a, dy_a) = elliptical_point(ooxml_angle(pt_ang), rw3, rh3, frame);
    let (x_a, y_a) = (hc + dx_a, vc + dy_a);
    let (dx_g, dy_g) = (
        thh * ooxml_angle(pt_ang).cos(),
        thh * ooxml_angle(pt_ang).sin(),
    );
    let (x_g, y_g) = (x_h + dx_g, y_h + dy_g);
    let (x_b, y_b) = (x_h - dx_g, y_h - dy_g);
    let (sx1, sy1) = (x_b - hc, y_b - vc);
    let (sx2, sy2) = (x_g - hc, y_g - vc);
    let intersect = |rw: f64, rh: f64, r_min: f64, from_second: bool| -> (f64, f64) {
        let (x1o, y1o) = (divide(sx1 * r_min, rw), divide(sy1 * r_min, rh));
        let (x2o, y2o) = (divide(sx2 * r_min, rw), divide(sy2 * r_min, rh));
        let (dxo, dyo) = (x2o - x1o, y2o - y1o);
        let d_o = dxo.hypot(dyo);
        let det = x1o * y2o - x2o * y1o;
        let q4 = d_o * d_o;
        let sdel = (r_min * r_min * q4 - det * det).max(0.0).sqrt();
        let sdy = if -dyo > 0.0 { -1.0 } else { 1.0 };
        let dxf1 = divide(det * dyo + sdy * dxo * sdel, q4);
        let dxf2 = divide(det * dyo - sdy * dxo * sdel, q4);
        let dyf1 = divide(-det * dxo + dyo.abs() * sdel, q4);
        let dyf2 = divide(-det * dxo - dyo.abs() * sdel, q4);
        let (bx, by) = if from_second { (x2o, y2o) } else { (x1o, y1o) };
        let q20 = (bx - dxf1).hypot(by - dyf1);
        let q21 = (bx - dxf2).hypot(by - dyf2);
        let (dxf, dyf) = if q21 - q20 > 0.0 {
            (dxf1, dyf1)
        } else {
            (dxf2, dyf2)
        };
        (hc + divide(dxf * rw, r_min), vc + divide(dyf * rh, r_min))
    };
    let r_o = rw1.min(rh1);
    let (x_f, y_f) = intersect(rw1, rh1, r_o, true);
    let (x_c, y_c) = intersect(rw2, rh2, r_i, false);
    let (dx_i2, dy_i2) = elliptical_point(ooxml_angle(st_ang), rw3, rh3, frame);
    let (x_i, y_i) = (hc + dx_i2, vc + dy_i2);
    let lpt_ang = st_ang - a_ang;
    let (dx_l, dy_l) = elliptical_point(ooxml_angle(lpt_ang), rw3, rh3, frame);
    let (x_l, y_l) = (hc + dx_l, vc + dy_l);
    let (dx_k, dy_k) = (
        thh * ooxml_angle(lpt_ang).cos(),
        thh * ooxml_angle(lpt_ang).sin(),
    );
    let (x_k, y_k) = (x_i + dx_k, y_i + dy_k);
    let (x_j, y_j) = (x_i - dx_k, y_i - dy_k);
    let p5 = (x_f - x_c).hypot(y_f - y_c) / 2.0 - thh;
    let (x_gp, y_gp) = if p5 > 0.0 { (x_f, y_f) } else { (x_g, y_g) };
    let (x_bp, y_bp) = if p5 > 0.0 { (x_c, y_c) } else { (x_b, y_b) };
    let to_units = |x: f64, y: f64| {
        let v = y.atan2(x).to_degrees() * 60000.0;
        if v > 0.0 { v } else { v + 21600000.0 }
    };
    let (sdx_f, sdy_f) = (x_f - hc, y_f - vc);
    let (sdx_c, sdy_c) = (x_c - hc, y_c - vc);
    let en2 = to_units(sdx_f, sdy_f);
    let od0 = en2 - en_ang;
    let od2 = if od0 > 0.0 { od0 } else { od0 + 21600000.0 };
    let st0 = st_ang - od2;
    let st2 = if st0 > 0.0 { st0 } else { st0 + 21600000.0 };
    let sw0 = en2 - st2;
    let sw_ang = if sw0 > 0.0 { sw0 } else { sw0 + 21600000.0 };
    let ist_ang = to_units(sdx_c, sdy_c);
    let id0 = ist_ang - en_ang;
    let id2 = if id0 > 0.0 { id0 - 21600000.0 } else { id0 };
    let ien0 = st_ang - id2;
    let ien1 = ien0 - 21600000.0;
    let ien2 = if ien1 > 0.0 { ien1 } else { ien0 };
    let isw1 = ien2 - ist_ang;
    let isw_ang = if isw1 > 0.0 { isw1 - 21600000.0 } else { isw1 };
    let (dx_e, dy_e) = elliptical_point(ooxml_angle(st2), rw1, rh1, frame);
    let (x_e, y_e) = (hc + dx_e, vc + dy_e);
    let (dx_d, dy_d) = elliptical_point(ooxml_angle(ien2), rw2, rh2, frame);
    let (x_d, y_d) = (hc + dx_d, vc + dy_d);
    let (x_kp, y_kp) = if p5 > 0.0 { (x_e, y_e) } else { (x_k, y_k) };
    let (x_jp, y_jp) = if p5 > 0.0 { (x_d, y_d) } else { (x_j, y_j) };
    Outline::new(frame)
        .move_to(x_l, y_l)
        .line_to(x_kp, y_kp)
        .line_to(x_e, y_e)
        .arc_to(rw1, rh1, ooxml_angle(st2), ooxml_angle(sw_ang))
        .line_to(x_gp, y_gp)
        .line_to(x_a, y_a)
        .line_to(x_bp, y_bp)
        .line_to(x_c, y_c)
        .arc_to(rw2, rh2, ooxml_angle(ist_ang), ooxml_angle(isw_ang))
        .line_to(x_jp, y_jp)
        .close()
        .finish()
}
