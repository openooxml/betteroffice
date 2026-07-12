use serde::Serialize;
use skrifa::instance::{LocationRef, Size};
use skrifa::outline::{DrawSettings, OutlinePen};
use skrifa::{FontRef, GlyphId, MetadataProvider};

use crate::font_store::{FontError, FontId, FontStore};

#[derive(Debug, Clone, Copy, PartialEq, Serialize)]
#[serde(tag = "t")]
pub enum PathCmd {
    /// Start a new contour at `(x, y)`.
    #[serde(rename = "M")]
    MoveTo { x: f32, y: f32 },
    /// Straight segment to `(x, y)`.
    #[serde(rename = "L")]
    LineTo { x: f32, y: f32 },
    /// Quadratic Bézier through control `(cx, cy)` to `(x, y)` — TrueType.
    #[serde(rename = "Q")]
    QuadTo { cx: f32, cy: f32, x: f32, y: f32 },
    /// Cubic Bézier through controls `(c1x, c1y)`, `(c2x, c2y)` to `(x, y)` —
    /// CFF/CFF2.
    #[serde(rename = "C")]
    CubicTo {
        c1x: f32,
        c1y: f32,
        c2x: f32,
        c2y: f32,
        x: f32,
        y: f32,
    },
    /// Close the current contour.
    #[serde(rename = "Z")]
    Close,
}

#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct GlyphOutline {
    pub upem: u16,
    /// Path commands in draw order, or empty for a blank glyph.
    pub cmds: Vec<PathCmd>,
}

/// Collects skrifa's draw callbacks into [`PathCmd`]s, verbatim, in font units.
#[derive(Default)]
struct CmdPen {
    cmds: Vec<PathCmd>,
}

impl OutlinePen for CmdPen {
    fn move_to(&mut self, x: f32, y: f32) {
        self.cmds.push(PathCmd::MoveTo { x, y });
    }

    fn line_to(&mut self, x: f32, y: f32) {
        self.cmds.push(PathCmd::LineTo { x, y });
    }

    fn quad_to(&mut self, cx: f32, cy: f32, x: f32, y: f32) {
        self.cmds.push(PathCmd::QuadTo { cx, cy, x, y });
    }

    fn curve_to(&mut self, c1x: f32, c1y: f32, c2x: f32, c2y: f32, x: f32, y: f32) {
        self.cmds.push(PathCmd::CubicTo {
            c1x,
            c1y,
            c2x,
            c2y,
            x,
            y,
        });
    }

    fn close(&mut self) {
        self.cmds.push(PathCmd::Close);
    }
}

impl FontStore {
    /// Outline of `glyph_id` in font `id`, in font design units.
    ///
    /// Returns empty `cmds` for a blank glyph (space). Errors with
    /// [`FontError::UnknownFont`] on an unregistered `id`,
    /// [`FontError::GlyphOutOfRange`] when the font has no such glyph, and
    /// [`FontError::Outline`] if the draw itself fails — never panics on
    /// attacker-controlled bytes.
    pub fn outline_glyph(&self, id: FontId, glyph_id: u16) -> Result<GlyphOutline, FontError> {
        // upem from the metrics captured at register(); cheap infallible lookup.
        let upem = self.metrics(id)?.units_per_em;

        // register() already proved these bytes parse; reparse still stays
        // panic-free (no unwrap) since a re-read is the trust boundary here.
        let bytes = self.font_bytes(id)?;
        let font = FontRef::new(bytes).map_err(|e| FontError::Parse(e.to_string()))?;

        let glyphs = font.outline_glyphs();
        let glyph = glyphs
            .get(GlyphId::from(glyph_id))
            .ok_or(FontError::GlyphOutOfRange(glyph_id))?;

        // Unscaled → coords come out in font units, matching FontStore metrics.
        // No hinting: hinting is a pixel-grid operation and would be wrong at
        // design-space size.
        let mut pen = CmdPen::default();
        glyph
            .draw(
                DrawSettings::unhinted(Size::unscaled(), LocationRef::default()),
                &mut pen,
            )
            .map_err(|e| FontError::Outline(e.to_string()))?;

        Ok(GlyphOutline {
            upem,
            cmds: pen.cmds,
        })
    }

    pub fn outline_glyph_json(&self, id: FontId, glyph_id: u16) -> Result<String, FontError> {
        let outline = self.outline_glyph(id, glyph_id)?;
        // Our own well-formed struct (finite font-unit floats); a serializer
        // failure is not reachable from valid outlines but is surfaced rather
        // than unwrapped to keep the untrusted-font path panic-free.
        serde_json::to_string(&outline).map_err(|e| FontError::Outline(e.to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const LIBERATION_SANS: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

    fn store() -> (FontStore, FontId) {
        let mut store = FontStore::new();
        let id = store
            .register(LIBERATION_SANS.to_vec())
            .expect("fixture font registers");
        (store, id)
    }

    #[test]
    fn outlines_a_letter_with_a_plausible_bbox() {
        let (store, id) = store();
        let gid = store
            .glyph_id(id, 'A')
            .expect("known font")
            .expect("cmap covers 'A'");
        assert_eq!(gid, 36, "LiberationSans 'A' glyph id");

        let outline = store.outline_glyph(id, gid).expect("outline 'A'");

        assert_eq!(outline.upem, 2048, "LiberationSans units per em");
        assert!(!outline.cmds.is_empty(), "'A' has contours");
        assert!(
            matches!(outline.cmds.first(), Some(PathCmd::MoveTo { .. })),
            "a contour starts with MoveTo"
        );
        assert!(
            matches!(outline.cmds.last(), Some(PathCmd::Close)),
            "a contour ends with Close"
        );

        // LiberationSans is TrueType (glyf): the pen must receive quadratics,
        // never cubics — we pass the curve kind through verbatim.
        assert!(
            outline
                .cmds
                .iter()
                .any(|c| matches!(c, PathCmd::QuadTo { .. })),
            "'A' has quadratic curves"
        );
        assert!(
            !outline
                .cmds
                .iter()
                .any(|c| matches!(c, PathCmd::CubicTo { .. })),
            "TrueType glyf emits no cubics"
        );

        // Every coordinate is finite and inside a plausible design-space box.
        let upem = outline.upem as f32;
        for cmd in &outline.cmds {
            for v in coords(cmd) {
                assert!(v.is_finite(), "coord finite");
                assert!(
                    v >= -upem && v <= upem * 1.5,
                    "coord {v} within design-space box (upem {upem})"
                );
            }
        }
    }

    #[test]
    fn blank_glyph_has_no_commands() {
        let (store, id) = store();
        let space = store
            .glyph_id(id, ' ')
            .expect("known font")
            .expect("cmap covers space");
        let outline = store.outline_glyph(id, space).expect("outline space");
        assert_eq!(outline.upem, 2048);
        assert!(outline.cmds.is_empty(), "space glyph paints nothing");
    }

    #[test]
    fn unknown_font_errors() {
        let (store, _) = store();
        let bogus = FontId::from_u32(999);
        assert_eq!(store.outline_glyph(bogus, 36), Err(FontError::UnknownFont));
    }

    #[test]
    fn glyph_out_of_range_errors() {
        let (store, id) = store();
        assert_eq!(
            store.outline_glyph(id, u16::MAX),
            Err(FontError::GlyphOutOfRange(u16::MAX))
        );
    }

    #[test]
    fn json_wire_shape_matches_contract() {
        let (store, id) = store();
        let gid = store.glyph_id(id, 'A').unwrap().unwrap();
        let json = store.outline_glyph_json(id, gid).expect("json");

        let v: serde_json::Value = serde_json::from_str(&json).expect("valid json");
        assert_eq!(v["upem"], 2048);

        let cmds = v["cmds"].as_array().expect("cmds array");
        assert!(!cmds.is_empty());

        // First command: MoveTo → {"t":"M","x":..,"y":..}
        let first = &cmds[0];
        assert_eq!(first["t"], "M");
        assert!(first["x"].is_number(), "M has x");
        assert!(first["y"].is_number(), "M has y");

        // Every tag is one of the pinned set with the pinned coordinate keys.
        for cmd in cmds {
            match cmd["t"].as_str().expect("tag") {
                "M" | "L" => {
                    assert!(cmd["x"].is_number() && cmd["y"].is_number());
                }
                "Q" => {
                    assert!(cmd["cx"].is_number() && cmd["cy"].is_number());
                    assert!(cmd["x"].is_number() && cmd["y"].is_number());
                }
                "C" => {
                    assert!(cmd["c1x"].is_number() && cmd["c1y"].is_number());
                    assert!(cmd["c2x"].is_number() && cmd["c2y"].is_number());
                    assert!(cmd["x"].is_number() && cmd["y"].is_number());
                }
                "Z" => {}
                other => panic!("unexpected path tag {other}"),
            }
        }
        assert_eq!(cmds.last().unwrap()["t"], "Z");
    }

    fn coords(cmd: &PathCmd) -> Vec<f32> {
        match *cmd {
            PathCmd::MoveTo { x, y } | PathCmd::LineTo { x, y } => vec![x, y],
            PathCmd::QuadTo { cx, cy, x, y } => vec![cx, cy, x, y],
            PathCmd::CubicTo {
                c1x,
                c1y,
                c2x,
                c2y,
                x,
                y,
            } => vec![c1x, c1y, c2x, c2y, x, y],
            PathCmd::Close => vec![],
        }
    }
}
