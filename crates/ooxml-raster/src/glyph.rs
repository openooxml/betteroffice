use std::collections::{HashMap, VecDeque};

use ooxml_text::{FontId, FontStore, GlyphOutline, PathCmd};
use tiny_skia::{FillRule, Mask, Paint, Path, Pixmap, Transform};

const MAX_CACHED_GLYPH_OUTLINES: usize = 8_192;

/// Store-bound glyph outlines with FIFO eviction.
#[derive(Default)]
pub struct GlyphCache {
    entries: HashMap<(u32, u32), CachedGlyph>,
    order: VecDeque<(u32, u32)>,
    store: Option<u64>,
    outlines: u64,
}

/// A cached path, its units per em, and its synthetic-bold widening.
#[derive(Debug)]
pub struct CachedGlyph {
    pub path: Option<Path>,
    pub upem: f32,
    /// Widened geometry for synthetic bold, resolved on the first emboldened
    /// lookup. `None` means widening is unavailable, so callers paint `path`.
    pub bold_path: Option<Path>,
    bold_resolved: bool,
}

impl GlyphCache {
    /// Looks up an outline, converting its id and building its path on misses.
    pub fn lookup(
        &mut self,
        fonts: &FontStore,
        font: FontId,
        glyph_id: u32,
        outline_id: impl FnOnce() -> Result<u16, String>,
        build_path: impl Fn(&[PathCmd]) -> Result<Option<Path>, String>,
    ) -> Result<&CachedGlyph, String> {
        self.lookup_with(fonts, font, glyph_id, false, outline_id, build_path)
    }

    /// Looks up an outline together with the geometry synthetic bold paints:
    /// the same glyph widened by [`ooxml_text::embolden_glyph_outline`]. The
    /// widened path is built once per cached glyph and reused. When widening
    /// is unavailable (blank or degenerate outlines) `bold_path` stays `None`
    /// and callers paint the crisp `path` instead.
    pub fn lookup_emboldened(
        &mut self,
        fonts: &FontStore,
        font: FontId,
        glyph_id: u32,
        outline_id: impl FnOnce() -> Result<u16, String>,
        build_path: impl Fn(&[PathCmd]) -> Result<Option<Path>, String>,
    ) -> Result<&CachedGlyph, String> {
        self.lookup_with(fonts, font, glyph_id, true, outline_id, build_path)
    }

    fn lookup_with(
        &mut self,
        fonts: &FontStore,
        font: FontId,
        glyph_id: u32,
        embolden: bool,
        outline_id: impl FnOnce() -> Result<u16, String>,
        build_path: impl Fn(&[PathCmd]) -> Result<Option<Path>, String>,
    ) -> Result<&CachedGlyph, String> {
        self.bind(fonts.id(), fonts.outline_revision())?;
        let key = (font.to_u32(), glyph_id);
        if self.entries.contains_key(&key) {
            if embolden
                && let Some(cached) = self.entries.get_mut(&key)
                && !cached.bold_resolved
            {
                let id = outline_id()?;
                let outline = fonts
                    .outline_glyph(font, id)
                    .map_err(|error| error.to_string())?;
                cached.bold_path = widened_path(&outline, &build_path);
                cached.bold_resolved = true;
            }
        } else {
            let id = outline_id()?;
            let outline = fonts
                .outline_glyph(font, id)
                .map_err(|error| error.to_string())?;
            let path = build_path(&outline.cmds)?;
            self.store = Some(fonts.id());
            self.remember(
                key,
                CachedGlyph {
                    path,
                    upem: f32::from(outline.upem),
                    bold_path: if embolden {
                        widened_path(&outline, &build_path)
                    } else {
                        None
                    },
                    bold_resolved: embolden,
                },
            );
        }
        Ok(self
            .entries
            .get(&key)
            .expect("a just-inserted or occupied key is present"))
    }

    /// Binds to `store`, refusing another one, and drops every path cached
    /// under an earlier outline revision of it.
    fn bind(&mut self, store: u64, outlines: u64) -> Result<(), String> {
        if self.store.is_some_and(|bound| bound != store) {
            return Err("glyph cache is bound to another font store".to_string());
        }
        if self.outlines != outlines {
            self.entries.clear();
            self.order.clear();
            self.outlines = outlines;
        }
        Ok(())
    }

    fn remember(&mut self, key: (u32, u32), glyph: CachedGlyph) {
        if self.entries.insert(key, glyph).is_none() {
            self.order.push_back(key);
        }
        while self.order.len() > MAX_CACHED_GLYPH_OUTLINES {
            let oldest = self
                .order
                .pop_front()
                .expect("the queue tracks every cached entry");
            self.entries.remove(&oldest);
        }
    }
}

/// The synthetic-bold path for `outline`, or `None` when the outline cannot
/// be widened. A widened path the caller's policy rejects also falls back to
/// the crisp path rather than failing the page.
fn widened_path(
    outline: &GlyphOutline,
    build_path: &impl Fn(&[PathCmd]) -> Result<Option<Path>, String>,
) -> Option<Path> {
    let commands = ooxml_text::embolden_glyph_outline(outline)?;
    build_path(&commands).ok().flatten()
}

/// Fills a baseline-placed glyph using a precomputed size/em scale.
#[allow(clippy::too_many_arguments)]
pub fn fill_glyph(
    pixmap: &mut Pixmap,
    path: &Path,
    scale: f32,
    x: f32,
    y: f32,
    transform: Transform,
    paint: &Paint<'_>,
    mask: Option<&Mask>,
) {
    let glyph_transform = Transform::from_row(scale, 0.0, 0.0, -scale, x, y);
    pixmap.fill_path(
        path,
        paint,
        FillRule::Winding,
        transform.pre_concat(glyph_transform),
        mask,
    );
}

#[cfg(test)]
mod tests;
