use std::collections::{HashMap, VecDeque};

use ooxml_text::{FontId, FontStore, PathCmd};
use tiny_skia::{FillRule, Mask, Paint, Path, Pixmap, Transform};

const MAX_CACHED_GLYPH_OUTLINES: usize = 8_192;

/// Store-bound glyph outlines with FIFO eviction.
#[derive(Default)]
pub struct GlyphCache {
    entries: HashMap<(u32, u32), CachedGlyph>,
    order: VecDeque<(u32, u32)>,
    store: Option<u64>,
}

/// A cached path and units per em.
#[derive(Debug)]
pub struct CachedGlyph {
    pub path: Option<Path>,
    pub upem: f32,
}

impl GlyphCache {
    /// Looks up an outline, converting its id and building its path on misses.
    pub fn lookup(
        &mut self,
        fonts: &FontStore,
        font: FontId,
        glyph_id: u32,
        outline_id: impl FnOnce() -> Result<u16, String>,
        build_path: impl FnOnce(&[PathCmd]) -> Result<Option<Path>, String>,
    ) -> Result<&CachedGlyph, String> {
        if self.store.is_some_and(|bound| bound != fonts.id()) {
            return Err("glyph cache is bound to another font store".to_string());
        }
        let key = (font.to_u32(), glyph_id);
        if !self.entries.contains_key(&key) {
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
                },
            );
        }
        Ok(self
            .entries
            .get(&key)
            .expect("a just-inserted or occupied key is present"))
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
