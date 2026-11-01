use std::collections::HashMap;
use std::sync::Arc;

use crate::DeckSession;

#[doc(hidden)]
pub struct MetafileSvgCache {
    entries: HashMap<String, Option<Arc<[u8]>>>,
    source_bytes: usize,
    svg_bytes: usize,
    replay: ooxml_metafile::ReplayBudget,
}

impl Default for MetafileSvgCache {
    fn default() -> Self {
        let picture = ooxml_metafile::ReplayBudget::default();
        Self {
            entries: HashMap::new(),
            source_bytes: 64 * 1024 * 1024,
            svg_bytes: 64 * 1024 * 1024,
            replay: ooxml_metafile::ReplayBudget {
                work: 4 * picture.work,
                pixels: 4 * picture.pixels,
            },
        }
    }
}

impl MetafileSvgCache {
    pub fn resolve(&mut self, key: &str, bytes: &[u8]) -> Option<Arc<[u8]>> {
        if let Some(svg) = self.entries.get(key) {
            return svg.clone();
        }
        if !ooxml_metafile::is_metafile(bytes) {
            return None;
        }
        let svg = if bytes.len() > 16 * 1024 * 1024
            || bytes.len() > self.source_bytes
            || self.svg_bytes == 0
            || self.replay.work == 0
        {
            None
        } else {
            self.source_bytes -= bytes.len();
            ooxml_metafile::to_svg_with_limits(bytes, &mut self.replay, &mut self.svg_bytes)
                .ok()
                .map(|svg| Arc::from(svg.markup.into_bytes()))
        };
        self.entries.insert(key.to_owned(), svg.clone());
        svg
    }
}

impl DeckSession {
    /// Resolves a cached display copy without changing the source media.
    pub fn metafile_svg(&self, part_path: &str) -> Option<Arc<[u8]>> {
        if let Some(svg) = self.metafile_svgs.borrow().entries.get(part_path) {
            return svg.clone();
        }
        let bytes = self.media_bytes(part_path).ok()?;
        self.metafile_svgs.borrow_mut().resolve(part_path, &bytes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/metafile-text.emf");
    const BITMAP: &[u8] =
        include_bytes!("../../pptx-render/tests/fixtures/metafile-bitmap-line.emf");

    #[test]
    fn replay_is_lazy_and_caches_successes_and_refusals() {
        let mut cache = MetafileSvgCache::default();
        let initial = cache.replay;
        assert!(cache.entries.is_empty());
        for (key, bytes) in [("text", TEXT), ("bitmap", BITMAP)] {
            assert!(ooxml_metafile::decode(bytes).is_none());
            let svg = cache.resolve(key, bytes).unwrap();
            assert!(std::str::from_utf8(&svg).unwrap().contains("<svg"));
            let remaining = (cache.replay, cache.svg_bytes, cache.source_bytes);
            assert!(Arc::ptr_eq(&svg, &cache.resolve(key, bytes).unwrap()));
            assert_eq!(
                remaining,
                (cache.replay, cache.svg_bytes, cache.source_bytes)
            );
        }
        assert!(cache.replay.work < initial.work);
        let mut refused = TEXT.to_vec();
        refused.truncate(44);
        assert!(cache.resolve("refused", &refused).is_none());
        let remaining = (cache.replay, cache.svg_bytes, cache.source_bytes);
        assert!(cache.resolve("refused", &refused).is_none());
        assert_eq!(
            remaining,
            (cache.replay, cache.svg_bytes, cache.source_bytes)
        );
    }

    #[test]
    fn replay_and_output_exhaustion_draw_nothing() {
        for limit in ["work", "pixels", "output", "source"] {
            let mut cache = MetafileSvgCache::default();
            match limit {
                "work" => cache.replay.work = 0,
                "pixels" => cache.replay.pixels = 0,
                "output" => cache.svg_bytes = 1,
                _ => cache.source_bytes = 0,
            }
            let bytes = if limit == "pixels" { BITMAP } else { TEXT };
            assert!(cache.resolve(limit, bytes).is_none(), "{limit}");
            assert!(cache.resolve(limit, bytes).is_none(), "{limit}");
        }
        let mut oversized = TEXT.to_vec();
        oversized.resize(16 * 1024 * 1024 + 1, 0);
        assert!(
            MetafileSvgCache::default()
                .resolve("large", &oversized)
                .is_none()
        );
    }
}
