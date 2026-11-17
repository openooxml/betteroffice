#[cfg(feature = "metafile-svg")]
use std::collections::HashMap;
use std::sync::Arc;

use crate::DeckSession;

#[doc(hidden)]
#[derive(Default)]
pub struct MetafileSvgCache {
    #[cfg(feature = "metafile-svg")]
    entries: HashMap<String, Option<Arc<[u8]>>>,
    #[cfg(feature = "metafile-svg")]
    session: ooxml_metafile::SvgSession,
}

impl MetafileSvgCache {
    #[cfg(not(feature = "metafile-svg"))]
    pub fn resolve(&mut self, _key: &str, _bytes: &[u8]) -> Option<Arc<[u8]>> {
        None
    }

    #[cfg(feature = "metafile-svg")]
    pub fn resolve(&mut self, key: &str, bytes: &[u8]) -> Option<Arc<[u8]>> {
        if let Some(svg) = self.entries.get(key) {
            return svg.clone();
        }
        if !ooxml_metafile::is_metafile(bytes) {
            return None;
        }
        let svg = self.session.convert(bytes).map(Arc::from);
        self.entries.insert(key.to_owned(), svg.clone());
        svg
    }
}

impl DeckSession {
    /// Resolves a cached display copy without changing the source media.
    pub fn metafile_svg(&self, part_path: &str) -> Option<Arc<[u8]>> {
        #[cfg(not(feature = "metafile-svg"))]
        {
            self.metafile_svgs.borrow_mut().resolve(part_path, &[])
        }
        #[cfg(feature = "metafile-svg")]
        {
            if let Some(svg) = self.metafile_svgs.borrow().entries.get(part_path) {
                return svg.clone();
            }
            let bytes = self.media_bytes(part_path).ok()?;
            self.metafile_svgs.borrow_mut().resolve(part_path, &bytes)
        }
    }
}

#[cfg(all(test, feature = "metafile-svg"))]
mod tests {
    use super::*;

    const TEXT: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/metafile-text.emf");
    const BITMAP: &[u8] =
        include_bytes!("../../pptx-render/tests/fixtures/metafile-bitmap-line.emf");

    #[test]
    fn replay_is_lazy_and_caches_successes_and_refusals() {
        let mut cache = MetafileSvgCache::default();
        let initial = cache.session.clone();
        assert!(cache.entries.is_empty());
        for (key, bytes) in [("text", TEXT), ("bitmap", BITMAP)] {
            assert!(ooxml_metafile::decode(bytes).is_none());
            let svg = cache.resolve(key, bytes).unwrap();
            assert!(std::str::from_utf8(&svg).unwrap().contains("<svg"));
            let remaining = cache.session.clone();
            assert!(Arc::ptr_eq(&svg, &cache.resolve(key, bytes).unwrap()));
            assert_eq!(remaining, cache.session);
        }
        assert_ne!(cache.session, initial);
        let mut refused = TEXT.to_vec();
        refused.truncate(44);
        assert!(cache.resolve("refused", &refused).is_none());
        let remaining = cache.session.clone();
        assert!(cache.resolve("refused", &refused).is_none());
        assert_eq!(remaining, cache.session);
    }

    #[test]
    fn oversized_sources_draw_nothing() {
        let mut oversized = TEXT.to_vec();
        oversized.resize(16 * 1024 * 1024 + 1, 0);
        assert!(
            MetafileSvgCache::default()
                .resolve("large", &oversized)
                .is_none()
        );
    }
}

#[cfg(all(test, not(feature = "metafile-svg")))]
mod lean_tests {
    use super::*;

    #[test]
    fn svg_is_unavailable() {
        let bytes = include_bytes!("../../pptx-render/tests/fixtures/metafile-text.emf");
        assert!(MetafileSvgCache::default().resolve("text", bytes).is_none());
        let session = DeckSession::open(
            include_bytes!("../../pptx-render/tests/fixtures/metafile-svg-pictures.pptx"),
            1,
        )
        .unwrap();
        assert!(session.metafile_svg("ppt/media/image1.emf").is_none());
    }
}
