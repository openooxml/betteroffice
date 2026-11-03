use crate::{ReplayBudget, is_metafile, to_svg_with_limits};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SvgSession {
    source_bytes: usize,
    svg_bytes: usize,
    replay: ReplayBudget,
}

impl Default for SvgSession {
    fn default() -> Self {
        let picture = ReplayBudget::default();
        Self {
            source_bytes: 64 * 1024 * 1024,
            svg_bytes: 64 * 1024 * 1024,
            replay: ReplayBudget {
                work: 4 * picture.work,
                pixels: 4 * picture.pixels,
            },
        }
    }
}

impl SvgSession {
    pub fn convert(&mut self, bytes: &[u8]) -> Option<Vec<u8>> {
        if !is_metafile(bytes)
            || bytes.len() > 16 * 1024 * 1024
            || bytes.len() > self.source_bytes
            || self.svg_bytes == 0
            || self.replay.work == 0
        {
            return None;
        }
        self.source_bytes -= bytes.len();
        to_svg_with_limits(bytes, &mut self.replay, &mut self.svg_bytes)
            .ok()
            .map(|svg| svg.markup.into_bytes())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_records::{Emf, cropped_rle_emf, i32s, rect};

    #[test]
    fn default_allowances_match_the_session_limits() {
        let session = SvgSession::default();
        let picture = ReplayBudget::default();
        assert_eq!(session.source_bytes, 64 * 1024 * 1024);
        assert_eq!(session.svg_bytes, 64 * 1024 * 1024);
        assert_eq!(session.replay.work, 4 * picture.work);
        assert_eq!(session.replay.pixels, 4 * picture.pixels);
    }

    #[test]
    fn distinct_pictures_share_the_cumulative_pixel_allowance() {
        let mut session = SvgSession::default();
        let side = 1024;
        let count = session.replay.pixels / (side * side);
        for index in 0..count {
            let mut bytes = cropped_rle_emf(side as u32);
            bytes[8..12].copy_from_slice(&(index as i32).to_le_bytes());
            assert!(session.convert(&bytes).is_some(), "picture {index}");
        }
        assert_eq!(session.replay.pixels, 0);
        assert!(session.replay.work > 0);
        assert!(session.source_bytes > 0);
        assert!(session.svg_bytes > 0);
        let next = cropped_rle_emf(2);
        assert!(SvgSession::default().convert(&next).is_some());
        assert!(session.convert(&next).is_none());
    }

    #[test]
    fn replay_and_byte_exhaustion_refuse() {
        let text = Emf::new(10, 10).recs(vec![rect(0, 0, 5, 5)]).bytes();
        let bitmap = cropped_rle_emf(2);
        for limit in ["work", "pixels", "output", "source"] {
            let mut session = SvgSession::default();
            match limit {
                "work" => session.replay.work = 0,
                "pixels" => session.replay.pixels = 0,
                "output" => session.svg_bytes = 1,
                _ => session.source_bytes = 0,
            }
            let bytes = if limit == "pixels" { &bitmap } else { &text };
            assert!(session.convert(bytes).is_none(), "{limit}");
        }
        let mut oversized = text;
        oversized.resize(16 * 1024 * 1024 + 1, 0);
        assert!(SvgSession::default().convert(&oversized).is_none());
    }

    #[test]
    fn source_and_svg_allowances_are_cumulative() {
        let bytes = Emf::new(10, 10).recs(vec![rect(0, 0, 5, 5)]).bytes();
        let svg_len = SvgSession::default().convert(&bytes).unwrap().len();
        for limit in ["source", "svg"] {
            let mut session = SvgSession::default();
            if limit == "source" {
                session.source_bytes = 2 * bytes.len();
            } else {
                session.svg_bytes = 2 * svg_len;
            }
            assert!(session.convert(&bytes).is_some(), "{limit}");
            assert!(session.convert(&bytes).is_some(), "{limit}");
            assert!(session.convert(&bytes).is_none(), "{limit}");
        }
    }

    #[test]
    fn refused_sources_spend_cumulative_allowances() {
        let mut broken = cropped_rle_emf(2);
        broken.truncate(broken.len() - 20);
        let mut session = SvgSession::default();
        let initial = session.clone();
        assert!(session.convert(&broken).is_none());
        assert_eq!(session.source_bytes, initial.source_bytes - broken.len());
        assert_eq!(session.replay.pixels, initial.replay.pixels - 4);
        assert!(session.replay.work < initial.replay.work);
        let remaining = session.clone();
        assert!(session.convert(&i32s(&[0, 1, 2])).is_none());
        assert_eq!(session, remaining);
    }
}
