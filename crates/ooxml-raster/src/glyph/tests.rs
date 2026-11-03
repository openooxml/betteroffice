use super::*;
use crate::outline_path;

const CARLITO: &[u8] = include_bytes!("../../../xlsx-raster/assets/Carlito-Regular.ttf");

fn store() -> (FontStore, FontId) {
    let mut store = FontStore::new();
    let font = store.register(CARLITO.to_vec()).unwrap();
    (store, font)
}

fn cached_glyph<'a>(
    cache: &'a mut GlyphCache,
    fonts: &FontStore,
    font: FontId,
    glyph_id: u32,
) -> Result<&'a CachedGlyph, String> {
    cache.lookup(
        fonts,
        font,
        glyph_id,
        || {
            u16::try_from(glyph_id)
                .map_err(|_| format!("glyph id {glyph_id} exceeds the font outline range"))
        },
        strict_outline_path,
    )
}

fn strict_outline_path(commands: &[PathCmd]) -> Result<Option<Path>, String> {
    if commands.is_empty() {
        return Ok(None);
    }
    outline_path(commands)
        .map(Some)
        .ok_or_else(|| "glyph outline has no drawable commands".to_string())
}

#[test]
fn a_hit_reuses_the_path_without_converting_or_extracting_again() {
    let (fonts, font) = store();
    let id = fonts.glyph_id(font, 'A').unwrap().unwrap();
    let mut cache = GlyphCache::default();
    let first = cached_glyph(&mut cache, &fonts, font, u32::from(id)).unwrap();
    let fresh = fonts.outline_glyph(font, id).unwrap();
    assert_eq!(first.upem, f32::from(fresh.upem));
    assert_eq!(first.path, outline_path(&fresh.cmds));
    let first_points = first.path.as_ref().unwrap().points().as_ptr();
    let second = cache
        .lookup(
            &fonts,
            font,
            u32::from(id),
            || panic!("a cache hit must not convert the glyph id"),
            |_| panic!("a cache hit must not rebuild the outline"),
        )
        .unwrap();
    assert_eq!(
        second.path.as_ref().unwrap().points().as_ptr(),
        first_points
    );
    assert_eq!(
        cache.order,
        VecDeque::from([(font.to_u32(), u32::from(id))])
    );
}

#[test]
fn eviction_is_fifo_even_after_a_hit_or_replacement() {
    let (fonts, font) = store();
    let id = fonts.glyph_id(font, 'A').unwrap().unwrap();
    let mut cache = GlyphCache::default();
    for glyph_id in 0..MAX_CACHED_GLYPH_OUTLINES as u32 {
        cache
            .lookup(
                &fonts,
                font,
                glyph_id,
                || Ok(id),
                |commands| Ok(outline_path(commands)),
            )
            .unwrap();
    }
    let first = (font.to_u32(), 0);
    let hit = cached_glyph(&mut cache, &fonts, font, 0).unwrap();
    assert!(hit.path.is_some());
    cache.remember(
        first,
        CachedGlyph {
            path: None,
            upem: 1_000.0,
        },
    );
    assert_eq!(cache.entries.len(), MAX_CACHED_GLYPH_OUTLINES);
    assert_eq!(cache.order.len(), MAX_CACHED_GLYPH_OUTLINES);
    assert_eq!(cache.order.front(), Some(&first));
    cache
        .lookup(
            &fonts,
            font,
            MAX_CACHED_GLYPH_OUTLINES as u32,
            || Ok(id),
            |commands| Ok(outline_path(commands)),
        )
        .unwrap();
    assert_eq!(cache.entries.len(), MAX_CACHED_GLYPH_OUTLINES);
    assert!(!cache.entries.contains_key(&first));
    assert_eq!(cache.order.front(), Some(&(font.to_u32(), 1)));
}

#[test]
fn foreign_stores_fail_before_id_conversion_with_the_exact_error() {
    let (first, font) = store();
    let id = first.glyph_id(font, 'A').unwrap().unwrap();
    let mut cache = GlyphCache::default();
    cached_glyph(&mut cache, &first, font, u32::from(id)).unwrap();
    let (second, other) = store();
    for glyph_id in [u32::from(id), u32::MAX] {
        let error = cache
            .lookup(
                &second,
                other,
                glyph_id,
                || panic!("a foreign store must fail before id conversion"),
                |_| panic!("a foreign store must fail before path building"),
            )
            .unwrap_err();
        assert_eq!(error, "glyph cache is bound to another font store");
    }
    assert_eq!(cache.store, Some(first.id()));
    assert_eq!(cache.entries.len(), 1);
}

#[test]
fn empty_and_undrawable_outlines_keep_the_callers_policy() {
    assert!(strict_outline_path(&[]).unwrap().is_none());
    let undrawable = [PathCmd::MoveTo { x: 0.0, y: 0.0 }];
    assert_eq!(
        strict_outline_path(&undrawable).unwrap_err(),
        "glyph outline has no drawable commands"
    );
    assert!(outline_path(&undrawable).is_none());

    let (fonts, font) = store();
    let space = fonts.glyph_id(font, ' ').unwrap().unwrap();
    let mut cache = GlyphCache::default();
    assert!(
        cached_glyph(&mut cache, &fonts, font, u32::from(space))
            .unwrap()
            .path
            .is_none()
    );
    assert_eq!(cache.store, Some(fonts.id()));
    let id = fonts.glyph_id(font, 'A').unwrap().unwrap();
    let mut rejected = GlyphCache::default();
    assert_eq!(
        rejected
            .lookup(
                &fonts,
                font,
                u32::from(id),
                || Ok(id),
                |_| strict_outline_path(&undrawable),
            )
            .unwrap_err(),
        "glyph outline has no drawable commands"
    );
    assert!(rejected.store.is_none());
    assert!(rejected.entries.is_empty());
    assert!(rejected.order.is_empty());
    let skipped = rejected
        .lookup(
            &fonts,
            font,
            u32::from(id),
            || Ok(id),
            |_| Ok(outline_path(&undrawable)),
        )
        .unwrap();
    assert!(skipped.path.is_none());
    assert_eq!(rejected.entries.len(), 1);
}

#[test]
fn outline_errors_do_not_bind_or_fill_the_cache() {
    let (fonts, font) = store();
    let mut cache = GlyphCache::default();
    let expected = fonts.outline_glyph(font, u16::MAX).unwrap_err().to_string();
    assert_eq!(
        cached_glyph(&mut cache, &fonts, font, u32::from(u16::MAX)).unwrap_err(),
        expected
    );
    assert_eq!(
        cached_glyph(&mut cache, &fonts, font, u32::MAX).unwrap_err(),
        "glyph id 4294967295 exceeds the font outline range"
    );
    assert!(cache.store.is_none());
    assert!(cache.entries.is_empty());
    assert!(cache.order.is_empty());
}

#[test]
fn shared_fill_matches_the_original_transform_rule_and_mask() {
    let (fonts, font) = store();
    let id = fonts.glyph_id(font, 'A').unwrap().unwrap();
    let outline = fonts.outline_glyph(font, id).unwrap();
    let path = outline_path(&outline.cmds).unwrap();
    let scale = 43.0 / f32::from(outline.upem);
    let mut paint = Paint::default();
    paint.set_color_rgba8(20, 70, 130, 190);
    paint.anti_alias = true;
    let mut mask = Mask::new(64, 64).unwrap();
    mask.data_mut().fill(192);
    for row in mask.data_mut().chunks_mut(64) {
        row[..17].fill(0);
    }
    for transform in [
        Transform::identity(),
        Transform::from_row(1.07, 0.13, -0.19, 0.91, 4.0, 2.0),
    ] {
        for mask in [None, Some(&mask)] {
            let mut expected = Pixmap::new(64, 64).unwrap();
            let mut actual = expected.clone();
            let glyph_transform = Transform::from_row(scale, 0.0, 0.0, -scale, 13.25, 47.5);
            expected.fill_path(
                &path,
                &paint,
                FillRule::Winding,
                transform.pre_concat(glyph_transform),
                mask,
            );
            fill_glyph(
                &mut actual,
                &path,
                scale,
                13.25,
                47.5,
                transform,
                &paint,
                mask,
            );
            assert_eq!(actual.data(), expected.data());
        }
    }
}

#[test]
fn the_glyph_cache_evicts_its_oldest_entries_past_the_cap() {
    let mut cache = GlyphCache::default();
    for index in 0..MAX_CACHED_GLYPH_OUTLINES as u32 + 2 {
        cache.remember(
            (index, 0),
            CachedGlyph {
                path: None,
                upem: 1000.0,
            },
        );
    }
    assert_eq!(cache.entries.len(), MAX_CACHED_GLYPH_OUTLINES);
    assert_eq!(cache.order.len(), MAX_CACHED_GLYPH_OUTLINES);
    assert!(!cache.entries.contains_key(&(0, 0)));
    assert!(cache.entries.contains_key(&(2, 0)));
    assert!(
        cache
            .entries
            .contains_key(&(MAX_CACHED_GLYPH_OUTLINES as u32 + 1, 0))
    );
}

/// Distinct outlines past the cap evict earlier ones; a render that asks
/// for an evicted outline gets a fresh extraction, not a stale or missing
/// path.
#[test]
fn an_evicted_glyph_outline_is_extracted_again_on_demand() {
    let mut fonts = FontStore::new();
    let mut ids = Vec::new();
    for _ in 0..4 {
        ids.push(fonts.register(CARLITO.to_vec()).expect("font"));
    }
    let mut cache = GlyphCache::default();
    let mut inserted = 0usize;
    for id in &ids {
        for gid in 0..u16::MAX {
            if cached_glyph(&mut cache, &fonts, *id, u32::from(gid)).is_err() {
                break;
            }
            inserted += 1;
        }
    }
    assert!(inserted > MAX_CACHED_GLYPH_OUTLINES);
    assert_eq!(cache.entries.len(), MAX_CACHED_GLYPH_OUTLINES);
    assert!(!cache.entries.contains_key(&(ids[0].to_u32(), 0)));
    let refilled = cached_glyph(&mut cache, &fonts, ids[0], 0).expect("re-extract after eviction");
    let upem = fonts.metrics(ids[0]).expect("metrics").units_per_em;
    assert_eq!(refilled.upem, f32::from(upem));
    assert!(cache.entries.contains_key(&(ids[0].to_u32(), 0)));
}
