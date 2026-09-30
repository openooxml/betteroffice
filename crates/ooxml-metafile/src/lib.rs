//! Bounded replay of EMF, EMF+ and WMF metafiles.
//!
//! [`decode`] is the strict vector replay PowerPoint pictures draw from: it
//! refuses any metafile holding something it cannot represent. [`replay`]
//! draws everything it can, text and bitmaps included, and lists what it drew
//! without; with the `svg` feature [`to_svg`] serializes that drawing.

mod blit;
mod dib;
pub mod drawing;
mod emf;
mod emfplus;
mod gradient;
mod player;
mod read;
mod shapes;
#[cfg(feature = "svg")]
mod svg;
mod text;
mod transform;
mod wmf;

#[cfg(test)]
#[path = "../tests/common/mod.rs"]
mod test_records;

pub use drawing::{Drawing, Omission, Refusal};
pub use shapes::{MetafileDrawing, MetafileOp, MetafileStroke, decode};
#[cfg(feature = "svg")]
pub use svg::{
    MAX_SVG_BYTES, Svg, placeholder_svg, to_svg, to_svg_with_budget, to_svg_with_limits,
};

use std::cell::Cell;
use std::rc::Rc;

use player::{Player, SharedBudget};
use read::{i32_at, u16_at, u32_at};

/// Whether `bytes` start like an EMF or a WMF.
pub fn is_metafile(bytes: &[u8]) -> bool {
    emf::is_emf(bytes) || wmf::is_wmf(bytes)
}

/// Whether `bytes` start like a WMF rather than an EMF.
pub fn is_wmf(bytes: &[u8]) -> bool {
    wmf::is_wmf(bytes)
}

/// The size in CSS pixels a metafile's header gives its picture.
pub fn picture_size(bytes: &[u8]) -> Option<(f64, f64)> {
    if emf::is_emf(bytes) {
        return Some(emf_size(bytes, emf::emf_frame(bytes)?));
    }
    if u32_at(bytes, 0) != Some(wmf::WMF_PLACEABLE_KEY) {
        return None;
    }
    let side = |at| read::i16_at(bytes, at).map(f64::from);
    let inch = f64::from(u16_at(bytes, 14).filter(|inch| *inch > 0)?);
    let (width, height) = (side(10)? - side(6)?, side(12)? - side(8)?);
    (width != 0.0 && height != 0.0).then(|| fit(width / inch * 96.0, height / inch * 96.0))
}

/// Remaining cumulative replay work and bitmap pixels.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReplayBudget {
    pub work: u64,
    pub pixels: u64,
}

impl Default for ReplayBudget {
    fn default() -> Self {
        Self {
            work: (player::FULL_LIMITS.records
                + player::FULL_LIMITS.ops
                + player::FULL_LIMITS.commands) as u64,
            pixels: player::FULL_LIMITS.bitmap_pixels,
        }
    }
}

/// Replays every record of an EMF, EMF+ or WMF metafile it can reproduce.
pub fn replay(bytes: &[u8]) -> Result<Drawing, Refusal> {
    replay_with_budget(bytes, &mut ReplayBudget::default())
}

/// Replays with cumulative allowances, spending them even on refusal.
pub fn replay_with_budget(bytes: &[u8], budget: &mut ReplayBudget) -> Result<Drawing, Refusal> {
    with_budget(budget, |shared| {
        play_nested(bytes, 0, shared).map_err(Refusal)
    })
}

fn with_budget<T>(
    budget: &mut ReplayBudget,
    replay: impl FnOnce(Rc<SharedBudget>) -> Result<T, Refusal>,
) -> Result<T, Refusal> {
    let limits = ReplayBudget::default();
    let allowance = ReplayBudget {
        work: budget.work.min(limits.work),
        pixels: budget.pixels.min(limits.pixels),
    };
    let shared = Rc::new(SharedBudget {
        remaining: Cell::new(allowance),
        exceeded: Cell::new(false),
    });
    let result = replay(Rc::clone(&shared));
    let remaining = shared.remaining.get();
    budget.work -= allowance.work - remaining.work;
    budget.pixels -= allowance.pixels - remaining.pixels;
    if shared.exceeded.get() {
        return Err(Refusal(
            "the metafile draws more than the replay limits".to_owned(),
        ));
    }
    result
}

pub(crate) fn play_nested(
    bytes: &[u8],
    depth: usize,
    budget: Rc<SharedBudget>,
) -> Result<Drawing, String> {
    let player = if emf::is_emf(bytes) {
        play_full_emf(bytes, depth, budget)
    } else if wmf::is_wmf(bytes) {
        play_wmf::<true>(bytes, depth, Some(budget))
    } else {
        Err("the bytes are not an EMF or WMF metafile".to_owned())
    }?;
    Ok(Drawing {
        width: player.unit.0,
        height: player.unit.1,
        omissions: player
            .omissions
            .iter()
            .map(|(what, count)| Omission {
                what,
                count: *count,
            })
            .collect(),
        ops: player.ops,
    })
}

/// An EMF with EMF+ records plays them, as Office does, when they draw
/// without omissions; otherwise, and for plain EMF, its GDI records play.
fn play_full_emf(
    bytes: &[u8],
    depth: usize,
    budget: Rc<SharedBudget>,
) -> Result<Player<true>, String> {
    if !has_plus_header(bytes, &budget) {
        return play_emf::<true>(bytes, depth, false, Some(budget));
    }
    let plus = play_emf::<true>(bytes, depth, true, Some(Rc::clone(&budget)));
    if budget.exceeded.get() {
        return Err("the metafile draws more than the replay limits".to_owned());
    }
    if let Ok(player) = &plus
        && player.omissions.is_empty()
        && !player.ops.is_empty()
    {
        return plus;
    }
    let gdi = play_emf::<true>(bytes, depth, false, Some(Rc::clone(&budget)));
    if budget.exceeded.get() {
        return Err("the metafile draws more than the replay limits".to_owned());
    }
    match gdi {
        Ok(gdi) if !gdi.ops.is_empty() => Ok(gdi),
        gdi => plus.or(gdi),
    }
}

/// Whether an EMF+ header comment comes before the EMF's first drawing record.
fn has_plus_header(bytes: &[u8], budget: &SharedBudget) -> bool {
    let mut offset = 0usize;
    for index in 0..player::FULL_LIMITS.records {
        if !budget.spend(1, 0) {
            return false;
        }
        let (Some(kind), Some(size)) = (
            u32_at(bytes, offset),
            u32_at(bytes, offset.saturating_add(4)),
        ) else {
            return false;
        };
        if size < 8 {
            return false;
        }
        match kind {
            emf::EMF_COMMENT if size >= 20 && u32_at(bytes, offset + 12) == Some(emf::EMF_PLUS) => {
                return u16_at(bytes, offset + 16) == Some(0x4001);
            }
            emf::EMF_COMMENT => {}
            1 if index == 0 => {}
            _ => return false,
        }
        offset = offset.saturating_add(size as usize);
    }
    false
}

/// Largest side of a replayed drawing, in CSS pixels.
const MAX_SIDE: f64 = 16_384.0;

/// A picture size in CSS pixels, scaled down to fit [`MAX_SIDE`].
fn fit(width: f64, height: f64) -> (f64, f64) {
    let (width, height) = (width.abs(), height.abs());
    let scale = (MAX_SIDE / width.max(height)).min(1.0);
    ((width * scale).max(1.0), (height * scale).max(1.0))
}

/// The CSS pixel size of an EMF's `rclFrame`, or of its device bounds.
fn emf_size(bytes: &[u8], frame: (f64, f64, f64, f64)) -> (f64, f64) {
    let read = |at| i32_at(bytes, at).map(f64::from).unwrap_or(0.0);
    let (width, height) = (read(32) - read(24), read(36) - read(28));
    if width > 0.0 && height > 0.0 {
        return fit(width / 2540.0 * 96.0, height / 2540.0 * 96.0);
    }
    let per_mm = emf::device_per_mm(bytes);
    fit(
        frame.2 / per_mm.0 * 96.0 / 25.4,
        frame.3 / per_mm.1 * 96.0 / 25.4,
    )
}

pub(crate) const MAX_NESTING: usize = 2;

pub(crate) fn play_emf<const FULL: bool>(
    bytes: &[u8],
    depth: usize,
    prefer_plus: bool,
    budget: Option<Rc<SharedBudget>>,
) -> Result<Player<FULL>, String> {
    let malformed = || "the EMF header is malformed".to_owned();
    if u32_at(bytes, 0) != Some(1)
        || u32_at(bytes, 40) != Some(emf::EMF_SIGNATURE)
        || u32_at(bytes, 4).is_none_or(|size| size < 88)
    {
        return Err(malformed());
    }
    let handles = u16_at(bytes, 56).ok_or_else(malformed)? as usize;
    let frame = emf::emf_frame(bytes).ok_or("the EMF frame is empty")?;
    let unit = if FULL {
        emf_size(bytes, frame)
    } else {
        (1.0, 1.0)
    };
    let mut player = Player::<FULL>::new(frame, handles + 1, unit);
    player.budget = budget;
    player.depth = depth;
    player.prefer_plus = prefer_plus;
    if FULL {
        player.device_per_mm = emf::device_per_mm(bytes);
        player.gdi_records =
            has_gdi_records(bytes, player.limits.records, player.budget.as_deref());
    }
    let mut offset = 0usize;
    for index in 0..player.limits.records {
        player
            .spend(1, 0)
            .ok_or_else(|| "the metafile draws more than the replay limits".to_owned())?;
        let (Some(kind), Some(size)) = (
            u32_at(bytes, offset),
            u32_at(bytes, offset.saturating_add(4)),
        ) else {
            return Err("the EMF ends before its EOF record".to_owned());
        };
        let size = size as usize;
        if size < 8 || !size.is_multiple_of(4) {
            return Err(format!("EMF record {} has an invalid size", index + 1));
        }
        let Some(record) = offset
            .checked_add(size)
            .and_then(|end| bytes.get(offset..end))
        else {
            return Err(format!(
                "EMF record {} runs past the end of the file",
                index + 1
            ));
        };
        if kind == emf::EMF_EOF {
            if size < 20 {
                return Err("the EMF EOF record is malformed".to_owned());
            }
            player.flush_pending();
            if player.overflowed {
                return Err("the metafile draws more than the replay limits".to_owned());
            }
            return Ok(player);
        }
        let skipped = FULL && player.plus_only && !player.plus_gdi && kind != emf::EMF_COMMENT;
        if !skipped && emf::emf_record(&mut player, record, kind, 8).is_none() {
            return Err(player.refusal.take().unwrap_or_else(|| {
                format!(
                    "EMF record {} (type {kind}) could not be replayed",
                    index + 1
                )
            }));
        }
        if player.overflowed {
            return Err("the metafile draws more than the replay limits".to_owned());
        }
        offset += size;
    }
    Err("the EMF holds more records than the replay limit".to_owned())
}

/// Whether an EMF holds records besides its header, comments and EOF: a
/// dual EMF+ metafile's GDI rendition, or plain GDI content.
fn has_gdi_records(bytes: &[u8], limit: usize, budget: Option<&SharedBudget>) -> bool {
    let mut offset = 0usize;
    for _ in 0..limit {
        if budget.is_some_and(|budget| !budget.spend(1, 0)) {
            return false;
        }
        let (Some(kind), Some(size)) = (
            u32_at(bytes, offset),
            u32_at(bytes, offset.saturating_add(4)),
        ) else {
            return false;
        };
        match kind {
            emf::EMF_EOF => return false,
            1 | emf::EMF_COMMENT => {}
            _ => return true,
        }
        if size < 8 {
            return false;
        }
        offset = offset.saturating_add(size as usize);
    }
    false
}

pub(crate) fn play_wmf<const FULL: bool>(
    bytes: &[u8],
    depth: usize,
    budget: Option<Rc<SharedBudget>>,
) -> Result<Player<FULL>, String> {
    let limits = if FULL {
        player::FULL_LIMITS
    } else {
        player::SHAPES_LIMITS
    };
    let records = wmf::wmf_records(bytes, limits.records, budget.as_deref()).ok_or_else(|| {
        if budget.as_ref().is_some_and(|budget| budget.exceeded.get()) {
            "the metafile draws more than the replay limits"
        } else {
            "the WMF is malformed"
        }
        .to_owned()
    })?;
    if FULL {
        if let Some(emf) = wmf::embedded_emf(bytes, &records.records)
            && let Ok(player) = play_emf::<FULL>(&emf, depth, false, budget.clone())
        {
            return Ok(player);
        }
        if budget.as_ref().is_some_and(|budget| budget.exceeded.get()) {
            return Err("the metafile draws more than the replay limits".to_owned());
        }
        if !records.framed {
            return Err("the WMF has neither a placeable header nor a window extent".to_owned());
        }
    }
    let unit = if FULL {
        let (width, height) = records.size.unwrap_or_else(|| {
            let (width, height) = (records.frame.2.abs(), records.frame.3.abs());
            let scale = 480.0 / width.max(height);
            (width * scale, height * scale)
        });
        fit(width, height)
    } else {
        (1.0, 1.0)
    };
    let mut player = Player::<FULL>::new(records.frame, records.handles, unit);
    player.budget = budget;
    player.depth = depth;
    player.pixel = FULL.then_some(1.0);
    match wmf::play_wmf(&mut player, bytes, &records.records) {
        Some(()) => Ok(player),
        None => Err(player
            .refusal
            .take()
            .unwrap_or_else(|| "the WMF could not be replayed".to_owned())),
    }
}

#[cfg(feature = "fuzzing")]
pub mod fuzzing {
    pub use crate::{decode, replay};
}
