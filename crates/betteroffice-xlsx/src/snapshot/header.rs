use xlsx_model::{CellRef, SheetId};

use crate::{CalculationOptions, CalculationResult, CellAddress};

use super::wire::{ChunkKind, Reader, Writer};
use super::{SnapshotError, SnapshotResult};

pub(crate) const NONCE_BYTES: usize = 32;
pub(crate) const GUID_BYTES: usize = 36;

pub(crate) fn identity_byte(byte: u8, index: usize, guid: bool) -> bool {
    if guid && matches!(index, 8 | 13 | 18 | 23) {
        byte == b'-'
    } else if index == if guid { 14 } else { 12 } {
        byte == b'4'
    } else if index == if guid { 19 } else { 16 } {
        matches!(byte, b'8' | b'9' | b'a' | b'b')
    } else {
        byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SnapshotMode {
    Standalone,
}

#[derive(Debug, Clone)]
pub(crate) struct SnapshotHeader {
    pub(crate) snapshot_id: u64,
    pub(crate) mode: SnapshotMode,
    pub(crate) edited_since_open: bool,
    pub(crate) recalculated_since_open: bool,
    pub(crate) moved_references_since_open: bool,
    pub(crate) active_sheet: SheetId,
    pub(crate) rand_seed: Option<u32>,
    pub(crate) model_epoch: u64,
    pub(crate) version_nonce: String,
    pub(crate) committed_changes: u64,
    pub(crate) last_calculation: CalculationResult,
    pub(crate) calculation_context: Option<CalculationOptions>,
    pub(crate) client_id: u64,
    pub(crate) guid: String,
    pub(crate) next_sheet_id: u64,
    pub(crate) state_vector: Vec<u8>,
    pub(crate) chunk_counts: [u64; 9],
}

impl SnapshotHeader {
    pub(crate) fn encode(&self) -> Vec<u8> {
        let Self {
            snapshot_id,
            mode,
            edited_since_open,
            recalculated_since_open,
            moved_references_since_open,
            active_sheet,
            rand_seed,
            model_epoch,
            version_nonce,
            committed_changes,
            last_calculation,
            calculation_context,
            client_id,
            guid,
            next_sheet_id,
            state_vector,
            chunk_counts,
        } = self;
        let mut w = Writer::new();
        w.var_u64(*snapshot_id);
        match mode {
            SnapshotMode::Standalone => w.u8(0),
        }
        w.bool(*edited_since_open);
        w.bool(*recalculated_since_open);
        w.bool(*moved_references_since_open);
        let SheetId(active_sheet) = active_sheet;
        w.var_u32(*active_sheet);
        w.option(*rand_seed, Writer::var_u32);
        w.var_u64(*model_epoch);
        w.str(version_nonce);
        w.var_u64(*committed_changes);
        let CalculationResult {
            changed,
            cycle_cells,
            limited_cells,
        } = last_calculation;
        for addresses in [changed, cycle_cells, limited_cells] {
            w.var_usize(addresses.len());
            for address in addresses {
                let CellAddress { sheet, cell } = address;
                let SheetId(sheet) = sheet;
                let CellRef {
                    row,
                    col,
                    abs_row,
                    abs_col,
                } = cell;
                w.var_u32(*sheet);
                w.var_u32(*row);
                w.var_u32(*col);
                w.bool(*abs_row);
                w.bool(*abs_col);
            }
        }
        w.option(*calculation_context, |w, options| {
            let CalculationOptions { now_serial } = options;
            w.option(now_serial, Writer::f64);
        });
        w.var_u64(*client_id);
        w.str(guid);
        w.var_u64(*next_sheet_id);
        w.bytes(state_vector);
        for count in chunk_counts {
            w.var_u64(*count);
        }
        w.into_bytes()
    }

    pub(crate) fn decode(payload: &[u8]) -> SnapshotResult<Self> {
        let mut r = Reader::new(payload);
        let snapshot_id = r.var_u64()?;
        let mode = match r.u8()? {
            0 => SnapshotMode::Standalone,
            _ => {
                return Err(SnapshotError::new(
                    "only standalone snapshots are supported",
                ));
            }
        };
        let edited_since_open = r.bool()?;
        let recalculated_since_open = r.bool()?;
        let moved_references_since_open = r.bool()?;
        let active_sheet = SheetId(r.var_u32()?);
        let rand_seed = r.option(Reader::var_u32)?;
        let model_epoch = r.var_u64()?;
        let version_nonce = r.str()?.to_owned();
        let committed_changes = r.var_u64()?;
        let mut addresses = || {
            let count = r.var_usize()?;
            let mut addresses = Vec::new();
            for _ in 0..count {
                addresses.push(CellAddress {
                    sheet: SheetId(r.var_u32()?),
                    cell: CellRef {
                        row: r.var_u32()?,
                        col: r.var_u32()?,
                        abs_row: r.bool()?,
                        abs_col: r.bool()?,
                    },
                });
            }
            Ok::<_, SnapshotError>(addresses)
        };
        let last_calculation = CalculationResult {
            changed: addresses()?,
            cycle_cells: addresses()?,
            limited_cells: addresses()?,
        };
        let calculation_context = r.option(|r| {
            Ok(CalculationOptions {
                now_serial: r.option(Reader::f64)?,
            })
        })?;
        let client_id = r.var_u64()?;
        let guid = r.str()?.to_owned();
        let next_sheet_id = r.var_u64()?;
        let state_vector = r.bytes()?.to_vec();
        let mut chunk_counts = [0; 9];
        for count in &mut chunk_counts {
            *count = r.var_u64()?;
        }
        r.finish()?;
        Ok(Self {
            snapshot_id,
            mode,
            edited_since_open,
            recalculated_since_open,
            moved_references_since_open,
            active_sheet,
            rand_seed,
            model_epoch,
            version_nonce,
            committed_changes,
            last_calculation,
            calculation_context,
            client_id,
            guid,
            next_sheet_id,
            state_vector,
            chunk_counts,
        })
    }

    pub(crate) fn chunk_count(&self, kind: ChunkKind) -> u64 {
        self.chunk_counts[match kind {
            ChunkKind::Header => 0,
            ChunkKind::AuthorityBase => 1,
            ChunkKind::Yrs => 2,
            ChunkKind::Model => 3,
            ChunkKind::Cells => 4,
            ChunkKind::Preserved => 5,
            ChunkKind::Facts => 6,
            ChunkKind::Source => 7,
            ChunkKind::End => 8,
        }]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn header_preserves_nested_options_and_float_bits() {
        let mut header = SnapshotHeader {
            snapshot_id: u64::MAX,
            mode: SnapshotMode::Standalone,
            edited_since_open: true,
            recalculated_since_open: false,
            moved_references_since_open: true,
            active_sheet: SheetId(7),
            rand_seed: Some(0),
            model_epoch: u64::MAX,
            version_nonce: String::new(),
            committed_changes: 42,
            last_calculation: CalculationResult {
                changed: vec![
                    CellAddress {
                        sheet: SheetId(2),
                        cell: CellRef::parse_a1("$C$4").unwrap(),
                    },
                    CellAddress {
                        sheet: SheetId(0),
                        cell: CellRef::new(0, 0),
                    },
                ],
                cycle_cells: vec![],
                limited_cells: vec![CellAddress {
                    sheet: SheetId(9),
                    cell: CellRef::parse_a1("D$5").unwrap(),
                }],
            },
            calculation_context: None,
            client_id: 123,
            guid: "snapshot-guid".into(),
            next_sheet_id: u64::MAX,
            state_vector: vec![0, 128, 255],
            chunk_counts: [1, 2, 3, 4, 5, 6, 7, 8, 1],
        };
        for context in [
            None,
            Some(CalculationOptions { now_serial: None }),
            Some(CalculationOptions {
                now_serial: Some(-0.0),
            }),
            Some(CalculationOptions {
                now_serial: Some(f64::from_bits(0x7ff8_0000_0000_0042)),
            }),
        ] {
            header.calculation_context = context;
            for seed in [None, Some(0), Some(u32::MAX)] {
                header.rand_seed = seed;
                let bytes = header.encode();
                let decoded = SnapshotHeader::decode(&bytes).unwrap();
                assert_eq!(decoded.encode(), bytes);
                assert_eq!(decoded.last_calculation, header.last_calculation);
                assert_eq!(decoded.rand_seed, seed);
                assert_eq!(decoded.chunk_count(ChunkKind::Yrs), 3);
                assert_eq!(
                    decoded
                        .calculation_context
                        .map(|options| options.now_serial.map(f64::to_bits)),
                    context.map(|options| options.now_serial.map(f64::to_bits)),
                );
            }
        }
        let mut bytes = header.encode();
        let mut r = Reader::new(&bytes);
        r.var_u64().unwrap();
        let mode_offset = bytes.len() - r.rest().len();
        bytes[mode_offset] = 1;
        assert!(SnapshotHeader::decode(&bytes).is_err());
    }
}
