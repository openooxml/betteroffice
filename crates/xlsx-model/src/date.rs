//! excel date serials: two epochs (workbookPr/@date1904) plus the deliberate
//! 1900 leap-year bug (serial 60 = phantom 1900-02-29, lotus 1-2-3 compat).

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
pub enum DateSystem {
    /// serial 1 = 1900-01-01; serial 60 = the phantom 1900-02-29.
    #[default]
    V1900,
    /// serial 0 = 1904-01-01. no leap bug.
    V1904,
}

impl DateSystem {
    /// serial to days since the unix epoch, ignoring time-of-day. None where
    /// no calendar date exists (1900 system: serial < 1 and the phantom 60).
    pub fn serial_to_unix_days(&self, serial: f64) -> Option<i64> {
        let whole = serial.floor() as i64;
        match self {
            DateSystem::V1900 => {
                if whole < 1 {
                    return None;
                }
                // serial 60 is the phantom 1900-02-29; serials past it shift back one day
                if whole == 60 {
                    return None;
                }
                let adjusted = if whole > 60 { whole - 1 } else { whole };
                // serial 1 = 1900-01-01 = -25567 unix days, so base is -25568
                Some(adjusted - 25_568)
            }
            DateSystem::V1904 => Some(whole - 24_107),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn epoch_anchors() {
        assert_eq!(DateSystem::V1900.serial_to_unix_days(1.0), Some(-25_567));
        assert_eq!(DateSystem::V1900.serial_to_unix_days(61.0), Some(-25_508));
        assert_eq!(DateSystem::V1900.serial_to_unix_days(60.0), None);
        assert_eq!(DateSystem::V1904.serial_to_unix_days(0.0), Some(-24_107));
        assert_eq!(
            DateSystem::V1900.serial_to_unix_days(43_831.0),
            Some(18_262)
        );
        assert_eq!(
            DateSystem::V1904.serial_to_unix_days(42_369.0),
            Some(18_262)
        );
    }
}
