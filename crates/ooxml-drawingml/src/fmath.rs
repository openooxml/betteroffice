#[cfg(feature = "chart")]
pub(crate) fn log10(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.log10()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::log10(x)
    }
}

pub(crate) fn atan2(y: f64, x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        y.atan2(x)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::atan2(y, x)
    }
}

pub(crate) fn cos(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.cos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::cos(x)
    }
}

pub(crate) fn hypot(x: f64, y: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.hypot(y)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        correctly_rounded_hypot(x, y)
    }
}

pub fn correctly_rounded_hypot(x: f64, y: f64) -> f64 {
    if x.is_infinite() || y.is_infinite() {
        return f64::INFINITY;
    }
    if x.is_nan() || y.is_nan() {
        return f64::NAN;
    }
    let (a, b) = if x.abs() >= y.abs() {
        (x.abs(), y.abs())
    } else {
        (y.abs(), x.abs())
    };
    if b == 0.0 || a > b * 134_217_728.0 {
        return a;
    }

    let (a, ea) = hypot_significand(a);
    let (b, eb) = hypot_significand(b);
    let squared = HypotInteger::<3>::square(a)
        .shl(2 * (ea - eb) as u32)
        .add(HypotInteger::square(b));
    let root_bits = squared.bit_len().div_ceil(2) as i32;
    let precision = (eb + root_bits + 1074).clamp(1, 53);
    let shift = 2 * (precision + 1 - root_bits);
    let mut sticky = false;
    let radicand = if shift >= 0 {
        squared.shl(shift as u32).low_u128()
    } else {
        let truncated = squared.shr((-shift) as u32);
        sticky = truncated.shl((-shift) as u32) != squared;
        truncated.low_u128()
    };
    let (root, remainder) = hypot_integer_sqrt(radicand);
    sticky |= remainder != 0;
    let round_up = root & 1 != 0 && (sticky || root & 2 != 0);
    let mut rounded = (root >> 1) + u64::from(round_up);
    let mut exponent = eb - shift / 2 + 1;
    if rounded == 1 << 53 {
        rounded >>= 1;
        exponent += 1;
    }
    if exponent == -1074 && rounded < 1 << 52 {
        return f64::from_bits(rounded);
    }
    if exponent + 1075 >= 2047 {
        return f64::INFINITY;
    }
    f64::from_bits(((exponent + 1075) as u64) << 52 | (rounded & 0x000f_ffff_ffff_ffff))
}

fn hypot_significand(value: f64) -> (u64, i32) {
    let bits = value.to_bits();
    let exponent = (bits >> 52) as i32;
    let fraction = bits & 0x000f_ffff_ffff_ffff;
    if exponent == 0 {
        let shift = fraction.leading_zeros() - 11;
        (fraction << shift, -1074 - shift as i32)
    } else {
        (fraction | (1 << 52), exponent - 1075)
    }
}

fn hypot_integer_sqrt(value: u128) -> (u64, u128) {
    let mut remainder = value;
    let mut root = 0;
    let mut place = 1u128 << ((127 - value.leading_zeros()) & !1);
    while place != 0 {
        let trial = root + place;
        if remainder >= trial {
            remainder -= trial;
            root = (root >> 1) + place;
        } else {
            root >>= 1;
        }
        place >>= 2;
    }
    (root as u64, remainder)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct HypotInteger<const N: usize>([u64; N]);

impl<const N: usize> HypotInteger<N> {
    fn square(value: u64) -> Self {
        let squared = u128::from(value) * u128::from(value);
        let mut result = Self([0; N]);
        result.0[0] = squared as u64;
        result.0[1] = (squared >> 64) as u64;
        result
    }

    fn shl(self, shift: u32) -> Self {
        let words = (shift / 64) as usize;
        let bits = shift % 64;
        let mut result = Self([0; N]);
        for (i, word) in self.0.into_iter().enumerate() {
            if i + words < N {
                result.0[i + words] |= word << bits;
            }
            if bits != 0 && i + words + 1 < N {
                result.0[i + words + 1] |= word >> (64 - bits);
            }
        }
        result
    }

    fn shr(self, shift: u32) -> Self {
        let words = (shift / 64) as usize;
        let bits = shift % 64;
        let mut result = Self([0; N]);
        for (i, word) in result.0.iter_mut().enumerate() {
            if i + words < N {
                *word = self.0[i + words] >> bits;
            }
            if bits != 0 && i + words + 1 < N {
                *word |= self.0[i + words + 1] << (64 - bits);
            }
        }
        result
    }

    fn add(self, other: Self) -> Self {
        let mut result = Self([0; N]);
        let mut carry = 0;
        for ((word, a), b) in result.0.iter_mut().zip(self.0).zip(other.0) {
            let sum = u128::from(a) + u128::from(b) + carry;
            *word = sum as u64;
            carry = sum >> 64;
        }
        result
    }

    fn bit_len(self) -> u32 {
        for (i, word) in self.0.into_iter().enumerate().rev() {
            if word != 0 {
                return i as u32 * 64 + 64 - word.leading_zeros();
            }
        }
        0
    }

    fn low_u128(self) -> u128 {
        u128::from(self.0[0]) | (u128::from(self.0[1]) << 64)
    }
}

pub(crate) fn pow(x: f64, y: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.powf(y)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::pow(x, y)
    }
}

pub(crate) fn sin(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.sin()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::sin(x)
    }
}

pub(crate) fn tan(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.tan()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::tan(x)
    }
}

#[cfg(feature = "chart")]
pub(crate) fn log(x: f64, base: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.log(base)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::log(x) / libm::log(base)
    }
}

pub(crate) fn sin_cos(x: f64) -> (f64, f64) {
    #[cfg(target_arch = "wasm32")]
    {
        x.sin_cos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        (sin(x), cos(x))
    }
}

#[cfg(test)]
pub(crate) fn sin_cosf(x: f32) -> (f32, f32) {
    #[cfg(target_arch = "wasm32")]
    {
        x.sin_cos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        (libm::sinf(x), libm::cosf(x))
    }
}

#[cfg(feature = "chart")]
pub(crate) fn powi(a: f64, b: i32) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        a.powi(b)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        let (mut a, mut b) = (a, b);
        let recip = b < 0;
        let mut r = 1.0;
        loop {
            if b & 1 != 0 {
                r *= a;
            }
            b /= 2;
            if b == 0 {
                break;
            }
            a *= a;
        }
        if recip { 1.0 / r } else { r }
    }
}

#[cfg(all(test, not(target_arch = "wasm32")))]
mod hypot_tests {
    use super::{HypotInteger, correctly_rounded_hypot, hypot};
    use std::cmp::Ordering;

    const WASM_CASES: &[(u64, u64, u64)] = &[
        (0x3fb205d63c782980, 0x3fdfdcfa1040aa9a, 0x3fe017104644c02e),
        (0x3fb60ea4cfa11cb0, 0xbfdd1000a481232c, 0x3fdd94bdfee609f2),
        (0x3fc56dc95b458960, 0xbfdb00bad71fc174, 0x3fdd0d05a710fb8a),
        (0x3fd87f930ba8d0cc, 0x3fb8aa1d3fd617e8, 0x3fd9432db2a03922),
        (0x3fe0634bd57b776e, 0x3fd0d2eb77a1cf18, 0x3fe26bb531ae4221),
        (0x3fe0b88d1960f3b4, 0xbfd70570b264409c, 0x3fe44cc1a773da5f),
        (0x3fe4ce8a84489e3e, 0x3fc84008ca501dde, 0x3fe5ac0ac3760a52),
        (0x3fe4cea0c114eeb2, 0x3fc84008ca501dde, 0x3fe5ac201cfb823a),
        (0x3feb25da34eadc6c, 0xbfc24acfeb8adeb4, 0x3feb87c3e0c5938b),
        (0x3ff042662eca4053, 0x3fd910bdd942f802, 0x3ff16cd3b6073f67),
        (0x3ff5d7b9a8ef16fa, 0xbfe9cc6203939b98, 0x3ff95dfc84ac5b72),
        (0x3ffa849fa979e71b, 0xbfe39d590fb26010, 0x3ffc460470f44dd8),
        (0x40b6baef0157315b, 0xc0e352cf947f7cf8, 0x40e3880000000000),
        (0x40b6baef04fbd633, 0xc0e352cf93327c84, 0x40e387fffec78000),
        (0x40bb8027f7627340, 0x40e84f478584ce8e, 0x40e88d3131fd230d),
        (0x40bf8343390a64b1, 0x40e32141761da159, 0x40e387fffec78000),
        (0x40bf83434005f03d, 0x40e32141772ead3b, 0x40e3880000000000),
        (0x40c0d40c124a61a0, 0xc0e62c4574e34dd0, 0x40e6918b2b8975f4),
        (0x40c17d597e0de252, 0xc0e08562fd7fab7d, 0x40e1170015b3c78f),
        (0x40c17d598590ae67, 0xc0e08562f7f369e0, 0x40e1170010d1c7ae),
        (0x40d05952ab2934c0, 0xc0e499fe908b0028, 0x40e629ff6338ef1f),
        (0x40d28ed1f62aa9b4, 0x40e12fb523bdbc0d, 0x40e387fffec78000),
        (0x40d28ed1f6335992, 0x40e12fb5251e86d9, 0x40e3880000000000),
        (0x40d47d1b6309327d, 0xc0e0a104ee6c77ec, 0x40e387fffec78000),
        (0x40d47d1b6567e868, 0xc0e0a104ef20a045, 0x40e3880000000000),
        (0x40d83e2d9d97bc1a, 0xc0d8182395df4bcd, 0x40e11700a1e860a0),
        (0x40d83e2da9146485, 0xc0d81823982af553, 0x40e11700a6ca60de),
        (0x40dd31e34c2f311a, 0xc0d9f3de8a0aa6b6, 0x40e387fffec78000),
        (0x40dd31e34d28add0, 0xc0d9f3de8c9eb667, 0x40e3880000000000),
        (0x40dea0c2d0fa3de2, 0x40d83ea59d8901e9, 0x40e387fffec78001),
        (0x40dea0c2d3af9925, 0x40d83ea59e0c15fb, 0x40e3880000000000),
        (0x40e07e6dde62d1d0, 0xc0c1e54779d2cf9d, 0x40e117012ff04f27),
        (0x40e07e6de240729f, 0xc0c1e5478b6c8499, 0x40e1170134d24f08),
        (0x40e227bf491c99ba, 0x40cccdf311b7ee40, 0x40e387fffec77fff),
        (0x40e227bf4a02b2c0, 0x40cccdf315e5c7e9, 0x40e3880000000000),
        (0x40e2895a0d38e525, 0xc0c89d48c4613808, 0x40e387fffec78000),
        (0x40e2895a0e9514f3, 0xc0c89d48c37d481e, 0x40e3880000000000),
        (0x40e2b0d4caee42ea, 0x40c2d149541b8108, 0x40e34610b7530e8c),
        (0x40e34610b7530e8c, 0x0000000000000000, 0x40e34610b7530e8c),
        (0x40e385f6ac99ee1a, 0xc091d606aa2884b0, 0x40e387fffec77fff),
        (0x40e385f6adcda0fb, 0xc091d606bfbcd2ae, 0x40e3880000000000),
        (0x40e629ff6338ef1f, 0x0000000000000000, 0x40e629ff6338ef1f),
        (0x40e6918b2b8975f4, 0x0000000000000000, 0x40e6918b2b8975f4),
        (0x40e84dd1fb56089c, 0x0000000000000000, 0x40e84dd1fb56089c),
        (0x40e88d3131fd230d, 0x0000000000000000, 0x40e88d3131fd230d),
        (0x40e90183b82669ba, 0x40d9abd6a2856c22, 0x40ec1b997e089997),
        (0x40e9839a5d3979a5, 0xc0e190569be654fc, 0x40eef99f15265fe1),
        (0x40ec1b997e089997, 0x0000000000000000, 0x40ec1b997e089997),
        (0x40eef99f15265fe0, 0x0000000000000000, 0x40eef99f15265fe0),
        (0x40efbfa82874c2dc, 0x40d2805ac0e889d7, 0x40f088d1fd0669b1),
        (0x40efbfca16fa1254, 0x40d2805ac0e889d7, 0x40f088e246e313bd),
        (0x40f088d1fd0669b1, 0x0000000000000000, 0x40f088d1fd0669b1),
        (0x40f088e246e313bd, 0x0000000000000000, 0x40f088e246e313bd),
        (0x40f4199f8043923a, 0x0000000000000000, 0x40f4199f8043923a),
        (0x40f4b650fd996ed5, 0xc0cbe96785779b60, 0x40f5010494c2df3f),
        (0x40f5010494c2df3f, 0x0000000000000000, 0x40f5010494c2df3f),
        (0x40f6494a6c40bae7, 0x0000000000000000, 0x40f6494a6c40bae7),
        (0x40f793f5778d70ab, 0x0000000000000000, 0x40f793f5778d70ab),
        (0x40f8cf51392fb629, 0x40e31f95e644f61b, 0x40fa96ae9b53d1a3),
        (0x40fa96ae9b53d1a3, 0x0000000000000000, 0x40fa96ae9b53d1a3),
        (0x40fda3d39b0b39d5, 0x0000000000000000, 0x40fda3d39b0b39d5),
        (0x4100aa25ced1b4eb, 0xc0f3aebe87761441, 0x41035a84b2133d8a),
        (0x4101f3793cbb5291, 0x0000000000000000, 0x4101f3793cbb5291),
        (0x41035a84b2133d89, 0x0000000000000000, 0x41035a84b2133d89),
        (0x41043b4f13a4580a, 0xc0ededf7f7789df4, 0x4105922b45a1b26c),
        (0x4105922b45a1b26c, 0x0000000000000000, 0x4105922b45a1b26c),
        (0xbfd7686dcf380f4c, 0x3feac38d312daf74, 0x3fed3609da02ab91),
        (0xbfdcd6663eef676c, 0xbfcb11b818b5f774, 0x3fdfdb1071a6f1ab),
        (0xbfe5bc62aebd26a8, 0x3fddc6780c7c821e, 0x3fea5878ab232667),
        (0xbfebdc52d2d0b944, 0xbfdabf32d0745e4e, 0x3feee773ade76e8f),
        (0xbff28838e848b565, 0x3fd748c67bc0658d, 0x3ff36cc210f2a13b),
        (0xbff5b4684704d960, 0x3fe22ac6f2c5d09e, 0x3ff787610585f1f5),
        (0xc0e1dbdc44ed7cc0, 0x40f46b51b1ec7486, 0x40f6494a6c40bae7),
        (0xc0e60052e1dc2eee, 0xc0d4a6f4c2b98260, 0x40e84dd1fb56089c),
        (0xc0f09549fe677e8c, 0x40e6b77b60c750fb, 0x40f4199f8043923b),
        (0xc0f54187f5f0be0c, 0xc0e467ff7948e918, 0x40f793f5778d70ab),
        (0xc0fc471bd14743fa, 0x40e1c3b5eda1b6d2, 0x40fda3d39b0b39d5),
        (0xc1008f33d45b67e1, 0x40ebb885e69de3a5, 0x4101f3793cbb5291),
    ];

    fn exact_square(bits: u64) -> HypotInteger<68> {
        let exponent = ((bits >> 52) & 0x7ff) as u32;
        let fraction = bits & 0x000f_ffff_ffff_ffff;
        if exponent == 0 {
            HypotInteger::square(fraction)
        } else {
            HypotInteger::square(fraction | (1 << 52)).shl(2 * (exponent - 1))
        }
    }

    fn reference_hypot(x: f64, y: f64) -> f64 {
        if x.is_infinite() || y.is_infinite() {
            return f64::INFINITY;
        }
        if x.is_nan() || y.is_nan() {
            return f64::NAN;
        }
        let squared = exact_square(x.to_bits()).add(exact_square(y.to_bits()));
        let mut low = 0;
        let mut high = f64::INFINITY.to_bits();
        while high - low > 1 {
            let mid = low + (high - low) / 2;
            if exact_square(mid).0.iter().rev().cmp(squared.0.iter().rev()) != Ordering::Greater {
                low = mid;
            } else {
                high = mid;
            }
        }
        let exponent = ((low >> 52) & 0x7ff) as u32;
        let fraction = low & 0x000f_ffff_ffff_ffff;
        let (significand, shift) = if exponent == 0 {
            (fraction, 0)
        } else {
            (fraction | (1 << 52), 2 * (exponent - 1))
        };
        let midpoint_squared = HypotInteger::<68>::square(2 * significand + 1).shl(shift);
        let rounded = match squared
            .shl(2)
            .0
            .iter()
            .rev()
            .cmp(midpoint_squared.0.iter().rev())
        {
            Ordering::Less => low,
            Ordering::Greater => high,
            Ordering::Equal => {
                if low & 1 == 0 {
                    low
                } else {
                    high
                }
            }
        };
        f64::from_bits(rounded)
    }

    fn assert_matches_reference(x: f64, y: f64) {
        let expected = reference_hypot(x, y).to_bits();
        assert_eq!(
            correctly_rounded_hypot(x, y).to_bits(),
            expected,
            "x={:#018x}, y={:#018x}",
            x.to_bits(),
            y.to_bits(),
        );
    }

    #[test]
    fn hypot_matches_wasm_corpus() {
        for &(x, y, expected) in WASM_CASES {
            assert_eq!(
                hypot(f64::from_bits(x), f64::from_bits(y)).to_bits(),
                expected,
                "x={x:#018x}, y={y:#018x}",
            );
        }
    }

    #[test]
    fn hypot_handles_special_values_and_boundaries() {
        for (x, y, expected) in [
            (3.0, 4.0, 5.0),
            (5.0, 12.0, 13.0),
            (0.0, 0.0, 0.0),
            (-0.0, -0.0, 0.0),
            (-3.0, 4.0, 5.0),
            (f64::INFINITY, f64::NAN, f64::INFINITY),
            (f64::NAN, f64::NEG_INFINITY, f64::INFINITY),
            (f64::MAX, f64::MAX, f64::INFINITY),
        ] {
            assert_eq!(hypot(x, y).to_bits(), expected.to_bits());
        }
        assert!(hypot(f64::NAN, 1.0).is_nan());
        assert!(hypot(1.0, f64::NAN).is_nan());
        for (x, y) in [
            (f64::MIN_POSITIVE * 0.5, f64::MIN_POSITIVE * 0.25),
            (f64::from_bits(1), f64::from_bits(1)),
            (f64::from_bits(3), f64::from_bits(4)),
            (f64::MIN_POSITIVE, f64::from_bits((1 << 52) - 1)),
            (134_217_728.0, 1.0),
            (f64::from_bits(134_217_728.0f64.to_bits() + 1), 1.0),
            (f64::MAX, f64::from_bits(1993 << 52)),
        ] {
            assert_matches_reference(x, y);
            assert_matches_reference(y, x);
        }
    }

    #[test]
    fn hypot_rounds_halfway_to_even() {
        let m = (1u64 << 25) + 1;
        let n = 1u64 << 25;
        for multiplier in [5, 7] {
            let x = multiplier * (m * m - n * n);
            let y = multiplier * 2 * m * n;
            let norm = multiplier * (m * m + n * n);
            let expected = ((norm + 1) & !3) as f64;
            assert_eq!(hypot(x as f64, y as f64).to_bits(), expected.to_bits());
            assert_matches_reference(x as f64, y as f64);
        }
    }

    #[test]
    fn hypot_matches_exact_reference_sweep() {
        let mut state = 0x6a09_e667_f3bc_c909u64;
        let mut next = || {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state
        };
        for i in 0..4096 {
            let mut x = next();
            let mut y = next();
            let (ex, ey) = match i % 4 {
                0 => (1024 + next() % 1023, next() % 512),
                1 => {
                    let exponent = next() % 2047;
                    (exponent, exponent.saturating_sub(next() % 28))
                }
                2 => (next() % 2, next() % 2),
                _ => (next() % 2047, next() % 2047),
            };
            x = (x & 0x800f_ffff_ffff_ffff) | (ex << 52);
            y = (y & 0x800f_ffff_ffff_ffff) | (ey << 52);
            assert_matches_reference(f64::from_bits(x), f64::from_bits(y));
            assert_matches_reference(f64::from_bits(y), f64::from_bits(x));
        }
    }
}
