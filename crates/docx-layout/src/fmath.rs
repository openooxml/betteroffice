pub(crate) fn sin_cos(x: f64) -> (f64, f64) {
    #[cfg(target_arch = "wasm32")]
    {
        x.sin_cos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        (libm::sin(x), libm::cos(x))
    }
}

#[cfg(all(test, not(target_arch = "wasm32")))]
mod tests {
    use super::*;

    #[test]
    fn math_matches_libm_bits() {
        for x in [-3.0, -0.0, 0.0, 0.37, 1.37, 7.0] {
            let (sin, cos) = sin_cos(x);
            assert_eq!(sin.to_bits(), libm::sin(x).to_bits());
            assert_eq!(cos.to_bits(), libm::cos(x).to_bits());
        }
    }
}
