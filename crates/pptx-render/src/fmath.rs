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
            assert_eq!(sin(x).to_bits(), libm::sin(x).to_bits());
            assert_eq!(cos(x).to_bits(), libm::cos(x).to_bits());
            let (sin, cos) = sin_cos(x);
            assert_eq!(sin.to_bits(), libm::sin(x).to_bits());
            assert_eq!(cos.to_bits(), libm::cos(x).to_bits());
        }
        for x in [-3.0, -0.0, 0.0, 0.37, 1.37, 7.0] {
            let (sin, cos) = sin_cosf(x);
            assert_eq!(sin.to_bits(), libm::sinf(x).to_bits());
            assert_eq!(cos.to_bits(), libm::cosf(x).to_bits());
        }
    }
}
