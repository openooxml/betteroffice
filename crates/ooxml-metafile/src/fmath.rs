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
        ooxml_drawingml::correctly_rounded_hypot(x, y)
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
            assert_eq!(atan2(x, 3.0).to_bits(), libm::atan2(x, 3.0).to_bits());
            assert_eq!(
                hypot(x, 3.0).to_bits(),
                ooxml_drawingml::correctly_rounded_hypot(x, 3.0).to_bits()
            );
            let (sin, cos) = sin_cos(x);
            assert_eq!(sin.to_bits(), libm::sin(x).to_bits());
            assert_eq!(cos.to_bits(), libm::cos(x).to_bits());
        }
    }
}
