pub(crate) fn acos(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.acos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::acos(x)
    }
}

pub(crate) fn asin(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.asin()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::asin(x)
    }
}

pub(crate) fn atan(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.atan()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::atan(x)
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

pub(crate) fn cosh(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.cosh()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::cosh(x)
    }
}

pub(crate) fn exp(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.exp()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::exp(x)
    }
}

pub(crate) fn ln(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.ln()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::log(x)
    }
}

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

pub(crate) fn sinh(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.sinh()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::sinh(x)
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

pub(crate) fn tanh(x: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.tanh()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::tanh(x)
    }
}

pub(crate) fn log(x: f64, base: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.log(base)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        ln(x) / ln(base)
    }
}

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
