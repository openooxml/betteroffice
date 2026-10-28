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
