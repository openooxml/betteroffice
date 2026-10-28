pub(crate) fn cos(x: f32) -> f32 {
    #[cfg(target_arch = "wasm32")]
    {
        x.cos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::cosf(x)
    }
}

pub(crate) fn sin(x: f32) -> f32 {
    #[cfg(target_arch = "wasm32")]
    {
        x.sin()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::sinf(x)
    }
}
