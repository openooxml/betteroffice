pub(crate) fn sin_cos(x: f32) -> (f32, f32) {
    #[cfg(target_arch = "wasm32")]
    {
        x.sin_cos()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        (libm::sinf(x), libm::cosf(x))
    }
}
