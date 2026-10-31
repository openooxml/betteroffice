use super::MeasureFonts;

#[doc(hidden)]
pub fn with_measure_font_store<T>(store: &mut ooxml_text::FontStore, run: impl FnOnce() -> T) -> T {
    struct Restore<'a> {
        target: &'a mut ooxml_text::FontStore,
        fonts: MeasureFonts,
    }
    impl Drop for Restore<'_> {
        fn drop(&mut self) {
            *self.target = std::mem::take(&mut *self.fonts.0.borrow_mut());
        }
    }
    let restore = Restore {
        target: store,
        fonts: MeasureFonts::default(),
    };
    *restore.fonts.0.borrow_mut() = std::mem::take(&mut *restore.target);
    let _scope = restore.fonts.enter();
    run()
}
