# Third-party notices

This wheel embeds no third-party fonts or other assets. The extension module
statically links the Rust crates this package depends on, each under its own
permissive license (Apache-2.0 or MIT); their notices travel with those crates
on crates.io. The package's own code is Apache-2.0.

Document text is measured and rasterized with fonts the caller registers
through `register_font`, so no face is compiled into the wheel.
