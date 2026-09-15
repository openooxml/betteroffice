# @betteroffice/docx-native

Native Rust bindings for the BetterOffice DOCX engine in Node.js and Bun. The package is intended for servers, CLIs, Electron applications, and other Node.js or Bun hosts that need the native engine without a browser or WebAssembly runtime.

The API follows Node.js conventions while preserving the data model, limits, and editing semantics of the Rust facade and its Python binding. Facade operations return promises and run in call order on one native worker without blocking the JavaScript event loop. Document bytes use `Buffer`, and structured contracts use ordinary JavaScript objects.
