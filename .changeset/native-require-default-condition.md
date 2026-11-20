---
'@betteroffice/docx-native': patch
'@betteroffice/xlsx-native': patch
'@betteroffice/pptx-native': patch
---

`require()` of the native packages from CommonJS resolves, as it does for the WebAssembly packages: every export gains a `default` condition beside `import`. Before, it failed with `ERR_PACKAGE_PATH_NOT_EXPORTED`. Loading an ES module by `require()` needs Node 20.19+ or 22.12+.
