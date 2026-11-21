# @betteroffice/fonts-intl

## 0.5.0

### Minor Changes

- df33f6a: Adds optional Noto coverage faces for international scripts, loaded on demand by `@betteroffice/fonts`.

### Patch Changes

- e6d5c52: Package exports now carry a `default` condition next to `import`, so webpack consumers (`next build --webpack`, including editors loaded with `next/dynamic`) resolve every entry instead of failing with "Package path is not exported".
