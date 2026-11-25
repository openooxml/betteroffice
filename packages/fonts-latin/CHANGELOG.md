# @betteroffice/fonts-latin

## 0.5.0

### Minor Changes

- 37199d9: Ships the same Latin and RTL faces in a smaller compressed install by default. Existing providers decode faces on demand, and DOCX reuses each face across fallback registrations.

### Patch Changes

- e6d5c52: Package exports now carry a `default` condition next to `import`, so webpack consumers (`next build --webpack`, including editors loaded with `next/dynamic`) resolve every entry instead of failing with "Package path is not exported".
