# @betteroffice/drawingml

Shared DrawingML (`a:` namespace) primitives for OOXML formats — theme
parsing, color math (tint/shade, theme slots), shape fill/outline/transform
parsing, and `a:`-level serialization.

DrawingML is the graphics language shared by WordprocessingML (docx),
PresentationML (pptx), and SpreadsheetML (xlsx). This package holds the
format-agnostic pieces so each format host doesn't reimplement them.

## XML seam

The package has **no runtime dependencies** and no XML parser of its own.
Hosts parse XML however they like and pass in any tree whose nodes satisfy
the structural `XmlLike` interface:

```ts
interface XmlLike {
  type?: string; // element nodes are 'element'
  name?: string; // prefixed name, e.g. 'a:srgbClr'
  attributes?: Record<string, string | number | undefined>;
  elements?: XmlLike[];
}
```

An xml-js `Element` tree (non-compact mode) satisfies this out of the box.

## Entry points

- `@betteroffice/drawingml` — theme parsing (`parseTheme`, `getThemeColor`,
  `resolveThemeFontRef`), color math (`applyTint`, `applyShade`,
  `resolveColor`, `generateThemeTintShadeMatrix`), shape parsing
  (`parseFill`, `parseOutline`, `parseTransform`, `parseGradientFill`),
  and serialization (`serializeFill`, `serializeOutline`).
- `@betteroffice/drawingml/units` — EMU/point/pixel conversion helpers.

Used by [`@betteroffice/docx`](https://www.npmjs.com/package/@betteroffice/docx)
and the OpenOOXML pptx project. Versioned in lockstep with the other
`@openooxml` packages.
