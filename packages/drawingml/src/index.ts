/**
 * Shared DrawingML (`a:` namespace) primitives for OOXML formats — theme
 * parsing, color math, shape fill/outline/transform parsing, and a:-level
 * serialization. Format hosts (docx, pptx, xlsx) provide the parsed XML tree
 * through the structural `XmlLike` interface; the package itself has no
 * runtime dependencies.
 *
 * @packageDocumentation
 * @public
 */

export * from './xml';
export * from './color';
export * from './theme';
export * from './drawing';
export * from './shape';
export * from './serialize';
export * from './units';
