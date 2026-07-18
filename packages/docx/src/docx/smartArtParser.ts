/** Compatibility context retained in public leaf-parser signatures. */
export interface SmartArtContext {
  parts: Map<string, string>;
  warnings?: string[];
}
