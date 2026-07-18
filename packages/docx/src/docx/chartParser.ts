import type { Chart } from '../types/document';

/**
 * Compatibility context retained in public leaf-parser signatures.
 * Full package chart parsing is owned by Rust S9.
 */
export type ChartPartsMap = Map<string, Chart>;
