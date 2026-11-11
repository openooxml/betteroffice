// @betteroffice/docx declaration editor/index.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Shared editor orchestration for framework adapters. */
export { buildResidentRegionLayoutRequest, computeLayout, documentRenderEnv, workerLayoutComputation, getLayoutKernelInputs, } from './computeLayout';
export type { ComputeLayoutInputs, LayoutComputation } from './computeLayout';
export { resolvedFinalSectionProperties, updateFinalSectionProperties } from './finalSection';
