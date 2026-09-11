/** Shared editor orchestration for framework adapters. */

export {
  buildResidentRegionLayoutRequest,
  computeLayout,
  getLayoutKernelInputs,
} from './computeLayout';
export type { ComputeLayoutInputs, LayoutComputation } from './computeLayout';
export { resolvedFinalSectionProperties, updateFinalSectionProperties } from './finalSection';
