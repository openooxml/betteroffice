/** Shared editor orchestration for framework adapters. */

export {
  buildResidentRegionLayoutRequest,
  computeLayout,
  workerLayoutComputation,
  getLayoutKernelInputs,
} from './computeLayout';
export type { ComputeLayoutInputs, LayoutComputation } from './computeLayout';
export { resolvedFinalSectionProperties, updateFinalSectionProperties } from './finalSection';
