/**
 * `core/editor/` — shared stateful-orchestration spine for the React/Vue
 * adapters (issue #696 Tier 2). The pure layout COMPUTE pass + the
 * rAF-coalescing layout scheduler; the transaction loop, view lifecycle, and
 * session seam land here as the tier proceeds.
 */

export {
  buildResidentRegionLayoutRequest,
  computeLayout,
  getLayoutKernelInputs,
} from './computeLayout';
export type { ComputeLayoutInputs, LayoutComputation } from './computeLayout';
