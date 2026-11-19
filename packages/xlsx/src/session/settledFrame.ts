import type { DisplayList } from '../display-list/types';

/** The first frame `read` returns with no conditional format left undecided, for output no later frame replaces. */
export function settledDisplayListJson(read: () => string): string {
  for (;;) {
    const json = read();
    if (!(JSON.parse(json) as DisplayList).conditionalPending) return json;
  }
}
