import { useId } from 'react';
import type { ReactNode } from 'react';

export interface DisabledDescription {
  /** Keeps a described trigger focusable so its reason can be read. */
  triggerProps:
    | { disabled: boolean; 'aria-disabled'?: undefined; 'aria-describedby'?: undefined }
    | { disabled: false; 'aria-disabled': true; 'aria-describedby': string };
  /** Hidden description element referenced by the trigger. */
  node: ReactNode;
  /** Reason shown on hover, when disabled. */
  title: string | undefined;
  /** Whether the control is disabled with a stated reason. */
  described: boolean;
}

/** Accessible disabled-reason wiring for a control trigger. */
export function useDisabledDescription(
  disabled: boolean,
  description: string | undefined
): DisabledDescription {
  const id = useId();
  const reason = disabled && description ? description : undefined;
  return reason
    ? {
        triggerProps: { disabled: false, 'aria-disabled': true, 'aria-describedby': id },
        node: (
          <span id={id} hidden>
            {reason}
          </span>
        ),
        title: reason,
        described: true,
      }
    : { triggerProps: { disabled }, node: null, title: undefined, described: false };
}
