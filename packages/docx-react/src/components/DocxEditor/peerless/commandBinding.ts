import type {
  PeerlessAdmissionOptions,
  PeerlessAdmissionResult,
  PeerlessFormattingIntent,
  PeerlessInputAdmission,
} from './inputAdmission';
import type { PeerlessInputIntent } from './inputEvents';
import { admitIntent, submitFormattingIntent } from './inputAdmission';

/** @internal Toolbar and keyboard entry points share the admission tail. */
export interface PeerlessCommandBinding {
  submit(intent: PeerlessInputIntent, options?: PeerlessAdmissionOptions): PeerlessAdmissionResult;
  format(intent: PeerlessFormattingIntent, options?: PeerlessAdmissionOptions): PeerlessAdmissionResult;
}

/** @internal */
export function createPeerlessCommandBinding(_admission: PeerlessInputAdmission): PeerlessCommandBinding {
  return {
    submit: (intent, options) => admitIntent(_admission, intent, options),
    format: (intent, options) => submitFormattingIntent(_admission, intent, options),
  };
}
