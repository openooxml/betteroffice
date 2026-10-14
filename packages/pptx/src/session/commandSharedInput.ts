import type { CommandTicket } from '../../../../shared/office-session/commands/types';
import type {
  PresentationCommandAck, PresentationCommandBasis, PresentationCommandClient, PresentationCommandIntent,
} from './commandProtocol';

export interface PresentationSharedInput {
  subscribeSent(listener: () => void): () => void;
  submit(intent: PresentationCommandIntent, basis: PresentationCommandBasis, after?: CommandTicket): {
    ticket: CommandTicket;
    operation: Promise<PresentationCommandAck>;
  };
}

const sharedInputKey = Symbol.for('betteroffice.pptx.sharedInput');
type SharedClient = PresentationCommandClient & { [sharedInputKey]?: PresentationSharedInput };

export const presentationSharedInputs = {
  get(client: PresentationCommandClient): PresentationSharedInput | undefined {
    return (client as SharedClient)[sharedInputKey];
  },
  set(client: PresentationCommandClient, input: PresentationSharedInput): void {
    Object.defineProperty(client, sharedInputKey, { value: input, configurable: true });
  },
  delete(client: PresentationCommandClient): void {
    delete (client as SharedClient)[sharedInputKey];
  },
};
