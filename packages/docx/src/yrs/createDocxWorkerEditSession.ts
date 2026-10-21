import { createWorkerEditSession } from '../../../../shared/office-session/commands/session';
import { cloneData } from '../../../../shared/office-session/commands/clone';
import { SessionFailure } from '../../../../shared/office-session/types';
import { discardCommandPresentation } from './residentCommandPresentation';
import type {
  DocxCommandAck,
  DocxWorkerEditSession,
  DocxWorkerEditSessionOptions,
  ResidentCommandConnection,
} from './residentCommandProtocol';

export function createDocxWorkerEditSession<O, X = never>(
  options: DocxWorkerEditSessionOptions<O, X>
): DocxWorkerEditSession<X> {
  const connections = new WeakMap<object, ResidentCommandConnection>([[options.connection, options.connection]]);
  const ackListeners = new Set<(ack: DocxCommandAck) => void>();
  const session = createWorkerEditSession({
    ...options,
    onAck(ack) {
      options.onAck?.(ack);
      for (const listener of ackListeners) listener(ack as DocxCommandAck);
    },
    copyCommandForDispatch: command =>
      command.kind === 'applyRemoteUpdate' || command.kind === 'applyRemoteUpdateChunks' || command.kind === 'collabPipe'
        ? command : cloneData(command),
    discardFrame: discardCommandPresentation,
    sameCoverageFrames: true,
    transactionalPresentation: true,
    echoOverflow: 'decline',
    maxEchoEntries: 256,
    recovery: {
      ...options.recovery,
      async reopen(input) {
        const connection = await options.recovery.reopen(input);
        connections.set(connection, connection);
        return connection;
      },
    },
    requestFullFrame(connection) {
      const resident = connections.get(connection);
      if (!resident) {
        return Promise.reject(new SessionFailure('message', 'Command connection cannot request a full frame'));
      }
      return resident.requestFullFrame();
    },
  });
  return Object.assign(session, {
    subscribeAck(listener: (ack: DocxCommandAck) => void) {
      ackListeners.add(listener);
      return () => { ackListeners.delete(listener); };
    },
  });
}
