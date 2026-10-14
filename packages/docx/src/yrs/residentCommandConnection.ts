import { isCommandAck, isCommandOutcomeUnknownError } from '../../../../shared/office-session/commands/types';
import { SessionFailure } from '../../../../shared/office-session/types';
import { discardCommandPresentation } from './residentCommandPresentation';
import type {
  ResidentCommandClient,
  DocxCommandAck,
  ResidentCommandConnection,
  ResidentCommandInitReply,
  ResidentCommandInitRequest,
} from './residentCommandProtocol';

export async function createResidentCommandConnection(
  client: ResidentCommandClient,
  init: ResidentCommandInitRequest | ResidentCommandInitReply
): Promise<ResidentCommandConnection> {
  const generation = 'outcomes' in init ? client.commandInitialGeneration(init) : client.commandGeneration;
  if (generation === undefined || generation !== client.commandGeneration) {
    if ('outcomes' in init) discardCommandPresentation(init.frame);
    throw new SessionFailure('disposed', 'Resident command document was replaced');
  }
  let failure: SessionFailure | undefined;
  let disposal: Promise<void> | undefined;
  const listeners = new Set<(failure: SessionFailure) => void>();
  const pending = new Set<(failure: SessionFailure) => void>();
  const frameSubscriptions = new Set<() => void>();

  function fail(error: unknown, deferNotification = false): SessionFailure {
    if (failure) return failure;
    failure = error instanceof SessionFailure ? error : new SessionFailure(
      'crash', error instanceof Error ? error.message : String(error)
    );
    const terminal = failure;
    for (const unsubscribe of frameSubscriptions) unsubscribe();
    frameSubscriptions.clear();
    for (const reject of pending) reject(terminal);
    pending.clear();
    const notify = () => {
      for (const listener of [...listeners]) {
        try { listener(terminal); } catch {}
      }
    };
    if (deferNotification) queueMicrotask(notify);
    else notify();
    return failure;
  }

  function call<T>(request: () => Promise<T>, discard?: (value: T) => void): Promise<T> {
    if (client.commandGeneration !== generation) {
      return Promise.reject(fail(new SessionFailure('disposed', 'Resident command document was replaced')));
    }
    if (failure) return Promise.reject(failure);
    if (disposal) return Promise.reject(new SessionFailure('disposed', 'Command connection was disposed'));
    return new Promise<T>((resolve, reject) => {
      pending.add(reject);
      let reply: Promise<T>;
      try { reply = request(); }
      catch (error) { reply = Promise.reject(error); }
      void reply.then((value) => {
        pending.delete(reject);
        if (failure || disposal || client.commandGeneration !== generation) {
          discard?.(value);
          reject(failure ?? fail(new SessionFailure('disposed', 'Resident command document was replaced')));
        } else resolve(value);
      }, (error: unknown) => {
        pending.delete(reject);
        if (isCommandOutcomeUnknownError(error)) {
          reject(error);
          fail(error, true);
        } else {
          reject(fail(error));
        }
      });
    });
  }

  function dispose(): Promise<void> {
    if (disposal) return disposal;
    disposal = Promise.resolve().then(() => {
      const current = client.commandGeneration === generation;
      if (current) client.onFailure(null);
      fail(new SessionFailure('disposed', 'Command connection was disposed'));
      listeners.clear();
      if (current) client.destroy();
    });
    return disposal;
  }

  client.onFailure((error) => { fail(error); });
  if (client.commandFailure) fail(client.commandFailure);
  try {
    const initial = 'outcomes' in init ? init : await call(() => client.commandInit(init),
      (value) => discardCommandPresentation(value.frame));
    if (client.commandGeneration !== generation || client.commandInitialGeneration(initial) !== generation) {
      discardCommandPresentation(initial.frame);
      throw new SessionFailure('disposed', 'Resident command document was replaced');
    }
    if (!isCommandAck(initial) || initial.session !== init.session || initial.incarnation !== init.incarnation ||
      initial.epoch !== init.epoch || initial.version !== (init.version ?? 0) ||
      initial.throughSeq !== (init.throughSeq ?? 0) || initial.selection === undefined ||
      initial.editState === undefined || !initial.frame || initial.frame.baseRevision !== undefined ||
      initial.frame.coversThroughSeq !== initial.throughSeq) {
      discardCommandPresentation(initial.frame);
      throw new SessionFailure('message', 'Resident command initialization has invalid state');
    }
    function validateAck(ack: DocxCommandAck): DocxCommandAck {
      const frame = ack?.frame;
      if (!isCommandAck(ack)) {
        discardCommandPresentation(frame);
        throw new SessionFailure('message', 'Malformed resident command acknowledgement');
      }
      if (ack.session !== initial.session || ack.incarnation !== initial.incarnation || ack.epoch !== initial.epoch) {
        discardCommandPresentation(ack.frame);
        throw fail(new SessionFailure('disposed', 'Resident command document was replaced'));
      }
      return ack;
    }
    const discardAck = (ack: DocxCommandAck) => discardCommandPresentation(ack.frame);
    return {
      initial,
      requestFullFrame: () => call(() => client.requestFullFrame().then(validateAck), discardAck),
      requestViewport: client.requestViewport ? (request) => call(() => client.requestViewport!(request)) : undefined,
      client: {
        onFrame(listener) {
          if (failure || disposal) return () => {};
          const unsubscribe = client.onFrame?.((frame) => {
            if (!failure && !disposal && client.commandGeneration === generation) listener(frame);
            else if (client.commandGeneration === generation) discardCommandPresentation(frame);
          }) ?? (() => {});
          frameSubscriptions.add(unsubscribe);
          return () => { frameSubscriptions.delete(unsubscribe); unsubscribe(); };
        },
        call: {
          edit: (command) => call(() => client.edit(command).then(validateAck), discardAck),
          checkpoint: () => call(() => client.checkpoint()),
        },
        get failure() { return failure; },
        onFailure(listener) {
          if (failure) {
            listener(failure);
            return () => {};
          }
          listeners.add(listener);
          return () => { listeners.delete(listener); };
        },
        dispose,
      },
    };
  } catch (error) {
    await dispose();
    throw error;
  }
}
