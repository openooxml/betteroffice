import { describe, expect, test } from 'bun:test';
import { planDemoSession, type DemoSessionInputs } from './demoSession';

const ROOM = 'room-1';
const seeded = { seed: Uint8Array.of(1, 2) };
const local = { seed: null };

describe('planDemoSession', () => {
  test('withholds the editor until a document is loaded', () => {
    expect(
      planDemoSession({ document: null, room: ROOM, clientId: 7, identified: true })
    ).toEqual({ status: 'loading' });
  });

  test('mounts a seeded document with its room', () => {
    expect(
      planDemoSession({ document: seeded, room: ROOM, clientId: 7, identified: true })
    ).toEqual({ status: 'shared', room: ROOM, clientId: 7 });
  });

  const incomplete: Array<[string, Omit<DemoSessionInputs, 'document'>]> = [
    ['no room', { room: null, clientId: 7, identified: true }],
    ['no client id', { room: ROOM, clientId: null, identified: true }],
    ['no identity', { room: ROOM, clientId: 7, identified: false }],
  ];

  for (const [missing, rest] of incomplete) {
    test(`holds a seeded document back with ${missing} rather than mounting it unshared`, () => {
      expect(planDemoSession({ document: seeded, ...rest })).toEqual({ status: 'loading' });
    });
  }

  test('a locally opened file is never shared', () => {
    expect(
      planDemoSession({ document: local, room: ROOM, clientId: 7, identified: true })
    ).toEqual({ status: 'local' });
  });

  test('a locally opened file stays local while a room lingers in the URL', () => {
    expect(
      planDemoSession({ document: local, room: 'stale-room', clientId: 9, identified: true })
    ).toEqual({ status: 'local' });
  });
});
