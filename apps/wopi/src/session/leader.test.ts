import { describe, expect, test } from "bun:test";
import { electLeader, type LeaderCandidate } from "./leader";

const NOW = 1700000000000;
function user(partial: Partial<LeaderCandidate> & { id: string }): LeaderCandidate {
  return { joinedAt: NOW - 10000, connected: true, canWrite: true, expiresAt: null, ...partial };
}

describe("electLeader", () => {
  test("returns null for empty and for no eligible user", () => {
    expect(electLeader([], NOW)).toBeNull();
    expect(electLeader([user({ id: "a", connected: false })], NOW)).toBeNull();
    expect(electLeader([user({ id: "a", canWrite: false })], NOW)).toBeNull();
  });

  test("picks the earliest joined connected writer", () => {
    const late = user({ id: "late", joinedAt: NOW - 1000 });
    const early = user({ id: "early", joinedAt: NOW - 9000 });
    const mid = user({ id: "mid", joinedAt: NOW - 5000 });
    expect(electLeader([late, early, mid], NOW)?.id).toBe("early");
  });

  test("ties keep input order", () => {
    const a = user({ id: "a", joinedAt: NOW - 1000 });
    const b = user({ id: "b", joinedAt: NOW - 1000 });
    expect(electLeader([a, b], NOW)?.id).toBe("a");
    expect(electLeader([b, a], NOW)?.id).toBe("b");
  });

  test("skips disconnected and read only users even when earlier", () => {
    const gone = user({ id: "gone", joinedAt: NOW - 9000, connected: false });
    const viewer = user({ id: "viewer", joinedAt: NOW - 8000, canWrite: false });
    const writer = user({ id: "writer", joinedAt: NOW - 1000 });
    expect(electLeader([gone, viewer, writer], NOW)?.id).toBe("writer");
  });

  test("skips tokens expiring within two minutes, accepts exactly two minutes", () => {
    const soon = user({ id: "soon", joinedAt: NOW - 9000, expiresAt: NOW + 119999 });
    const exact = user({ id: "exact", joinedAt: NOW - 8000, expiresAt: NOW + 120000 });
    const later = user({ id: "later", joinedAt: NOW - 1000, expiresAt: NOW + 3600000 });
    expect(electLeader([soon, later], NOW)?.id).toBe("later");
    expect(electLeader([soon, exact], NOW)?.id).toBe("exact");
    expect(electLeader([user({ id: "expired", expiresAt: NOW - 1 })], NOW)).toBeNull();
  });

  test("unknown expiry counts as valid", () => {
    expect(electLeader([user({ id: "a", expiresAt: null })], NOW)?.id).toBe("a");
    expect(electLeader([user({ id: "b", expiresAt: undefined })], NOW)?.id).toBe("b");
    expect(electLeader([user({ id: "c", expiresAt: 0 })], NOW)?.id).toBe("c");
  });

  test("is pure and does not mutate its input", () => {
    const users = [user({ id: "b", joinedAt: 2 }), user({ id: "a", joinedAt: 1 })];
    const snapshot = users.map((u) => ({ ...u }));
    electLeader(users, NOW);
    expect(users.map((u) => ({ ...u }))).toEqual(snapshot);
    expect(users[0]?.id).toBe("b");
  });
});
