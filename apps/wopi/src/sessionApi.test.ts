import { expect, test } from "bun:test";
import { encodeStateVector } from "yjs";
import { waitForStateVector } from "./sessionApi";

const vector = (entries: Array<[number, number]>) => encodeStateVector(new Map(entries));

test("covered state vectors proceed immediately", async () => {
  await waitForStateVector(() => vector([[1, 3], [2, 4]]), vector([[1, 2], [2, 4]]), { timeoutMs: 0 });
});

test("state vectors wait until every requested clock is covered", async () => {
  let current = vector([[1, 1]]);
  let covered = false;
  const waiting = waitForStateVector(() => current, vector([[1, 2], [2, 1]]), { timeoutMs: 500 }).then(() => { covered = true; });
  await new Promise((resolve) => setTimeout(resolve, 60));
  expect(covered).toBe(false);
  current = vector([[1, 2], [2, 1]]);
  await waiting;
  expect(covered).toBe(true);
});

test("an uncovered vector times out with a typed notSynced failure", async () => {
  await expect(waitForStateVector(() => vector([]), vector([[7, 1]]), { timeoutMs: 1 })).rejects.toMatchObject({ kind: "notSynced" });
});
