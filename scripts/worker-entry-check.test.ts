import { expect, test } from 'bun:test';

type Check = { problems: string[]; target: string | null };
const validatePublishedWorker: ((source: string) => Promise<Check>) | undefined = await import('./worker-entry-check.mjs').then(
  (module) => module.validatePublishedWorker,
  () => undefined,
);

test('a worker entry that only imports one literal path dynamically passes with that path as its target', async () => {
  expect(validatePublishedWorker).toBeFunction();
  expect(await validatePublishedWorker!("import('./x.js');")).toEqual({ problems: [], target: './x.js' });
});

test('a worker entry whose dynamic import target is not a string literal fails', async () => {
  expect(validatePublishedWorker).toBeFunction();
  const { problems, target } = await validatePublishedWorker!('const name = globalThis.start;\nimport(name);');
  expect({ failed: problems.length > 0, target }).toEqual({ failed: true, target: null });
});

test('a worker entry with a static import behind a comment fails', async () => {
  expect(validatePublishedWorker).toBeFunction();
  const { problems } = await validatePublishedWorker!("import /*c*/ './x.js'; import('./y.js');");
  expect(problems.length).toBeGreaterThan(0);
});
