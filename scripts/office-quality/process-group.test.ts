import { expect, test } from 'bun:test';
import { killGroup, spawnGroup } from './process-group.mjs';

const alive = (pid: number) => {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
};

test('killing a group also stops the workers its leader started', async () => {
  const child = spawnGroup('sh', ['-c', 'sleep 30 & echo $!; wait'], { stdio: ['ignore', 'pipe', 'ignore'] });
  const worker = Number(await new Promise<string>((resolve) => child.stdout!.once('data', (data) => resolve(String(data)))));
  const closed = new Promise((resolve) => child.once('close', resolve));
  expect(alive(worker)).toBe(true);
  killGroup(child, 'SIGTERM');
  await closed;
  for (let attempt = 0; attempt < 50 && alive(worker); attempt++) await Bun.sleep(20);
  expect(alive(worker)).toBe(false);
});
