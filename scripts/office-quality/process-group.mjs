import { spawn } from 'node:child_process';

// A child that leads its own process group, so a kill also reaches its workers.
export const spawnGroup = (command, args, options = {}) =>
  spawn(command, args, { ...options, detached: true });

export function killGroup(child, signal) {
  try {
    process.kill(-child.pid, signal);
  } catch {}
}
