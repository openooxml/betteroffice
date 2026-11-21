import { spawn } from 'node:child_process';
import { closeSync, mkdirSync, openSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

export function privateMode(environment = process.env) {
  const value = environment.QUALITY_PRIVATE?.trim() ?? '';
  if (!['', 'false', 'true'].includes(value)) throw new Error('QUALITY_PRIVATE must be true, false or empty');
  return value === 'true';
}

export function privateLogDirectory(environment = process.env) {
  const directory = environment.QUALITY_PRIVATE_LOGS?.trim() ||
    (environment.RUNNER_TEMP?.trim() && resolve(environment.RUNNER_TEMP, 'private-logs'));
  if (!directory) throw new Error('QUALITY_PRIVATE_LOGS or RUNNER_TEMP is required in private mode');
  return resolve(directory);
}

export async function runStep(label, command, args, environment = process.env) {
  if (!/^[a-z0-9][a-z0-9-]{0,63}$/.test(label ?? '')) throw new Error('Invalid private step label');
  if (!command) throw new Error('Missing private step command');
  const isPrivate = privateMode(environment);
  let output;
  if (isPrivate) {
    const directory = privateLogDirectory(environment);
    mkdirSync(directory, { recursive: true });
    output = openSync(resolve(directory, `${label}.log`), 'a');
  }
  try {
    const code = await new Promise(done => {
      const child = spawn(command, args, { env: environment, stdio: isPrivate ? ['ignore', output, output] : 'inherit' });
      child.on('error', () => done(127));
      child.on('close', (status, signal) => done(signal ? 1 : status ?? 1));
    });
    if (isPrivate) process.stdout.write(`${label}: exit ${code}; output kept in the private log\n`);
    return code;
  } finally {
    if (output !== undefined) closeSync(output);
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [label, separator, command, ...args] = process.argv.slice(2);
  if (separator !== '--') {
    process.stderr.write('usage: private-step.mjs <label> -- <command> [args...]\n');
    process.exit(2);
  }
  try {
    process.exitCode = await runStep(label, command, args);
  } catch (error) {
    process.stderr.write(`private-step: ${error.message}\n`);
    process.exitCode = 2;
  }
}
