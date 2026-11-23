import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { auditNpmPackages, missingHeldNatives, releaseEntries } from './check-publish-targets.mjs';
import { workspaceEntries } from './published-packages.mjs';

export function releaseTag({ name, version }) {
  return `${name}@${version}`;
}

export function releaseNotes(changelog, version) {
  const lines = changelog.split(/\r?\n/);
  let start;
  let fence;
  for (const [index, line] of lines.entries()) {
    const marker = /^ {0,3}(`{3,}|~{3,})(.*)$/.exec(line);
    if (fence) {
      if (marker && marker[1][0] === fence[0] && marker[1].length >= fence.length && !marker[2].trim()) fence = undefined;
      continue;
    }
    if (marker) {
      fence = marker[1];
      continue;
    }
    const heading = /^ {0,3}(#{1,2})[ \t]+(.+?)(?:[ \t]+#+)?[ \t]*$/.exec(line);
    if (!heading) continue;
    if (start !== undefined) return lines.slice(start, index).join('\n').trim();
    if (heading[1] === '##' && heading[2] === version) start = index + 1;
  }
  return start === undefined ? '' : lines.slice(start).join('\n').trim();
}

function tarballPath(destination, { name, version }) {
  return join(destination, name, `${version}.tgz`);
}

function isPrerelease(version) {
  return version.split('+')[0].includes('-');
}

function packHeld(entries, destination) {
  for (const entry of entries) {
    const { name, version, directory, manifest } = entry;
    const tarball = tarballPath(destination, entry);
    const packDirectory = dirname(tarball);
    mkdirSync(packDirectory, { recursive: true });
    const [packed] = JSON.parse(execFileSync('npm', ['pack', '--json', '--pack-destination', packDirectory], {
      cwd: directory, encoding: 'utf8', stdio: ['ignore', 'pipe', 'inherit']
    }));
    if (packed?.name !== name || packed.version !== version) throw new Error(`npm pack did not produce ${releaseTag(entry)}`);
    renameSync(join(packDirectory, packed.filename), tarball);
    writeFileSync(join(directory, 'package.json'), JSON.stringify({ ...manifest, private: true }, null, 2) + '\n');
  }
}

function ensureRelease(entry) {
  const { directory, version } = entry;
  const tag = releaseTag(entry);
  const ref = `refs/tags/${tag}`;
  if (!execFileSync('git', ['ls-remote', '--refs', 'origin', ref], { encoding: 'utf8' }).trim()) {
    execFileSync('git', [
      '-c', 'user.name=github-actions[bot]',
      '-c', 'user.email=41898282+github-actions[bot]@users.noreply.github.com',
      'tag', '-a', tag, '-m', tag, process.env.GITHUB_SHA
    ], { stdio: 'inherit' });
    execFileSync('git', ['push', 'origin', ref], { stdio: 'inherit' });
  }
  if (spawnSync('gh', ['release', 'view', tag], { stdio: 'ignore' }).status === 0) return;
  const changelog = join(directory, 'CHANGELOG.md');
  const notes = existsSync(changelog) ? releaseNotes(readFileSync(changelog, 'utf8'), version) : '';
  const prerelease = isPrerelease(version) ? ['--prerelease'] : [];
  execFileSync('gh', ['release', 'create', tag, '--verify-tag', '--title', tag, '--notes', notes, ...prerelease], {
    stdio: 'inherit'
  });
}

async function publishHeld(entries, workspaces, destination) {
  const skipped = process.env.NATIVE_NPM_RESULT === 'skipped';
  const missing = await missingHeldNatives(entries, { wait: !skipped, workspaces });
  if (missing.length) {
    const message = `Held packages are waiting for npm: ${missing.map(releaseTag).join(', ')}.`;
    if (!skipped) throw new Error(message);
    console.log(message);
    return;
  }

  const preFile = new URL('../.changeset/pre.json', import.meta.url);
  const preState = existsSync(preFile) ? JSON.parse(readFileSync(preFile, 'utf8')) : undefined;
  const preTag = preState?.mode === 'pre' && typeof preState.tag === 'string' ? preState.tag.trim() : '';
  const tag = preTag ? ['--tag', preTag] : [];
  const audit = await auditNpmPackages(entries);
  for (const [index, entry] of entries.entries()) {
    if (audit[index].state === 'published') {
      console.log(`${releaseTag(entry)} is already on npm.`);
    } else {
      if (isPrerelease(entry.version) && (!preTag || preTag.toLowerCase() === 'latest')) {
        throw new Error(`Cannot publish ${releaseTag(entry)} without a prerelease tag in .changeset/pre.json.`);
      }
      execFileSync('npm', ['publish', tarballPath(destination, entry), '--provenance', '--access', 'public', ...tag], {
        stdio: 'inherit'
      });
    }
    ensureRelease(entry);
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const [mode, destination, ...rest] = process.argv.slice(2);
  if (!['--pack', '--publish'].includes(mode) || !destination || rest.length) {
    throw new Error('usage: held-npm-packages.mjs --pack|--publish <tarball directory>');
  }
  const packages = JSON.parse(process.env.HELD_NPM_PACKAGES || '[]');
  const workspaces = workspaceEntries();
  const entries = releaseEntries(packages, workspaces);
  if (mode === '--pack') packHeld(entries, resolve(destination));
  else await publishHeld(entries, workspaces, resolve(destination));
}
