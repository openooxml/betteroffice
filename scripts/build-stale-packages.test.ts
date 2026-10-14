import { describe, expect, test } from 'bun:test';
import { order, readTargets } from './build-stale-packages.ts';

const targets = await readTargets();
const names = order(targets).map((target) => target.name);
const position = new Map(names.map((name, index) => [name, index]));

describe('workspace build order', () => {
  test('every workspace dependency and peer builds before its dependent', () => {
    const late: string[] = [];
    for (const target of targets) {
      for (const dep of target.deps) {
        if (!position.has(dep)) continue;
        if (position.get(dep)! > position.get(target.name)!) {
          late.push(`${dep} builds after ${target.name}`);
        }
      }
    }
    expect(late).toEqual([]);
  });

  // docx peer-depends on fonts, which peer-depends on fonts-cjk. Reading only
  // `dependencies` left docx first and its dts build could not resolve fonts.
  test('a workspace peer is a build dependency', () => {
    expect(targets.find((target) => target.name === '@betteroffice/docx')?.deps).toContain(
      '@betteroffice/fonts'
    );
    expect(targets.find((target) => target.name === '@betteroffice/fonts')?.deps).toContain(
      '@betteroffice/fonts-cjk'
    );
    expect(position.get('@betteroffice/fonts')!).toBeLessThan(position.get('@betteroffice/docx')!);
    expect(position.get('@betteroffice/fonts-cjk')!).toBeLessThan(
      position.get('@betteroffice/fonts')!
    );
  });

  test('every package that builds is ordered exactly once', () => {
    expect(names.length).toBe(targets.length);
    expect(new Set(names).size).toBe(names.length);
  });
});
