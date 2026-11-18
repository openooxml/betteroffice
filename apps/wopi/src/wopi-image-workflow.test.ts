import { describe, expect, test } from "bun:test";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const repoRoot = join(import.meta.dir, "..", "..", "..");
const wopiWorkflowPath = join(repoRoot, ".github", "workflows", "wopi-image.yml");
const releaseWorkflowPath = join(repoRoot, ".github", "workflows", "release.yml");

type Step = {
  id?: string;
  name?: string;
  if?: string;
  uses?: string;
  run?: string;
  with?: Record<string, unknown>;
  env?: Record<string, unknown>;
};
type Job = {
  if?: string;
  uses?: string;
  needs?: string | string[];
  outputs?: Record<string, string>;
  permissions?: Record<string, string>;
  with?: Record<string, unknown>;
  steps?: Step[];
};
type Workflow = {
  on?: unknown;
  permissions?: Record<string, string>;
  jobs: Record<string, Job>;
};

function loadWorkflow(path: string): Workflow {
  const parsed = Bun.YAML.parse(readFileSync(path, "utf8")) as Record<string, unknown>;
  const triggers = (parsed.on ?? parsed[true as unknown as string]) as Workflow["on"];
  return { ...(parsed as unknown as Workflow), on: triggers };
}

function buildPushSteps(job: Job): Step[] {
  return (job.steps ?? []).filter((step) => step.uses?.startsWith("docker/build-push-action@"));
}

const PINNED = /^[^@\s]+@[0-9a-f]{40}$/;
const ACTION_PINS: Record<string, string> = {
  "actions/checkout": "3d3c42e5aac5ba805825da76410c181273ba90b1 # v7.0.1",
  "docker/build-push-action": "10e90e3645eae34f1e60eeb005ba3a3d33f178e8 # v6.19.2",
  "docker/login-action": "c94ce9fb468520275223c153574b00df6fe4bcc9 # v3.7.0",
  "docker/setup-buildx-action": "8d2750c68a42422c14e847fe6c8ac0403b4cbd6f # v3.12.0",
  "docker/setup-qemu-action": "c7c53464625b32c7a7e944ae62b3e17d2b600130 # v3.7.0",
  "actions/attest-build-provenance": "e8998f949152b193b063cb0ec769d69d929409be # v2.4.0",
};
const EXPECTED_PATHS = [
  "apps/wopi/**",
  "packages/**",
  "crates/**",
  "shared/**",
  "scripts/**",
  "bun.lock",
  "Cargo.lock",
  "package.json",
  "Cargo.toml",
  ".dockerignore",
  ".github/workflows/wopi-image.yml",
];
const RELEASE_CONTEXT_IF =
  "github.repository == 'openooxml/betteroffice' && github.event_name == 'push' && github.ref == 'refs/heads/main' && github.event.repository.fork == false && startsWith(github.workflow_ref, format('{0}/.github/workflows/release.yml@refs/heads/main', github.repository))";
const PUBLISH_IF =
  "${{ inputs.publish && inputs.sha != '' && github.repository == 'openooxml/betteroffice' && github.event_name == 'workflow_dispatch' && github.ref == 'refs/heads/main' && github.event.repository.fork == false && vars.WOPI_IMAGE_PUBLISH == 'true' }}";
const RELEASE_WOPI_IF =
  "needs.release.result == 'success' && needs.release.outputs.published == 'true' && " +
  RELEASE_CONTEXT_IF + " && vars.WOPI_IMAGE_PUBLISH == 'true'";
const VERSION_IMAGE = "ghcr.io/openooxml/betteroffice-wopi:0.1.0";

function inspectVersionTag({ status = 0, stdout = "", stderr = "" } = {}) {
  const workflow = loadWorkflow(wopiWorkflowPath);
  const check = workflow.jobs.publish!.steps!.find((step) => step.id === "exists")!;
  const directory = mkdtempSync(join(tmpdir(), "wopi-image-inspect-"));
  try {
    const output = join(directory, "output");
    const call = join(directory, "docker-call");
    writeFileSync(output, "");
    writeFileSync(
      join(directory, "docker"),
      [
        "#!/bin/sh",
        'printf "%s\\n" "$*" > "$DOCKER_CALL"',
        'printf "%s" "$INSPECT_STDOUT"',
        'printf "%s" "$INSPECT_STDERR" >&2',
        'exit "$INSPECT_STATUS"',
      ].join("\n"),
      { mode: 0o755 },
    );
    const run = check.run!
      .replaceAll("${{ github.repository_owner }}", "openooxml")
      .replaceAll("${{ steps.version.outputs.version }}", "0.1.0");
    const result = spawnSync("bash", ["--noprofile", "--norc", "-e", "-o", "pipefail", "-c", run], {
      encoding: "utf8",
      env: {
        ...process.env,
        PATH: `${directory}:${process.env.PATH}`,
        GITHUB_OUTPUT: output,
        DOCKER_CALL: call,
        INSPECT_STDOUT: stdout,
        INSPECT_STDERR: stderr,
        INSPECT_STATUS: String(status),
      },
    });
    expect(result.error).toBeUndefined();
    expect(readFileSync(call, "utf8").trim()).toBe(`buildx imagetools inspect ${VERSION_IMAGE}`);
    return { ...result, output: readFileSync(output, "utf8") };
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}

function checkCurrentVersion(version: string, mainVersion: string) {
  const workflow = loadWorkflow(wopiWorkflowPath);
  const check = workflow.jobs.publish!.steps!.find((step) => step.id === "current")!;
  const directory = mkdtempSync(join(tmpdir(), "wopi-image-current-"));
  try {
    writeFileSync(
      join(directory, "gh"),
      ["#!/bin/sh", `printf '%s' '${JSON.stringify({ version: mainVersion })}'`].join("\n"),
      { mode: 0o755 },
    );
    return spawnSync("bash", ["--noprofile", "--norc", "-e", "-o", "pipefail", "-c", check.run!], {
      encoding: "utf8",
      env: { ...process.env, PATH: `${directory}:${process.env.PATH}`, GH_REPO: "o/r", VERSION: version },
    });
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}

describe("wopi-image workflow", () => {
  test("triggers on pull requests and main pushes with the image paths", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    const on = workflow.on as Record<string, { branches?: string[]; paths?: string[] }>;
    expect(on.pull_request.paths).toEqual(EXPECTED_PATHS);
    expect(on.push.branches).toEqual(["main"]);
    expect(on.push.paths).toEqual(EXPECTED_PATHS);
    const dispatch = on.workflow_dispatch as { inputs?: Record<string, { type?: string }> };
    expect(dispatch.inputs?.publish.type).toBe("boolean");
    expect(dispatch.inputs?.sha.type).toBe("string");
  });

  test("never pushes on the pull request or main push build path", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    const build = workflow.jobs.build!;
    expect(build.if).toBe("${{ !inputs.publish }}");
    expect(build.if).not.toContain("||");
    const builds = buildPushSteps(build);
    expect(builds).toHaveLength(1);
    expect(builds[0]!.with?.push).toBe(false);
    expect(builds[0]!.with?.file).toBe("apps/wopi/Dockerfile");
    expect(builds[0]!.with?.context).toBe(".");
    const smoke = (build.steps ?? []).find((step) => step.run?.includes("/healthz"));
    expect(smoke).toBeDefined();
    expect(smoke!.run).toContain("/hosting/discovery");
    expect(smoke!.run).toContain("proof-key");
    expect(smoke!.run).toContain("WOPI_PUBLIC_URL");
    expect(smoke!.run).toContain("WOPI_HOSTS");
    expect(smoke!.run).toMatch(/-v\s+"\$[A-Za-z_]+:\/data"/);
    expect(JSON.stringify(build)).not.toContain("packages: write");
    expect(JSON.stringify(build)).not.toContain("ghcr.io");
  });

  test("publishes multi-arch to ghcr only behind the repository variable", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    const publish = workflow.jobs.publish!;
    expect(publish.if).toBe(PUBLISH_IF);
    expect(publish.if).toContain("github.repository == 'openooxml/betteroffice'");
    expect(publish.if).not.toContain("||");
    const builds = buildPushSteps(publish);
    expect(builds).toHaveLength(1);
    const built = builds[0]!;
    expect(built.with?.push).toBe(true);
    expect(built.with?.platforms).toBe("linux/amd64,linux/arm64");
    expect(String(built.with?.tags)).toContain(
      "ghcr.io/${{ github.repository_owner }}/betteroffice-wopi:${{ steps.version.outputs.version }}",
    );
    expect(String(built.with?.tags)).toContain(
      "ghcr.io/${{ github.repository_owner }}/betteroffice-wopi:latest",
    );
    expect(String(built.with?.labels)).toContain("org.opencontainers.image.source=");
    expect(String(built.with?.labels)).toContain("org.opencontainers.image.revision=${{ inputs.sha }}");
    expect((publish.steps ?? [])[0]!.with?.ref).toBe("${{ inputs.sha }}");
    expect(String(built.with?.labels)).toContain("org.opencontainers.image.licenses=Apache-2.0");
    expect(built.with?.sbom).toBe(true);
    expect(String(built.with?.provenance)).toContain("mode=max");
    expect((publish.steps ?? []).some((step) => step.uses?.startsWith("docker/login-action@"))).toBe(true);
    expect((publish.steps ?? []).some((step) => step.uses?.startsWith("docker/setup-qemu-action@"))).toBe(true);
    const version = (publish.steps ?? []).find((step) => step.run?.includes("apps/wopi/package.json"));
    expect(version).toBeDefined();
  });

  test("skips the push when the version tag already exists", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    const publish = workflow.jobs.publish!;
    const steps = publish.steps ?? [];
    const check = steps.find((step) => step.id === "exists");
    expect(check).toBeDefined();
    expect(check!.run).toContain("docker buildx imagetools inspect");
    expect(check!.run).toContain(
      "ghcr.io/${{ github.repository_owner }}/betteroffice-wopi:${{ steps.version.outputs.version }}",
    );
    expect(check!.run).toContain('echo "exists=true" >> "$GITHUB_OUTPUT"');
    expect(check!.run).toContain('echo "exists=false" >> "$GITHUB_OUTPUT"');
    const build = steps.find((step) => step.id === "build");
    expect(build).toBeDefined();
    expect(build!.if).toBe("steps.exists.outputs.exists == 'false'");
    const attest = steps.find((step) => step.uses?.startsWith("actions/attest-build-provenance@"));
    expect(attest).toBeDefined();
    expect(attest!.if).toBe("steps.exists.outputs.exists == 'false'");
    expect(steps.indexOf(check!)).toBeLessThan(steps.indexOf(build!));
  });

  test("marks a successful inspection as an existing version", () => {
    const result = inspectVersionTag({ stdout: `Name: ${VERSION_IMAGE}\n` });
    expect(result.status).toBe(0);
    expect(result.output).toBe("exists=true\n");
  });

  test("permits a push only after the version manifest is confirmed missing", () => {
    const result = inspectVersionTag({ status: 1, stderr: `ERROR: ${VERSION_IMAGE}: not found\n` });
    expect(result.status).toBe(0);
    expect(result.output).toBe("exists=false\n");
  });

  test.each([
    ["authentication", "ERROR: failed to authorize: 401 Unauthorized"],
    ["authorization", "ERROR: unexpected status from HEAD request: 403 Forbidden"],
    ["registry server", "ERROR: unexpected status from HEAD request: 500 Internal Server Error"],
    ["registry outage", "ERROR: unexpected status from HEAD request: 503 Service Unavailable"],
    ["network", "ERROR: failed to do request: dial tcp: i/o timeout"],
    ["DNS", "ERROR: lookup ghcr.io: no such host"],
    ["unrelated not found", "ERROR: credential helper not found"],
    ["generic 404", "ERROR: 404 Not Found"],
    ["wrong version", "ERROR: ghcr.io/openooxml/betteroffice-wopi:0.0.1: not found"],
    ["mixed errors", `ERROR: ${VERSION_IMAGE}: not found\nERROR: 503 Service Unavailable`],
    ["unexpected output", "unexpected registry response"],
    ["empty output", ""],
  ])("fails closed on %s inspection failures", (_name, stderr) => {
    const result = inspectVersionTag({ status: 1, stderr });
    expect(result.status).toBe(1);
    expect(result.output).toBe("");
    expect(result.stdout).toContain("::error::Could not confirm whether version tag");
  });

  test("rejects unexpected output alongside a manifest-not-found error", () => {
    const result = inspectVersionTag({
      status: 1,
      stdout: "unexpected output\n",
      stderr: `ERROR: ${VERSION_IMAGE}: not found\n`,
    });
    expect(result.status).toBe(1);
    expect(result.output).toBe("");
  });

  test("refuses a version main has moved past before anything can be pushed", () => {
    expect(checkCurrentVersion("0.2.0", "0.2.0").status).toBe(0);
    const stale = checkCurrentVersion("0.1.0", "0.2.0");
    expect(stale.status).toBe(1);
    expect(stale.stdout).toContain("latest would move backward");
    const steps = loadWorkflow(wopiWorkflowPath).jobs.publish!.steps!;
    const current = steps.findIndex((step) => step.id === "current");
    expect(current).toBeGreaterThan(-1);
    expect(current).toBeLessThan(steps.findIndex((step) => step.id === "exists"));
    expect(current).toBeLessThan(steps.findIndex((step) => step.id === "build"));
  });

  test("a publish waits on approval in its own concurrency group", () => {
    const source = Bun.YAML.parse(readFileSync(wopiWorkflowPath, "utf8")) as {
      concurrency: { group: string; "cancel-in-progress": string };
    };
    expect(source.concurrency.group).toContain("inputs.publish && 'publish'");
    expect(source.concurrency.group).not.toBe("release");
    expect(source.concurrency["cancel-in-progress"]).toBe("${{ github.event_name == 'pull_request' }}");
  });

  test("publishes only after approval of the ghcr-wopi-image environment", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    const jobs = workflow.jobs as Record<string, { environment?: unknown; permissions?: { packages?: string } }>;
    expect(jobs.publish!.environment).toBe("ghcr-wopi-image");
    for (const [name, job] of Object.entries(jobs)) {
      if (job.permissions?.packages === "write" && name !== "publish") throw new Error(`${name} has packages: write`);
    }
  });

  test("keeps permissions minimal per job", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    expect(workflow.permissions).toEqual({ contents: "read" });
    expect(workflow.jobs.build!.permissions).toEqual({ contents: "read" });
    expect(workflow.jobs.publish!.permissions).toEqual({
      contents: "read",
      packages: "write",
      "id-token": "write",
      attestations: "write",
    });
    for (const [name, job] of Object.entries(workflow.jobs)) {
      if (name === "publish") continue;
      expect(job.permissions?.packages).toBeUndefined();
      expect(job.permissions?.["id-token"]).toBeUndefined();
      expect(job.permissions?.attestations).toBeUndefined();
    }
  });

  test("pins every uses line to the required full SHA with a version comment", () => {
    const source = readFileSync(wopiWorkflowPath, "utf8");
    const uses = source.split("\n").filter((line) => /^\s*(?:-\s*)?uses:/.test(line));
    expect(uses.length).toBeGreaterThan(0);
    for (const line of uses) {
      const action = line.trim().replace(/^(?:-\s*)?uses:\s*/, "");
      expect(action.split(/\s/)[0]).toMatch(PINNED);
      const name = action.split("@")[0]!;
      expect(ACTION_PINS[name]).toBeDefined();
      expect(action).toBe(`${name}@${ACTION_PINS[name]}`);
    }
  });

  test("contains no token logging and no em or en dashes", () => {
    for (const path of [wopiWorkflowPath, releaseWorkflowPath]) {
      const source = readFileSync(path, "utf8");
      expect(source.includes("\u2014")).toBe(false);
      expect(source.includes("\u2013")).toBe(false);
      expect(source).not.toMatch(/echo[^\n]*secrets\./);
    }
  });
});

describe("release workflow WOPI image job", () => {
  test("calls the publish path only after publishing and only behind the variable", () => {
    const workflow = loadWorkflow(releaseWorkflowPath);
    const caller = workflow.jobs["wopi-image"]!;
    expect(Object.values(workflow.jobs).some((job) => job.uses?.includes("wopi-image.yml"))).toBe(false);
    const needs = Array.isArray(caller.needs) ? caller.needs : [caller.needs];
    expect(needs).toContain("release");
    expect(workflow.jobs.release!.outputs).toEqual({
      publishing: "${{ steps.pending.outputs.publishing }}",
      published: "${{ steps.changesets.outputs.published }}",
    });
    expect(caller.if).toBe(RELEASE_WOPI_IF);
    expect(caller.if).not.toContain("||");
    expect(caller.if).not.toContain("outputs.publishing");
    expect(caller.if).toContain(RELEASE_CONTEXT_IF);
    expect(caller.permissions).toEqual({ actions: "write" });
    const dispatch = caller.steps![0]!;
    expect(dispatch.run).toContain("actions/workflows/wopi-image.yml/dispatches");
    expect(dispatch.run).toContain('-F "inputs[publish]=true"');
    expect(dispatch.run).toContain('-f "inputs[sha]=$SHA"');
    expect(dispatch.env?.SHA).toBe("${{ github.sha }}");
    expect(JSON.stringify(caller)).not.toContain("gh run watch");
  });
});

describe("wopi package version", () => {
  test("stays private with a real version for the image tags", () => {
    const pkg = JSON.parse(readFileSync(join(repoRoot, "apps", "wopi", "package.json"), "utf8")) as {
      name: string;
      private?: boolean;
      version: string;
    };
    expect(pkg.name).toBe("@betteroffice/wopi");
    expect(pkg.private).toBe(true);
    expect(pkg.version).toMatch(/^\d+\.\d+\.\d+$/);
    expect(pkg.version).not.toBe("0.0.0");
    const lock = readFileSync(join(repoRoot, "bun.lock"), "utf8");
    expect(lock).toContain(
      `"apps/wopi": {\n      "name": "@betteroffice/wopi",\n      "version": "${pkg.version}",`,
    );
  });
});
