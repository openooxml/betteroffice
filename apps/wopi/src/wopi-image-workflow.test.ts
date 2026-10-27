import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
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

function stepsWithUses(job: Job): Step[] {
  return (job.steps ?? []).filter((step) => typeof step.uses === "string");
}

function buildPushSteps(job: Job): Step[] {
  return (job.steps ?? []).filter((step) => step.uses?.startsWith("docker/build-push-action@"));
}

const PINNED = /^[^@\s]+@(?:v\d+(?:\.\d+){0,2}(?:[-.][0-9A-Za-z.-]+)?|[0-9a-f]{40})$/;
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
const PUBLISH_IF =
  "${{ inputs.publish && github.event_name == 'workflow_call' && github.ref == 'refs/heads/main' && vars.WOPI_IMAGE_PUBLISH == 'true' }}";
const RELEASE_WOPI_IF =
  "needs.release.result == 'success' && needs.release.outputs.published == 'true' && vars.WOPI_IMAGE_PUBLISH == 'true'";

describe("wopi-image workflow", () => {
  test("triggers on pull requests and main pushes with the image paths", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    const on = workflow.on as Record<string, { branches?: string[]; paths?: string[] }>;
    expect(on.pull_request.paths).toEqual(EXPECTED_PATHS);
    expect(on.push.branches).toEqual(["main"]);
    expect(on.push.paths).toEqual(EXPECTED_PATHS);
    const workflowCall = on.workflow_call as { inputs?: Record<string, { type?: string }> };
    expect(workflowCall.inputs?.publish.type).toBe("boolean");
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
    expect(String(built.with?.labels)).toContain("org.opencontainers.image.revision=${{ github.sha }}");
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
    expect(build!.if).toBe("steps.exists.outputs.exists != 'true'");
    const attest = steps.find((step) => step.uses?.startsWith("actions/attest-build-provenance@"));
    expect(attest).toBeDefined();
    expect(attest!.if).toBe("steps.exists.outputs.exists != 'true'");
    expect(steps.indexOf(check!)).toBeLessThan(steps.indexOf(build!));
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

  test("pins every action to a version tag or SHA", () => {
    const workflow = loadWorkflow(wopiWorkflowPath);
    for (const job of Object.values(workflow.jobs)) {
      for (const step of stepsWithUses(job)) {
        if (step.uses!.startsWith("./")) continue;
        expect(step.uses).toMatch(PINNED);
      }
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
    const callers = Object.values(workflow.jobs).filter(
      (job) => job.uses === "./.github/workflows/wopi-image.yml",
    );
    expect(callers).toHaveLength(1);
    const caller = callers[0]!;
    const needs = Array.isArray(caller.needs) ? caller.needs : [caller.needs];
    expect(needs).toContain("release");
    expect(workflow.jobs.release!.outputs).toEqual({
      publishing: "${{ steps.pending.outputs.publishing }}",
      published: "${{ steps.changesets.outputs.published }}",
    });
    expect(caller.if).toBe(RELEASE_WOPI_IF);
    expect(caller.if).not.toContain("||");
    expect(caller.if).not.toContain("outputs.publishing");
    expect(caller.with?.publish).toBe(true);
    expect(caller.permissions).toEqual({
      contents: "read",
      packages: "write",
      "id-token": "write",
      attestations: "write",
    });
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
