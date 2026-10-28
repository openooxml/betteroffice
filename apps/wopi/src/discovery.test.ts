import { readFileSync } from "node:fs";
import { describe, expect, test } from "bun:test";
import { buildDiscovery, capabilities } from "./discovery";
import { generateKeyPair, proofKeyAttributes } from "./proof";

function decodeEntities(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function parseAttributes(tag: string): Record<string, string> {
  const attributes: Record<string, string> = {};
  for (const match of tag.matchAll(/([\w-]+)="([^"]*)"/g)) {
    attributes[match[1]!] = decodeEntities(match[2]!);
  }
  return attributes;
}

interface ParsedApp {
  name: string;
  favIconUrl?: string;
  actions: Array<Record<string, string>>;
}

function parseDiscovery(xml: string): { apps: ParsedApp[]; proofKey: Record<string, string> } {
  const apps: ParsedApp[] = [];
  for (const match of xml.matchAll(/<app\s([^>]*)>([\s\S]*?)<\/app>/g)) {
    const appAttributes = parseAttributes(match[1]!);
    const actions = [...match[2]!.matchAll(/<action\s([^>]*?)\/>/g)].map((action) =>
      parseAttributes(action[1]!),
    );
    apps.push({ name: appAttributes.name!, favIconUrl: appAttributes.favIconUrl, actions });
  }
  const proofMatch = xml.match(/<proof-key\s([^>]*?)\/>/);
  expect(proofMatch).not.toBeNull();
  return { apps, proofKey: parseAttributes(proofMatch![1]!) };
}

const publicUrl = "https://office.example.com";
const formats = [
  { app: "Word", ext: "docx", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
  { app: "Excel", ext: "xlsx", mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  { app: "PowerPoint", ext: "pptx", mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
] as const;

describe("buildDiscovery", () => {
  test("has view, edit and editnew for every format in both spellings", () => {
    const xml = buildDiscovery({ publicUrl, currentKey: generateKeyPair().publicKey });
    const { apps } = parseDiscovery(xml);
    expect(apps).toHaveLength(6);

    for (const format of formats) {
      const msApp = apps.find((app) => app.name === format.app);
      expect(msApp).toBeDefined();
      expect(msApp!.favIconUrl).toBe(`${publicUrl}/favicon.ico`);
      expect(msApp!.actions.map((action) => action.name).sort()).toEqual(["edit", "editnew", "view"]);
      for (const action of msApp!.actions) {
        expect(action.ext).toBe(format.ext);
      }

      const collaboraApp = apps.find((app) => app.name === format.mime);
      expect(collaboraApp).toBeDefined();
      expect(collaboraApp!.actions.map((action) => action.name).sort()).toEqual(["edit", "editnew", "view"]);
      for (const action of collaboraApp!.actions) {
        expect(action.ext).toBe("");
      }
    }
  });

  test("MIME-named apps list edit first, because Nextcloud richdocuments takes the first action", () => {
    const xml = buildDiscovery({ publicUrl, currentKey: generateKeyPair().publicKey });
    const { apps } = parseDiscovery(xml);
    for (const format of formats) {
      const collaboraApp = apps.find((app) => app.name === format.mime)!;
      expect(collaboraApp.actions[0]!.name).toBe("edit");
      expect(collaboraApp.actions[0]!.urlsrc).toContain("/wopi/edit?");
    }
  });

  test("builds urlsrc placeholders and requires attributes", () => {
    const xml = buildDiscovery({ publicUrl, currentKey: generateKeyPair().publicKey });
    const { apps } = parseDiscovery(xml);

    for (const app of apps) {
      for (const action of app.actions) {
        const endpoint = action.name === "view" ? "view" : "edit";
        expect(action.urlsrc).toBe(
          `${publicUrl}/wopi/${endpoint}?<ui=UI_LLCC&><rs=DC_LLCC&><dchat=DISABLE_CHAT&>`,
        );
        if (action.name === "view") {
          expect(action.requires).toBeUndefined();
        } else {
          expect(action.requires).toBe("locks,update");
        }
      }
    }
    expect(xml).toContain("<net-zone name=\"external-https\">");
  });

  test("embeds current and old proof keys", () => {
    const current = generateKeyPair();
    const old = generateKeyPair();
    const xml = buildDiscovery({
      publicUrl,
      currentKey: current.publicKey,
      oldKey: old.publicKey,
    });
    const { proofKey } = parseDiscovery(xml);
    const currentAttributes = proofKeyAttributes(current.publicKey);
    const oldAttributes = proofKeyAttributes(old.publicKey);
    expect(proofKey.value).toBe(currentAttributes.value);
    expect(proofKey.modulus).toBe(currentAttributes.modulus);
    expect(proofKey.exponent).toBe(currentAttributes.exponent);
    expect(proofKey.oldvalue).toBe(oldAttributes.value);
    expect(proofKey.oldmodulus).toBe(oldAttributes.modulus);
    expect(proofKey.oldexponent).toBe(oldAttributes.exponent);
  });

  test("publishes the current key as the old key without an old key", () => {
    const current = generateKeyPair();
    const xml = buildDiscovery({ publicUrl, currentKey: current.publicKey });
    const { proofKey } = parseDiscovery(xml);
    const attributes = proofKeyAttributes(current.publicKey);
    expect(proofKey.oldvalue).toBe(attributes.value);
    expect(proofKey.oldmodulus).toBe(attributes.modulus);
    expect(proofKey.oldexponent).toBe(attributes.exponent);
  });

  test("escapes special characters in attribute values", () => {
    const trickyUrl = "https://office.example.com/a&b\"c<d>";
    const xml = buildDiscovery({ publicUrl: trickyUrl, currentKey: generateKeyPair().publicKey });
    expect(xml).not.toContain("a&b\"c<d>");
    expect(xml).toContain("a&amp;b&quot;c&lt;d&gt;");
    const { apps } = parseDiscovery(xml);
    expect(apps[0]!.favIconUrl).toBe(`${trickyUrl}/favicon.ico`);
    expect(apps[0]!.actions[0]!.urlsrc).toContain(trickyUrl);
  });
});

describe("capabilities", () => {
  test("returns the minimal honest capability object", () => {
    const result = capabilities();
    const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
      version: string;
    };
    expect(result.productName).toBe("BetterOffice");
    expect(result.productVersion).toBe(pkg.version);
    expect(result.hasTemplateSaveAs).toBe(false);
    expect(result.hasTemplateSource).toBe(false);
    expect(result.hasWOPIExtendedLockLength).toBe(false);
    for (const [key, value] of Object.entries(result)) {
      if (key.startsWith("has") || key.startsWith("enable")) {
        expect(value).toBe(false);
      }
    }
  });
});
