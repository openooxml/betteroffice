import { readFileSync } from "node:fs";
import { proofKeyAttributes, type ProofKeyInput } from "./proof";

export interface DiscoveryOptions {
  publicUrl: string;
  currentKey: ProofKeyInput;
  oldKey?: ProofKeyInput;
  favIconUrl?: string;
}

export interface HostingCapabilities {
  productName: string;
  productVersion: string;
  hasTemplateSaveAs: boolean;
  hasTemplateSource: boolean;
  hasWOPIExtendedLockLength: boolean;
  hasMobileSupport: boolean;
  hasProxyPrefix: boolean;
  hasSettingIframeSupport: boolean;
  hasWOPIFilePostMessage: boolean;
  hasInsertRemoteImage: boolean;
  hasInsertRemoteFile: boolean;
  hasDocumentSigningSupport: boolean;
  enableDrawing: boolean;
  [key: string]: string | boolean;
}

const FORMATS = [
  {
    appName: "Word",
    ext: "docx",
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  {
    appName: "Excel",
    ext: "xlsx",
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  },
  {
    appName: "PowerPoint",
    ext: "pptx",
    mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  },
] as const;

function escapeXmlAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function actionXml(publicUrl: string, name: "view" | "edit" | "editnew", ext: string): string {
  const endpoint = name === "view" ? "view" : "edit";
  const urlsrc = `${publicUrl}/wopi/${endpoint}?<ui=UI_LLCC&><rs=DC_LLCC&>`;
  const requires = name === "view" ? "" : ` requires="locks,update"`;
  return `      <action name="${name}" ext="${escapeXmlAttribute(ext)}"${requires} urlsrc="${escapeXmlAttribute(urlsrc)}"/>`;
}

function appXml(
  publicUrl: string,
  favIconUrl: string,
  appName: string,
  ext: string,
): string {
  const actions = (["view", "edit", "editnew"] as const)
    .map((name) => actionXml(publicUrl, name, ext))
    .join("\n");
  return [
    `    <app name="${escapeXmlAttribute(appName)}" favIconUrl="${escapeXmlAttribute(favIconUrl)}">`,
    actions,
    "    </app>",
  ].join("\n");
}

/** WOPI discovery XML with both MS and Collabora app spellings. */
export function buildDiscovery(options: DiscoveryOptions): string {
  const publicUrl = options.publicUrl.replace(/\/+$/, "");
  const favIconUrl = options.favIconUrl ?? `${publicUrl}/favicon.ico`;
  const current = proofKeyAttributes(options.currentKey);
  const apps = FORMATS.flatMap((format) => [
    appXml(publicUrl, favIconUrl, format.appName, format.ext),
    appXml(publicUrl, favIconUrl, format.mime, ""),
  ]).join("\n");
  const old = options.oldKey ? proofKeyAttributes(options.oldKey) : current;
  const oldAttributes = ` oldvalue="${escapeXmlAttribute(old.value)}" oldmodulus="${escapeXmlAttribute(old.modulus)}" oldexponent="${escapeXmlAttribute(old.exponent)}"`;
  return [
    '<?xml version="1.0" encoding="utf-8"?>',
    "<wopi-discovery>",
    '  <net-zone name="external-https">',
    apps,
    "  </net-zone>",
    `  <proof-key value="${escapeXmlAttribute(current.value)}" modulus="${escapeXmlAttribute(current.modulus)}" exponent="${escapeXmlAttribute(current.exponent)}"${oldAttributes}/>`,
    "</wopi-discovery>",
    "",
  ].join("\n");
}

function packageVersion(): string {
  try {
    const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
      version?: string;
    };
    return pkg.version ?? "0.0.1";
  } catch {
    return "0.0.1";
  }
}

export const PRODUCT_VERSION = packageVersion();

/** Minimal capabilities object for GET /hosting/capabilities. */
export function capabilities(productVersion: string = PRODUCT_VERSION): HostingCapabilities {
  return {
    productName: "BetterOffice",
    productVersion,
    hasTemplateSaveAs: false,
    hasTemplateSource: false,
    hasWOPIExtendedLockLength: false,
    hasMobileSupport: false,
    hasProxyPrefix: false,
    hasSettingIframeSupport: false,
    hasWOPIFilePostMessage: false,
    hasInsertRemoteImage: false,
    hasInsertRemoteFile: false,
    hasDocumentSigningSupport: false,
    enableDrawing: false,
  };
}
