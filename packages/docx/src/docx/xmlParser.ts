/**
 * Thin xml-js compatibility helpers for published XmlElement leaf APIs.
 *
 * Rust owns all package XML parsing. This module remains solely because the
 * stable ./docx API accepts xml-js Element values.
 */

import { xml2js, type Element as XmlElement } from 'xml-js';

export type { Element as XmlElement } from 'xml-js';

const STRAY_AMPERSAND_RE = /&(?!(?:[A-Za-z][A-Za-z0-9]*|#[0-9]+|#x[0-9a-fA-F]+);)/g;

export function parseXml(xml: string): XmlElement {
  const sanitized = xml.replace(STRAY_AMPERSAND_RE, '&amp;');
  try {
    return xml2js(sanitized, {
      compact: false,
      ignoreComment: true,
      ignoreInstruction: true,
      ignoreDoctype: true,
      alwaysArray: false,
      trim: false,
      captureSpacesBetweenElements: true,
      attributesKey: 'attributes',
      textKey: 'text',
    }) as XmlElement;
  } catch (error) {
    if (!(error instanceof Error)) throw error;
    const column = error.message.match(/Column:\s*(\d+)/)?.[1];
    if (!column) throw error;
    const offset = Number(column);
    const snippet = JSON.stringify(sanitized.slice(Math.max(0, offset - 30), offset + 30));
    const wrapped = new Error(`${error.message}\nNear: ${snippet}`);
    wrapped.stack = error.stack;
    throw wrapped;
  }
}

export function parseXmlDocument(xml: string): XmlElement | null {
  try {
    const parsed = parseXml(xml);
    return parsed.elements?.find((element) => element.type === 'element') ?? parsed;
  } catch (error) {
    console.warn('Failed to parse XML:', error);
    return null;
  }
}

export function getLocalName(name: string): string {
  const colon = name.indexOf(':');
  return colon >= 0 ? name.slice(colon + 1) : name;
}

export function findChild(
  parent: XmlElement | null | undefined,
  namespace: string,
  localName: string
): XmlElement | null {
  const fullName = `${namespace}:${localName}`;
  return (
    getChildElements(parent).find(
      (child) => child.name === fullName || getLocalName(child.name ?? '') === localName
    ) ?? null
  );
}

export function findChildren(
  parent: XmlElement | null | undefined,
  namespace: string,
  localName: string
): XmlElement[] {
  const fullName = `${namespace}:${localName}`;
  return getChildElements(parent).filter(
    (child) => child.name === fullName || getLocalName(child.name ?? '') === localName
  );
}

export function findChildrenByLocalName(
  parent: XmlElement | null | undefined,
  localName: string
): XmlElement[] {
  return getChildElements(parent).filter(
    (child) => getLocalName(child.name ?? '') === localName
  );
}

export function findByFullName(
  parent: XmlElement | null | undefined,
  fullName: string
): XmlElement | null {
  return getChildElements(parent).find((child) => child.name === fullName) ?? null;
}

export function getChildElements(parent: XmlElement | null | undefined): XmlElement[] {
  return parent?.elements?.filter((child) => child.type === 'element') ?? [];
}

export function getAttribute(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string
): string | null {
  if (!element?.attributes) return null;
  const attributes = element.attributes as Record<string, string>;
  if (namespace && `${namespace}:${name}` in attributes) {
    return attributes[`${namespace}:${name}`];
  }
  return name in attributes ? attributes[name] : null;
}

export function getTextContent(element: XmlElement | null | undefined): string {
  if (!element) return '';
  if (typeof element.text === 'string') return element.text;
  return (element.elements ?? [])
    .map((child) => {
      if (child.type === 'text') return typeof child.text === 'string' ? child.text : '';
      return child.type === 'element' ? getTextContent(child) : '';
    })
    .join('');
}

export function parseNumericAttribute(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string,
  scale = 1
): number | undefined {
  const raw = getAttribute(element, namespace, name);
  if (raw === null) return undefined;
  const value = parseInt(raw, 10);
  return Number.isNaN(value) ? undefined : value * scale;
}

export function parseBooleanElement(
  element: XmlElement | null | undefined,
  namespace = 'w'
): boolean {
  if (!element) return false;
  const value = getAttribute(element, namespace, 'val');
  return value !== '0' && value !== 'false' && value !== 'off';
}
