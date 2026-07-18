/**
 * Minimal structural XML seam for DrawingML parsing.
 *
 * This package deliberately has no XML-parser dependency. Format hosts
 * (docx/pptx/xlsx) parse XML however they like and pass in any tree whose
 * nodes satisfy {@link XmlLike} structurally — an xml-js `Element` tree
 * (non-compact mode) satisfies it out of the box, with no adapter objects.
 */

/**
 * Structural interface for an XML node. Mirrors the xml-js non-compact
 * `Element` shape: element nodes carry `type: 'element'`, a prefixed `name`
 * (e.g. `a:srgbClr`), an attribute map, and interleaved child nodes.
 *
 * @public
 */
export interface XmlLike {
  /** Node kind marker; element nodes are `'element'` (xml-js convention) */
  type?: string;
  /** Element name including namespace prefix (e.g. `a:srgbClr`) */
  name?: string;
  /** Attribute map keyed by (possibly prefixed) attribute name */
  attributes?: Record<string, string | number | undefined>;
  /** Child nodes (elements and text nodes interleaved) */
  elements?: XmlLike[];
}

/**
 * Get local name from a prefixed element name
 * e.g., "w:p" -\> "p", "a:graphic" -\> "graphic"
 *
 * @public
 */
export function getLocalName(name: string): string {
  const colonIndex = name.indexOf(':');
  return colonIndex >= 0 ? name.substring(colonIndex + 1) : name;
}

/**
 * Get all child elements (excludes text nodes, etc.)
 *
 * @public
 */
export function getChildElements(parent: XmlLike | null | undefined): XmlLike[] {
  if (!parent || !parent.elements) return [];
  return parent.elements.filter((child) => child.type === 'element');
}

/**
 * Get an attribute value from an element, trying the namespaced name first
 * and falling back to the bare name.
 *
 * @public
 */
export function getAttribute(
  element: XmlLike | null | undefined,
  namespace: string | null,
  name: string
): string | null {
  if (!element || !element.attributes) return null;

  const attrs = element.attributes;

  if (namespace) {
    const prefixedName = `${namespace}:${name}`;
    if (prefixedName in attrs) {
      const value = attrs[prefixedName];
      return value == null ? null : String(value);
    }
  }

  if (name in attrs) {
    const value = attrs[name];
    return value == null ? null : String(value);
  }

  return null;
}

/**
 * Find first child element matching the given namespaced name (also matches
 * on bare local name).
 *
 * @public
 */
export function findChild(
  parent: XmlLike | null | undefined,
  namespace: string,
  localName: string
): XmlLike | null {
  if (!parent || !parent.elements) return null;

  const fullName = `${namespace}:${localName}`;

  for (const child of parent.elements) {
    if (child.type !== 'element') continue;

    if (child.name === fullName) {
      return child;
    }

    if (getLocalName(child.name || '') === localName) {
      return child;
    }
  }

  return null;
}

/**
 * Find all child elements matching the given namespaced name (also matches
 * on bare local name).
 *
 * @public
 */
export function findChildren(
  parent: XmlLike | null | undefined,
  namespace: string,
  localName: string
): XmlLike[] {
  if (!parent || !parent.elements) return [];

  const fullName = `${namespace}:${localName}`;
  const results: XmlLike[] = [];

  for (const child of parent.elements) {
    if (child.type !== 'element') continue;

    if (child.name === fullName || getLocalName(child.name || '') === localName) {
      results.push(child);
    }
  }

  return results;
}

/**
 * Find all child elements by local name only (ignoring namespace).
 *
 * @public
 */
export function findChildrenByLocalName(
  parent: XmlLike | null | undefined,
  localName: string
): XmlLike[] {
  if (!parent || !parent.elements) return [];

  return parent.elements.filter(
    (child) => child.type === 'element' && getLocalName(child.name || '') === localName
  );
}

/**
 * Find first child element by full name (including namespace prefix).
 *
 * @public
 */
export function findByFullName(
  parent: XmlLike | null | undefined,
  fullName: string
): XmlLike | null {
  if (!parent || !parent.elements) return null;

  for (const child of parent.elements) {
    if (child.type !== 'element') continue;
    if (child.name === fullName) return child;
  }

  return null;
}

/**
 * Parse a numeric value from an attribute, with optional scale.
 *
 * @public
 */
export function parseNumericAttribute(
  element: XmlLike | null | undefined,
  namespace: string | null,
  name: string,
  scale: number = 1
): number | undefined {
  const value = getAttribute(element, namespace, name);
  if (value === null) return undefined;

  const num = parseInt(value, 10);
  if (isNaN(num)) return undefined;

  return num * scale;
}
