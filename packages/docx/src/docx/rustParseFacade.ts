import type {
  Chart,
  Document,
  DocumentBody,
  DocxPackage,
  Endnote,
  Footnote,
  HeaderFooter,
  MediaFile,
  Relationship,
  RelationshipMap,
} from '../types/document';
import { parseDocxS9Wire, parseRelationshipsXmlWire } from './parseWasm';

const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

export interface RustS9ParseOptions {
  parseHeadersFooters: boolean;
  parseNotes: boolean;
  detectVariables: boolean;
  determinismSeed?: string;
  includeCanonical?: boolean;
}

export interface RustS9Result {
  wireVersion: 1;
  document: Document;
  embeddedFonts: Map<string, ArrayBuffer>;
  fontTableRelationshipsXml?: string;
  rustCanonicalBytes?: Uint8Array;
  rustCanonicalSha256?: string;
}

/** Production-capable S9 facade. Rehydrates maps, aliased buffers, and the public Document. */
export function parseDocumentWithRust(
  data: ArrayBuffer,
  options: RustS9ParseOptions
): RustS9Result {
  return decodeS9Envelope(parseDocxS9Wire(new Uint8Array(data), JSON.stringify(options)), data);
}

export function decodeS9Envelope(json: string, originalBuffer: ArrayBuffer): RustS9Result {
  const value: unknown = JSON.parse(json);
  assertSafeTree(value, 'wire');
  const envelope = objectAt(value, 'wire');
  exactKeys(
    envelope,
    [
      'wireVersion',
      'document',
      'embeddedFontParts',
      'fontTableRelationshipsXml',
      'canonicalBase64',
      'canonicalSha256',
    ],
    'wire',
    ['fontTableRelationshipsXml', 'canonicalBase64', 'canonicalSha256']
  );
  if (envelope.wireVersion !== 1) throw new TypeError('wire.wireVersion must be 1');
  const document = decodeS9Document(envelope.document, originalBuffer);
  const embeddedFonts = decodeBinaryParts(envelope.embeddedFontParts);
  const hasCanonical =
    envelope.canonicalBase64 !== undefined || envelope.canonicalSha256 !== undefined;
  if (
    hasCanonical &&
    (envelope.canonicalBase64 === undefined || envelope.canonicalSha256 === undefined)
  ) {
    throw new TypeError('wire canonical fields must be present together');
  }
  const sha =
    envelope.canonicalSha256 === undefined
      ? undefined
      : stringAt(envelope.canonicalSha256, 'wire.canonicalSha256');
  if (sha !== undefined && !/^[0-9a-f]{64}$/.test(sha)) {
    throw new TypeError('wire.canonicalSha256 is invalid');
  }
  return {
    wireVersion: 1,
    document,
    embeddedFonts,
    ...(envelope.fontTableRelationshipsXml === undefined
      ? {}
      : {
          fontTableRelationshipsXml: stringAt(
            envelope.fontTableRelationshipsXml,
            'wire.fontTableRelationshipsXml'
          ),
        }),
    ...(envelope.canonicalBase64 === undefined
      ? {}
      : {
          rustCanonicalBytes: decodeBase64(
            stringAt(envelope.canonicalBase64, 'wire.canonicalBase64')
          ),
          rustCanonicalSha256: sha,
        }),
  };
}

function decodeS9Document(value: unknown, originalBuffer: ArrayBuffer): Document {
  const wireDocument = objectAt(value, 'wire.document');
  exactKeys(wireDocument, ['package', 'templateVariables', 'warnings'], 'wire.document', [
    'templateVariables',
    'warnings',
  ]);
  const pkg = decodeS9Package(wireDocument.package);
  const document: Document = { package: pkg, originalBuffer };
  if (wireDocument.templateVariables !== undefined) {
    document.templateVariables = stringArrayAt(
      wireDocument.templateVariables,
      'wire.document.templateVariables'
    );
  }
  if (wireDocument.warnings !== undefined) {
    document.warnings = stringArrayAt(wireDocument.warnings, 'wire.document.warnings');
  }
  return document;
}

function decodeS9Package(value: unknown): DocxPackage {
  const wirePackage = objectAt(value, 'wire.document.package');
  const optional = [
    'styles',
    'headerEntries',
    'footerEntries',
    'footnotes',
    'endnotes',
    'footnoteSeparators',
    'endnoteSeparators',
  ];
  exactKeys(
    wirePackage,
    [
      'document',
      'styles',
      'theme',
      'numbering',
      'settings',
      'fontTable',
      'headerEntries',
      'footerEntries',
      'footnotes',
      'endnotes',
      'footnoteSeparators',
      'endnoteSeparators',
      'relationshipEntries',
      'mediaEntries',
      'chartEntries',
    ],
    'wire.document.package',
    optional
  );
  const documentBody = decodeS9DocumentBody(wirePackage.document);
  const pkg: DocxPackage = {
    document: documentBody,
    theme: objectAt(wirePackage.theme, 'wire.document.package.theme') as DocxPackage['theme'],
    numbering: objectAt(
      wirePackage.numbering,
      'wire.document.package.numbering'
    ) as unknown as DocxPackage['numbering'],
    settings: objectAt(
      wirePackage.settings,
      'wire.document.package.settings'
    ) as unknown as DocxPackage['settings'],
    fontTable: objectAt(
      wirePackage.fontTable,
      'wire.document.package.fontTable'
    ) as unknown as DocxPackage['fontTable'],
    relationships: decodeRelationshipEntries(
      wirePackage.relationshipEntries,
      'wire.document.package.relationshipEntries'
    ),
    media: decodeMediaEntries(wirePackage.mediaEntries),
    charts: decodeMapEntries<Chart>(
      wirePackage.chartEntries,
      'wire.document.package.chartEntries',
      (entry, path) => objectAt(entry, path) as unknown as Chart
    ),
  };
  if (wirePackage.styles !== undefined) {
    pkg.styles = objectAt(
      wirePackage.styles,
      'wire.document.package.styles'
    ) as unknown as DocxPackage['styles'];
  }
  for (const [wireName, publicName] of [
    ['headerEntries', 'headers'],
    ['footerEntries', 'footers'],
  ] as const) {
    const entries = wirePackage[wireName];
    if (entries !== undefined) {
      pkg[publicName] = decodeMapEntries(
        entries,
        `wire.document.package.${wireName}`,
        (entry, path) => objectAt(entry, path) as unknown as HeaderFooter
      );
    }
  }
  if (wirePackage.footnotes !== undefined) {
    pkg.footnotes = decodeObjectArray<Footnote>(
      wirePackage.footnotes,
      'wire.document.package.footnotes'
    );
  }
  if (wirePackage.endnotes !== undefined) {
    pkg.endnotes = decodeObjectArray<Endnote>(
      wirePackage.endnotes,
      'wire.document.package.endnotes'
    );
  }
  if (wirePackage.footnoteSeparators !== undefined) {
    pkg.footnoteSeparators = decodeObjectArray<Footnote>(
      wirePackage.footnoteSeparators,
      'wire.document.package.footnoteSeparators'
    );
  }
  if (wirePackage.endnoteSeparators !== undefined) {
    pkg.endnoteSeparators = decodeObjectArray<Endnote>(
      wirePackage.endnoteSeparators,
      'wire.document.package.endnoteSeparators'
    );
  }
  return pkg;
}

function decodeS9DocumentBody(value: unknown): DocumentBody {
  const path = 'wire.document.package.document';
  const body = objectAt(value, path);
  exactKeys(body, ['content', 'sections', 'finalSectionProperties', 'comments'], path, [
    'sections',
    'finalSectionProperties',
    'comments',
  ]);
  if (!Array.isArray(body.content)) throw new TypeError(`${path}.content must be an array`);
  body.content.forEach((block, index) => objectAt(block, `${path}.content[${index}]`));
  const documentBody: DocumentBody = {
    content: body.content as DocumentBody['content'],
  };
  if (body.sections !== undefined) {
    if (!Array.isArray(body.sections)) throw new TypeError(`${path}.sections must be an array`);
    documentBody.sections = body.sections.map((value, index) => {
      const sectionPath = `${path}.sections[${index}]`;
      const section = objectAt(value, sectionPath);
      exactKeys(section, ['id', 'properties', 'contentStart', 'contentEnd'], sectionPath, ['id']);
      const start = integerAt(section.contentStart, `${sectionPath}.contentStart`);
      const end = integerAt(section.contentEnd, `${sectionPath}.contentEnd`);
      if (start > end || end > documentBody.content.length) {
        throw new TypeError(`${sectionPath} has an invalid content range`);
      }
      return {
        ...(section.id === undefined ? {} : { id: stringAt(section.id, `${sectionPath}.id`) }),
        properties: objectAt(
          section.properties,
          `${sectionPath}.properties`
        ) as unknown as NonNullable<DocumentBody['sections']>[number]['properties'],
        content: documentBody.content.slice(start, end),
      };
    });
  }
  if (body.finalSectionProperties !== undefined) {
    documentBody.finalSectionProperties = objectAt(
      body.finalSectionProperties,
      `${path}.finalSectionProperties`
    ) as unknown as NonNullable<DocumentBody['finalSectionProperties']>;
  }
  if (body.comments !== undefined) {
    documentBody.comments = decodeObjectArray<NonNullable<DocumentBody['comments']>[number]>(
      body.comments,
      `${path}.comments`
    );
  }
  return documentBody;
}

function decodeRelationshipEntries(value: unknown, path: string): RelationshipMap {
  return decodeMapEntries(value, path, (entry, entryPath, key) =>
    decodeRelationship(entry, entryPath, key)
  );
}

function decodeMediaEntries(value: unknown): Map<string, MediaFile> {
  const path = 'wire.document.package.mediaEntries';
  if (!Array.isArray(value)) throw new TypeError(`${path} must be an array`);
  const media = new Map<string, MediaFile>();
  const bySourcePath = new Map<string, MediaFile>();
  for (let index = 0; index < value.length; index += 1) {
    const entryPath = `${path}[${index}]`;
    const entry = value[index];
    if (!Array.isArray(entry) || entry.length !== 2) {
      throw new TypeError(`${entryPath} must be a key/value pair`);
    }
    const alias = stringAt(entry[0], `${entryPath}[0]`);
    if (media.has(alias))
      throw new TypeError(`${path} contains duplicate key ${JSON.stringify(alias)}`);
    const filePath = `${entryPath}[1]`;
    const wireFile = objectAt(entry[1], filePath);
    exactKeys(wireFile, ['path', 'filename', 'mimeType', 'base64', 'dataUrl'], filePath, [
      'filename',
    ]);
    const sourcePath = stringAt(wireFile.path, `${filePath}.path`);
    let file = bySourcePath.get(sourcePath);
    if (!file) {
      const bytes = decodeBase64(stringAt(wireFile.base64, `${filePath}.base64`));
      const created: MediaFile = {
        path: sourcePath,
        ...(wireFile.filename === undefined
          ? {}
          : { filename: stringAt(wireFile.filename, `${filePath}.filename`) }),
        mimeType: stringAt(wireFile.mimeType, `${filePath}.mimeType`),
        data: exactArrayBuffer(bytes),
        dataUrl: stringAt(wireFile.dataUrl, `${filePath}.dataUrl`),
      };
      file = created;
      bySourcePath.set(sourcePath, created);
    }
    media.set(alias, file);
  }
  return media;
}

function decodeBinaryParts(value: unknown): Map<string, ArrayBuffer> {
  const path = 'wire.embeddedFontParts';
  if (!Array.isArray(value)) throw new TypeError(`${path} must be an array`);
  const parts = new Map<string, ArrayBuffer>();
  value.forEach((entry, index) => {
    const entryPath = `${path}[${index}]`;
    const part = objectAt(entry, entryPath);
    exactKeys(part, ['path', 'base64'], entryPath);
    const partPath = stringAt(part.path, `${entryPath}.path`);
    if (parts.has(partPath)) {
      throw new TypeError(`${path} contains duplicate key ${JSON.stringify(partPath)}`);
    }
    parts.set(
      partPath,
      exactArrayBuffer(decodeBase64(stringAt(part.base64, `${entryPath}.base64`)))
    );
  });
  return parts;
}

function decodeMapEntries<T>(
  value: unknown,
  path: string,
  decodeValue: (value: unknown, path: string, key: string) => T
): Map<string, T> {
  if (!Array.isArray(value)) throw new TypeError(`${path} must be an array`);
  const map = new Map<string, T>();
  value.forEach((entry, index) => {
    const entryPath = `${path}[${index}]`;
    if (!Array.isArray(entry) || entry.length !== 2) {
      throw new TypeError(`${entryPath} must be a key/value pair`);
    }
    const key = stringAt(entry[0], `${entryPath}[0]`);
    if (map.has(key)) throw new TypeError(`${path} contains duplicate key ${JSON.stringify(key)}`);
    map.set(key, decodeValue(entry[1], `${entryPath}[1]`, key));
  });
  return map;
}

function stringArrayAt(value: unknown, path: string): string[] {
  if (!Array.isArray(value)) throw new TypeError(`${path} must be an array`);
  return value.map((entry, index) => stringAt(entry, `${path}[${index}]`));
}

function decodeObjectArray<T>(value: unknown, path: string): T[] {
  if (!Array.isArray(value)) throw new TypeError(`${path} must be an array`);
  value.forEach((entry, index) => objectAt(entry, `${path}[${index}]`));
  return value as T[];
}

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return Uint8Array.from(bytes).buffer;
}

function integerAt(value: unknown, path: string): number {
  const number = numberAt(value, path);
  if (!Number.isSafeInteger(number) || number < 0) {
    throw new TypeError(`${path} must be a non-negative safe integer`);
  }
  return number;
}

/** Focused Rust relationship adapter used by embedded-font compatibility code. */
export function parseRelationshipsXmlWithRust(
  xml: string,
  partPath = 'word/_rels/document.xml.rels'
): RelationshipMap {
  const value: unknown = JSON.parse(
    parseRelationshipsXmlWire(new TextEncoder().encode(xml), partPath)
  );
  assertSafeTree(value, 'relationshipPart');
  const part = objectAt(value, 'relationshipPart');
  exactKeys(
    part,
    ['path', 'relationships', 'canonicalBase64', 'canonicalSha256'],
    'relationshipPart'
  );
  stringAt(part.path, 'relationshipPart.path');
  return decodeRelationshipEntries(part.relationships, 'relationshipPart.relationships');
}

function decodeRelationship(value: unknown, path: string, mapId: string): Relationship {
  const relationship = objectAt(value, path);
  exactKeys(relationship, ['id', 'type', 'target', 'targetMode'], path, ['targetMode']);
  const id = stringAt(relationship.id, `${path}.id`);
  if (id !== mapId) throw new TypeError(`${path}.id does not match its map key`);
  const decoded: Relationship = {
    id,
    type: stringAt(relationship.type, `${path}.type`),
    target: stringAt(relationship.target, `${path}.target`),
  };
  if (relationship.targetMode !== undefined) {
    if (relationship.targetMode !== 'External' && relationship.targetMode !== 'Internal') {
      throw new TypeError(`${path}.targetMode is invalid`);
    }
    decoded.targetMode = relationship.targetMode;
  }
  return decoded;
}

function assertSafeTree(value: unknown, path: string, active = new WeakSet<object>()): void {
  if (value === null || typeof value !== 'object') return;
  if (active.has(value)) throw new TypeError(`${path} contains a cycle`);
  active.add(value);
  try {
    for (const key of Object.keys(value)) {
      if (FORBIDDEN_KEYS.has(key)) throw new TypeError(`${path} contains forbidden key ${key}`);
      assertSafeTree((value as Record<string, unknown>)[key], `${path}.${key}`, active);
    }
  } finally {
    active.delete(value);
  }
}

function objectAt(value: unknown, path: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${path} must be an object`);
  }
  return value as Record<string, unknown>;
}

function exactKeys(
  object: Record<string, unknown>,
  allowed: string[],
  path: string,
  optional: string[] = []
): void {
  const allowedSet = new Set(allowed);
  for (const key of Object.keys(object)) {
    if (!allowedSet.has(key)) throw new TypeError(`${path} contains unknown key ${key}`);
  }
  const optionalSet = new Set(optional);
  for (const key of allowed) {
    if (!optionalSet.has(key) && !Object.prototype.hasOwnProperty.call(object, key)) {
      throw new TypeError(`${path} is missing ${key}`);
    }
  }
}

function stringAt(value: unknown, path: string): string {
  if (typeof value !== 'string') throw new TypeError(`${path} must be a string`);
  return value;
}

function numberAt(value: unknown, path: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${path} must be a finite number`);
  }
  return value;
}

function decodeBase64(value: string): Uint8Array {
  let binary: string;
  try {
    binary = atob(value);
  } catch {
    throw new TypeError('wire canonicalBase64 is invalid');
  }
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}
