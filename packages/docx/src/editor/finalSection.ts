import type { DocumentBody, SectionProperties } from '../types/document';

/**
 * The last section as the engine lays it out, with references and `titlePg`
 * inherited forward; `finalSectionProperties` is the body sectPr as authored.
 */
export function resolvedFinalSectionProperties(
  body: DocumentBody | null | undefined
): SectionProperties | undefined {
  return body?.sections?.at(-1)?.properties ?? body?.finalSectionProperties;
}

/**
 * Applies an edit to the last section on both views: the resolved entry the
 * engine lays out and the authored sectPr the serializer saves.
 */
export function updateFinalSectionProperties(
  body: DocumentBody,
  update: (properties: SectionProperties) => SectionProperties
): DocumentBody {
  const sections = body.sections?.map((section, index, all) =>
    index === all.length - 1 ? { ...section, properties: update(section.properties) } : section
  );
  return {
    ...body,
    ...(sections ? { sections } : {}),
    finalSectionProperties: update(body.finalSectionProperties ?? {}),
  };
}
