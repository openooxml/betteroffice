const TEXT_PLACEHOLDER_TYPES = new Set(['title', 'ctrTitle', 'subTitle', 'body', 'obj']);

export function isTextPlaceholderType(kind: string | undefined): kind is string {
  return kind !== undefined && TEXT_PLACEHOLDER_TYPES.has(kind);
}
