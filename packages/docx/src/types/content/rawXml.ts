export interface RawXml {
  type: 'rawXml';
  xml: string;
  /** The source occurrence of its first `w:p`, on documents parsed for an editing session. */
  sourceOrdinal?: number;
}

export function isRawXml(value: { type: string }): value is RawXml {
  return value.type === 'rawXml';
}
