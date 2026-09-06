export interface RawXml {
  type: 'rawXml';
  xml: string;
}

export function isRawXml(value: { type: string }): value is RawXml {
  return value.type === 'rawXml';
}
