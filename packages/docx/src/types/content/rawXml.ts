export interface RawXml {
  type: 'rawXml';
  xml: string;
}

export function isRawXml(value: { type: string }): boolean {
  return value.type === 'rawXml';
}
