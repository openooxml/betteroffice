/** The `w:numId` and `w:abstractNumId` a save gives a list definition the editor created. */
export interface DocxListDefinitionIds {
  numId: number;
  abstractNumId: number;
}

/**
 * The ids a save gives the list definitions the editor created, by key, whatever order the keys
 * come in: a definition an earlier save numbered, in `previous`, keeps its ids, and every other
 * one takes the next ids free above the source's and those already given.
 */
export function savedListDefinitionIds(
  _source: { numIds: readonly number[]; abstractNumIds: readonly number[] },
  _keys: readonly string[],
  _previous: ReadonlyMap<string, DocxListDefinitionIds> = new Map()
): Map<string, DocxListDefinitionIds> {
  return new Map();
}
