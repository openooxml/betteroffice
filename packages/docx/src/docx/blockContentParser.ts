/** Pure list-marker compatibility helper retained outside the Rust story parser. */

const SYMBOL_BULLET_MAP: Record<number, string> = {
  0x00b7: '•',
  0x006f: '○',
  0x00a7: '■',
  0x00fc: '✓',
  0x006e: '■',
  0x0071: '○',
  0x0075: '◆',
  0x0076: '❖',
  0x00a8: '✓',
  0x00fb: '✓',
  0x00fe: '✓',
  0xf0b7: '•',
  0xf06e: '■',
  0xf06f: '○',
  0xf0a7: '■',
  0xf0fc: '✓',
  0x2022: '•',
  0x25cf: '●',
  0x25cb: '○',
  0x25a0: '■',
  0x25a1: '□',
  0x25c6: '◆',
  0x25c7: '◇',
  0x2013: '–',
  0x2014: '—',
  0x003e: '>',
  0x002d: '-',
};

export function convertBulletToUnicode(bulletChar: string): string {
  if (!bulletChar || bulletChar.trim() === '') return '•';
  const charCode = bulletChar.charCodeAt(0);
  if (SYMBOL_BULLET_MAP[charCode]) return SYMBOL_BULLET_MAP[charCode];
  if (charCode >= 0xe000 && charCode <= 0xf8ff) return '•';
  if (charCode < 32 || (charCode >= 127 && charCode < 160)) return '•';
  return bulletChar;
}
