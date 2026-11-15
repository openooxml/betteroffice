/** 0-based sheet column to its A1 letters: 0 -> A, 25 -> Z, 26 -> AA. */
export function columnLetter(col: number): string {
  let out = '';
  for (let n = col; n >= 0; n = Math.floor(n / 26) - 1) out = String.fromCharCode(65 + (n % 26)) + out;
  return out;
}
