/** @internal */
export function localNowSerial(ms: number): number {
  return (ms - new Date(ms).getTimezoneOffset() * 60_000) / 86_400_000 + 25_569;
}
