const DIRECT = new Set(
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'(),-./:? ".split(""),
);

/** Encodes text as RFC 2152 UTF-7 for WOPI name headers. */
export function encodeUtf7(value: string): string {
  let out = "";
  let run = "";
  const flush = () => {
    if (run.length === 0) return;
    const bytes: number[] = [];
    for (let i = 0; i < run.length; i += 1) {
      const unit = run.charCodeAt(i);
      bytes.push((unit >> 8) & 0xff, unit & 0xff);
    }
    out += `+${Buffer.from(bytes).toString("base64").replace(/=+$/, "")}-`;
    run = "";
  };
  for (const char of value) {
    if (char === "+") {
      flush();
      out += "+-";
    } else if (char.length === 1 && char.charCodeAt(0) < 128 && DIRECT.has(char)) {
      flush();
      out += char;
    } else {
      run += char;
    }
  }
  flush();
  return out;
}

/** Decodes RFC 2152 UTF-7 text from a WOPI header. */
export function decodeUtf7(value: string): string {
  let out = "";
  let i = 0;
  while (i < value.length) {
    const char = value[i] as string;
    if (char !== "+") {
      out += char;
      i += 1;
      continue;
    }
    if (value[i + 1] === "-") {
      out += "+";
      i += 2;
      continue;
    }
    let end = i + 1;
    while (end < value.length && /[A-Za-z0-9+/]/.test(value[end] as string)) end += 1;
    if (end === i + 1) {
      out += "+";
      i += 1;
      continue;
    }
    let encoded = value.slice(i + 1, end);
    while (encoded.length % 4 !== 0) encoded += "=";
    const bytes = Buffer.from(encoded, "base64");
    for (let b = 0; b + 1 < bytes.length; b += 2) {
      out += String.fromCharCode(((bytes[b] as number) << 8) | (bytes[b + 1] as number));
    }
    i = end;
    if (value[i] === "-") i += 1;
  }
  return out;
}

export const encodeUTF7 = encodeUtf7;
export const decodeUTF7 = decodeUtf7;
