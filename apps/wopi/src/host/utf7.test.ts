import { describe, expect, test } from "bun:test";
import { decodeUtf7, encodeUtf7 } from "./utf7";

describe("utf7", () => {
  test("round trips representative strings", () => {
    const cases = [
      "report.docx",
      "plain ASCII file 123.txt",
      "München Über file.docx",
      "文件名称.docx",
      "emoji 😀 file.docx",
      "a+b",
      "a-b",
      "a+-b -- ++",
      "plus + and minus - mixed Ünïcödé 文件 😀.xlsx",
      "",
      "+",
      "-",
      "Jürgen 文件 +-.pptx",
    ];
    for (const value of cases) {
      expect(decodeUtf7(encodeUtf7(value))).toBe(value);
    }
  });

  test("keeps plain ASCII readable and escapes a literal plus", () => {
    expect(encodeUtf7("report.docx")).toBe("report.docx");
    expect(encodeUtf7("a+b")).toBe("a+-b");
    expect(encodeUtf7("a-b")).toBe("a-b");
  });

  test("decodes shifted sequences with and without a terminator", () => {
    expect(decodeUtf7("+AGU-")).toBe("e");
    expect(decodeUtf7(encodeUtf7("ü"))).toBe("ü");
    expect(decodeUtf7("a+-b")).toBe("a+b");
  });
});
