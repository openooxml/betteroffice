/**
 * Math equations (`m:oMath`, `m:oMathPara`). OMML XML is round-tripped
 * verbatim to preserve fidelity Word/Pages/Docs can disagree on.
 */

/**
 * Math equation content (m:oMath or m:oMathPara)
 */
export interface MathEquation {
  type: 'mathEquation';
  /** Whether this is a block (oMathPara) or inline (oMath) equation */
  display: 'inline' | 'block';
  /** Raw OMML XML for round-trip preservation */
  ommlXml: string;
  /** Plain text representation for accessibility/fallback */
  plainText?: string;
}
