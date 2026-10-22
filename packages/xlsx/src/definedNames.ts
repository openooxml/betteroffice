export type XlsxDefinedNameOp = {
  type: 'setDefinedName';
  name: string;
  refers_to: string;
  scope?: number | null;
  hidden?: boolean;
} | {
  type: 'deleteDefinedName';
  name: string;
  scope?: number | null;
};

export interface XlsxDefinedName {
  name: string;
  refersTo: string;
  scope: number | null;
  hidden: boolean;
}

export interface XlsxDefinedNamesRequest {
  /** Omitted selects all scopes; null selects workbook scope. */
  scope?: number | null;
  offset?: number;
  limit?: number;
}

export interface XlsxDefinedNamesResult {
  version: string;
  items: XlsxDefinedName[];
  total: number;
  nextOffset: number | null;
}
