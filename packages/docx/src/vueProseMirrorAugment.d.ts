import '@betteroffice/docx/prosemirror';
import '@betteroffice/docx/prosemirror/imageCommit';

declare module '@betteroffice/docx/prosemirror/imageCommit' {
  export function toolbarValueToLayoutTarget(
    value: string
  ):
    | 'inline'
    | 'square'
    | 'tight'
    | 'through'
    | 'topAndBottom'
    | 'behind'
    | 'inFront'
    | 'squareLeft'
    | 'squareRight'
    | undefined;
}

declare module '@betteroffice/docx/prosemirror' {
  export function enclosingSdtGroupIds(doc: unknown, from: number, to: number): Set<string>;
}
