import type { RustParagraphIds } from '../docx/rustSaveFacade';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import type { DocxSessionSave } from './saveYrsDocx';
import type { EditSession } from './wasm/index';

type Plan = (
  session: YrsSession, output: Document, capture: DocxSessionSave, sourceBuffer: ArrayBuffer
) => Promise<{ plan?: NonNullable<RustParagraphIds['sectionEdits']>; lossy: boolean; strict?: boolean } | undefined>;

export const sectionSavePlanTestSeam: {
  hook?: {
    attach: (session: YrsSession, raw: EditSession) => void;
    check: (...args: [...Parameters<Plan>, useLocalProvenance: boolean]) => ReturnType<Plan>;
  };
  reference?: {
    plan: Plan;
    metadata: (document: Document) => string;
    withoutSynthetic: (document: Document, identities: DocxSessionSave['identities']) => Document;
  };
} = {};
