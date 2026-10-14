import { commentsAndAnchors } from './comments-and-anchors';
import { editingSession } from './editing-session';
import { formattingAndStyles } from './formatting-and-styles';
import { paginationPressure } from './pagination-pressure';
import { pythonLayoutParity } from './python-layout-parity';
import { pythonRoundtrip } from './python-roundtrip';
import { saveLoadCycles } from './save-load-cycles';
import { searchAndReplaceSweep } from './search-and-replace-sweep';
import { tablesDeep } from './tables-deep';
import { threeEditorsSuggesting } from './three-editors-suggesting';
import { twoEditorsConverge } from './two-editors-converge';
import { typingBurst } from './typing-burst';
import type { DocxScenario } from './context';

export const scenarios: DocxScenario[] = [
  editingSession,
  typingBurst,
  paginationPressure,
  tablesDeep,
  formattingAndStyles,
  searchAndReplaceSweep,
  twoEditorsConverge,
  threeEditorsSuggesting,
  commentsAndAnchors,
  pythonRoundtrip,
  pythonLayoutParity,
  saveLoadCycles,
];
