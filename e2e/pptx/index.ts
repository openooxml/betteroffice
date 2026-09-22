import { commentsThread } from './comments-thread';
import { deckBuildFromScratch } from './deck-build-from-scratch';
import { editingSession } from './editing-session';
import { formattingSweep } from './formatting-sweep';
import { layoutAllSlides } from './layout-all-slides';
import { proposalsReview } from './proposals-review';
import { pythonRoundtrip } from './python-roundtrip';
import { pythonWebLiveCollab } from './python-web-live-collab';
import { shapesAndPictures } from './shapes-and-pictures';
import { slideReorderAndDelete } from './slide-reorder-and-delete';
import { threeEditorsMesh } from './three-editors-mesh';
import { twoEditorsConverge } from './two-editors-converge';
import { typingBurst } from './typing-burst';
import type { PptxScenario } from './context';

export const scenarios: PptxScenario[] = [
  editingSession,
  typingBurst,
  deckBuildFromScratch,
  slideReorderAndDelete,
  formattingSweep,
  shapesAndPictures,
  commentsThread,
  twoEditorsConverge,
  threeEditorsMesh,
  proposalsReview,
  pythonRoundtrip,
  pythonWebLiveCollab,
  layoutAllSlides,
];
