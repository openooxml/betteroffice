/** @internal */
export { createPresentationCommandClient } from './session/commandClient';
/** @internal */
export { createPptxWorkerSession, PptxWorkerUnsupportedError } from './session/commandSession';
/** @internal */
export type { PptxWorkerSession, PptxWorkerSessionOptions } from './session/commandSession';
/** @internal */
export { presentationSharedInputs } from './session/commandSharedInput';
/** @internal */
export { CommandOutcomeUnknown, PRESENTATION_COMMAND_ADMISSION, PRESENTATION_COMMAND_METHODS, PRESENTATION_COMMAND_MUTATIONS } from './session/commandProtocol';
/** @internal */
export { frameAssetIds } from './session/frame';
/** @internal */
export { decodeCommandDisplayList, indexGeometry, indexHitGeometry, findLine, transformedRect,
  selectionGeometry, hitGeometry, navigateGeometry } from './session/commandGeometry';
/** @internal */
export type { CommandRect, CommandGeometryLine, CommandGeometryIndex, CommandHitGeometryIndex } from './session/commandGeometry';
/** @internal */
export type {
  PresentationCommandMutationMethods, PresentationCommandMutation,
  PresentationCommandAccessMethods, PresentationCommandAccessMethod, PresentationCommandAccessResult, PresentationCommandAck, PresentationCommandBasis, PresentationCommandClient,
  PresentationCommandEditState, PresentationCommandEnvelope, PresentationCommandEvents,
  PresentationCommandFailure, PresentationCommandFence, PresentationCommandFrame, PresentationCommandFrameTag,
  PresentationCommandMotion, PresentationCommandIdentity,
  PresentationCommandInput, PresentationCommandIntent, PresentationCommandKind,
  PresentationCommandMethods, PresentationCommandOpenOptions, PresentationCommandOpenResult,
  PresentationCommandOutcome, PresentationCommandPayloads, PresentationCommandReadResult,
  PresentationCommandRefusalReason, PresentationCommandSaveResult, PresentationCommandSelection,
  PresentationCommandSelected, PresentationCommandObjectSelection, PresentationCommandObjectTarget,
  PresentationCommandSelectionReference, PresentationCommandTextInput,
  PresentationCommandThumbnailRequest, PresentationCommandThumbnailResult, PresentationCommandToolbarContext,
} from './session/commandProtocol';

/** @internal */
export { presentationCollaborationPipes } from './session/commandCollaboration';
