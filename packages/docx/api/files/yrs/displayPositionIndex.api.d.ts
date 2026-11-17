// @betteroffice/docx declaration yrs/displayPositionIndex.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsSession } from './index';
import { type YrsInputPositionMap } from './inputPositionMap';
import { type YrsPositionProjection } from './yrsPositionProjection';
/** @internal What resolving display positions against the live document reads. */
export type DisplayPositionReader = Pick<YrsSession, 'version' | 'storyIds' | 'paragraphs' | 'hasStory' | 'storySegments' | 'paragraphSpans' | 'locateParagraph' | 'selectionText' | 'resolveComment' | 'listRevisions'>;
/**
 * @internal Display-position projections and input maps of one document version, rebuilt when
 * the version moves on. Display positions are those of the layout of that version.
 */
export declare class DisplayPositionIndex {
    readonly reader: DisplayPositionReader;
    private version;
    private readonly projections;
    private readonly inputMaps;
    constructor(reader: DisplayPositionReader);
    private current;
    /** The projection of root story `rootStory` (`body`, `hf:<rId>`, `fn:<id>` or `en:<id>`). */
    projection(rootStory: string): YrsPositionProjection | null;
    inputMap(story: string): YrsInputPositionMap | null;
    /** The display position of `loc` in the layout of root story `rootStory`. */
    positionOf(loc: YrsLoc, rootStory: string): number | null;
}
