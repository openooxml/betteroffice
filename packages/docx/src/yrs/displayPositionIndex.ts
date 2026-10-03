import type { YrsLoc, YrsSession } from './index';
import { createYrsInputPositionMap, type YrsInputPositionMap } from './inputPositionMap';
import {
  createYrsPositionProjection,
  yrsLocToProjectedDisplayPosition,
  type YrsPositionProjection,
} from './yrsPositionProjection';

/** @internal What resolving display positions against the live document reads. */
export type DisplayPositionReader = Pick<
  YrsSession,
  | 'version'
  | 'storyIds'
  | 'paragraphs'
  | 'hasStory'
  | 'storySegments'
  | 'paragraphSpans'
  | 'locateParagraph'
  | 'selectionText'
  | 'resolveComment'
  | 'listRevisions'
>;

/**
 * @internal Display-position projections and input maps of one document version, rebuilt when
 * the version moves on. Display positions are those of the layout of that version.
 */
export class DisplayPositionIndex {
  private version: string | null = null;
  private readonly projections = new Map<string, YrsPositionProjection | null>();
  private readonly inputMaps = new Map<string, YrsInputPositionMap | null>();

  constructor(readonly reader: DisplayPositionReader) {}

  private current(): void {
    const version = this.reader.version();
    if (version === this.version) return;
    this.version = version;
    this.projections.clear();
    this.inputMaps.clear();
  }

  /** The projection of root story `rootStory` (`body`, `hf:<rId>`, `fn:<id>` or `en:<id>`). */
  projection(rootStory: string): YrsPositionProjection | null {
    this.current();
    let projection = this.projections.get(rootStory);
    if (projection === undefined) {
      projection = createYrsPositionProjection(this.reader, rootStory);
      this.projections.set(rootStory, projection);
    }
    return projection;
  }

  inputMap(story: string): YrsInputPositionMap | null {
    this.current();
    let map = this.inputMaps.get(story);
    if (map === undefined) {
      map = this.reader.hasStory(story)
        ? createYrsInputPositionMap(story, this.reader.paragraphSpans(story))
        : null;
      this.inputMaps.set(story, map);
    }
    return map;
  }

  /** The display position of `loc` in the layout of root story `rootStory`. */
  positionOf(loc: YrsLoc, rootStory: string): number | null {
    return yrsLocToProjectedDisplayPosition(
      this.reader,
      (root) => this.projection(root),
      loc,
      rootStory,
      (story) => this.inputMap(story)
    );
  }
}
