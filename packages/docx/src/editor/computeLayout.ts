import type { ResidentMeasurementConfig } from '../layout/measure';
import type {
  BlockExtent,
  Layout,
  LayoutBlock,
  LayoutOptions,
  MeasuredBlock,
} from '../layout/pagination';
import type { DisplayListHeadersFooters } from '../layout/render/rustDisplayList';
import type { Document, NoteKind, SectionProperties } from '../types/document';
import type { YrsRenderEnv, YrsSession } from '../yrs';

interface ResidentRegionLayoutOutput {
  measured: MeasuredBlock[];
  options: LayoutOptions;
  layout: Layout;
  headersFooters?: DisplayListHeadersFooters;
  notesConverged: boolean;
}

export interface ResidentRegionLayoutRequest {
  bodyStory: 'body';
  options: Pick<LayoutOptions, 'contractVersion' | 'pageGap'>;
  regions: {
    sections: Array<{
      sectionId?: string;
      properties: SectionProperties;
    }>;
    settings?: Document['package']['settings'];
    watermark?: SectionProperties['watermark'];
  };
  notes: {
    contents: Array<{
      id: number;
      noteKind: NoteKind;
      height: 0;
    }>;
  };
  renderEnv: YrsRenderEnv;
  measurement?: ResidentMeasurementConfig;
}

export interface ComputeLayoutInputs {
  document: Document | null;
  pageGap: number;
  session: Pick<YrsSession, 'layoutDocumentWithRegionsJson'>;
  renderEnv: YrsRenderEnv;
  measurement: ResidentMeasurementConfig;
}

export interface LayoutComputation {
  blocks: LayoutBlock[];
  measures: BlockExtent[];
  layout: Layout;
  notesConverged: boolean;
}

const kernelInputsByLayout = new WeakMap<
  Layout,
  {
    measured: MeasuredBlock[];
    options: LayoutOptions;
    headersFooters?: DisplayListHeadersFooters;
  }
>();

function orderedSections(document: Document | null): ResidentRegionLayoutRequest['regions']['sections'] {
  const body = document?.package.document;
  if (!body) return [{ properties: {} }];
  const sections = (body.sections ?? []).map((section) => ({
    sectionId: section.id ?? section.properties.sectionId,
    properties: section.properties,
  }));
  sections.push({
    sectionId: body.finalSectionProperties?.sectionId,
    properties: body.finalSectionProperties ?? {},
  });
  return sections;
}

export function buildResidentRegionLayoutRequest(
  document: Document | null,
  pageGap: number,
  renderEnv: YrsRenderEnv
): ResidentRegionLayoutRequest {
  const contents: ResidentRegionLayoutRequest['notes']['contents'] = [];
  for (const note of document?.package.footnotes ?? []) {
    if (note.noteType && note.noteType !== 'normal') continue;
    contents.push({ id: note.id, noteKind: 'footnote', height: 0 });
  }
  for (const note of document?.package.endnotes ?? []) {
    if (note.noteType && note.noteType !== 'normal') continue;
    contents.push({ id: note.id, noteKind: 'endnote', height: 0 });
  }
  return {
    bodyStory: 'body',
    options: {
      contractVersion: document?.package.contractVersion,
      pageGap,
    },
    regions: {
      sections: orderedSections(document),
      settings: document?.package.settings,
      watermark: document?.package.document.finalSectionProperties?.watermark,
    },
    notes: { contents },
    renderEnv,
  };
}

export function getLayoutKernelInputs(layout: Layout):
  | {
      measured: MeasuredBlock[];
      options: unknown;
      headersFooters?: DisplayListHeadersFooters;
    }
  | undefined {
  return kernelInputsByLayout.get(layout);
}

export function computeLayout(inputs: ComputeLayoutInputs): LayoutComputation {
  const request = buildResidentRegionLayoutRequest(
    inputs.document,
    inputs.pageGap,
    inputs.renderEnv
  );
  request.measurement = inputs.measurement;
  const output = JSON.parse(
    inputs.session.layoutDocumentWithRegionsJson(JSON.stringify(request))
  ) as ResidentRegionLayoutOutput;
  kernelInputsByLayout.set(output.layout, {
    measured: output.measured,
    options: output.options,
    ...(output.headersFooters ? { headersFooters: output.headersFooters } : {}),
  });
  return {
    blocks: output.measured.map((item) => item.block),
    measures: output.measured.map((item) => item.measure),
    layout: output.layout,
    notesConverged: output.notesConverged,
  };
}
