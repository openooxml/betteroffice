import { fields } from './projectionParity.allowlist.fields';
import { markers } from './projectionParity.allowlist.markers';
import { storiesAux } from './projectionParity.allowlist.storiesAux';
import { tables } from './projectionParity.allowlist.tables';

export interface ProjectionAllowance {
  kind: string;
  path: string;
  reason: string;
}

const lanes = { fields, tables, markers, storiesAux };
export type ProjectionFamily = keyof typeof lanes;

export const projectionAllowlist = (Object.keys(lanes) as ProjectionFamily[]).flatMap((family) =>
  lanes[family].map((entry) => ({ family, ...entry }))
);
