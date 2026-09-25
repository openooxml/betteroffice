import * as os from 'node:os';
import * as path from 'node:path';

import {
  fetchAsset,
  resolveAssetCacheDir,
} from '../scripts/office-quality/asset-cache.mjs';
import { download } from '../scripts/office-quality/download.mjs';
import { CORPUS_ORIGIN } from '../scripts/office-quality/samples.mjs';

export type Format = 'docx' | 'xlsx' | 'pptx';

export interface PinnedSample {
  id: string;
  format: Format;
  bytes: number;
  sha256: string;
  /** Source part file name when the corpus does not use `source.<format>`. */
  file?: string;
}

export const SAMPLES: readonly PinnedSample[] = [
  {
    id: 'betteroffice-demo',
    format: 'docx',
    bytes: 6383,
    sha256: '5b272a248fe899d3b90259297dc20158b1bdfc7c0e1d653c7e29985869e17c51',
  },
  {
    id: 'oxi-en-administrative-02',
    format: 'docx',
    bytes: 33671,
    sha256: '524b7d77b2ccd46921176081636f728a5be66560662e820d46e531700cabdb75',
  },
  {
    id: 'bo-corpus-1',
    format: 'docx',
    bytes: 107223,
    sha256: '2288816e29b2cf66ef228f4772609977fc9146fd9640e8940aca4ff7e6318c32',
    file: 'bo-corpus-1.docx',
  },
  {
    id: 'betteroffice-workbook',
    format: 'xlsx',
    bytes: 20744,
    sha256: 'd6d28b00ca4352cf38124c81cb686c8848f25af89b339e8dbee9ffb67120fa60',
  },
  {
    id: 'sheetpedia-7347e836f3f2',
    format: 'xlsx',
    bytes: 10176,
    sha256: '7347e836f3f2a0c340c12b1ae453831c5b04ef54b8e70a8a48ecc1f90d0ed08b',
  },
  {
    id: 'sheetpedia-06dfb94ff719',
    format: 'xlsx',
    bytes: 16132,
    sha256: '06dfb94ff719235496b78ec9d9fd9e4c305138500597d4f769bd8385a51f542d',
  },
  {
    id: 'betteroffice-slides',
    format: 'pptx',
    bytes: 12395,
    sha256: 'bfdcad9f47c9b614ef6579351ec632c5137f013c1e4aca9b162326bef5bbbca7',
  },
  {
    id: 'pptarena-002-original',
    format: 'pptx',
    bytes: 334145,
    sha256: 'd7bafc1afc1ab5486e3d8019d087f690c19b5594c1ece7827a129d14e572d5db',
  },
  {
    id: 'pptarena-003-original',
    format: 'pptx',
    bytes: 471140,
    sha256: '01238cae60124a40312629aa2f8ec5c6735ef4e38bc56f0f85691336cf1fe400',
  },
];

/** `QUALITY_ASSET_CACHE` wins; otherwise the pinned bytes live under the user cache. */
const DEFAULT_CACHE_DIR = path.join(
  os.homedir(),
  '.cache',
  'betteroffice',
  'e2e-assets'
);

export function samplesFor(format: Format): PinnedSample[] {
  return SAMPLES.filter((sample) => sample.format === format);
}

/** The pinned bytes, from the local asset cache when it holds them. */
export async function loadSample(sample: PinnedSample): Promise<Uint8Array> {
  const entry = {
    url: `${CORPUS_ORIGIN}/${sample.id}/${
      sample.file ?? `source.${sample.format}`
    }`,
    bytes: sample.bytes,
    sha256: sample.sha256,
  };
  const bytes: Buffer = await fetchAsset(entry, sample.id, download, {
    cacheDir: resolveAssetCacheDir(process.env) ?? DEFAULT_CACHE_DIR,
    origin: CORPUS_ORIGIN,
    maximum: 128 * 1024 * 1024,
  });
  return new Uint8Array(bytes);
}
