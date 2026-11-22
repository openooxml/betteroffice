import jp from '../../fonts-cjk/assets/NotoSansJP-Regular/slices.json';
import kr from '../../fonts-cjk/assets/NotoSansKR-Regular/slices.json';
import sc from '../../fonts-cjk/assets/NotoSansSC-Regular/slices.json';
import tc from '../../fonts-cjk/assets/NotoSansTC-Regular/slices.json';
import serif from '../../fonts-cjk/assets/NotoSerifSC-Regular/slices.json';
import { addonManifests } from './packed';

export const cjkSliceFiles = [jp, kr, sc, tc, serif].flatMap(({ skeleton, core, chunks }) =>
  [skeleton.file, core.file, ...chunks.map(({ file }) => file)]);
cjkSliceFiles.push(...addonManifests.cjk.faces.flatMap(({ slices }) => slices ? [slices.index] : []));
