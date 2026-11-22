import { setFontAssetSource } from '../../packages/fonts/src/worker';
import { installWorkerFontPacks } from './optionalFontPacks';

installWorkerFontPacks(self, setFontAssetSource);
