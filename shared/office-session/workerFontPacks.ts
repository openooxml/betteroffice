import { setWorkerFontAssets } from '../../packages/fonts/src/workerAssets';
import { installWorkerFontPacks } from './optionalFontPacks';

installWorkerFontPacks(self, setWorkerFontAssets);
