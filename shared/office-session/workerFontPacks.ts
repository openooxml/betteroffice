import { setOptionalFontPacks } from '@betteroffice/fonts/worker';
import { installWorkerFontPacks } from './optionalFontPacks';

installWorkerFontPacks(self, setOptionalFontPacks);
