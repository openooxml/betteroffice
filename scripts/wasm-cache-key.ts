// Prints the wasm fingerprint for CI to key its cache on. Deriving the key from
// the same digest the build checks keeps an immutable cache entry from outliving
// an input the workflow forgot to list.
import { sourcesFingerprint } from './wasm.ts';

console.log(await sourcesFingerprint());
