import { useMemo, useRef } from 'react';
import {
  validReservationsPayload,
  type DocxReservationsPayload,
} from '@betteroffice/docx/yrs';
import type { DocxPluginActivation, DocxPluginHost } from './createDocxPluginHost';
import type { DocxReservationSet } from './types';

interface Evaluated {
  definition: DocxPluginActivation['plugin'];
  generate: DocxPluginActivation['plugin']['reservations'];
  state: unknown;
  version: string;
  mode: string;
  readOnly: boolean;
  set: DocxReservationSet | null;
}

/** The set a plugin returned, refused as `invalid-payload` unless well formed. */
function checkedSet(value: unknown): DocxReservationSet | null {
  if (value === null || value === undefined) return null;
  const set = value as DocxReservationSet;
  if (
    typeof set !== 'object' ||
    !validReservationsPayload({ version: set.version, sets: [{ owner: 'plugin', items: set.items }] })
  ) {
    throw new TypeError('invalid-payload');
  }
  return set;
}

/** Combines current sets with each active owner's last current set while its replacement is stale. */
export function usePluginReservations(
  host: DocxPluginHost,
  activations: readonly DocxPluginActivation[]
): DocxReservationsPayload | null {
  const evaluated = useRef(new WeakMap<object, Evaluated>());
  const lastSets = useRef(new Map<string, DocxReservationSet>());
  const generation = useRef<string | null>(null);
  const previous = useRef<{ json: string; payload: DocxReservationsPayload | null }>({
    json: 'null',
    payload: null,
  });
  return useMemo(() => {
    if (generation.current !== host.generation()) {
      lastSets.current.clear();
      generation.current = host.generation();
    }
    const version = host.version();
    const sets: DocxReservationsPayload['sets'][number][] = [];
    const seen = new Set<string>();
    let defined = false;
    for (const activation of activations) {
      const generate = activation.plugin.reservations;
      if (!generate) continue;
      defined = true;
      seen.add(activation.pluginId);
      const { snapshot } = activation.context;
      const state = activation.context.state;
      let entry = evaluated.current.get(activation.activation);
      if (
        !entry ||
        entry.definition !== activation.plugin ||
        entry.generate !== generate ||
        entry.state !== state ||
        entry.version !== snapshot.version ||
        entry.mode !== snapshot.mode ||
        entry.readOnly !== snapshot.readOnly
      ) {
        entry = {
          definition: activation.plugin,
          generate,
          state,
          version: snapshot.version,
          mode: snapshot.mode,
          readOnly: snapshot.readOnly,
          set: host.guard(
            activation.pluginId,
            'reservations',
            (context) => checkedSet(generate.call(activation.plugin, context)),
            null
          ),
        };
        evaluated.current.set(activation.activation, entry);
      }
      if (!entry.set) {
        lastSets.current.delete(activation.pluginId);
        continue;
      }
      if (entry.set.version === version) {
        lastSets.current.set(activation.pluginId, entry.set);
        sets.push({ owner: activation.pluginId, items: entry.set.items });
      } else {
        const prior = lastSets.current.get(activation.pluginId);
        if (prior) sets.push({ owner: activation.pluginId, version: prior.version, items: prior.items });
      }
    }
    for (const owner of lastSets.current.keys()) {
      if (!seen.has(owner)) lastSets.current.delete(owner);
    }
    const payload = defined ? { version, sets } : null;
    const json = JSON.stringify(payload);
    if (json !== previous.current.json) previous.current = { json, payload };
    return previous.current.payload;
  }, [host, activations]);
}
