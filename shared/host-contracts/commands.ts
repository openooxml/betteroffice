/**
 * A value that survives `JSON.stringify` / `JSON.parse` unchanged.
 * @experimental
 */
export type JsonValue =
  | null
  | boolean
  | number
  | string
  | readonly JsonValue[]
  | { readonly [key: string]: JsonValue };

/**
 * A machine-readable code with a localized explanation for people.
 * @experimental
 */
export interface CommandReason<Code extends string = string> {
  code: Code;
  message: string;
}

/**
 * Serializable availability of one editor command. A disabled command always
 * says why; branch on `disabledReason.code`, show `disabledReason.message`.
 * @experimental
 */
export type CommandState<Value extends JsonValue = JsonValue, Code extends string = string> = {
  /** Toggle state; `'mixed'` when the selection disagrees. */
  active?: boolean | 'mixed';
  /** The command's current value, such as the selected style or font size. */
  value?: Value;
} & (
  | { enabled: true; disabledReason?: never }
  | { enabled: false; disabledReason: CommandReason<Code> }
);
