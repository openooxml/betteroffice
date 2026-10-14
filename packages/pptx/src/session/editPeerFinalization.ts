// FinalizationRegistry is ES2021, newer than this package's lib, so it is typed locally.
interface Registry<T> {
  register(target: object, heldValue: T, unregisterToken?: object): void;
  unregister(unregisterToken: object): boolean;
}

const scope = globalThis as unknown as {
  FinalizationRegistry: new <T>(cleanup: (heldValue: T) => void) => Registry<T>;
};
const FinalizationRegistry = scope.FinalizationRegistry;

export const finalizationRegistrations: { registry: string; live: number }[] = [];

class CountingFinalizationRegistry<T> extends FinalizationRegistry<T> {
  private readonly tokens = new WeakMap<object, number>();
  private readonly count: { registry: string; live: number };

  constructor(cleanup: (value: T) => void) {
    super(cleanup);
    const registry = cleanup.toString().match(/__wbg_(\w+)_free/)?.[1] ?? cleanup.toString();
    this.count = { registry, live: 0 };
    finalizationRegistrations.push(this.count);
  }

  override register(...args: Parameters<Registry<T>['register']>): void {
    super.register(...args);
    this.count.live += 1;
    const token = args[2];
    if (token !== undefined) this.tokens.set(token, (this.tokens.get(token) ?? 0) + 1);
  }

  override unregister(...args: Parameters<Registry<T>['unregister']>): boolean {
    const removed = super.unregister(...args);
    if (removed) {
      this.count.live -= this.tokens.get(args[0]) ?? 0;
      this.tokens.delete(args[0]);
    }
    return removed;
  }
}

scope.FinalizationRegistry = CountingFinalizationRegistry;

export function restoreFinalizationRegistry(): void {
  scope.FinalizationRegistry = FinalizationRegistry;
}
