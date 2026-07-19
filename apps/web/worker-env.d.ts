interface KVNamespace {
  get<T = unknown>(key: string, type: "json"): Promise<T | null>;
  get(key: string): Promise<string | null>;
  put(key: string, value: string): Promise<void>;
}

interface CloudflareEnv {
  ASSETS: Fetcher;
  STATS_KV: KVNamespace;
}
