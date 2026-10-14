import openNext from "./.open-next/worker.js";
import { refreshDownloadStats } from "./lib/download-stats";

export * from "./.open-next/worker.js";

export default {
  fetch: openNext.fetch,
  async scheduled(_controller, env, ctx) {
    ctx.waitUntil(
      refreshDownloadStats(env.STATS_KV).then((result) => console.log("download stats", JSON.stringify(result))),
    );
  },
};
