import { createRoot } from "react-dom/client";
import { configureDefaultFonts } from "@betteroffice/docx/layout";
import { SessionApi } from "./api";
import { displayUserName, readWopiConfig } from "./config";
import { PostMessageBridge } from "./postMessage";
import WopiApp, { loadBaseline, wopiDebug, type Loaded } from "./page";
import "./styles.css";

async function boot(): Promise<void> {
  wopiDebug({ step: "boot" });
  const config = readWopiConfig(document);
  wopiDebug({ step: "config", format: config.format, postMessageOrigin: config.postMessageOrigin, canWrite: config.canWrite });
  const api = new SessionApi(config.sessionHandle);
  const actions: Loaded["actions"] = { views: [{ ViewId: 0, UserName: displayUserName(config.userFriendlyName) }] };
  const bridge = new PostMessageBridge(window, config.postMessageOrigin, {
    onSave: (options) => {
      if (!actions.save) return Promise.reject(new Error("editorNotReady"));
      return actions.save(options);
    },
    onClose: () => {
      void actions.close?.();
    },
    onBlurFocus: () => {
      if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    },
    onGrabFocus: () => {
      window.focus();
      actions.grabFocus?.();
    },
    onGetViews: () => actions.views,
  });
  wopiDebug({ bridgeEnabled: bridge.enabled, bridgeOrigin: config.postMessageOrigin });
  bridge.frameReady();
  wopiDebug({ frameReadySent: true });

  let root: ReturnType<typeof createRoot> | null = null;
  try {
    configureDefaultFonts({ load: () => import("@betteroffice/fonts"), baseUrl: "/fonts/" });
    let rootElement = document.getElementById("root");
    if (!rootElement) {
      rootElement = document.createElement("div");
      rootElement.id = "root";
      document.body.appendChild(rootElement);
    }
    root = createRoot(rootElement);
    const baseline = await loadBaseline(api);
    actions.views = [{ ViewId: baseline.clientId, UserName: displayUserName(config.userFriendlyName) }];
    const loaded: Loaded = { config, api, bridge, ...baseline, actions };
    root.render(<WopiApp loaded={loaded} />);
    wopiDebug({ step: "rendered" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Could not open the document.";
    wopiDebug({ step: "bootError", bootError: message });
    bridge.bootFailed();
    if (root) root.render(<div className="wopi-error" role="alert">{message}</div>);
    else document.body.textContent = message;
  }
}

void boot();
