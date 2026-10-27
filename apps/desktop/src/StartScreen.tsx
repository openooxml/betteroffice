import type { Configuration, RecentFile } from "./bridge";
import { t, type DesktopKey } from "./i18n";

export function StartScreen({
  config,
  open,
  openRecent,
  dragging,
  busy,
}: {
  config: Configuration;
  open(): void;
  openRecent(file: RecentFile): void;
  dragging: boolean;
  busy: boolean;
}) {
  const heading: DesktopKey = {
    docx: "start.openDocument",
    xlsx: "start.openSpreadsheet",
    pptx: "start.openPresentation",
  }[config.format];
  return (
    <main className={`start-screen${dragging ? " dragging" : ""}`}>
      <div className="start-content">
        <svg
          className="document-mark"
          width="38"
          height="44"
          viewBox="0 0 38 44"
          fill="none"
          aria-hidden="true"
        >
          <path
            d="M8 2h15l11 11v27a2 2 0 0 1-2 2H8a4 4 0 0 1-4-4V6a4 4 0 0 1 4-4Z"
            stroke="currentColor"
            strokeWidth="1.5"
          />
          <path
            d="M23 2v11h11M12 23h14M12 29h14M12 35h9"
            stroke="currentColor"
            strokeWidth="1.5"
          />
        </svg>
        <h1>{dragging ? t("start.dropFileHere") : t(heading)}</h1>
        <p className="intro">{t("start.intro")}</p>
        <button className="open-primary" onClick={open} disabled={busy}>
          <span>{busy ? t("start.opening") : t("start.openFile")}</span>
          <kbd>⌘ O</kbd>
        </button>
        <p className="drop-hint">
          {t("start.dropHint")} <span>·</span> .{config.format}
        </p>
        {config.recent.length > 0 && (
          <section className="recent" aria-label={t("start.recentFiles")}>
            <h2>{t("start.recent")}</h2>
            {config.recent.map((file) => (
              <button
                key={file.id}
                className="recent-file"
                onClick={() => openRecent(file)}
                disabled={busy}
              >
                <span className="recent-name">{file.name}</span>
                <span className="recent-folder">{file.folder}</span>
              </button>
            ))}
          </section>
        )}
      </div>
    </main>
  );
}
