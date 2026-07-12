import { FormatCard } from "./components/FormatCard";
import { formats } from "../lib/formats";

export default function Home() {
  return (
    <div>
      <h1 className="max-w-2xl text-3xl font-medium tracking-tight text-fg">
        BetterOffice demos
      </h1>
      <p className="mt-3 max-w-2xl text-sm leading-relaxed text-ink">
        Try the engines in your browser — Word documents, spreadsheets, and
        slides, rendered by native OOXML engines written in Rust and compiled to
        WebAssembly. Pick a format to start.
      </p>

      <div className="mt-10 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
        {formats.map((f) => (
          <FormatCard key={f.id} format={f} />
        ))}
      </div>
    </div>
  );
}
