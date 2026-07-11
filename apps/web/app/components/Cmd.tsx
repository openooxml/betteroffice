import { CopyButton } from "./CopyButton";

// single shell command with prompt glyph and copy button
export function Cmd({ command }: { command: string }) {
  return (
    <div className="flex max-w-[34rem] items-center gap-3 rounded-md border border-line bg-surface py-3 pr-3.5 pl-4.5 font-mono text-[0.8125rem]">
      <span className="text-dim select-none">$</span>
      <code className="flex-1 overflow-x-auto py-1 whitespace-nowrap [scrollbar-width:none] [&::-webkit-scrollbar]:hidden">
        {command}
      </code>
      <CopyButton text={command} />
    </div>
  );
}
