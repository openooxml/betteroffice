// the betteroffice mark, fill follows currentColor
export function Logo({
  height = 20,
  className,
}: {
  height?: number;
  className?: string;
}) {
  return (
    <svg
      height={height}
      width={height}
      viewBox="0 0 64 64"
      fill="currentColor"
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      aria-hidden="true"
    >
      <path d="M5 7 H28 L31 25 H8 Z" />
      <path d="M9 33 H32 L36 56 H13 Z" />
      <path d="M36 7 H51 L59 56 H44 Z" />
    </svg>
  );
}
