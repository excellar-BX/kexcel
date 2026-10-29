import type { SVGProps } from "react";

type IconProps = SVGProps<SVGSVGElement> & { size?: number };

function Svg({ size = 18, children, ...rest }: IconProps & { children: React.ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={2}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      {...rest}
    >
      {children}
    </svg>
  );
}

const icon = (paths: string[]) =>
  function Icon(props: IconProps) {
    return (
      <Svg {...props}>
        {paths.map((d) => (
          <path key={d} d={d} />
        ))}
      </Svg>
    );
  };

export const UndoIcon = icon(["M9 14 4 9l5-5", "M4 9h10.5a5.5 5.5 0 0 1 0 11H11"]);
export const RedoIcon = icon(["m15 14 5-5-5-5", "M20 9H9.5a5.5 5.5 0 0 0 0 11H13"]);
export const ChevronDown = icon(["m6 9 6 6 6-6"]);
export const ChevronRight = icon(["m9 18 6-6-6-6"]);
export const PlusIcon = icon(["M5 12h14", "M12 5v14"]);
export const MinusIcon = icon(["M5 12h14"]);
export const CheckIcon = icon(["M20 6 9 17l-5-5"]);
export const AlignLeftIcon = icon(["M3 6h18", "M3 12h12", "M3 18h15"]);
export const AlignCenterIcon = icon(["M3 6h18", "M6 12h12", "M4.5 18h15"]);
export const AlignRightIcon = icon(["M3 6h18", "M9 12h12", "M6 18h15"]);
export const SortIcon = icon(["m3 16 4 4 4-4", "M7 20V4", "M11 4h10", "M11 8h7", "M11 12h4"]);
export const EraserIcon = icon([
  "m7 21-4.3-4.3a2.4 2.4 0 0 1 0-3.4l9.6-9.6a2.4 2.4 0 0 1 3.4 0l5.6 5.6a2.4 2.4 0 0 1 0 3.4L13 21",
  "M22 21H7",
  "m5 11 9 9",
]);
export const FillIcon = icon([
  "m19 11-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2a2 2 0 0 0 2.8 0L19 11Z",
  "m5 2 5 5",
  "M2 13h15",
  "M22 20a2 2 0 1 1-4 0c0-1.6 1.7-2.4 2-4 .3 1.6 2 2.4 2 4Z",
]);
export const PrintIcon = icon([
  "M6 9V3h12v6",
  "M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2",
  "M6 14h12v7H6z",
]);
export const SparklesIcon = icon([
  "M9.94 15.5A2 2 0 0 0 8.5 14.06l-6.13-1.58a.5.5 0 0 1 0-.96L8.5 9.94A2 2 0 0 0 9.94 8.5l1.58-6.14a.5.5 0 0 1 .96 0l1.58 6.14a2 2 0 0 0 1.44 1.44l6.13 1.58a.5.5 0 0 1 0 .96l-6.13 1.58a2 2 0 0 0-1.44 1.44l-1.58 6.14a.5.5 0 0 1-.96 0Z",
  "M20 3v4",
  "M22 5h-4",
]);
export const XIcon = icon(["M18 6 6 18", "m6 6 12 12"]);
export const ChevronUp = icon(["m18 15-6-6-6 6"]);
export const DecimalDecreaseIcon = icon([
  "m13 21-3-3 3-3",
  "M20 18H10",
  "M3 11h.01",
  "M10 3a3 3 0 0 1 3 3v2a3 3 0 0 1-6 0V6a3 3 0 0 1 3-3Z",
]);
export const DecimalIncreaseIcon = icon([
  "M10 18h10",
  "m17 21 3-3-3-3",
  "M3 11h.01",
  "M15 3a3 3 0 0 1 3 3v2a3 3 0 0 1-6 0V6a3 3 0 0 1 3-3Z",
  "M6 3a3 3 0 0 1 3 3v2a3 3 0 0 1-6 0V6a3 3 0 0 1 3-3Z",
]);
export const WrapIcon = icon(["M3 6h18", "M3 12h15a3 3 0 1 1 0 6h-4", "m16 16-2 2 2 2", "M3 18h7"]);
export const MergeIcon = icon(["M3 4v16", "M21 4v16", "M6 12h5", "m9 9 3 3-3 3", "M18 12h-5", "m15 9-3 3 3 3"]);
export const ChartIcon = icon(["M3 3v16a2 2 0 0 0 2 2h16", "M18 17V9", "M13 17V5", "M8 17v-3"]);
export const SearchIcon = icon(["M11 18a7 7 0 1 0 0-14 7 7 0 0 0 0 14Z", "m21 21-4.3-4.3"]);

/** A letter with a colour bar underneath, for text and fill colour buttons. */
export function ColorBarIcon({ color, children }: { color: string; children: React.ReactNode }) {
  return (
    <span className="flex flex-col items-center leading-none">
      <span className="flex h-[15px] items-end">{children}</span>
      <span className="mt-0.5 h-[3px] w-4 rounded-sm" style={{ background: color }} />
    </span>
  );
}

export type BorderMode = "all" | "outer" | "inner" | "top" | "bottom" | "left" | "right" | "none";

/** A small grid with the sides a border option draws picked out. */
export function BorderIcon({ mode, size = 18 }: { mode: BorderMode; size?: number }) {
  const strong = "currentColor";
  const faint = "currentColor";
  const lines: { d: string; on: boolean }[] = [
    { d: "M4 4h16", on: ["all", "outer", "top"].includes(mode) },
    { d: "M4 20h16", on: ["all", "outer", "bottom"].includes(mode) },
    { d: "M4 4v16", on: ["all", "outer", "left"].includes(mode) },
    { d: "M20 4v16", on: ["all", "outer", "right"].includes(mode) },
    { d: "M4 12h16", on: ["all", "inner"].includes(mode) },
    { d: "M12 4v16", on: ["all", "inner"].includes(mode) },
  ];
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      {lines.map((l) => (
        <path
          key={l.d}
          d={l.d}
          stroke={l.on ? strong : faint}
          strokeWidth={l.on ? 2 : 1}
          strokeOpacity={l.on ? 1 : 0.3}
          strokeDasharray={l.on ? undefined : "2 2"}
        />
      ))}
    </svg>
  );
}
