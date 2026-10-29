import AccountMenu from "@/components/auth/AccountMenu";
import { SheetList } from "@/components/ClientOnly";

const FEATURES = [
  { title: "Live formulas", body: "SUM, IF, VLOOKUP, COUNTIF and 40+ functions, across tabs with Sheet2!A1." },
  { title: "Format like a pro", body: "Currency, percents, fonts, colours, borders and alignment from one toolbar." },
  { title: "Saved everywhere", body: "Sign in and your spreadsheets follow you to any device, and still work offline." },
];

export default function Home() {
  return (
    <main className="mx-auto flex min-h-dvh max-w-3xl flex-col px-4 py-12 sm:py-20">
      <div className="mb-10 flex items-center gap-3">
        <span className="grid h-10 w-10 place-items-center rounded-lg bg-[var(--accent)] text-lg font-bold text-white">
          K
        </span>
        <span className="flex-1 text-xl font-semibold">Kexcel</span>
        <AccountMenu />
      </div>

      <h1 className="text-3xl font-bold tracking-tight sm:text-5xl">Spreadsheets that just work.</h1>
      <p className="mt-4 max-w-xl text-lg text-[var(--muted)]">
        A fast, simple spreadsheet for small businesses, students and teams. Open it and start typing.
      </p>

      <div className="mt-8">
        <SheetList />
      </div>

      <ul className="mt-16 grid gap-4 sm:grid-cols-3">
        {FEATURES.map((f) => (
          <li key={f.title} className="rounded-lg border border-[var(--border)] p-4">
            <h3 className="font-medium">{f.title}</h3>
            <p className="mt-1 text-sm text-[var(--muted)]">{f.body}</p>
          </li>
        ))}
      </ul>
    </main>
  );
}
