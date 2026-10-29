"use client";

import Link from "next/link";
import { useRouter, useSearchParams } from "next/navigation";
import { useState } from "react";
import { CloudError, setSessionAccount, signIn, signUp } from "@/lib/cloud";

const input =
  "w-full rounded-lg border border-[var(--border)] bg-[var(--bg)] px-3 py-2.5 text-sm outline-none focus:border-[var(--accent)] focus:ring-2 focus:ring-[var(--accent)]/20";

/** Only same-site paths are allowed as the place to return to after signing in. */
function safeNext(next: string | null) {
  return next && next.startsWith("/") && !next.startsWith("//") ? next : "/";
}

export default function AuthForm() {
  const router = useRouter();
  const params = useSearchParams();
  const [mode, setMode] = useState<"signin" | "signup">(params.get("mode") === "signup" ? "signup" : "signin");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setError(null);
    try {
      const account = mode === "signup" ? await signUp(email, password, name) : await signIn(email, password);
      setSessionAccount(account);
      router.push(safeNext(params.get("next")));
    } catch (err) {
      setError(err instanceof CloudError ? err.message : "Something went wrong. Please try again.");
      setBusy(false);
    }
  };

  return (
    <div className="w-full max-w-sm">
      <Link href="/" className="mb-8 flex items-center gap-2">
        <span className="grid h-9 w-9 place-items-center rounded-lg bg-[var(--accent)] font-bold text-white">K</span>
        <span className="text-lg font-semibold">Kexcel</span>
      </Link>
      <h1 className="text-2xl font-bold tracking-tight">{mode === "signup" ? "Create your account" : "Sign in"}</h1>
      <p className="mt-1 text-sm text-[var(--muted)]">
        {mode === "signup"
          ? "Your spreadsheets will be saved to your account and open on any device."
          : "Open your spreadsheets on any device."}
      </p>

      <form onSubmit={submit} className="mt-6 space-y-3">
        {mode === "signup" && (
          <label className="block">
            <span className="mb-1 block text-sm font-medium">Name</span>
            <input value={name} onChange={(e) => setName(e.target.value)} autoComplete="name" className={input} />
          </label>
        )}
        <label className="block">
          <span className="mb-1 block text-sm font-medium">Email</span>
          <input
            type="email"
            required
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            autoComplete="email"
            className={input}
          />
        </label>
        <label className="block">
          <span className="mb-1 block text-sm font-medium">Password</span>
          <input
            type="password"
            required
            minLength={mode === "signup" ? 8 : undefined}
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoComplete={mode === "signup" ? "new-password" : "current-password"}
            className={input}
          />
          {mode === "signup" && <span className="mt-1 block text-xs text-[var(--muted)]">At least 8 characters.</span>}
        </label>
        {error && (
          <p role="alert" className="text-sm text-[var(--error)]">
            {error}
          </p>
        )}
        <button
          type="submit"
          disabled={busy}
          className="w-full rounded-lg bg-[var(--accent)] px-4 py-2.5 font-medium text-white transition hover:opacity-90 disabled:opacity-50"
        >
          {busy ? "One moment…" : mode === "signup" ? "Create account" : "Sign in"}
        </button>
      </form>

      <p className="mt-6 text-center text-sm text-[var(--muted)]">
        {mode === "signup" ? "Already have an account?" : "New to Kexcel?"}{" "}
        <button
          type="button"
          onClick={() => {
            setMode(mode === "signup" ? "signin" : "signup");
            setError(null);
          }}
          className="font-medium text-[var(--accent)] hover:underline"
        >
          {mode === "signup" ? "Sign in" : "Create an account"}
        </button>
      </p>
    </div>
  );
}
