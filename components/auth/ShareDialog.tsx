"use client";

import { useEffect, useRef, useState } from "react";
import { CheckIcon, XIcon } from "@/components/ui/icons";
import { useDismiss } from "@/components/ui/Menu";
import { CloudError, getSharing, setLinkAccess, shareWith, unshare, type Share } from "@/lib/cloud";

const field =
  "rounded-md border border-[var(--border)] bg-[var(--bg)] px-2.5 py-1.5 text-sm outline-none focus:border-[var(--accent)]";

/** Who a spreadsheet is shared with, and whether its link opens it. For the owner. */
export default function ShareDialog({ docId, title, onClose }: { docId: string; title: string; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  useDismiss(ref, onClose);
  const [shares, setShares] = useState<Share[] | null>(null);
  const [linkAccess, setLink] = useState<"none" | "view">("none");
  const [email, setEmail] = useState("");
  const [role, setRole] = useState<"view" | "edit">("edit");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  const reload = () =>
    getSharing(docId)
      .then((s) => {
        setShares(s.shares);
        setLink(s.linkAccess);
      })
      .catch((e) => setError(e instanceof CloudError ? e.message : "Couldn't load sharing."));

  useEffect(() => {
    void reload();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [docId]);

  const run = async (fn: () => Promise<unknown>) => {
    setBusy(true);
    setError(null);
    try {
      await fn();
      await reload();
    } catch (e) {
      setError(e instanceof CloudError ? e.message : "Something went wrong.");
    } finally {
      setBusy(false);
    }
  };

  const link = typeof window !== "undefined" ? `${window.location.origin}/sheet/${docId}` : "";

  return (
    <div className="fixed inset-0 z-50 grid place-items-center bg-black/30 p-4">
      <div
        ref={ref}
        role="dialog"
        aria-modal="true"
        aria-labelledby="share-title"
        className="w-full max-w-md space-y-4 rounded-xl border border-[var(--border)] bg-[var(--panel)] p-5 shadow-2xl"
      >
        <div className="flex items-center gap-2">
          <h2 id="share-title" className="min-w-0 flex-1 truncate text-lg font-semibold">
            Share &ldquo;{title}&rdquo;
          </h2>
          <button type="button" onClick={onClose} aria-label="Close" className="grid h-7 w-7 place-items-center rounded hover:bg-[var(--hover)]">
            <XIcon size={16} />
          </button>
        </div>

        <form
          className="flex gap-2"
          onSubmit={(e) => {
            e.preventDefault();
            if (!email.trim()) return;
            void run(async () => {
              await shareWith(docId, email, role);
              setEmail("");
            });
          }}
        >
          <input
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="Their account email"
            aria-label="Email of the person to share with"
            className={`${field} min-w-0 flex-1`}
          />
          <select value={role} onChange={(e) => setRole(e.target.value as "view" | "edit")} aria-label="Access" className={field}>
            <option value="edit">Can edit</option>
            <option value="view">Can view</option>
          </select>
          <button
            type="submit"
            disabled={busy || !email.trim()}
            className="rounded-md bg-[var(--accent)] px-3 py-1.5 text-sm font-medium text-white hover:opacity-90 disabled:opacity-40"
          >
            Share
          </button>
        </form>
        {error && (
          <p role="alert" className="text-sm text-[var(--error)]">
            {error}
          </p>
        )}

        <div>
          <h3 className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-[var(--muted)]">People with access</h3>
          {shares === null ? (
            <p className="text-sm text-[var(--muted)]">Loading…</p>
          ) : shares.length === 0 ? (
            <p className="text-sm text-[var(--muted)]">Only you, so far.</p>
          ) : (
            <ul className="space-y-1.5">
              {shares.map((s) => (
                <li key={s.userId} className="flex items-center gap-2 text-sm">
                  <span className="grid h-7 w-7 shrink-0 place-items-center rounded-full bg-[var(--header-selected)] text-xs font-semibold text-[var(--accent)]">
                    {s.name.slice(0, 1).toUpperCase()}
                  </span>
                  <span className="min-w-0 flex-1">
                    <span className="block truncate">{s.name}</span>
                    <span className="block truncate text-xs text-[var(--muted)]">{s.email}</span>
                  </span>
                  <select
                    value={s.role}
                    disabled={busy}
                    onChange={(e) => void run(() => shareWith(docId, s.email, e.target.value as "view" | "edit"))}
                    aria-label={`Access for ${s.name}`}
                    className={field}
                  >
                    <option value="edit">Can edit</option>
                    <option value="view">Can view</option>
                  </select>
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void run(() => unshare(docId, s.userId))}
                    className="rounded px-1.5 py-1 text-xs text-[var(--muted)] hover:bg-[var(--hover)] hover:text-[var(--error)]"
                  >
                    Remove
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="space-y-2 border-t border-[var(--border)] pt-4">
          <label className="flex items-start gap-2 text-sm">
            <input
              type="checkbox"
              checked={linkAccess === "view"}
              disabled={busy}
              onChange={(e) => void run(() => setLinkAccess(docId, e.target.checked ? "view" : "none"))}
              className="mt-0.5"
            />
            <span>
              Anyone with the link can view
              <span className="block text-xs text-[var(--muted)]">They can look but not change anything, even without an account.</span>
            </span>
          </label>
          <div className="flex gap-2">
            <input readOnly value={link} aria-label="Link to this spreadsheet" className={`${field} min-w-0 flex-1 font-mono text-xs`} onFocus={(e) => e.target.select()} />
            <button
              type="button"
              onClick={() => {
                void navigator.clipboard?.writeText(link).then(() => {
                  setCopied(true);
                  setTimeout(() => setCopied(false), 2000);
                });
              }}
              className="flex items-center gap-1.5 rounded-md border border-[var(--border)] px-3 py-1.5 text-sm hover:bg-[var(--hover)]"
            >
              {copied && <CheckIcon size={14} />}
              {copied ? "Copied" : "Copy link"}
            </button>
          </div>
          <p className="text-xs text-[var(--muted)]">People you share with open this link after signing in.</p>
        </div>
      </div>
    </div>
  );
}
