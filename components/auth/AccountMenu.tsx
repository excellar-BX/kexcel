"use client";

import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { useCallback, useRef, useState } from "react";
import { menuPanelClass, useDismiss } from "@/components/ui/Menu";
import { setSessionAccount, signOut } from "@/lib/cloud";
import { useSession } from "./useSession";

/** "Sign in" when signed out; otherwise the account's initial, opening a small menu. */
export default function AccountMenu() {
  const session = useSession();
  const router = useRouter();
  const pathname = usePathname();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setOpen(false), []);
  useDismiss(ref, close, open);

  if (session.status === "loading") return <span className="h-8 w-8" aria-hidden="true" />;
  const account = session.account;
  if (!account) {
    return (
      <Link
        href={`/login?next=${encodeURIComponent(pathname)}`}
        className="rounded-lg border border-[var(--border)] px-3 py-1.5 text-sm font-medium hover:bg-[var(--hover)]"
      >
        Sign in
      </Link>
    );
  }
  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={`Account: ${account.name}`}
        className="grid h-8 w-8 place-items-center rounded-full bg-[var(--header-selected)] text-sm font-semibold text-[var(--accent)] hover:ring-2 hover:ring-[var(--accent)]/30"
      >
        {account.name.slice(0, 1).toUpperCase()}
      </button>
      {open && (
        <div className="absolute right-0 top-full z-50 mt-1">
          <div className={`${menuPanelClass} p-1`} role="menu">
            <div className="px-3 py-2">
              <p className="text-sm font-medium">{account.name}</p>
              <p className="text-xs text-[var(--muted)]">{account.email}</p>
            </div>
            <div className="my-1 border-t border-[var(--border)]" />
            <button
              type="button"
              role="menuitem"
              onClick={async () => {
                setOpen(false);
                await signOut().catch(() => undefined);
                setSessionAccount(null);
                router.push("/");
              }}
              className="w-full rounded px-3 py-1.5 text-left text-sm hover:bg-[var(--hover)]"
            >
              Sign out
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
