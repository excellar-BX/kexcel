// Browser side of accounts and cloud save: thin wrappers around the /api routes.
import type { WorkbookJSON } from "@/lib/engine";

export interface Account {
  id: string;
  email: string;
  name: string;
}

export interface CloudWorkbook {
  id: string;
  title: string;
  workbook: WorkbookJSON;
  activeSheet?: string;
  version: number;
  updatedAt: number;
}

export interface CloudSummary {
  id: string;
  title: string;
  version: number;
  updatedAt: number;
}

/** A request that failed: `offline` means the server couldn't be reached at all. */
export class CloudError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly offline = false,
  ) {
    super(message);
  }
}

async function call<T>(path: string, init?: RequestInit): Promise<T> {
  let res: Response;
  try {
    res = await fetch(path, {
      ...init,
      headers: { "content-type": "application/json", ...init?.headers },
      cache: "no-store",
    });
  } catch {
    throw new CloudError("You're offline. Changes are saved on this device.", 0, true);
  }
  const body = (await res.json().catch(() => ({}))) as T & { error?: string };
  if (!res.ok && res.status !== 409) throw new CloudError(body.error ?? `Request failed (${res.status})`, res.status);
  return body;
}

export const fetchAccount = () => call<{ user: Account | null }>("/api/auth/me").then((r) => r.user);

export const signUp = (email: string, password: string, name: string) =>
  call<{ user: Account }>("/api/auth/signup", { method: "POST", body: JSON.stringify({ email, password, name }) }).then((r) => r.user);

export const signIn = (email: string, password: string) =>
  call<{ user: Account }>("/api/auth/login", { method: "POST", body: JSON.stringify({ email, password }) }).then((r) => r.user);

export const signOut = () => call<{ ok: true }>("/api/auth/logout", { method: "POST", body: "{}" });

export const listCloud = () => call<{ workbooks: CloudSummary[] }>("/api/workbooks").then((r) => r.workbooks);

/** A spreadsheet from the account, or null if there's none with this id. */
export async function fetchCloud(id: string): Promise<CloudWorkbook | null> {
  try {
    return await call<CloudWorkbook>(`/api/workbooks/${encodeURIComponent(id)}`);
  } catch (e) {
    if (e instanceof CloudError && e.status === 404) return null;
    throw e;
  }
}

export type PushResult = { version: number } | { conflict: CloudWorkbook };

/** Saves to the account. `baseVersion` is the version these changes started from (null if new). */
export const pushCloud = (
  id: string,
  doc: { title: string; workbook: WorkbookJSON; activeSheet?: string },
  baseVersion: number | null,
) =>
  call<PushResult>(`/api/workbooks/${encodeURIComponent(id)}`, {
    method: "PUT",
    body: JSON.stringify({ ...doc, baseVersion }),
  });

export const deleteCloud = (id: string) => call<{ ok: true }>(`/api/workbooks/${encodeURIComponent(id)}`, { method: "DELETE" });

// ---- The signed-in account, shared by every component ----------------------

type Session = { status: "loading" } | { status: "ready"; account: Account | null };

let session: Session = { status: "loading" };
const listeners = new Set<() => void>();
let started = false;

const emit = () => listeners.forEach((l) => l());

export function setSessionAccount(account: Account | null) {
  session = { status: "ready", account };
  emit();
}

/** For useSyncExternalStore: subscribe, and load the account the first time anyone asks. */
export function subscribeSession(listener: () => void) {
  listeners.add(listener);
  if (!started) {
    started = true;
    fetchAccount()
      .then(setSessionAccount)
      // Offline or the server is down: carry on signed out, using this device.
      .catch(() => setSessionAccount(null));
  }
  return () => listeners.delete(listener);
}

export const getSession = () => session;
const loadingSession: Session = { status: "loading" };
export const getServerSession = () => loadingSession;
