"use client";

import { useSyncExternalStore } from "react";
import { getServerSession, getSession, subscribeSession } from "@/lib/cloud";

/** The signed-in account, loaded once and shared across the app. */
export function useSession() {
  return useSyncExternalStore(subscribeSession, getSession, getServerSession);
}
