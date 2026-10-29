import { Suspense } from "react";
import AuthForm from "@/components/auth/AuthForm";

export const metadata = { title: "Sign in · Kexcel" };

export default function LoginPage() {
  return (
    <main className="flex min-h-dvh items-center justify-center px-4 py-12">
      {/* The form reads ?next= and ?mode= from the URL, which is only known in the browser. */}
      <Suspense>
        <AuthForm />
      </Suspense>
    </main>
  );
}
