import { auth } from "@/lib/auth";
import { NextResponse } from "next/server";

export async function requireAdmin() {
  const session = await auth();

  if (!session?.user) {
    return { error: NextResponse.json({ error: "No autorizado" }, { status: 401 }), session: null };
  }

  if ((session.user as { role?: string }).role !== "ADMIN") {
    return { error: NextResponse.json({ error: "Acceso denegado" }, { status: 403 }), session: null };
  }

  return { error: null, session };
}
