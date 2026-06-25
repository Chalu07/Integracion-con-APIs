import { NextResponse } from "next/server";
import { prisma } from "@/lib/prisma";
import { contactSchema } from "@/lib/validations";
import { rateLimit } from "@/lib/rate-limit";
import { sanitizeHtml } from "@/lib/utils";
import { headers } from "next/headers";

export async function POST(request: Request) {
  try {
    const headersList = await headers();
    const ip = headersList.get("x-forwarded-for") ?? "unknown";

    if (!rateLimit(ip, 5, 60000)) {
      return NextResponse.json(
        { error: "Demasiadas solicitudes. Intenta más tarde." },
        { status: 429 }
      );
    }

    const body = await request.json();
    const result = contactSchema.safeParse(body);

    if (!result.success) {
      return NextResponse.json(
        { error: "Datos inválidos", details: result.error.flatten() },
        { status: 400 }
      );
    }

    const data = result.data;

    const submission = await prisma.contactSubmission.create({
      data: {
        name: sanitizeHtml(data.name),
        email: sanitizeHtml(data.email),
        subject: data.subject ? sanitizeHtml(data.subject) : null,
        message: sanitizeHtml(data.message),
      },
    });

    return NextResponse.json(
      { success: true, id: submission.id },
      { status: 201 }
    );
  } catch {
    return NextResponse.json(
      { error: "Error interno del servidor" },
      { status: 500 }
    );
  }
}
