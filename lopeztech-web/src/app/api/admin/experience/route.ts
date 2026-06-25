import { NextResponse } from "next/server";
import { prisma } from "@/lib/prisma";
import { requireAdmin } from "@/lib/admin-auth";
import { experienceSchema } from "@/lib/validations";

export async function GET() {
  const { error } = await requireAdmin();
  if (error) return error;

  const experiences = await prisma.experience.findMany({ orderBy: { order: "asc" } });
  return NextResponse.json(experiences);
}

export async function POST(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const result = experienceSchema.safeParse(body);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos", details: result.error.flatten() }, { status: 400 });
  }

  const experience = await prisma.experience.create({ data: result.data });
  return NextResponse.json(experience, { status: 201 });
}

export async function PUT(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const { id, ...data } = body;
  const result = experienceSchema.safeParse(data);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos" }, { status: 400 });
  }

  const experience = await prisma.experience.update({ where: { id }, data: result.data });
  return NextResponse.json(experience);
}

export async function DELETE(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const { searchParams } = new URL(request.url);
  const id = searchParams.get("id");

  if (!id) return NextResponse.json({ error: "ID requerido" }, { status: 400 });

  await prisma.experience.delete({ where: { id } });
  return NextResponse.json({ success: true });
}
