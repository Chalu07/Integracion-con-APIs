import { NextResponse } from "next/server";
import { prisma } from "@/lib/prisma";
import { requireAdmin } from "@/lib/admin-auth";
import { projectSchema } from "@/lib/validations";

export async function GET() {
  const { error } = await requireAdmin();
  if (error) return error;

  const projects = await prisma.project.findMany({ orderBy: { order: "asc" } });
  return NextResponse.json(projects);
}

export async function POST(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const result = projectSchema.safeParse(body);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos", details: result.error.flatten() }, { status: 400 });
  }

  const project = await prisma.project.create({ data: result.data });
  return NextResponse.json(project, { status: 201 });
}

export async function PUT(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const { id, ...data } = body;
  const result = projectSchema.safeParse(data);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos" }, { status: 400 });
  }

  const project = await prisma.project.update({ where: { id }, data: result.data });
  return NextResponse.json(project);
}

export async function DELETE(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const { searchParams } = new URL(request.url);
  const id = searchParams.get("id");

  if (!id) return NextResponse.json({ error: "ID requerido" }, { status: 400 });

  await prisma.project.delete({ where: { id } });
  return NextResponse.json({ success: true });
}
