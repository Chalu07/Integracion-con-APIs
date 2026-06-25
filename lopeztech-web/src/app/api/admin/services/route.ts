import { NextResponse } from "next/server";
import { prisma } from "@/lib/prisma";
import { requireAdmin } from "@/lib/admin-auth";
import { serviceSchema } from "@/lib/validations";

export async function GET() {
  const { error } = await requireAdmin();
  if (error) return error;

  const services = await prisma.service.findMany({ orderBy: { order: "asc" } });
  return NextResponse.json(services);
}

export async function POST(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const result = serviceSchema.safeParse(body);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos", details: result.error.flatten() }, { status: 400 });
  }

  const service = await prisma.service.create({ data: result.data });
  return NextResponse.json(service, { status: 201 });
}

export async function PUT(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const { id, ...data } = body;
  const result = serviceSchema.safeParse(data);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos" }, { status: 400 });
  }

  const service = await prisma.service.update({ where: { id }, data: result.data });
  return NextResponse.json(service);
}

export async function DELETE(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const { searchParams } = new URL(request.url);
  const id = searchParams.get("id");

  if (!id) return NextResponse.json({ error: "ID requerido" }, { status: 400 });

  await prisma.service.delete({ where: { id } });
  return NextResponse.json({ success: true });
}
