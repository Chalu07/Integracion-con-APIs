import { NextResponse } from "next/server";
import { prisma } from "@/lib/prisma";
import { requireAdmin } from "@/lib/admin-auth";
import { blogPostSchema } from "@/lib/validations";

export async function GET() {
  const { error } = await requireAdmin();
  if (error) return error;

  const posts = await prisma.blogPost.findMany({ orderBy: { createdAt: "desc" } });
  return NextResponse.json(posts);
}

export async function POST(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const result = blogPostSchema.safeParse(body);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos", details: result.error.flatten() }, { status: 400 });
  }

  const post = await prisma.blogPost.create({
    data: {
      ...result.data,
      publishedAt: result.data.publishedAt ? new Date(result.data.publishedAt) : null,
    },
  });
  return NextResponse.json(post, { status: 201 });
}

export async function PUT(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const body = await request.json();
  const { id, ...data } = body;
  const result = blogPostSchema.safeParse(data);

  if (!result.success) {
    return NextResponse.json({ error: "Datos inválidos" }, { status: 400 });
  }

  const post = await prisma.blogPost.update({
    where: { id },
    data: {
      ...result.data,
      publishedAt: result.data.publishedAt ? new Date(result.data.publishedAt) : null,
    },
  });
  return NextResponse.json(post);
}

export async function DELETE(request: Request) {
  const { error } = await requireAdmin();
  if (error) return error;

  const { searchParams } = new URL(request.url);
  const id = searchParams.get("id");

  if (!id) return NextResponse.json({ error: "ID requerido" }, { status: 400 });

  await prisma.blogPost.delete({ where: { id } });
  return NextResponse.json({ success: true });
}
