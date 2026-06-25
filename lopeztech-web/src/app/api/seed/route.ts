import { NextResponse } from "next/server";
import { prisma } from "@/lib/prisma";
import bcrypt from "bcryptjs";

export async function POST() {
  if (process.env.NODE_ENV === "production") {
    return NextResponse.json({ error: "Not available in production" }, { status: 403 });
  }

  try {
    const hashedPassword = await bcrypt.hash(
      process.env.ADMIN_PASSWORD || "admin123",
      12
    );

    await prisma.user.upsert({
      where: { email: process.env.ADMIN_EMAIL || "admin@lopeztech.dev" },
      update: { password: hashedPassword, role: "ADMIN" },
      create: {
        email: process.env.ADMIN_EMAIL || "admin@lopeztech.dev",
        name: "Duvan López",
        password: hashedPassword,
        role: "ADMIN",
      },
    });

    await prisma.service.createMany({
      skipDuplicates: true,
      data: [
        {
          title: "Desarrollo de Software",
          titleEn: "Software Development",
          description: "Aplicaciones web y móviles con tecnologías modernas como React, Next.js, TypeScript y más.",
          descriptionEn: "Web and mobile applications with modern technologies like React, Next.js, TypeScript and more.",
          icon: "Code",
          slug: "desarrollo-software",
          order: 1,
        },
        {
          title: "Automatizaciones",
          titleEn: "Automation",
          description: "Flujos automatizados con n8n, Power Automate y APIs REST para optimizar procesos empresariales.",
          descriptionEn: "Automated flows with n8n, Power Automate and REST APIs to optimize business processes.",
          icon: "Zap",
          slug: "automatizaciones",
          order: 2,
        },
        {
          title: "Ciberseguridad",
          titleEn: "Cybersecurity",
          description: "Análisis de vulnerabilidades, pruebas de penetración y estrategias de seguridad para aplicaciones.",
          descriptionEn: "Vulnerability analysis, penetration testing and security strategies for applications.",
          icon: "Shield",
          slug: "ciberseguridad",
          order: 3,
        },
      ],
    });

    await prisma.project.createMany({
      skipDuplicates: true,
      data: [
        {
          title: "Music Box",
          titleEn: "Music Box",
          description: "Aplicación musical que recomienda canciones personalizadas según el estado de ánimo.",
          descriptionEn: "Music app that recommends personalized songs based on mood.",
          tags: ["React", "API", "Auth"],
          codeUrl: "https://github.com/DuvanLope/PROYECTO-MUSICBOX.git",
          slug: "music-box",
          featured: true,
          order: 1,
        },
        {
          title: "Plataforma Web de Onboarding",
          titleEn: "Onboarding Web Platform",
          description: "Plataforma web para estandarizar el entrenamiento técnico de nuevos ingresos.",
          descriptionEn: "Web platform to standardize technical training for new employees.",
          tags: ["HTML/CSS", "JavaScript"],
          result: "Capacitación técnica 100% autónoma.",
          resultEn: "100% autonomous technical training.",
          codeUrl: "https://github.com/DuvanLope/We-Capacitacion.git",
          slug: "plataforma-onboarding",
          order: 2,
        },
      ],
    });

    await prisma.siteSettings.upsert({
      where: { id: "main" },
      update: {},
      create: {
        siteName: "LopezTech",
        siteUrl: "https://lopeztech.dev",
        description: "Portafolio profesional de Duvan López",
        email: "rendonfredy31@gmail.com",
        github: "https://github.com/DuvanLope",
        linkedin: "https://linkedin.com/in/duvanlopez",
      },
    });

    return NextResponse.json({ success: true, message: "Seed completed" });
  } catch (error) {
    console.error("Seed error:", error);
    return NextResponse.json(
      { error: "Seed failed", details: String(error) },
      { status: 500 }
    );
  }
}
