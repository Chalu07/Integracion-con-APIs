"use client";

import Link from "next/link";
import { useEffect, useSyncExternalStore } from "react";
import { useRouter } from "next/navigation";
import { useSession, signOut } from "next-auth/react";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import {
  Briefcase,
  FolderGit2,
  Cpu,
  MessageSquare,
  Users,
  FileText,
  Mail,
  GraduationCap,
  Award,
  LogOut,
  LayoutDashboard,
} from "lucide-react";

const adminSections = [
  { href: "/admin/servicios", icon: Briefcase, label: "Servicios", desc: "Gestionar servicios ofrecidos" },
  { href: "/admin/proyectos", icon: FolderGit2, label: "Proyectos", desc: "Gestionar portafolio de proyectos" },
  { href: "/admin/tecnologias", icon: Cpu, label: "Tecnologías", desc: "Gestionar stack tecnológico" },
  { href: "/admin/testimonios", icon: MessageSquare, label: "Testimonios", desc: "Gestionar testimonios de clientes" },
  { href: "/admin/clientes", icon: Users, label: "Clientes", desc: "Gestionar lista de clientes" },
  { href: "/admin/blog", icon: FileText, label: "Blog", desc: "Crear y gestionar artículos" },
  { href: "/admin/contacto", icon: Mail, label: "Formularios", desc: "Ver solicitudes recibidas" },
  { href: "/admin/experiencia", icon: GraduationCap, label: "Experiencia", desc: "Gestionar experiencia profesional" },
  { href: "/admin/certificados", icon: Award, label: "Certificados", desc: "Gestionar certificaciones" },
];

const emptySubscribe = () => () => {};

export default function AdminDashboard() {
  const { data: session, status } = useSession();
  const router = useRouter();
  const mounted = useSyncExternalStore(
    emptySubscribe,
    () => true,
    () => false
  );

  useEffect(() => {
    if (status === "unauthenticated") {
      router.push("/admin/login");
    }
  }, [status, router]);

  if (!mounted || status === "loading") {
    return (
      <div className="min-h-screen flex items-center justify-center bg-background">
        <div className="animate-spin w-8 h-8 border-2 border-accent border-t-transparent rounded-full" />
      </div>
    );
  }

  if (!session) return null;

  return (
    <div className="min-h-screen bg-background">
      <header className="sticky top-0 z-50 border-b border-border bg-background/80 backdrop-blur-xl">
        <div className="mx-auto max-w-6xl flex items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <LayoutDashboard size={20} className="text-accent" />
            <h1 className="text-lg font-bold">
              <span className="text-accent">Lopez</span>Tech CMS
            </h1>
          </div>
          <div className="flex items-center gap-4">
            <span className="text-sm text-muted-foreground hidden sm:block">
              {session.user?.name || session.user?.email}
            </span>
            <Button
              variant="ghost"
              size="sm"
              onClick={() => signOut({ callbackUrl: "/admin/login" })}
            >
              <LogOut size={16} /> Salir
            </Button>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-6 py-10">
        <div className="mb-8">
          <h2 className="text-2xl font-bold text-foreground">Dashboard</h2>
          <p className="text-muted-foreground mt-1">Gestiona todo el contenido de tu portafolio.</p>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {adminSections.map((section) => (
            <Link key={section.href} href={section.href}>
              <Card className="h-full group cursor-pointer">
                <div className="flex items-start gap-4 p-2">
                  <div className="w-10 h-10 rounded-xl bg-accent/10 flex items-center justify-center shrink-0 group-hover:bg-accent/20 transition-colors">
                    <section.icon size={20} className="text-accent" />
                  </div>
                  <div>
                    <h3 className="font-semibold text-foreground group-hover:text-accent transition-colors">
                      {section.label}
                    </h3>
                    <p className="text-xs text-muted-foreground mt-0.5">{section.desc}</p>
                  </div>
                </div>
              </Card>
            </Link>
          ))}
        </div>
      </main>
    </div>
  );
}
