"use client";

import { motion } from "framer-motion";
import Link from "next/link";
import { ExternalLink, Code2 } from "lucide-react";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { SectionHeader } from "@/components/ui/section-header";

const projects = [
  {
    title: "Music Box",
    category: "React + Música",
    description: "Aplicación musical que recomienda canciones personalizadas según el estado de ánimo y la actividad del usuario, con autenticación y experiencia visual moderna.",
    tags: ["React", "API", "Auth"],
    codeUrl: "https://github.com/DuvanLope/PROYECTO-MUSICBOX.git",
    demoUrl: "#",
    featured: true,
  },
  {
    title: "Plataforma Web de Onboarding",
    category: "Web + Capacitación",
    description: "Plataforma web para estandarizar el entrenamiento técnico de nuevos ingresos, con experiencia de aprendizaje autónoma y escalable.",
    tags: ["HTML/CSS", "JavaScript"],
    result: "Capacitación técnica 100% autónoma.",
    codeUrl: "https://github.com/DuvanLope/We-Capacitacion.git",
  },
  {
    title: "Automatizaciones Inteligentes con n8n",
    category: "Automatización + Seguridad",
    description: "Pipeline de creación de usuarios y analista de ciberseguridad con IA que monitorea el SIEM y responde a amenazas automáticamente.",
    tags: ["n8n", "SharePoint", "Active Directory", "SIEM", "AI"],
    result: "Aprovisionamiento sin intervención manual.",
    codeUrl: "https://github.com/DuvanLope/N8N-Creacion-de-Usuarios.git",
  },
];

export function ProjectsPreview() {
  return (
    <section className="py-24 px-6 bg-card/30">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Proyectos"
          title="Proyectos Seleccionados"
          description="Resultados reales con contexto de negocio y tecnologías aplicadas."
        />

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {projects.map((project, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 30 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 0.5, delay: i * 0.1 }}
              className={i === 0 ? "lg:col-span-2" : ""}
            >
              <Card className="h-full">
                <div className="p-2">
                  <div className="flex items-center justify-between mb-3">
                    <span className="text-xs text-accent font-mono">{project.category}</span>
                  </div>
                  <h3 className="text-xl font-semibold text-foreground mb-3">
                    {project.title}
                  </h3>
                  <p className="text-sm text-muted-foreground mb-4 leading-relaxed">
                    {project.description}
                  </p>
                  {project.result && (
                    <p className="text-sm text-accent mb-4 font-medium">
                      {project.result}
                    </p>
                  )}
                  <div className="flex flex-wrap gap-2 mb-4">
                    {project.tags.map((tag) => (
                      <Badge key={tag}>{tag}</Badge>
                    ))}
                  </div>
                  <div className="flex gap-3">
                    {project.codeUrl && (
                      <a href={project.codeUrl} target="_blank" rel="noopener noreferrer">
                        <Button variant="secondary" size="sm">
                          <Code2 size={16} /> Code
                        </Button>
                      </a>
                    )}
                    {project.demoUrl && (
                      <a href={project.demoUrl} target="_blank" rel="noopener noreferrer">
                        <Button size="sm">
                          <ExternalLink size={16} /> Demo
                        </Button>
                      </a>
                    )}
                  </div>
                </div>
              </Card>
            </motion.div>
          ))}
        </div>

        <div className="mt-10 text-center">
          <Link href="/proyectos">
            <Button variant="secondary">Ver todos los proyectos</Button>
          </Link>
        </div>
      </div>
    </section>
  );
}
