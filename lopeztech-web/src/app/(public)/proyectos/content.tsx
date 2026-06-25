"use client";

import { motion } from "framer-motion";
import { SectionHeader } from "@/components/ui/section-header";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Code2, ExternalLink } from "lucide-react";

const projects = [
  {
    title: "Music Box",
    category: "React + Música",
    description: "Aplicación musical que recomienda canciones personalizadas según el estado de ánimo y la actividad del usuario, con autenticación y una experiencia visual enfocada en descubrir música de forma sencilla.",
    tags: ["React", "API", "Auth", "UX"],
    codeUrl: "https://github.com/DuvanLope/PROYECTO-MUSICBOX.git",
    demoUrl: "#",
    featured: true,
  },
  {
    title: "Plataforma Web de Onboarding y Capacitación",
    category: "Web + Capacitación",
    description: "Diseño e implementación de una plataforma web para estandarizar el entrenamiento técnico de nuevos ingresos, con una experiencia de aprendizaje autónoma y escalable.",
    tags: ["HTML/CSS", "JavaScript"],
    result: "Capacitación técnica 100% autónoma.",
    codeUrl: "https://github.com/DuvanLope/We-Capacitacion.git",
  },
  {
    title: "Automatizaciones Inteligentes con n8n",
    category: "Automatización + Seguridad",
    description: "Dos flujos de automatización empresarial: (1) Pipeline de creación de usuarios que lee datos desde SharePoint y provisiona cuentas en Active Directory mediante APIs REST. (2) Analista de ciberseguridad con IA que monitorea el SIEM y genera informes de incidentes automáticamente.",
    tags: ["n8n", "SharePoint", "Active Directory", "REST API", "SIEM", "AI"],
    result: "Aprovisionamiento sin intervención manual + respuesta automática a amenazas.",
    codeUrl: "https://github.com/DuvanLope/N8N-Creacion-de-Usuarios.git",
  },
];

export function ProyectosContent() {
  return (
    <section className="py-24 px-6">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Proyectos"
          title="Proyectos Seleccionados"
          description="Estos proyectos fueron redactados para mostrar resultados, contexto de negocio y tecnologías, no solo una lista de herramientas."
        />

        <div className="space-y-6">
          {projects.map((project, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 30 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.5, delay: i * 0.1 }}
            >
              <Card className={project.featured ? "border-accent/20 bg-gradient-to-br from-accent/5 to-card/50" : ""}>
                <div className="p-2">
                  <div className="flex items-center gap-3 mb-3">
                    <span className="text-xs text-accent font-mono">{project.category}</span>
                    {project.featured && <Badge variant="accent">Destacado</Badge>}
                  </div>
                  <h3 className="text-xl font-semibold text-foreground mb-3">{project.title}</h3>
                  <p className="text-sm text-muted-foreground mb-4 leading-relaxed max-w-3xl">{project.description}</p>
                  {project.result && (
                    <p className="text-sm text-accent font-medium mb-4">{project.result}</p>
                  )}
                  <div className="flex flex-wrap gap-2 mb-5">
                    {project.tags.map((tag) => <Badge key={tag}>{tag}</Badge>)}
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
      </div>
    </section>
  );
}
