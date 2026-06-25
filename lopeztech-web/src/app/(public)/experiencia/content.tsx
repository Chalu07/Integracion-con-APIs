"use client";

import { motion } from "framer-motion";
import { SectionHeader } from "@/components/ui/section-header";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { ExternalLink } from "lucide-react";
import { Button } from "@/components/ui/button";

const laboralExp = [
  {
    period: "Nov 2025 — Jun 2026",
    title: "Analista de Automatizaciones",
    company: "Socia BPO",
    description: "Diseño e implementación de flujos de automatización para optimizar procesos operativos internos. Integración de herramientas empresariales mediante APIs y plataformas de automatización de bajo código.",
    tags: ["Power Automate", "Power Apps", "N8N", "APIs REST", "SharePoint", "AD Activo"],
    functions: [
      "Levantamiento y análisis de procesos repetitivos susceptibles de automatización.",
      "Desarrollo de flujos en Power Automate y N8N para reducir carga operativa.",
      "Integración de sistemas internos mediante conectores y APIs REST.",
      "Construcción de aplicaciones internas con Power Apps para gestión de datos.",
      "Documentación técnica de flujos y soluciones implementadas.",
      "Soporte y mantenimiento de automatizaciones en producción.",
    ],
  },
];

const freelanceExp = [
  {
    badge: "01",
    type: "QA · Ciberseguridad",
    title: "Hummingbirds AI",
    description: "Rol de QA Engineer en una empresa enfocada en inteligencia artificial. Encargado de diseñar y automatizar casos de prueba, validar APIs e identificar vulnerabilidades en sistemas y aplicaciones web.",
    tags: ["Python", "Selenium", "Pytest", "Postman", "REST APIs", "Pruebas de Vulnerabilidad"],
    website: "https://hummingbirds.ai/",
  },
  {
    badge: "02",
    type: "Desarrollo Web",
    title: "Ycaycom",
    description: "Mejora integral de la presencia web del cliente: optimización de rendimiento, actualización de estructura de contenidos, correcciones de diseño responsivo y mejoras en la experiencia de usuario.",
    tags: ["HTML / CSS", "JavaScript", "Diseño Responsivo", "UX / UI"],
    website: "https://www.ycaycom.com/",
  },
  {
    badge: "03",
    type: "Automatización · N8N",
    title: "Cliente — Sector Comercial",
    description: "Implementación de flujos automatizados en N8N para gestión de pedidos y notificaciones. Integración con plataformas de mensajería y hojas de cálculo para seguimiento en tiempo real.",
    tags: ["N8N", "Webhooks", "Google Sheets", "WhatsApp API"],
  },
  {
    badge: "04",
    type: "Automatización · N8N",
    title: "Cliente — Servicios Profesionales",
    description: "Desarrollo de automatizaciones en N8N para centralizar la recepción y distribución de leads. Conexión entre formularios web, CRM y correo electrónico.",
    tags: ["N8N", "CRM", "Email API", "Formularios Web"],
  },
];

export function ExperienciaContent() {
  return (
    <div className="py-24 px-6">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Experiencia"
          title="Trayectoria profesional"
          description="Roles desempeñados en empresa y proyectos independientes que reflejan mi evolución como profesional."
        />

        <div className="space-y-8 mb-20">
          {laboralExp.map((exp, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, x: -30 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.5 }}
            >
              <Card>
                <div className="p-2">
                  <p className="text-xs text-accent font-mono mb-2">{exp.period}</p>
                  <h3 className="text-xl font-semibold text-foreground">{exp.title}</h3>
                  <p className="text-sm text-accent mt-1 mb-3">{exp.company}</p>
                  <p className="text-sm text-muted-foreground mb-4 leading-relaxed">{exp.description}</p>
                  <div className="flex flex-wrap gap-2 mb-4">
                    {exp.tags.map((tag) => <Badge key={tag}>{tag}</Badge>)}
                  </div>
                  <div>
                    <p className="text-sm font-semibold text-foreground mb-2">Funciones principales</p>
                    <ul className="space-y-1.5">
                      {exp.functions.map((fn, j) => (
                        <li key={j} className="text-sm text-muted-foreground flex gap-2">
                          <span className="text-accent mt-1 shrink-0">&#8226;</span>
                          {fn}
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>
              </Card>
            </motion.div>
          ))}
        </div>

        <SectionHeader
          kicker="Freelance"
          title="Experiencia independiente"
          description="Clientes y proyectos en los que he trabajado de forma independiente, aplicando soluciones web y de automatización a medida."
        />

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {freelanceExp.map((exp, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 30 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.5, delay: i * 0.1 }}
            >
              <Card className="h-full flex flex-col">
                <div className="p-2 flex-1">
                  <div className="flex items-center gap-3 mb-3">
                    <span className="text-xs font-mono text-accent bg-accent/10 px-2 py-1 rounded-full">
                      {exp.badge}
                    </span>
                    <span className="text-xs text-muted-foreground">{exp.type}</span>
                  </div>
                  <h3 className="text-lg font-semibold text-foreground mb-2">{exp.title}</h3>
                  <p className="text-sm text-muted-foreground mb-4 leading-relaxed">{exp.description}</p>
                  <div className="flex flex-wrap gap-2">
                    {exp.tags.map((tag) => <Badge key={tag} variant="muted">{tag}</Badge>)}
                  </div>
                </div>
                {exp.website && (
                  <div className="p-2 pt-4 border-t border-border mt-4">
                    <a href={exp.website} target="_blank" rel="noopener noreferrer">
                      <Button variant="secondary" size="sm">
                        <ExternalLink size={14} /> Ver sitio web
                      </Button>
                    </a>
                  </div>
                )}
              </Card>
            </motion.div>
          ))}
        </div>
      </div>
    </div>
  );
}
