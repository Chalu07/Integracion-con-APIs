"use client";

import { motion } from "framer-motion";
import { SectionHeader } from "@/components/ui/section-header";
import { Card } from "@/components/ui/card";

const certificates = [
  {
    badge: "01",
    title: "Tecnología en Desarrollo de Software",
    description: "Formación en programación, desarrollo web, bases de datos, análisis de requisitos y construcción de soluciones técnicas.",
    institution: "Universidad Pascual Bravo · En curso",
  },
  {
    badge: "02",
    title: "Introducción a los Algoritmos",
    description: "Fundamentos de lógica de programación, pseudocódigo, estructuras de control y pensamiento computacional.",
    institution: "TodoCode · Ene 2025",
  },
  {
    badge: "03",
    title: "Bases de Datos Relacionales con MySQL",
    description: "Modelado relacional, consultas SQL, joins, normalización y buenas prácticas para diseño de bases de datos.",
    institution: "TodoCode · Ene 2025",
  },
  {
    badge: "04",
    title: "Introducción al Big Data",
    description: "Conceptos fundamentales de ecosistemas de datos masivos, herramientas de análisis y almacenamiento distribuido.",
    institution: "TodoCode · Ene 2025",
  },
  {
    badge: "05",
    title: "Fundamentos de Seguridad de la Información",
    description: "Conceptos esenciales de ciberseguridad, gestión de riesgos, control de accesos y principios de seguridad.",
    institution: "TodoCode · Ene 2025",
  },
  {
    badge: "06",
    title: "OWASP Top 10 Vulnerabilidades",
    description: "Identificación y comprensión de las principales vulnerabilidades web, con enfoque en prevención y mitigación.",
    institution: "TodoCode · Ene 2025",
  },
];

export function CertificadosContent() {
  return (
    <section className="py-24 px-6">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Certificados"
          title="Formación"
          description="Certificaciones y estudios destacados para mostrar una ruta clara de crecimiento profesional."
        />

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {certificates.map((cert, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 30 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.5, delay: i * 0.08 }}
            >
              <Card className="h-full flex flex-col">
                <div className="p-2 flex-1">
                  <span className="inline-block text-xs font-mono text-accent bg-accent/10 px-2.5 py-1 rounded-full mb-4">
                    {cert.badge}
                  </span>
                  <h3 className="font-semibold text-foreground mb-2">{cert.title}</h3>
                  <p className="text-sm text-muted-foreground leading-relaxed">{cert.description}</p>
                </div>
                <div className="p-2 pt-4 border-t border-border mt-4">
                  <p className="text-xs text-muted-foreground">{cert.institution}</p>
                </div>
              </Card>
            </motion.div>
          ))}
        </div>
      </div>
    </section>
  );
}
