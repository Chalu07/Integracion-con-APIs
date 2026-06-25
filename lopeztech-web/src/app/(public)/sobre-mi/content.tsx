"use client";

import { motion } from "framer-motion";
import { SectionHeader } from "@/components/ui/section-header";
import { Card } from "@/components/ui/card";
import { Code, Zap, Shield, GraduationCap } from "lucide-react";

const highlights = [
  {
    icon: Code,
    title: "Desarrollo de Software",
    desc: "Experiencia en React, Next.js, TypeScript, Python, Java, C# y Flutter. Construcción de aplicaciones web modernas y escalables.",
  },
  {
    icon: Zap,
    title: "Automatizaciones Empresariales",
    desc: "Diseño e implementación de flujos con n8n, Power Automate, APIs REST. Integración de SharePoint, Active Directory y Microsoft 365.",
  },
  {
    icon: Shield,
    title: "QA & Ciberseguridad",
    desc: "Pruebas automatizadas con Selenium y Pytest. Análisis de vulnerabilidades, monitoreo SIEM y respuesta a incidentes de seguridad.",
  },
  {
    icon: GraduationCap,
    title: "Formación Continua",
    desc: "Tecnología en Desarrollo de Software — Universidad Pascual Bravo. Certificaciones en algoritmos, bases de datos, Big Data, seguridad y más.",
  },
];

export function SobreMiContent() {
  return (
    <section className="py-24 px-6">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Sobre mí"
          title="Duvan López"
          description="Desarrollador de Software y QA Engineer con experiencia en automatizaciones empresariales, APIs, SharePoint, Active Directory y ciberseguridad."
        />

        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          transition={{ duration: 0.5 }}
          className="prose prose-invert max-w-3xl mb-16"
        >
          <p className="text-lg text-muted-foreground leading-relaxed">
            Soy un profesional apasionado por la tecnología con enfoque en la resolución de problemas
            reales mediante desarrollo de software, automatización de procesos y seguridad informática.
            Actualmente trabajo como Analista de Automatizaciones en Socia BPO, donde diseño e implemento
            flujos de automatización que optimizan procesos operativos internos.
          </p>
          <p className="text-lg text-muted-foreground leading-relaxed mt-4">
            Mi experiencia abarca desde desarrollo web full-stack hasta integración de sistemas empresariales,
            pasando por QA engineering y análisis de ciberseguridad. He trabajado con empresas como Hummingbirds AI
            y Ycaycom, además de clientes independientes en diversos sectores.
          </p>
        </motion.div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {highlights.map((item, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 30 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.5, delay: i * 0.1 }}
            >
              <Card className="h-full">
                <div className="flex gap-4 p-2">
                  <div className="w-12 h-12 rounded-xl bg-accent/10 flex items-center justify-center shrink-0">
                    <item.icon size={24} className="text-accent" />
                  </div>
                  <div>
                    <h3 className="font-semibold text-foreground mb-1">{item.title}</h3>
                    <p className="text-sm text-muted-foreground leading-relaxed">{item.desc}</p>
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
