"use client";

import { motion } from "framer-motion";
import { Code, Zap, Shield, Database, Globe, GitBranch } from "lucide-react";
import { Card } from "@/components/ui/card";
import { SectionHeader } from "@/components/ui/section-header";

const services = [
  {
    icon: Code,
    title: "Desarrollo de Software",
    description: "Aplicaciones web y móviles con tecnologías modernas como React, Next.js, TypeScript y más.",
  },
  {
    icon: Zap,
    title: "Automatizaciones",
    description: "Flujos automatizados con n8n, Power Automate y APIs REST para optimizar procesos empresariales.",
  },
  {
    icon: Shield,
    title: "Ciberseguridad",
    description: "Análisis de vulnerabilidades, pruebas de penetración y estrategias de seguridad para aplicaciones.",
  },
  {
    icon: Database,
    title: "Integración de APIs",
    description: "Conexión de sistemas empresariales mediante APIs REST, webhooks y conectores personalizados.",
  },
  {
    icon: Globe,
    title: "SharePoint & AD",
    description: "Soluciones corporativas con SharePoint, Active Directory y herramientas Microsoft 365.",
  },
  {
    icon: GitBranch,
    title: "QA & Testing",
    description: "Diseño y automatización de pruebas con Selenium, Pytest y Postman para garantizar calidad.",
  },
];

export function ServicesPreview() {
  return (
    <section className="py-24 px-6" id="about-preview">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Servicios"
          title="Lo que hago"
          description="Soluciones integrales de desarrollo, automatización y seguridad para impulsar tu negocio digital."
        />

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {services.map((service, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 30 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 0.5, delay: i * 0.1 }}
            >
              <Card className="h-full">
                <div className="p-2">
                  <div className="w-12 h-12 rounded-xl bg-accent/10 flex items-center justify-center mb-4">
                    <service.icon size={24} className="text-accent" />
                  </div>
                  <h3 className="text-lg font-semibold text-foreground mb-2">
                    {service.title}
                  </h3>
                  <p className="text-sm text-muted-foreground leading-relaxed">
                    {service.description}
                  </p>
                </div>
              </Card>
            </motion.div>
          ))}
        </div>
      </div>
    </section>
  );
}
