"use client";

import { motion } from "framer-motion";
import { SectionHeader } from "@/components/ui/section-header";

const categories = [
  {
    num: "01",
    name: "Lenguajes de Programación",
    techs: ["Python", "Java", "C#", "JavaScript", "HTML", "CSS", "Dart"],
  },
  {
    num: "02",
    name: "Frameworks & Librerías",
    techs: ["React", "Next.js", "Flutter", "Tailwind CSS", "Node.js", "Express", ".NET"],
  },
  {
    num: "03",
    name: "Automatización & No-Code",
    techs: ["N8N", "Power Automate", "Power Apps", "Webhooks", "APIs REST"],
  },
  {
    num: "04",
    name: "Infraestructura & Cloud",
    techs: ["SharePoint", "Active Directory", "Azure AD", "Microsoft 365", "Git", "GitHub"],
  },
  {
    num: "05",
    name: "Bases de Datos",
    techs: ["MySQL", "PostgreSQL", "MongoDB", "SQL Server", "Firebase"],
  },
  {
    num: "06",
    name: "QA & Testing",
    techs: ["Selenium", "Pytest", "Postman", "JUnit", "Pruebas E2E", "Pruebas de Vulnerabilidad"],
  },
  {
    num: "07",
    name: "Ciberseguridad",
    techs: ["OWASP", "SIEM", "Análisis de Vulnerabilidades", "Pentesting", "Nmap", "Burp Suite"],
  },
  {
    num: "08",
    name: "Datos & BI",
    techs: ["Power BI", "Google Sheets API", "ETL", "Big Data", "Análisis de Datos"],
  },
];

export function HabilidadesContent() {
  return (
    <section className="py-24 px-6">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Habilidades"
          title="Herramientas para construir, automatizar y proteger."
          description="Stack tecnológico construido a partir de formación en Desarrollo de Software, experiencia como analista de automatizaciones y proyectos freelance."
        />

        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          {categories.map((cat, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 30 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.5, delay: i * 0.05 }}
              className="rounded-2xl border border-border bg-card/30 p-6 hover:border-accent/20 transition-all duration-300"
            >
              <div className="flex items-center gap-3 mb-5">
                <span className="text-xs font-mono text-accent bg-accent/10 px-2.5 py-1 rounded-full">
                  {cat.num}
                </span>
                <h3 className="font-semibold text-foreground">{cat.name}</h3>
              </div>
              <div className="flex flex-wrap gap-2">
                {cat.techs.map((tech) => (
                  <span
                    key={tech}
                    className="px-3 py-1.5 text-xs rounded-lg bg-background border border-border text-muted-foreground hover:text-accent hover:border-accent/30 transition-all cursor-default"
                  >
                    {tech}
                  </span>
                ))}
              </div>
            </motion.div>
          ))}
        </div>
      </div>
    </section>
  );
}
