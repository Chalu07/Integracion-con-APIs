"use client";

import { motion } from "framer-motion";
import { ArrowDown, Code, Zap, Shield } from "lucide-react";
import { Button } from "@/components/ui/button";
import Link from "next/link";

const roles = [
  { icon: Code, label: "Desarrollador de Software" },
  { icon: Zap, label: "Automatizaciones" },
  { icon: Shield, label: "Ciberseguridad" },
];

export function Hero() {
  return (
    <section
      className="relative min-h-[90vh] flex items-center justify-center overflow-hidden"
      aria-labelledby="hero-title"
    >
      <div className="absolute inset-0 bg-gradient-to-b from-accent/5 via-transparent to-transparent" />
      <div className="absolute top-1/4 -right-1/4 w-[600px] h-[600px] rounded-full bg-accent/5 blur-3xl" />
      <div className="absolute bottom-1/4 -left-1/4 w-[400px] h-[400px] rounded-full bg-blue-500/5 blur-3xl" />

      <div className="relative mx-auto max-w-6xl px-6 py-20 text-center">
        <motion.p
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5 }}
          className="text-accent font-mono text-sm tracking-widest uppercase mb-6"
        >
          &middot; Mi Portafolio &middot;
        </motion.p>

        <motion.h1
          id="hero-title"
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 0.1 }}
          className="text-5xl sm:text-6xl md:text-7xl lg:text-8xl font-bold tracking-tight"
        >
          DUVAN
          <br />
          <span className="text-accent">LÓPEZ</span>
        </motion.h1>

        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ duration: 0.5, delay: 0.3 }}
          className="w-16 h-px bg-accent/50 mx-auto my-8"
        />

        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5, delay: 0.4 }}
          className="flex flex-wrap items-center justify-center gap-4 md:gap-6 mb-10"
        >
          {roles.map((role, i) => (
            <div key={i} className="flex items-center gap-2 text-muted-foreground">
              <role.icon size={18} className="text-accent" />
              <span className="text-sm md:text-base">{role.label}</span>
              {i < roles.length - 1 && (
                <span className="hidden md:inline text-border ml-4">&middot;</span>
              )}
            </div>
          ))}
        </motion.div>

        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5, delay: 0.5 }}
          className="flex flex-wrap items-center justify-center gap-4"
        >
          <Link href="/contacto">
            <Button size="lg">Contactar</Button>
          </Link>
          <Link href="/proyectos">
            <Button variant="secondary" size="lg">Ver proyectos</Button>
          </Link>
        </motion.div>

        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ duration: 0.5, delay: 0.8 }}
          className="mt-16"
        >
          <a
            href="#about-preview"
            className="inline-flex items-center gap-2 text-muted-foreground hover:text-accent transition-colors text-sm"
            aria-label="Desplazar hacia abajo"
          >
            <ArrowDown size={16} className="animate-bounce" />
            Explorar
          </a>
        </motion.div>
      </div>
    </section>
  );
}
