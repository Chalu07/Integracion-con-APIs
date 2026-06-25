"use client";

import { motion } from "framer-motion";
import Link from "next/link";
import { Button } from "@/components/ui/button";
import { ArrowRight } from "lucide-react";

export function CTA() {
  return (
    <section className="py-24 px-6">
      <div className="mx-auto max-w-4xl">
        <motion.div
          initial={{ opacity: 0, y: 30 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          transition={{ duration: 0.6 }}
          className="relative rounded-3xl border border-accent/20 bg-gradient-to-br from-accent/5 via-card/50 to-card/50 p-12 md:p-16 text-center overflow-hidden"
        >
          <div className="absolute top-0 right-0 w-64 h-64 bg-accent/10 rounded-full blur-3xl" />
          <div className="relative">
            <h2 className="text-3xl md:text-4xl font-bold text-foreground mb-4">
              ¿Tienes un proyecto en mente?
            </h2>
            <p className="text-muted-foreground text-lg mb-8 max-w-xl mx-auto">
              Conversemos sobre cómo puedo ayudarte a construir, automatizar o proteger tus soluciones digitales.
            </p>
            <Link href="/contacto">
              <Button size="lg">
                Contactar <ArrowRight size={18} />
              </Button>
            </Link>
          </div>
        </motion.div>
      </div>
    </section>
  );
}
