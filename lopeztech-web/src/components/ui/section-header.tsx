"use client";

import { motion } from "framer-motion";

interface SectionHeaderProps {
  kicker: string;
  title: string;
  description?: string;
}

export function SectionHeader({ kicker, title, description }: SectionHeaderProps) {
  return (
    <motion.div
      initial={{ opacity: 0, y: 30 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true, margin: "-100px" }}
      transition={{ duration: 0.6 }}
      className="mb-16"
    >
      <p className="text-accent font-mono text-sm tracking-wider uppercase mb-3">
        {kicker}
      </p>
      <h2 className="text-3xl md:text-4xl lg:text-5xl font-bold text-foreground leading-tight">
        {title}
      </h2>
      {description && (
        <p className="mt-4 text-lg text-muted-foreground max-w-2xl">
          {description}
        </p>
      )}
    </motion.div>
  );
}
