import type { Metadata } from "next";
import { HabilidadesContent } from "./content";

export const metadata: Metadata = {
  title: "Habilidades",
  description: "Habilidades técnicas de Duvan López en programación, automatización, infraestructura, datos y ciberseguridad.",
};

export default function HabilidadesPage() {
  return <HabilidadesContent />;
}
