import type { Metadata } from "next";
import { ExperienciaContent } from "./content";

export const metadata: Metadata = {
  title: "Experiencia",
  description: "Experiencia laboral e independiente de Duvan López en automatizaciones, desarrollo web y soluciones corporativas.",
};

export default function ExperienciaPage() {
  return <ExperienciaContent />;
}
