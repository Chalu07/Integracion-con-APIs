import type { Metadata } from "next";
import { ProyectosContent } from "./content";

export const metadata: Metadata = {
  title: "Proyectos",
  description: "Proyectos destacados de automatización, desarrollo web, onboarding, Active Directory, APIs y GitHub de Duvan López.",
};

export default function ProyectosPage() {
  return <ProyectosContent />;
}
