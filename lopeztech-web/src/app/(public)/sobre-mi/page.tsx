import type { Metadata } from "next";
import { SobreMiContent } from "./content";

export const metadata: Metadata = {
  title: "Sobre mí",
  description: "Conoce a Duvan López — Desarrollador de Software, QA Engineer y especialista en automatización, APIs, SharePoint y ciberseguridad.",
};

export default function SobreMiPage() {
  return <SobreMiContent />;
}
