import type { Metadata } from "next";
import { CertificadosContent } from "./content";

export const metadata: Metadata = {
  title: "Certificados",
  description: "Certificados, estudios y formación destacada de Duvan López en desarrollo de software, datos y ciberseguridad.",
};

export default function CertificadosPage() {
  return <CertificadosContent />;
}
