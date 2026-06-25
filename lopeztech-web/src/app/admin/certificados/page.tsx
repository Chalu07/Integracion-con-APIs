"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "title", label: "Título", required: true },
  { name: "titleEn", label: "Título (EN)" },
  { name: "description", label: "Descripción", type: "textarea" as const, required: true },
  { name: "descriptionEn", label: "Descripción (EN)", type: "textarea" as const },
  { name: "institution", label: "Institución", required: true },
  { name: "date", label: "Fecha", placeholder: "Ene 2025" },
  { name: "url", label: "URL del certificado" },
  { name: "order", label: "Orden", type: "number" as const },
  { name: "published", label: "Publicado", type: "checkbox" as const },
];

export default function AdminCertificadosPage() {
  return (
    <CrudPage
      title="Certificados"
      apiPath="/api/admin/certificates"
      columns={[
        { key: "title", label: "Título" },
        { key: "institution", label: "Institución" },
      ]}
      renderForm={(item, onSave, onCancel) => (
        <SimpleForm fields={fields} initialData={item} onSave={onSave} onCancel={onCancel} />
      )}
    />
  );
}
