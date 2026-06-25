"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "title", label: "Título del cargo", required: true },
  { name: "titleEn", label: "Título (EN)" },
  { name: "company", label: "Empresa", required: true },
  { name: "description", label: "Descripción", type: "textarea" as const, required: true },
  { name: "descriptionEn", label: "Descripción (EN)", type: "textarea" as const },
  { name: "startDate", label: "Fecha inicio", required: true, placeholder: "Nov 2025" },
  { name: "endDate", label: "Fecha fin", placeholder: "Jun 2026 o vacío si actual" },
  { name: "type", label: "Tipo", placeholder: "laboral o freelance" },
  { name: "order", label: "Orden", type: "number" as const },
  { name: "published", label: "Publicado", type: "checkbox" as const },
];

export default function AdminExperienciaPage() {
  return (
    <CrudPage
      title="Experiencia"
      apiPath="/api/admin/experience"
      columns={[
        { key: "title", label: "Cargo" },
        { key: "company", label: "Empresa" },
      ]}
      renderForm={(item, onSave, onCancel) => (
        <SimpleForm fields={fields} initialData={item} onSave={onSave} onCancel={onCancel} />
      )}
    />
  );
}
