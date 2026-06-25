"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "title", label: "Título", required: true, placeholder: "Desarrollo de Software" },
  { name: "titleEn", label: "Título (EN)", placeholder: "Software Development" },
  { name: "description", label: "Descripción", type: "textarea" as const, required: true },
  { name: "descriptionEn", label: "Descripción (EN)", type: "textarea" as const },
  { name: "icon", label: "Icono", placeholder: "Code, Zap, Shield..." },
  { name: "slug", label: "Slug", required: true, placeholder: "desarrollo-software" },
  { name: "order", label: "Orden", type: "number" as const },
  { name: "published", label: "Publicado", type: "checkbox" as const },
];

export default function AdminServiciosPage() {
  return (
    <CrudPage
      title="Servicios"
      apiPath="/api/admin/services"
      columns={[
        { key: "title", label: "Título" },
        { key: "slug", label: "Slug" },
      ]}
      renderForm={(item, onSave, onCancel) => (
        <SimpleForm fields={fields} initialData={item} onSave={onSave} onCancel={onCancel} />
      )}
    />
  );
}
