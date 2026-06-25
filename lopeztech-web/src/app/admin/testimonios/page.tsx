"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "name", label: "Nombre", required: true },
  { name: "role", label: "Cargo" },
  { name: "company", label: "Empresa" },
  { name: "content", label: "Contenido", type: "textarea" as const, required: true },
  { name: "contentEn", label: "Contenido (EN)", type: "textarea" as const },
  { name: "avatar", label: "Avatar URL" },
  { name: "rating", label: "Calificación (1-5)", type: "number" as const },
  { name: "order", label: "Orden", type: "number" as const },
  { name: "published", label: "Publicado", type: "checkbox" as const },
];

export default function AdminTestimoniosPage() {
  return (
    <CrudPage
      title="Testimonios"
      apiPath="/api/admin/testimonials"
      columns={[
        { key: "name", label: "Nombre" },
        { key: "company", label: "Empresa" },
      ]}
      renderForm={(item, onSave, onCancel) => (
        <SimpleForm fields={fields} initialData={item} onSave={onSave} onCancel={onCancel} />
      )}
    />
  );
}
