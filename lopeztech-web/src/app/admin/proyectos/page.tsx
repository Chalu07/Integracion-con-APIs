"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "title", label: "Título", required: true },
  { name: "titleEn", label: "Título (EN)" },
  { name: "description", label: "Descripción", type: "textarea" as const, required: true },
  { name: "descriptionEn", label: "Descripción (EN)", type: "textarea" as const },
  { name: "image", label: "Imagen URL" },
  { name: "codeUrl", label: "URL del Código", type: "url" as const },
  { name: "demoUrl", label: "URL Demo", type: "url" as const },
  { name: "result", label: "Resultado" },
  { name: "slug", label: "Slug", required: true },
  { name: "order", label: "Orden", type: "number" as const },
  { name: "featured", label: "Destacado", type: "checkbox" as const },
  { name: "published", label: "Publicado", type: "checkbox" as const },
];

export default function AdminProyectosPage() {
  return (
    <CrudPage
      title="Proyectos"
      apiPath="/api/admin/projects"
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
