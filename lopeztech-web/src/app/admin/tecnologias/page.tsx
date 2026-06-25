"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "name", label: "Nombre", required: true, placeholder: "React" },
  { name: "icon", label: "Icono URL" },
  { name: "url", label: "URL", type: "url" as const },
  { name: "category", label: "Categoría", required: true, placeholder: "Frameworks" },
  { name: "order", label: "Orden", type: "number" as const },
];

export default function AdminTecnologiasPage() {
  return (
    <CrudPage
      title="Tecnologías"
      apiPath="/api/admin/technologies"
      columns={[
        { key: "name", label: "Nombre" },
        { key: "category", label: "Categoría" },
      ]}
      renderForm={(item, onSave, onCancel) => (
        <SimpleForm fields={fields} initialData={item} onSave={onSave} onCancel={onCancel} />
      )}
    />
  );
}
