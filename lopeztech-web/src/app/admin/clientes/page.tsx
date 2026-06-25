"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "name", label: "Nombre", required: true },
  { name: "logo", label: "Logo URL" },
  { name: "url", label: "URL", type: "url" as const },
  { name: "order", label: "Orden", type: "number" as const },
  { name: "published", label: "Publicado", type: "checkbox" as const },
];

export default function AdminClientesPage() {
  return (
    <CrudPage
      title="Clientes"
      apiPath="/api/admin/clients"
      columns={[
        { key: "name", label: "Nombre" },
        { key: "url", label: "URL" },
      ]}
      renderForm={(item, onSave, onCancel) => (
        <SimpleForm fields={fields} initialData={item} onSave={onSave} onCancel={onCancel} />
      )}
    />
  );
}
