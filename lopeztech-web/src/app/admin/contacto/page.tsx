"use client";

import { CrudPage } from "@/components/admin/crud-page";

export default function AdminContactoPage() {
  return (
    <CrudPage
      title="Solicitudes de Contacto"
      apiPath="/api/admin/contact"
      columns={[
        { key: "name", label: "Nombre" },
        { key: "email", label: "Email" },
        { key: "subject", label: "Asunto" },
        { key: "read", label: "Leído" },
      ]}
    />
  );
}
