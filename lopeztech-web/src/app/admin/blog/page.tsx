"use client";

import { CrudPage } from "@/components/admin/crud-page";
import { SimpleForm } from "@/components/admin/simple-form";

const fields = [
  { name: "title", label: "Título", required: true },
  { name: "titleEn", label: "Título (EN)" },
  { name: "slug", label: "Slug", required: true },
  { name: "excerpt", label: "Extracto", type: "textarea" as const },
  { name: "content", label: "Contenido", type: "textarea" as const, required: true },
  { name: "contentEn", label: "Contenido (EN)", type: "textarea" as const },
  { name: "coverImage", label: "Imagen de portada URL" },
  { name: "seoTitle", label: "Título SEO" },
  { name: "seoDescription", label: "Descripción SEO" },
  { name: "published", label: "Publicado", type: "checkbox" as const },
];

export default function AdminBlogPage() {
  return (
    <CrudPage
      title="Blog"
      apiPath="/api/admin/blog"
      columns={[
        { key: "title", label: "Título" },
        { key: "slug", label: "Slug" },
        { key: "published", label: "Publicado" },
      ]}
      renderForm={(item, onSave, onCancel) => (
        <SimpleForm fields={fields} initialData={item} onSave={onSave} onCancel={onCancel} />
      )}
    />
  );
}
