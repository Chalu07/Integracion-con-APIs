import type { Metadata } from "next";
import Link from "next/link";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { SectionHeader } from "@/components/ui/section-header";

export const metadata: Metadata = {
  title: "Blog",
  description: "Artículos sobre desarrollo de software, automatizaciones, ciberseguridad y tecnología por Duvan López.",
};

const placeholderPosts = [
  {
    slug: "introduccion-automatizaciones-n8n",
    title: "Introducción a las automatizaciones con n8n",
    excerpt: "Aprende cómo n8n puede transformar los procesos operativos de tu empresa mediante flujos de automatización visual.",
    tags: ["Automatización", "n8n"],
    date: "2025-12-15",
  },
  {
    slug: "seguridad-apis-rest",
    title: "Buenas prácticas de seguridad en APIs REST",
    excerpt: "Guía completa sobre autenticación, validación, rate limiting y protección contra ataques comunes en APIs.",
    tags: ["Seguridad", "APIs"],
    date: "2025-11-20",
  },
  {
    slug: "sharepoint-active-directory",
    title: "Integración SharePoint + Active Directory con n8n",
    excerpt: "Cómo construir un pipeline de aprovisionamiento de usuarios automatizado usando SharePoint y AD.",
    tags: ["SharePoint", "AD", "n8n"],
    date: "2025-10-10",
  },
];

export default function BlogPage() {
  return (
    <section className="py-24 px-6">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Blog"
          title="Artículos y recursos"
          description="Contenido técnico sobre desarrollo, automatización y seguridad."
        />

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {placeholderPosts.map((post) => (
            <Link key={post.slug} href={`/blog/${post.slug}`}>
              <Card className="h-full group cursor-pointer">
                <div className="p-2">
                  <p className="text-xs text-muted-foreground mb-3">{post.date}</p>
                  <h3 className="font-semibold text-foreground group-hover:text-accent transition-colors mb-2">
                    {post.title}
                  </h3>
                  <p className="text-sm text-muted-foreground leading-relaxed mb-4">
                    {post.excerpt}
                  </p>
                  <div className="flex flex-wrap gap-2">
                    {post.tags.map((tag) => (
                      <Badge key={tag} variant="muted">{tag}</Badge>
                    ))}
                  </div>
                </div>
              </Card>
            </Link>
          ))}
        </div>
      </div>
    </section>
  );
}
