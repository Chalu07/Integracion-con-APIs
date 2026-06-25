import type { Metadata } from "next";
import Link from "next/link";
import { ArrowLeft } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";

type Props = {
  params: Promise<{ slug: string }>;
};

export async function generateMetadata({ params }: Props): Promise<Metadata> {
  const { slug } = await params;
  const title = slug.replace(/-/g, " ").replace(/\b\w/g, (l) => l.toUpperCase());
  return {
    title,
    description: `Artículo: ${title}`,
  };
}

export default async function BlogPostPage({ params }: Props) {
  const { slug } = await params;
  const title = slug.replace(/-/g, " ").replace(/\b\w/g, (l) => l.toUpperCase());

  return (
    <article className="py-24 px-6">
      <div className="mx-auto max-w-3xl">
        <Link href="/blog">
          <Button variant="ghost" size="sm" className="mb-8">
            <ArrowLeft size={16} /> Volver al blog
          </Button>
        </Link>

        <div className="mb-8">
          <div className="flex flex-wrap gap-2 mb-4">
            <Badge>Artículo</Badge>
          </div>
          <h1 className="text-3xl md:text-4xl font-bold text-foreground mb-4">
            {title}
          </h1>
          <p className="text-muted-foreground">
            Publicado por Duvan López
          </p>
        </div>

        <div className="prose prose-invert max-w-none">
          <p className="text-muted-foreground text-lg leading-relaxed">
            Este artículo estará disponible próximamente. El contenido será gestionado desde el panel administrativo CMS.
          </p>
          <p className="text-muted-foreground leading-relaxed mt-4">
            Mientras tanto, puedes explorar los proyectos y experiencia profesional de Duvan López
            en las otras secciones del portafolio.
          </p>
        </div>
      </div>
    </article>
  );
}
