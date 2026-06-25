import Link from "next/link";
import { Code2, UserRound, Mail } from "lucide-react";

const footerLinks = [
  { href: "/sobre-mi", label: "Sobre mí" },
  { href: "/proyectos", label: "Proyectos" },
  { href: "/blog", label: "Blog" },
  { href: "/contacto", label: "Contacto" },
];

const socialLinks = [
  { href: "https://github.com/DuvanLope", icon: Code2, label: "GitHub" },
  { href: "https://linkedin.com/in/duvanlopez", icon: UserRound, label: "LinkedIn" },
  { href: "mailto:rendonfredy31@gmail.com", icon: Mail, label: "Email" },
];

export function Footer() {
  return (
    <footer className="border-t border-border bg-background/50">
      <div className="mx-auto max-w-6xl px-6 py-12">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          <div>
            <Link href="/" className="text-xl font-bold">
              <span className="text-accent">Lopez</span>Tech
            </Link>
            <p className="mt-3 text-sm text-muted-foreground max-w-xs">
              Desarrollador de Software, QA Engineer y especialista en automatización, APIs y ciberseguridad.
            </p>
          </div>

          <div>
            <h3 className="font-semibold text-foreground mb-4">Navegación</h3>
            <ul className="space-y-2">
              {footerLinks.map((link) => (
                <li key={link.href}>
                  <Link
                    href={link.href}
                    className="text-sm text-muted-foreground hover:text-accent transition-colors"
                  >
                    {link.label}
                  </Link>
                </li>
              ))}
            </ul>
          </div>

          <div>
            <h3 className="font-semibold text-foreground mb-4">Conectar</h3>
            <div className="flex gap-3">
              {socialLinks.map((social) => (
                <a
                  key={social.href}
                  href={social.href}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="p-2.5 rounded-full border border-border text-muted-foreground hover:text-accent hover:border-accent/40 transition-all"
                  aria-label={social.label}
                >
                  <social.icon size={18} />
                </a>
              ))}
            </div>
          </div>
        </div>

        <div className="mt-10 pt-6 border-t border-border flex flex-col sm:flex-row justify-between items-center gap-4">
          <p className="text-xs text-muted-foreground">
            &copy; {new Date().getFullYear()} Duvan López. Portafolio profesional.
          </p>
          <p className="text-xs text-muted-foreground">
            Construido con Next.js, TypeScript y Tailwind CSS
          </p>
        </div>
      </div>
    </footer>
  );
}
