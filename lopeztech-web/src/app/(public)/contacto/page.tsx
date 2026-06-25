import type { Metadata } from "next";
import { ContactForm } from "@/components/sections/contact-form";
import { SectionHeader } from "@/components/ui/section-header";
import { Mail, MapPin, Globe } from "lucide-react";

export const metadata: Metadata = {
  title: "Contacto",
  description: "Contacta a Duvan López para proyectos de desarrollo de software, automatización o ciberseguridad.",
};

const contactInfo = [
  { icon: Mail, label: "Email", value: "rendonfredy31@gmail.com", href: "mailto:rendonfredy31@gmail.com" },
  { icon: MapPin, label: "Ubicación", value: "Colombia" },
  { icon: Globe, label: "Web", value: "lopeztech.dev", href: "https://lopeztech.dev" },
];

export default function ContactoPage() {
  return (
    <section className="py-24 px-6">
      <div className="mx-auto max-w-6xl">
        <SectionHeader
          kicker="Contacto"
          title="Hablemos"
          description="¿Tienes un proyecto en mente? Envíame un mensaje y conversemos sobre cómo puedo ayudarte."
        />

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-12">
          <div className="lg:col-span-2">
            <ContactForm />
          </div>

          <div className="space-y-6">
            {contactInfo.map((info, i) => (
              <div key={i} className="flex items-start gap-4 p-4 rounded-xl border border-border bg-card/30">
                <div className="w-10 h-10 rounded-lg bg-accent/10 flex items-center justify-center shrink-0">
                  <info.icon size={18} className="text-accent" />
                </div>
                <div>
                  <p className="text-xs text-muted-foreground mb-1">{info.label}</p>
                  {info.href ? (
                    <a
                      href={info.href}
                      className="text-sm text-foreground hover:text-accent transition-colors"
                      target={info.href.startsWith("http") ? "_blank" : undefined}
                      rel={info.href.startsWith("http") ? "noopener noreferrer" : undefined}
                    >
                      {info.value}
                    </a>
                  ) : (
                    <p className="text-sm text-foreground">{info.value}</p>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}
