import type { Metadata } from "next";
import { Plus_Jakarta_Sans } from "next/font/google";
import { Providers } from "@/components/layout/providers";
import "./globals.css";

const jakarta = Plus_Jakarta_Sans({
  subsets: ["latin"],
  variable: "--font-sans",
});

export const metadata: Metadata = {
  metadataBase: new URL(process.env.NEXT_PUBLIC_SITE_URL || "https://lopeztech.dev"),
  title: {
    default: "Duvan López — Desarrollador de Software & Automatizaciones",
    template: "%s — LopezTech",
  },
  description:
    "Portafolio profesional de Duvan López — Desarrollador de Software, QA Engineer y especialista en automatización, APIs, SharePoint y ciberseguridad.",
  keywords: [
    "Duvan López",
    "Desarrollador de Software",
    "QA Engineer",
    "Automatizaciones",
    "n8n",
    "Power Automate",
    "APIs REST",
    "SharePoint",
    "Ciberseguridad",
    "React",
    "Next.js",
  ],
  authors: [{ name: "Duvan López" }],
  creator: "Duvan López",
  openGraph: {
    type: "website",
    locale: "es_CO",
    url: "/",
    title: "Duvan López — Portafolio Profesional",
    description:
      "Desarrollador de Software y QA Engineer especializado en automatizaciones n8n, APIs, SharePoint, Active Directory y ciberseguridad.",
    siteName: "LopezTech",
  },
  twitter: {
    card: "summary_large_image",
    title: "Duvan López — Portafolio Profesional",
    description:
      "Desarrollador de Software y QA Engineer especializado en automatizaciones n8n, APIs, SharePoint, Active Directory y ciberseguridad.",
  },
  robots: {
    index: true,
    follow: true,
    googleBot: { index: true, follow: true },
  },
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="es" suppressHydrationWarning>
      <body className={`${jakarta.variable} font-sans antialiased`}>
        <Providers>{children}</Providers>
      </body>
    </html>
  );
}
