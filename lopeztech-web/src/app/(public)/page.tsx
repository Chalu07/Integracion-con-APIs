import { Hero } from "@/components/sections/hero";
import { ServicesPreview } from "@/components/sections/services-preview";
import { ProjectsPreview } from "@/components/sections/projects-preview";
import { CTA } from "@/components/sections/cta";

export default function HomePage() {
  return (
    <>
      <Hero />
      <ServicesPreview />
      <ProjectsPreview />
      <CTA />
    </>
  );
}
