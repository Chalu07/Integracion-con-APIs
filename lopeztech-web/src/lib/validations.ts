import { z } from "zod";

export const contactSchema = z.object({
  name: z.string().min(2, "El nombre debe tener al menos 2 caracteres").max(100),
  email: z.string().email("Email inválido"),
  subject: z.string().min(3, "El asunto debe tener al menos 3 caracteres").max(200).optional(),
  message: z.string().min(10, "El mensaje debe tener al menos 10 caracteres").max(5000),
});

export const serviceSchema = z.object({
  title: z.string().min(2).max(200),
  titleEn: z.string().max(200).optional(),
  description: z.string().min(10),
  descriptionEn: z.string().optional(),
  icon: z.string().optional(),
  slug: z.string().min(2),
  order: z.number().int().default(0),
  published: z.boolean().default(true),
});

export const projectSchema = z.object({
  title: z.string().min(2).max(200),
  titleEn: z.string().max(200).optional(),
  description: z.string().min(10),
  descriptionEn: z.string().optional(),
  image: z.string().optional(),
  tags: z.array(z.string()).default([]),
  demoUrl: z.string().url().optional().or(z.literal("")),
  codeUrl: z.string().url().optional().or(z.literal("")),
  result: z.string().optional(),
  resultEn: z.string().optional(),
  slug: z.string().min(2),
  featured: z.boolean().default(false),
  order: z.number().int().default(0),
  published: z.boolean().default(true),
});

export const blogPostSchema = z.object({
  title: z.string().min(2).max(300),
  titleEn: z.string().max(300).optional(),
  slug: z.string().min(2),
  excerpt: z.string().optional(),
  excerptEn: z.string().optional(),
  content: z.string().min(10),
  contentEn: z.string().optional(),
  coverImage: z.string().optional(),
  tags: z.array(z.string()).default([]),
  published: z.boolean().default(false),
  publishedAt: z.string().datetime().optional().nullable(),
  seoTitle: z.string().optional(),
  seoDescription: z.string().optional(),
});

export const technologySchema = z.object({
  name: z.string().min(1).max(100),
  icon: z.string().optional(),
  url: z.string().url().optional().or(z.literal("")),
  category: z.string().min(1),
  order: z.number().int().default(0),
});

export const testimonialSchema = z.object({
  name: z.string().min(2).max(100),
  role: z.string().optional(),
  company: z.string().optional(),
  content: z.string().min(10),
  contentEn: z.string().optional(),
  avatar: z.string().optional(),
  rating: z.number().int().min(1).max(5).default(5),
  published: z.boolean().default(true),
  order: z.number().int().default(0),
});

export const clientSchema = z.object({
  name: z.string().min(2).max(100),
  logo: z.string().optional(),
  url: z.string().url().optional().or(z.literal("")),
  order: z.number().int().default(0),
  published: z.boolean().default(true),
});

export const experienceSchema = z.object({
  title: z.string().min(2).max(200),
  titleEn: z.string().max(200).optional(),
  company: z.string().min(1),
  description: z.string().min(10),
  descriptionEn: z.string().optional(),
  startDate: z.string(),
  endDate: z.string().optional(),
  tags: z.array(z.string()).default([]),
  type: z.string().default("laboral"),
  order: z.number().int().default(0),
  published: z.boolean().default(true),
});

export const certificateSchema = z.object({
  title: z.string().min(2).max(200),
  titleEn: z.string().max(200).optional(),
  description: z.string().min(10),
  descriptionEn: z.string().optional(),
  institution: z.string().min(1),
  date: z.string().optional(),
  url: z.string().optional(),
  order: z.number().int().default(0),
  published: z.boolean().default(true),
});

export type ContactFormData = z.infer<typeof contactSchema>;
export type ServiceFormData = z.infer<typeof serviceSchema>;
export type ProjectFormData = z.infer<typeof projectSchema>;
export type BlogPostFormData = z.infer<typeof blogPostSchema>;
