"use client";

import { useState } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { motion } from "framer-motion";
import { Send, CheckCircle, AlertCircle } from "lucide-react";
import { contactSchema, type ContactFormData } from "@/lib/validations";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";

export function ContactForm() {
  const [status, setStatus] = useState<"idle" | "loading" | "success" | "error">("idle");

  const {
    register,
    handleSubmit,
    reset,
    formState: { errors },
  } = useForm<ContactFormData>({
    resolver: zodResolver(contactSchema),
  });

  const onSubmit = async (data: ContactFormData) => {
    setStatus("loading");
    try {
      const res = await fetch("/api/contact", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });

      if (!res.ok) throw new Error("Error al enviar");
      setStatus("success");
      reset();
    } catch {
      setStatus("error");
    }
  };

  return (
    <motion.form
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true }}
      transition={{ duration: 0.5 }}
      onSubmit={handleSubmit(onSubmit)}
      className="space-y-6"
      noValidate
    >
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
        <Input
          id="name"
          label="Nombre"
          placeholder="Tu nombre"
          error={errors.name?.message}
          {...register("name")}
        />
        <Input
          id="email"
          label="Email"
          type="email"
          placeholder="tu@email.com"
          error={errors.email?.message}
          {...register("email")}
        />
      </div>

      <Input
        id="subject"
        label="Asunto"
        placeholder="¿De qué se trata?"
        error={errors.subject?.message}
        {...register("subject")}
      />

      <Textarea
        id="message"
        label="Mensaje"
        placeholder="Cuéntame sobre tu proyecto..."
        rows={6}
        error={errors.message?.message}
        {...register("message")}
      />

      <div className="flex items-center gap-4">
        <Button type="submit" disabled={status === "loading"}>
          {status === "loading" ? (
            <>
              <span className="animate-spin inline-block w-4 h-4 border-2 border-current border-t-transparent rounded-full" />
              Enviando...
            </>
          ) : (
            <>
              <Send size={16} /> Enviar mensaje
            </>
          )}
        </Button>

        {status === "success" && (
          <motion.span
            initial={{ opacity: 0, x: -10 }}
            animate={{ opacity: 1, x: 0 }}
            className="flex items-center gap-2 text-sm text-green-500"
          >
            <CheckCircle size={16} /> Mensaje enviado correctamente
          </motion.span>
        )}

        {status === "error" && (
          <motion.span
            initial={{ opacity: 0, x: -10 }}
            animate={{ opacity: 1, x: 0 }}
            className="flex items-center gap-2 text-sm text-red-500"
          >
            <AlertCircle size={16} /> Error al enviar. Intenta de nuevo.
          </motion.span>
        )}
      </div>
    </motion.form>
  );
}
