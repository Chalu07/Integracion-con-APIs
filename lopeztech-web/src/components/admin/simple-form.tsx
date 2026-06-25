"use client";

import { useMemo } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { useForm } from "react-hook-form";

interface FieldConfig {
  name: string;
  label: string;
  type?: "text" | "textarea" | "number" | "checkbox" | "url" | "email";
  required?: boolean;
  placeholder?: string;
}

interface SimpleFormProps {
  fields: FieldConfig[];
  initialData: Record<string, unknown> | null;
  onSave: (data: Record<string, unknown>) => void;
  onCancel: () => void;
}

export function SimpleForm({ fields, initialData, onSave, onCancel }: SimpleFormProps) {
  const defaultValues = useMemo(() => {
    if (initialData) return initialData;
    const defaults: Record<string, unknown> = {};
    fields.forEach((f) => {
      if (f.type === "number") defaults[f.name] = 0;
      else if (f.type === "checkbox") defaults[f.name] = false;
      else defaults[f.name] = "";
    });
    return defaults;
  }, [initialData, fields]);

  const { register, handleSubmit } = useForm({
    defaultValues: defaultValues as Record<string, string | number | boolean>,
  });

  const onSubmitHandler = (formData: Record<string, string | number | boolean>) => {
    const cleaned: Record<string, unknown> = {};
    fields.forEach((f) => {
      const val = formData[f.name];
      if (f.type === "number") cleaned[f.name] = Number(val) || 0;
      else if (f.type === "checkbox") cleaned[f.name] = Boolean(val);
      else cleaned[f.name] = val || (f.required ? "" : undefined);
    });
    onSave(cleaned);
  };

  return (
    <form onSubmit={handleSubmit(onSubmitHandler)} className="space-y-4">
      {fields.map((field) => {
        if (field.type === "checkbox") {
          return (
            <label key={field.name} className="flex items-center gap-2 text-sm">
              <input
                type="checkbox"
                {...register(field.name)}
                className="accent-accent"
              />
              {field.label}
            </label>
          );
        }

        if (field.type === "textarea") {
          return (
            <Textarea
              key={field.name}
              id={field.name}
              label={field.label}
              placeholder={field.placeholder}
              {...register(field.name)}
              required={field.required}
            />
          );
        }

        return (
          <Input
            key={field.name}
            id={field.name}
            label={field.label}
            type={field.type || "text"}
            placeholder={field.placeholder}
            {...register(field.name)}
            required={field.required}
          />
        );
      })}

      <div className="flex gap-3 pt-2">
        <Button type="submit">Guardar</Button>
        <Button type="button" variant="secondary" onClick={onCancel}>
          Cancelar
        </Button>
      </div>
    </form>
  );
}
