"use client";

import { useState, useEffect, useCallback } from "react";
import { useSession } from "next-auth/react";
import { useRouter } from "next/navigation";
import Link from "next/link";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { ArrowLeft, Plus, Pencil, Trash2, LayoutDashboard } from "lucide-react";

interface CrudPageProps {
  title: string;
  apiPath: string;
  columns: { key: string; label: string }[];
  renderForm?: (
    item: Record<string, unknown> | null,
    onSave: (data: Record<string, unknown>) => void,
    onCancel: () => void
  ) => React.ReactNode;
}

export function CrudPage({ title, apiPath, columns, renderForm }: CrudPageProps) {
  const { status } = useSession();
  const router = useRouter();
  const [items, setItems] = useState<Record<string, unknown>[]>([]);
  const [editing, setEditing] = useState<Record<string, unknown> | null>(null);
  const [creating, setCreating] = useState(false);
  const [loading, setLoading] = useState(true);

  const fetchItems = useCallback(async () => {
    try {
      const res = await fetch(apiPath);
      if (res.ok) {
        const data = await res.json();
        setItems(data);
      }
    } finally {
      setLoading(false);
    }
  }, [apiPath]);

  useEffect(() => {
    if (status === "unauthenticated") router.push("/admin/login");
    if (status === "authenticated") fetchItems();
  }, [status, router, fetchItems]);

  const handleSave = async (data: Record<string, unknown>) => {
    const isEdit = editing && editing.id;
    const res = await fetch(apiPath, {
      method: isEdit ? "PUT" : "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(isEdit ? { id: editing.id, ...data } : data),
    });

    if (res.ok) {
      setEditing(null);
      setCreating(false);
      fetchItems();
    }
  };

  const handleDelete = async (id: string) => {
    if (!confirm("¿Estás seguro de eliminar este elemento?")) return;
    const res = await fetch(`${apiPath}?id=${id}`, { method: "DELETE" });
    if (res.ok) fetchItems();
  };

  if (status === "loading" || loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-background">
        <div className="animate-spin w-8 h-8 border-2 border-accent border-t-transparent rounded-full" />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background">
      <header className="sticky top-0 z-50 border-b border-border bg-background/80 backdrop-blur-xl">
        <div className="mx-auto max-w-6xl flex items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <LayoutDashboard size={20} className="text-accent" />
            <h1 className="text-lg font-bold">
              <span className="text-accent">Lopez</span>Tech CMS
            </h1>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-6 py-10">
        <div className="flex items-center justify-between mb-8">
          <div className="flex items-center gap-4">
            <Link href="/admin">
              <Button variant="ghost" size="sm">
                <ArrowLeft size={16} /> Volver
              </Button>
            </Link>
            <h2 className="text-2xl font-bold text-foreground">{title}</h2>
          </div>
          {renderForm && !creating && !editing && (
            <Button onClick={() => setCreating(true)} size="sm">
              <Plus size={16} /> Nuevo
            </Button>
          )}
        </div>

        {(creating || editing) && renderForm && (
          <Card className="mb-8">
            <div className="p-4">
              <h3 className="font-semibold text-foreground mb-4">
                {editing ? "Editar" : "Crear"} {title.toLowerCase().replace(/s$/, "")}
              </h3>
              {renderForm(editing, handleSave, () => {
                setEditing(null);
                setCreating(false);
              })}
            </div>
          </Card>
        )}

        {items.length === 0 ? (
          <Card>
            <p className="text-center text-muted-foreground py-8">No hay elementos todavía.</p>
          </Card>
        ) : (
          <div className="space-y-3">
            {items.map((item) => (
              <Card key={item.id as string} hover={false} className="flex items-center justify-between gap-4">
                <div className="flex-1 min-w-0">
                  {columns.map((col) => (
                    <span key={col.key} className="text-sm text-foreground mr-4">
                      <span className="text-muted-foreground text-xs">{col.label}: </span>
                      {String(item[col.key] ?? "")}
                    </span>
                  ))}
                </div>
                <div className="flex gap-2 shrink-0">
                  {renderForm && (
                    <Button
                      variant="ghost"
                      size="sm"
                      onClick={() => { setEditing(item); setCreating(false); }}
                    >
                      <Pencil size={14} />
                    </Button>
                  )}
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => handleDelete(item.id as string)}
                    className="text-red-500 hover:text-red-600"
                  >
                    <Trash2 size={14} />
                  </Button>
                </div>
              </Card>
            ))}
          </div>
        )}
      </main>
    </div>
  );
}
