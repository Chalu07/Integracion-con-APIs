import { cn } from "@/lib/utils";
import type { HTMLAttributes } from "react";

interface BadgeProps extends HTMLAttributes<HTMLSpanElement> {
  variant?: "default" | "accent" | "muted";
}

export function Badge({ className, variant = "default", children, ...props }: BadgeProps) {
  return (
    <span
      className={cn(
        "inline-flex items-center rounded-full px-3 py-1 text-xs font-medium transition-colors",
        {
          "bg-accent/10 text-accent border border-accent/20": variant === "default",
          "bg-accent text-background": variant === "accent",
          "bg-muted-foreground/10 text-muted-foreground": variant === "muted",
        },
        className
      )}
      {...props}
    >
      {children}
    </span>
  );
}
