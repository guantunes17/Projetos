import { cva } from "class-variance-authority";

import { cn } from "@/lib/utils";

const badgeVariants = cva("inline-flex items-center rounded-full px-2.5 py-1 text-xs font-medium", {
  variants: {
    variant: {
      default: "bg-slate-800 text-slate-300",
      success: "bg-emerald-500/20 text-emerald-300",
      warning: "bg-amber-500/20 text-amber-300",
      info: "bg-indigo-500/20 text-indigo-300",
    },
  },
  defaultVariants: {
    variant: "default",
  },
});

export function Badge({ className, variant, ...props }) {
  return <span className={cn(badgeVariants({ variant }), className)} {...props} />;
}
