import { cva } from "class-variance-authority";

import { cn } from "@/lib/utils";

const buttonVariants = cva(
  "inline-flex items-center justify-center rounded-xl text-sm font-medium transition-all duration-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-400 disabled:pointer-events-none disabled:opacity-50 motion-safe:hover:-translate-y-0.5 active:translate-y-0",
  {
    variants: {
      variant: {
        default: "bg-blue-500 text-white shadow-sm shadow-blue-900/30 hover:bg-blue-400 hover:shadow-blue-700/30",
        secondary: "bg-slate-800 text-slate-100 hover:bg-slate-700 hover:shadow-lg hover:shadow-slate-950/30",
        ghost: "bg-transparent text-slate-300 hover:bg-slate-800 hover:text-slate-100",
        success: "bg-lime-500 text-slate-950 shadow-sm shadow-lime-900/20 hover:bg-lime-400 hover:shadow-lime-700/20",
      },
      size: {
        default: "h-10 px-4 py-2",
        sm: "h-8 px-3",
        lg: "h-11 px-6",
      },
    },
    defaultVariants: {
      variant: "default",
      size: "default",
    },
  },
);

export function Button({ className, variant, size, ...props }) {
  return <button className={cn(buttonVariants({ variant, size }), className)} {...props} />;
}
