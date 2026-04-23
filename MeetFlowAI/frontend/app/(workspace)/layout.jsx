"use client";

import { motion } from "framer-motion";

import { AppShell } from "@/components/meetflow/app-shell";
import { AuthGate } from "@/components/meetflow/auth-gate";
import { MeetFlowProvider } from "@/components/meetflow/app-context";

export default function WorkspaceLayout({ children }) {
  return (
    <MeetFlowProvider>
      <AuthGate>
        <AppShell>
          <motion.div
            initial={{ opacity: 0, y: 8 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.25, ease: "easeOut" }}
            className="h-full xl:mx-auto xl:w-full xl:max-w-[1120px]"
          >
            {children}
          </motion.div>
        </AppShell>
      </AuthGate>
    </MeetFlowProvider>
  );
}
