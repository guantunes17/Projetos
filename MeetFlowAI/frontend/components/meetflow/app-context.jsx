"use client";

import { createContext, useContext, useEffect, useMemo, useState } from "react";

import { api } from "@/lib/api";

const MeetFlowContext = createContext(null);

export function MeetFlowProvider({ children }) {
  const [token, setToken] = useState("");
  const [userName, setUserName] = useState("MeetFlow Owner");
  const [dashboard, setDashboard] = useState({ meetings_count: 0, tasks_count: 0, decisions_count: 0 });
  const [meetings, setMeetings] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [syncError, setSyncError] = useState("");

  useEffect(() => {
    const savedToken = localStorage.getItem("meetflow.token");
    const savedUser = localStorage.getItem("meetflow.userName");
    if (savedToken) {
      setToken(savedToken);
      setUserName(savedUser || "MeetFlow Owner");
      refreshAll(savedToken).catch((e) => {
        if (e?.status === 401) {
          localStorage.removeItem("meetflow.token");
          localStorage.removeItem("meetflow.userName");
          setToken("");
        }
        // Rede / backend fora: mantém sessão e mostra aviso; não apaga histórico em cache
      });
    }
  }, []);

  async function refreshAll(authToken = token) {
    if (!authToken) return;
    setIsLoading(true);
    setSyncError("");
    try {
      const [dash, list] = await Promise.all([api.dashboard(authToken), api.meetings(authToken)]);
      setDashboard(dash);
      setMeetings(list);
    } catch (e) {
      if (e?.status === 401) {
        setToken("");
        setMeetings([]);
        setDashboard({ meetings_count: 0, tasks_count: 0, decisions_count: 0 });
        localStorage.removeItem("meetflow.token");
        localStorage.removeItem("meetflow.userName");
        return;
      }
      setSyncError(
        e?.message ||
          "Não foi possível sincronizar com o servidor. Os dados apresentados podem estar desatualizados; tente novamente.",
      );
      // Nunca limpar meetings/dashboard em falha de rede: evita perder o histórico por ECONNRESET / proxy.
    } finally {
      setIsLoading(false);
    }
  }

  async function login(email, password) {
    const result = await api.login(email, password);
    setToken(result.access_token);
    setUserName(result.user_name || "MeetFlow Owner");
    localStorage.setItem("meetflow.token", result.access_token);
    localStorage.setItem("meetflow.userName", result.user_name || "MeetFlow Owner");
    await refreshAll(result.access_token);
  }

  function logout() {
    setToken("");
    setMeetings([]);
    setDashboard({ meetings_count: 0, tasks_count: 0, decisions_count: 0 });
    localStorage.removeItem("meetflow.token");
    localStorage.removeItem("meetflow.userName");
  }

  const value = useMemo(
    () => ({
      token,
      userName,
      dashboard,
      meetings,
      isLoading,
      syncError,
      clearSyncError: () => setSyncError(""),
      login,
      logout,
      refreshAll,
    }),
    [token, userName, dashboard, meetings, isLoading, syncError],
  );

  return <MeetFlowContext.Provider value={value}>{children}</MeetFlowContext.Provider>;
}

export function useMeetFlow() {
  const context = useContext(MeetFlowContext);
  if (!context) throw new Error("useMeetFlow must be used inside MeetFlowProvider");
  return context;
}
