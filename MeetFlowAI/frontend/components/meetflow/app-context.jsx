"use client";

import { createContext, useContext, useEffect, useMemo, useState } from "react";

import { api } from "@/lib/api";

const MeetFlowContext = createContext(null);
const INACTIVITY_LIMIT_MS = 60 * 60 * 1000;
const WARNING_WINDOW_MS = 5 * 60 * 1000;

function getStorage(mode) {
  return mode === "local" ? localStorage : sessionStorage;
}

function readSession() {
  const localToken = localStorage.getItem("meetflow.token");
  if (localToken) {
    return {
      mode: "local",
      token: localToken,
      userName: localStorage.getItem("meetflow.userName") || "",
      userEmail: localStorage.getItem("meetflow.userEmail") || "",
      lastActivity: Number(localStorage.getItem("meetflow.lastActivity") || "0"),
    };
  }
  const sessionToken = sessionStorage.getItem("meetflow.token");
  if (sessionToken) {
    return {
      mode: "session",
      token: sessionToken,
      userName: sessionStorage.getItem("meetflow.userName") || "",
      userEmail: sessionStorage.getItem("meetflow.userEmail") || "",
      lastActivity: Number(sessionStorage.getItem("meetflow.lastActivity") || "0"),
    };
  }
  return { mode: "session", token: "", userName: "", userEmail: "", lastActivity: 0 };
}

export function MeetFlowProvider({ children }) {
  const [token, setToken] = useState("");
  const [userName, setUserName] = useState("MeetFlow Owner");
  const [userEmail, setUserEmail] = useState("");
  const [sessionMode, setSessionMode] = useState("session");
  const [dashboard, setDashboard] = useState({ meetings_count: 0, tasks_count: 0, decisions_count: 0 });
  const [meetings, setMeetings] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [syncError, setSyncError] = useState("");
  const [sessionTimeLeftMs, setSessionTimeLeftMs] = useState(0);

  function touchActivity(mode = sessionMode) {
    if (!token) return;
    const activityStorage = getStorage(mode);
    activityStorage.setItem("meetflow.lastActivity", String(Date.now()));
  }

  useEffect(() => {
    const session = readSession();
    if (session.token) {
      if (session.lastActivity && Date.now() - session.lastActivity > INACTIVITY_LIMIT_MS) {
        localStorage.removeItem("meetflow.token");
        localStorage.removeItem("meetflow.userName");
        localStorage.removeItem("meetflow.userEmail");
        localStorage.removeItem("meetflow.lastActivity");
        sessionStorage.removeItem("meetflow.token");
        sessionStorage.removeItem("meetflow.userName");
        sessionStorage.removeItem("meetflow.userEmail");
        sessionStorage.removeItem("meetflow.lastActivity");
        return;
      }
      setSessionMode(session.mode);
      setToken(session.token);
      setUserName(session.userName || "MeetFlow Owner");
      setUserEmail(session.userEmail || "");
      refreshAll(session.token).catch((e) => {
        if (e?.status === 401) {
          logout();
        }
        // Rede / backend fora: mantém sessão e mostra aviso; não apaga histórico em cache
      });
    }
  }, []);

  useEffect(() => {
    if (!token) return;
    const activityStorage = getStorage(sessionMode);
    const markActivity = () => activityStorage.setItem("meetflow.lastActivity", String(Date.now()));
    const onActivity = () => touchActivity(sessionMode);
    markActivity();
    window.addEventListener("click", onActivity);
    window.addEventListener("keydown", onActivity);
    window.addEventListener("mousemove", onActivity);
    return () => {
      window.removeEventListener("click", onActivity);
      window.removeEventListener("keydown", onActivity);
      window.removeEventListener("mousemove", onActivity);
    };
  }, [token, sessionMode]);

  useEffect(() => {
    if (!token) {
      setSessionTimeLeftMs(0);
      return;
    }
    const syncSessionClock = () => {
      const activityStorage = getStorage(sessionMode);
      const lastActivity = Number(activityStorage.getItem("meetflow.lastActivity") || "0");
      if (!lastActivity) return;
      const elapsed = Date.now() - lastActivity;
      const left = Math.max(0, INACTIVITY_LIMIT_MS - elapsed);
      setSessionTimeLeftMs(left);
      if (elapsed > INACTIVITY_LIMIT_MS) {
        logout();
      }
    };
    syncSessionClock();
    const interval = window.setInterval(syncSessionClock, 15_000);
    return () => window.clearInterval(interval);
  }, [token, sessionMode]);

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
        logout();
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

  async function login(email, password, remember = false) {
    const result = await api.login(email, password);
    setToken(result.access_token);
    setUserName(result.user_name || "MeetFlow Owner");
    setSessionMode(remember ? "local" : "session");
    const profile = await api.me(result.access_token);
    setUserEmail(profile.email || "");
    const target = getStorage(remember ? "local" : "session");
    const other = getStorage(remember ? "session" : "local");
    other.removeItem("meetflow.token");
    other.removeItem("meetflow.userName");
    other.removeItem("meetflow.userEmail");
    other.removeItem("meetflow.lastActivity");
    target.setItem("meetflow.token", result.access_token);
    target.setItem("meetflow.userName", result.user_name || "MeetFlow Owner");
    target.setItem("meetflow.userEmail", profile.email || "");
    target.setItem("meetflow.lastActivity", String(Date.now()));
    await refreshAll(result.access_token);
  }

  async function register(payload, remember = false) {
    const result = await api.register(payload);
    setToken(result.access_token);
    setUserName(result.user_name || "MeetFlow Owner");
    setSessionMode(remember ? "local" : "session");
    const profile = await api.me(result.access_token);
    setUserEmail(profile.email || "");
    const target = getStorage(remember ? "local" : "session");
    const other = getStorage(remember ? "session" : "local");
    other.removeItem("meetflow.token");
    other.removeItem("meetflow.userName");
    other.removeItem("meetflow.userEmail");
    other.removeItem("meetflow.lastActivity");
    target.setItem("meetflow.token", result.access_token);
    target.setItem("meetflow.userName", result.user_name || "MeetFlow Owner");
    target.setItem("meetflow.userEmail", profile.email || "");
    target.setItem("meetflow.lastActivity", String(Date.now()));
    await refreshAll(result.access_token);
  }

  async function refreshProfile(authToken = token) {
    if (!authToken) return;
    const profile = await api.me(authToken);
    setUserName(profile.full_name || "MeetFlow Owner");
    setUserEmail(profile.email || "");
    const storage = getStorage(sessionMode);
    storage.setItem("meetflow.userName", profile.full_name || "MeetFlow Owner");
    storage.setItem("meetflow.userEmail", profile.email || "");
    return profile;
  }

  async function updateProfile(payload) {
    const profile = await api.updateMe(payload, token);
    setUserName(profile.full_name || "MeetFlow Owner");
    setUserEmail(profile.email || "");
    const storage = getStorage(sessionMode);
    storage.setItem("meetflow.userName", profile.full_name || "MeetFlow Owner");
    storage.setItem("meetflow.userEmail", profile.email || "");
    return profile;
  }

  async function changePassword(payload) {
    return api.changePassword(payload, token);
  }

  function logout() {
    setToken("");
    setMeetings([]);
    setDashboard({ meetings_count: 0, tasks_count: 0, decisions_count: 0 });
    setUserEmail("");
    setSessionMode("session");
    localStorage.removeItem("meetflow.token");
    localStorage.removeItem("meetflow.userName");
    localStorage.removeItem("meetflow.userEmail");
    localStorage.removeItem("meetflow.lastActivity");
    sessionStorage.removeItem("meetflow.token");
    sessionStorage.removeItem("meetflow.userName");
    sessionStorage.removeItem("meetflow.userEmail");
    sessionStorage.removeItem("meetflow.lastActivity");
  }

  function extendSession() {
    touchActivity(sessionMode);
    setSessionTimeLeftMs(INACTIVITY_LIMIT_MS);
  }

  const value = useMemo(
    () => ({
      token,
      userName,
      userEmail,
      dashboard,
      meetings,
      isLoading,
      syncError,
      clearSyncError: () => setSyncError(""),
      isSessionWarningVisible: token && sessionTimeLeftMs > 0 && sessionTimeLeftMs <= WARNING_WINDOW_MS,
      sessionTimeLeftMs,
      extendSession,
      login,
      register,
      logout,
      refreshAll,
      refreshProfile,
      updateProfile,
      changePassword,
    }),
    [token, userName, userEmail, dashboard, meetings, isLoading, syncError, sessionTimeLeftMs],
  );

  return <MeetFlowContext.Provider value={value}>{children}</MeetFlowContext.Provider>;
}

export function useMeetFlow() {
  const context = useContext(MeetFlowContext);
  if (!context) throw new Error("useMeetFlow must be used inside MeetFlowProvider");
  return context;
}
