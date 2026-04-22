import { useState } from "react";
import { getMe, login } from "./api";

const DEFAULT_USER = "admin";
const DEFAULT_PASSWORD = "admin123";

export default function App() {
  const [username, setUsername] = useState(DEFAULT_USER);
  const [password, setPassword] = useState(DEFAULT_PASSWORD);
  const [token, setToken] = useState("");
  const [userData, setUserData] = useState(null);
  const [message, setMessage] = useState("Fase 1: Backend + Auth inicial");
  const [isLoading, setIsLoading] = useState(false);

  async function handleLogin(event) {
    event.preventDefault();
    setIsLoading(true);
    setMessage("Autenticando...");

    try {
      const result = await login(username, password);
      setToken(result.token);
      setMessage(`Login OK: ${result.user.name} (${result.user.role})`);
    } catch (error) {
      setToken("");
      setUserData(null);
      setMessage(error.message);
    } finally {
      setIsLoading(false);
    }
  }

  async function handleLoadMe() {
    if (!token) {
      setMessage("Faca login antes de chamar /api/me");
      return;
    }

    setIsLoading(true);
    setMessage("Buscando dados do usuario...");
    try {
      const me = await getMe(token);
      setUserData(me);
      setMessage("Rota protegida OK");
    } catch (error) {
      setUserData(null);
      setMessage(error.message);
    } finally {
      setIsLoading(false);
    }
  }

  return (
    <main className="container">
      <h1>Novo Projeto Web</h1>
      <p className="subtitle">
        Execucao paralela com Tkinter: backend + frontend + postgres em Docker.
      </p>

      <section className="card">
        <h2>Login de teste</h2>
        <form onSubmit={handleLogin}>
          <label>
            Usuario
            <input
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              autoComplete="username"
            />
          </label>

          <label>
            Senha
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              autoComplete="current-password"
            />
          </label>

          <button type="submit" disabled={isLoading}>
            Entrar
          </button>
        </form>
      </section>

      <section className="card">
        <h2>Rota protegida</h2>
        <button onClick={handleLoadMe} disabled={isLoading || !token}>
          Testar /api/me
        </button>
        {userData && <pre>{JSON.stringify(userData, null, 2)}</pre>}
      </section>

      <p className="message">{message}</p>
    </main>
  );
}
