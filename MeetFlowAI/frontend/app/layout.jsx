import "./globals.css";

export const metadata = {
  title: "MeetFlow AI MVP",
  description: "Reunioes em atas profissionais com IA",
};

export default function RootLayout({ children }) {
  return (
    <html lang="pt-BR">
      <body>{children}</body>
    </html>
  );
}
