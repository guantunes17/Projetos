/** @type {import('next').NextConfig} */
// No Docker, o browser acessa o Next; o proxy no servidor encaminha para o backend.
const apiOrigin = process.env.MEETFLOW_API_ORIGIN || "http://127.0.0.1:8000";

const nextConfig = {
  output: "standalone",
  async rewrites() {
    return [
      {
        source: "/__meetflow/api/:path*",
        destination: `${apiOrigin.replace(/\/$/, "")}/api/:path*`,
      },
    ];
  },
};

export default nextConfig;
