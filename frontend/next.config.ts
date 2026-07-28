import type { NextConfig } from "next";

const RAILWAY_API =
  process.env.NEXT_PUBLIC_API_URL || "https://mytvs-invoice-excel-production.up.railway.app";

const nextConfig: NextConfig = {
  async rewrites() {
    // Local → FastAPI on :8000. Production / Vercel → Railway backend.
    const destination =
      process.env.NODE_ENV === "development"
        ? "http://127.0.0.1:8000/api/:path*"
        : `${RAILWAY_API.replace(/\/$/, "")}/api/:path*`;
    return [{ source: "/api/:path*", destination }];
  },
};

export default nextConfig;
