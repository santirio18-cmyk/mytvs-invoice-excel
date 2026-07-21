import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Local-only proxy. In production, set NEXT_PUBLIC_API_URL to your Railway backend.
  async rewrites() {
    if (process.env.NODE_ENV === "production" && process.env.NEXT_PUBLIC_API_URL) {
      return [];
    }
    return [
      {
        source: "/api/:path*",
        destination: "http://127.0.0.1:8000/api/:path*",
      },
    ];
  },
};

export default nextConfig;
