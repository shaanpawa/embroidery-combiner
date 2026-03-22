import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  async redirects() {
    return [
      {
        source: "/combo-builder",
        destination: "/stacker",
        permanent: true,
      },
    ];
  },
};

export default nextConfig;
