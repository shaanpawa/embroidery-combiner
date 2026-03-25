import type { NextConfig } from "next";

const isStaticBuild = process.env.BUILD_STATIC === "true";

const nextConfig: NextConfig = {
  output: isStaticBuild ? "export" : undefined,
  // trailingSlash creates /stacker/index.html instead of stacker.html
  // which is needed for FastAPI StaticFiles(html=True) to serve /stacker correctly
  trailingSlash: isStaticBuild ? true : undefined,
  async redirects() {
    // Redirects are not supported in static export mode
    if (isStaticBuild) return [];
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
