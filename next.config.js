/** @type {import('next').NextConfig} */
const nextConfig = {
  // Keep pdfjs-dist as a Node.js external dependency so its legacy worker
  // is loaded from node_modules instead of a broken route-relative bundle.
  serverExternalPackages: ["pdfjs-dist"],
  outputFileTracingIncludes: {
    "/api/outlook/audit": [
      "./node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs",
    ],
  },
};

module.exports = nextConfig;
