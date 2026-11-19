/** @type {import('next').NextConfig} */
const nextConfig = {
  async rewrites() {
    return [
      {
        source: '/eden/:path*',
        destination: 'http://127.0.0.1:8001/eden/:path*',
      },
    ];
  },
};
module.exports = nextConfig;
