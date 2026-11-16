/**
 * Konfigurasi Tailwind khusus halaman beranda (views/home.ejs).
 * Sebelumnya memakai cdn.tailwindcss.com (kompilasi di browser, lambat); kini dikompilasi sekali:
 *   npm run build-home-css
 */
module.exports = {
  content: ["./views/home.ejs"],
  theme: {
    extend: {
      colors: {
        primary: {
          50: "#f0f9ff",
          100: "#e0f2fe",
          200: "#bae6fd",
          300: "#7dd3fc",
          400: "#38bdf8",
          500: "#0ea5e9",
          600: "#0284c7",
          700: "#0369a1",
          800: "#075985",
          900: "#0c4a6e",
        },
      },
    },
  },
  plugins: [],
};
