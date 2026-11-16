const express = require("express");
const path = require("path");
const { getSettings } = require("../utils/settings");

const router = express.Router();

// Manifest memakai nama musholla dari pengaturan
router.get("/manifest.webmanifest", async (req, res) => {
  let nama = "Musholla Nurul Fitrah";
  try {
    nama = (await getSettings(req.app.locals.db)).masjid_nama || nama;
  } catch (error) {
    // pakai nama bawaan
  }
  res.setHeader("Content-Type", "application/manifest+json; charset=utf-8");
  res.setHeader("Cache-Control", "public, max-age=3600");
  res.json({
    name: nama,
    short_name: nama.replace(/^(musholla|masjid)\s+/i, "").slice(0, 14) || "Musholla",
    description: "Jadwal Ramadan, donasi, dan laporan kas terbuka " + nama,
    id: "/",
    start_url: "/?source=pwa",
    scope: "/",
    display: "standalone",
    orientation: "portrait",
    lang: "id",
    background_color: "#f4faf5",
    theme_color: "#14912a",
    categories: ["lifestyle", "productivity"],
    icons: [
      { src: "/icons/icon-192.png", sizes: "192x192", type: "image/png", purpose: "any" },
      { src: "/icons/icon-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
      { src: "/icons/icon-maskable-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
    ],
    shortcuts: [
      { name: "Jadwal Ramadan", url: "/jadwal-ramadan?source=shortcut", icons: [{ src: "/icons/icon-192.png", sizes: "192x192" }] },
      { name: "Donasi", url: "/donasi?source=shortcut", icons: [{ src: "/icons/icon-192.png", sizes: "192x192" }] },
      { name: "Laporan Kas", url: "/transparansi?source=shortcut", icons: [{ src: "/icons/icon-192.png", sizes: "192x192" }] },
    ],
  });
});

// Service worker harus dari root agar cakupannya seluruh situs, dan tidak boleh tersimpan lama
router.get("/sw.js", (req, res) => {
  res.setHeader("Content-Type", "application/javascript; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Service-Worker-Allowed", "/");
  res.sendFile(path.join(__dirname, "..", "public", "sw.js"));
});

router.get("/offline", async (req, res) => {
  let nama = "Musholla Nurul Fitrah";
  try {
    nama = (await getSettings(req.app.locals.db)).masjid_nama || nama;
  } catch (error) {
    // pakai nama bawaan
  }
  res.setHeader("Cache-Control", "no-store");
  res.render("offline", { nama, layout: false });
});

// Menandai halaman publik (tanpa data akun) agar boleh disimpan untuk dibuka tanpa internet
function markPublic(req, res, next) {
  res.setHeader("X-App-Public", req.session && req.session.user ? "0" : "1");
  next();
}

module.exports = { router, markPublic };
