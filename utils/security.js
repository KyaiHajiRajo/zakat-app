const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const DEFAULT_SECRET = "your-default-secret-key";
const SECRET_FILE = path.join(__dirname, "..", ".session-secret");

/**
 * Rahasia sesi: pakai SESSION_SECRET dari .env; jika kosong, buat sekali lalu simpan
 * di file `.session-secret` (di-ignore git) supaya login tidak hilang saat server restart.
 */
function getSessionSecret() {
  const fromEnv = process.env.SESSION_SECRET;
  if (fromEnv && fromEnv !== DEFAULT_SECRET && fromEnv.length >= 16) {
    return fromEnv;
  }

  try {
    const saved = fs.readFileSync(SECRET_FILE, "utf8").trim();
    if (saved.length >= 32) {
      return saved;
    }
  } catch (error) {
    // belum ada, dibuat di bawah
  }

  const secret = crypto.randomBytes(48).toString("hex");
  fs.writeFileSync(SECRET_FILE, secret + "\n", { mode: 0o600 });
  console.log("🔐 Rahasia sesi dibuat otomatis (.session-secret). Isi SESSION_SECRET di .env untuk menggantinya.");
  return secret;
}

/** Header keamanan dasar. CSP sengaja hanya membatasi hal yang tidak merusak inline script/style. */
function securityHeaders(req, res, next) {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "SAMEORIGIN");
  res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
  res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()");
  res.setHeader("Cross-Origin-Opener-Policy", "same-origin");
  res.setHeader(
    "Content-Security-Policy",
    "frame-ancestors 'self'; object-src 'none'; base-uri 'self'; form-action 'self'"
  );
  if (req.secure) {
    res.setHeader("Strict-Transport-Security", "max-age=15552000");
  }
  next();
}

const SAFE_METHODS = new Set(["GET", "HEAD", "OPTIONS"]);

function trustedOrigins() {
  return String(process.env.TRUSTED_ORIGINS || "")
    .split(",")
    .map((item) => item.trim().replace(/\/$/, ""))
    .filter(Boolean);
}

/**
 * Perlindungan CSRF tanpa token: permintaan yang mengubah data harus datang dari
 * halaman aplikasi ini sendiri (Sec-Fetch-Site / Origin / Referer). Dipadukan dengan
 * cookie SameSite=Lax. Klien non-browser (curl) tidak membawa cookie sesi orang lain, jadi dibiarkan.
 */
function sameOriginGuard(req, res, next) {
  if (SAFE_METHODS.has(req.method)) {
    return next();
  }

  const fetchSite = req.get("sec-fetch-site");
  if (fetchSite) {
    if (fetchSite === "same-origin" || fetchSite === "none") {
      return next();
    }
    return deny(req, res);
  }

  const source = req.get("origin") || req.get("referer");
  if (!source) {
    return next();
  }

  try {
    const url = new URL(source);
    if (url.host === req.get("host") || trustedOrigins().includes(url.origin)) {
      return next();
    }
  } catch (error) {
    // asal tidak valid -> tolak
  }
  return deny(req, res);
}

function deny(req, res) {
  const message = "Permintaan ditolak karena berasal dari luar aplikasi. Muat ulang halaman lalu coba lagi.";
  if (req.get("x-requested-with") === "fetch" || (req.get("accept") || "").includes("application/json")) {
    return res.status(403).json({ success: false, message });
  }
  return res.status(403).render("error", {
    title: "Permintaan Ditolak",
    message,
    error: { status: 403 },
  });
}

function isLoopback(req) {
  const address = req.socket && req.socket.remoteAddress;
  return address === "127.0.0.1" || address === "::1" || address === "::ffff:127.0.0.1";
}

module.exports = { getSessionSecret, securityHeaders, sameOriginGuard, isLoopback };
