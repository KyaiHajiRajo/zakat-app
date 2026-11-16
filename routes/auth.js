const express = require("express");
const bcrypt = require("bcryptjs");
const { createLimiter } = require("../utils/rate-limit");
const { isLoopback } = require("../utils/security");
const { logAktivitas } = require("../utils/audit");
const router = express.Router();

// 5 kali salah untuk satu nomor dari satu perangkat, atau 30 kali dari satu perangkat -> kunci 15 menit
const WINDOW_MS = 15 * 60 * 1000;
const accountLimiter = createLimiter({ windowMs: WINDOW_MS, max: 5 });
const ipLimiter = createLimiter({ windowMs: WINDOW_MS, max: 30 });
// pembanding palsu agar waktu respons sama baik nomor terdaftar maupun tidak
const DUMMY_HASH = bcrypt.hashSync("tidak-dipakai", 10);

// GET /auth/login - Show login form
router.get("/login", (req, res) => {
  // Redirect to dashboard if already logged in
  if (req.session.user) {
    return res.redirect("/");
  }

  // Get messages from query parameters
  const success = req.query.success;
  const error = req.query.error;

  res.render("auth/login", {
    title: "Login - Zakat Fitrah",
    layout: "layouts/main",
    success: success,
    error: error,
  });
});

// POST /auth/login - Process login
router.post("/login", async (req, res) => {
  const { nomor_wa, password } = req.body;

  try {
    // Input validation
    if (!nomor_wa || !password) {
      req.flash("error_msg", "Nomor WhatsApp dan password harus diisi");
      return res.redirect("/auth/login");
    }

    const ipKey = req.ip;
    const accountKey = `${req.ip}|${String(nomor_wa).trim()}`;
    const lockedFor = Math.max(ipLimiter.blockedFor(ipKey), accountLimiter.blockedFor(accountKey));
    if (lockedFor > 0) {
      req.flash("error_msg", `Terlalu banyak percobaan login. Coba lagi dalam ${Math.ceil(lockedFor / 60)} menit.`);
      return res.redirect("/auth/login");
    }

    // Get database connection
    const db = req.app.locals.db;

    // Find user by nomor_wa
    const [users] = await db.execute("SELECT * FROM users WHERE nomor_wa = ?", [
      nomor_wa,
    ]);

    const user = users[0];

    // Check password (tetap dihitung walau user tidak ada)
    const isPasswordValid = await bcrypt.compare(
      String(password),
      user ? user.password : DUMMY_HASH
    );

    if (!user || !isPasswordValid) {
      ipLimiter.fail(ipKey);
      accountLimiter.fail(accountKey);
      logAktivitas(db, { user: user ? { id: user.id, name: user.name } : null, aksi: "Login gagal", modul: "Akun", ringkasan: `Login gagal untuk nomor ${String(nomor_wa).slice(0, 20)}`, ip: req.ip });
      req.flash("error_msg", "Nomor WhatsApp atau password salah");
      return res.redirect("/auth/login");
    }

    accountLimiter.reset(accountKey);

    // Sesi baru setelah login (mencegah session fixation)
    req.session.regenerate((regenError) => {
      if (regenError) {
        console.error("Session regenerate error:", regenError);
        req.flash("error_msg", "Terjadi kesalahan saat login");
        return res.redirect("/auth/login");
      }

      req.session.user = {
        id: user.id,
        name: user.name,
        nomor_wa: user.nomor_wa,
        role: user.role,
      };

      logAktivitas(db, { user: { id: user.id, name: user.name }, aksi: "Login", modul: "Akun", ringkasan: "Masuk ke aplikasi", ip: req.ip });
      req.flash("success", `Selamat datang, ${user.name}!`);
      res.redirect("/dashboard");
    });
  } catch (error) {
    console.error("Login error:", error);
    req.flash("error_msg", "Terjadi kesalahan saat login");
    res.redirect("/auth/login");
  }
});

// GET /auth/logout - Logout user
router.get("/logout", (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      console.error("Logout error:", err);
      return res.redirect("/?error=Terjadi kesalahan saat logout");
    }
    res.redirect("/?success=Anda telah berhasil logout");
  });
});

// POST /auth/logout - Logout user (for form submission)
router.post("/logout", (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      console.error("Logout error:", err);
      return res.redirect("/?error=Terjadi kesalahan saat logout");
    }
    res.redirect("/?success=Anda telah berhasil logout");
  });
});

// GET /auth/register - Show register form (only for development)
router.get("/register", (req, res) => {
  // Hanya mode development dan hanya dari komputer ini sendiri
  if (process.env.NODE_ENV !== "development" || !isLoopback(req)) {
    return res.status(404).send("Not Found");
  }

  res.render("auth/register", {
    title: "Register - Zakat Fitrah",
    layout: "layouts/main",
  });
});

// POST /auth/register - Process registration (only for development)
router.post("/register", async (req, res) => {
  // Hanya mode development dan hanya dari komputer ini sendiri
  if (process.env.NODE_ENV !== "development" || !isLoopback(req)) {
    return res.status(404).send("Not Found");
  }

  const { name, nomor_wa, password, confirm_password } = req.body;

  try {
    // Input validation
    if (!name || !nomor_wa || !password || !confirm_password) {
      req.flash("error_msg", "Semua field harus diisi");
      return res.redirect("/auth/register");
    }

    if (password !== confirm_password) {
      req.flash("error_msg", "Password dan konfirmasi password tidak cocok");
      return res.redirect("/auth/register");
    }

    if (password.length < 8) {
      req.flash("error_msg", "Password minimal 8 karakter");
      return res.redirect("/auth/register");
    }

    // Check if nomor_wa already exists
    const db = req.app.locals.db;
    const [existingUsers] = await db.execute(
      "SELECT id FROM users WHERE nomor_wa = ?",
      [nomor_wa]
    );

    if (existingUsers.length > 0) {
      req.flash("error_msg", "Nomor WhatsApp sudah terdaftar");
      return res.redirect("/auth/register");
    }

    // Hash password
    const hashedPassword = await bcrypt.hash(password, 10);

    // Akun pertama menjadi admin; selanjutnya selalu panitia (admin dibuat lewat menu Pengguna)
    const [[{ total }]] = await db.execute("SELECT COUNT(*) AS total FROM users");
    await db.execute(
      "INSERT INTO users (name, nomor_wa, password, role) VALUES (?, ?, ?, ?)",
      [name, nomor_wa, hashedPassword, total === 0 ? "admin" : "panitia"]
    );

    req.flash("success_msg", "Registrasi berhasil! Silakan login");
    res.redirect("/auth/login");
  } catch (error) {
    console.error("Registration error:", error);
    req.flash("error_msg", "Terjadi kesalahan saat registrasi");
    res.redirect("/auth/register");
  }
});

module.exports = router;
