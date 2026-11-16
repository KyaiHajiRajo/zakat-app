const express = require("express");
const bcrypt = require("bcryptjs");

const router = express.Router();

const BACK = "/users/profile";

router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const [rows] = await db.execute(
      "SELECT id, name, nomor_wa, role, created_at FROM users WHERE id = ?",
      [req.session.user.id]
    );

    if (rows.length === 0) {
      return req.session.destroy(() => res.redirect("/auth/login"));
    }

    res.render("users/profile", {
      title: "Profil Saya - Zakat Fitrah",
      currentPage: "profile",
      user: req.session.user,
      profile: rows[0],
    });
  } catch (error) {
    console.error("Error loading profile:", error);
    req.flash("error_msg", "Gagal memuat profil");
    res.redirect("/dashboard");
  }
});

router.post("/info", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const name = String(req.body.name || "").trim();
    const nomorWa = String(req.body.nomor_wa || "").trim();

    if (!name || name.length > 100) {
      req.flash("error_msg", "Nama wajib diisi (maksimal 100 karakter)");
      return res.redirect(BACK);
    }

    if (!/^[0-9]{10,15}$/.test(nomorWa)) {
      req.flash("error_msg", "Nomor WhatsApp harus 10-15 digit angka");
      return res.redirect(BACK);
    }

    const [existing] = await db.execute(
      "SELECT id FROM users WHERE nomor_wa = ? AND id != ?",
      [nomorWa, req.session.user.id]
    );

    if (existing.length > 0) {
      req.flash("error_msg", "Nomor WhatsApp sudah digunakan user lain");
      return res.redirect(BACK);
    }

    await db.execute("UPDATE users SET name = ?, nomor_wa = ? WHERE id = ?", [
      name,
      nomorWa,
      req.session.user.id,
    ]);

    req.session.user.name = name;
    req.session.user.nomor_wa = nomorWa;

    req.flash("success_msg", "Profil berhasil diperbarui");
    res.redirect(BACK);
  } catch (error) {
    console.error("Error updating profile:", error);
    req.flash("error_msg", "Gagal memperbarui profil");
    res.redirect(BACK);
  }
});

router.post("/password", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { current_password, new_password, confirm_password } = req.body;

    if (!current_password || !new_password || !confirm_password) {
      req.flash("error_msg", "Semua kolom password wajib diisi");
      return res.redirect(BACK);
    }

    if (new_password.length < 8) {
      req.flash("error_msg", "Password baru minimal 8 karakter");
      return res.redirect(BACK);
    }

    if (new_password !== confirm_password) {
      req.flash("error_msg", "Konfirmasi password baru tidak cocok");
      return res.redirect(BACK);
    }

    const [rows] = await db.execute("SELECT password FROM users WHERE id = ?", [
      req.session.user.id,
    ]);

    if (
      rows.length === 0 ||
      !(await bcrypt.compare(current_password, rows[0].password))
    ) {
      req.flash("error_msg", "Password lama salah");
      return res.redirect(BACK);
    }

    const hashed = await bcrypt.hash(new_password, 10);
    await db.execute("UPDATE users SET password = ? WHERE id = ?", [
      hashed,
      req.session.user.id,
    ]);

    req.flash("success_msg", "Password berhasil diubah");
    res.redirect(BACK);
  } catch (error) {
    console.error("Error changing password:", error);
    req.flash("error_msg", "Gagal mengubah password");
    res.redirect(BACK);
  }
});

module.exports = router;
