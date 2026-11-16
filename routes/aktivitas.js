const express = require("express");
const { getSettings, lokasi } = require("../utils/settings");
const { validIso } = require("../utils/dates");

const PER_PAGE = 30;
const router = express.Router();

router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const settings = await getSettings(db);
    const tz = lokasi(settings).tz;

    const modul = String(req.query.modul || "").slice(0, 40);
    const aksi = String(req.query.aksi || "").slice(0, 40);
    const userId = parseInt(req.query.user, 10) || 0;
    const q = String(req.query.q || "").trim().slice(0, 80);
    const dari = validIso(req.query.dari) ? req.query.dari : "";
    const sampai = validIso(req.query.sampai) ? req.query.sampai : "";
    const page = Math.max(1, parseInt(req.query.page, 10) || 1);

    const where = [];
    const params = [];
    if (modul) { where.push("modul = ?"); params.push(modul); }
    if (aksi) { where.push("aksi = ?"); params.push(aksi); }
    if (userId) { where.push("user_id = ?"); params.push(userId); }
    if (q) { where.push("(ringkasan LIKE ? OR user_name LIKE ?)"); params.push(`%${q}%`, `%${q}%`); }
    // rentang tanggal mengikuti zona waktu masjid (tengah malam setempat -> instan absolut)
    if (dari) { where.push("created_at >= ?"); params.push(new Date(Date.parse(dari + "T00:00:00Z") - tz * 3600e3)); }
    if (sampai) { where.push("created_at < ?"); params.push(new Date(Date.parse(sampai + "T00:00:00Z") + 864e5 - tz * 3600e3)); }
    const sql = where.length ? "WHERE " + where.join(" AND ") : "";

    const [[{ total }]] = await db.execute(`SELECT COUNT(*) AS total FROM log_aktivitas ${sql}`, params);
    const totalPages = Math.max(1, Math.ceil(total / PER_PAGE));
    const cur = Math.min(page, totalPages);
    const [logs] = await db.query(
      `SELECT id, user_id, user_name, aksi, modul, objek_id, ringkasan, ip, created_at FROM log_aktivitas ${sql}
        ORDER BY id DESC LIMIT ${PER_PAGE} OFFSET ${(cur - 1) * PER_PAGE}`,
      params
    );

    const [[stat]] = await db.execute(
      `SELECT
         COALESCE(SUM(created_at >= (NOW() - INTERVAL 1 DAY)),0) AS hari,
         COALESCE(SUM(created_at >= (NOW() - INTERVAL 7 DAY)),0) AS minggu,
         COALESCE(SUM(aksi = 'Menghapus' AND created_at >= (NOW() - INTERVAL 7 DAY)),0) AS hapus,
         COALESCE(SUM(aksi = 'Login gagal' AND created_at >= (NOW() - INTERVAL 7 DAY)),0) AS gagal,
         COUNT(DISTINCT CASE WHEN created_at >= (NOW() - INTERVAL 1 DAY) THEN user_id END) AS aktif
       FROM log_aktivitas`
    );
    const [moduls] = await db.execute("SELECT DISTINCT modul FROM log_aktivitas ORDER BY modul");
    const [aksis] = await db.execute("SELECT DISTINCT aksi FROM log_aktivitas ORDER BY aksi");
    const [users] = await db.execute("SELECT id, name FROM users ORDER BY name");

    res.render("aktivitas/index", {
      title: "Log Aktivitas - " + settings.masjid_nama,
      user: req.session.user,
      currentPage: "aktivitas",
      settings,
      tz,
      f: { modul, aksi, user: userId, q, dari, sampai },
      logs,
      total,
      page: cur,
      totalPages,
      perPage: PER_PAGE,
      stat: Object.fromEntries(Object.entries(stat).map(([k, v]) => [k, Number(v)])),
      moduls: moduls.map((m) => m.modul),
      aksis: aksis.map((a) => a.aksi),
      users,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error log aktivitas:", error);
    req.flash("error", "Gagal memuat log aktivitas");
    res.redirect("/dashboard");
  }
});

module.exports = { router };
