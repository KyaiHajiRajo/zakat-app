const express = require("express");
const { getSettings, lokasi } = require("../utils/settings");
const { reply, safeReturn } = require("../utils/respond");
const { validIso, dateToIso, todayIso, labelTanggal, hariTanggal } = require("../utils/dates");

const JENIS = {
  pengajian: { label: "Pengajian", icon: "fa-book-open", tone: "green" },
  kerja_bakti: { label: "Kerja Bakti", icon: "fa-broom", tone: "gold" },
  peringatan: { label: "Peringatan Hari Besar", icon: "fa-star-and-crescent", tone: "blue" },
  rapat: { label: "Rapat Pengurus", icon: "fa-people-group", tone: "gray" },
  sosial: { label: "Kegiatan Sosial", icon: "fa-hand-holding-heart", tone: "red" },
  ramadan: { label: "Kegiatan Ramadan", icon: "fa-moon", tone: "green" },
  lainnya: { label: "Lainnya", icon: "fa-calendar-day", tone: "gray" },
};
const TIME = /^([01]\d|2[0-3]):[0-5]\d$/;

const mapAgenda = (r) => ({
  ...r,
  tanggal: dateToIso(r.tanggal),
  jenisLabel: (JENIS[r.jenis] || JENIS.lainnya).label,
  icon: (JENIS[r.jenis] || JENIS.lainnya).icon,
  tone: (JENIS[r.jenis] || JENIS.lainnya).tone,
});
const mapPengumuman = (r) => ({ ...r, berlaku_sampai: dateToIso(r.berlaku_sampai), penting: !!r.penting, tampil_publik: !!r.tampil_publik });

const AGENDA_COLS = "id, judul, jenis, tanggal, TIME_FORMAT(jam_mulai,'%H:%i') AS jam_mulai, TIME_FORMAT(jam_selesai,'%H:%i') AS jam_selesai, lokasi, pemateri, deskripsi, tampil_publik";

async function loadData(db, today, { publicOnly }) {
  const pub = publicOnly ? "AND tampil_publik = 1" : "";
  const [mendatang] = await db.execute(`SELECT ${AGENDA_COLS} FROM agenda WHERE tanggal >= ? ${pub} ORDER BY tanggal ASC, jam_mulai ASC, id ASC LIMIT 60`, [today]);
  const [lampau] = publicOnly
    ? [[]]
    : await db.execute(`SELECT ${AGENDA_COLS} FROM agenda WHERE tanggal < ? ORDER BY tanggal DESC, id DESC LIMIT 15`, [today]);
  const [aktif] = await db.execute(
    `SELECT id, judul, isi, penting, berlaku_sampai, tampil_publik, created_at FROM pengumuman
      WHERE (berlaku_sampai IS NULL OR berlaku_sampai >= ?) ${pub} ORDER BY penting DESC, created_at DESC LIMIT 50`,
    [today]
  );
  const [arsip] = publicOnly
    ? [[]]
    : await db.execute(
        `SELECT id, judul, isi, penting, berlaku_sampai, tampil_publik, created_at FROM pengumuman
          WHERE berlaku_sampai IS NOT NULL AND berlaku_sampai < ? ORDER BY berlaku_sampai DESC LIMIT 15`,
        [today]
      );
  return { mendatang: mendatang.map(mapAgenda), lampau: lampau.map(mapAgenda), aktif: aktif.map(mapPengumuman), arsip: arsip.map(mapPengumuman) };
}

function base(req, settings, today) {
  return { settings, JENIS, today, labelTanggal, hariTanggal, success: req.flash("success"), error: req.flash("error") };
}

// ---------- Halaman publik ----------
const publicRouter = express.Router();
publicRouter.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const settings = await getSettings(db);
    const today = todayIso(lokasi(settings).tz);
    const data = await loadData(db, today, { publicOnly: true });
    res.render("info/index", { ...base(req, settings, today), ...data, title: "Info " + settings.masjid_nama, user: req.session.user || null, currentPage: "agenda" });
  } catch (error) {
    console.error("Error info publik:", error);
    res.status(500).render("error", { title: "Server Error", message: "Gagal memuat informasi", error: {} });
  }
});

// ---------- Pengelola ----------
const router = express.Router();

router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const settings = await getSettings(db);
    const today = todayIso(lokasi(settings).tz);
    const data = await loadData(db, today, { publicOnly: false });
    res.render("agenda/index", { ...base(req, settings, today), ...data, title: "Agenda & Pengumuman - " + settings.masjid_nama, user: req.session.user, currentPage: "agenda", isAdmin: req.session.user.role === "admin" });
  } catch (error) {
    console.error("Error agenda:", error);
    req.flash("error", "Gagal memuat agenda");
    res.redirect("/dashboard");
  }
});

function validateAgenda(b) {
  const judul = String(b.judul || "").trim();
  const jenis = String(b.jenis || "");
  const tanggal = String(b.tanggal || "").trim();
  const jam_mulai = String(b.jam_mulai || "").trim();
  const jam_selesai = String(b.jam_selesai || "").trim();
  const lokasiTxt = String(b.lokasi || "").trim();
  const pemateri = String(b.pemateri || "").trim();
  const deskripsi = String(b.deskripsi || "").trim();
  if (!judul || judul.length > 150) return { error: "Judul wajib diisi (maks. 150 karakter)" };
  if (!JENIS[jenis]) return { error: "Jenis agenda tidak valid" };
  if (!validIso(tanggal)) return { error: "Tanggal tidak valid" };
  if (jam_mulai && !TIME.test(jam_mulai)) return { error: "Jam mulai tidak valid (format 24 jam, mis. 19:30)" };
  if (jam_selesai && !TIME.test(jam_selesai)) return { error: "Jam selesai tidak valid" };
  if (jam_mulai && jam_selesai && jam_selesai < jam_mulai) return { error: "Jam selesai tidak boleh sebelum jam mulai" };
  if (lokasiTxt.length > 150 || pemateri.length > 120 || deskripsi.length > 2000) return { error: "Lokasi/pemateri/deskripsi terlalu panjang" };
  return {
    value: { judul, jenis, tanggal, jam_mulai: jam_mulai || null, jam_selesai: jam_selesai || null, lokasi: lokasiTxt || null, pemateri: pemateri || null, deskripsi: deskripsi || null, tampil_publik: b.tampil_publik === "1" ? 1 : 0 },
  };
}

function validatePengumuman(b) {
  const judul = String(b.judul || "").trim();
  const isi = String(b.isi || "").trim();
  const berlaku = String(b.berlaku_sampai || "").trim();
  if (!judul || judul.length > 150) return { error: "Judul wajib diisi (maks. 150 karakter)" };
  if (!isi || isi.length > 2000) return { error: "Isi pengumuman wajib diisi (maks. 2000 karakter)" };
  if (berlaku && !validIso(berlaku)) return { error: "Tanggal berlaku tidak valid" };
  return { value: { judul, isi, penting: b.penting === "1" ? 1 : 0, berlaku_sampai: berlaku || null, tampil_publik: b.tampil_publik === "1" ? 1 : 0 } };
}

const BACK = (req) => safeReturn(req.body.return_to, "/agenda");

router.get("/agenda/form-data", async (req, res) => {
  try {
    const [rows] = await req.app.locals.db.execute(
      `SELECT id, judul, jenis, DATE_FORMAT(tanggal,'%Y-%m-%d') AS tanggal, TIME_FORMAT(jam_mulai,'%H:%i') AS jam_mulai, TIME_FORMAT(jam_selesai,'%H:%i') AS jam_selesai, lokasi, pemateri, deskripsi, tampil_publik FROM agenda WHERE id = ?`,
      [req.query.id]
    );
    if (!rows.length) return res.status(404).json({ ok: false, message: "Agenda tidak ditemukan" });
    res.json({ ok: true, agenda: rows[0] });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, message: "Gagal memuat data" });
  }
});

router.post("/agenda", async (req, res) => {
  const { value: v, error } = validateAgenda(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: BACK(req) });
  try {
    await req.app.locals.db.execute(
      "INSERT INTO agenda (judul, jenis, tanggal, jam_mulai, jam_selesai, lokasi, pemateri, deskripsi, tampil_publik, user_id) VALUES (?,?,?,?,?,?,?,?,?,?)",
      [v.judul, v.jenis, v.tanggal, v.jam_mulai, v.jam_selesai, v.lokasi, v.pemateri, v.deskripsi, v.tampil_publik, req.session.user.id]
    );
    reply(req, res, { ok: true, message: `Agenda “${v.judul}” ditambahkan`, redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal menyimpan agenda", redirect: BACK(req), status: 500 });
  }
});

router.post("/agenda/:id/edit", async (req, res) => {
  const { value: v, error } = validateAgenda(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: BACK(req) });
  try {
    const [r] = await req.app.locals.db.execute(
      "UPDATE agenda SET judul=?, jenis=?, tanggal=?, jam_mulai=?, jam_selesai=?, lokasi=?, pemateri=?, deskripsi=?, tampil_publik=? WHERE id=?",
      [v.judul, v.jenis, v.tanggal, v.jam_mulai, v.jam_selesai, v.lokasi, v.pemateri, v.deskripsi, v.tampil_publik, req.params.id]
    );
    if (!r.affectedRows) return reply(req, res, { ok: false, message: "Agenda tidak ditemukan", redirect: BACK(req), status: 404 });
    reply(req, res, { ok: true, message: "Agenda diperbarui", redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal memperbarui agenda", redirect: BACK(req), status: 500 });
  }
});

router.post("/agenda/:id/delete", async (req, res) => {
  try {
    await req.app.locals.db.execute("DELETE FROM agenda WHERE id = ?", [req.params.id]);
    reply(req, res, { ok: true, message: "Agenda dihapus", redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal menghapus agenda", redirect: BACK(req), status: 500 });
  }
});

router.get("/pengumuman/form-data", async (req, res) => {
  try {
    const [rows] = await req.app.locals.db.execute(
      "SELECT id, judul, isi, penting, DATE_FORMAT(berlaku_sampai,'%Y-%m-%d') AS berlaku_sampai, tampil_publik FROM pengumuman WHERE id = ?",
      [req.query.id]
    );
    if (!rows.length) return res.status(404).json({ ok: false, message: "Pengumuman tidak ditemukan" });
    res.json({ ok: true, pengumuman: rows[0] });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, message: "Gagal memuat data" });
  }
});

router.post("/pengumuman", async (req, res) => {
  const { value: v, error } = validatePengumuman(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: BACK(req) });
  try {
    await req.app.locals.db.execute(
      "INSERT INTO pengumuman (judul, isi, penting, berlaku_sampai, tampil_publik, user_id) VALUES (?,?,?,?,?,?)",
      [v.judul, v.isi, v.penting, v.berlaku_sampai, v.tampil_publik, req.session.user.id]
    );
    reply(req, res, { ok: true, message: "Pengumuman diterbitkan", redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal menyimpan pengumuman", redirect: BACK(req), status: 500 });
  }
});

router.post("/pengumuman/:id/edit", async (req, res) => {
  const { value: v, error } = validatePengumuman(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: BACK(req) });
  try {
    const [r] = await req.app.locals.db.execute(
      "UPDATE pengumuman SET judul=?, isi=?, penting=?, berlaku_sampai=?, tampil_publik=? WHERE id=?",
      [v.judul, v.isi, v.penting, v.berlaku_sampai, v.tampil_publik, req.params.id]
    );
    if (!r.affectedRows) return reply(req, res, { ok: false, message: "Pengumuman tidak ditemukan", redirect: BACK(req), status: 404 });
    reply(req, res, { ok: true, message: "Pengumuman diperbarui", redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal memperbarui pengumuman", redirect: BACK(req), status: 500 });
  }
});

router.post("/pengumuman/:id/delete", async (req, res) => {
  try {
    await req.app.locals.db.execute("DELETE FROM pengumuman WHERE id = ?", [req.params.id]);
    reply(req, res, { ok: true, message: "Pengumuman dihapus", redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal menghapus pengumuman", redirect: BACK(req), status: 500 });
  }
});

module.exports = { router, publicRouter, JENIS };
