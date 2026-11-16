const express = require("express");
const { getSettings } = require("../utils/settings");
const { reply, safeReturn } = require("../utils/respond");
const { validIso, dateToIso, labelTanggal } = require("../utils/dates");

const KATEGORI = {
  perlengkapan_ibadah: { label: "Perlengkapan Ibadah", icon: "fa-mosque" },
  elektronik: { label: "Elektronik & Audio", icon: "fa-volume-high" },
  furnitur: { label: "Furnitur", icon: "fa-couch" },
  kebersihan: { label: "Kebersihan", icon: "fa-broom" },
  konsumsi: { label: "Peralatan Konsumsi", icon: "fa-utensils" },
  lainnya: { label: "Lainnya", icon: "fa-box" },
};
const KONDISI = {
  baik: { label: "Baik", tone: "green" },
  perlu_perbaikan: { label: "Perlu Perbaikan", tone: "gold" },
  rusak: { label: "Rusak", tone: "red" },
};

const router = express.Router();

router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const settings = await getSettings(db);
    const kategori = KATEGORI[req.query.kategori] ? req.query.kategori : "";
    const kondisi = KONDISI[req.query.kondisi] ? req.query.kondisi : "";
    const q = String(req.query.q || "").trim().slice(0, 80);

    const where = [];
    const params = [];
    if (kategori) { where.push("kategori = ?"); params.push(kategori); }
    if (kondisi) { where.push("kondisi = ?"); params.push(kondisi); }
    if (q) { where.push("(nama LIKE ? OR lokasi LIKE ? OR catatan LIKE ?)"); params.push(`%${q}%`, `%${q}%`, `%${q}%`); }

    const [items] = await db.execute(
      `SELECT id, nama, kategori, jumlah, satuan, kondisi, lokasi, tanggal_perolehan, sumber, catatan
         FROM inventaris ${where.length ? "WHERE " + where.join(" AND ") : ""} ORDER BY kategori, nama LIMIT 500`,
      params
    );
    const [[totals]] = await db.execute(
      `SELECT COUNT(*) AS jenis, COALESCE(SUM(jumlah),0) AS unit,
              COALESCE(SUM(CASE WHEN kondisi='perlu_perbaikan' THEN 1 END),0) AS perbaikan,
              COALESCE(SUM(CASE WHEN kondisi='rusak' THEN 1 END),0) AS rusak FROM inventaris`
    );

    res.render("inventaris/index", {
      title: "Inventaris - " + settings.masjid_nama,
      user: req.session.user,
      currentPage: "inventaris",
      isAdmin: req.session.user.role === "admin",
      settings,
      KATEGORI,
      KONDISI,
      f: { kategori, kondisi, q },
      items: items.map((i) => ({ ...i, tanggal_perolehan: dateToIso(i.tanggal_perolehan), kategoriLabel: KATEGORI[i.kategori].label, icon: KATEGORI[i.kategori].icon, kondisiLabel: KONDISI[i.kondisi].label, tone: KONDISI[i.kondisi].tone })),
      totals: { jenis: Number(totals.jenis), unit: Number(totals.unit), perbaikan: Number(totals.perbaikan), rusak: Number(totals.rusak) },
      labelTanggal,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error inventaris:", error);
    req.flash("error", "Gagal memuat inventaris");
    res.redirect("/dashboard");
  }
});

function validate(b) {
  const nama = String(b.nama || "").trim();
  const kategori = String(b.kategori || "");
  const kondisi = String(b.kondisi || "");
  const satuan = String(b.satuan || "unit").trim() || "unit";
  const lokasi = String(b.lokasi || "").trim();
  const sumber = String(b.sumber || "").trim();
  const catatan = String(b.catatan || "").trim();
  const perolehan = String(b.tanggal_perolehan || "").trim();
  const jumlah = Number(b.jumlah);
  if (!nama || nama.length > 120) return { error: "Nama barang wajib diisi (maks. 120 karakter)" };
  if (!KATEGORI[kategori]) return { error: "Kategori tidak valid" };
  if (!KONDISI[kondisi]) return { error: "Kondisi tidak valid" };
  if (!Number.isInteger(jumlah) || jumlah < 0 || jumlah > 1000000) return { error: "Jumlah harus berupa angka bulat 0 atau lebih" };
  if (satuan.length > 20) return { error: "Satuan maksimal 20 karakter" };
  if (lokasi.length > 100 || sumber.length > 100 || catatan.length > 255) return { error: "Lokasi/sumber/catatan terlalu panjang" };
  if (perolehan && !validIso(perolehan)) return { error: "Tanggal perolehan tidak valid" };
  return { value: { nama, kategori, kondisi, satuan, jumlah, lokasi: lokasi || null, sumber: sumber || null, catatan: catatan || null, perolehan: perolehan || null } };
}
const BACK = (req) => safeReturn(req.body.return_to, "/inventaris");

router.get("/form-data", async (req, res) => {
  try {
    const [rows] = await req.app.locals.db.execute(
      "SELECT id, nama, kategori, jumlah, satuan, kondisi, lokasi, DATE_FORMAT(tanggal_perolehan,'%Y-%m-%d') AS tanggal_perolehan, sumber, catatan FROM inventaris WHERE id = ?",
      [req.query.id]
    );
    if (!rows.length) return res.status(404).json({ ok: false, message: "Barang tidak ditemukan" });
    res.json({ ok: true, barang: rows[0] });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, message: "Gagal memuat data" });
  }
});

router.post("/", async (req, res) => {
  const { value: v, error } = validate(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: BACK(req) });
  try {
    await req.app.locals.db.execute(
      "INSERT INTO inventaris (nama, kategori, jumlah, satuan, kondisi, lokasi, tanggal_perolehan, sumber, catatan) VALUES (?,?,?,?,?,?,?,?,?)",
      [v.nama, v.kategori, v.jumlah, v.satuan, v.kondisi, v.lokasi, v.perolehan, v.sumber, v.catatan]
    );
    reply(req, res, { ok: true, message: `“${v.nama}” ditambahkan ke inventaris`, redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal menyimpan barang", redirect: BACK(req), status: 500 });
  }
});

router.post("/:id/edit", async (req, res) => {
  const { value: v, error } = validate(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: BACK(req) });
  try {
    const [r] = await req.app.locals.db.execute(
      "UPDATE inventaris SET nama=?, kategori=?, jumlah=?, satuan=?, kondisi=?, lokasi=?, tanggal_perolehan=?, sumber=?, catatan=? WHERE id=?",
      [v.nama, v.kategori, v.jumlah, v.satuan, v.kondisi, v.lokasi, v.perolehan, v.sumber, v.catatan, req.params.id]
    );
    if (!r.affectedRows) return reply(req, res, { ok: false, message: "Barang tidak ditemukan", redirect: BACK(req), status: 404 });
    reply(req, res, { ok: true, message: "Barang diperbarui", redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal memperbarui barang", redirect: BACK(req), status: 500 });
  }
});

router.post("/:id/delete", async (req, res) => {
  try {
    await req.app.locals.db.execute("DELETE FROM inventaris WHERE id = ?", [req.params.id]);
    reply(req, res, { ok: true, message: "Barang dihapus", redirect: BACK(req) });
  } catch (e) {
    console.error(e);
    reply(req, res, { ok: false, message: "Gagal menghapus barang", redirect: BACK(req), status: 500 });
  }
});

module.exports = { router };
