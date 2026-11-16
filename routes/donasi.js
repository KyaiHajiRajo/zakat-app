const express = require("express");
const multer = require("multer");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { getSettings, saveSettings, lokasi } = require("../utils/settings");
const { createLimiter } = require("../utils/rate-limit");
const { normalizeWa, displayWa } = require("../utils/phone");
const { reply } = require("../utils/respond");
const { logAktivitas } = require("../utils/audit");
const { labelTanggal, todayIso, dateToIso } = require("../utils/dates");
const { KATEGORI } = require("./kas");

const TUJUAN = {
  donasi: { label: "Donasi Umum", icon: "fa-hand-holding-heart", desc: "Operasional & kegiatan musholla" },
  infaq_ramadan: { label: "Infaq Ramadan", icon: "fa-moon", desc: "Takjil, buka bersama & tarawih" },
  wakaf: { label: "Wakaf", icon: "fa-mosque", desc: "Pembangunan & perawatan" },
};
const NOMINAL_CEPAT = [20000, 50000, 100000, 250000, 500000];
const MIN_DONASI = 1000;
const MAX_DONASI = 1000000000;

const rupiah = (n) => "Rp " + new Intl.NumberFormat("id-ID").format(Math.round(Number(n) || 0));
const IMAGE_TYPES = { "image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp" };

// Pastikan isi berkas benar-benar gambar (bukan hanya nama/tipe yang dikirim)
function sniffImage(buf) {
  if (buf.length < 12) return null;
  if (buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return ".jpg";
  if (buf.slice(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) return ".png";
  if (buf.slice(0, 4).toString("ascii") === "RIFF" && buf.slice(8, 12).toString("ascii") === "WEBP") return ".webp";
  return null;
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 4 * 1024 * 1024, files: 1, fields: 20 },
  fileFilter: (req, file, cb) => (IMAGE_TYPES[file.mimetype] ? cb(null, true) : cb(new Error("Bukti harus berupa gambar JPG, PNG, atau WebP"))),
});

function runUpload(field) {
  return (req, res, next) =>
    upload.single(field)(req, res, (error) => {
      if (!error) return next();
      const message = error.code === "LIMIT_FILE_SIZE" ? "Ukuran gambar maksimal 4 MB" : error.message || "Gagal mengunggah gambar";
      return res.status(400).json({ ok: false, message });
    });
}

function saveImage(dir, prefix, file) {
  const ext = sniffImage(file.buffer);
  if (!ext) return null;
  fs.mkdirSync(dir, { recursive: true });
  const name = `${prefix}-${Date.now()}-${crypto.randomBytes(8).toString("hex")}${ext}`;
  fs.writeFileSync(path.join(dir, name), file.buffer, { mode: 0o644 });
  return name;
}

const BUKTI_DIR = path.join(__dirname, "..", "public", "uploads", "bukti-donasi");
const QRIS_DIR = path.join(__dirname, "..", "public", "qris");

const KODE_CHARS = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";
function buatKode() {
  let out = "DN-";
  const bytes = crypto.randomBytes(8);
  for (let i = 0; i < 8; i++) out += KODE_CHARS[bytes[i] % KODE_CHARS.length];
  return out;
}

function waktuLalu(date) {
  const diff = Math.max(0, Date.now() - new Date(date).getTime());
  const m = Math.floor(diff / 60000);
  if (m < 1) return "baru saja";
  if (m < 60) return `${m} menit lalu`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h} jam lalu`;
  const d = Math.floor(h / 24);
  if (d < 30) return `${d} hari lalu`;
  return new Date(date).toLocaleDateString("id-ID", { day: "numeric", month: "short", year: "numeric" });
}

// Jumlah donasi yang menunggu (untuk lencana menu), disimpan 15 detik agar ringan
let pendingCache = { at: 0, n: 0 };
async function pendingCount(db) {
  if (Date.now() - pendingCache.at < 15000) return pendingCache.n;
  const [[row]] = await db.execute("SELECT COUNT(*) AS n FROM donasi WHERE status = 'menunggu'");
  pendingCache = { at: Date.now(), n: Number(row.n) };
  return pendingCache.n;
}
const resetPending = () => { pendingCache.at = 0; };

function paymentReady(s) {
  return !!(s.donasi_qris || (s.donasi_bank_nama && s.donasi_bank_norek));
}

// ---------------------------------------------------------------- Halaman jamaah (tanpa login)
const publicRouter = express.Router();
const submitLimiter = createLimiter({ windowMs: 60 * 60 * 1000, max: 8 });

publicRouter.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const settings = await getSettings(db);
    const [terbaru] = await db.execute(
      "SELECT nama, anonim, jumlah, tujuan, created_at, diproses_at FROM donasi WHERE status = 'diterima' ORDER BY diproses_at DESC, id DESC LIMIT 6"
    );
    const [[total]] = await db.execute("SELECT COUNT(*) AS n, COALESCE(SUM(jumlah),0) AS total FROM donasi WHERE status = 'diterima'");
    res.render("donasi/public", {
      title: "Donasi & Infaq - " + settings.masjid_nama,
      user: req.session.user || null,
      settings,
      aktif: settings.donasi_aktif === "1" && paymentReady(settings),
      tujuan: TUJUAN,
      nominalCepat: NOMINAL_CEPAT,
      min: MIN_DONASI,
      waPengurus: normalizeWa(settings.donasi_wa),
      terbaru: terbaru.map((d) => ({ nama: d.anonim ? "Hamba Allah" : d.nama, jumlah: Number(d.jumlah), tujuan: TUJUAN[d.tujuan] ? TUJUAN[d.tujuan].label : "Donasi", lalu: waktuLalu(d.diproses_at || d.created_at) })),
      jumlahDonatur: Number(total.n),
      totalDonasi: Number(total.total),
      rupiah,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error halaman donasi:", error);
    res.status(500).render("error", { title: "Server Error", message: "Gagal memuat halaman donasi", error: {} });
  }
});

publicRouter.post("/kirim", runUpload("bukti"), async (req, res) => {
  const db = req.app.locals.db;
  const fail = (status, message) => res.status(status).json({ ok: false, message });
  const ipKey = req.ip || "-";
  try {
    const blocked = submitLimiter.blockedFor(ipKey);
    if (blocked) return fail(429, `Terlalu banyak pengiriman. Coba lagi ${Math.ceil(blocked / 60)} menit lagi.`);

    const settings = await getSettings(db);
    if (settings.donasi_aktif !== "1" || !paymentReady(settings)) return fail(403, "Donasi online sedang ditutup.");

    const b = req.body || {};
    if (b.website) return res.json({ ok: true, redirect: "/donasi" }); // kolom jebakan untuk bot

    const nama = String(b.nama || "").trim().replace(/\s+/g, " ");
    const anonim = b.anonim === "1" || b.anonim === "on" ? 1 : 0;
    const pesan = String(b.pesan || "").trim().replace(/\s+/g, " ");
    const tujuan = String(b.tujuan || "");
    const jumlah = Math.round(Number(String(b.jumlah || "").replace(/[^\d]/g, "")));
    const nomorMentah = String(b.nomor_wa || "").trim();
    const nomor = nomorMentah ? normalizeWa(nomorMentah) : null;

    if (!nama || nama.length > 100) return fail(400, "Nama wajib diisi (maksimal 100 karakter)");
    if (!TUJUAN[tujuan]) return fail(400, "Pilih tujuan donasi");
    if (!(jumlah >= MIN_DONASI) || jumlah > MAX_DONASI) return fail(400, `Nominal minimal ${rupiah(MIN_DONASI)}`);
    if (nomorMentah && !nomor) return fail(400, "Nomor WhatsApp tidak valid (contoh: 0812 3456 7890)");
    if (pesan.length > 255) return fail(400, "Pesan maksimal 255 karakter");
    if (!req.file) return fail(400, "Unggah foto bukti transfer");

    const berkas = saveImage(BUKTI_DIR, "donasi", req.file);
    if (!berkas) return fail(400, "Berkas bukan gambar yang valid (JPG, PNG, atau WebP)");

    submitLimiter.fail(ipKey);
    let kode, inserted = false;
    for (let i = 0; i < 4 && !inserted; i++) {
      kode = buatKode();
      try {
        await db.execute(
          "INSERT INTO donasi (kode, nama, anonim, nomor_wa, jumlah, tujuan, pesan, bukti, ip) VALUES (?,?,?,?,?,?,?,?,?)",
          [kode, nama, anonim, nomor, jumlah, tujuan, pesan || null, berkas, ipKey.slice(0, 64)]
        );
        inserted = true;
      } catch (error) {
        if (error.code !== "ER_DUP_ENTRY") throw error;
      }
    }
    if (!inserted) return fail(500, "Gagal membuat kode donasi, coba lagi");
    resetPending();
    logAktivitas(db, { user: null, aksi: "Menambah", modul: "Donasi", ringkasan: `Konfirmasi donasi baru ${kode}: ${anonim ? "Hamba Allah" : nama} · ${rupiah(jumlah)} · ${TUJUAN[tujuan].label}`, ip: req.ip });
    res.json({ ok: true, redirect: `/donasi/status/${kode}?baru=1` });
  } catch (error) {
    console.error("Error kirim donasi:", error);
    fail(500, "Terjadi kesalahan, coba lagi sebentar lagi");
  }
});

publicRouter.get("/cek", (req, res) => {
  const kode = String(req.query.kode || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!/^DN-[A-Z0-9]{8}$/.test(kode)) {
    req.flash("error", "Kode tidak valid. Contoh: DN-AB12CD34");
    return res.redirect("/donasi");
  }
  res.redirect(`/donasi/status/${kode}`);
});

publicRouter.get("/status/:kode", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const kode = String(req.params.kode || "").toUpperCase();
    const settings = await getSettings(db);
    const [rows] = await db.execute(
      "SELECT kode, nama, anonim, jumlah, tujuan, pesan, status, alasan_tolak, created_at, diproses_at FROM donasi WHERE kode = ?",
      [kode]
    );
    res.status(rows.length ? 200 : 404).render("donasi/status", {
      title: "Status Donasi - " + settings.masjid_nama,
      user: req.session.user || null,
      settings,
      d: rows.length ? { ...rows[0], jumlah: Number(rows[0].jumlah), tujuanLabel: (TUJUAN[rows[0].tujuan] || TUJUAN.donasi).label } : null,
      baru: req.query.baru === "1",
      waPengurus: normalizeWa(settings.donasi_wa),
      rupiah,
      waktuLalu,
      success: [],
      error: [],
    });
  } catch (error) {
    console.error("Error status donasi:", error);
    res.status(500).render("error", { title: "Server Error", message: "Gagal memuat status donasi", error: {} });
  }
});

// ---------------------------------------------------------------- Halaman pengurus (login)
const router = express.Router();
const FILTERS = ["menunggu", "diterima", "ditolak", "semua"];
const PER_PAGE = 12;

router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const settings = await getSettings(db);
    const [[hitung]] = await db.execute(
      `SELECT SUM(status='menunggu') AS menunggu, SUM(status='diterima') AS diterima, SUM(status='ditolak') AS ditolak,
              COALESCE(SUM(CASE WHEN status='diterima' THEN jumlah END),0) AS total_diterima,
              COALESCE(SUM(CASE WHEN status='menunggu' THEN jumlah END),0) AS total_menunggu,
              COUNT(*) AS semua FROM donasi`
    );
    const counts = { menunggu: Number(hitung.menunggu || 0), diterima: Number(hitung.diterima || 0), ditolak: Number(hitung.ditolak || 0), semua: Number(hitung.semua || 0) };
    let filter = String(req.query.status || "");
    if (!FILTERS.includes(filter)) filter = counts.menunggu ? "menunggu" : "semua";
    const page = Math.max(1, parseInt(req.query.page, 10) || 1);
    const where = filter === "semua" ? "" : "WHERE status = ?";
    const params = filter === "semua" ? [] : [filter];
    const total = filter === "semua" ? counts.semua : counts[filter];
    const totalPages = Math.max(1, Math.ceil(total / PER_PAGE));
    const cur = Math.min(page, totalPages);
    const [rows] = await db.query(
      `SELECT d.*, u.name AS oleh FROM donasi d LEFT JOIN users u ON u.id = d.diproses_oleh ${where}
        ORDER BY (d.status='menunggu') DESC, d.created_at DESC, d.id DESC LIMIT ${PER_PAGE} OFFSET ${(cur - 1) * PER_PAGE}`,
      params
    );
    res.render("donasi/admin", {
      title: "Donasi Masuk - " + settings.masjid_nama,
      user: req.session.user,
      currentPage: "donasi",
      isAdmin: req.session.user.role === "admin",
      settings,
      siap: paymentReady(settings),
      counts,
      totalDiterima: Number(hitung.total_diterima),
      totalMenunggu: Number(hitung.total_menunggu),
      filter,
      page: cur,
      totalPages,
      total,
      rows: rows.map((r) => ({
        ...r,
        jumlah: Number(r.jumlah),
        tujuanLabel: (TUJUAN[r.tujuan] || TUJUAN.donasi).label,
        tujuanIcon: (TUJUAN[r.tujuan] || TUJUAN.donasi).icon,
        lalu: waktuLalu(r.created_at),
        wa: r.nomor_wa,
        waTampil: displayWa(r.nomor_wa),
      })),
      TUJUAN,
      KATEGORI_MASUK: KATEGORI.masuk,
      hariIni: todayIso(lokasi(settings).tz),
      rupiah,
      labelTanggal,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error donasi masuk:", error);
    req.flash("error", "Gagal memuat donasi masuk");
    res.redirect("/dashboard");
  }
});

router.post("/:id/terima", async (req, res) => {
  const db = req.app.locals.db;
  const back = "/donasi-masuk";
  const id = parseInt(req.params.id, 10);
  const tanggal = String(req.body.tanggal || "").trim();
  const kategori = String(req.body.kategori || "");
  const { validIso } = require("../utils/dates");
  if (!validIso(tanggal)) return reply(req, res, { ok: false, message: "Tanggal tidak valid", redirect: back });
  if (!KATEGORI.masuk[kategori]) return reply(req, res, { ok: false, message: "Kategori kas tidak valid", redirect: back });

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();
    const [[d]] = await conn.execute("SELECT * FROM donasi WHERE id = ? FOR UPDATE", [id]);
    if (!d) { await conn.rollback(); return reply(req, res, { ok: false, message: "Donasi tidak ditemukan", redirect: back, status: 404 }); }
    if (d.status !== "menunggu") { await conn.rollback(); return reply(req, res, { ok: false, message: "Donasi ini sudah diproses sebelumnya", redirect: back }); }
    const pihak = d.anonim ? "Hamba Allah" : d.nama;
    const [kas] = await conn.execute(
      "INSERT INTO kas_transaksi (tanggal, tipe, kategori, jumlah, pihak, keterangan, user_id) VALUES (?, 'masuk', ?, ?, ?, ?, ?)",
      [tanggal, kategori, d.jumlah, pihak.slice(0, 120), `Donasi online ${d.kode}`, req.session.user.id]
    );
    await conn.execute(
      "UPDATE donasi SET status='diterima', kas_id=?, diproses_oleh=?, diproses_at=NOW() WHERE id=?",
      [kas.insertId, req.session.user.id, id]
    );
    await conn.commit();
    resetPending();
    logAktivitas(db, { user: req.session.user, aksi: "Menerima", modul: "Donasi", objekId: id, ringkasan: `Menerima donasi ${d.kode}: ${pihak} · ${rupiah(d.jumlah)} → kas (${KATEGORI.masuk[kategori].label})`, ip: req.ip });
    reply(req, res, { ok: true, message: `Donasi ${rupiah(d.jumlah)} dari ${pihak} diterima dan dicatat ke kas`, redirect: back + (req.body.filter ? "?status=" + encodeURIComponent(String(req.body.filter).slice(0, 10)) : "") });
  } catch (error) {
    await conn.rollback().catch(() => {});
    console.error("Error terima donasi:", error);
    reply(req, res, { ok: false, message: "Gagal memproses donasi", redirect: back, status: 500 });
  } finally {
    conn.release();
  }
});

router.post("/:id/tolak", async (req, res) => {
  const db = req.app.locals.db;
  const back = "/donasi-masuk";
  const id = parseInt(req.params.id, 10);
  const alasan = String(req.body.alasan || "").trim().replace(/\s+/g, " ");
  if (!alasan) return reply(req, res, { ok: false, message: "Tulis alasan penolakan agar donatur paham", redirect: back });
  if (alasan.length > 255) return reply(req, res, { ok: false, message: "Alasan maksimal 255 karakter", redirect: back });
  try {
    const [[d]] = await db.execute("SELECT kode, nama, anonim, jumlah, status FROM donasi WHERE id = ?", [id]);
    if (!d) return reply(req, res, { ok: false, message: "Donasi tidak ditemukan", redirect: back, status: 404 });
    const [r] = await db.execute(
      "UPDATE donasi SET status='ditolak', alasan_tolak=?, diproses_oleh=?, diproses_at=NOW() WHERE id=? AND status='menunggu'",
      [alasan, req.session.user.id, id]
    );
    if (!r.affectedRows) return reply(req, res, { ok: false, message: "Donasi ini sudah diproses sebelumnya", redirect: back });
    resetPending();
    logAktivitas(db, { user: req.session.user, aksi: "Menolak", modul: "Donasi", objekId: id, ringkasan: `Menolak donasi ${d.kode}: ${d.anonim ? "Hamba Allah" : d.nama} · ${rupiah(d.jumlah)} (${alasan.slice(0, 80)})`, ip: req.ip });
    reply(req, res, { ok: true, message: "Donasi ditolak, alasan akan terlihat oleh donatur", redirect: back });
  } catch (error) {
    console.error("Error tolak donasi:", error);
    reply(req, res, { ok: false, message: "Gagal memproses donasi", redirect: back, status: 500 });
  }
});

function adminOnly(req, res, next) {
  if (req.session.user && req.session.user.role === "admin") return next();
  return res.status(403).json({ ok: false, message: "Hanya admin yang dapat mengubah pengaturan donasi" });
}

router.post("/pengaturan", adminOnly, async (req, res) => {
  const back = "/donasi-masuk";
  const b = req.body || {};
  const val = (k, max) => String(b[k] || "").trim().replace(/\s+/g, " ").slice(0, max);
  const norek = val("donasi_bank_norek", 40).replace(/[^\d\- ]/g, "");
  const waMentah = val("donasi_wa", 20);
  const wa = waMentah ? normalizeWa(waMentah) : "";
  if (waMentah && !wa) return reply(req, res, { ok: false, message: "Nomor WhatsApp pengurus tidak valid", redirect: back });
  const values = {
    donasi_aktif: b.donasi_aktif === "1" ? "1" : "0",
    donasi_bank_nama: val("donasi_bank_nama", 60),
    donasi_bank_norek: norek,
    donasi_bank_atas_nama: val("donasi_bank_atas_nama", 80),
    donasi_wa: wa,
    donasi_catatan: val("donasi_catatan", 255),
    transparansi_aktif: b.transparansi_aktif === "1" ? "1" : "0",
    transparansi_nama_donatur: b.transparansi_nama_donatur === "1" ? "1" : "0",
  };
  if ((values.donasi_bank_nama || values.donasi_bank_norek) && !(values.donasi_bank_nama && values.donasi_bank_norek)) {
    return reply(req, res, { ok: false, message: "Isi nama bank dan nomor rekening sekaligus", redirect: back });
  }
  try {
    await saveSettings(req.app.locals.db, values);
    logAktivitas(req.app.locals.db, { user: req.session.user, aksi: "Mengubah", modul: "Donasi", ringkasan: "Mengubah pengaturan donasi & transparansi publik", ip: req.ip });
    reply(req, res, { ok: true, message: "Pengaturan donasi disimpan", redirect: back });
  } catch (error) {
    console.error("Error pengaturan donasi:", error);
    reply(req, res, { ok: false, message: "Gagal menyimpan pengaturan", redirect: back, status: 500 });
  }
});

router.post("/qris", adminOnly, runUpload("qris"), async (req, res) => {
  const db = req.app.locals.db;
  try {
    if (!req.file) return res.status(400).json({ ok: false, message: "Pilih gambar QRIS terlebih dahulu" });
    const name = saveImage(QRIS_DIR, "qris", req.file);
    if (!name) return res.status(400).json({ ok: false, message: "Berkas bukan gambar yang valid (JPG, PNG, atau WebP)" });
    const settings = await getSettings(db);
    await saveSettings(db, { donasi_qris: name });
    if (settings.donasi_qris) fs.unlink(path.join(QRIS_DIR, path.basename(settings.donasi_qris)), () => {});
    logAktivitas(db, { user: req.session.user, aksi: "Mengubah", modul: "Donasi", ringkasan: "Mengunggah gambar QRIS donasi", ip: req.ip });
    res.json({ ok: true, url: "/qris/" + name });
  } catch (error) {
    console.error("Error unggah QRIS:", error);
    res.status(500).json({ ok: false, message: "Gagal menyimpan gambar QRIS" });
  }
});

router.post("/qris/hapus", adminOnly, async (req, res) => {
  const db = req.app.locals.db;
  try {
    const settings = await getSettings(db);
    await saveSettings(db, { donasi_qris: "" });
    if (settings.donasi_qris) fs.unlink(path.join(QRIS_DIR, path.basename(settings.donasi_qris)), () => {});
    logAktivitas(db, { user: req.session.user, aksi: "Menghapus", modul: "Donasi", ringkasan: "Menghapus gambar QRIS donasi", ip: req.ip });
    res.json({ ok: true });
  } catch (error) {
    console.error("Error hapus QRIS:", error);
    res.status(500).json({ ok: false, message: "Gagal menghapus QRIS" });
  }
});

module.exports = { router, publicRouter, pendingCount, TUJUAN };
