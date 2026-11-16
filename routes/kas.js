const express = require("express");
const { getSettings, lokasi } = require("../utils/settings");
const { reply, safeReturn } = require("../utils/respond");
const { validIso, labelTanggal, dateToIso, bulanRange, todayIso, BULAN } = require("../utils/dates");

const KATEGORI = {
  masuk: {
    kotak_amal: { label: "Kotak Amal", icon: "fa-box-archive" },
    donasi: { label: "Donasi / Sumbangan", icon: "fa-hand-holding-heart" },
    infaq_ramadan: { label: "Infaq Ramadan", icon: "fa-moon" },
    wakaf: { label: "Wakaf", icon: "fa-mosque" },
    lainnya: { label: "Lainnya", icon: "fa-circle-plus" },
  },
  keluar: {
    listrik_air: { label: "Listrik & Air", icon: "fa-bolt" },
    kebersihan: { label: "Kebersihan", icon: "fa-broom" },
    perawatan: { label: "Perawatan & Perbaikan", icon: "fa-screwdriver-wrench" },
    perlengkapan: { label: "Perlengkapan", icon: "fa-box-open" },
    konsumsi: { label: "Konsumsi & Takjil", icon: "fa-utensils" },
    honor: { label: "Honor / Insentif", icon: "fa-user-tie" },
    kegiatan: { label: "Kegiatan & Acara", icon: "fa-calendar-check" },
    sosial: { label: "Bantuan Sosial", icon: "fa-people-carry-box" },
    lainnya: { label: "Lainnya", icon: "fa-circle-minus" },
  },
};
const PER_PAGE = 20;

const kategoriLabel = (tipe, key) => (KATEGORI[tipe] && KATEGORI[tipe][key] ? KATEGORI[tipe][key].label : key);
const rupiah = (n) => "Rp " + new Intl.NumberFormat("id-ID").format(Math.round(Number(n) || 0));

async function currentMonth(req) {
  const settings = await getSettings(req.app.locals.db);
  return { settings, ym: todayIso(lokasi(settings).tz).slice(0, 7) };
}

// WHERE dasar untuk rentang bulan + filter
function buildFilter({ range, tipe, kategori, q }) {
  const where = [];
  const params = [];
  if (range) {
    where.push("k.tanggal BETWEEN ? AND ?");
    params.push(range.awal, range.akhir);
  }
  if (tipe === "masuk" || tipe === "keluar") {
    where.push("k.tipe = ?");
    params.push(tipe);
  }
  if (kategori && /^[a-z_]+$/.test(kategori)) {
    where.push("k.kategori = ?");
    params.push(kategori);
  }
  if (q) {
    where.push("(k.pihak LIKE ? OR k.keterangan LIKE ?)");
    params.push(`%${q}%`, `%${q}%`);
  }
  return { sql: where.length ? "WHERE " + where.join(" AND ") : "", params };
}

function parseQuery(req, ym) {
  const bulanParam = String(req.query.bulan || "");
  const semua = bulanParam === "semua";
  const bulan = semua ? "semua" : bulanRange(bulanParam) ? bulanParam : ym;
  return {
    bulan,
    range: semua ? null : bulanRange(bulan),
    tipe: ["masuk", "keluar"].includes(req.query.tipe) ? req.query.tipe : "",
    kategori: String(req.query.kategori || "").slice(0, 40),
    q: String(req.query.q || "").trim().slice(0, 80),
    page: Math.max(1, parseInt(req.query.page, 10) || 1),
  };
}

const router = express.Router();

router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { settings, ym } = await currentMonth(req);
    const f = parseQuery(req, ym);

    const [[saldoRow]] = await db.execute(
      "SELECT COALESCE(SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE -jumlah END),0) AS saldo FROM kas_transaksi"
    );
    const rangeFilter = buildFilter({ range: f.range });
    const [[ringkas]] = await db.execute(
      `SELECT COALESCE(SUM(CASE WHEN k.tipe='masuk' THEN k.jumlah END),0) AS masuk,
              COALESCE(SUM(CASE WHEN k.tipe='keluar' THEN k.jumlah END),0) AS keluar,
              COUNT(*) AS total
         FROM kas_transaksi k ${rangeFilter.sql}`,
      rangeFilter.params
    );
    let saldoAwal = 0;
    if (f.range) {
      const [[row]] = await db.execute(
        "SELECT COALESCE(SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE -jumlah END),0) AS saldo FROM kas_transaksi WHERE tanggal < ?",
        [f.range.awal]
      );
      saldoAwal = Number(row.saldo);
    }

    const [breakdown] = await db.execute(
      `SELECT k.tipe, k.kategori, SUM(k.jumlah) AS total, COUNT(*) AS n
         FROM kas_transaksi k ${rangeFilter.sql} GROUP BY k.tipe, k.kategori ORDER BY total DESC`,
      rangeFilter.params
    );

    // grafik 6 bulan terakhir (berakhir di bulan terpilih / bulan ini)
    const endYm = f.range ? f.bulan : ym;
    const [ey, em] = endYm.split("-").map(Number);
    const months = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(Date.UTC(ey, em - 1 - i, 1));
      months.push(d.toISOString().slice(0, 7));
    }
    const [series] = await db.execute(
      `SELECT DATE_FORMAT(tanggal,'%Y-%m') AS ym,
              SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE 0 END) AS masuk,
              SUM(CASE WHEN tipe='keluar' THEN jumlah ELSE 0 END) AS keluar
         FROM kas_transaksi WHERE tanggal BETWEEN ? AND ? GROUP BY ym`,
      [`${months[0]}-01`, bulanRange(months[5]).akhir]
    );
    const chart = months.map((m) => {
      const row = series.find((s) => s.ym === m);
      const [y, mo] = m.split("-").map(Number);
      return { ym: m, label: BULAN[mo - 1].slice(0, 3) + " " + String(y).slice(2), masuk: Number(row ? row.masuk : 0), keluar: Number(row ? row.keluar : 0) };
    });

    const listFilter = buildFilter(f);
    const [[{ total }]] = await db.execute(`SELECT COUNT(*) AS total FROM kas_transaksi k ${listFilter.sql}`, listFilter.params);
    const totalPages = Math.max(1, Math.ceil(total / PER_PAGE));
    const page = Math.min(f.page, totalPages);
    const [rows] = await db.query(
      `SELECT k.id, k.tanggal, k.tipe, k.kategori, k.jumlah, k.pihak, k.keterangan, u.name AS oleh
         FROM kas_transaksi k LEFT JOIN users u ON u.id = k.user_id ${listFilter.sql}
        ORDER BY k.tanggal DESC, k.id DESC LIMIT ${PER_PAGE} OFFSET ${(page - 1) * PER_PAGE}`,
      listFilter.params
    );

    const [monthRows] = await db.execute("SELECT DISTINCT DATE_FORMAT(tanggal,'%Y-%m') AS ym FROM kas_transaksi ORDER BY ym DESC LIMIT 36");
    const monthSet = new Set(monthRows.map((r) => r.ym));
    monthSet.add(ym);
    if (f.range) monthSet.add(f.bulan);
    const bulanList = [...monthSet].sort().reverse().map((m) => ({ ym: m, label: bulanRange(m).label }));

    res.render("kas/index", {
      title: "Kas Masjid - " + settings.masjid_nama,
      user: req.session.user,
      isAdmin: req.session.user.role === "admin",
      settings,
      f,
      periodeLabel: f.range ? f.range.label : "Semua periode",
      saldo: Number(saldoRow.saldo),
      saldoAwal,
      masuk: Number(ringkas.masuk),
      keluar: Number(ringkas.keluar),
      jumlahTransaksi: Number(ringkas.total),
      breakdown: breakdown.map((b) => ({ ...b, total: Number(b.total), label: kategoriLabel(b.tipe, b.kategori), icon: (KATEGORI[b.tipe][b.kategori] || {}).icon || "fa-circle" })),
      chart,
      rows: rows.map((r) => ({ ...r, tanggal: dateToIso(r.tanggal), jumlah: Number(r.jumlah), kategoriLabel: kategoriLabel(r.tipe, r.kategori), icon: (KATEGORI[r.tipe][r.kategori] || {}).icon || "fa-circle" })),
      total,
      page,
      totalPages,
      perPage: PER_PAGE,
      bulanList,
      KATEGORI,
      labelTanggal,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error kas:", error);
    req.flash("error", "Gagal memuat kas masjid");
    res.redirect("/dashboard");
  }
});

function validate(body) {
  const tanggal = String(body.tanggal || "").trim();
  const tipe = String(body.tipe || "");
  const kategori = String(body.kategori || "");
  const pihak = String(body.pihak || "").trim();
  const keterangan = String(body.keterangan || "").trim();
  const jumlah = Number(String(body.jumlah || "").replace(",", "."));
  if (!validIso(tanggal)) return { error: "Tanggal tidak valid" };
  if (!["masuk", "keluar"].includes(tipe)) return { error: "Jenis transaksi tidak valid" };
  if (!KATEGORI[tipe][kategori]) return { error: "Kategori tidak valid" };
  if (!(jumlah > 0) || jumlah > 99999999999) return { error: "Jumlah harus lebih dari 0" };
  if (pihak.length > 120) return { error: "Nama pihak maksimal 120 karakter" };
  if (keterangan.length > 255) return { error: "Keterangan maksimal 255 karakter" };
  return { value: { tanggal, tipe, kategori, jumlah: Math.round(jumlah * 100) / 100, pihak: pihak || null, keterangan: keterangan || null } };
}

router.get("/form-data", async (req, res) => {
  try {
    const [rows] = await req.app.locals.db.execute(
      "SELECT id, DATE_FORMAT(tanggal,'%Y-%m-%d') AS tanggal, tipe, kategori, jumlah, pihak, keterangan FROM kas_transaksi WHERE id = ?",
      [req.query.id]
    );
    if (!rows.length) return res.status(404).json({ ok: false, message: "Transaksi tidak ditemukan" });
    res.json({ ok: true, kas: { ...rows[0], jumlah: Number(rows[0].jumlah) } });
  } catch (error) {
    console.error("Error form kas:", error);
    res.status(500).json({ ok: false, message: "Gagal memuat data" });
  }
});

router.post("/", async (req, res) => {
  const back = safeReturn(req.body.return_to, "/kas");
  const { value, error } = validate(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: back });
  try {
    await req.app.locals.db.execute(
      "INSERT INTO kas_transaksi (tanggal, tipe, kategori, jumlah, pihak, keterangan, user_id) VALUES (?, ?, ?, ?, ?, ?, ?)",
      [value.tanggal, value.tipe, value.kategori, value.jumlah, value.pihak, value.keterangan, req.session.user.id]
    );
    reply(req, res, { ok: true, message: `${value.tipe === "masuk" ? "Pemasukan" : "Pengeluaran"} ${rupiah(value.jumlah)} dicatat`, redirect: back });
  } catch (err) {
    console.error("Error simpan kas:", err);
    reply(req, res, { ok: false, message: "Gagal menyimpan transaksi", redirect: back, status: 500 });
  }
});

router.post("/:id/edit", async (req, res) => {
  const back = safeReturn(req.body.return_to, "/kas");
  const { value, error } = validate(req.body);
  if (error) return reply(req, res, { ok: false, message: error, redirect: back });
  try {
    const [result] = await req.app.locals.db.execute(
      "UPDATE kas_transaksi SET tanggal=?, tipe=?, kategori=?, jumlah=?, pihak=?, keterangan=? WHERE id=?",
      [value.tanggal, value.tipe, value.kategori, value.jumlah, value.pihak, value.keterangan, req.params.id]
    );
    if (!result.affectedRows) return reply(req, res, { ok: false, message: "Transaksi tidak ditemukan", redirect: back, status: 404 });
    reply(req, res, { ok: true, message: "Transaksi diperbarui", redirect: back });
  } catch (err) {
    console.error("Error ubah kas:", err);
    reply(req, res, { ok: false, message: "Gagal memperbarui transaksi", redirect: back, status: 500 });
  }
});

// Hanya admin yang boleh menghapus catatan keuangan
router.post("/:id/delete", async (req, res) => {
  const back = safeReturn(req.body.return_to, "/kas");
  if (req.session.user.role !== "admin") {
    return reply(req, res, { ok: false, message: "Hanya admin yang dapat menghapus transaksi kas", redirect: back, status: 403 });
  }
  try {
    await req.app.locals.db.execute("DELETE FROM kas_transaksi WHERE id = ?", [req.params.id]);
    reply(req, res, { ok: true, message: "Transaksi dihapus", redirect: back });
  } catch (err) {
    console.error("Error hapus kas:", err);
    reply(req, res, { ok: false, message: "Gagal menghapus transaksi", redirect: back, status: 500 });
  }
});

// Laporan bulanan siap cetak
router.get("/laporan", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { settings, ym } = await currentMonth(req);
    const range = bulanRange(String(req.query.bulan || "")) || bulanRange(ym);
    const bulan = range.awal.slice(0, 7);

    const [[awalRow]] = await db.execute(
      "SELECT COALESCE(SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE -jumlah END),0) AS saldo FROM kas_transaksi WHERE tanggal < ?",
      [range.awal]
    );
    const [rows] = await db.execute(
      `SELECT k.id, k.tanggal, k.tipe, k.kategori, k.jumlah, k.pihak, k.keterangan
         FROM kas_transaksi k WHERE k.tanggal BETWEEN ? AND ? ORDER BY k.tanggal ASC, k.id ASC`,
      [range.awal, range.akhir]
    );
    const transaksi = rows.map((r) => ({ ...r, tanggal: dateToIso(r.tanggal), jumlah: Number(r.jumlah), kategoriLabel: kategoriLabel(r.tipe, r.kategori) }));
    const sum = (tipe) => transaksi.filter((t) => t.tipe === tipe).reduce((a, t) => a + t.jumlah, 0);
    const group = (tipe) => {
      const map = new Map();
      transaksi.filter((t) => t.tipe === tipe).forEach((t) => map.set(t.kategoriLabel, (map.get(t.kategoriLabel) || 0) + t.jumlah));
      return [...map.entries()].map(([label, total]) => ({ label, total })).sort((a, b) => b.total - a.total);
    };
    const saldoAwal = Number(awalRow.saldo);
    const masuk = sum("masuk");
    const keluar = sum("keluar");

    res.render("kas/laporan", {
      title: `Laporan Kas ${range.label} - ${settings.masjid_nama}`,
      user: req.session.user,
      currentPage: "kas",
      settings,
      bulan,
      periode: range.label,
      saldoAwal,
      masuk,
      keluar,
      saldoAkhir: saldoAwal + masuk - keluar,
      grupMasuk: group("masuk"),
      grupKeluar: group("keluar"),
      transaksi,
      labelTanggal,
      cetakTanggal: labelTanggal(todayIso(lokasi(settings).tz)),
    });
  } catch (error) {
    console.error("Error laporan kas:", error);
    req.flash("error", "Gagal memuat laporan kas");
    res.redirect("/kas");
  }
});

// Ekspor Excel (nilai teks diamankan dari rumus berbahaya)
router.get("/export", async (req, res) => {
  try {
    const ExcelJS = require("exceljs");
    const db = req.app.locals.db;
    const { settings, ym } = await currentMonth(req);
    const f = parseQuery(req, ym);
    const filter = buildFilter({ range: f.range, tipe: f.tipe, kategori: f.kategori, q: f.q });
    const [rows] = await db.execute(
      `SELECT k.tanggal, k.tipe, k.kategori, k.jumlah, k.pihak, k.keterangan, u.name AS oleh
         FROM kas_transaksi k LEFT JOIN users u ON u.id = k.user_id ${filter.sql} ORDER BY k.tanggal ASC, k.id ASC`,
      filter.params
    );
    const safe = (v) => (typeof v === "string" && /^[=+\-@\t\r]/.test(v) ? "'" + v : v || "");

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Kas");
    ws.addRow([`Kas ${settings.masjid_nama} — ${f.range ? f.range.label : "Semua periode"}`]).font = { bold: true, size: 14 };
    ws.addRow([]);
    const head = ws.addRow(["Tanggal", "Jenis", "Kategori", "Pihak", "Keterangan", "Pemasukan", "Pengeluaran", "Dicatat oleh"]);
    head.font = { bold: true, color: { argb: "FFFFFFFF" } };
    head.eachCell((c) => (c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF14912A" } }));
    let masuk = 0, keluar = 0;
    rows.forEach((r) => {
      const n = Number(r.jumlah);
      if (r.tipe === "masuk") masuk += n; else keluar += n;
      ws.addRow([dateToIso(r.tanggal), r.tipe === "masuk" ? "Pemasukan" : "Pengeluaran", kategoriLabel(r.tipe, r.kategori), safe(r.pihak), safe(r.keterangan), r.tipe === "masuk" ? n : null, r.tipe === "keluar" ? n : null, safe(r.oleh)]);
    });
    const total = ws.addRow(["", "", "", "", "TOTAL", masuk, keluar, ""]);
    total.font = { bold: true };
    const saldo = ws.addRow(["", "", "", "", "SELISIH", masuk - keluar, "", ""]);
    saldo.font = { bold: true };
    [6, 7].forEach((c) => (ws.getColumn(c).numFmt = '#,##0'));
    [14, 12, 24, 26, 36, 16, 16, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="kas-${f.range ? f.bulan : "semua"}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Error ekspor kas:", error);
    if (!res.headersSent) {
      req.flash("error", "Gagal mengekspor kas");
      res.redirect("/kas");
    }
  }
});

module.exports = { router, KATEGORI, rupiah };
