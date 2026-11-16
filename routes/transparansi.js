const express = require("express");
const { getSettings, lokasi } = require("../utils/settings");
const { dateToIso, bulanRange, todayIso, labelTanggal, BULAN } = require("../utils/dates");
const { KATEGORI } = require("./kas");

const rupiah = (n) => "Rp " + new Intl.NumberFormat("id-ID").format(Math.round(Number(n) || 0));
const MAX_ROWS = 120;

const router = express.Router();

router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const settings = await getSettings(db);
    const base = {
      title: "Laporan Kas Terbuka - " + settings.masjid_nama,
      user: req.session.user || null,
      settings,
      rupiah,
      success: [],
      error: [],
    };
    if (settings.transparansi_aktif !== "1") {
      return res.render("transparansi/index", { ...base, tutup: true });
    }

    const ym = todayIso(lokasi(settings).tz).slice(0, 7);
    const param = String(req.query.bulan || "");
    const semua = param === "semua";
    const bulan = semua ? "semua" : bulanRange(param) ? param : ym;
    const range = semua ? null : bulanRange(bulan);
    const showNames = settings.transparansi_nama_donatur === "1";

    const [[saldoRow]] = await db.execute("SELECT COALESCE(SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE -jumlah END),0) AS saldo FROM kas_transaksi");
    const where = range ? "WHERE tanggal BETWEEN ? AND ?" : "";
    const params = range ? [range.awal, range.akhir] : [];
    const [[ring]] = await db.execute(
      `SELECT COALESCE(SUM(CASE WHEN tipe='masuk' THEN jumlah END),0) AS masuk, COALESCE(SUM(CASE WHEN tipe='keluar' THEN jumlah END),0) AS keluar, COUNT(*) AS n FROM kas_transaksi ${where}`,
      params
    );
    const [grup] = await db.execute(`SELECT tipe, kategori, SUM(jumlah) AS total FROM kas_transaksi ${where} GROUP BY tipe, kategori ORDER BY total DESC`, params);
    const [rows] = await db.query(
      `SELECT tanggal, tipe, kategori, jumlah, pihak, keterangan FROM kas_transaksi ${where} ORDER BY tanggal DESC, id DESC LIMIT ${MAX_ROWS}`,
      params
    );

    // grafik 6 bulan berakhir di bulan terpilih
    const endYm = range ? bulan : ym;
    const [ey, em] = endYm.split("-").map(Number);
    const months = [];
    for (let i = 5; i >= 0; i--) months.push(new Date(Date.UTC(ey, em - 1 - i, 1)).toISOString().slice(0, 7));
    const [series] = await db.execute(
      `SELECT DATE_FORMAT(tanggal,'%Y-%m') AS ym, SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE 0 END) AS masuk, SUM(CASE WHEN tipe='keluar' THEN jumlah ELSE 0 END) AS keluar
         FROM kas_transaksi WHERE tanggal BETWEEN ? AND ? GROUP BY ym`,
      [`${months[0]}-01`, bulanRange(months[5]).akhir]
    );
    const chart = months.map((m) => {
      const r = series.find((s) => s.ym === m);
      const [y, mo] = m.split("-").map(Number);
      return { ym: m, label: BULAN[mo - 1].slice(0, 3) + " " + String(y).slice(2), masuk: Number(r ? r.masuk : 0), keluar: Number(r ? r.keluar : 0), aktif: m === bulan };
    });

    const [monthRows] = await db.execute("SELECT DISTINCT DATE_FORMAT(tanggal,'%Y-%m') AS ym FROM kas_transaksi ORDER BY ym DESC LIMIT 12");
    const monthSet = new Set(monthRows.map((r) => r.ym));
    monthSet.add(ym);
    if (range) monthSet.add(bulan);
    const bulanList = [...monthSet].sort().reverse().slice(0, 12).map((m) => ({ ym: m, label: bulanRange(m).label }));

    const label = (tipe, key) => (KATEGORI[tipe] && KATEGORI[tipe][key] ? KATEGORI[tipe][key] : { label: key, icon: "fa-circle" });
    const bagi = (tipe) => {
      const items = grup.filter((g) => g.tipe === tipe).map((g) => ({ label: label(tipe, g.kategori).label, icon: label(tipe, g.kategori).icon, total: Number(g.total) }));
      const sum = items.reduce((a, b) => a + b.total, 0);
      return { sum, items: items.map((i) => ({ ...i, persen: sum ? Math.round((i.total / sum) * 100) : 0 })) };
    };

    // Zakat fitrah: tahun terbaru yang ada datanya
    const [[tahunRow]] = await db.execute("SELECT YEAR(MAX(COALESCE(tanggal, DATE(created_at)))) AS tahun FROM muzakki");
    const tahun = tahunRow.tahun || null;
    let zakat = null;
    if (tahun) {
      const [[z]] = await db.execute(
        `SELECT COUNT(*) AS muzakki, COALESCE(SUM(jumlah_jiwa),0) AS jiwa,
                COALESCE(SUM(CASE WHEN jenis_zakat='uang' THEN jumlah_uang END),0) AS uang,
                COALESCE(SUM(CASE WHEN jenis_zakat='beras' THEN jumlah_beras_kg END),0) AS beras
           FROM muzakki WHERE YEAR(COALESCE(tanggal, DATE(created_at))) = ?`,
        [tahun]
      );
      const [[d]] = await db.execute(
        `SELECT COUNT(*) AS penerima, COALESCE(SUM(CASE WHEN jenis_zakat='uang' THEN jumlah END),0) AS uang, COALESCE(SUM(CASE WHEN jenis_zakat='beras' THEN jumlah END),0) AS beras
           FROM distribusi_zakat WHERE status IN ('disalurkan','diterima') AND YEAR(created_at) = ?`,
        [tahun]
      );
      if (Number(z.muzakki) > 0) {
        zakat = { tahun, muzakki: Number(z.muzakki), jiwa: Number(z.jiwa), uang: Number(z.uang), beras: Number(z.beras), penerima: Number(d.penerima), uangSalur: Number(d.uang), berasSalur: Number(d.beras) };
      }
    }

    const masuk = Number(ring.masuk), keluar = Number(ring.keluar);
    res.render("transparansi/index", {
      ...base,
      tutup: false,
      bulan,
      periodeLabel: range ? range.label : "Semua periode",
      saldo: Number(saldoRow.saldo),
      masuk,
      keluar,
      selisih: masuk - keluar,
      jumlahTransaksi: Number(ring.n),
      chart,
      chartMax: Math.max(1, ...chart.map((c) => Math.max(c.masuk, c.keluar))),
      bulanList,
      grupMasuk: bagi("masuk"),
      grupKeluar: bagi("keluar"),
      rows: rows.map((r) => {
        const k = label(r.tipe, r.kategori);
        const pihak = showNames && r.pihak ? r.pihak : "";
        return { tanggal: dateToIso(r.tanggal), tipe: r.tipe, kategoriLabel: k.label, icon: k.icon, jumlah: Number(r.jumlah), pihak, keterangan: r.keterangan || "" };
      }),
      terpotong: Number(ring.n) > MAX_ROWS,
      maxRows: MAX_ROWS,
      zakat,
      labelTanggal,
      hariIni: labelTanggal(todayIso(lokasi(settings).tz)),
    });
  } catch (error) {
    console.error("Error transparansi:", error);
    res.status(500).render("error", { title: "Server Error", message: "Gagal memuat laporan kas", error: {} });
  }
});

module.exports = { router };
