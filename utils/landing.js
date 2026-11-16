// Data untuk landing page publik. Setiap bagian gagal-senyap agar halaman tetap tampil walau satu tabel bermasalah.
const { prayerTimes } = require("./prayer-times");
const { getSettings, lokasi } = require("./settings");
const { addDays, validIso, todayIso, labelTanggal, hariTanggal, dateToIso, bulanRange, BULAN } = require("./dates");
const { normalizeWa } = require("./phone");

const AGENDA_ICON = {
  pengajian: "fa-book-open", kerja_bakti: "fa-broom", peringatan: "fa-star-and-crescent", rapat: "fa-people-group",
  sosial: "fa-hand-holding-heart", ramadan: "fa-moon", lainnya: "fa-calendar-day",
};
const PERAN = { imam_tarawih: ["Imam Tarawih", "fa-mosque"], bilal: ["Bilal", "fa-bullhorn"], penceramah: ["Penceramah / Kultum", "fa-microphone-lines"] };
const TUJUAN = { donasi: "Donasi Umum", infaq_ramadan: "Infaq Ramadan", wakaf: "Wakaf" };

async function safe(fn, fallback) {
  try {
    return await fn();
  } catch (error) {
    console.error("Landing: bagian gagal dimuat:", error.message);
    return fallback;
  }
}

async function landingData(db) {
  const settings = await getSettings(db);
  const lok = lokasi(settings);
  const today = todayIso(lok.tz);
  const ihtiyati = Number(settings.ihtiyati_menit) || 0;
  const jumlahHari = Number(settings.ramadan_hari) === 29 ? 29 : 30;
  const mulai = validIso(settings.ramadan_mulai) ? settings.ramadan_mulai : "2027-02-08";
  const akhir = addDays(mulai, jumlahHari - 1);
  const dayNo = today >= mulai && today <= akhir ? Math.round((Date.parse(today) - Date.parse(mulai)) / 864e5) + 1 : 0;
  const selisihMulai = Math.round((Date.parse(mulai) - Date.parse(today)) / 864e5);
  const times = prayerTimes(today, { lat: lok.lat, lng: lok.lng, tz: lok.tz, ihtiyati });

  const ramadan = { dayNo, total: jumlahHari, selisihMulai, tahunH: settings.ramadan_tahun_h, mulaiLabel: labelTanggal(mulai), akhirLabel: labelTanggal(akhir), selesai: today > akhir };

  // petugas malam ini (hanya saat Ramadan)
  const petugas = dayNo
    ? await safe(async () => {
        const [rows] = await db.execute(
          "SELECT jenis, petugas FROM ramadan_kegiatan WHERE tanggal = ? AND jenis IN ('imam_tarawih','bilal','penceramah') ORDER BY FIELD(jenis,'imam_tarawih','bilal','penceramah'), id",
          [today]
        );
        const map = new Map();
        rows.forEach((r) => {
          if (!map.has(r.jenis)) map.set(r.jenis, []);
          map.get(r.jenis).push(r.petugas);
        });
        return [...map.entries()].map(([jenis, names]) => ({ label: PERAN[jenis][0], icon: PERAN[jenis][1], names }));
      }, [])
    : [];

  // donasi
  const donasiAktif = settings.donasi_aktif === "1" && !!(settings.donasi_qris || (settings.donasi_bank_nama && settings.donasi_bank_norek));
  const donasi = donasiAktif
    ? await safe(async () => {
        const [[tot]] = await db.execute("SELECT COUNT(*) AS n, COALESCE(SUM(jumlah),0) AS total FROM donasi WHERE status = 'diterima'");
        const [terbaru] = await db.execute("SELECT nama, anonim, jumlah, tujuan FROM donasi WHERE status = 'diterima' ORDER BY diproses_at DESC, id DESC LIMIT 4");
        return {
          jumlah: Number(tot.n),
          total: Number(tot.total),
          qris: settings.donasi_qris || "",
          bank: settings.donasi_bank_nama || "",
          norek: settings.donasi_bank_norek || "",
          atasNama: settings.donasi_bank_atas_nama || "",
          catatan: settings.donasi_catatan || "",
          wa: normalizeWa(settings.donasi_wa) || "",
          terbaru: terbaru.map((d) => ({ nama: d.anonim ? "Hamba Allah" : d.nama, jumlah: Number(d.jumlah), tujuan: TUJUAN[d.tujuan] || "Donasi" })),
        };
      }, null)
    : null;

  // laporan kas terbuka
  const kas =
    settings.transparansi_aktif === "1"
      ? await safe(async () => {
          const month = bulanRange(today.slice(0, 7));
          const [[row]] = await db.execute(
            `SELECT COALESCE(SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE -jumlah END),0) AS saldo,
                    COALESCE(SUM(CASE WHEN tipe='masuk' AND tanggal BETWEEN ? AND ? THEN jumlah END),0) AS masuk,
                    COALESCE(SUM(CASE WHEN tipe='keluar' AND tanggal BETWEEN ? AND ? THEN jumlah END),0) AS keluar
               FROM kas_transaksi`,
            [month.awal, month.akhir, month.awal, month.akhir]
          );
          const [ey, em] = today.slice(0, 7).split("-").map(Number);
          const months = [];
          for (let i = 5; i >= 0; i--) months.push(new Date(Date.UTC(ey, em - 1 - i, 1)).toISOString().slice(0, 7));
          const [series] = await db.execute(
            `SELECT DATE_FORMAT(tanggal,'%Y-%m') AS ym, SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE 0 END) AS masuk, SUM(CASE WHEN tipe='keluar' THEN jumlah ELSE 0 END) AS keluar
               FROM kas_transaksi WHERE tanggal BETWEEN ? AND ? GROUP BY ym`,
            [`${months[0]}-01`, bulanRange(months[5]).akhir]
          );
          const chart = months.map((m) => {
            const r = series.find((s) => s.ym === m);
            return { label: BULAN[Number(m.slice(5)) - 1].slice(0, 3), masuk: Number(r ? r.masuk : 0), keluar: Number(r ? r.keluar : 0) };
          });
          return {
            saldo: Number(row.saldo), masuk: Number(row.masuk), keluar: Number(row.keluar), bulanLabel: month.label,
            chart, chartMax: Math.max(1, ...chart.map((c) => Math.max(c.masuk, c.keluar))),
          };
        }, null)
      : null;

  // agenda & pengumuman publik
  const agenda = await safe(async () => {
    const [rows] = await db.execute(
      "SELECT id, judul, jenis, tanggal, TIME_FORMAT(jam_mulai,'%H:%i') AS jam_mulai, lokasi, pemateri FROM agenda WHERE tanggal >= ? AND tampil_publik = 1 ORDER BY tanggal ASC, jam_mulai ASC LIMIT 4",
      [today]
    );
    return rows.map((a) => {
      const iso = dateToIso(a.tanggal);
      return { ...a, tanggal: iso, label: labelTanggal(iso), hari: hariTanggal(iso), tgl: Number(iso.slice(8)), bln: BULAN[Number(iso.slice(5, 7)) - 1].slice(0, 3).toUpperCase(), icon: AGENDA_ICON[a.jenis] || "fa-calendar-day", hariIni: iso === today };
    });
  }, []);
  const pengumuman = await safe(async () => {
    const [rows] = await db.execute(
      "SELECT id, judul, isi, penting FROM pengumuman WHERE (berlaku_sampai IS NULL OR berlaku_sampai >= ?) AND tampil_publik = 1 ORDER BY penting DESC, created_at DESC LIMIT 3",
      [today]
    );
    return rows.map((p) => ({ ...p, penting: !!p.penting, ringkas: String(p.isi || "").replace(/\s+/g, " ").slice(0, 140) }));
  }, []);

  // zakat fitrah (tahun terbaru yang ada datanya)
  const zakat = await safe(async () => {
    const [[t]] = await db.execute("SELECT YEAR(MAX(COALESCE(tanggal, DATE(created_at)))) AS tahun FROM muzakki");
    if (!t.tahun) return null;
    const [[z]] = await db.execute(
      `SELECT COUNT(*) AS muzakki, COALESCE(SUM(jumlah_jiwa),0) AS jiwa,
              COALESCE(SUM(CASE WHEN jenis_zakat='uang' THEN jumlah_uang END),0) AS uang,
              COALESCE(SUM(CASE WHEN jenis_zakat='beras' THEN jumlah_beras_kg END),0) AS beras
         FROM muzakki WHERE YEAR(COALESCE(tanggal, DATE(created_at))) = ?`,
      [t.tahun]
    );
    const [[d]] = await db.execute("SELECT COUNT(*) AS penerima FROM distribusi_zakat WHERE status IN ('disalurkan','diterima') AND YEAR(created_at) = ?", [t.tahun]);
    const [[rt]] = await db.execute("SELECT COUNT(*) AS n FROM rt");
    return Number(z.muzakki) ? { tahun: t.tahun, muzakki: Number(z.muzakki), jiwa: Number(z.jiwa), uang: Number(z.uang), beras: Number(z.beras), penerima: Number(d.penerima), rt: Number(rt.n) } : null;
  }, null);

  return {
    settings, lok, today,
    todayLabel: `${hariTanggal(today)}, ${labelTanggal(today)}`,
    tipe: settings.masjid_tipe === "masjid" ? "Masjid" : "Musholla",
    times: times.minutes, timesText: times.text,
    ramadan, petugas, donasi, kas, agenda, pengumuman, zakat,
  };
}

module.exports = { landingData };
