// Ringkasan modul masjid untuk widget dashboard. Gagal senyap (dashboard zakat tetap jalan bila tabel belum ada).
const { prayerTimes } = require("./prayer-times");
const { getSettings, lokasi } = require("./settings");
const { addDays, validIso, todayIso, labelTanggal, hariTanggal, dateToIso, bulanRange } = require("./dates");

async function masjidSummary(db) {
  try {
    const settings = await getSettings(db);
    const lok = lokasi(settings);
    const today = todayIso(lok.tz);
    const ihtiyati = Number(settings.ihtiyati_menit) || 0;
    const jumlahHari = Number(settings.ramadan_hari) === 29 ? 29 : 30;
    const mulai = validIso(settings.ramadan_mulai) ? settings.ramadan_mulai : "2027-02-08";
    const akhir = addDays(mulai, jumlahHari - 1);
    const dayNo = today >= mulai && today <= akhir ? Math.round((Date.parse(today) - Date.parse(mulai)) / 864e5) + 1 : 0;
    const selisihMulai = Math.round((Date.parse(mulai) - Date.parse(today)) / 864e5);
    const times = prayerTimes(today, { lat: lok.lat, lng: lok.lng, tz: lok.tz, ihtiyati }).minutes;

    const month = bulanRange(today.slice(0, 7));
    const [[kas]] = await db.execute(
      `SELECT COALESCE(SUM(CASE WHEN tipe='masuk' THEN jumlah ELSE -jumlah END),0) AS saldo,
              COALESCE(SUM(CASE WHEN tipe='masuk' AND tanggal BETWEEN ? AND ? THEN jumlah END),0) AS masuk,
              COALESCE(SUM(CASE WHEN tipe='keluar' AND tanggal BETWEEN ? AND ? THEN jumlah END),0) AS keluar
         FROM kas_transaksi`,
      [month.awal, month.akhir, month.awal, month.akhir]
    );
    const [agenda] = await db.execute(
      "SELECT id, judul, tanggal, TIME_FORMAT(jam_mulai,'%H:%i') AS jam_mulai, lokasi FROM agenda WHERE tanggal >= ? ORDER BY tanggal ASC, jam_mulai ASC LIMIT 3",
      [today]
    );
    const [pengumuman] = await db.execute(
      "SELECT id, judul, penting FROM pengumuman WHERE berlaku_sampai IS NULL OR berlaku_sampai >= ? ORDER BY penting DESC, created_at DESC LIMIT 2",
      [today]
    );
    const [[inv]] = await db.execute("SELECT COALESCE(SUM(CASE WHEN kondisi <> 'baik' THEN 1 END),0) AS bermasalah FROM inventaris");
    let donasiMenunggu = 0;
    try {
      const [[dn]] = await db.execute("SELECT COUNT(*) AS n FROM donasi WHERE status = 'menunggu'");
      donasiMenunggu = Number(dn.n);
    } catch (error) {
      donasiMenunggu = 0;
    }
    let petugas = [];
    if (dayNo) {
      const [rows] = await db.execute("SELECT jenis, petugas FROM ramadan_kegiatan WHERE tanggal = ? AND jenis IN ('imam_tarawih','bilal','penceramah') ORDER BY id", [today]);
      petugas = rows;
    }

    return {
      settings,
      lok,
      today,
      todayLabel: `${hariTanggal(today)}, ${labelTanggal(today)}`,
      ramadan: { dayNo, total: jumlahHari, selisihMulai, mulaiLabel: labelTanggal(mulai), tahunH: settings.ramadan_tahun_h, petugas },
      times,
      kas: { saldo: Number(kas.saldo), masuk: Number(kas.masuk), keluar: Number(kas.keluar), bulanLabel: month.label },
      agenda: agenda.map((a) => ({ ...a, tanggal: dateToIso(a.tanggal), label: labelTanggal(dateToIso(a.tanggal)) })),
      pengumuman,
      invBermasalah: Number(inv.bermasalah),
      donasiMenunggu,
    };
  } catch (error) {
    console.error("Ringkasan masjid gagal:", error.message);
    return null;
  }
}

module.exports = { masjidSummary };
