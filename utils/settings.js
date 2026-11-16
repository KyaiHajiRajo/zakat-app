// Pengaturan aplikasi (profil masjid/musholla, lokasi, Ramadan) disimpan sebagai pasangan kunci-nilai.
const { byKey } = require("./kota-indonesia");

const DEFAULTS = {
  masjid_nama: "Musholla Nurul Fitrah",
  masjid_tipe: "musholla",
  masjid_alamat: "",
  lokasi_kota: "jakarta",
  lokasi_lat: "-6.2088",
  lokasi_lng: "106.8456",
  lokasi_zona: "7",
  ihtiyati_menit: "2",
  ramadan_tahun_h: "1448",
  ramadan_mulai: "2027-02-08", // perkiraan 1 Ramadan 1448 H; ubah setelah sidang isbat
  ramadan_hari: "30",
  // donasi online
  donasi_aktif: "1",
  donasi_bank_nama: "",
  donasi_bank_norek: "",
  donasi_bank_atas_nama: "",
  donasi_qris: "", // nama berkas QRIS di public/qris
  donasi_wa: "", // nomor WA pengurus untuk konfirmasi
  donasi_catatan: "",
  // transparansi publik
  transparansi_aktif: "1",
  transparansi_nama_donatur: "0",
};

let cache = null;

async function getSettings(db) {
  if (cache && Date.now() - cache.at < 30000) {
    return cache.data;
  }
  const [rows] = await db.execute("SELECT kunci, nilai FROM pengaturan");
  const data = { ...DEFAULTS };
  for (const row of rows) {
    if (row.kunci in DEFAULTS && row.nilai !== null) {
      data[row.kunci] = row.nilai;
    }
  }
  cache = { at: Date.now(), data };
  return data;
}

async function saveSettings(db, values) {
  for (const [kunci, nilai] of Object.entries(values)) {
    if (kunci in DEFAULTS) {
      await db.execute(
        "INSERT INTO pengaturan (kunci, nilai) VALUES (?, ?) ON DUPLICATE KEY UPDATE nilai = VALUES(nilai)",
        [kunci, String(nilai)]
      );
    }
  }
  cache = null;
}

// Lokasi efektif untuk perhitungan: kota preset, atau koordinat kustom bila kota = "kustom"
function lokasi(settings) {
  const preset = byKey(settings.lokasi_kota);
  if (preset) {
    return { nama: preset.nama, lat: preset.lat, lng: preset.lng, tz: preset.zona };
  }
  return {
    nama: "Lokasi kustom",
    lat: Number(settings.lokasi_lat),
    lng: Number(settings.lokasi_lng),
    tz: Number(settings.lokasi_zona),
  };
}

module.exports = { DEFAULTS, getSettings, saveSettings, lokasi };
