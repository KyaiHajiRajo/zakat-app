// Helper tanggal berbasis string ISO "YYYY-MM-DD" (aman dari pergeseran zona waktu server).
const HARI = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
const BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
const ISO = /^\d{4}-\d{2}-\d{2}$/;

const toUtc = (iso) => {
  const [y, m, d] = iso.split("-").map(Number);
  return Date.UTC(y, m - 1, d);
};
const toIso = (ms) => new Date(ms).toISOString().slice(0, 10);
const addDays = (iso, n) => toIso(toUtc(iso) + n * 864e5);
const validIso = (value) => ISO.test(value || "") && toIso(toUtc(value)) === value;
const labelTanggal = (iso) => {
  const d = new Date(toUtc(iso));
  return `${d.getUTCDate()} ${BULAN[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
};
const hariTanggal = (iso) => HARI[new Date(toUtc(iso)).getUTCDay()];

// Nilai kolom DATE dari mysql2 (objek Date lokal) -> "YYYY-MM-DD" tanpa bergeser hari
const dateToIso = (value) => {
  if (!value) return null;
  if (typeof value === "string") return value.slice(0, 10);
  const pad = (n) => String(n).padStart(2, "0");
  return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
};

// "YYYY-MM" valid? -> {awal, akhir, label}
function bulanRange(ym) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(ym || "")) return null;
  const [y, m] = ym.split("-").map(Number);
  const awal = `${ym}-01`;
  const akhir = toIso(Date.UTC(y, m, 0));
  return { awal, akhir, label: `${BULAN[m - 1]} ${y}` };
}

// Tanggal "hari ini" menurut zona waktu tertentu (jam)
const todayIso = (tz = 7) => toIso(Date.now() + tz * 3600e3);

module.exports = { HARI, BULAN, toUtc, toIso, addDays, validIso, labelTanggal, hariTanggal, dateToIso, bulanRange, todayIso };
