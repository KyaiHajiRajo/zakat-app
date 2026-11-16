const express = require("express");
const { prayerTimes } = require("../utils/prayer-times");
const { getSettings, saveSettings, lokasi } = require("../utils/settings");
const { KOTA, byKey } = require("../utils/kota-indonesia");
const { reply } = require("../utils/respond");
const { logAktivitas } = require("../utils/audit");
const { normalizeWa, displayWa } = require("../utils/phone");

const JENIS = {
  imam_tarawih: { label: "Imam Tarawih", icon: "fa-mosque", kolom: "imam_tarawih" },
  bilal: { label: "Bilal Tarawih", icon: "fa-bullhorn", kolom: "bilal" },
  penceramah: { label: "Penceramah / Kultum", icon: "fa-microphone-lines", kolom: "penceramah" },
};
const KOLOM = [
  { key: "imam_tarawih", label: "Imam Tarawih", icon: "fa-mosque", placeholder: "Nama imam" },
  { key: "bilal", label: "Bilal", icon: "fa-bullhorn", placeholder: "Nama bilal" },
  { key: "penceramah", label: "Penceramah / Kultum", icon: "fa-microphone-lines", placeholder: "Nama penceramah" },
];

const { HARI, BULAN, toUtc, toIso, addDays, validIso, labelTanggal } = require("../utils/dates");

async function buildPage(req, { readonly }) {
  const db = req.app.locals.db;
  const settings = await getSettings(db);
  const lok = lokasi(settings);
  const ihtiyati = Number(settings.ihtiyati_menit) || 0;
  const jumlahHari = Number(settings.ramadan_hari) === 29 ? 29 : 30;
  const mulai = validIso(settings.ramadan_mulai) ? settings.ramadan_mulai : "2027-02-08";

  const local = new Date(Date.now() + lok.tz * 3600e3);
  const today = toIso(local.getTime());
  const nowMinutes = local.getUTCHours() * 60 + local.getUTCMinutes();

  const [kegiatan] = await db.execute(
    "SELECT id, tanggal, jenis, petugas, keterangan, jumlah FROM ramadan_kegiatan WHERE jenis IN ('imam_tarawih','bilal','penceramah') ORDER BY tanggal ASC, id ASC"
  );
  const byDay = new Map();
  const outside = [];
  const akhir = addDays(mulai, jumlahHari - 1);
  for (const item of kegiatan) {
    const tanggal = typeof item.tanggal === "string" ? item.tanggal : toIso(Date.UTC(item.tanggal.getFullYear(), item.tanggal.getMonth(), item.tanggal.getDate()));
    const row = { ...item, tanggal, kolom: JENIS[item.jenis].kolom, jenisLabel: JENIS[item.jenis].label };
    if (tanggal < mulai || tanggal > akhir) {
      outside.push(row);
      continue;
    }
    if (!byDay.has(tanggal)) byDay.set(tanggal, []);
    byDay.get(tanggal).push(row);
  }

  const days = [];
  for (let i = 0; i < jumlahHari; i++) {
    const tanggal = addDays(mulai, i);
    const d = new Date(toUtc(tanggal));
    const times = prayerTimes(tanggal, { lat: lok.lat, lng: lok.lng, tz: lok.tz, ihtiyati });
    const no = i + 1;
    days.push({
      no,
      tanggal,
      hari: HARI[d.getUTCDay()],
      label: `${d.getUTCDate()} ${BULAN[d.getUTCMonth()].slice(0, 3)}`,
      times: times.text,
      minutes: times.minutes,
      isToday: tanggal === today,
      isPast: tanggal < today,
      isJumat: d.getUTCDay() === 5,
      badge: no === 17 ? "Nuzulul Qur'an" : no >= 21 && no % 2 === 1 ? "Malam ganjil" : null,
      kegiatan: byDay.get(tanggal) || [],
    });
  }

  const hariIni = days.find((d) => d.isToday) || null;
  const selisihMulai = Math.round((toUtc(mulai) - toUtc(today)) / 864e5);
  const todayTimes = hariIni
    ? hariIni.minutes
    : prayerTimes(today, { lat: lok.lat, lng: lok.lng, tz: lok.tz, ihtiyati }).minutes;

  return {
    title: "Jadwal Ramadan - " + settings.masjid_nama,
    user: req.session.user || null,
    currentPage: "ramadan",
    readonly,
    isAdmin: !!(req.session.user && req.session.user.role === "admin"),
    settings,
    lok,
    kota: KOTA,
    jenis: JENIS,
    kolom: KOLOM,
    days,
    outside,
    today,
    nowMinutes,
    todayTimes,
    todayLabel: labelTanggal(today),
    hariIni,
    selisihMulai,
    mulaiLabel: labelTanggal(mulai),
    akhirLabel: labelTanggal(akhir),
    idulFitriLabel: labelTanggal(addDays(mulai, jumlahHari)),
    totalKegiatan: kegiatan.length,
    success: req.flash("success"),
    error: req.flash("error"),
  };
}

// ---------- Halaman publik (tanpa login, hanya baca) ----------
const publicRouter = express.Router();
publicRouter.get("/", async (req, res) => {
  try {
    res.render("ramadan/index", await buildPage(req, { readonly: true }));
  } catch (error) {
    console.error("Error ramadan publik:", error);
    res.status(500).render("error", { title: "Server Error", message: "Gagal memuat jadwal Ramadan", error: {} });
  }
});

// ---------- Halaman pengelola ----------
const router = express.Router();

router.get("/", async (req, res) => {
  try {
    res.render("ramadan/index", await buildPage(req, { readonly: false }));
  } catch (error) {
    console.error("Error ramadan:", error);
    req.flash("error", "Gagal memuat jadwal Ramadan");
    res.redirect("/dashboard");
  }
});

// Simpan nama petugas langsung dari sel tabel (tanpa popup).
//  - ada id + nama kosong  -> hapus
//  - ada id + nama baru    -> ubah nama
//  - tanpa id + nama       -> tambah
router.post("/kegiatan/inline", async (req, res) => {
  const db = req.app.locals.db;
  const user = req.session.user;
  const tanggal = String(req.body.tanggal || "").trim();
  const jenis = String(req.body.jenis || "").trim();
  const petugas = String(req.body.petugas || "").trim().replace(/\s+/g, " ");
  const id = parseInt(req.body.id, 10) || 0;

  const fail = (status, message) => res.status(status).json({ ok: false, message });
  if (!validIso(tanggal)) return fail(400, "Tanggal tidak valid");
  if (!JENIS[jenis]) return fail(400, "Jenis petugas tidak valid");
  if (petugas.length > 120) return fail(400, "Nama maksimal 120 karakter");

  const label = JENIS[jenis].label;
  try {
    if (id) {
      const [rows] = await db.execute("SELECT petugas FROM ramadan_kegiatan WHERE id = ? AND jenis = ?", [id, jenis]);
      if (!rows.length) return fail(404, "Data tidak ditemukan, muat ulang halaman");
      const lama = rows[0].petugas;
      if (!petugas) {
        await db.execute("DELETE FROM ramadan_kegiatan WHERE id = ?", [id]);
        logAktivitas(db, { user, aksi: "Menghapus", modul: "Ramadan", objekId: id, ringkasan: `Menghapus petugas ${label.toLowerCase()}: ${lama} · ${labelTanggal(tanggal)}`, ip: req.ip });
        return res.json({ ok: true, deleted: true });
      }
      if (petugas !== lama) {
        await db.execute("UPDATE ramadan_kegiatan SET petugas = ? WHERE id = ?", [petugas, id]);
        logAktivitas(db, { user, aksi: "Mengubah", modul: "Ramadan", objekId: id, ringkasan: `Mengubah petugas ${label.toLowerCase()} ${labelTanggal(tanggal)}: ${lama} → ${petugas}`, ip: req.ip });
      }
      return res.json({ ok: true, id, petugas });
    }

    if (!petugas) return res.json({ ok: true, noop: true });
    const [result] = await db.execute(
      "INSERT INTO ramadan_kegiatan (tanggal, jenis, petugas, user_id) VALUES (?, ?, ?, ?)",
      [tanggal, jenis, petugas, user.id]
    );
    logAktivitas(db, { user, aksi: "Menambah", modul: "Ramadan", objekId: result.insertId, ringkasan: `Menambah petugas ${label.toLowerCase()} ${labelTanggal(tanggal)}: ${petugas}`, ip: req.ip });
    res.json({ ok: true, id: result.insertId, petugas });
  } catch (error) {
    console.error("Error simpan petugas inline:", error);
    fail(500, "Gagal menyimpan, coba lagi");
  }
});

// ---------- Pengingat petugas lewat WhatsApp ----------
const namaKey = (nama) => String(nama || "").trim().toLowerCase().replace(/\s+/g, " ").slice(0, 120);

// Daftar petugas pada satu tanggal + nomor WA tersimpan + pesan siap kirim
router.get("/pengingat", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const tanggal = String(req.query.tanggal || "").trim();
    if (!validIso(tanggal)) return res.status(400).json({ ok: false, message: "Tanggal tidak valid" });
    const settings = await getSettings(db);
    const lok = lokasi(settings);
    const tipe = settings.masjid_tipe === "masjid" ? "Masjid" : "Musholla";
    const today = toIso(Date.now() + lok.tz * 3600e3);
    const selisih = Math.round((toUtc(tanggal) - toUtc(today)) / 864e5);
    const kapan = selisih === 0 ? "hari ini" : selisih === 1 ? "besok" : "pada";
    const d = new Date(toUtc(tanggal));
    const hariLabel = `${HARI[d.getUTCDay()]}, ${labelTanggal(tanggal)}`;
    const mulai = validIso(settings.ramadan_mulai) ? settings.ramadan_mulai : "2027-02-08";
    const hariKe = Math.round((toUtc(tanggal) - toUtc(mulai)) / 864e5) + 1;
    const ihtiyati = Number(settings.ihtiyati_menit) || 0;
    const t = prayerTimes(tanggal, { lat: lok.lat, lng: lok.lng, tz: lok.tz, ihtiyati }).text;
    const malam = hariKe >= 1 && hariKe <= 30 ? ` (malam ke-${hariKe} Ramadan)` : "";

    const [rows] = await db.execute(
      `SELECT k.id, k.jenis, k.petugas, p.nomor_wa
         FROM ramadan_kegiatan k LEFT JOIN petugas_kontak p ON p.nama_key = LOWER(TRIM(k.petugas))
        WHERE k.tanggal = ? AND k.jenis IN ('imam_tarawih','bilal','penceramah')
        ORDER BY FIELD(k.jenis,'imam_tarawih','bilal','penceramah'), k.id`,
      [tanggal]
    );
    const items = rows.map((r) => {
      const peran = JENIS[r.jenis].label;
      const waktu = r.jenis === "penceramah" ? `kultum setelah tarawih/subuh sesuai jadwal` : `sholat Tarawih setelah Isya (${t.isya})`;
      const pesan = `Assalamu'alaikum warahmatullahi wabarakatuh, ${r.petugas}.\n\nMengingatkan, ${kapan === "pada" ? "pada " : kapan + " "}*${hariLabel}*${malam} Bapak/Ibu dijadwalkan bertugas sebagai *${peran}* di *${settings.masjid_nama}*.\n\nWaktu: ${waktu}.\n\nJika berhalangan, mohon kabari pengurus agar bisa dicarikan pengganti. Jazakumullahu khairan 🙏`;
      return { id: r.id, jenis: r.jenis, peran, icon: JENIS[r.jenis].icon, petugas: r.petugas, nomor: r.nomor_wa || "", tampil: displayWa(r.nomor_wa), pesan };
    });
    const baris = items.map((i) => `• ${i.peran}: ${i.petugas}`).join("\n");
    const grup = `🌙 *Jadwal Petugas ${tipe} ${settings.masjid_nama.replace(/^(musholla|masjid)\s+/i, "")}*\n📅 ${hariLabel}${malam}\n\n${baris || "Belum ada petugas yang dijadwalkan."}\n\nSholat Tarawih setelah Isya (${t.isya}). Mohon hadir tepat waktu. Jazakumullahu khairan 🙏`;
    res.json({ ok: true, tanggal, hariLabel, kapan: kapan === "pada" ? hariLabel : kapan, items, grup });
  } catch (error) {
    console.error("Error pengingat petugas:", error);
    res.status(500).json({ ok: false, message: "Gagal memuat daftar petugas" });
  }
});

// Simpan nomor WA petugas (sekali untuk semua jadwal dengan nama yang sama). Nomor kosong = hapus.
router.post("/kontak", async (req, res) => {
  const db = req.app.locals.db;
  const nama = String(req.body.nama || "").trim().replace(/\s+/g, " ");
  const mentah = String(req.body.nomor || "").trim();
  if (!nama || nama.length > 120) return res.status(400).json({ ok: false, message: "Nama petugas tidak valid" });
  try {
    if (!mentah) {
      await db.execute("DELETE FROM petugas_kontak WHERE nama_key = ?", [namaKey(nama)]);
      logAktivitas(db, { user: req.session.user, aksi: "Menghapus", modul: "Ramadan", ringkasan: `Menghapus nomor WA petugas ${nama}`, ip: req.ip });
      return res.json({ ok: true, nomor: "", tampil: "" });
    }
    const nomor = normalizeWa(mentah);
    if (!nomor) return res.status(400).json({ ok: false, message: "Nomor WhatsApp tidak valid (contoh: 0812 3456 7890)" });
    await db.execute(
      "INSERT INTO petugas_kontak (nama_key, nama, nomor_wa) VALUES (?, ?, ?) ON DUPLICATE KEY UPDATE nama = VALUES(nama), nomor_wa = VALUES(nomor_wa)",
      [namaKey(nama), nama, nomor]
    );
    logAktivitas(db, { user: req.session.user, aksi: "Mengubah", modul: "Ramadan", ringkasan: `Menyimpan nomor WA petugas ${nama}`, ip: req.ip });
    res.json({ ok: true, nomor, tampil: displayWa(nomor) });
  } catch (error) {
    console.error("Error simpan kontak petugas:", error);
    res.status(500).json({ ok: false, message: "Gagal menyimpan nomor" });
  }
});

router.post("/pengaturan", async (req, res) => {
  const back = "/ramadan";
  if (!req.session.user || req.session.user.role !== "admin") {
    return reply(req, res, { ok: false, message: "Hanya admin yang dapat mengubah pengaturan", redirect: back, status: 403 });
  }
  const b = req.body;
  const nama = String(b.masjid_nama || "").trim();
  const tipe = b.masjid_tipe === "masjid" ? "masjid" : "musholla";
  const alamat = String(b.masjid_alamat || "").trim();
  const kota = String(b.lokasi_kota || "");
  const mulai = String(b.ramadan_mulai || "").trim();
  const hari = Number(b.ramadan_hari) === 29 ? "29" : "30";
  const ihtiyati = Number(b.ihtiyati_menit);
  const tahun = Number(b.ramadan_tahun_h);

  if (!nama || nama.length > 100) return reply(req, res, { ok: false, message: "Nama masjid/musholla wajib diisi (maks. 100 karakter)", redirect: back });
  if (alamat.length > 255) return reply(req, res, { ok: false, message: "Alamat maksimal 255 karakter", redirect: back });
  if (!validIso(mulai)) return reply(req, res, { ok: false, message: "Tanggal 1 Ramadan tidak valid", redirect: back });
  if (!Number.isInteger(ihtiyati) || ihtiyati < 0 || ihtiyati > 5) return reply(req, res, { ok: false, message: "Ihtiyati harus 0–5 menit", redirect: back });
  if (!Number.isInteger(tahun) || tahun < 1400 || tahun > 1600) return reply(req, res, { ok: false, message: "Tahun Hijriah tidak valid", redirect: back });

  const values = { masjid_nama: nama, masjid_tipe: tipe, masjid_alamat: alamat, ramadan_mulai: mulai, ramadan_hari: hari, ihtiyati_menit: ihtiyati, ramadan_tahun_h: tahun };
  if (byKey(kota)) {
    values.lokasi_kota = kota;
  } else if (kota === "kustom") {
    const lat = Number(b.lokasi_lat), lng = Number(b.lokasi_lng), zona = Number(b.lokasi_zona);
    if (!(lat >= -12 && lat <= 7) || !(lng >= 94 && lng <= 142) || ![7, 8, 9].includes(zona)) {
      return reply(req, res, { ok: false, message: "Koordinat kustom di luar wilayah Indonesia atau zona waktu tidak valid", redirect: back });
    }
    Object.assign(values, { lokasi_kota: "kustom", lokasi_lat: lat, lokasi_lng: lng, lokasi_zona: zona });
  } else {
    return reply(req, res, { ok: false, message: "Kota tidak dikenal", redirect: back });
  }

  try {
    await saveSettings(req.app.locals.db, values);
    reply(req, res, { ok: true, message: "Pengaturan disimpan, jadwal dihitung ulang", redirect: back });
  } catch (err) {
    console.error("Error simpan pengaturan:", err);
    reply(req, res, { ok: false, message: "Gagal menyimpan pengaturan", redirect: back, status: 500 });
  }
});

module.exports = { router, publicRouter };
