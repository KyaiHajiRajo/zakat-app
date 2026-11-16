// Log aktivitas: mencatat siapa mengubah apa beserta rinciannya. Dipasang sekali sebagai middleware.
const MODUL = {
  kas: "Kas", distribusi: "Distribusi", muzakki: "Muzakki", mustahik: "Mustahik", infak: "Infak",
  rt: "RT", "rt-rw": "RT", rw: "RW", users: "Pengguna", ramadan: "Ramadan", agenda: "Agenda", inventaris: "Inventaris",
};
const JENIS_KEGIATAN = { imam_tarawih: "Imam tarawih", bilal: "Bilal tarawih", penceramah: "Penceramah", takjil: "Takjil", buka_bersama: "Buka bersama", tilawah: "Tilawah", itikaf: "I'tikaf", lainnya: "Lainnya" };
const KATEGORI_KAS = {
  kotak_amal: "Kotak Amal", donasi: "Donasi", infaq_ramadan: "Infaq Ramadan", wakaf: "Wakaf", listrik_air: "Listrik & Air", kebersihan: "Kebersihan",
  perawatan: "Perawatan", perlengkapan: "Perlengkapan", konsumsi: "Konsumsi", honor: "Honor", kegiatan: "Kegiatan", sosial: "Bantuan Sosial", lainnya: "Lainnya",
};
const rp = (n) => "Rp " + new Intl.NumberFormat("id-ID").format(Math.round(Number(n) || 0));
const cut = (v, n = 60) => String(v === undefined || v === null ? "" : v).trim().replace(/\s+/g, " ").slice(0, n);
const first = (v) => (Array.isArray(v) ? v[0] : v);

function describe(method, path) {
  const parts = path.split("?")[0].split("/").filter(Boolean);
  const modulKey = parts[0];
  if (!MODUL[modulKey]) return null;
  const rest = parts.slice(1);
  const sub = rest.find((p) => ["kegiatan", "pengumuman", "pengaturan", "profile", "password", "status", "reset-password"].includes(p)) || "";

  let aksi = "Menambah";
  if (method === "DELETE" || rest.includes("delete") || rest.includes("hapus")) aksi = "Menghapus";
  else if (sub === "status") aksi = "Mengubah status";
  else if (sub === "reset-password") aksi = "Mereset password";
  else if (sub === "password") aksi = "Mengganti password";
  else if (sub === "pengaturan") aksi = "Mengubah pengaturan";
  else if (method === "PUT" || rest.includes("edit") || rest.includes("update")) aksi = "Mengubah";

  let objek = MODUL[modulKey].toLowerCase();
  if (modulKey === "ramadan" && sub === "kegiatan") objek = "jadwal petugas Ramadan";
  else if (modulKey === "ramadan") objek = "pengaturan Ramadan & lokasi";
  else if (modulKey === "agenda" && sub === "pengumuman") objek = "pengumuman";
  else if (modulKey === "agenda") objek = "agenda";
  else if (modulKey === "users" && rest[0] === "profile") objek = "profil sendiri";
  else if (modulKey === "users") objek = "pengguna";
  else if (modulKey === "kas") objek = "transaksi kas";
  else if (modulKey === "inventaris") objek = "barang inventaris";
  else if (modulKey === "distribusi") objek = "distribusi zakat";
  else if (modulKey === "muzakki") objek = "data muzakki";
  else if (modulKey === "mustahik") objek = "data mustahik";

  const idPart = rest.find((p) => /^\d+$/.test(p));
  return { aksi, modul: MODUL[modulKey], modulKey, sub, objek, objekId: idPart ? Number(idPart) : null };
}

// Rincian dari data yang dikirim (tanpa kata sandi / data sensitif)
function detailFromBody(d, body) {
  const b = body || {};
  switch (d.modulKey) {
    case "kas": {
      if (!b.tipe) return "";
      return [b.tipe === "masuk" ? "Pemasukan" : "Pengeluaran", rp(String(b.jumlah || "").replace(",", ".")), KATEGORI_KAS[b.kategori] || b.kategori, cut(b.pihak, 40)].filter(Boolean).join(" · ");
    }
    case "agenda": return cut(b.judul);
    case "inventaris": return [cut(b.nama), b.jumlah ? `${cut(b.jumlah, 8)} ${cut(b.satuan, 12)}` : "", b.kondisi ? cut(b.kondisi).replace("_", " ") : ""].filter(Boolean).join(" · ");
    case "ramadan":
      if (d.sub === "kegiatan") return [JENIS_KEGIATAN[b.jenis] || cut(b.jenis), cut(b.petugas, 40), cut(b.tanggal, 10)].filter(Boolean).join(" · ");
      return [cut(b.masjid_nama), b.lokasi_kota ? "kota " + cut(b.lokasi_kota, 20) : "", b.ramadan_mulai ? "1 Ramadan " + cut(b.ramadan_mulai, 10) : ""].filter(Boolean).join(" · ");
    case "users": return d.sub === "password" || d.sub === "reset-password" ? "" : [cut(b.name), b.role ? `(${cut(b.role, 10)})` : ""].filter(Boolean).join(" ");
    case "distribusi": return [b.status ? "status → " + cut(b.status, 20) : "", b.jenis_zakat ? cut(b.jenis_zakat, 10) : "", b.jumlah ? cut(b.jumlah, 12) : ""].filter(Boolean).join(" · ");
    case "muzakki": {
      const nama = first(b.nama_muzakki) || first(b.nama);
      const n = Array.isArray(b.nama_muzakki) ? b.nama_muzakki.filter(Boolean).length : 0;
      return nama ? cut(nama, 40) + (n > 1 ? ` (+${n - 1} lainnya)` : "") : "";
    }
    case "mustahik": return [cut(b.nama), cut(b.kategori, 20)].filter(Boolean).join(" · ");
    case "rt": case "rt-rw": return [cut(b.nomor_rt, 12), b.ketua_rt ? "ketua " + cut(b.ketua_rt, 30) : ""].filter(Boolean).join(" · ");
    case "rw": return [cut(b.nomor_rw, 12), b.ketua_rw ? "ketua " + cut(b.ketua_rw, 30) : ""].filter(Boolean).join(" · ");
    case "infak": return b.jumlah ? cut(b.jumlah, 14) : "";
    default: return "";
  }
}

// Label data yang sudah ada (diambil SEBELUM diproses agar penghapusan tetap bisa dijelaskan)
async function existingLabel(db, d) {
  if (!d.objekId) return "";
  const one = async (sql) => { const [rows] = await db.execute(sql, [d.objekId]); return rows[0]; };
  try {
    switch (d.modulKey) {
      case "kas": { const r = await one("SELECT tipe, jumlah, kategori, pihak FROM kas_transaksi WHERE id = ?"); return r ? [r.tipe === "masuk" ? "Pemasukan" : "Pengeluaran", rp(r.jumlah), KATEGORI_KAS[r.kategori] || r.kategori, cut(r.pihak, 40)].filter(Boolean).join(" · ") : ""; }
      case "agenda": { const r = await one(d.sub === "pengumuman" ? "SELECT judul FROM pengumuman WHERE id = ?" : "SELECT judul FROM agenda WHERE id = ?"); return r ? cut(r.judul) : ""; }
      case "inventaris": { const r = await one("SELECT nama FROM inventaris WHERE id = ?"); return r ? cut(r.nama) : ""; }
      case "users": { const r = await one("SELECT name FROM users WHERE id = ?"); return r ? cut(r.name) : ""; }
      case "ramadan": { if (d.sub !== "kegiatan") return ""; const r = await one("SELECT jenis, petugas FROM ramadan_kegiatan WHERE id = ?"); return r ? `${JENIS_KEGIATAN[r.jenis] || r.jenis} · ${cut(r.petugas, 40)}` : ""; }
      case "muzakki": { const r = await one("SELECT nama_muzakki FROM muzakki WHERE id = ?"); return r ? cut(r.nama_muzakki, 40) : ""; }
      case "mustahik": { const r = await one("SELECT nama FROM mustahik WHERE id = ?"); return r ? cut(r.nama) : ""; }
      case "rt": case "rt-rw": { const r = await one("SELECT nomor_rt FROM rt WHERE id = ?"); return r ? cut(r.nomor_rt) : ""; }
      case "rw": { const r = await one("SELECT nomor_rw FROM rw WHERE id = ?"); return r ? cut(r.nomor_rw) : ""; }
      case "distribusi": { const r = await one("SELECT m.nama, d.jumlah, d.jenis_zakat FROM distribusi_zakat d JOIN mustahik m ON m.id = d.mustahik_id WHERE d.id = ?"); return r ? `${cut(r.nama, 40)} · ${r.jenis_zakat === "uang" ? rp(r.jumlah) : r.jumlah + " kg beras"}` : ""; }
      case "infak": { const r = await one("SELECT jumlah FROM infak WHERE id = ?"); return r ? rp(r.jumlah) : ""; }
      default: return "";
    }
  } catch (error) {
    return "";
  }
}

async function logAktivitas(db, { user, aksi, modul, objekId = null, ringkasan, ip = null }) {
  try {
    await db.execute(
      "INSERT INTO log_aktivitas (user_id, user_name, aksi, modul, objek_id, ringkasan, ip) VALUES (?, ?, ?, ?, ?, ?, ?)",
      [user ? user.id : null, user ? String(user.name || "").slice(0, 100) : null, aksi.slice(0, 40), modul.slice(0, 40), objekId, ringkasan.slice(0, 255), ip ? String(ip).slice(0, 64) : null]
    );
  } catch (error) {
    console.error("Gagal mencatat log aktivitas:", error.message);
  }
}

function auditMiddleware(db) {
  return async (req, res, next) => {
    const original = (req.method || "GET").toUpperCase();
    if (original === "GET" || original === "HEAD" || original === "OPTIONS") return next();
    const originalUrl = req.originalUrl || req.url;
    const path = originalUrl.split("?")[0];
    if (!req.session || !req.session.user || path.startsWith("/auth")) return next();
    if (path === "/ramadan/kegiatan/inline") return next(); // route ini mencatat log sendiri (lebih rinci)

    const override = (originalUrl.match(/[?&]_method=([A-Za-z]+)/) || [])[1];
    const d = describe((override || original).toUpperCase(), path);
    if (!d) return next();
    const sebelum = d.objekId ? await existingLabel(db, d) : "";

    res.on("finish", () => {
      const user = req.session && req.session.user;
      if (!user || res.statusCode >= 400) return;
      // redirect yang membawa pesan galat (validasi gagal) tidak dihitung sebagai perubahan
      const flash = (req.session && req.session.flash) || {};
      if ((flash.error && flash.error.length) || (flash.error_msg && flash.error_msg.length)) return;

      const detail = d.aksi === "Menghapus" ? sebelum : detailFromBody(d, req.body) || sebelum;
      const ringkasan = `${d.aksi} ${d.objek}${detail ? ": " + detail : ""}`;
      logAktivitas(db, { user, aksi: d.aksi, modul: d.modul, objekId: d.objekId, ringkasan, ip: req.ip });
    });
    next();
  };
}

module.exports = { logAktivitas, auditMiddleware, describe };
