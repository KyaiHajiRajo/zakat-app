/**
 * Migration: sinkronkan database lama (database.sql) ke skema yang dipakai kode saat ini.
 *
 * Menambahkan (semua idempoten, aman dijalankan berulang):
 * - tabel `rw` + kolom `rt.rw_id`
 * - tabel `master_zakat` (+ 2 data bawaan)
 * - `muzakki`: nama_muzakki, bin_binti, nama_orang_tua, master_zakat_id, tanggal
 * - `distribusi_zakat.bukti_foto`
 *
 * Usage: node migrations/sync-schema-current.js
 */
const mysql = require("mysql2/promise");
require("dotenv").config();

const config = {
  host: process.env.DB_HOST || "localhost",
  user: process.env.DB_USER || "root",
  password: process.env.DB_PASSWORD || process.env.DB_PASS || "",
  database: process.env.DB_NAME || "zakat_fitrah",
};

async function main() {
  const db = await mysql.createConnection(config);

  const exists = async (sql, params) => (await db.query(sql, params))[0].length > 0;
  const hasTable = (t) =>
    exists(
      "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?",
      [config.database, t]
    );
  const hasColumn = (t, c) =>
    exists(
      "SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? AND COLUMN_NAME = ?",
      [config.database, t, c]
    );
  const hasIndex = (t, i) =>
    exists(
      "SELECT 1 FROM INFORMATION_SCHEMA.STATISTICS WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? AND INDEX_NAME = ?",
      [config.database, t, i]
    );
  const hasConstraint = (t, c) =>
    exists(
      "SELECT 1 FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? AND CONSTRAINT_NAME = ?",
      [config.database, t, c]
    );
  const step = (msg) => console.log("•", msg);

  try {
    // --- rw ---
    if (!(await hasTable("rw"))) {
      await db.query(`
        CREATE TABLE rw (
          id INT NOT NULL AUTO_INCREMENT,
          nomor_rw VARCHAR(10) NOT NULL,
          ketua_rw VARCHAR(100) NOT NULL,
          keterangan VARCHAR(100) DEFAULT NULL,
          created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
          updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
          PRIMARY KEY (id),
          UNIQUE KEY nomor_rw (nomor_rw)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
      `);
      step("tabel rw dibuat");
    }

    const [[{ total: rwTotal }]] = await db.query("SELECT COUNT(*) AS total FROM rw");
    if (rwTotal === 0) {
      await db.query("INSERT INTO rw (nomor_rw, ketua_rw) VALUES ('RW 01', 'Budi Santoso')");
      step("data contoh RW 01 ditambahkan (ubah lewat menu Data RW/RT)");
    }

    if (!(await hasColumn("rt", "rw_id"))) {
      await db.query(`
        ALTER TABLE rt
          ADD COLUMN rw_id INT DEFAULT 1,
          ADD CONSTRAINT rt_ibfk_2 FOREIGN KEY (rw_id) REFERENCES rw (id)
      `);
      step("kolom rt.rw_id ditambahkan");
    }

    // --- master_zakat ---
    if (!(await hasTable("master_zakat"))) {
      await db.query(`
        CREATE TABLE master_zakat (
          id INT PRIMARY KEY AUTO_INCREMENT,
          nama VARCHAR(100) NOT NULL,
          harga DECIMAL(15, 2) NOT NULL DEFAULT 0,
          kg DECIMAL(10, 2) NOT NULL DEFAULT 0,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
          created_by INT NULL,
          updated_by INT NULL,
          FOREIGN KEY (created_by) REFERENCES users(id) ON DELETE SET NULL ON UPDATE CASCADE,
          FOREIGN KEY (updated_by) REFERENCES users(id) ON DELETE SET NULL ON UPDATE CASCADE,
          INDEX idx_nama (nama)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
      `);
      await db.query(`
        INSERT INTO master_zakat (nama, harga, kg, created_by, updated_by)
        SELECT 'Beras Standar', 45000, 2.5, MIN(id), MIN(id) FROM users
        UNION ALL
        SELECT 'Uang Standar', 45000, 0, MIN(id), MIN(id) FROM users
      `);
      step("tabel master_zakat dibuat dengan 2 data bawaan");
    }

    // --- muzakki ---
    if (!(await hasColumn("muzakki", "nama_muzakki"))) {
      await db.query("ALTER TABLE muzakki ADD COLUMN nama_muzakki VARCHAR(100) NULL AFTER rt_id");
      step("kolom muzakki.nama_muzakki ditambahkan");
    }
    const [filled] = await db.query(
      "UPDATE muzakki SET nama_muzakki = nama WHERE nama_muzakki IS NULL OR TRIM(nama_muzakki) = ''"
    );
    if (filled.affectedRows) step(`nama_muzakki diisi dari kolom nama (${filled.affectedRows} baris)`);

    // Kolom lama `nama` tidak lagi diisi aplikasi (nama disimpan di nama_muzakki), jadi harus boleh kosong.
    const [[legacyNama]] = await db.query(
      "SELECT IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = ? AND TABLE_NAME = 'muzakki' AND COLUMN_NAME = 'nama'",
      [config.database]
    );
    if (legacyNama && legacyNama.IS_NULLABLE === "NO") {
      await db.query("ALTER TABLE muzakki MODIFY COLUMN nama VARCHAR(100) NULL");
      step("kolom lama muzakki.nama dibuat boleh kosong");
    }

    if (!(await hasColumn("muzakki", "bin_binti"))) {
      await db.query("ALTER TABLE muzakki ADD COLUMN bin_binti ENUM('bin','binti') NULL AFTER nama_muzakki");
      step("kolom muzakki.bin_binti ditambahkan");
    }
    if (!(await hasColumn("muzakki", "nama_orang_tua"))) {
      await db.query("ALTER TABLE muzakki ADD COLUMN nama_orang_tua VARCHAR(150) NULL AFTER bin_binti");
      step("kolom muzakki.nama_orang_tua ditambahkan");
    }
    if (!(await hasColumn("muzakki", "master_zakat_id"))) {
      await db.query("ALTER TABLE muzakki ADD COLUMN master_zakat_id INT NULL AFTER jenis_zakat");
      step("kolom muzakki.master_zakat_id ditambahkan");
    }
    if (!(await hasIndex("muzakki", "idx_muzakki_master_zakat_id"))) {
      await db.query("ALTER TABLE muzakki ADD INDEX idx_muzakki_master_zakat_id (master_zakat_id)");
    }
    if (!(await hasConstraint("muzakki", "fk_muzakki_master_zakat"))) {
      await db.query(`
        ALTER TABLE muzakki
          ADD CONSTRAINT fk_muzakki_master_zakat FOREIGN KEY (master_zakat_id)
          REFERENCES master_zakat(id) ON DELETE SET NULL ON UPDATE CASCADE
      `);
      step("relasi muzakki -> master_zakat ditambahkan");
    }

    if (!(await hasColumn("muzakki", "tanggal"))) {
      await db.query("ALTER TABLE muzakki ADD COLUMN tanggal DATE NULL AFTER user_id");
      step("kolom muzakki.tanggal ditambahkan");
    }
    await db.query("UPDATE muzakki SET tanggal = DATE(created_at) WHERE tanggal IS NULL");
    await db.query("ALTER TABLE muzakki MODIFY COLUMN tanggal DATE NOT NULL");
    if (!(await hasIndex("muzakki", "idx_muzakki_tanggal"))) {
      await db.query("ALTER TABLE muzakki ADD INDEX idx_muzakki_tanggal (tanggal)");
    }

    // --- distribusi_zakat ---
    if (!(await hasColumn("distribusi_zakat", "bukti_foto"))) {
      await db.query("ALTER TABLE distribusi_zakat ADD COLUMN bukti_foto VARCHAR(255) NULL AFTER status");
      step("kolom distribusi_zakat.bukti_foto ditambahkan");
    }

    // --- modul masjid & Ramadan ---
    await db.query(`
      CREATE TABLE IF NOT EXISTS pengaturan (
        kunci VARCHAR(60) NOT NULL PRIMARY KEY,
        nilai TEXT NULL,
        updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);
    await db.query(`
      CREATE TABLE IF NOT EXISTS ramadan_kegiatan (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        tanggal DATE NOT NULL,
        jenis ENUM('imam_tarawih','bilal','penceramah','takjil','buka_bersama','tilawah','itikaf','lainnya') NOT NULL,
        petugas VARCHAR(120) NOT NULL,
        keterangan VARCHAR(255) NULL,
        jumlah INT NULL,
        user_id INT NULL,
        created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_ramadan_tanggal_jenis (tanggal, jenis)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);

    // jenis "bilal" ditambahkan ke jadwal petugas Ramadan (tabel lama tidak punya nilai ini)
    const [[jenisCol]] = await db.query(
      "SELECT COLUMN_TYPE AS t FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = ? AND TABLE_NAME = 'ramadan_kegiatan' AND COLUMN_NAME = 'jenis'",
      [config.database]
    );
    if (jenisCol && !String(jenisCol.t).includes("'bilal'")) {
      await db.query("ALTER TABLE ramadan_kegiatan MODIFY COLUMN jenis ENUM('imam_tarawih','bilal','penceramah','takjil','buka_bersama','tilawah','itikaf','lainnya') NOT NULL");
      step("jenis 'bilal' ditambahkan ke ramadan_kegiatan");
    }

    // --- kas masjid, agenda, pengumuman, inventaris ---
    await db.query(`
      CREATE TABLE IF NOT EXISTS kas_transaksi (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        tanggal DATE NOT NULL,
        tipe ENUM('masuk','keluar') NOT NULL,
        kategori VARCHAR(40) NOT NULL,
        jumlah DECIMAL(14,2) NOT NULL,
        pihak VARCHAR(120) NULL,
        keterangan VARCHAR(255) NULL,
        user_id INT NULL,
        created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_kas_tanggal (tanggal),
        INDEX idx_kas_tipe_kategori (tipe, kategori)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);
    await db.query(`
      CREATE TABLE IF NOT EXISTS agenda (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        judul VARCHAR(150) NOT NULL,
        jenis ENUM('pengajian','kerja_bakti','peringatan','rapat','sosial','ramadan','lainnya') NOT NULL DEFAULT 'pengajian',
        tanggal DATE NOT NULL,
        jam_mulai TIME NULL,
        jam_selesai TIME NULL,
        lokasi VARCHAR(150) NULL,
        pemateri VARCHAR(120) NULL,
        deskripsi TEXT NULL,
        tampil_publik TINYINT(1) NOT NULL DEFAULT 1,
        user_id INT NULL,
        created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_agenda_tanggal (tanggal)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);
    await db.query(`
      CREATE TABLE IF NOT EXISTS pengumuman (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        judul VARCHAR(150) NOT NULL,
        isi TEXT NOT NULL,
        penting TINYINT(1) NOT NULL DEFAULT 0,
        berlaku_sampai DATE NULL,
        tampil_publik TINYINT(1) NOT NULL DEFAULT 1,
        user_id INT NULL,
        created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_pengumuman_berlaku (berlaku_sampai)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);
    await db.query(`
      CREATE TABLE IF NOT EXISTS inventaris (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        nama VARCHAR(120) NOT NULL,
        kategori ENUM('perlengkapan_ibadah','elektronik','furnitur','kebersihan','konsumsi','lainnya') NOT NULL DEFAULT 'lainnya',
        jumlah INT NOT NULL DEFAULT 1,
        satuan VARCHAR(20) NOT NULL DEFAULT 'unit',
        kondisi ENUM('baik','perlu_perbaikan','rusak') NOT NULL DEFAULT 'baik',
        lokasi VARCHAR(100) NULL,
        tanggal_perolehan DATE NULL,
        sumber VARCHAR(100) NULL,
        catatan VARCHAR(255) NULL,
        created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_inventaris_kategori (kategori),
        INDEX idx_inventaris_kondisi (kondisi)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);

    // --- log aktivitas ---
    await db.query(`
      CREATE TABLE IF NOT EXISTS log_aktivitas (
        id BIGINT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        user_id INT NULL,
        user_name VARCHAR(100) NULL,
        aksi VARCHAR(40) NOT NULL,
        modul VARCHAR(40) NOT NULL,
        objek_id INT NULL,
        ringkasan VARCHAR(255) NOT NULL,
        ip VARCHAR(64) NULL,
        created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        INDEX idx_log_created (created_at),
        INDEX idx_log_modul (modul, created_at),
        INDEX idx_log_user (user_id, created_at)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);

    // --- donasi online (konfirmasi dari jamaah, diverifikasi pengurus lalu masuk ke kas) ---
    await db.query(`
      CREATE TABLE IF NOT EXISTS donasi (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        kode VARCHAR(12) NOT NULL,
        nama VARCHAR(100) NOT NULL,
        anonim TINYINT(1) NOT NULL DEFAULT 0,
        nomor_wa VARCHAR(20) NULL,
        jumlah DECIMAL(14,2) NOT NULL,
        tujuan VARCHAR(30) NOT NULL DEFAULT 'donasi',
        pesan VARCHAR(255) NULL,
        bukti VARCHAR(255) NULL,
        status ENUM('menunggu','diterima','ditolak') NOT NULL DEFAULT 'menunggu',
        alasan_tolak VARCHAR(255) NULL,
        kas_id INT NULL,
        diproses_oleh INT NULL,
        diproses_at DATETIME NULL,
        ip VARCHAR(64) NULL,
        created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE KEY uq_donasi_kode (kode),
        INDEX idx_donasi_status (status, created_at)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);

    // --- nomor WhatsApp petugas (sekali simpan, dipakai untuk semua hari) ---
    await db.query(`
      CREATE TABLE IF NOT EXISTS petugas_kontak (
        id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        nama_key VARCHAR(120) NOT NULL,
        nama VARCHAR(120) NOT NULL,
        nomor_wa VARCHAR(20) NOT NULL,
        updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        UNIQUE KEY uq_petugas_kontak (nama_key)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    `);

    // --- indeks untuk kecepatan (aman dijalankan berulang) ---
    const speedIndexes = [
      ["muzakki", "idx_muzakki_rt_tanggal", "(rt_id, tanggal)"],
      ["distribusi_zakat", "idx_distribusi_status_created", "(status, created_at)"],
      ["distribusi_zakat", "idx_distribusi_created", "(created_at)"],
      ["infak", "idx_infak_created", "(created_at)"],
      ["mustahik", "idx_mustahik_kategori", "(kategori)"],
      ["mustahik", "idx_mustahik_created", "(created_at)"],
    ];
    for (const [table, name, columns] of speedIndexes) {
      if (!(await hasIndex(table, name))) {
        await db.query(`ALTER TABLE ${table} ADD INDEX ${name} ${columns}`);
        step(`indeks ${name} ditambahkan`);
      }
    }

    console.log("Selesai: skema sudah sinkron.");
  } finally {
    await db.end();
  }
}

main().catch((error) => {
  console.error("Migrasi gagal:", error.message);
  process.exit(1);
});
