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

    console.log("Selesai: skema sudah sinkron.");
  } finally {
    await db.end();
  }
}

main().catch((error) => {
  console.error("Migrasi gagal:", error.message);
  process.exit(1);
});
