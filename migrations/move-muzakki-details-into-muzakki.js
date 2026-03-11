/**
 * Migration: move muzakki_details into muzakki
 *
 * Usage:
 *   node migrations/move-muzakki-details-into-muzakki.js
 *   node migrations/move-muzakki-details-into-muzakki.js down
 *
 * Notes:
 * - Setelah migration ini, satu baris `muzakki` mewakili satu nama muzakki.
 * - Data lama yang punya banyak baris di `muzakki_details` akan di-split menjadi
 *   beberapa baris `muzakki`, sambil menjaga total jumlah_jiwa, pembayaran,
 *   kembalian, dan infak tetap sama secara agregat.
 */

const mysql = require("mysql2/promise");
require("dotenv").config();

function getDbConfig() {
  return {
    host: process.env.DB_HOST || "localhost",
    user: process.env.DB_USER || "root",
    password: process.env.DB_PASSWORD || process.env.DB_PASS || "",
    database: process.env.DB_NAME || "zakat",
    multipleStatements: true,
  };
}

async function tableExists(connection, tableName) {
  const [rows] = await connection.execute(
    `
    SELECT 1
    FROM INFORMATION_SCHEMA.TABLES
    WHERE TABLE_SCHEMA = ?
      AND TABLE_NAME = ?
    LIMIT 1
    `,
    [getDbConfig().database, tableName]
  );

  return rows.length > 0;
}

async function columnExists(connection, tableName, columnName) {
  const [rows] = await connection.execute(
    `
    SELECT 1
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_SCHEMA = ?
      AND TABLE_NAME = ?
      AND COLUMN_NAME = ?
    LIMIT 1
    `,
    [getDbConfig().database, tableName, columnName]
  );

  return rows.length > 0;
}

async function indexExists(connection, tableName, indexName) {
  const [rows] = await connection.execute(
    `
    SELECT 1
    FROM INFORMATION_SCHEMA.STATISTICS
    WHERE TABLE_SCHEMA = ?
      AND TABLE_NAME = ?
      AND INDEX_NAME = ?
    LIMIT 1
    `,
    [getDbConfig().database, tableName, indexName]
  );

  return rows.length > 0;
}

async function foreignKeyName(connection, tableName, columnName, referencedTable) {
  const [rows] = await connection.execute(
    `
    SELECT CONSTRAINT_NAME
    FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
    WHERE TABLE_SCHEMA = ?
      AND TABLE_NAME = ?
      AND COLUMN_NAME = ?
      AND REFERENCED_TABLE_NAME = ?
    LIMIT 1
    `,
    [getDbConfig().database, tableName, columnName, referencedTable]
  );

  return rows[0]?.CONSTRAINT_NAME || null;
}

function splitDecimal(value, parts, decimals = 2) {
  if (value === null || value === undefined) {
    return Array(parts).fill(null);
  }

  const factor = 10 ** decimals;
  const total = Math.round(Number(value || 0) * factor);
  const base = Math.floor(total / parts);
  let remainder = total - base * parts;

  return Array.from({ length: parts }, (_, index) => {
    const share = base + (index === parts - 1 ? remainder : 0);
    remainder = index === parts - 1 ? 0 : remainder;
    return share / factor;
  });
}

function splitInteger(value, parts) {
  const total = Number.isFinite(Number(value)) ? parseInt(value, 10) : 0;

  if (total <= 0) {
    return Array(parts).fill(1);
  }

  const base = Math.floor(total / parts);
  let remainder = total - base * parts;

  return Array.from({ length: parts }, (_, index) => {
    const share = base + (remainder > 0 ? 1 : 0);
    if (remainder > 0) remainder -= 1;
    return share;
  });
}

async function ensureTargetColumns(connection) {
  const additions = [
    {
      name: "nama_muzakki",
      sql: `
        ALTER TABLE muzakki
        ADD COLUMN nama_muzakki VARCHAR(100) NULL AFTER rt_id
      `,
    },
    {
      name: "bin_binti",
      sql: `
        ALTER TABLE muzakki
        ADD COLUMN bin_binti ENUM('bin', 'binti') NULL AFTER nama_muzakki
      `,
    },
    {
      name: "nama_orang_tua",
      sql: `
        ALTER TABLE muzakki
        ADD COLUMN nama_orang_tua VARCHAR(150) NULL AFTER bin_binti
      `,
    },
    {
      name: "master_zakat_id",
      sql: `
        ALTER TABLE muzakki
        ADD COLUMN master_zakat_id INT NULL
          COMMENT 'Relasi ke master_zakat untuk jenis zakat yang dipilih'
          AFTER jenis_zakat
      `,
    },
  ];

  for (const addition of additions) {
    if (!(await columnExists(connection, "muzakki", addition.name))) {
      console.log(`Adding column muzakki.${addition.name}...`);
      await connection.execute(addition.sql);
    }
  }
}

async function splitInfakRows(connection, originalMuzakkiId, targetIds) {
  const [infakRows] = await connection.execute(
    `
    SELECT id, jumlah, keterangan, created_at, updated_at
    FROM infak
    WHERE muzakki_id = ?
    ORDER BY id ASC
    `,
    [originalMuzakkiId]
  );

  for (const infak of infakRows) {
    const splitJumlah = splitDecimal(infak.jumlah, targetIds.length, 2);

    await connection.execute(
      `
      UPDATE infak
      SET muzakki_id = ?, jumlah = ?
      WHERE id = ?
      `,
      [targetIds[0], splitJumlah[0], infak.id]
    );

    for (let index = 1; index < targetIds.length; index += 1) {
      await connection.execute(
        `
        INSERT INTO infak (muzakki_id, jumlah, keterangan, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?)
        `,
        [
          targetIds[index],
          splitJumlah[index],
          infak.keterangan,
          infak.created_at,
          infak.updated_at,
        ]
      );
    }
  }
}

async function migrateDetailsIntoMuzakki(connection) {
  if (!(await tableExists(connection, "muzakki_details"))) {
    console.log("Table muzakki_details not found. Skipping data migration.");
    return;
  }

  const [muzakkiRows] = await connection.execute(
    `
    SELECT *
    FROM muzakki
    ORDER BY id ASC
    `
  );

  for (const row of muzakkiRows) {
    const [details] = await connection.execute(
      `
      SELECT *
      FROM muzakki_details
      WHERE muzakki_id = ?
      ORDER BY id ASC
      `,
      [row.id]
    );

    if (details.length === 0) {
      continue;
    }

    if (details.length === 1) {
      const detail = details[0];
      await connection.execute(
        `
        UPDATE muzakki
        SET nama_muzakki = ?,
            bin_binti = ?,
            nama_orang_tua = ?,
            master_zakat_id = COALESCE(master_zakat_id, ?)
        WHERE id = ?
        `,
        [
          detail.nama_muzakki || null,
          detail.bin_binti || null,
          detail.nama_orang_tua || null,
          detail.master_zakat_id || null,
          row.id,
        ]
      );

      continue;
    }

    const jumlahJiwaSplit = splitInteger(row.jumlah_jiwa || details.length, details.length);
    const jumlahBerasSplit = splitDecimal(row.jumlah_beras_kg, details.length, 2);
    const jumlahUangSplit = splitDecimal(row.jumlah_uang, details.length, 2);
    const jumlahBayarSplit = splitDecimal(row.jumlah_bayar, details.length, 2);
    const kembalianSplit = splitDecimal(row.kembalian, details.length, 2);

    const targetIds = [row.id];

    for (let index = 0; index < details.length; index += 1) {
      const detail = details[index];
      const values = [
        detail.nama_muzakki || null,
        detail.bin_binti || null,
        detail.nama_orang_tua || null,
        detail.master_zakat_id || null,
        jumlahJiwaSplit[index],
        jumlahBerasSplit[index],
        jumlahUangSplit[index],
        jumlahBayarSplit[index],
        kembalianSplit[index],
      ];

      if (index === 0) {
        await connection.execute(
          `
          UPDATE muzakki
          SET nama_muzakki = ?,
              bin_binti = ?,
              nama_orang_tua = ?,
              master_zakat_id = ?,
              jumlah_jiwa = ?,
              jumlah_beras_kg = ?,
              jumlah_uang = ?,
              jumlah_bayar = ?,
              kembalian = ?
          WHERE id = ?
          `,
          [...values, row.id]
        );
      } else {
        const [inserted] = await connection.execute(
          `
          INSERT INTO muzakki (
            rt_id,
            nama_muzakki,
            bin_binti,
            nama_orang_tua,
            jumlah_jiwa,
            jenis_zakat,
            master_zakat_id,
            jumlah_beras_kg,
            jumlah_uang,
            jumlah_bayar,
            kembalian,
            catatan,
            user_id,
            tanggal,
            created_at,
            updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            row.rt_id,
            detail.nama_muzakki || null,
            detail.bin_binti || null,
            detail.nama_orang_tua || null,
            jumlahJiwaSplit[index],
            row.jenis_zakat,
            detail.master_zakat_id || null,
            jumlahBerasSplit[index],
            jumlahUangSplit[index],
            jumlahBayarSplit[index],
            kembalianSplit[index],
            row.catatan,
            row.user_id,
            row.tanggal,
            row.created_at,
            row.updated_at,
          ]
        );

        targetIds.push(inserted.insertId);
      }
    }

    await splitInfakRows(connection, row.id, targetIds);
  }
}

async function ensureMasterZakatRelation(connection) {
  if (!(await indexExists(connection, "muzakki", "idx_muzakki_master_zakat_id"))) {
    console.log("Adding index idx_muzakki_master_zakat_id...");
    await connection.execute(`
      ALTER TABLE muzakki
      ADD INDEX idx_muzakki_master_zakat_id (master_zakat_id)
    `);
  }

  const fkName = await foreignKeyName(
    connection,
    "muzakki",
    "master_zakat_id",
    "master_zakat"
  );

  if (!fkName) {
    console.log("Adding foreign key fk_muzakki_master_zakat...");
    await connection.execute(`
      ALTER TABLE muzakki
      ADD CONSTRAINT fk_muzakki_master_zakat
      FOREIGN KEY (master_zakat_id) REFERENCES master_zakat(id)
      ON DELETE SET NULL
      ON UPDATE CASCADE
    `);
  }
}

async function up() {
  const connection = await mysql.createConnection(getDbConfig());

  try {
    console.log("Starting migration: move muzakki_details into muzakki");
    await connection.beginTransaction();

    await ensureTargetColumns(connection);
    await migrateDetailsIntoMuzakki(connection);
    await ensureMasterZakatRelation(connection);

    if (await tableExists(connection, "muzakki_details")) {
      console.log("Dropping table muzakki_details...");
      await connection.execute("DROP TABLE muzakki_details");
    }

    await connection.commit();
    console.log("Migration completed successfully.");
  } catch (error) {
    await connection.rollback();
    console.error("Migration failed:", error.message);
    throw error;
  } finally {
    await connection.end();
  }
}

async function down() {
  const connection = await mysql.createConnection(getDbConfig());

  try {
    console.log("Rolling back migration: recreate muzakki_details");
    await connection.beginTransaction();

    if (!(await tableExists(connection, "muzakki_details"))) {
      await connection.execute(`
        CREATE TABLE muzakki_details (
          id INT NOT NULL AUTO_INCREMENT,
          muzakki_id INT NOT NULL,
          nama_muzakki VARCHAR(100) NULL,
          bin_binti ENUM('bin', 'binti') NULL,
          nama_orang_tua VARCHAR(150) NULL,
          master_zakat_id INT NULL COMMENT 'Relasi ke master_zakat untuk jenis zakat yang dipilih',
          created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
          updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
          PRIMARY KEY (id),
          KEY idx_muzakki_details_muzakki_id (muzakki_id),
          KEY idx_muzakki_details_master_zakat_id (master_zakat_id),
          CONSTRAINT fk_muzakki_details_muzakki
            FOREIGN KEY (muzakki_id) REFERENCES muzakki(id)
            ON DELETE CASCADE
            ON UPDATE CASCADE,
          CONSTRAINT fk_muzakki_details_master_zakat
            FOREIGN KEY (master_zakat_id) REFERENCES master_zakat(id)
            ON DELETE SET NULL
            ON UPDATE CASCADE
        )
      `);
    }

    await connection.execute("DELETE FROM muzakki_details");

    await connection.execute(`
      INSERT INTO muzakki_details (
        muzakki_id,
        nama_muzakki,
        bin_binti,
        nama_orang_tua,
        master_zakat_id,
        created_at,
        updated_at
      )
      SELECT
        id,
        nama_muzakki,
        bin_binti,
        nama_orang_tua,
        master_zakat_id,
        created_at,
        updated_at
      FROM muzakki
      WHERE nama_muzakki IS NOT NULL
        AND TRIM(nama_muzakki) <> ''
    `);

    const fkName = await foreignKeyName(
      connection,
      "muzakki",
      "master_zakat_id",
      "master_zakat"
    );

    if (fkName) {
      await connection.execute(
        `ALTER TABLE muzakki DROP FOREIGN KEY \`${fkName}\``
      );
    }

    if (await indexExists(connection, "muzakki", "idx_muzakki_master_zakat_id")) {
      await connection.execute(`
        ALTER TABLE muzakki
        DROP INDEX idx_muzakki_master_zakat_id
      `);
    }

    const removableColumns = [
      "master_zakat_id",
      "nama_orang_tua",
      "bin_binti",
      "nama_muzakki",
    ];

    for (const columnName of removableColumns) {
      if (await columnExists(connection, "muzakki", columnName)) {
        await connection.execute(
          `ALTER TABLE muzakki DROP COLUMN \`${columnName}\``
        );
      }
    }

    await connection.commit();
    console.log("Rollback completed successfully.");
  } catch (error) {
    await connection.rollback();
    console.error("Rollback failed:", error.message);
    throw error;
  } finally {
    await connection.end();
  }
}

if (require.main === module) {
  const command = process.argv[2];
  const runner = command === "down" || command === "rollback" ? down : up;

  runner()
    .then(() => process.exit(0))
    .catch(() => process.exit(1));
}

module.exports = { up, down };
