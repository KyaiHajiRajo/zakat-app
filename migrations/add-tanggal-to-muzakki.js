/**
 * Migration: add tanggal column to muzakki table
 * Created: 2026-03-10
 * Description:
 * - Menambahkan kolom tanggal untuk tanggal pembayaran zakat
 * - Mengisi data lama dari DATE(created_at)
 * - Menambahkan index untuk kebutuhan grouping dan filter
 */

const mysql = require("mysql2/promise");
require("dotenv").config();

function getDbConfig() {
  return {
    host: process.env.DB_HOST || "localhost",
    user: process.env.DB_USER || "root",
    password: process.env.DB_PASSWORD || process.env.DB_PASS || "",
    database: process.env.DB_NAME || "zakat",
  };
}

async function up() {
  const connection = await mysql.createConnection(getDbConfig());

  try {
    console.log("Starting migration: add tanggal column to muzakki");
    await connection.beginTransaction();

    const [columns] = await connection.execute(
      `
      SELECT COLUMN_NAME
      FROM INFORMATION_SCHEMA.COLUMNS
      WHERE TABLE_SCHEMA = ?
        AND TABLE_NAME = 'muzakki'
        AND COLUMN_NAME = 'tanggal'
      `,
      [getDbConfig().database]
    );

    if (columns.length === 0) {
      await connection.execute(`
        ALTER TABLE muzakki
        ADD COLUMN tanggal DATE NULL
        COMMENT 'Tanggal pembayaran zakat'
        AFTER user_id
      `);

      console.log("Column tanggal added");
    } else {
      console.log("Column tanggal already exists, skipping add column");
    }

    await connection.execute(`
      UPDATE muzakki
      SET tanggal = DATE(created_at)
      WHERE tanggal IS NULL
    `);

    await connection.execute(`
      ALTER TABLE muzakki
      MODIFY COLUMN tanggal DATE NOT NULL
      COMMENT 'Tanggal pembayaran zakat'
    `);

    const [indexes] = await connection.execute(
      `
      SELECT INDEX_NAME
      FROM INFORMATION_SCHEMA.STATISTICS
      WHERE TABLE_SCHEMA = ?
        AND TABLE_NAME = 'muzakki'
        AND INDEX_NAME = 'idx_muzakki_tanggal'
      `,
      [getDbConfig().database]
    );

    if (indexes.length === 0) {
      await connection.execute(`
        ALTER TABLE muzakki
        ADD INDEX idx_muzakki_tanggal (tanggal)
      `);

      console.log("Index idx_muzakki_tanggal added");
    } else {
      console.log("Index idx_muzakki_tanggal already exists, skipping");
    }

    await connection.commit();
    console.log("Migration completed successfully");
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
    console.log("Rollback migration: remove tanggal column from muzakki");
    await connection.beginTransaction();

    const [columns] = await connection.execute(
      `
      SELECT COLUMN_NAME
      FROM INFORMATION_SCHEMA.COLUMNS
      WHERE TABLE_SCHEMA = ?
        AND TABLE_NAME = 'muzakki'
        AND COLUMN_NAME = 'tanggal'
      `,
      [getDbConfig().database]
    );

    if (columns.length === 0) {
      console.log("Column tanggal does not exist, nothing to rollback");
      await connection.rollback();
      return;
    }

    const [indexes] = await connection.execute(
      `
      SELECT INDEX_NAME
      FROM INFORMATION_SCHEMA.STATISTICS
      WHERE TABLE_SCHEMA = ?
        AND TABLE_NAME = 'muzakki'
        AND INDEX_NAME = 'idx_muzakki_tanggal'
      `,
      [getDbConfig().database]
    );

    if (indexes.length > 0) {
      await connection.execute(`
        ALTER TABLE muzakki
        DROP INDEX idx_muzakki_tanggal
      `);
    }

    await connection.execute(`
      ALTER TABLE muzakki
      DROP COLUMN tanggal
    `);

    await connection.commit();
    console.log("Rollback completed successfully");
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
