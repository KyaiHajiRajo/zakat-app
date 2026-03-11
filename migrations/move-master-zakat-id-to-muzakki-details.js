/**
 * Migration: move master_zakat_id from muzakki to muzakki_details
 *
 * Usage:
 *   node migrations/move-master-zakat-id-to-muzakki-details.js
 *   node migrations/move-master-zakat-id-to-muzakki-details.js down
 */

const mysql = require("mysql2/promise");

const dbConfig = {
  host: process.env.DB_HOST || "localhost",
  user: process.env.DB_USER || "root",
  password: process.env.DB_PASSWORD || "",
  database: process.env.DB_NAME || "zakat",
  multipleStatements: true,
};

async function columnExists(connection, tableName, columnName) {
  const [rows] = await connection.query(
    `
      SELECT 1
      FROM INFORMATION_SCHEMA.COLUMNS
      WHERE TABLE_SCHEMA = ?
        AND TABLE_NAME = ?
        AND COLUMN_NAME = ?
      LIMIT 1
    `,
    [dbConfig.database, tableName, columnName]
  );

  return rows.length > 0;
}

async function foreignKeyName(connection, tableName, columnName, referencedTable) {
  const [rows] = await connection.query(
    `
      SELECT CONSTRAINT_NAME
      FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
      WHERE TABLE_SCHEMA = ?
        AND TABLE_NAME = ?
        AND COLUMN_NAME = ?
        AND REFERENCED_TABLE_NAME = ?
      LIMIT 1
    `,
    [dbConfig.database, tableName, columnName, referencedTable]
  );

  return rows[0]?.CONSTRAINT_NAME || null;
}

async function indexExists(connection, tableName, indexName) {
  const [rows] = await connection.query(
    `
      SELECT 1
      FROM INFORMATION_SCHEMA.STATISTICS
      WHERE TABLE_SCHEMA = ?
        AND TABLE_NAME = ?
        AND INDEX_NAME = ?
      LIMIT 1
    `,
    [dbConfig.database, tableName, indexName]
  );

  return rows.length > 0;
}

async function up() {
  let connection;

  try {
    connection = await mysql.createConnection(dbConfig);
    console.log("Starting migration: move master_zakat_id to muzakki_details");

    await connection.beginTransaction();

    if (!(await columnExists(connection, "muzakki_details", "master_zakat_id"))) {
      console.log("Adding master_zakat_id column to muzakki_details...");
      await connection.query(`
        ALTER TABLE muzakki_details
        ADD COLUMN master_zakat_id INT NULL
          COMMENT 'Relasi ke master_zakat untuk jenis zakat yang dipilih'
          AFTER nama_orang_tua
      `);
    }

    console.log("Backfilling muzakki_details.master_zakat_id from muzakki...");
    await connection.query(`
      UPDATE muzakki_details md
      INNER JOIN muzakki m ON m.id = md.muzakki_id
      SET md.master_zakat_id = m.master_zakat_id
      WHERE md.master_zakat_id IS NULL
        AND m.master_zakat_id IS NOT NULL
    `);

    if (!(await indexExists(connection, "muzakki_details", "idx_muzakki_details_master_zakat_id"))) {
      console.log("Adding index on muzakki_details.master_zakat_id...");
      await connection.query(`
        ALTER TABLE muzakki_details
        ADD INDEX idx_muzakki_details_master_zakat_id (master_zakat_id)
      `);
    }

    const detailsForeignKey = await foreignKeyName(
      connection,
      "muzakki_details",
      "master_zakat_id",
      "master_zakat"
    );

    if (!detailsForeignKey) {
      console.log("Adding foreign key on muzakki_details.master_zakat_id...");
      await connection.query(`
        ALTER TABLE muzakki_details
        ADD CONSTRAINT fk_muzakki_details_master_zakat
        FOREIGN KEY (master_zakat_id) REFERENCES master_zakat(id)
        ON DELETE SET NULL
        ON UPDATE CASCADE
      `);
    }

    const muzakkiForeignKey = await foreignKeyName(
      connection,
      "muzakki",
      "master_zakat_id",
      "master_zakat"
    );

    if (muzakkiForeignKey) {
      console.log("Dropping foreign key from muzakki.master_zakat_id...");
      await connection.query(
        `ALTER TABLE muzakki DROP FOREIGN KEY \`${muzakkiForeignKey}\``
      );
    }

    if (await columnExists(connection, "muzakki", "master_zakat_id")) {
      console.log("Dropping master_zakat_id column from muzakki...");
      await connection.query(`
        ALTER TABLE muzakki
        DROP COLUMN master_zakat_id
      `);
    }

    await connection.commit();
    console.log("Migration completed successfully.");
  } catch (error) {
    if (connection) {
      await connection.rollback();
    }

    console.error("Migration failed:", error.message);
    throw error;
  } finally {
    if (connection) {
      await connection.end();
    }
  }
}

async function down() {
  let connection;

  try {
    connection = await mysql.createConnection(dbConfig);
    console.log("Rolling back migration: move master_zakat_id to muzakki_details");

    await connection.beginTransaction();

    if (!(await columnExists(connection, "muzakki", "master_zakat_id"))) {
      console.log("Adding master_zakat_id column back to muzakki...");
      await connection.query(`
        ALTER TABLE muzakki
        ADD COLUMN master_zakat_id INT NULL
          COMMENT 'Relasi ke master_zakat untuk jenis zakat yang dipilih'
          AFTER jenis_zakat
      `);
    }

    console.log("Backfilling muzakki.master_zakat_id from muzakki_details...");
    await connection.query(`
      UPDATE muzakki m
      INNER JOIN (
        SELECT muzakki_id, MAX(master_zakat_id) AS master_zakat_id
        FROM muzakki_details
        WHERE master_zakat_id IS NOT NULL
        GROUP BY muzakki_id
      ) md ON md.muzakki_id = m.id
      SET m.master_zakat_id = md.master_zakat_id
      WHERE m.master_zakat_id IS NULL
    `);

    if (!(await indexExists(connection, "muzakki", "idx_master_zakat_id"))) {
      console.log("Adding index on muzakki.master_zakat_id...");
      await connection.query(`
        ALTER TABLE muzakki
        ADD INDEX idx_master_zakat_id (master_zakat_id)
      `);
    }

    const muzakkiForeignKey = await foreignKeyName(
      connection,
      "muzakki",
      "master_zakat_id",
      "master_zakat"
    );

    if (!muzakkiForeignKey) {
      console.log("Adding foreign key on muzakki.master_zakat_id...");
      await connection.query(`
        ALTER TABLE muzakki
        ADD CONSTRAINT fk_muzakki_master_zakat
        FOREIGN KEY (master_zakat_id) REFERENCES master_zakat(id)
        ON DELETE SET NULL
        ON UPDATE CASCADE
      `);
    }

    const detailsForeignKey = await foreignKeyName(
      connection,
      "muzakki_details",
      "master_zakat_id",
      "master_zakat"
    );

    if (detailsForeignKey) {
      console.log("Dropping foreign key from muzakki_details.master_zakat_id...");
      await connection.query(
        `ALTER TABLE muzakki_details DROP FOREIGN KEY \`${detailsForeignKey}\``
      );
    }

    if (await columnExists(connection, "muzakki_details", "master_zakat_id")) {
      console.log("Dropping master_zakat_id column from muzakki_details...");
      await connection.query(`
        ALTER TABLE muzakki_details
        DROP COLUMN master_zakat_id
      `);
    }

    await connection.commit();
    console.log("Rollback completed successfully.");
  } catch (error) {
    if (connection) {
      await connection.rollback();
    }

    console.error("Rollback failed:", error.message);
    throw error;
  } finally {
    if (connection) {
      await connection.end();
    }
  }
}

if (require.main === module) {
  const command = process.argv[2];

  if (command === "down" || command === "rollback") {
    down()
      .then(() => process.exit(0))
      .catch((error) => {
        console.error(error);
        process.exit(1);
      });
  } else {
    up()
      .then(() => process.exit(0))
      .catch((error) => {
        console.error(error);
        process.exit(1);
      });
  }
}

module.exports = { up, down };
