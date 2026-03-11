-- Schema SQL for moving identity fields from `muzakki_details` to `muzakki`
-- Safe full data migration (including row split + infak split) is implemented in:
--   migrations/move-muzakki-details-into-muzakki.js
--
-- This SQL file provides the DDL/query references for the target schema.

ALTER TABLE muzakki
  ADD COLUMN IF NOT EXISTS nama_muzakki VARCHAR(100) NULL AFTER rt_id,
  ADD COLUMN IF NOT EXISTS bin_binti ENUM('bin', 'binti') NULL AFTER nama_muzakki,
  ADD COLUMN IF NOT EXISTS nama_orang_tua VARCHAR(150) NULL AFTER bin_binti,
  ADD COLUMN IF NOT EXISTS master_zakat_id INT NULL COMMENT 'Relasi ke master_zakat untuk jenis zakat yang dipilih' AFTER jenis_zakat;

ALTER TABLE muzakki
  ADD INDEX idx_muzakki_master_zakat_id (master_zakat_id);

ALTER TABLE muzakki
  ADD CONSTRAINT fk_muzakki_master_zakat
  FOREIGN KEY (master_zakat_id) REFERENCES master_zakat(id)
  ON DELETE SET NULL
  ON UPDATE CASCADE;

-- Reference backfill for rows that only have one detail row.
UPDATE muzakki m
INNER JOIN (
  SELECT
    md.muzakki_id,
    md.nama_muzakki,
    md.bin_binti,
    md.nama_orang_tua,
    md.master_zakat_id
  FROM muzakki_details md
  INNER JOIN (
    SELECT muzakki_id, MIN(id) AS first_detail_id, COUNT(*) AS detail_count
    FROM muzakki_details
    GROUP BY muzakki_id
  ) summary
    ON summary.muzakki_id = md.muzakki_id
   AND summary.first_detail_id = md.id
  WHERE summary.detail_count = 1
) src ON src.muzakki_id = m.id
SET
  m.nama_muzakki = src.nama_muzakki,
  m.bin_binti = src.bin_binti,
  m.nama_orang_tua = src.nama_orang_tua,
  m.master_zakat_id = src.master_zakat_id;

-- Verification queries
SELECT COUNT(*) AS total_muzakki FROM muzakki;
SELECT COUNT(*) AS total_named_muzakki FROM muzakki WHERE nama_muzakki IS NOT NULL AND TRIM(nama_muzakki) <> '';
SELECT COUNT(*) AS total_muzakki_details FROM muzakki_details;

-- IMPORTANT:
-- Use the JS migration to safely split rows with multiple muzakki_details and preserve infak totals:
--   node migrations/move-muzakki-details-into-muzakki.js
--
-- Final cleanup handled by the JS migration:
--   DROP TABLE muzakki_details;
