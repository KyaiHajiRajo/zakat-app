const express = require("express");

const router = express.Router();

const REPORT_YEAR = 2026;

function getMuzakkiTanggalSql(alias = "m") {
  return `COALESCE(${alias}.tanggal, DATE(${alias}.created_at))`;
}

router.get("/zakat-fitrah-2026", async (req, res) => {
  try {
    const db = req.app.locals.db;

    const [statsRows] = await db.execute(
      `
      SELECT
        COUNT(*) as total_data_muzakki,
        COUNT(CASE
          WHEN m.nama_muzakki IS NOT NULL AND TRIM(m.nama_muzakki) <> ''
          THEN 1
        END) as total_nama_muzakki,
        COALESCE(SUM(COALESCE(m.jumlah_bayar, 0)), 0) as total_jumlah_bayar,
        COALESCE(SUM(COALESCE(m.jumlah_beras_kg, 0)), 0) as total_jumlah_beras_kg,
        COUNT(CASE WHEN m.jenis_zakat = 'uang' THEN 1 END) as total_muzakki_uang,
        COUNT(CASE WHEN m.jenis_zakat = 'beras' THEN 1 END) as total_muzakki_beras,
        COUNT(DISTINCT CASE
          WHEN m.master_zakat_id IS NOT NULL THEN m.master_zakat_id
        END) as total_master_zakat_aktif,
        MIN(${getMuzakkiTanggalSql("m")}) as periode_awal,
        MAX(${getMuzakkiTanggalSql("m")}) as periode_akhir
      FROM muzakki m
      WHERE YEAR(${getMuzakkiTanggalSql("m")}) = ?
      `,
      [REPORT_YEAR]
    );

    const [reportRows] = await db.execute(
      `
      SELECT
        m.id,
        COALESCE(NULLIF(TRIM(m.nama_muzakki), ''), CONCAT('Muzakki #', m.id)) as nama_muzakki,
        COALESCE(NULLIF(TRIM(r.nomor_rt), ''), '-') as nomor_rt,
        m.master_zakat_id,
        CASE
          WHEN m.master_zakat_id IS NULL THEN '-'
          ELSE COALESCE(
            NULLIF(TRIM(mz.nama), ''),
            CONCAT('Master Zakat #', m.master_zakat_id)
          )
        END as master_zakat_nama,
        m.jenis_zakat,
        COALESCE(m.jumlah_beras_kg, 0) as jumlah_beras_kg,
        COALESCE(m.jumlah_bayar, 0) as jumlah_bayar,
        DATE_FORMAT(${getMuzakkiTanggalSql("m")}, '%Y-%m-%d') as tanggal
      FROM muzakki m
      LEFT JOIN rt r ON r.id = m.rt_id
      LEFT JOIN master_zakat mz ON mz.id = m.master_zakat_id
      WHERE YEAR(${getMuzakkiTanggalSql("m")}) = ?
      ORDER BY ${getMuzakkiTanggalSql("m")} ASC, m.created_at ASC, m.id ASC
      `,
      [REPORT_YEAR]
    );

    const [mustahikRows] = await db.execute(
      `
      SELECT
        ms.id,
        COALESCE(NULLIF(TRIM(ms.nama), ''), CONCAT('Mustahik #', ms.id)) as nama,
        COALESCE(NULLIF(TRIM(r.nomor_rt), ''), '-') as nomor_rt,
        ms.kategori
      FROM mustahik ms
      LEFT JOIN rt r ON r.id = ms.rt_id
      ORDER BY ms.nama ASC, ms.id ASC
      `
    );

    res.render("public/zakat-fitrah-2026", {
      title: `Report Zakat Fitrah ${REPORT_YEAR}`,
      reportYear: REPORT_YEAR,
      stats: statsRows[0] || {},
      reportRows: reportRows || [],
      mustahikRows: mustahikRows || [],
      generatedAt: new Date(),
    });
  } catch (error) {
    console.error("Public report error:", error);
    res.status(500).render("error", {
      title: "Laporan Publik Error",
      message: "Terjadi kesalahan saat memuat laporan publik zakat fitrah.",
      error: process.env.NODE_ENV === "development" ? error : {},
    });
  }
});

module.exports = router;
