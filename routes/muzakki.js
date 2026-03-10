const express = require("express");
const router = express.Router();

// Constants for zakat calculation
const ZAKAT_BERAS_PER_JIWA = 2.5; // kg per jiwa
const ZAKAT_UANG_PER_JIWA = 45000; // Rp per jiwa

function sanitizeForExcel(value) {
  if (value === null || value === undefined) return "";
  let safe = String(value);
  safe = safe.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");

  if (safe.length > 30000) {
    safe = `${safe.substring(0, 30000)}... (truncated)`;
  }

  return safe;
}

function normalizeExcelSheetName(name) {
  let sheetName = sanitizeForExcel(name).trim();

  if (!sheetName) {
    sheetName = "Sheet1";
  }

  sheetName = sheetName.replace(/[\/\\\?\*\[\]\:]/g, "_");

  if (sheetName.length > 31) {
    sheetName = sheetName.substring(0, 31);
  }

  return sheetName;
}

function formatDateForExcel(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return date.toLocaleDateString("id-ID");
}

function safeRtForFileName(value) {
  const safe = sanitizeForExcel(value)
    .replace(/[^a-zA-Z0-9_-]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");

  return safe || "unknown";
}

const EXCEL_NUMFMT = {
  INTEGER: "#,##0",
  WEIGHT: "#,##0.00",
  CURRENCY_RP: '[$Rp-421] #,##0',
};

function applyMuzakkiExcelNumberFormat(row) {
  row.getCell(3).numFmt = EXCEL_NUMFMT.INTEGER; // Jumlah jiwa
  row.getCell(5).numFmt = EXCEL_NUMFMT.WEIGHT; // Beras (kg)
  row.getCell(6).numFmt = EXCEL_NUMFMT.CURRENCY_RP; // Jumlah uang
  row.getCell(7).numFmt = EXCEL_NUMFMT.CURRENCY_RP; // Jumlah bayar
  row.getCell(8).numFmt = EXCEL_NUMFMT.CURRENCY_RP; // Kembalian
}

function getMuzakkiTanggalSql(alias = "m") {
  return `COALESCE(${alias}.tanggal, DATE(${alias}.created_at))`;
}

function getMuzakkiTanggalSelectSql(alias = "m") {
  return `DATE_FORMAT(${getMuzakkiTanggalSql(alias)}, '%Y-%m-%d')`;
}

function normalizeMuzakkiDateInput(value) {
  if (!value) return null;

  const normalized = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    return null;
  }

  const [year, month, day] = normalized.split("-").map(Number);
  const parsedDate = new Date(Date.UTC(year, month - 1, day));
  if (
    Number.isNaN(parsedDate.getTime()) ||
    parsedDate.toISOString().split("T")[0] !== normalized
  ) {
    return null;
  }

  return normalized;
}

function normalizeMuzakkiPayload(muzakki) {
  if (!muzakki) return [];
  if (Array.isArray(muzakki)) return muzakki;
  if (typeof muzakki === "object") return Object.values(muzakki);
  return [];
}

function formatTanggalLabel(value) {
  if (!value) return null;

  return new Date(`${value}T00:00:00`).toLocaleDateString("id-ID", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
  });
}

// GET /muzakki - List all muzakki grouped by tanggal
router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;

    // Get data grouped by tanggal pembayaran zakat
    const [tanggalData] = await db.execute(`
            SELECT
                DATE_FORMAT(grouped.tanggal, '%Y-%m-%d') as tanggal,
                CAST(
                    SUBSTRING_INDEX(
                        GROUP_CONCAT(DISTINCT grouped.rt_id ORDER BY grouped.nomor_rt SEPARATOR ','),
                        ',',
                        1
                    ) AS UNSIGNED
                ) as primary_rt_id,
                COUNT(*) as total_muzakki,
                COUNT(DISTINCT grouped.rt_id) as total_rt,
                GROUP_CONCAT(
                    DISTINCT CONCAT('RT ', grouped.nomor_rt)
                    ORDER BY grouped.nomor_rt
                    SEPARATOR ', '
                ) as daftar_rt,
                SUM(grouped.total_nama_muzakki) as total_nama_muzakki,
                SUM(grouped.jumlah_jiwa) as total_jiwa,
                SUM(CASE WHEN grouped.jenis_zakat = 'beras' THEN grouped.jumlah_beras_kg ELSE 0 END) as total_beras,
                SUM(CASE WHEN grouped.jenis_zakat = 'uang' THEN grouped.jumlah_bayar ELSE 0 END) as total_uang,
                SUM(
                    CASE
                        WHEN grouped.kembalian > 0 AND grouped.kembalian IS NOT NULL THEN grouped.kembalian
                        ELSE 0
                    END
                ) as total_kembalian_saat_ini,
                SUM(grouped.total_infak) as total_infak
            FROM (
                SELECT
                    m.id,
                    ${getMuzakkiTanggalSelectSql("m")} as tanggal,
                    m.rt_id,
                    r.nomor_rt,
                    m.jumlah_jiwa,
                    m.jenis_zakat,
                    m.jumlah_beras_kg,
                    m.jumlah_bayar,
                    m.kembalian,
                    (
                        SELECT COUNT(*)
                        FROM muzakki_details md
                        WHERE md.muzakki_id = m.id
                    ) as total_nama_muzakki,
                    COALESCE((
                        SELECT SUM(i.jumlah)
                        FROM infak i
                        WHERE i.muzakki_id = m.id
                    ), 0) as total_infak
                FROM muzakki m
                LEFT JOIN rt r ON m.rt_id = r.id
            ) grouped
            GROUP BY grouped.tanggal
            ORDER BY grouped.tanggal DESC
        `);

    // Get overall statistics
    const [stats] = await db.execute(`
            SELECT 
                COUNT(DISTINCT m.id) as total_muzakki_records,
                COUNT(DISTINCT m.rt_id) as total_rt_all,
                SUM(m.jumlah_jiwa) as total_jiwa_all,
                SUM(CASE WHEN m.jenis_zakat = 'beras' THEN m.jumlah_beras_kg ELSE 0 END) as total_beras_all,
                SUM(CASE WHEN m.jenis_zakat = 'uang' THEN m.jumlah_bayar ELSE 0 END) as total_uang_all,
                SUM(m.kembalian) as total_kembalian_all,
                COUNT(DISTINCT md.id) as total_nama_muzakki_all
            FROM muzakki m
            LEFT JOIN muzakki_details md ON m.id = md.muzakki_id
        `);

    res.render("muzakki/index", {
      title: "Data Muzakki - Zakat Fitrah",
      layout: "layouts/main",
      tanggalData,
      stats: stats[0] || {},
    });
  } catch (error) {
    console.error("Error fetching muzakki:", error);
    req.flash("error_msg", "Terjadi kesalahan saat mengambil data muzakki");
    res.redirect("/");
  }
});

// GET /muzakki/tanggal/:tanggal - Show detail muzakki for specific tanggal
router.get("/tanggal/:tanggal", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const selectedTanggal = normalizeMuzakkiDateInput(req.params.tanggal);

    if (!selectedTanggal) {
      req.flash("error_msg", "Format tanggal tidak valid");
      return res.redirect("/muzakki");
    }

    const [muzakki] = await db.execute(
      `
      SELECT
        m.*,
        ${getMuzakkiTanggalSelectSql("m")} as tanggal,
        r.nomor_rt,
        r.ketua_rt,
        u.name as pencatat_name,
        GROUP_CONCAT(md.nama_muzakki SEPARATOR ', ') as nama_muzakki_list,
        COALESCE(COUNT(md.id), 0) as jumlah_muzakki,
        MAX(i.id) as infak_id,
        COALESCE(SUM(i.jumlah), 0) as infak_jumlah
      FROM muzakki m
      LEFT JOIN rt r ON m.rt_id = r.id
      LEFT JOIN users u ON m.user_id = u.id
      LEFT JOIN muzakki_details md ON m.id = md.muzakki_id
      LEFT JOIN infak i ON m.id = i.muzakki_id
      WHERE ${getMuzakkiTanggalSql("m")} = ?
      GROUP BY m.id, r.nomor_rt, r.ketua_rt, u.name
      ORDER BY r.nomor_rt ASC, m.created_at DESC
      `,
      [selectedTanggal]
    );

    if (muzakki.length === 0) {
      req.flash("error_msg", "Tidak ada data muzakki pada tanggal tersebut");
      return res.redirect("/muzakki");
    }

    const summary = muzakki.reduce(
      (accumulator, item) => {
        accumulator.totalMuzakki += 1;
        accumulator.totalNama += parseInt(item.jumlah_muzakki, 10) || 0;
        accumulator.totalJiwa += parseInt(item.jumlah_jiwa, 10) || 0;
        accumulator.totalRT.add(item.rt_id);
        return accumulator;
      },
      {
        totalMuzakki: 0,
        totalNama: 0,
        totalJiwa: 0,
        totalRT: new Set(),
      }
    );

    res.render("muzakki/date-detail", {
      title: `Detail Muzakki Tanggal ${selectedTanggal} - Zakat Fitrah`,
      layout: "layouts/main",
      selectedTanggal,
      selectedTanggalLabel: formatTanggalLabel(selectedTanggal),
      muzakki,
      summary: {
        totalMuzakki: summary.totalMuzakki,
        totalNama: summary.totalNama,
        totalJiwa: summary.totalJiwa,
        totalRT: summary.totalRT.size,
      },
    });
  } catch (error) {
    console.error("Error fetching tanggal detail:", error);
    req.flash("error_msg", "Terjadi kesalahan saat mengambil detail per tanggal");
    res.redirect("/muzakki");
  }
});

// GET /muzakki/rt/:rt_id - Show detail muzakki for specific RT
router.get("/rt/:rt_id", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const rtId = req.params.rt_id;
    const selectedTanggal = normalizeMuzakkiDateInput(req.query.tanggal);

    // Get RT info
    const [rtInfo] = await db.execute("SELECT * FROM rt WHERE id = ?", [rtId]);

    if (rtInfo.length === 0) {
      req.flash("error_msg", "Data RT tidak ditemukan");
      return res.redirect("/muzakki");
    }

    let navTabs = [];

    if (selectedTanggal) {
      const [tanggalTabs] = await db.execute(
        `
        SELECT
          r.id,
          r.nomor_rt,
          r.ketua_rt,
          COUNT(DISTINCT m.id) as total_muzakki
        FROM muzakki m
        INNER JOIN rt r ON m.rt_id = r.id
        WHERE ${getMuzakkiTanggalSql("m")} = ?
        GROUP BY r.id, r.nomor_rt, r.ketua_rt
        ORDER BY r.nomor_rt ASC
        `,
        [selectedTanggal]
      );

      navTabs = tanggalTabs || [];

      if (navTabs.length === 0) {
        req.flash("error_msg", "Tidak ada data RT pada tanggal tersebut");
        return res.redirect("/muzakki");
      }

      const isRtIncluded = navTabs.some((item) => String(item.id) === String(rtId));
      if (!isRtIncluded) {
        return res.redirect(`/muzakki/rt/${navTabs[0].id}?tanggal=${selectedTanggal}`);
      }
    }

    // Get all muzakki for this RT
    const muzakkiParams = [rtId];
    let muzakkiWhereClause = "WHERE m.rt_id = ?";

    if (selectedTanggal) {
      muzakkiWhereClause += ` AND ${getMuzakkiTanggalSql("m")} = ?`;
      muzakkiParams.push(selectedTanggal);
    }

    const [muzakki] = await db.execute(
      `
      SELECT 
        m.*,
        ${getMuzakkiTanggalSelectSql("m")} as tanggal,
        u.name as pencatat_name,
        GROUP_CONCAT(md.nama_muzakki SEPARATOR ', ') as nama_muzakki_list,
        COALESCE(COUNT(md.id), 0) as jumlah_muzakki,
        MAX(i.id) as infak_id,
        SUM(i.jumlah) as infak_jumlah
      FROM muzakki m
      LEFT JOIN users u ON m.user_id = u.id
      LEFT JOIN muzakki_details md ON m.id = md.muzakki_id
      LEFT JOIN infak i ON m.id = i.muzakki_id
      ${muzakkiWhereClause}
      GROUP BY m.id, u.name
      ORDER BY ${getMuzakkiTanggalSql("m")} DESC, m.created_at DESC
    `,
      muzakkiParams
    );

    // Debug: Log data untuk troubleshooting
    console.log('=== RT DETAIL DEBUG ===');
    console.log('RT ID:', rtId);
    console.log('Total muzakki:', muzakki.length);
    muzakki.forEach((m, idx) => {
      console.log(`Muzakki ${idx + 1}:`, {
        id: m.id,
        nama: m.nama_muzakki_list,
        kembalian: m.kembalian,
        infak_id: m.infak_id,
        jumlah_bayar: m.jumlah_bayar
      });
    });

    res.render("muzakki/rt-detail", {
      title: `Detail Muzakki RT ${rtInfo[0].nomor_rt} - Zakat Fitrah`,
      layout: "layouts/main",
      rt: rtInfo[0],
      muzakki,
      navTabs,
      selectedTanggal,
      selectedTanggalLabel: formatTanggalLabel(selectedTanggal),
    });
  } catch (error) {
    console.error("Error fetching RT detail:", error);
    req.flash("error_msg", "Terjadi kesalahan saat mengambil detail RT");
    res.redirect("/muzakki");
  }
});

// GET /muzakki/rt/:rt_id/export-excel - Export data muzakki khusus RT tertentu
router.get("/rt/:rt_id/export-excel", async (req, res) => {
  let ExcelJS;

  try {
    ExcelJS = require("exceljs");
  } catch (error) {
    console.error("ExcelJS not installed:", error);
    return res.status(500).json({
      success: false,
      message:
        "ExcelJS library tidak tersedia. Silakan install dengan: npm install exceljs",
    });
  }

  try {
    const db = req.app.locals.db;
    const rtId = req.params.rt_id;
    const selectedTanggal = normalizeMuzakkiDateInput(req.query.tanggal);

    const [rtRows] = await db.execute(
      `
      SELECT id, nomor_rt, ketua_rt
      FROM rt
      WHERE id = ?
      LIMIT 1
      `,
      [rtId]
    );

    if (rtRows.length === 0) {
      return res.status(404).json({
        success: false,
        message: "Data RT tidak ditemukan",
      });
    }

    const rt = rtRows[0];
    const exportParams = [rtId];
    let exportWhereClause = "WHERE m.rt_id = ?";

    if (selectedTanggal) {
      exportWhereClause += ` AND ${getMuzakkiTanggalSql("m")} = ?`;
      exportParams.push(selectedTanggal);
    }

    const [muzakkiData] = await db.execute(
      `
      SELECT
        m.id,
        m.jumlah_jiwa,
        m.jenis_zakat,
        m.jumlah_beras_kg,
        m.jumlah_uang,
        m.jumlah_bayar,
        m.kembalian,
        ${getMuzakkiTanggalSelectSql("m")} as tanggal,
        u.name as pencatat_name,
        (
          SELECT GROUP_CONCAT(nama_muzakki SEPARATOR ', ')
          FROM muzakki_details
          WHERE muzakki_id = m.id
        ) as nama_muzakki_list
      FROM muzakki m
      LEFT JOIN users u ON m.user_id = u.id
      ${exportWhereClause}
      ORDER BY ${getMuzakkiTanggalSql("m")} DESC, m.created_at DESC
      `,
      exportParams
    );

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Zakat Fitrah App";
    workbook.lastModifiedBy = "System";
    workbook.created = new Date();
    workbook.modified = new Date();

    const worksheet = workbook.addWorksheet(
      normalizeExcelSheetName(`RT ${rt.nomor_rt}`)
    );

    worksheet.mergeCells("A1:J1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = `DATA MUZAKKI RT ${sanitizeForExcel(rt.nomor_rt)}`;
    titleCell.font = { bold: true, size: 14, color: { argb: "FF1EAF2F" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE8F5E9" },
    };
    worksheet.getRow(1).height = 25;

    worksheet.mergeCells("A2:J2");
    const ketuaCell = worksheet.getCell("A2");
    ketuaCell.value = `Ketua RT: ${sanitizeForExcel(rt.ketua_rt) || "-"}`;
    ketuaCell.font = { italic: true, size: 11 };
    ketuaCell.alignment = { horizontal: "center", vertical: "middle" };
    ketuaCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF5F5F5" },
    };
    worksheet.getRow(2).height = 20;

    worksheet.addRow([]);

    const headerRow = worksheet.addRow([
      "No",
      "Nama Muzakki",
      "Jumlah Jiwa",
      "Jenis Zakat",
      "Jumlah Beras (kg)",
      "Jumlah Uang (Rp)",
      "Jumlah Bayar (Rp)",
      "Kembalian (Rp)",
      "Pencatat",
      "Tanggal",
    ]);

    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.alignment = { horizontal: "center", vertical: "middle" };
    headerRow.height = 20;

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF1EAF2F" },
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    worksheet.getColumn(1).width = 5;
    worksheet.getColumn(2).width = 35;
    worksheet.getColumn(3).width = 12;
    worksheet.getColumn(4).width = 12;
    worksheet.getColumn(5).width = 18;
    worksheet.getColumn(6).width = 18;
    worksheet.getColumn(7).width = 18;
    worksheet.getColumn(8).width = 18;
    worksheet.getColumn(9).width = 20;
    worksheet.getColumn(10).width = 15;
    worksheet.getColumn(3).numFmt = EXCEL_NUMFMT.INTEGER;
    worksheet.getColumn(5).numFmt = EXCEL_NUMFMT.WEIGHT;
    worksheet.getColumn(6).numFmt = EXCEL_NUMFMT.CURRENCY_RP;
    worksheet.getColumn(7).numFmt = EXCEL_NUMFMT.CURRENCY_RP;
    worksheet.getColumn(8).numFmt = EXCEL_NUMFMT.CURRENCY_RP;

    let totalJiwa = 0;
    let totalBeras = 0;
    let totalUang = 0;
    let totalBayar = 0;
    let totalKembalian = 0;

    muzakkiData.forEach((item, index) => {
      const jiwa = parseInt(item.jumlah_jiwa, 10) || 0;
      const beras = parseFloat(item.jumlah_beras_kg) || 0;
      const uang = parseFloat(item.jumlah_uang) || 0;
      const bayar = parseFloat(item.jumlah_bayar) || 0;
      const kembalian = parseFloat(item.kembalian) || 0;

      totalJiwa += jiwa;
      totalBeras += beras;
      totalUang += uang;
      totalBayar += bayar;
      totalKembalian += kembalian;

      const row = worksheet.addRow([
        index + 1,
        sanitizeForExcel(item.nama_muzakki_list || "-"),
        jiwa,
        item.jenis_zakat === "beras" ? "Beras" : "Uang",
        beras,
        uang,
        bayar,
        kembalian,
        sanitizeForExcel(item.pencatat_name || "-"),
        formatDateForExcel(item.tanggal),
      ]);

      row.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE0E0E0" } },
          left: { style: "thin", color: { argb: "FFE0E0E0" } },
          bottom: { style: "thin", color: { argb: "FFE0E0E0" } },
          right: { style: "thin", color: { argb: "FFE0E0E0" } },
        };

        if ([1, 3, 4, 5, 6, 7, 8, 10].includes(colNumber)) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else {
          cell.alignment = { vertical: "middle" };
        }

        if (index % 2 === 0) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF9FAFB" },
          };
        }
      });

      applyMuzakkiExcelNumberFormat(row);
    });

    if (muzakkiData.length > 0) {
      worksheet.addRow([]);
      const totalRow = worksheet.addRow([
        "",
        "TOTAL",
        totalJiwa,
        "",
        Math.round(totalBeras * 100) / 100,
        Math.round(totalUang),
        Math.round(totalBayar),
        Math.round(totalKembalian),
        "",
        "",
      ]);

      totalRow.font = { bold: true, color: { argb: "FF1EAF2F" } };
      totalRow.height = 22;

      totalRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFE8F5E9" },
        };
        cell.border = {
          top: { style: "double", color: { argb: "FF1EAF2F" } },
          left: { style: "thin" },
          bottom: { style: "double", color: { argb: "FF1EAF2F" } },
          right: { style: "thin" },
        };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });

      applyMuzakkiExcelNumberFormat(totalRow);
    } else {
      const noDataRow = worksheet.addRow([
        "",
        "Belum ada data muzakki untuk RT ini",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
      ]);
      worksheet.mergeCells(`B${noDataRow.number}:J${noDataRow.number}`);
      const noDataCell = worksheet.getCell(`B${noDataRow.number}`);
      noDataCell.font = { italic: true, color: { argb: "FF9CA3AF" } };
      noDataCell.alignment = { horizontal: "center", vertical: "middle" };
    }

    const rawBuffer = await workbook.xlsx.writeBuffer();
    const nodeBuffer = Buffer.from(rawBuffer);

    if (!nodeBuffer || nodeBuffer.length === 0) {
      throw new Error("Buffer generation failed: empty buffer");
    }

    const dateString = new Date().toISOString().split("T")[0];
    const safeRt = safeRtForFileName(rt.nomor_rt);
    const filename = `Data_Muzakki_RT_${safeRt}_${dateString}.xlsx`;

    res.status(200);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Length", nodeBuffer.length);
    res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
    res.setHeader("X-Content-Type-Options", "nosniff");

    return res.end(nodeBuffer);
  } catch (error) {
    console.error("Error exporting RT to Excel:", error);

    if (!res.headersSent) {
      return res.status(500).json({
        success: false,
        message: "Terjadi kesalahan saat export data RT ke Excel",
        error: error.message,
      });
    }
  }
});

// GET /muzakki/create - Show create form
router.get("/create", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const requestedTanggal =
      normalizeMuzakkiDateInput(req.query.tanggal) ||
      new Date(Date.now() - new Date().getTimezoneOffset() * 60000)
        .toISOString()
        .split("T")[0];

    // Get all RT for dropdown
    const [rtList] = await db.execute("SELECT * FROM rt ORDER BY nomor_rt");
    
    // Get all master_zakat for dropdown
    const [masterZakatList] = await db.execute(
      "SELECT id, nama, harga, kg FROM master_zakat ORDER BY nama"
    );

    res.render("muzakki/create", {
      title: "Tambah Muzakki - Zakat Fitrah",
      layout: "layouts/main",
      rtList,
      masterZakatList,
      defaultTanggal: requestedTanggal,
    });
  } catch (error) {
    console.error("Error loading create form:", error);
    req.flash("error_msg", "Terjadi kesalahan saat memuat form");
    res.redirect("/muzakki");
  }
});

// POST /muzakki - Create new muzakki
router.post("/", async (req, res) => {
  const {
    rt_id,
    tanggal,
    muzakki,
    jumlah_jiwa,
    master_zakat_id,
    jumlah_bayar,
    catatan,
  } = req.body;

  try {
    // Debug log untuk melihat data yang diterima
    console.log("Received data:", {
      rt_id,
      tanggal,
      muzakki,
      jumlah_jiwa,
      master_zakat_id,
      jumlah_bayar,
      catatan,
    });

    // Input validation
    if (!rt_id || !tanggal || !jumlah_jiwa || !master_zakat_id || !jumlah_bayar) {
      req.flash("error_msg", "Semua field yang wajib harus diisi");
      return res.redirect("/muzakki/create");
    }

    const tanggalZakat = normalizeMuzakkiDateInput(tanggal);
    if (!tanggalZakat) {
      req.flash("error_msg", "Tanggal zakat tidak valid");
      return res.redirect("/muzakki/create");
    }

    const muzakkiArray = normalizeMuzakkiPayload(muzakki);

    console.log("Processed muzakkiArray:", muzakkiArray);

    if (!muzakkiArray || muzakkiArray.length === 0) {
      req.flash("error_msg", "Minimal harus ada satu data muzakki");
      return res.redirect("/muzakki/create");
    }

    // Validate each muzakki has nama
    for (let i = 0; i < muzakkiArray.length; i++) {
      if (
        !muzakkiArray[i] ||
        !muzakkiArray[i].nama ||
        !muzakkiArray[i].nama.trim()
      ) {
        req.flash("error_msg", `Nama muzakki #${i + 1} tidak boleh kosong`);
        return res.redirect("/muzakki/create");
      }
    }

    const jiwa = parseInt(jumlah_jiwa);
    const bayar = parseFloat(jumlah_bayar);

    if (jiwa <= 0 || bayar <= 0) {
      req.flash("error_msg", "Jumlah jiwa dan jumlah bayar harus lebih dari 0");
      return res.redirect("/muzakki/create");
    }

    const db = req.app.locals.db;
    const userId = req.session.user ? req.session.user.id : null;

    if (!userId) {
      req.flash("error_msg", "Session tidak valid. Silakan login kembali.");
      return res.redirect("/auth/login");
    }

    const connection = await db.getConnection();

    try {
      const [masterZakat] = await connection.execute(
        "SELECT * FROM master_zakat WHERE id = ?",
        [master_zakat_id]
      );

      if (masterZakat.length === 0) {
        req.flash("error_msg", "Jenis zakat tidak valid");
        return res.redirect("/muzakki/create");
      }

      const zakatData = masterZakat[0];
      const jenis_zakat = zakatData.kg > 0 ? "beras" : "uang";

      let jumlah_beras_kg = null;
      let jumlah_uang = null;
      let kewajiban = 0;

      if (jenis_zakat === "beras") {
        jumlah_beras_kg = jiwa * parseFloat(zakatData.kg);
        kewajiban = jiwa * parseFloat(zakatData.harga);
      } else if (jenis_zakat === "uang") {
        jumlah_uang = jiwa * parseFloat(zakatData.harga);
        kewajiban = jumlah_uang;
      }

      const kembalian = Math.max(0, bayar - kewajiban);

      console.log("About to start transaction with data:", {
        rt_id,
        tanggalZakat,
        jiwa,
        jenis_zakat,
        jumlah_beras_kg,
        jumlah_uang,
        bayar,
        kembalian,
        catatan,
        userId,
        muzakkiCount: muzakkiArray.length,
      });

      await connection.beginTransaction();

      // Insert main muzakki record (without nama)
      const [result] = await connection.execute(
        `
        INSERT INTO muzakki 
        (rt_id, jumlah_jiwa, jenis_zakat, jumlah_beras_kg, jumlah_uang, jumlah_bayar, kembalian, catatan, user_id, tanggal, master_zakat_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `,
        [
          rt_id,
          jiwa,
          jenis_zakat,
          jumlah_beras_kg,
          jumlah_uang,
          bayar,
          kembalian,
          catatan || null,
          userId,
          tanggalZakat,
          master_zakat_id,
        ]
      );

      const muzakkiId = result.insertId;
      console.log("Inserted muzakki with ID:", muzakkiId);

      // Insert muzakki details for each individual
      for (let i = 0; i < muzakkiArray.length; i++) {
        const muzakkiData = muzakkiArray[i];
        console.log(`Inserting muzakki detail ${i + 1}:`, muzakkiData);

        await connection.execute(
          `
          INSERT INTO muzakki_details 
          (muzakki_id, nama_muzakki, bin_binti, nama_orang_tua)
          VALUES (?, ?, ?, ?)
          `,
          [
            muzakkiId,
            muzakkiData.nama.trim(),
            muzakkiData.bin_binti || null,
            muzakkiData.nama_orang_tua || null,
          ]
        );
      }

      // Commit transaction
      await connection.commit();
      console.log("Transaction committed successfully");

      req.flash(
        "success_msg",
        `Data muzakki berhasil ditambahkan dengan ${muzakkiArray.length} orang muzakki`
      );
      res.redirect(`/muzakki/${muzakkiId}`);
    } catch (error) {
      // Rollback on error
      await connection.rollback();
      console.error("Error in transaction:", error);
      throw error;
    } finally {
      connection.release();
    }
  } catch (error) {
    console.error("Error creating muzakki:", error);
    req.flash(
      "error_msg",
      `Terjadi kesalahan saat menyimpan data muzakki: ${error.message}`
    );
    res.redirect("/muzakki/create");
  }
});

// GET /muzakki/:id/edit - Show edit form
router.get("/:id/edit", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const id = req.params.id;

    // Get muzakki data
    const [muzakki] = await db.execute(
      `
            SELECT 
                m.*,
                ${getMuzakkiTanggalSelectSql("m")} as tanggal,
                r.nomor_rt,
                r.ketua_rt
            FROM muzakki m
            LEFT JOIN rt r ON m.rt_id = r.id
            WHERE m.id = ?
        `,
      [id]
    );

    if (muzakki.length === 0) {
      req.flash("error_msg", "Data muzakki tidak ditemukan");
      return res.redirect("/muzakki");
    }

    // Get muzakki details
    const [muzakkiDetails] = await db.execute(
      "SELECT * FROM muzakki_details WHERE muzakki_id = ? ORDER BY id",
      [id]
    );

    // Get all RT for dropdown
    const [rtList] = await db.execute("SELECT * FROM rt ORDER BY nomor_rt");
    
    // Get all master_zakat for dropdown
    const [masterZakatList] = await db.execute(
      "SELECT id, nama, harga, kg FROM master_zakat ORDER BY nama"
    );

    res.render("muzakki/edit", {
      title: "Edit Muzakki - Zakat Fitrah",
      layout: "layouts/main",
      muzakki: muzakki[0],
      muzakkiDetails: muzakkiDetails, // Pass all details
      rtList,
      masterZakatList,
    });
  } catch (error) {
    console.error("Error loading edit form:", error);
    req.flash("error_msg", "Terjadi kesalahan saat memuat form");
    res.redirect("/muzakki");
  }
});

// PUT /muzakki/:id - Update muzakki
router.put("/:id", async (req, res) => {
  const id = req.params.id;
  const {
    rt_id,
    tanggal,
    muzakki,
    jumlah_jiwa,
    master_zakat_id,
    jumlah_bayar,
    catatan,
  } = req.body;

  try {
    // Input validation
    if (!rt_id || !tanggal || !jumlah_jiwa || !master_zakat_id || !jumlah_bayar) {
      req.flash("error_msg", "Semua field yang wajib harus diisi");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    const tanggalZakat = normalizeMuzakkiDateInput(tanggal);
    if (!tanggalZakat) {
      req.flash("error_msg", "Tanggal zakat tidak valid");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    const muzakkiArray = normalizeMuzakkiPayload(muzakki);

    // Validate muzakki array
    if (!muzakkiArray || muzakkiArray.length === 0) {
      req.flash("error_msg", "Minimal harus ada satu data muzakki");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    // Validate each muzakki has nama
    for (let i = 0; i < muzakkiArray.length; i++) {
      if (!muzakkiArray[i].nama || !muzakkiArray[i].nama.trim()) {
        req.flash("error_msg", `Nama muzakki #${i + 1} tidak boleh kosong`);
        return res.redirect(`/muzakki/${id}/edit`);
      }
    }

    const jiwa = parseInt(jumlah_jiwa);
    const bayar = parseFloat(jumlah_bayar);

    if (jiwa <= 0 || bayar <= 0) {
      req.flash("error_msg", "Jumlah jiwa dan jumlah bayar harus lebih dari 0");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    const db = req.app.locals.db;
    const connection = await db.getConnection();

    try {
      const [masterZakat] = await connection.execute(
        "SELECT * FROM master_zakat WHERE id = ?",
        [master_zakat_id]
      );

      if (masterZakat.length === 0) {
        req.flash("error_msg", "Jenis zakat tidak valid");
        return res.redirect(`/muzakki/${id}/edit`);
      }

      const zakatData = masterZakat[0];
      const jenis_zakat = zakatData.kg > 0 ? "beras" : "uang";

      let jumlah_beras_kg = null;
      let jumlah_uang = null;
      let kewajiban = 0;

      if (jenis_zakat === "beras") {
        jumlah_beras_kg = jiwa * parseFloat(zakatData.kg);
        kewajiban = jiwa * parseFloat(zakatData.harga);
      } else if (jenis_zakat === "uang") {
        jumlah_uang = jiwa * parseFloat(zakatData.harga);
        kewajiban = jumlah_uang;
      }

      const kembalian = Math.max(0, bayar - kewajiban);

      await connection.beginTransaction();

      // Update muzakki main record
      await connection.execute(
        `
        UPDATE muzakki 
        SET rt_id = ?, jumlah_jiwa = ?, jenis_zakat = ?, 
            jumlah_beras_kg = ?, jumlah_uang = ?, jumlah_bayar = ?, 
            kembalian = ?, catatan = ?, tanggal = ?, master_zakat_id = ?
        WHERE id = ?
        `,
        [
          rt_id,
          jiwa,
          jenis_zakat,
          jumlah_beras_kg,
          jumlah_uang,
          bayar,
          kembalian,
          catatan,
          tanggalZakat,
          master_zakat_id,
          id,
        ]
      );

      // Delete existing muzakki details
      await connection.execute("DELETE FROM muzakki_details WHERE muzakki_id = ?", [
        id,
      ]);

      // Insert new muzakki details
      for (let i = 0; i < muzakkiArray.length; i++) {
        const muzakkiData = muzakkiArray[i];

        await connection.execute(
          `
          INSERT INTO muzakki_details 
          (muzakki_id, nama_muzakki, bin_binti, nama_orang_tua)
          VALUES (?, ?, ?, ?)
          `,
          [
            id,
            muzakkiData.nama.trim(),
            muzakkiData.bin_binti || null,
            muzakkiData.nama_orang_tua || null,
          ]
        );
      }

      // Update infak if kembalian changed
      const [existingInfak] = await connection.execute(
        "SELECT id FROM infak WHERE muzakki_id = ? AND keterangan = 'Kembalian zakat fitrah'",
        [id]
      );

      if (kembalian > 0) {
        if (existingInfak.length > 0) {
          // Update existing infak
          await connection.execute(
            "UPDATE infak SET jumlah = ? WHERE muzakki_id = ? AND keterangan = 'Kembalian zakat fitrah'",
            [kembalian, id]
          );
        } else {
          // Insert new infak
          await connection.execute(
            `
            INSERT INTO infak (muzakki_id, jumlah, keterangan)
            VALUES (?, ?, ?)
            `,
            [id, kembalian, "Kembalian zakat fitrah"]
          );
        }
      } else if (existingInfak.length > 0) {
        // Delete infak if no kembalian
        await connection.execute(
          "DELETE FROM infak WHERE muzakki_id = ? AND keterangan = 'Kembalian zakat fitrah'",
          [id]
        );
      }

      // Commit transaction
      await connection.commit();

      req.flash(
        "success_msg",
        `Data muzakki berhasil diupdate dengan ${muzakkiArray.length} orang muzakki`
      );
      res.redirect(`/muzakki/${id}`);
    } catch (error) {
      // Rollback transaction on error
      await connection.rollback();
      throw error;
    } finally {
      connection.release();
    }
  } catch (error) {
    console.error("Error updating muzakki:", error);
    req.flash("error_msg", "Terjadi kesalahan saat mengupdate data muzakki");
    res.redirect(`/muzakki/${id}/edit`);
  }
});

// DELETE /muzakki/:id - Delete muzakki
router.delete("/:id", async (req, res) => {
  const id = req.params.id;

  try {
    const db = req.app.locals.db;

    // Check if muzakki exists and get RT ID for redirect
    const [muzakki] = await db.execute("SELECT rt_id FROM muzakki WHERE id = ?", [
      id,
    ]);

    if (muzakki.length === 0) {
      req.flash("error_msg", "Data muzakki tidak ditemukan");
      return res.redirect("/muzakki");
    }

    const rtId = muzakki[0].rt_id;

    // Delete muzakki details first
    await db.execute("DELETE FROM muzakki_details WHERE muzakki_id = ?", [id]);

    // Delete related infak records
    await db.execute("DELETE FROM infak WHERE muzakki_id = ?", [id]);

    // Delete muzakki
    await db.execute("DELETE FROM muzakki WHERE id = ?", [id]);

    req.flash("success_msg", "Data muzakki berhasil dihapus");

    // Redirect back to RT detail if came from RT page, otherwise to main muzakki page
    const referer = req.get('Referer') || '';
    if (referer.includes('/muzakki/rt/') && rtId) {
      res.redirect(`/muzakki/rt/${rtId}`);
    } else {
      res.redirect("/muzakki");
    }
  } catch (error) {
    console.error("Error deleting muzakki:", error);
    req.flash("error_msg", "Terjadi kesalahan saat menghapus data");

    // Try to redirect back to referer or muzakki index
    const referer = req.get('Referer') || '/muzakki';
    res.redirect(referer.replace(/^https?:\/\/[^\/]+/, ''));
  }
});

// POST /muzakki/:id/sedekahkan-kembalian - Sedekahkan kembalian sebagai infak
router.post("/:id/sedekahkan-kembalian", async (req, res) => {
  const id = req.params.id;

  try {
    const db = req.app.locals.db;

    // Get muzakki data
    const [muzakki] = await db.execute("SELECT * FROM muzakki WHERE id = ?", [
      id,
    ]);

    if (muzakki.length === 0) {
      req.flash("error_msg", "Data muzakki tidak ditemukan");
      return res.redirect("/muzakki");
    }

    const muzakkiData = muzakki[0];

    if (muzakkiData.kembalian <= 0) {
      req.flash("error_msg", "Tidak ada kembalian untuk disedekahkan");
      return res.redirect("/muzakki");
    }

    // Insert infak
    await db.execute(
      `
            INSERT INTO infak (muzakki_id, jumlah, keterangan)
            VALUES (?, ?, ?)
        `,
      [id, muzakkiData.kembalian, "Kembalian dari zakat fitrah"]
    );

    // Reset kembalian to 0
    await db.execute("UPDATE muzakki SET kembalian = 0 WHERE id = ?", [id]);

    req.flash(
      "success_msg",
      `Kembalian Rp ${muzakkiData.kembalian.toLocaleString(
        "id-ID"
      )} berhasil disedekahkan sebagai infak`
    );
    res.redirect("/muzakki");
  } catch (error) {
    console.error("Error sedekahkan kembalian:", error);
    req.flash("error_msg", "Terjadi kesalahan saat menyedekahkan kembalian");
    res.redirect("/muzakki");
  }
});

// POST /muzakki/rt/:rtId/batch-infak - Batch insert infak dari kembalian
router.post("/rt/:rtId/batch-infak", async (req, res) => {
  const rtId = req.params.rtId;
  const { muzakki_ids } = req.body;

  try {
    const db = req.app.locals.db;

    // Validate input
    if (!muzakki_ids || muzakki_ids.length === 0) {
      req.flash("error_msg", "Pilih minimal 1 muzakki untuk ditambahkan sebagai infak");
      return res.redirect(`/muzakki/rt/${rtId}`);
    }

    // Ensure muzakki_ids is array
    const ids = Array.isArray(muzakki_ids) ? muzakki_ids : [muzakki_ids];
    
    let successCount = 0;
    let totalInfak = 0;
    const errors = [];

    // Process each muzakki
    for (const muzakkiId of ids) {
      try {
        // Get muzakki data
        const [muzakki] = await db.execute(
          "SELECT * FROM muzakki WHERE id = ? AND rt_id = ?",
          [muzakkiId, rtId]
        );

        if (muzakki.length === 0) {
          errors.push(`Muzakki ID ${muzakkiId} tidak ditemukan`);
          continue;
        }

        const muzakkiData = muzakki[0];

        // Check if already has infak
        const [existingInfak] = await db.execute(
          "SELECT id FROM infak WHERE muzakki_id = ?",
          [muzakkiId]
        );

        if (existingInfak.length > 0) {
          errors.push(`Muzakki ${muzakkiData.nama_muzakki_list || muzakkiId} sudah memiliki infak`);
          continue;
        }

        if (muzakkiData.kembalian <= 0) {
          errors.push(`Muzakki ${muzakkiData.nama_muzakki_list || muzakkiId} tidak memiliki kembalian`);
          continue;
        }

        // Get keterangan for this muzakki
        const keterangan = req.body[`keterangan_${muzakkiId}`] || "Kembalian dari zakat fitrah";

        // Insert infak
        await db.execute(
          `INSERT INTO infak (muzakki_id, jumlah, keterangan) VALUES (?, ?, ?)`,
          [muzakkiId, muzakkiData.kembalian, keterangan]
        );

        // Reset kembalian to 0
        await db.execute("UPDATE muzakki SET kembalian = 0 WHERE id = ?", [muzakkiId]);

        successCount++;
        totalInfak += muzakkiData.kembalian;
      } catch (err) {
        console.error(`Error processing muzakki ${muzakkiId}:`, err);
        errors.push(`Error pada muzakki ID ${muzakkiId}`);
      }
    }

    // Show results
    if (successCount > 0) {
      req.flash(
        "success_msg",
        `Berhasil menambahkan ${successCount} infak dengan total Rp ${totalInfak.toLocaleString("id-ID")}`
      );
    }

    if (errors.length > 0) {
      req.flash("error_msg", errors.join(", "));
    }

    res.redirect(`/muzakki/rt/${rtId}`);
  } catch (error) {
    console.error("Error batch infak:", error);
    req.flash("error_msg", "Terjadi kesalahan saat menambahkan infak");
    res.redirect(`/muzakki/rt/${rtId}`);
  }
});

// GET /muzakki/export-excel - Export all muzakki data to Excel with sheets per RT
// IMPORTANT: This route MUST be defined before /:id to prevent "export-excel" matching as an :id param
router.get("/export-excel", async (req, res) => {
  let ExcelJS;

  try {
    ExcelJS = require('exceljs');
  } catch (error) {
    console.error("ExcelJS not installed:", error);
    return res.status(500).json({
      success: false,
      message: "ExcelJS library tidak tersedia. Silakan install dengan: npm install exceljs"
    });
  }

  try {
    const db = req.app.locals.db;

    const sanitizeForExcel = (str) => {
      if (str === null || str === undefined) return '';
      let s = String(str);
      s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
      if (s.length > 30000) {
        s = s.substring(0, 30000) + '... (truncated)';
      }
      return s;
    };

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Zakat Fitrah App';
    workbook.lastModifiedBy = 'System';
    workbook.created = new Date();
    workbook.modified = new Date();

    const [rtList] = await db.execute(`
      SELECT id, nomor_rt, ketua_rt
      FROM rt
      ORDER BY nomor_rt
    `);

    if (rtList.length === 0) {
      const worksheet = workbook.addWorksheet('Data Kosong');
      worksheet.addRow(['Tidak ada data RT tersedia']);
      worksheet.getCell('A1').font = { bold: true, color: { argb: 'FFFF0000' } };
    } else {

    for (const rt of rtList) {
      const [muzakkiData] = await db.execute(`
        SELECT
          m.id,
          m.jumlah_jiwa,
          m.jenis_zakat,
          m.jumlah_beras_kg,
          m.jumlah_uang,
          m.jumlah_bayar,
          m.kembalian,
          ${getMuzakkiTanggalSelectSql("m")} as tanggal,
          u.name as pencatat_name,
          (
            SELECT GROUP_CONCAT(nama_muzakki SEPARATOR ', ')
            FROM muzakki_details
            WHERE muzakki_id = m.id
          ) as nama_muzakki_list
        FROM muzakki m
        LEFT JOIN users u ON m.user_id = u.id
        WHERE m.rt_id = ?
        ORDER BY ${getMuzakkiTanggalSql("m")} DESC, m.created_at DESC
      `, [rt.id]);

      let sheetName = `RT ${rt.nomor_rt}`.trim();
      sheetName = sheetName.replace(/[\/\\\?\*\[\]\:]/g, '_');
      if (sheetName.length > 31) {
        sheetName = sheetName.substring(0, 31);
      }

      let uniqueName = sheetName;
      let counter = 1;
      while (workbook.getWorksheet(uniqueName)) {
        uniqueName = `${sheetName.substring(0, 28)}(${counter})`;
        counter++;
      }

      const worksheet = workbook.addWorksheet(uniqueName);

      worksheet.mergeCells('A1:J1');
      const titleCell = worksheet.getCell('A1');
      titleCell.value = `DATA MUZAKKI RT ${sanitizeForExcel(rt.nomor_rt)}`;
      titleCell.font = { bold: true, size: 14, color: { argb: 'FF1EAF2F' } };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE8F5E9' }
      };
      worksheet.getRow(1).height = 25;

      worksheet.mergeCells('A2:J2');
      const ketuaCell = worksheet.getCell('A2');
      ketuaCell.value = `Ketua RT: ${sanitizeForExcel(rt.ketua_rt) || '-'}`;
      ketuaCell.font = { italic: true, size: 11 };
      ketuaCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ketuaCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF5F5F5' }
      };
      worksheet.getRow(2).height = 20;

      worksheet.addRow([]);

      const headerRow = worksheet.addRow([
        'No',
        'Nama Muzakki',
        'Jumlah Jiwa',
        'Jenis Zakat',
        'Jumlah Beras (kg)',
        'Jumlah Uang (Rp)',
        'Jumlah Bayar (Rp)',
        'Kembalian (Rp)',
        'Pencatat',
        'Tanggal'
      ]);

      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
      headerRow.height = 20;

      headerRow.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF1EAF2F' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      worksheet.getColumn(1).width = 5;
      worksheet.getColumn(2).width = 35;
      worksheet.getColumn(3).width = 12;
      worksheet.getColumn(4).width = 12;
      worksheet.getColumn(5).width = 18;
      worksheet.getColumn(6).width = 18;
      worksheet.getColumn(7).width = 18;
      worksheet.getColumn(8).width = 18;
      worksheet.getColumn(9).width = 20;
      worksheet.getColumn(10).width = 15;
      worksheet.getColumn(3).numFmt = EXCEL_NUMFMT.INTEGER;
      worksheet.getColumn(5).numFmt = EXCEL_NUMFMT.WEIGHT;
      worksheet.getColumn(6).numFmt = EXCEL_NUMFMT.CURRENCY_RP;
      worksheet.getColumn(7).numFmt = EXCEL_NUMFMT.CURRENCY_RP;
      worksheet.getColumn(8).numFmt = EXCEL_NUMFMT.CURRENCY_RP;

      let totalJiwa = 0;
      let totalBeras = 0;
      let totalUang = 0;
      let totalBayar = 0;
      let totalKembalian = 0;

      muzakkiData.forEach((item, index) => {
        const jiwa = parseInt(item.jumlah_jiwa, 10) || 0;
        const beras = parseFloat(item.jumlah_beras_kg) || 0;
        const uang = parseFloat(item.jumlah_uang) || 0;
        const bayar = parseFloat(item.jumlah_bayar) || 0;
        const kembalian = parseFloat(item.kembalian) || 0;

        totalJiwa += jiwa;
        totalBeras += beras;
        totalUang += uang;
        totalBayar += bayar;
        totalKembalian += kembalian;

        const dataRow = worksheet.addRow([
          index + 1,
          sanitizeForExcel(item.nama_muzakki_list || '-'),
          jiwa,
          item.jenis_zakat === 'beras' ? 'Beras' : 'Uang',
          beras,
          uang,
          bayar,
          kembalian,
          sanitizeForExcel(item.pencatat_name || '-'),
          formatDateForExcel(item.tanggal)
        ]);

        dataRow.eachCell((cell, colNumber) => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
            left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
            bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
            right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
          };

          if ([1, 3, 4, 5, 6, 7, 8, 10].includes(colNumber)) {
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          } else {
            cell.alignment = { vertical: 'middle' };
          }

          if (index % 2 === 0) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFF9FAFB' }
            };
          }
        });

        applyMuzakkiExcelNumberFormat(dataRow);
      });

      if (muzakkiData.length > 0) {
        worksheet.addRow([]);
        const totalRow = worksheet.addRow([
          '',
          'TOTAL',
          totalJiwa,
          '',
          Math.round(totalBeras * 100) / 100,
          Math.round(totalUang),
          Math.round(totalBayar),
          Math.round(totalKembalian),
          '',
          ''
        ]);

        totalRow.font = { bold: true, color: { argb: 'FF1EAF2F' } };
        totalRow.height = 22;

        totalRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE8F5E9' }
          };
          cell.border = {
            top: { style: 'double', color: { argb: 'FF1EAF2F' } },
            left: { style: 'thin' },
            bottom: { style: 'double', color: { argb: 'FF1EAF2F' } },
            right: { style: 'thin' }
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });

        applyMuzakkiExcelNumberFormat(totalRow);
      } else {
        const noDataRow = worksheet.addRow(['', 'Belum ada data muzakki untuk RT ini', '', '', '', '', '', '', '', '']);
        worksheet.mergeCells(`B${noDataRow.number}:J${noDataRow.number}`);
        const noDataCell = worksheet.getCell(`B${noDataRow.number}`);
        noDataCell.font = { italic: true, color: { argb: 'FF9CA3AF' } };
        noDataCell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    }
    }

    console.log('📊 Generating Excel buffer...');

    const rawBuffer = await workbook.xlsx.writeBuffer();
    const nodeBuffer = Buffer.from(rawBuffer);

    if (!nodeBuffer || nodeBuffer.length === 0) {
      throw new Error('Buffer generation failed: empty buffer');
    }

    console.log(`✅ Excel generated successfully. Size: ${nodeBuffer.length} bytes (${(nodeBuffer.length / 1024).toFixed(2)} KB)`);

    if (res.headersSent) {
      console.error('❌ ABORT: Headers already sent, cannot send file');
      return;
    }

    const dateString = new Date().toISOString().split('T')[0];
    const filename = `Data_Muzakki_${dateString}.xlsx`;

    res.status(200);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', nodeBuffer.length);
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('X-Content-Type-Options', 'nosniff');

    res.end(nodeBuffer);

    console.log('✅ Excel file sent successfully');
    return;

  } catch (error) {
    console.error("❌ Error exporting to Excel:", error);
    console.error("Stack trace:", error.stack);

    if (!res.headersSent) {
      try {
        res.status(500).json({
          success: false,
          message: "Terjadi kesalahan saat export ke Excel",
          error: error.message
        });
      } catch (sendError) {
        console.error("❌ Failed to send error JSON:", sendError);
        res.status(500).send('Error generating Excel file');
      }
    }
  }
});

// GET /muzakki/:id - Show detail muzakki
router.get("/:id", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const id = req.params.id;

    // Get muzakki data with RT and user info
    const [muzakki] = await db.execute(
      `
            SELECT 
                m.*,
                ${getMuzakkiTanggalSelectSql("m")} as tanggal,
                r.nomor_rt,
                r.ketua_rt,
                u.name as pencatat_name
            FROM muzakki m
            LEFT JOIN rt r ON m.rt_id = r.id
            LEFT JOIN users u ON m.user_id = u.id
            WHERE m.id = ?
        `,
      [id]
    );

    if (muzakki.length === 0) {
      req.flash("error_msg", "Data muzakki tidak ditemukan");
      return res.redirect("/muzakki");
    }

    // Get muzakki details
    const [muzakkiDetails] = await db.execute(
      "SELECT * FROM muzakki_details WHERE muzakki_id = ? ORDER BY id",
      [id]
    );

    res.render("muzakki/detail", {
      title: "Detail Muzakki - Zakat Fitrah",
      layout: "layouts/main",
      muzakki: muzakki[0],
      muzakkiDetails: muzakkiDetails,
    });
  } catch (error) {
    console.error("Error fetching muzakki detail:", error);
    req.flash("error_msg", "Terjadi kesalahan saat mengambil detail muzakki");
    res.redirect("/muzakki");
  }
});

// ========================================
// API Routes for Master Zakat CRUD
// ========================================

// GET /api/master-zakat - Get all master zakat
router.get("/api/master-zakat", async (req, res) => {
  try {
    const db = req.app.locals.db;
    
    const [masterZakat] = await db.execute(
      `SELECT id, nama, harga, kg, created_at, updated_at 
       FROM master_zakat 
       ORDER BY id ASC`
    );

    res.json({
      success: true,
      data: masterZakat
    });
  } catch (error) {
    console.error("Error fetching master zakat:", error);
    res.status(500).json({
      success: false,
      message: "Terjadi kesalahan saat mengambil data master zakat"
    });
  }
});

// POST /api/master-zakat - Create new master zakat
router.post("/api/master-zakat", async (req, res) => {
  const { nama, harga, kg } = req.body;

  try {
    // Validation
    if (!nama || nama.trim() === '') {
      return res.status(400).json({
        success: false,
        message: "Nama jenis zakat harus diisi"
      });
    }

    if (harga === undefined || harga === null || harga === '' || parseFloat(harga) < 0) {
      return res.status(400).json({
        success: false,
        message: "Harga tidak boleh kosong atau kurang dari 0"
      });
    }

    const db = req.app.locals.db;
    const userId = req.session.user ? req.session.user.id : null;

    if (!userId) {
      return res.status(401).json({
        success: false,
        message: "Session tidak valid. Silakan login kembali."
      });
    }

    // Insert to database
    const [result] = await db.execute(
      `INSERT INTO master_zakat (nama, harga, kg, created_by, updated_by)
       VALUES (?, ?, ?, ?, ?)`,
      [nama.trim(), parseFloat(harga), parseFloat(kg) || 0, userId, userId]
    );

    res.json({
      success: true,
      message: "Data jenis zakat berhasil ditambahkan",
      data: { id: result.insertId }
    });
  } catch (error) {
    console.error("Error creating master zakat:", error);
    res.status(500).json({
      success: false,
      message: "Terjadi kesalahan saat menyimpan data: " + error.message
    });
  }
});

// PUT /api/master-zakat/:id - Update master zakat
router.put("/api/master-zakat/:id", async (req, res) => {
  const { id } = req.params;
  const { nama, harga, kg } = req.body;

  try {
    // Validation
    if (!nama || nama.trim() === '') {
      return res.status(400).json({
        success: false,
        message: "Nama jenis zakat harus diisi"
      });
    }

    if (harga === undefined || harga === null || harga === '' || parseFloat(harga) < 0) {
      return res.status(400).json({
        success: false,
        message: "Harga tidak boleh kosong atau kurang dari 0"
      });
    }

    const db = req.app.locals.db;
    const userId = req.session.user ? req.session.user.id : null;

    if (!userId) {
      return res.status(401).json({
        success: false,
        message: "Session tidak valid. Silakan login kembali."
      });
    }

    // Check if exists
    const [existing] = await db.execute(
      "SELECT id FROM master_zakat WHERE id = ?",
      [id]
    );

    if (existing.length === 0) {
      return res.status(404).json({
        success: false,
        message: "Data tidak ditemukan"
      });
    }

    // Update database
    await db.execute(
      `UPDATE master_zakat 
       SET nama = ?, harga = ?, kg = ?, updated_by = ?
       WHERE id = ?`,
      [nama.trim(), parseFloat(harga), parseFloat(kg) || 0, userId, id]
    );

    res.json({
      success: true,
      message: "Data jenis zakat berhasil diupdate"
    });
  } catch (error) {
    console.error("Error updating master zakat:", error);
    res.status(500).json({
      success: false,
      message: "Terjadi kesalahan saat mengupdate data: " + error.message
    });
  }
});

// DELETE /api/master-zakat/:id - Delete master zakat
router.delete("/api/master-zakat/:id", async (req, res) => {
  const { id } = req.params;

  try {
    const db = req.app.locals.db;

    // Check if exists
    const [existing] = await db.execute(
      "SELECT id, nama FROM master_zakat WHERE id = ?",
      [id]
    );

    if (existing.length === 0) {
      return res.status(404).json({
        success: false,
        message: "Data tidak ditemukan"
      });
    }

    // Check if being used
    const [usageCount] = await db.execute(
      "SELECT COUNT(*) as count FROM muzakki WHERE master_zakat_id = ?",
      [id]
    );

    if (usageCount[0].count > 0) {
      return res.status(400).json({
        success: false,
        message: `Data ini sedang digunakan oleh ${usageCount[0].count} muzakki. Tidak dapat dihapus.`
      });
    }

    // Delete from database
    await db.execute("DELETE FROM master_zakat WHERE id = ?", [id]);

    res.json({
      success: true,
      message: "Data jenis zakat berhasil dihapus"
    });
  } catch (error) {
    console.error("Error deleting master zakat:", error);
    res.status(500).json({
      success: false,
      message: "Terjadi kesalahan saat menghapus data: " + error.message
    });
  }
});

module.exports = router;

