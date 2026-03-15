const express = require("express");
const router = express.Router();

// Constants for zakat calculation
const ZAKAT_BERAS_PER_JIWA = 2.5; // kg per jiwa
const ZAKAT_UANG_PER_JIWA = 45000; // Rp per jiwa

function getMasterZakatJenis(masterZakat = {}) {
  const kg = parseFloat(masterZakat.kg) || 0;
  const nama = String(masterZakat.nama || "").trim();

  return kg > 0 || /\bberas\b/i.test(nama) ? "beras" : "uang";
}

function getMasterZakatBerasKg(masterZakat = {}) {
  const kg = parseFloat(masterZakat.kg) || 0;
  return kg > 0 ? kg : ZAKAT_BERAS_PER_JIWA;
}

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

function parseNullableDecimal(value) {
  if (value === null || value === undefined) return null;

  const normalized = String(value).trim();
  if (!normalized) {
    return null;
  }

  const parsed = parseFloat(normalized);
  return Number.isNaN(parsed) ? null : parsed;
}

function getNamaMuzakkiSql(alias = "m") {
  return `COALESCE(NULLIF(TRIM(${alias}.nama_muzakki), ''), '-')`;
}

function getJumlahNamaMuzakkiSql(alias = "m") {
  return `CASE WHEN ${alias}.nama_muzakki IS NOT NULL AND TRIM(${alias}.nama_muzakki) <> '' THEN 1 ELSE 0 END`;
}

function getPrimaryMuzakkiIdentity(body = {}) {
  const muzakkiArray = normalizeMuzakkiPayload(body.muzakki);
  const firstMuzakki =
    muzakkiArray.find(
      (item) => item && typeof item === "object" && item.nama && item.nama.trim()
    ) || muzakkiArray[0] || {};

  const nama_muzakki = String(body.nama_muzakki || firstMuzakki.nama || "").trim();
  const bin_binti = body.bin_binti || firstMuzakki.bin_binti || null;
  const nama_orang_tua = String(
    body.nama_orang_tua || firstMuzakki.nama_orang_tua || ""
  ).trim();

  return {
    nama_muzakki,
    bin_binti,
    nama_orang_tua: nama_orang_tua || null,
  };
}

function getSubmittedMuzakkiEntries(body = {}) {
  const muzakkiArray = normalizeMuzakkiPayload(body.muzakki);
  const rawEntries =
    muzakkiArray.length > 0
      ? muzakkiArray
      : [
          {
            nama: body.nama_muzakki,
            bin_binti: body.bin_binti,
            nama_orang_tua: body.nama_orang_tua,
            master_zakat_id: body.master_zakat_id,
            jumlah_bayar: body.jumlah_bayar,
          },
        ];

  return rawEntries
    .filter((item) => item && typeof item === "object")
    .map((item) => {
      const nama_muzakki = String(item.nama || item.nama_muzakki || "").trim();
      const bin_binti = item.bin_binti || null;
      const nama_orang_tua = String(item.nama_orang_tua || "").trim();
      const master_zakat_id = String(item.master_zakat_id || "").trim();
      const jumlah_bayar = parseNullableDecimal(item.jumlah_bayar);

      return {
        nama_muzakki,
        bin_binti,
        nama_orang_tua: nama_orang_tua || null,
        master_zakat_id,
        jumlah_bayar,
      };
    })
    .filter(
      (item) =>
        item.nama_muzakki ||
        item.bin_binti ||
        item.nama_orang_tua ||
        item.master_zakat_id ||
        item.jumlah_bayar !== null
    );
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

async function getTanggalRtTabs(db, selectedTanggal) {
  if (!selectedTanggal) {
    return [];
  }

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

  return tanggalTabs || [];
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
                    ${getJumlahNamaMuzakkiSql("m")} as total_nama_muzakki,
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
                SUM(${getJumlahNamaMuzakkiSql("m")}) as total_nama_muzakki_all
            FROM muzakki m
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

    const navTabs = await getTanggalRtTabs(db, selectedTanggal);
    if (navTabs.length === 0) {
      req.flash("error_msg", "Tidak ada data RT pada tanggal tersebut");
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
        ${getNamaMuzakkiSql("m")} as nama_muzakki_list,
        ${getJumlahNamaMuzakkiSql("m")} as jumlah_muzakki,
        infak_agg.infak_id,
        COALESCE(infak_agg.infak_jumlah, 0) as infak_jumlah
      FROM muzakki m
      LEFT JOIN rt r ON m.rt_id = r.id
      LEFT JOIN users u ON m.user_id = u.id
      LEFT JOIN (
        SELECT
          muzakki_id,
          MAX(id) as infak_id,
          SUM(jumlah) as infak_jumlah
        FROM infak
        GROUP BY muzakki_id
      ) infak_agg ON m.id = infak_agg.muzakki_id
      WHERE ${getMuzakkiTanggalSql("m")} = ?
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
      navTabs,
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
    const isAllTanggalView = Boolean(selectedTanggal) && req.query.scope !== "rt";

    // Get RT info
    const [rtInfo] = await db.execute("SELECT * FROM rt WHERE id = ?", [rtId]);

    if (rtInfo.length === 0) {
      req.flash("error_msg", "Data RT tidak ditemukan");
      return res.redirect("/muzakki");
    }

    let navTabs = [];

    if (selectedTanggal) {
      navTabs = await getTanggalRtTabs(db, selectedTanggal);

      if (navTabs.length === 0) {
        req.flash("error_msg", "Tidak ada data RT pada tanggal tersebut");
        return res.redirect("/muzakki");
      }

      const isRtIncluded = navTabs.some((item) => String(item.id) === String(rtId));
      if (!isAllTanggalView && !isRtIncluded) {
        return res.redirect(`/muzakki/rt/${navTabs[0].id}?tanggal=${selectedTanggal}&scope=rt`);
      }
    }

    // Get muzakki data for this RT or all RT on selected tanggal
    const muzakkiParams = [];
    let muzakkiWhereClause = "";

    if (isAllTanggalView) {
      muzakkiWhereClause = `WHERE ${getMuzakkiTanggalSql("m")} = ?`;
      muzakkiParams.push(selectedTanggal);
    } else {
      muzakkiWhereClause = "WHERE m.rt_id = ?";
      muzakkiParams.push(rtId);

      if (selectedTanggal) {
        muzakkiWhereClause += ` AND ${getMuzakkiTanggalSql("m")} = ?`;
        muzakkiParams.push(selectedTanggal);
      }
    }

    const [muzakki] = await db.execute(
      `
      SELECT
        m.*,
        ${getMuzakkiTanggalSelectSql("m")} as tanggal,
        r.nomor_rt,
        r.ketua_rt,
        u.name as pencatat_name,
        ${getNamaMuzakkiSql("m")} as nama_muzakki_list,
        ${getJumlahNamaMuzakkiSql("m")} as jumlah_muzakki,
        infak_agg.infak_id,
        COALESCE(infak_agg.infak_jumlah, 0) as infak_jumlah
      FROM muzakki m
      LEFT JOIN rt r ON m.rt_id = r.id
      LEFT JOIN users u ON m.user_id = u.id
      LEFT JOIN (
        SELECT
          muzakki_id,
          MAX(id) as infak_id,
          SUM(jumlah) as infak_jumlah
        FROM infak
        GROUP BY muzakki_id
      ) infak_agg ON m.id = infak_agg.muzakki_id
      ${muzakkiWhereClause}
      ORDER BY ${isAllTanggalView ? "r.nomor_rt ASC," : ""} ${getMuzakkiTanggalSql("m")} DESC, m.created_at DESC
    `,
      muzakkiParams
    );
    const displayMuzakki = muzakki.map((item) => ({
      ...item,
      group_id: item.id,
      nama_muzakki_display: item.nama_muzakki_list || "-",
      nama_muzakki_search: item.nama_muzakki_list || "",
      nama_muzakki_parent_list: item.nama_muzakki_list || "-",
      jumlah_muzakki_parent: parseInt(item.jumlah_muzakki, 10) || 0,
      detail_index: 1,
      detail_count: 1,
    }));

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
      title: isAllTanggalView
        ? `Detail Muzakki Semua RT ${selectedTanggal ? `- ${selectedTanggal}` : ""} - Zakat Fitrah`
        : `Detail Muzakki RT ${rtInfo[0].nomor_rt} - Zakat Fitrah`,
      layout: "layouts/main",
      rt: rtInfo[0],
      muzakki,
      displayMuzakki,
      navTabs,
      selectedTanggal,
      selectedTanggalLabel: formatTanggalLabel(selectedTanggal),
      isAllTanggalView,
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
    const isAllTanggalView = Boolean(selectedTanggal) && req.query.scope !== "rt";

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
    const exportParams = [];
    let exportWhereClause = "";

    if (isAllTanggalView) {
      exportWhereClause = `WHERE ${getMuzakkiTanggalSql("m")} = ?`;
      exportParams.push(selectedTanggal);
    } else {
      exportWhereClause = "WHERE m.rt_id = ?";
      exportParams.push(rtId);

      if (selectedTanggal) {
        exportWhereClause += ` AND ${getMuzakkiTanggalSql("m")} = ?`;
        exportParams.push(selectedTanggal);
      }
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
        r.nomor_rt,
        u.name as pencatat_name,
        ${getNamaMuzakkiSql("m")} as nama_muzakki_list
      FROM muzakki m
      LEFT JOIN rt r ON m.rt_id = r.id
      LEFT JOIN users u ON m.user_id = u.id
      ${exportWhereClause}
      ORDER BY ${isAllTanggalView ? "r.nomor_rt ASC," : ""} ${getMuzakkiTanggalSql("m")} DESC, m.created_at DESC
      `,
      exportParams
    );

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Zakat Fitrah App";
    workbook.lastModifiedBy = "System";
    workbook.created = new Date();
    workbook.modified = new Date();

    const worksheet = workbook.addWorksheet(
      normalizeExcelSheetName(isAllTanggalView ? `Semua RT ${selectedTanggal}` : `RT ${rt.nomor_rt}`)
    );

    worksheet.mergeCells("A1:J1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = isAllTanggalView
      ? `DATA MUZAKKI SEMUA RT${selectedTanggal ? ` ${sanitizeForExcel(selectedTanggal)}` : ""}`
      : `DATA MUZAKKI RT ${sanitizeForExcel(rt.nomor_rt)}`;
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
    ketuaCell.value = isAllTanggalView
      ? `Tanggal: ${sanitizeForExcel(formatTanggalLabel(selectedTanggal) || selectedTanggal || "-")}`
      : `Ketua RT: ${sanitizeForExcel(rt.ketua_rt) || "-"}`;
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
    jumlah_jiwa,
    catatan,
  } = req.body;

  try {
    const submittedMuzakki = getSubmittedMuzakkiEntries(req.body);

    // Debug log untuk melihat data yang diterima
    console.log("Received data:", {
      rt_id,
      tanggal,
      muzakki: req.body.muzakki,
      jumlah_jiwa,
      submittedMuzakki,
      catatan,
    });

    // Input validation
    if (!rt_id || !tanggal || !jumlah_jiwa) {
      req.flash("error_msg", "Semua field yang wajib harus diisi");
      return res.redirect("/muzakki/create");
    }

    const tanggalZakat = normalizeMuzakkiDateInput(tanggal);
    if (!tanggalZakat) {
      req.flash("error_msg", "Tanggal zakat tidak valid");
      return res.redirect("/muzakki/create");
    }

    if (submittedMuzakki.length === 0) {
      req.flash("error_msg", "Minimal harus ada 1 muzakki");
      return res.redirect("/muzakki/create");
    }

    const invalidEntryIndex = submittedMuzakki.findIndex(
      (item) =>
        !item.nama_muzakki ||
        !item.master_zakat_id ||
        (item.jumlah_bayar !== null && item.jumlah_bayar < 0)
    );

    if (invalidEntryIndex !== -1) {
      req.flash(
        "error_msg",
        `Data muzakki #${invalidEntryIndex + 1} belum lengkap atau jumlah bayar tidak valid`
      );
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
      const masterZakatIds = [
        ...new Set(submittedMuzakki.map((item) => item.master_zakat_id)),
      ];

      const placeholders = masterZakatIds.map(() => "?").join(", ");
      const [masterZakat] = await connection.execute(
        `SELECT * FROM master_zakat WHERE id IN (${placeholders})`,
        masterZakatIds
      );

      if (masterZakat.length !== masterZakatIds.length) {
        req.flash("error_msg", "Jenis zakat tidak valid");
        return res.redirect("/muzakki/create");
      }

      const masterZakatMap = new Map(
        masterZakat.map((item) => [String(item.id), item])
      );

      await connection.beginTransaction();

      const insertedIds = [];

      for (const [index, entry] of submittedMuzakki.entries()) {
        const zakatData = masterZakatMap.get(entry.master_zakat_id);
        if (!zakatData) {
          throw new Error(`Jenis zakat untuk muzakki #${index + 1} tidak ditemukan`);
        }

        const jiwa = 1;
        const bayar = entry.jumlah_bayar;
        const jenis_zakat = getMasterZakatJenis(zakatData);

        let jumlah_beras_kg = null;
        let jumlah_uang = null;
        let kewajiban = 0;

        if (jenis_zakat === "beras") {
          jumlah_beras_kg = jiwa * getMasterZakatBerasKg(zakatData);
          kewajiban = jiwa * parseFloat(zakatData.harga);
        } else {
          jumlah_uang = jiwa * parseFloat(zakatData.harga);
          kewajiban = jumlah_uang;
        }

        const kembalian = bayar === null ? null : Math.max(0, bayar - kewajiban);

        const [result] = await connection.execute(
          `
          INSERT INTO muzakki 
          (rt_id, nama_muzakki, bin_binti, nama_orang_tua, jumlah_jiwa, jenis_zakat, master_zakat_id, jumlah_beras_kg, jumlah_uang, jumlah_bayar, kembalian, catatan, user_id, tanggal)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            rt_id,
            entry.nama_muzakki,
            entry.bin_binti,
            entry.nama_orang_tua,
            jiwa,
            jenis_zakat,
            entry.master_zakat_id,
            jumlah_beras_kg,
            jumlah_uang,
            bayar,
            kembalian,
            catatan || null,
            userId,
            tanggalZakat,
          ]
        );

        insertedIds.push(result.insertId);
      }

      // Commit transaction
      await connection.commit();
      console.log("Transaction committed successfully");

      req.flash(
        "success_msg",
        submittedMuzakki.length > 1
          ? `${submittedMuzakki.length} data muzakki berhasil ditambahkan`
          : "Data muzakki berhasil ditambahkan"
      );

      if (insertedIds.length === 1) {
        return res.redirect(`/muzakki/${insertedIds[0]}`);
      }

      return res.redirect(`/muzakki/tanggal/${tanggalZakat}`);
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

    const muzakkiDetails = [
      {
        nama_muzakki: muzakki[0].nama_muzakki,
        bin_binti: muzakki[0].bin_binti,
        nama_orang_tua: muzakki[0].nama_orang_tua,
        master_zakat_id: muzakki[0].master_zakat_id,
      },
    ];

    const selectedMasterZakatId = muzakki[0].master_zakat_id || null;

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
      selectedMasterZakatId,
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
    jumlah_jiwa,
    master_zakat_id,
    jumlah_bayar,
    catatan,
  } = req.body;

  try {
    // Input validation
    if (!rt_id || !tanggal || !jumlah_jiwa || !master_zakat_id) {
      req.flash("error_msg", "Semua field yang wajib harus diisi");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    const tanggalZakat = normalizeMuzakkiDateInput(tanggal);
    if (!tanggalZakat) {
      req.flash("error_msg", "Tanggal zakat tidak valid");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    const identity = getPrimaryMuzakkiIdentity(req.body);
    if (!identity.nama_muzakki) {
      req.flash("error_msg", "Nama muzakki tidak boleh kosong");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    const jiwa = 1;
    const bayar = parseNullableDecimal(jumlah_bayar);

    if (jumlah_bayar !== undefined && String(jumlah_bayar).trim() !== "" && bayar === null) {
      req.flash("error_msg", "Jumlah bayar tidak valid");
      return res.redirect(`/muzakki/${id}/edit`);
    }

    if (bayar !== null && bayar < 0) {
      req.flash("error_msg", "Jumlah bayar tidak boleh negatif");
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
      const jenis_zakat = getMasterZakatJenis(zakatData);

      let jumlah_beras_kg = null;
      let jumlah_uang = null;
      let kewajiban = 0;

      if (jenis_zakat === "beras") {
        jumlah_beras_kg = jiwa * getMasterZakatBerasKg(zakatData);
        kewajiban = jiwa * parseFloat(zakatData.harga);
      } else if (jenis_zakat === "uang") {
        jumlah_uang = jiwa * parseFloat(zakatData.harga);
        kewajiban = jumlah_uang;
      }

      const kembalian = bayar === null ? null : Math.max(0, bayar - kewajiban);

      await connection.beginTransaction();

      // Update muzakki main record
      await connection.execute(
        `
        UPDATE muzakki 
        SET rt_id = ?, nama_muzakki = ?, bin_binti = ?, nama_orang_tua = ?,
            jumlah_jiwa = ?, jenis_zakat = ?, master_zakat_id = ?,
            jumlah_beras_kg = ?, jumlah_uang = ?, jumlah_bayar = ?, 
            kembalian = ?, catatan = ?, tanggal = ?
        WHERE id = ?
        `,
        [
          rt_id,
          identity.nama_muzakki,
          identity.bin_binti,
          identity.nama_orang_tua,
          jiwa,
          jenis_zakat,
          master_zakat_id,
          jumlah_beras_kg,
          jumlah_uang,
          bayar,
          kembalian,
          catatan,
          tanggalZakat,
          id,
        ]
      );

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
        "Data muzakki berhasil diupdate"
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
          errors.push(`Muzakki ${muzakkiData.nama_muzakki || muzakkiId} sudah memiliki infak`);
          continue;
        }

        if (muzakkiData.kembalian <= 0) {
          errors.push(`Muzakki ${muzakkiData.nama_muzakki || muzakkiId} tidak memiliki kembalian`);
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
          ${getNamaMuzakkiSql("m")} as nama_muzakki_list
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
    const muzakkiDetails = [
      {
        nama_muzakki: muzakki[0].nama_muzakki,
        bin_binti: muzakki[0].bin_binti,
        nama_orang_tua: muzakki[0].nama_orang_tua,
        master_zakat_id: muzakki[0].master_zakat_id,
      },
    ];

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

