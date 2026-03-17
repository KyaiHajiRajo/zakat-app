const express = require("express");
const router = express.Router();
const multer = require("multer");
const path = require("path");
const fs = require("fs");

// Configure multer untuk upload bukti foto
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const uploadDir = "public/uploads/bukti-distribusi";
    // Create directory if it doesn't exist
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: function (req, file, cb) {
    // Generate unique filename: distribusi-{timestamp}-{random}.{ext}
    const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);
    const ext = path.extname(file.originalname);
    cb(null, "distribusi-" + uniqueSuffix + ext);
  },
});

// File filter untuk hanya accept gambar
const fileFilter = (req, file, cb) => {
  const allowedTypes = /jpeg|jpg|png|gif/;
  const extname = allowedTypes.test(
    path.extname(file.originalname).toLowerCase()
  );
  const mimetype = allowedTypes.test(file.mimetype);

  if (mimetype && extname) {
    return cb(null, true);
  } else {
    cb(new Error("Hanya file gambar (JPG, JPEG, PNG, GIF) yang diperbolehkan!"));
  }
};

// Initialize multer
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 5 * 1024 * 1024, // Max 5MB
  },
  fileFilter: fileFilter,
});

// Middleware untuk auth
function isAuthenticated(req, res, next) {
  if (req.session && req.session.user) {
    return next();
  }
  res.redirect("/auth/login");
}

// Apply auth middleware to all routes
router.use(isAuthenticated);

// Get all distribusi zakat
router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;

    // Pagination parameters
    const page = parseInt(req.query.page) || 1;
    const limit = 20; // items per page
    const offset = (page - 1) * limit;

    // Get distribusi with mustahik and user info with pagination
    const [distribusi] = await db.execute(
      `
      SELECT dz.*, m.nama as mustahik_nama, m.kategori as mustahik_kategori,
             r.nomor_rt as rt_nama, u.name as user_nama
      FROM distribusi_zakat dz
      LEFT JOIN mustahik m ON dz.mustahik_id = m.id
      LEFT JOIN rt r ON m.rt_id = r.id
      LEFT JOIN users u ON dz.user_id = u.id
      ORDER BY dz.created_at DESC
      LIMIT ${limit} OFFSET ${offset}
    `
    );

    // ✅ Tambahkan query untuk menghitung total distribusi
    const [countResult] = await db.execute(`
      SELECT COUNT(*) as total FROM distribusi_zakat
    `);
    const total = countResult[0].total;
    const totalPages = Math.ceil(total / limit);

    // Get summary statistics
    const [stats] = await db.execute(`
      SELECT 
        COUNT(*) as total_distribusi,
        SUM(CASE WHEN jenis_zakat = 'beras' THEN jumlah ELSE 0 END) as total_beras,
        SUM(CASE WHEN jenis_zakat = 'uang' THEN jumlah ELSE 0 END) as total_uang,
        COUNT(CASE WHEN status = 'pending' THEN 1 END) as pending,
        COUNT(CASE WHEN status = 'disalurkan' THEN 1 END) as disalurkan,
        COUNT(CASE WHEN status = 'diterima' THEN 1 END) as diterima,
        COUNT(CASE WHEN status = 'batal' THEN 1 END) as batal
      FROM distribusi_zakat
    `);

    // Get mustahik list for modal dropdown
    const [mustahiks] = await db.execute(`
      SELECT m.id, m.nama, m.kategori, r.nomor_rt as rt_nama
      FROM mustahik m 
      LEFT JOIN rt r ON m.rt_id = r.id 
      ORDER BY r.nomor_rt, m.nama
    `);

    // ✅ total ikut dikirim agar bisa dipakai di index.ejs
    res.render("distribusi/index", {
      title: "Distribusi Zakat - Zakat Fitrah App",
      user: req.session.user,
      distributions: distribusi,
      stats: stats[0],
      total,
      page,
      totalPages,
      search: req.query.search || "",
      status: req.query.status || "",
      jenis: req.query.jenis || "",
      currentPage: "distribusi",
      mustahiks: mustahiks,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error fetching distribusi:", error);
    req.flash("error", "Gagal mengambil data distribusi zakat");
    res.redirect("/dashboard");
  }
});

// Show create form
router.get("/create", async (req, res) => {
  try {
    const db = req.app.locals.db;

    // Get mustahik list with RT info
    const [mustahikList] = await db.execute(`
      SELECT m.*, r.nomor_rt as rt_nama 
      FROM mustahik m 
      LEFT JOIN rt r ON m.rt_id = r.id 
      ORDER BY r.nomor_rt, m.nama
    `);

    // Get available zakat from muzakki (total collected zakat)
    const [zakatStats] = await db.execute(`
      SELECT 
        SUM(jumlah_uang) as total_uang_tersedia,
        SUM(jumlah_beras_kg) as total_beras_tersedia
      FROM muzakki
    `);

    // Get already distributed amounts
    const [distributedStats] = await db.execute(`
      SELECT 
        SUM(CASE WHEN jenis_zakat = 'uang' AND status != 'batal' THEN jumlah ELSE 0 END) as uang_terdistribusi,
        SUM(CASE WHEN jenis_zakat = 'beras' AND status != 'batal' THEN jumlah ELSE 0 END) as beras_terdistribusi
      FROM distribusi_zakat
    `);

    const availableZakat = {
      uang:
        (zakatStats[0].total_uang_tersedia || 0) -
        (distributedStats[0].uang_terdistribusi || 0),
      beras:
        (zakatStats[0].total_beras_tersedia || 0) -
        (distributedStats[0].beras_terdistribusi || 0),
    };

    res.render("distribusi/create", {
      title: "Tambah Distribusi Zakat - Zakat Fitrah App",
      user: req.session.user,
      mustahiks: mustahikList,
      stockInfo: availableZakat,
      selectedMustahik: "",
      currentPage: "distribusi",
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error loading create form:", error);
    req.flash("error", "Gagal memuat form distribusi zakat");
    res.redirect("/distribusi");
  }
});

// Create new distribusi
router.post("/create", upload.single("bukti_foto"), async (req, res) => {
  try {
    const { mustahik_id, jenis_zakat, jumlah } = req.body;
    const user_id = req.session.user.id;
    const bukti_foto = req.file ? `/uploads/bukti-distribusi/${req.file.filename}` : null;

    // Validation
    if (!mustahik_id || !jenis_zakat || !jumlah) {
      req.flash("error", "Semua field harus diisi");
      return res.redirect("/distribusi/create");
    }

    if (parseFloat(jumlah) <= 0) {
      req.flash("error", "Jumlah distribusi harus lebih dari 0");
      return res.redirect("/distribusi/create");
    }

    const db = req.app.locals.db;

    // Check available zakat
    const stockColumn = jenis_zakat === "uang" ? "jumlah_uang" : "jumlah_beras_kg";
    const [zakatStats] = await db.execute(
      `
      SELECT
        COALESCE(SUM(${stockColumn}), 0) as total_tersedia
      FROM muzakki
      WHERE jenis_zakat = ?
    `,
      [jenis_zakat]
    );

    const [distributedStats] = await db.execute(
      `
      SELECT 
        SUM(jumlah) as total_terdistribusi
      FROM distribusi_zakat 
      WHERE jenis_zakat = ? AND status != 'batal'
    `,
      [jenis_zakat]
    );

    const available =
      (zakatStats[0].total_tersedia || 0) -
      (distributedStats[0].total_terdistribusi || 0);

    if (parseFloat(jumlah) > available) {
      // Delete uploaded file if validation fails
      if (req.file) {
        fs.unlinkSync(req.file.path);
      }
      req.flash(
        "error",
        `Jumlah ${jenis_zakat} yang tersedia hanya ${available.toLocaleString(
          "id-ID"
        )} ${jenis_zakat === "beras" ? "kg" : ""}`
      );
      return res.redirect("/distribusi/create");
    }

    // Insert distribusi
    await db.execute(
      "INSERT INTO distribusi_zakat (mustahik_id, jenis_zakat, jumlah, bukti_foto, user_id, status) VALUES (?, ?, ?, ?, ?, ?)",
      [mustahik_id, jenis_zakat, jumlah, bukti_foto, user_id, "pending"]
    );

    req.flash("success", "Distribusi zakat berhasil ditambahkan");
    res.redirect("/distribusi");
  } catch (error) {
    // Delete uploaded file if error occurs
    if (req.file) {
      try {
        fs.unlinkSync(req.file.path);
      } catch (unlinkError) {
        console.error("Error deleting file:", unlinkError);
      }
    }
    console.error("Error creating distribusi:", error);
    req.flash("error", "Gagal menambahkan distribusi zakat");
    res.redirect("/distribusi/create");
  }
});

// API endpoint to get distribusi by ID (for edit modal)
router.get("/api/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const db = req.app.locals.db;

    const [distribusi] = await db.execute(
      `SELECT dz.*, m.nama as mustahik_nama 
       FROM distribusi_zakat dz
       LEFT JOIN mustahik m ON dz.mustahik_id = m.id
       WHERE dz.id = ?`,
      [id]
    );

    if (distribusi.length === 0) {
      return res.status(404).json({ 
        success: false, 
        message: "Distribusi tidak ditemukan" 
      });
    }

    res.json({ 
      success: true, 
      data: distribusi[0] 
    });
  } catch (error) {
    console.error("Error fetching distribusi:", error);
    res.status(500).json({ 
      success: false, 
      message: "Gagal memuat data distribusi" 
    });
  }
});

// Update distribusi
router.post("/update/:id", upload.single("bukti_foto"), async (req, res) => {
  try {
    const { id } = req.params;
    const { mustahik_id, jenis_zakat, jumlah, status } = req.body;
    const db = req.app.locals.db;

    // Get existing distribusi data
    const [existingData] = await db.execute(
      "SELECT bukti_foto FROM distribusi_zakat WHERE id = ?",
      [id]
    );

    if (existingData.length === 0) {
      if (req.file) {
        fs.unlinkSync(req.file.path);
      }
      req.flash("error", "Distribusi tidak ditemukan");
      return res.redirect("/distribusi");
    }

    // Handle bukti_foto
    let bukti_foto = existingData[0].bukti_foto;
    
    if (req.file) {
      // Delete old file if exists
      if (bukti_foto) {
        const oldFilePath = path.join("public", bukti_foto);
        if (fs.existsSync(oldFilePath)) {
          fs.unlinkSync(oldFilePath);
        }
      }
      bukti_foto = `/uploads/bukti-distribusi/${req.file.filename}`;
    }

    // Validation
    if (!mustahik_id || !jenis_zakat || !jumlah) {
      if (req.file) {
        fs.unlinkSync(req.file.path);
      }
      req.flash("error", "Semua field harus diisi");
      return res.redirect("/distribusi");
    }

    if (parseFloat(jumlah) <= 0) {
      if (req.file) {
        fs.unlinkSync(req.file.path);
      }
      req.flash("error", "Jumlah distribusi harus lebih dari 0");
      return res.redirect("/distribusi");
    }

    // Update distribusi
    await db.execute(
      "UPDATE distribusi_zakat SET mustahik_id = ?, jenis_zakat = ?, jumlah = ?, bukti_foto = ?, status = ? WHERE id = ?",
      [mustahik_id, jenis_zakat, jumlah, bukti_foto, status || "pending", id]
    );

    req.flash("success", "Distribusi zakat berhasil diupdate");
    res.redirect("/distribusi");
  } catch (error) {
    if (req.file) {
      try {
        fs.unlinkSync(req.file.path);
      } catch (unlinkError) {
        console.error("Error deleting file:", unlinkError);
      }
    }
    console.error("Error updating distribusi:", error);
    req.flash("error", "Gagal mengupdate distribusi zakat");
    res.redirect("/distribusi");
  }
});

// Update status distribusi
router.put("/status/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;

    // Validation
    const validStatus = ["pending", "disalurkan", "diterima", "batal"];
    if (!validStatus.includes(status)) {
      req.flash("error", "Status tidak valid");
      return res.redirect("/distribusi");
    }

    const db = req.app.locals.db;

    // Update status
    await db.execute("UPDATE distribusi_zakat SET status = ? WHERE id = ?", [
      status,
      id,
    ]);

    req.flash("success", `Status distribusi berhasil diubah menjadi ${status}`);
    res.redirect("/distribusi");
  } catch (error) {
    console.error("Error updating status:", error);
    req.flash("error", "Gagal mengubah status distribusi");
    res.redirect("/distribusi");
  }
});

// Delete distribusi
router.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const db = req.app.locals.db;

    // Check if distribusi can be deleted (only pending or batal)
    const [distribusi] = await db.execute(
      "SELECT status FROM distribusi_zakat WHERE id = ?",
      [id]
    );

    if (distribusi.length === 0) {
      req.flash("error", "Data distribusi tidak ditemukan");
      return res.redirect("/distribusi");
    }

    if (!["pending", "batal"].includes(distribusi[0].status)) {
      req.flash(
        "error",
        "Hanya distribusi dengan status pending atau batal yang dapat dihapus"
      );
      return res.redirect("/distribusi");
    }

    // Delete distribusi
    await db.execute("DELETE FROM distribusi_zakat WHERE id = ?", [id]);

    req.flash("success", "Data distribusi berhasil dihapus");
    res.redirect("/distribusi");
  } catch (error) {
    console.error("Error deleting distribusi:", error);
    req.flash("error", "Gagal menghapus data distribusi");
    res.redirect("/distribusi");
  }
});

// Show distribusi detail
router.get("/detail/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const db = req.app.locals.db;

    // Get distribusi with complete info
    const [distribusi] = await db.execute(
      `
      SELECT dz.*, m.nama as mustahik_nama, m.kategori as mustahik_kategori,
             r.nomor_rt as rt_nama, u.name as user_nama
      FROM distribusi_zakat dz
      LEFT JOIN mustahik m ON dz.mustahik_id = m.id
      LEFT JOIN rt r ON m.rt_id = r.id
      LEFT JOIN users u ON dz.user_id = u.id
      WHERE dz.id = ?
    `,
      [id]
    );

    if (distribusi.length === 0) {
      req.flash("error", "Data distribusi tidak ditemukan");
      return res.redirect("/distribusi");
    }

    res.render("distribusi/detail", {
      title: `Detail Distribusi - Zakat Fitrah App`,
      user: req.session.user,
      distribusi: distribusi[0],
      currentPage: "distribusi",
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error fetching distribusi detail:", error);
    req.flash("error", "Gagal mengambil detail distribusi");
    res.redirect("/distribusi");
  }
});

// Laporan distribusi per kategori mustahik
router.get("/laporan", async (req, res) => {
  try {
    const db = req.app.locals.db;

    // Get distribution summary by category
    const [categoryStats] = await db.execute(`
      SELECT 
        m.kategori,
        COUNT(dz.id) as total_distribusi,
        SUM(CASE WHEN dz.jenis_zakat = 'beras' THEN dz.jumlah ELSE 0 END) as total_beras,
        SUM(CASE WHEN dz.jenis_zakat = 'uang' THEN dz.jumlah ELSE 0 END) as total_uang,
        COUNT(CASE WHEN dz.status = 'diterima' THEN 1 END) as sudah_diterima,
        COUNT(CASE WHEN dz.status = 'pending' THEN 1 END) as pending
      FROM mustahik m
      LEFT JOIN distribusi_zakat dz ON m.id = dz.mustahik_id AND dz.status != 'batal'
      GROUP BY m.kategori
      ORDER BY m.kategori
    `);

    // Get distribution by RT
    const [rtStats] = await db.execute(`
      SELECT 
        COALESCE(r.nomor_rt, 'Tidak Ada RT') as rt_nama,
        COUNT(dz.id) as total_distribusi,
        SUM(CASE WHEN dz.jenis_zakat = 'beras' THEN dz.jumlah ELSE 0 END) as total_beras,
        SUM(CASE WHEN dz.jenis_zakat = 'uang' THEN dz.jumlah ELSE 0 END) as total_uang
      FROM mustahik m
      LEFT JOIN rt r ON m.rt_id = r.id
      LEFT JOIN distribusi_zakat dz ON m.id = dz.mustahik_id AND dz.status != 'batal'
      GROUP BY r.id, r.nomor_rt
      ORDER BY r.nomor_rt
    `);

    res.render("distribusi/laporan", {
      title: "Laporan Distribusi Zakat - Zakat Fitrah App",
      user: req.session.user,
      categoryStats,
      rtStats,
      currentPage: "distribusi",
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error generating report:", error);
    req.flash("error", "Gagal membuat laporan distribusi");
    res.redirect("/distribusi");
  }
});

module.exports = router;
