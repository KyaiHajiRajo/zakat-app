const express = require("express");
const mysql = require("mysql2/promise");
const dotenv = require("dotenv");
const session = require("express-session");
const flash = require("connect-flash");
const methodOverride = require("method-override");
const path = require("path");

// Load environment variables
dotenv.config();

const app = express();

// Setup database connection pool
const dbConfig = {
  host: process.env.DB_HOST || "localhost",
  user: process.env.DB_USER || "root",
  password: process.env.DB_PASSWORD || "",
  database: process.env.DB_NAME || "zakat_fitrah",
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
};

const pool = mysql.createPool(dbConfig);

// Test database connection
async function testConnection() {
  try {
    const connection = await pool.getConnection();
    console.log("✅ Database connected successfully");
    connection.release();
  } catch (error) {
    console.error("❌ Database connection failed:", error.message);
  }
}

testConnection();

// Make database available to all routes
app.locals.db = pool;

// View engine setup
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// Middleware
app.use(express.urlencoded({ extended: true })); // Changed to true to support complex objects

app.use(express.json());

app.use(methodOverride("_method"));
app.use(express.static(path.join(__dirname, "public")));
app.use(
  "/vendor/flatpickr",
  express.static(path.join(__dirname, "node_modules", "flatpickr", "dist"))
);

// Session middleware (stored in MySQL so logins survive server restarts)
const MySQLStore = require("express-mysql-session")(session);
const sessionStore = new MySQLStore({ clearExpired: true }, pool);

app.use(
  session({
    store: sessionStore,
    secret: process.env.SESSION_SECRET || "your-default-secret-key",
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: false, // set to true if using HTTPS
      maxAge: 24 * 60 * 60 * 1000, // 24 hours
    },
  })
);

// CSS mode gelap per halaman (dibuat oleh `npm run build-dark`)
const darkManifestPath = path.join(__dirname, "public", "css", "dark", "manifest.json");
let darkManifest = { mtime: 0, data: {} };

function getDarkManifest() {
  try {
    const { mtimeMs } = require("fs").statSync(darkManifestPath);
    if (mtimeMs !== darkManifest.mtime) {
      darkManifest = {
        mtime: mtimeMs,
        data: JSON.parse(require("fs").readFileSync(darkManifestPath, "utf8")),
      };
    }
  } catch (error) {
    darkManifest = { mtime: 0, data: {} };
  }
  return darkManifest.data;
}

app.use((req, res, next) => {
  const render = res.render.bind(res);
  res.render = (view, ...rest) => {
    const key = String(view).replace(/^\/+/, "").split("/").join("__");
    const version = getDarkManifest()[key];
    if (version) {
      res.locals.darkCss = `/css/dark/${key}.css?v=${version}`;
    }
    return render(view, ...rest);
  };
  next();
});

// Pastikan sesi sudah tersimpan di MySQL sebelum redirect dikirim ke browser
app.use((req, res, next) => {
  const redirect = res.redirect.bind(res);
  res.redirect = (...args) => {
    if (!req.session) {
      return redirect(...args);
    }
    return req.session.save(() => redirect(...args));
  };
  next();
});

// Flash messages middleware
app.use(flash());

// Global variables for all routes
app.use((req, res, next) => {
  // Baca flash sekali, lalu biarkan route yang memanggil req.flash(tipe) mendapat pesan yang sama
  // (jika tidak, pesan sudah terhapus sebelum route sempat membacanya).
  const readFlash = req.flash.bind(req);
  const consumed = {
    success_msg: readFlash("success_msg"),
    error_msg: readFlash("error_msg"),
    error: readFlash("error"),
  };

  req.flash = (type, ...rest) =>
    rest.length === 0 && consumed[type] ? consumed[type] : readFlash(type, ...rest);

  res.locals.success_msg = consumed.success_msg;
  res.locals.error_msg = consumed.error_msg;
  res.locals.error = consumed.error;
  res.locals.user = req.session.user || null;
  next();
});

// Authentication middleware
function isAuthenticated(req, res, next) {
  if (req.session.user) {
    return next();
  }
  req.flash("error_msg", "Silakan login untuk mengakses halaman ini");
  res.redirect("/auth/login");
}

// Admin middleware
function isAdmin(req, res, next) {
  if (req.session.user && req.session.user.role === "admin") {
    return next();
  }
  req.flash(
    "error_msg",
    "Akses ditolak. Hanya admin yang dapat mengakses halaman ini"
  );
  res.redirect("/");
}

// Make middleware available globally
app.locals.isAuthenticated = isAuthenticated;
app.locals.isAdmin = isAdmin;

// Routes
const authRoutes = require("./routes/auth");
const publicReportRoutes = require("./routes/public-report");
const usersRoutes = require("./routes/users");
const profileRoutes = require("./routes/profile");
const muzakkiRoutes = require("./routes/muzakki");
const infakRoutes = require("./routes/infak");
const laporanRoutes = require("./routes/laporan");
const rtRoutes = require("./routes/rt");
const rwRoutes = require("./routes/rw");
const mustahikRoutes = require("./routes/mustahik");
const distribusiRoutes = require("./routes/distribusi");

// Route middleware
app.use("/auth", authRoutes);
app.use("/", publicReportRoutes);
app.use("/users/profile", isAuthenticated, profileRoutes);
app.use("/users", isAuthenticated, isAdmin, usersRoutes);

app.use("/muzakki", isAuthenticated, muzakkiRoutes);

app.use("/infak", isAuthenticated, infakRoutes);
app.use("/laporan", isAuthenticated, laporanRoutes);
app.use("/rt-rw", isAuthenticated, rtRoutes);
app.use("/rt", isAuthenticated, rtRoutes);
app.use("/rw", isAuthenticated, rwRoutes);
app.use("/mustahik", isAuthenticated, mustahikRoutes);
app.use("/distribusi", isAuthenticated, distribusiRoutes);

// Dashboard route
app.get("/dashboard", isAuthenticated, async (req, res) => {
  try {
    const db = req.app.locals.db;
    // Get statistics
    const [statsResult] = await db.execute(`
      SELECT
        (
          SELECT COUNT(*)
          FROM muzakki m2
          WHERE m2.nama_muzakki IS NOT NULL
            AND TRIM(m2.nama_muzakki) <> ''
        ) as total_muzakki,
        COALESCE(SUM(COALESCE(m.jumlah_beras_kg, 0)), 0) as total_beras_kg,
        COALESCE(SUM(COALESCE(m.jumlah_bayar, 0)), 0) as total_jumlah_bayar,
        COUNT(DISTINCT CASE WHEN m.rt_id IS NOT NULL THEN m.rt_id END) as total_rt_aktif,
        COALESCE(SUM(CASE 
          WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
          ELSE m.jumlah_beras_kg * 12000 
        END), 0) as total_zakat,
        COUNT(CASE 
          WHEN m.jumlah_bayar < CASE 
            WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
            ELSE m.jumlah_beras_kg * 12000 
          END THEN 1 
        END) as belum_bayar
      FROM muzakki m
    `);

    // Get infak total separately
    const [infakResult] = await db.execute(`
      SELECT COALESCE(SUM(jumlah), 0) as total_infak FROM infak
    `);

    // Get total users
    const [usersResult] = await db.execute(`
      SELECT COUNT(*) as total_users FROM users
    `);

    // Get recent trend for dashboard chart
    const [trendRows] = await db.execute(`
      SELECT
        DATE_FORMAT(chart.payment_date, '%Y-%m-%d') as tanggal,
        chart.total_muzakki,
        chart.total_bayar,
        chart.total_beras_kg
      FROM (
        SELECT
          COALESCE(m.tanggal, DATE(m.created_at)) as payment_date,
          COUNT(*) as total_muzakki,
          COALESCE(SUM(COALESCE(m.jumlah_bayar, 0)), 0) as total_bayar,
          COALESCE(SUM(COALESCE(m.jumlah_beras_kg, 0)), 0) as total_beras_kg
        FROM muzakki m
        GROUP BY COALESCE(m.tanggal, DATE(m.created_at))
        ORDER BY payment_date DESC
        LIMIT 14
      ) chart
      ORDER BY chart.payment_date ASC
    `);

    // Get recent muzakki
    const [recentMuzakki] = await db.execute(`
      SELECT m.*, rt.nomor_rt as rt_nama,
        COALESCE(NULLIF(TRIM(m.nama_muzakki), ''), CONCAT('Muzakki #', m.id)) as nama,
        CASE
          WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang
          ELSE m.jumlah_beras_kg * 12000
        END as jumlah_zakat,
        CASE
          WHEN m.jumlah_bayar >= CASE
            WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang
            ELSE m.jumlah_beras_kg * 12000
          END THEN 'lunas'
          ELSE 'belum_lunas'
        END as status
      FROM muzakki m
      LEFT JOIN rt ON m.rt_id = rt.id
      ORDER BY COALESCE(m.tanggal, DATE(m.created_at)) DESC, m.created_at DESC
      LIMIT 5
    `);

    // Get RT statistics
    const [rtStats] = await db.execute(`
      SELECT
        rt.nomor_rt,
        COUNT(m.id) as total_muzakki,
        COALESCE(SUM(CASE 
          WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
          ELSE m.jumlah_beras_kg * 12000 
        END), 0) as total_zakat,
        COUNT(CASE 
          WHEN m.jumlah_bayar >= CASE 
            WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
            ELSE m.jumlah_beras_kg * 12000 
          END THEN 1 
        END) as lunas,
        COUNT(CASE 
          WHEN m.jumlah_bayar < CASE 
            WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
            ELSE m.jumlah_beras_kg * 12000 
          END THEN 1 
        END) as belum_bayar
      FROM rt
      LEFT JOIN muzakki m ON rt.id = m.rt_id
      GROUP BY rt.id, rt.nomor_rt
      ORDER BY rt.nomor_rt
      LIMIT 10
    `);

    // Combine stats
    const stats = {
      ...statsResult[0],
      total_infak: infakResult[0].total_infak,
      total_users: usersResult[0].total_users,
    };

    const dashboardTrend = (trendRows || []).map((row) => {
      const tanggal = row.tanggal;
      const parsedDate = tanggal ? new Date(`${tanggal}T00:00:00`) : null;

      return {
        tanggal,
        label:
          parsedDate && !Number.isNaN(parsedDate.getTime())
            ? parsedDate.toLocaleDateString("id-ID", {
                day: "numeric",
                month: "short",
              })
            : tanggal || "-",
        total_muzakki: Number(row.total_muzakki || 0),
        total_bayar: Number(row.total_bayar || 0),
        total_beras_kg: Number(row.total_beras_kg || 0),
      };
    });

    const emptyInsight = {
      tanggal: null,
      label: "-",
      total_muzakki: 0,
      total_bayar: 0,
      total_beras_kg: 0,
    };

    const trendPeakMuzakki =
      dashboardTrend.reduce((best, row) => {
        if (!best || row.total_muzakki > best.total_muzakki) {
          return row;
        }
        return best;
      }, null) || emptyInsight;

    const trendPeakBayar =
      dashboardTrend.reduce((best, row) => {
        if (!best || row.total_bayar > best.total_bayar) {
          return row;
        }
        return best;
      }, null) || emptyInsight;

    const dashboardInsights = {
      active_days: dashboardTrend.length,
      avg_bayar_per_muzakki:
        Number(stats.total_muzakki || 0) > 0
          ? Number(stats.total_jumlah_bayar || 0) / Number(stats.total_muzakki || 1)
          : 0,
      peak_muzakki: trendPeakMuzakki,
      peak_bayar: trendPeakBayar,
    };

    res.render("dashboard", {
      title: "Dashboard - Zakat Fitrah App",
      user: req.session.user,
      stats: stats || {},
      dashboardTrend: dashboardTrend || [],
      dashboardInsights,
      recentMuzakki: recentMuzakki || [],
      rtStats: rtStats || [],
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Dashboard error:", error);
    req.flash("error", "Gagal memuat dashboard");
    res.redirect("/muzakki");
  }
});

// Home route - landing page
app.get("/", async (req, res) => {
  if (req.session.user) {
    res.redirect("/dashboard");
  } else {
    try {
      const db = req.app.locals.db;

      // Get some basic statistics for the landing page
      const [statsResult] = await db.execute(`
        SELECT
          COUNT(DISTINCT m.id) as total_muzakki,
          COALESCE(SUM(CASE 
            WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
            ELSE m.jumlah_beras_kg * 12000 
          END), 0) as total_zakat_terkumpul,
          COUNT(DISTINCT m.rt_id) as total_rt_aktif
        FROM muzakki m
      `);

      const [infakResult] = await db.execute(`
        SELECT COALESCE(SUM(jumlah), 0) as total_infak FROM infak
      `);

      const stats = {
        ...statsResult[0],
        total_infak: infakResult[0].total_infak,
      };

      // Handle query parameters for flash messages
      const success = req.query.success;
      const error = req.query.error;

      res.render("home", {
        title: "Sistem Manajemen Zakat Fitrah",
        stats: stats || {},
        success: success || null,
        error: error || null,
        layout: false, // Gunakan layout khusus untuk home
      });
    } catch (error) {
      console.error("Home page error:", error);
      res.render("home", {
        title: "Sistem Manajemen Zakat Fitrah",
        stats: {
          total_muzakki: 0,
          total_zakat_terkumpul: 0,
          total_rt_aktif: 0,
          total_infak: 0,
        },
        success: req.query.success || null,
        error: req.query.error || null,
        layout: false,
      });
    }
  }
});

// 404 handler
app.use((req, res) => {
  res.status(404).render("error", {
    title: "Halaman Tidak Ditemukan",
    message: "Halaman yang Anda cari tidak ditemukan",
    error: { status: 404 },
  });
});

// Error handler
app.use((error, req, res, next) => {
  console.error("Error:", error);
  res.status(500).render("error", {
    title: "Server Error",
    message: "Terjadi kesalahan pada server",
    error: process.env.NODE_ENV === "development" ? error : {},
  });
});

module.exports = app;
