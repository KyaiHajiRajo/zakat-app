const express = require("express");
const { reply, safeReturn } = require("../utils/respond");
const router = express.Router();

// GET /rw - Tampilkan daftar RW
router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const [rows] = await db.execute(`
      SELECT 
        rw.*,
        COUNT(DISTINCT rt.id) as total_rt,
        COUNT(DISTINCT m.id) as total_muzakki
      FROM rw 
      LEFT JOIN rt ON rw.id = rt.rw_id 
      LEFT JOIN muzakki m ON rt.id = m.rt_id 
      GROUP BY rw.id 
      ORDER BY rw.nomor_rw ASC
    `);

    res.render("rw/index", {
      title: "Data RW - Zakat Fitrah App",
      user: req.session.user,
      rwList: rows,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error fetching RW data:", error);
    req.flash("error", "Gagal mengambil data RW");
    res.redirect("/rw");
  }
});

// GET /rw/form-data?id= - Data untuk modal tambah/ubah RW (JSON)
router.get("/form-data", async (req, res) => {
  try {
    const db = req.app.locals.db;
    let rw = null;
    if (req.query.id) {
      const [rows] = await db.execute(
        "SELECT id, nomor_rw, ketua_rw, keterangan FROM rw WHERE id = ?",
        [req.query.id]
      );
      if (rows.length === 0) {
        return res.status(404).json({ ok: false, message: "RW tidak ditemukan" });
      }
      rw = rows[0];
    }
    res.json({ ok: true, rw });
  } catch (error) {
    console.error("Error loading RW form data:", error);
    res.status(500).json({ ok: false, message: "Gagal memuat data form" });
  }
});

// GET /rw/create - Form tambah RW kini berupa modal di halaman daftar
router.get("/create", (req, res) => {
  res.redirect("/rw?new=1");
});

// Validasi bersama untuk tambah/ubah RW
async function validateRw(db, body, excludeId) {
  const nomorRw = String(body.nomor_rw || "").trim();
  const ketuaRw = String(body.ketua_rw || "").trim();
  const keterangan = String(body.keterangan || "").trim();

  if (!nomorRw || !ketuaRw) return { error: "Nomor RW dan Ketua RW harus diisi" };
  if (nomorRw.length > 10) return { error: "Nomor RW maksimal 10 karakter" };
  if (ketuaRw.length > 100) return { error: "Nama ketua RW maksimal 100 karakter" };
  if (keterangan.length > 100) return { error: "Keterangan maksimal 100 karakter" };

  const [existing] = await db.execute(
    "SELECT id FROM rw WHERE nomor_rw = ? AND id != ?",
    [nomorRw, excludeId || 0]
  );
  if (existing.length > 0) return { error: "Nomor RW sudah terdaftar" };

  return { nomorRw, ketuaRw, keterangan: keterangan || null };
}

// POST /rw/create - Simpan RW baru
router.post("/create", async (req, res) => {
  const back = safeReturn(req.body.return_to, "/rw");
  try {
    const db = req.app.locals.db;
    const data = await validateRw(db, req.body);
    if (data.error) {
      return reply(req, res, { ok: false, message: data.error, redirect: back });
    }

    await db.execute(
      "INSERT INTO rw (nomor_rw, ketua_rw, keterangan) VALUES (?, ?, ?)",
      [data.nomorRw, data.ketuaRw, data.keterangan]
    );

    reply(req, res, {
      ok: true,
      message: `${data.nomorRw} berhasil ditambahkan`,
      redirect: back,
    });
  } catch (error) {
    console.error("Error creating RW:", error);
    reply(req, res, {
      ok: false,
      message: "Gagal menambahkan RW",
      redirect: back,
      status: 500,
    });
  }
});

// GET /rw/:id/edit - Form edit RW kini berupa modal di halaman daftar
router.get("/:id/edit", (req, res) => {
  if (!/^\d+$/.test(req.params.id)) return res.redirect("/rw");
  res.redirect(`/rw?edit=${req.params.id}`);
});

// POST /rw/:id/edit - Update RW
router.post("/:id/edit", async (req, res) => {
  const back = safeReturn(req.body.return_to, "/rw");
  try {
    const db = req.app.locals.db;
    const [current] = await db.execute("SELECT id FROM rw WHERE id = ?", [
      req.params.id,
    ]);
    if (current.length === 0) {
      return reply(req, res, {
        ok: false,
        message: "RW tidak ditemukan",
        redirect: back,
        status: 404,
      });
    }

    const data = await validateRw(db, req.body, req.params.id);
    if (data.error) {
      return reply(req, res, { ok: false, message: data.error, redirect: back });
    }

    await db.execute(
      "UPDATE rw SET nomor_rw = ?, ketua_rw = ?, keterangan = ? WHERE id = ?",
      [data.nomorRw, data.ketuaRw, data.keterangan, req.params.id]
    );

    reply(req, res, {
      ok: true,
      message: `${data.nomorRw} berhasil diperbarui`,
      redirect: back,
    });
  } catch (error) {
    console.error("Error updating RW:", error);
    reply(req, res, {
      ok: false,
      message: "Gagal memperbarui RW",
      redirect: back,
      status: 500,
    });
  }
});

// POST /rw/:id/delete - Hapus RW
router.post("/:id/delete", async (req, res) => {
  try {
    const db = req.app.locals.db;
    // Cek apakah RW masih memiliki RT
    const [rt] = await db.execute(
      "SELECT COUNT(*) as count FROM rt WHERE rw_id = ?",
      [req.params.id]
    );

    if (rt[0].count > 0) {
      req.flash(
        "error",
        "RW tidak dapat dihapus karena masih memiliki data RT"
      );
      return res.redirect("/rw");
    }

    // Ambil data RW untuk pesan konfirmasi
    const [rwData] = await db.execute("SELECT nomor_rw FROM rw WHERE id = ?", [
      req.params.id,
    ]);

    if (rwData.length === 0) {
      req.flash("error", "RW tidak ditemukan");
      return res.redirect("/rw");
    }

    // Hapus RW
    await db.execute("DELETE FROM rw WHERE id = ?", [req.params.id]);

    req.flash("success", `RW ${rwData[0].nomor_rw} berhasil dihapus`);
    res.redirect("/rw");
  } catch (error) {
    console.error("Error deleting RW:", error);
    req.flash("error", "Gagal menghapus RW");
    res.redirect("/rw");
  }
});

// GET /rw/:id/detail - Detail RW dengan daftar RT
router.get("/:id/detail", async (req, res) => {
  try {
    const db = req.app.locals.db;
    // Ambil data RW
    const [rwData] = await db.execute("SELECT * FROM rw WHERE id = ?", [
      req.params.id,
    ]);

    if (rwData.length === 0) {
      req.flash("error", "RW tidak ditemukan");
      return res.redirect("/rw");
    }

    // Ambil daftar RT di RW ini
    const [rtData] = await db.execute(
      `
      SELECT 
        rt.*,
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
        END) as muzakki_lunas
      FROM rt 
      LEFT JOIN muzakki m ON rt.id = m.rt_id 
      WHERE rt.rw_id = ? 
      GROUP BY rt.id 
      ORDER BY rt.nomor_rt ASC
    `,
      [req.params.id]
    );

    // Hitung statistik
    const stats = {
      total_rt: rtData.length,
      total_muzakki: rtData.reduce((sum, rt) => sum + rt.total_muzakki, 0),
      total_zakat: rtData.reduce((sum, rt) => sum + parseFloat(rt.total_zakat || 0), 0),
      muzakki_lunas: rtData.reduce((sum, rt) => sum + rt.muzakki_lunas, 0),
    };

    res.render("rw/detail", {
      title: `Detail RW ${rwData[0].nomor_rw} - Zakat Fitrah App`,
      user: req.session.user,
      rw: rwData[0],
      rtList: rtData,
      stats: stats,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error fetching RW detail:", error);
    req.flash("error", "Gagal mengambil detail RW");
    res.redirect("/rw");
  }
});

module.exports = router;