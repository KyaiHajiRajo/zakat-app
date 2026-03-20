const express = require("express");
const router = express.Router();

// GET /rt - Tampilkan daftar RT
router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const [rtRows] = await db.execute(`
      SELECT 
        rt.*,
        rw.nomor_rw,
        rw.ketua_rw,
        COUNT(m.id) as total_muzakki,
        COALESCE(SUM(CASE 
          WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
          ELSE m.jumlah_beras_kg * 12000 
        END), 0) as total_zakat,
        COALESCE(SUM(COALESCE(m.jumlah_bayar, 0)), 0) as total_jumlah_bayar,
        COUNT(CASE WHEN m.jenis_zakat = 'uang' THEN 1 END) as total_muzakki_uang,
        COUNT(CASE WHEN m.jenis_zakat = 'beras' THEN 1 END) as total_muzakki_beras,
        COUNT(CASE 
          WHEN m.jumlah_bayar >= CASE 
            WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
            ELSE m.jumlah_beras_kg * 12000 
          END THEN 1 
        END) as muzakki_lunas,
        COUNT(CASE 
          WHEN m.jumlah_bayar < CASE 
            WHEN m.jenis_zakat = 'uang' THEN m.jumlah_uang 
            ELSE m.jumlah_beras_kg * 12000 
          END THEN 1 
        END) as muzakki_belum_lunas
      FROM rt 
      LEFT JOIN rw ON rt.rw_id = rw.id
      LEFT JOIN muzakki m ON rt.id = m.rt_id 
      GROUP BY rt.id 
      ORDER BY rt.nomor_rt ASC
    `);

    const [rwRows] = await db.execute(`
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

    res.render("rt/index", {
      title: "Data RT - Zakat Fitrah App",
      user: req.session.user,
      rtList: rtRows,
      rwList: rwRows,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error fetching RT data:", error);
    req.flash("error", "Gagal mengambil data RT");
    res.redirect("/dashboard");
  }
});

// GET /rt/create - Form tambah RT
router.get("/create", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const [rwRows] = await db.execute(
      "SELECT id, nomor_rw, ketua_rw FROM rw ORDER BY nomor_rw ASC"
    );

    res.render("rt/create", {
      title: "Tambah RT - Zakat Fitrah App",
      user: req.session.user,
      rwList: rwRows,
      error: req.flash("error"),
      success: req.flash("success"),
    });
  } catch (error) {
    console.error("Error fetching RW data for RT create:", error);
    req.flash("error", "Gagal mengambil data RW");
    res.redirect("/rt");
  }
});

// POST /rt/create - Simpan RT baru
router.post("/create", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { rw_id, nomor_rt, ketua_rt, keterangan } = req.body;

    // Validasi input
    if (!rw_id || !nomor_rt || !ketua_rt) {
      req.flash("error", "RW, Nomor RT dan Ketua RT harus diisi");
      return res.redirect("/rt/create");
    }

    // Cek apakah nomor RT sudah ada
    const [existing] = await db.execute(
      "SELECT id FROM rt WHERE nomor_rt = ?",
      [nomor_rt]
    );

    if (existing.length > 0) {
      req.flash("error", "Nomor RT sudah terdaftar");
      return res.redirect("/rt/create");
    }

    // Insert RT baru
    await db.execute(
      "INSERT INTO rt (rw_id, nomor_rt, ketua_rt, keterangan) VALUES (?, ?, ?, ?)",
      [rw_id, nomor_rt, ketua_rt, keterangan || null]
    );

    req.flash("success", `RT ${nomor_rt} berhasil ditambahkan`);
    res.redirect("/rt");
  } catch (error) {
    console.error("Error creating RT:", error);
    req.flash("error", "Gagal menambahkan RT");
    res.redirect("/rt/create");
  }
});

// GET /rt/:id/edit - Form edit RT
router.get("/:id/edit", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const [rtRows] = await db.execute("SELECT * FROM rt WHERE id = ?", [
      req.params.id,
    ]);

    if (rtRows.length === 0) {
      req.flash("error", "RT tidak ditemukan");
      return res.redirect("/rt");
    }

    const [rwRows] = await db.execute(
      "SELECT id, nomor_rw, ketua_rw FROM rw ORDER BY nomor_rw ASC"
    );

    res.render("rt/edit", {
      title: "Edit RT - Zakat Fitrah App",
      user: req.session.user,
      rt: rtRows[0],
      rwList: rwRows,
      error: req.flash("error"),
      success: req.flash("success"),
    });
  } catch (error) {
    console.error("Error fetching RT for edit:", error);
    req.flash("error", "Gagal mengambil data RT");
    res.redirect("/rt");
  }
});

// POST /rt/:id/edit - Update RT
router.post("/:id/edit", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const { rw_id, nomor_rt, ketua_rt, keterangan } = req.body;

    // Validasi input
    if (!rw_id || !nomor_rt || !ketua_rt) {
      req.flash("error", "RW, Nomor RT dan Ketua RT harus diisi");
      return res.redirect(`/rt/${req.params.id}/edit`);
    }

    // Cek apakah nomor RT sudah ada (kecuali untuk RT yang sedang diedit)
    const [existing] = await db.execute(
      "SELECT id FROM rt WHERE nomor_rt = ? AND id != ?",
      [nomor_rt, req.params.id]
    );

    if (existing.length > 0) {
      req.flash("error", "Nomor RT sudah terdaftar");
      return res.redirect(`/rt/${req.params.id}/edit`);
    }

    // Update RT
    await db.execute(
      "UPDATE rt SET rw_id = ?, nomor_rt = ?, ketua_rt = ?, keterangan = ? WHERE id = ?",
      [rw_id, nomor_rt, ketua_rt, keterangan || null, req.params.id]
    );

    req.flash("success", `RT ${nomor_rt} berhasil diperbarui`);
    res.redirect("/rt");
  } catch (error) {
    console.error("Error updating RT:", error);
    req.flash("error", "Gagal memperbarui RT");
    res.redirect(`/rt/${req.params.id}/edit`);
  }
});

// POST /rt/:id/delete - Hapus RT
router.post("/:id/delete", async (req, res) => {
  try {
    const db = req.app.locals.db;
    // Cek apakah RT masih memiliki muzakki
    const [muzakki] = await db.execute(
      "SELECT COUNT(*) as count FROM muzakki WHERE rt_id = ?",
      [req.params.id]
    );

    if (muzakki[0].count > 0) {
      req.flash(
        "error",
        "RT tidak dapat dihapus karena masih memiliki data muzakki"
      );
      return res.redirect("/rt");
    }

    // Ambil data RT untuk pesan konfirmasi
    const [rtData] = await db.execute("SELECT nomor_rt FROM rt WHERE id = ?", [
      req.params.id,
    ]);

    if (rtData.length === 0) {
      req.flash("error", "RT tidak ditemukan");
      return res.redirect("/rt");
    }

    // Hapus RT
    await db.execute("DELETE FROM rt WHERE id = ?", [req.params.id]);

    req.flash("success", `RT ${rtData[0].nomor_rt} berhasil dihapus`);
    res.redirect("/rt");
  } catch (error) {
    console.error("Error deleting RT:", error);
    req.flash("error", "Gagal menghapus RT");
    res.redirect("/rt");
  }
});

// GET /rt/:id/detail - Detail RT dengan daftar muzakki
router.get("/:id/detail", async (req, res) => {
  try {
    const db = req.app.locals.db;
    // Ambil data RT
    const [rtData] = await db.execute("SELECT * FROM rt WHERE id = ?", [
      req.params.id,
    ]);

    if (rtData.length === 0) {
      req.flash("error", "RT tidak ditemukan");
      return res.redirect("/rt");
    }

    // Ambil daftar muzakki di RT ini
    const [muzakkiData] = await db.execute(
      `
      SELECT 
        m.*,
        COALESCE(NULLIF(TRIM(m.nama_muzakki), ''), CONCAT('Muzakki #', m.id)) as nama_muzakki_list
      FROM muzakki m 
      WHERE m.rt_id = ? 
      ORDER BY m.id ASC
    `,
      [req.params.id]
    );

    // Ambil daftar mustahik di RT ini
    const [mustahikData] = await db.execute(
      `
      SELECT
        ms.*,
        COALESCE(NULLIF(TRIM(ms.nama), ''), CONCAT('Mustahik #', ms.id)) as nama_mustahik_list
      FROM mustahik ms
      WHERE ms.rt_id = ?
      ORDER BY ms.created_at DESC, ms.id DESC
    `,
      [req.params.id]
    );

    const normalizedMuzakkiData = (muzakkiData || []).map((item) => {
      const jumlahBayar = parseFloat(item.jumlah_bayar) || 0;
      const jumlahUang = parseFloat(item.jumlah_uang) || 0;
      const jumlahBerasKg = parseFloat(item.jumlah_beras_kg) || 0;
      const jumlahZakatRupiah =
        item.jenis_zakat === "uang" ? jumlahUang : jumlahBerasKg * 12000;
      const kembalian = parseFloat(item.kembalian) || 0;
      const status = jumlahBayar >= jumlahZakatRupiah ? "lunas" : "belum_lunas";

      return {
        ...item,
        jumlah_bayar_safe: jumlahBayar,
        jumlah_uang_safe: jumlahUang,
        jumlah_beras_kg_safe: jumlahBerasKg,
        jumlah_zakat_rupiah_safe: jumlahZakatRupiah,
        kembalian_safe: kembalian,
        status,
      };
    });

    // Hitung statistik
    const stats = {
      total_muzakki: normalizedMuzakkiData.length,
      total_mustahik: (mustahikData || []).length,
      total_zakat: normalizedMuzakkiData.reduce(
        (sum, m) => sum + parseFloat(m.jumlah_zakat_rupiah_safe || 0),
        0
      ),
      total_bayar: normalizedMuzakkiData.reduce(
        (sum, m) => sum + parseFloat(m.jumlah_bayar_safe || 0),
        0
      ),
      total_jumlah_beras_kg: normalizedMuzakkiData.reduce(
        (sum, m) => sum + parseFloat(m.jumlah_beras_kg_safe || 0),
        0
      ),
      total_muzakki_uang: normalizedMuzakkiData.filter(
        (m) => String(m.jenis_zakat || "").toLowerCase() === "uang"
      ).length,
      total_muzakki_beras: normalizedMuzakkiData.filter(
        (m) => String(m.jenis_zakat || "").toLowerCase() === "beras"
      ).length,
      lunas: normalizedMuzakkiData.filter((m) => m.status === "lunas").length,
      belum_lunas: normalizedMuzakkiData.filter((m) => m.status === "belum_lunas").length,
      total_kembalian: normalizedMuzakkiData.reduce(
        (sum, m) => sum + parseFloat(m.kembalian_safe || 0),
        0
      ),
    };

    res.render("rt/detail", {
      title: `Detail RT ${rtData[0].nomor_rt} - Zakat Fitrah App`,
      user: req.session.user,
      rt: rtData[0],
      muzakkiList: normalizedMuzakkiData,
      mustahikList: mustahikData || [],
      stats: stats,
      success: req.flash("success"),
      error: req.flash("error"),
    });
  } catch (error) {
    console.error("Error fetching RT detail:", error);
    req.flash("error", "Gagal mengambil detail RT");
    res.redirect("/rt");
  }
});

module.exports = router;
