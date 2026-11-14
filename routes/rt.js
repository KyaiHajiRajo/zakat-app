const express = require("express");
const { reply, safeReturn } = require("../utils/respond");
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

// GET /rt/form-data?id= - Data untuk modal tambah/ubah RT (JSON)
router.get("/form-data", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const [rwList] = await db.execute(
      "SELECT id, nomor_rw, ketua_rw FROM rw ORDER BY nomor_rw ASC"
    );

    let rt = null;
    if (req.query.id) {
      const [rows] = await db.execute(
        "SELECT id, rw_id, nomor_rt, ketua_rt, keterangan FROM rt WHERE id = ?",
        [req.query.id]
      );
      if (rows.length === 0) {
        return res.status(404).json({ ok: false, message: "RT tidak ditemukan" });
      }
      rt = rows[0];
    }

    res.json({ ok: true, rwList, rt });
  } catch (error) {
    console.error("Error loading RT form data:", error);
    res.status(500).json({ ok: false, message: "Gagal memuat data form" });
  }
});

// GET /rt/create - Form tambah RT kini berupa modal di halaman daftar
router.get("/create", (req, res) => {
  const rw = /^\d+$/.test(req.query.rw || "") ? `&rw=${req.query.rw}` : "";
  res.redirect(`/rt?new=1${rw}`);
});

// Validasi bersama untuk tambah/ubah RT
async function validateRt(db, body, excludeId) {
  const rwId = String(body.rw_id || "").trim();
  const nomorRt = String(body.nomor_rt || "").trim();
  const ketuaRt = String(body.ketua_rt || "").trim();
  const keterangan = String(body.keterangan || "").trim();

  if (!rwId || !nomorRt || !ketuaRt) {
    return { error: "RW, Nomor RT dan Ketua RT harus diisi" };
  }
  if (nomorRt.length > 10) return { error: "Nomor RT maksimal 10 karakter" };
  if (ketuaRt.length > 100) return { error: "Nama ketua RT maksimal 100 karakter" };
  if (keterangan.length > 100) return { error: "Keterangan maksimal 100 karakter" };

  const [rw] = await db.execute("SELECT id FROM rw WHERE id = ?", [rwId]);
  if (rw.length === 0) return { error: "RW yang dipilih tidak ditemukan" };

  const [existing] = await db.execute(
    "SELECT id FROM rt WHERE nomor_rt = ? AND id != ?",
    [nomorRt, excludeId || 0]
  );
  if (existing.length > 0) return { error: "Nomor RT sudah terdaftar" };

  return { rwId, nomorRt, ketuaRt, keterangan: keterangan || null };
}

// POST /rt/create - Simpan RT baru
router.post("/create", async (req, res) => {
  const back = safeReturn(req.body.return_to, "/rt");
  try {
    const db = req.app.locals.db;
    const data = await validateRt(db, req.body);
    if (data.error) {
      return reply(req, res, { ok: false, message: data.error, redirect: back });
    }

    await db.execute(
      "INSERT INTO rt (rw_id, nomor_rt, ketua_rt, keterangan) VALUES (?, ?, ?, ?)",
      [data.rwId, data.nomorRt, data.ketuaRt, data.keterangan]
    );

    reply(req, res, {
      ok: true,
      message: `${data.nomorRt} berhasil ditambahkan`,
      redirect: back,
    });
  } catch (error) {
    console.error("Error creating RT:", error);
    reply(req, res, {
      ok: false,
      message: "Gagal menambahkan RT",
      redirect: back,
      status: 500,
    });
  }
});

// GET /rt/:id/edit - Form edit RT kini berupa modal di halaman daftar
router.get("/:id/edit", (req, res) => {
  if (!/^\d+$/.test(req.params.id)) return res.redirect("/rt");
  res.redirect(`/rt?edit=${req.params.id}`);
});

// POST /rt/:id/edit - Update RT
router.post("/:id/edit", async (req, res) => {
  const back = safeReturn(req.body.return_to, "/rt");
  try {
    const db = req.app.locals.db;
    const [current] = await db.execute("SELECT id FROM rt WHERE id = ?", [
      req.params.id,
    ]);
    if (current.length === 0) {
      return reply(req, res, {
        ok: false,
        message: "RT tidak ditemukan",
        redirect: back,
        status: 404,
      });
    }

    const data = await validateRt(db, req.body, req.params.id);
    if (data.error) {
      return reply(req, res, { ok: false, message: data.error, redirect: back });
    }

    await db.execute(
      "UPDATE rt SET rw_id = ?, nomor_rt = ?, ketua_rt = ?, keterangan = ? WHERE id = ?",
      [data.rwId, data.nomorRt, data.ketuaRt, data.keterangan, req.params.id]
    );

    reply(req, res, {
      ok: true,
      message: `${data.nomorRt} berhasil diperbarui`,
      redirect: back,
    });
  } catch (error) {
    console.error("Error updating RT:", error);
    reply(req, res, {
      ok: false,
      message: "Gagal memperbarui RT",
      redirect: back,
      status: 500,
    });
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
