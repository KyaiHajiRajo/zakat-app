const express = require("express");
const router = express.Router();

// GET /infak - List all infak
router.get("/", async (req, res) => {
  try {
    const db = req.app.locals.db;

    // Get all infak with muzakki info
    const [infak] = await db.execute(`
            SELECT 
                i.id,
                i.muzakki_id,
                COALESCE(i.jumlah, 0) as jumlah,
                i.keterangan,
                i.created_at,
                i.updated_at,
                COALESCE(NULLIF(TRIM(m.nama_muzakki), ''), CONCAT('Muzakki #', m.id)) as muzakki_name,
                r.nomor_rt,
                m.jumlah_jiwa
            FROM infak i
            LEFT JOIN muzakki m ON i.muzakki_id = m.id
            LEFT JOIN rt r ON m.rt_id = r.id
            ORDER BY i.created_at DESC
        `);

    const [availableMuzakki] = await db.execute(`
            SELECT
                m.id,
                COALESCE(NULLIF(TRIM(m.nama_muzakki), ''), CONCAT('Muzakki #', m.id)) as nama,
                r.nomor_rt
            FROM muzakki m
            LEFT JOIN infak i ON i.muzakki_id = m.id
            LEFT JOIN rt r ON m.rt_id = r.id
            WHERE i.id IS NULL
            ORDER BY
                CASE WHEN r.nomor_rt IS NULL THEN 1 ELSE 0 END,
                r.nomor_rt ASC,
                nama ASC
        `);

    res.render("infak/index", {
      title: "Data Infak - Zakat Fitrah",
      layout: "layouts/main",
      infak,
      availableMuzakki: availableMuzakki || [],
    });
  } catch (error) {
    console.error("Error fetching infak:", error);
    req.flash("error_msg", "Terjadi kesalahan saat mengambil data infak");
    res.redirect("/");
  }
});

// POST /infak/create - Simpan infak baru
router.post("/create", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const incomingIds = Array.isArray(req.body.muzakki_id)
      ? req.body.muzakki_id
      : req.body.muzakki_id
      ? [req.body.muzakki_id]
      : [];
    const incomingJumlah = Array.isArray(req.body.jumlah)
      ? req.body.jumlah
      : req.body.jumlah
      ? [req.body.jumlah]
      : [];
    const incomingNotes = Array.isArray(req.body.keterangan)
      ? req.body.keterangan
      : req.body.keterangan
      ? [req.body.keterangan]
      : [];

    const entries = incomingIds
      .map((muzakkiId, index) => {
        const parsedId = parseInt(muzakkiId, 10);
        const parsedJumlah = parseFloat(incomingJumlah[index] || 0);
        const note = (incomingNotes[index] || "").trim();

        if (!parsedId || Number.isNaN(parsedJumlah) || parsedJumlah <= 0) {
          return null;
        }

        return {
          muzakki_id: parsedId,
          jumlah: Math.round(parsedJumlah * 100) / 100,
          keterangan: note || null,
        };
      })
      .filter(Boolean);

    if (!entries.length) {
      req.flash("error_msg", "Tidak ada data infak yang valid untuk disimpan");
      return res.redirect("/infak");
    }

    let successCount = 0;
    const errors = [];

    for (const entry of entries) {
      const [existing] = await db.execute(
        "SELECT id FROM infak WHERE muzakki_id = ?",
        [entry.muzakki_id]
      );

      if (existing.length > 0) {
        errors.push(`Muzakki ID ${entry.muzakki_id} sudah tercatat`);
        continue;
      }

      await db.execute(
        `INSERT INTO infak (muzakki_id, jumlah, keterangan) VALUES (?, ?, ?)`,
        [entry.muzakki_id, entry.jumlah, entry.keterangan]
      );

      successCount += 1;
    }

    if (successCount > 0) {
      req.flash("success_msg", `Berhasil menyimpan ${successCount} infak baru`);
    }

    if (errors.length) {
      req.flash("error_msg", errors.join("; "));
    }

    res.redirect("/infak");
  } catch (error) {
    console.error("Error creating infak:", error);
    req.flash("error_msg", "Terjadi kesalahan saat menyimpan data infak");
    res.redirect("/infak");
  }
});

// DELETE /infak/:id - Hapus data infak
router.delete("/:id", async (req, res) => {
  try {
    const db = req.app.locals.db;
    const infakId = parseInt(req.params.id, 10);

    if (!infakId) {
      req.flash("error_msg", "ID infak tidak valid");
      return res.redirect("/infak");
    }

    const [existingInfak] = await db.execute(
      `
        SELECT i.id, i.muzakki_id,
          COALESCE(NULLIF(TRIM(m.nama_muzakki), ''), CONCAT('Muzakki #', m.id)) as muzakki_name
        FROM infak i
        LEFT JOIN muzakki m ON i.muzakki_id = m.id
        WHERE i.id = ?
      `,
      [infakId]
    );

    if (!existingInfak.length) {
      req.flash("error_msg", "Data infak tidak ditemukan");
      return res.redirect("/infak");
    }

    await db.execute("DELETE FROM infak WHERE id = ?", [infakId]);

    req.flash(
      "success_msg",
      `Data infak untuk ${existingInfak[0].muzakki_name || `ID ${infakId}`} berhasil dihapus`
    );
    res.redirect("/infak");
  } catch (error) {
    console.error("Error deleting infak:", error);
    req.flash("error_msg", "Terjadi kesalahan saat menghapus data infak");
    res.redirect("/infak");
  }
});

module.exports = router;
