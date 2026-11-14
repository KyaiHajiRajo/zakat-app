// Helper untuk form yang bisa dikirim lewat modal (fetch) maupun form biasa.

function wantsJson(req) {
  return req.get("X-Requested-With") === "fetch";
}

// Hanya izinkan path internal agar tidak jadi open redirect.
function safeReturn(value, fallback) {
  if (
    typeof value === "string" &&
    value.startsWith("/") &&
    !value.startsWith("//") &&
    !value.includes("\\")
  ) {
    return value;
  }
  return fallback;
}

function reply(req, res, { ok, message, redirect, status }) {
  if (wantsJson(req)) {
    if (ok) {
      req.flash("success", message);
    }
    return res
      .status(ok ? 200 : status || 400)
      .json({ ok, message, redirect });
  }

  req.flash(ok ? "success" : "error", message);
  return res.redirect(redirect);
}

module.exports = { wantsJson, safeReturn, reply };
