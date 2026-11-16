/**
 * Pembatas percobaan sederhana di memori (tanpa paket tambahan, ringan untuk laptop).
 * Hanya kegagalan yang dihitung; `reset` dipanggil saat berhasil.
 */
function createLimiter({ windowMs, max }) {
  const entries = new Map();

  const sweeper = setInterval(() => {
    const now = Date.now();
    for (const [key, entry] of entries) {
      if (entry.resetAt <= now) {
        entries.delete(key);
      }
    }
  }, 60 * 1000);
  sweeper.unref();

  function current(key) {
    const entry = entries.get(key);
    if (!entry || entry.resetAt <= Date.now()) {
      entries.delete(key);
      return null;
    }
    return entry;
  }

  return {
    // sisa detik penguncian, atau 0 jika boleh mencoba
    blockedFor(key) {
      const entry = current(key);
      return entry && entry.count >= max ? Math.ceil((entry.resetAt - Date.now()) / 1000) : 0;
    },
    fail(key) {
      const entry = current(key) || { count: 0, resetAt: Date.now() + windowMs };
      entry.count += 1;
      entries.set(key, entry);
    },
    reset(key) {
      entries.delete(key);
    },
  };
}

module.exports = { createLimiter };
