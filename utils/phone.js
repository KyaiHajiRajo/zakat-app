// Nomor WhatsApp Indonesia: terima 0812…, +62 812…, 62812… lalu ubah ke format internasional tanpa tanda (62812…)
function normalizeWa(input) {
  let digits = String(input || "").replace(/[^\d+]/g, "");
  if (digits.startsWith("+")) digits = digits.slice(1);
  digits = digits.replace(/\D/g, "");
  if (digits.startsWith("0")) digits = "62" + digits.slice(1);
  else if (digits.startsWith("8")) digits = "62" + digits;
  if (!/^628\d{7,12}$/.test(digits)) return null;
  return digits;
}

// 6281234567890 -> 0812-3456-7890 (untuk ditampilkan)
function displayWa(normalized) {
  if (!normalized) return "";
  const local = "0" + String(normalized).slice(2);
  return local.replace(/^(\d{4})(\d{4})(\d+)$/, "$1-$2-$3");
}

module.exports = { normalizeWa, displayWa };
