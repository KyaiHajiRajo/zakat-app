/**
 * Membuat stylesheet mode gelap per halaman dari CSS aslinya.
 *
 * - Membaca <style> dan atribut style="" tiap view (+ partial yang di-include).
 * - Latar/teks/garis terang dibalik kecerahannya (hue tetap) jadi versi gelap;
 *   warna aksen (hijau, biru, dsb. yang sudah cukup kontras) dibiarkan.
 * - Hasil: public/css/dark/<view>.css dan manifest.json, dimuat otomatis oleh header
 *   (lihat middleware di app.js). Aturan dark-mode yang ditulis manual di halaman
 *   tetap menang karena spesifisitasnya sama dan dimuat belakangan.
 *
 * Jalankan ulang setelah mengubah CSS halaman:  npm run build-dark
 */
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const postcss = require("postcss");

const ROOT = path.join(__dirname, "..");
const VIEWS = path.join(ROOT, "views");
const OUT = path.join(ROOT, "public", "css", "dark");

// Modul masjid (kas, agenda, inventaris, info, ramadan) sudah berbasis variabel tema sehingga mode gelap otomatis.
const SKIP_DIRS = ["layouts", "auth", "kas", "agenda", "inventaris", "info", "ramadan", "masjid", "donasi", "transparansi"];
const SKIP_FILES = ["home.ejs", "error.ejs"];

/* ---------- warna ---------- */
const clamp = (v, a, b) => Math.min(b, Math.max(a, v));

function parseColor(token) {
  let m;
  token = token.trim().toLowerCase();
  if (token === "white") return { r: 255, g: 255, b: 255, a: 1 };
  if (token === "black") return { r: 0, g: 0, b: 0, a: 1 };
  if ((m = token.match(/^#([0-9a-f]{3,8})$/))) {
    let h = m[1];
    if (h.length === 3 || h.length === 4) h = h.split("").map((c) => c + c).join("");
    if (h.length !== 6 && h.length !== 8) return null;
    return {
      r: parseInt(h.slice(0, 2), 16),
      g: parseInt(h.slice(2, 4), 16),
      b: parseInt(h.slice(4, 6), 16),
      a: h.length === 8 ? parseInt(h.slice(6, 8), 16) / 255 : 1,
    };
  }
  if ((m = token.match(/^rgba?\(([^)]+)\)$/))) {
    const p = m[1].split(/[\s,\/]+/).filter(Boolean);
    if (p.length < 3) return null;
    const n = (v) => (v.endsWith("%") ? (parseFloat(v) / 100) * 255 : parseFloat(v));
    const a = p[3] === undefined ? 1 : p[3].endsWith("%") ? parseFloat(p[3]) / 100 : parseFloat(p[3]);
    const c = { r: n(p[0]), g: n(p[1]), b: n(p[2]), a };
    return [c.r, c.g, c.b, c.a].some(Number.isNaN) ? null : c;
  }
  if ((m = token.match(/^hsla?\(([^)]+)\)$/))) {
    const p = m[1].split(/[\s,\/]+/).filter(Boolean);
    if (p.length < 3) return null;
    const h = parseFloat(p[0]), s = parseFloat(p[1]) / 100, l = parseFloat(p[2]) / 100;
    const a = p[3] === undefined ? 1 : p[3].endsWith("%") ? parseFloat(p[3]) / 100 : parseFloat(p[3]);
    const rgb = hslToRgb({ h, s, l });
    return Number.isNaN(h + s + l + a) ? null : { ...rgb, a };
  }
  return null;
}

function rgbToHsl({ r, g, b }) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0, s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h *= 60;
  }
  return { h, s, l };
}

function hslToRgb({ h, s, l }) {
  h = ((h % 360) + 360) % 360;
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  let r = 0, g = 0, b = 0;
  if (h < 60) [r, g, b] = [c, x, 0];
  else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x];
  else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];
  return { r: (r + m) * 255, g: (g + m) * 255, b: (b + m) * 255 };
}

function luminance({ r, g, b }) {
  const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
  return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b);
}

function contrast(a, b) {
  const l1 = luminance(a), l2 = luminance(b);
  return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
}

function format(c) {
  const to = (v) => Math.round(clamp(v, 0, 255));
  if (c.a !== undefined && c.a < 1) return `rgba(${to(c.r)}, ${to(c.g)}, ${to(c.b)}, ${Math.round(c.a * 100) / 100})`;
  return "#" + [c.r, c.g, c.b].map((v) => to(v).toString(16).padStart(2, "0")).join("");
}

const SURFACE = { r: 15, g: 23, b: 42 }; // #0f172a, sama dengan --surface mode gelap
const SLATE = { h: 222, s: 0.47 };

/* ---------- pemetaan per peran ---------- */
function mapBackground(c) {
  if (c.a < 0.35) return c;
  const hsl = rgbToHsl(c);
  if (c.a < 0.5 && hsl.l > 0.85) return c; // lapisan putih tipis di atas warna (kaca)
  if (hsl.l < 0.72) return c;
  const neutral = hsl.s < 0.15;
  const l2 = neutral ? 0.11 + (1 - hsl.l) * 0.5 : 0.12 + (1 - hsl.l) * 0.6;
  const rgb = hslToRgb({
    h: neutral ? SLATE.h : hsl.h,
    s: neutral ? SLATE.s : clamp(hsl.s * 0.55, 0.25, 0.7),
    l: l2,
  });
  return { ...rgb, a: c.a };
}

function mapBorder(c) {
  if (c.a < 0.35) return c;
  const hsl = rgbToHsl(c);
  if (c.a < 0.9 && hsl.l > 0.85) return c;
  if (hsl.l < 0.7) return c;
  const neutral = hsl.s < 0.15;
  const rgb = hslToRgb({
    h: neutral ? SLATE.h : hsl.h,
    s: neutral ? 0.3 : clamp(hsl.s * 0.5, 0.2, 0.6),
    l: 0.14 + (1 - hsl.l) * 0.55,
  });
  return { ...rgb, a: c.a };
}

function mapForeground(c, bgRgb) {
  if (c.a < 0.4) return c;
  const against = bgRgb || SURFACE;
  const hsl = rgbToHsl(c);
  if (bgRgb && luminance(bgRgb) > 0.2) {
    // latar tetap terang/menengah: pertahankan jika cukup kontras, selain itu pilih terbaik
    if (contrast(c, against) >= 3) return c;
    const dark = { r: 17, g: 24, b: 39 }, light = { r: 255, g: 255, b: 255 };
    return { ...(contrast(dark, against) >= contrast(light, against) ? dark : light), a: c.a };
  }
  if (hsl.l < 0.3) {
    const l2 = clamp(1 - hsl.l, 0.74, 0.95);
    return { ...hslToRgb({ h: hsl.h, s: hsl.s < 0.15 ? 0.25 : clamp(hsl.s, 0.3, 0.9), l: l2 }), a: c.a };
  }
  if (contrast(c, against) >= 4.5) return c;
  let lo = hsl.l, hi = 0.97;
  for (let i = 0; i < 14; i++) {
    const mid = (lo + hi) / 2;
    const test = hslToRgb({ h: hsl.h, s: hsl.s, l: mid });
    if (contrast(test, against) >= 4.8) hi = mid; else lo = mid;
  }
  return { ...hslToRgb({ h: hsl.h, s: hsl.s, l: hi }), a: c.a };
}

const COLOR_RE = /#[0-9a-fA-F]{3,8}\b|rgba?\([^)]*\)|hsla?\([^)]*\)|\b(?:white|black)\b/g;

function mapValue(value, mapper) {
  if (/var\(|url\(/.test(value) && !/gradient/.test(value)) return value;
  return value.replace(COLOR_RE, (tok) => {
    const c = parseColor(tok);
    if (!c) return tok;
    const out = mapper(c);
    return out === c ? tok : format(out);
  });
}

const FG_PROPS = new Set(["color", "fill", "stroke", "caret-color", "text-decoration-color", "-webkit-text-fill-color"]);
const BG_PROPS = new Set(["background", "background-color", "background-image"]);
const isBorderProp = (p) => /^(border|outline)(-(top|right|bottom|left))?(-color)?$/.test(p) || p === "column-rule-color";

/** Hitung nilai gelap untuk sekumpulan deklarasi {prop, value} yang berada dalam satu aturan. */
function darkenDeclarations(decls) {
  const out = [];
  let newBg = null;
  let changed = false;

  // 1) latar dahulu agar teks bisa menyesuaikan
  for (const d of decls) {
    if (!BG_PROPS.has(d.prop)) continue;
    const value = mapValue(d.value, mapBackground);
    // warna latar efektif: warna solid terakhir / rata-rata stop gradien
    const toks = (value.match(COLOR_RE) || []).map(parseColor).filter((c) => c && c.a > 0.5);
    if (toks.length) {
      newBg = {
        r: toks.reduce((s, c) => s + c.r, 0) / toks.length,
        g: toks.reduce((s, c) => s + c.g, 0) / toks.length,
        b: toks.reduce((s, c) => s + c.b, 0) / toks.length,
      };
    }
    if (value !== d.value) changed = true;
    out.push({ ...d, value });
  }

  for (const d of decls) {
    let value = null;
    if (FG_PROPS.has(d.prop)) value = mapValue(d.value, (c) => mapForeground(c, newBg));
    else if (isBorderProp(d.prop)) value = mapValue(d.value, mapBorder);
    if (value === null) continue;
    if (value !== d.value) changed = true;
    out.push({ ...d, value });
  }

  // Aturan tanpa perubahan tetap dikeluarkan ("pin") hanya jika ada perubahan lain di aturan yang sama
  // atau ia menyetel warna; ini menjaga urutan prioritas antar aturan halaman.
  return out.length && (changed || out.length) ? out : [];
}

/* ---------- membaca view ---------- */
function resolveIncludes(file, seen = new Set()) {
  if (seen.has(file) || !fs.existsSync(file)) return "";
  seen.add(file);
  const text = fs.readFileSync(file, "utf8");
  let all = text;
  const re = /include\(\s*['"]([^'"]+)['"]/g;
  let m;
  while ((m = re.exec(text))) {
    if (/layouts\//.test(m[1])) continue;
    let target = path.resolve(path.dirname(file), m[1]);
    if (!target.endsWith(".ejs")) target += ".ejs";
    all += "\n" + resolveIncludes(target, seen);
  }
  return all;
}

function listPages(dir = VIEWS, acc = []) {
  for (const name of fs.readdirSync(dir)) {
    const full = path.join(dir, name);
    const rel = path.relative(VIEWS, full);
    if (fs.statSync(full).isDirectory()) {
      if (SKIP_DIRS.includes(name) || name === "partials") continue;
      listPages(full, acc);
    } else if (name.endsWith(".ejs") && !SKIP_FILES.includes(name) && !/backup/.test(name)) {
      acc.push(rel);
    }
  }
  return acc;
}

function cssBlocks(source) {
  const blocks = [];
  const re = /<style[^>]*>([\s\S]*?)<\/style>/gi;
  let m;
  while ((m = re.exec(source))) blocks.push(m[1].replace(/<%[\s\S]*?%>/g, ""));
  return blocks;
}

function inlineStyles(source) {
  const found = new Set();
  const re = /\sstyle="([^"<>]*)"/g;
  let m;
  while ((m = re.exec(source))) if (/(background|color|border)/i.test(m[1])) found.add(m[1]);
  return [...found];
}

function prefixSelector(sel) {
  sel = sel.trim();
  if (!sel || /dark-mode/.test(sel)) return null;
  if (/^(html|:root)\b/.test(sel)) {
    if (sel === "html" || sel === ":root") return "html.dark-mode";
    return null;
  }
  return "html.dark-mode " + sel;
}

function buildForPage(rel) {
  const source = resolveIncludes(path.join(VIEWS, rel));
  const out = postcss.root();
  let ruleCount = 0;

  for (const css of cssBlocks(source)) {
    let root;
    try { root = postcss.parse(css); } catch (e) { console.warn(`  ! ${rel}: CSS tidak bisa di-parse (${e.message.split("\n")[0]})`); continue; }

    root.walkRules((rule) => {
      for (let p = rule.parent; p; p = p.parent) {
        if (p.type === "atrule" && /keyframes|font-face|page/i.test(p.name)) return;
        if (p.type === "atrule" && /prefers-color-scheme\s*:\s*dark/i.test(p.params || "")) return;
      }
      const decls = [];
      rule.walkDecls((d) => decls.push({ prop: d.prop.toLowerCase(), value: d.value, important: d.important }));
      const darkened = darkenDeclarations(decls);
      if (!darkened.length) return;

      const selectors = rule.selectors.map(prefixSelector).filter(Boolean);
      if (!selectors.length) return;

      const clone = postcss.rule({ selectors });
      darkened.forEach((d) => clone.append(postcss.decl({ prop: d.prop, value: d.value, important: d.important })));

      // bungkus dengan at-rule induk (@media dll.)
      let node = clone;
      for (let p = rule.parent; p && p.type !== "root"; p = p.parent) {
        if (p.type === "atrule") node = postcss.atRule({ name: p.name, params: p.params }).append(node);
      }
      out.append(node);
      ruleCount++;
    });
  }

  for (const style of inlineStyles(source)) {
    const decls = style.split(";").map((s) => s.trim()).filter(Boolean).map((s) => {
      const i = s.indexOf(":");
      return i < 0 ? null : { prop: s.slice(0, i).trim().toLowerCase(), value: s.slice(i + 1).trim() };
    }).filter(Boolean);
    const darkened = darkenDeclarations(decls);
    if (!darkened.length) continue;
    const attr = style.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
    const clone = postcss.rule({ selector: `html.dark-mode [style="${attr}"]` });
    darkened.forEach((d) => clone.append(postcss.decl({ prop: d.prop, value: d.value, important: true })));
    out.append(clone);
    ruleCount++;
  }

  return { css: out.toString(), ruleCount };
}

/* ---------- main ---------- */
function main() {
  fs.mkdirSync(OUT, { recursive: true });
  for (const f of fs.readdirSync(OUT)) fs.unlinkSync(path.join(OUT, f));

  const manifest = {};
  let total = 0;
  for (const rel of listPages()) {
    const key = rel.replace(/\.ejs$/, "").split(path.sep).join("__");
    const { css, ruleCount } = buildForPage(rel);
    if (!ruleCount) continue;
    const body = `/* Dibuat otomatis oleh scripts/build-dark-css.js dari views/${rel} */\n${css}\n`;
    fs.writeFileSync(path.join(OUT, key + ".css"), body);
    manifest[key] = crypto.createHash("md5").update(body).digest("hex").slice(0, 8);
    total += ruleCount;
    console.log(`  ${rel.padEnd(34)} ${String(ruleCount).padStart(4)} aturan  ${(Buffer.byteLength(body) / 1024).toFixed(1)} KB`);
  }
  fs.writeFileSync(path.join(OUT, "manifest.json"), JSON.stringify(manifest, null, 2));
  console.log(`Selesai: ${Object.keys(manifest).length} halaman, ${total} aturan.`);
}

main();
