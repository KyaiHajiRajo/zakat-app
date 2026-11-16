/**
 * Waktu sholat & imsakiyah dihitung offline (tanpa internet / API) memakai rumus astronomi
 * standar (PrayTimes.org) dengan parameter Kemenag RI:
 *   Subuh 20°, Isya 18°, Ashar (Syafi'i), Maghrib/Terbit −0,833°,
 *   Imsak = Subuh − 10 menit, ihtiyati +2 menit (Terbit −2 menit).
 */
const D2R = Math.PI / 180;
const sin = (d) => Math.sin(d * D2R);
const cos = (d) => Math.cos(d * D2R);
const tan = (d) => Math.tan(d * D2R);
const arcsin = (x) => Math.asin(x) / D2R;
const arccos = (x) => Math.acos(x) / D2R;
const arctan2 = (y, x) => Math.atan2(y, x) / D2R;
const arccot = (x) => Math.atan(1 / x) / D2R;
const fixAngle = (a) => ((a % 360) + 360) % 360;
const fixHour = (h) => ((h % 24) + 24) % 24;

function julian(year, month, day) {
  if (month <= 2) {
    year -= 1;
    month += 12;
  }
  const A = Math.floor(year / 100);
  const B = 2 - A + Math.floor(A / 4);
  return Math.floor(365.25 * (year + 4716)) + Math.floor(30.6001 * (month + 1)) + day + B - 1524.5;
}

function sunPosition(jd) {
  const D = jd - 2451545.0;
  const g = fixAngle(357.529 + 0.98560028 * D);
  const q = fixAngle(280.459 + 0.98564736 * D);
  const L = fixAngle(q + 1.915 * sin(g) + 0.02 * sin(2 * g));
  const e = 23.439 - 0.00000036 * D;
  const RA = fixHour(arctan2(cos(e) * sin(L), cos(L)) / 15);
  return { declination: arcsin(sin(e) * sin(L)), equation: q / 15 - RA };
}

function compute(year, month, day, { lat, lng, tz, fajrAngle = 20, ishaAngle = 18 }) {
  const jd = julian(year, month, day) - lng / (15 * 24);
  const midDay = (t) => fixHour(12 - sunPosition(jd + t).equation);
  const angleTime = (angle, t, ccw) => {
    const decl = sunPosition(jd + t).declination;
    const value = (-sin(angle) - sin(decl) * sin(lat)) / (cos(decl) * cos(lat));
    const span = arccos(Math.max(-1, Math.min(1, value))) / 15;
    return midDay(t) + (ccw ? -span : span);
  };
  const asrTime = (t) => {
    const decl = sunPosition(jd + t).declination;
    return angleTime(-arccot(1 + tan(Math.abs(lat - decl))), t, false);
  };

  let t = { fajr: 5 / 24, sunrise: 6 / 24, dhuhr: 12 / 24, asr: 13 / 24, sunset: 18 / 24, isha: 18 / 24 };
  let r;
  for (let i = 0; i < 2; i++) {
    r = {
      fajr: angleTime(fajrAngle, t.fajr, true),
      sunrise: angleTime(0.833, t.sunrise, true),
      dhuhr: midDay(t.dhuhr),
      asr: asrTime(t.asr),
      sunset: angleTime(0.833, t.sunset, false),
      isha: angleTime(ishaAngle, t.isha, false),
    };
    t = Object.fromEntries(Object.entries(r).map(([k, v]) => [k, v / 24]));
  }

  const shift = tz - lng / 15; // dari waktu universal ke zona waktu setempat
  return Object.fromEntries(Object.entries(r).map(([k, v]) => [k, v + shift]));
}

function toMinutes(hours) {
  return Math.round(hours * 60);
}

function format(minutes) {
  const m = ((minutes % 1440) + 1440) % 1440;
  return String(Math.floor(m / 60)).padStart(2, "0") + ":" + String(m % 60).padStart(2, "0");
}

/**
 * @param {string} isoDate  "YYYY-MM-DD"
 * @returns {{imsak,subuh,terbit,dzuhur,ashar,maghrib,isya}} teks "HH:MM" dan `menit` (menit sejak 00:00)
 */
function prayerTimes(isoDate, { lat, lng, tz, ihtiyati = 2 }) {
  const [y, m, d] = isoDate.split("-").map(Number);
  const raw = compute(y, m, d, { lat, lng, tz });
  const minutes = {
    imsak: toMinutes(raw.fajr) + ihtiyati - 10,
    subuh: toMinutes(raw.fajr) + ihtiyati,
    terbit: toMinutes(raw.sunrise) - ihtiyati,
    dzuhur: toMinutes(raw.dhuhr) + ihtiyati,
    ashar: toMinutes(raw.asr) + ihtiyati,
    maghrib: toMinutes(raw.sunset) + ihtiyati,
    isya: toMinutes(raw.isha) + ihtiyati,
  };
  const text = {};
  for (const key of Object.keys(minutes)) {
    text[key] = format(minutes[key]);
  }
  return { text, minutes };
}

module.exports = { prayerTimes };
