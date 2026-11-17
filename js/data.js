/** DTspace 하자 엑셀 → 내부 레코드 정규화, 조회 문법, 필터·그룹 */

/** 엑셀 헤더 → 내부 키 (DTspace 하자관리 Export Excel 36개 열 중 사용하는 열) */
export const COLS = {
  "순번": "id", "단계": "stage", "접수일": "rcv", "동": "dong", "호수": "ho", "평형": "py",
  "위치": "loc", "부위": "part", "대공종": "big", "공종": "trade", "세부공종": "sub", "유형": "type",
  "시공업체명": "vendor", "처리상태": "dts", "처리경과일": "days", "하자내용": "text", "이미지여부": "img",
  "작업완료일": "wdone", "최종완료일": "fdone", "재접수일": "rercv", "재접수여부": "rerc",
  "E/V호기": "ev", "AS예약일시": "as", "참고": "note", "입주/키불": "movein", "입주예약일": "moveinDate",
};
export const REQUIRED = ["순번", "동", "호수", "유형", "하자내용"];
const DATE_KEYS = new Set(["rcv", "wdone", "fdone", "rercv", "moveinDate"]);
export const SRC_KEYS = Object.values(COLS).filter((k) => k !== "id");

function str(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\r/g, "").trim();
}
export function fmtDate(v) {
  const s = str(v);
  if (!s) return "";
  if (/^\d{8}$/.test(s)) return `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`;
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    // 엑셀 일련번호
    const d = new Date(Math.round((Number(s) - 25569) * 86400000));
    return d.toISOString().slice(0, 10);
  }
  return s.replace(/\./g, "-").slice(0, 16);
}
export function normDong(v) {
  const s = str(v).replace(/동$/, "");
  const n = parseInt(s, 10);
  if (!Number.isFinite(n)) return s;
  return String(n < 100 ? 300 + n : n);
}
/**
 * 호수 정규화 → 4자리 문자열. "901"·"0901"·901·"901호"·"301-901"·"301동 0901호"·"3010901" → "0901".
 * 엑셀 셀이 숫자(901)로 오든 문자("0901")로 오든 같은 세대가 되게 한다.
 */
export function normHo(v) {
  const s = str(v);
  if (!s) return "";
  const groups = s.match(/\d+/g);
  if (!groups) return s;
  let t = groups[groups.length - 1];
  if (t.length >= 6 && groups.length === 1) t = t.slice(-4); // 동+호 붙은 값(3010901)
  if (!/^\d{1,4}$/.test(t)) return s;
  return t.padStart(4, "0");
}
/** 호수 → 층(숫자). "0901" → 9, "1203" → 12. 숫자가 아니면 NaN */
export function floorOfHo(ho) {
  const n = Number(ho);
  return /^\d{3,4}$/.test(String(ho)) && n >= 100 ? Math.floor(n / 100) : NaN;
}

/** 헤더행과 데이터행(배열) → 레코드 배열 */
export function rowsToRecords(header, rows) {
  const idx = {};
  header.forEach((h, i) => {
    const k = COLS[str(h)];
    if (k && idx[k] === undefined) idx[k] = i;
  });
  const missing = REQUIRED.filter((h) => idx[COLS[h]] === undefined);
  if (missing.length) throw new Error(`DTspace 하자 엑셀 형식이 아닙니다. 없는 열: ${missing.join(", ")}`);
  const out = [];
  for (const r of rows) {
    if (!r) continue;
    const id = Number(str(r[idx.id]));
    if (!Number.isFinite(id) || id <= 0) continue;
    const rec = { id };
    for (const [k, i] of Object.entries(idx)) {
      if (k === "id") continue;
      let v = r[i];
      if (k === "dong") v = normDong(v);
      else if (k === "ho") v = normHo(v);
      else if (k === "days") v = v === null || v === "" || v === undefined ? "" : Number(v);
      else if (DATE_KEYS.has(k)) v = fmtDate(v);
      else v = str(v);
      if (v !== "") rec[k] = v;
    }
    out.push(rec);
  }
  out.keys = Object.keys(idx);
  return out;
}

/** 원본 비교(처리경과일은 매일 바뀌므로 제외). keys: 이번 엑셀에 있는 열만 비교 */
export function sameSource(a, b, keys = SRC_KEYS) {
  for (const k of keys) if (k !== "days" && k !== "id" && (a[k] ?? "") !== (b[k] ?? "")) return false;
  return true;
}

/* ---------- 판정 ---------- */
export const MI_RE = /미시공|미설치/;
export const isMi = (r) => r.type === "미시공" || MI_RE.test(r.text || "");
export const isPa = (r) => r.type === "파손" || /파손/.test(r.text || "");

export const PRESETS = {
  mipa: { label: "미시공·파손", test: (r) => isMi(r) || isPa(r) },
  mi: { label: "미시공", test: isMi },
  pa: { label: "파손", test: isPa },
  all: { label: "전체", test: () => true },
};

export const LOCAL_STATUS = ["미처리", "작업중", "완료", "재접수"];

/** 내 상태: 내가 지정한 값 우선, 없으면 DTspace 처리상태로 추정 */
export function statusOf(r, edit) {
  if (edit?.status) return edit.status;
  const d = r.dts || "";
  if (/최종완료|작업완료/.test(d)) return "완료";
  if (/재접수/.test(d)) return "재접수";
  return "미처리";
}

/* ---------- 조회 문법: "9동 1203", "309-1203", "12 203", "309동", "1203" ---------- */
export function parseQuery(q) {
  const s = String(q || "").trim().replace(/\s+/g, " ");
  if (!s) return null;
  let m = s.match(/^(\d{1,3})\s*동?\s*[-\s./]?\s*(\d{3,4})\s*호?$/);
  if (m && /동|[-\s./]/.test(s)) return { dong: normDong(m[1]), ho: normHo(m[2]) };
  m = s.match(/^(\d{1,3})\s*동$/);
  if (m) return { dong: normDong(m[1]) };
  m = s.match(/^(\d{3,4})\s*호$/);
  if (m) return { ho: normHo(m[1]) };
  m = s.match(/^(\d{5,7})$/); // 3091203 / 91203
  if (m) {
    const t = m[1];
    const ho = t.slice(-4), d = t.slice(0, -4);
    if (+normDong(d) >= 301 && +normDong(d) <= 399) return { dong: normDong(d), ho: normHo(ho) };
  }
  m = s.match(/^(\d{1,3})$/);
  if (m && (+m[1] <= 12 || (+m[1] >= 301 && +m[1] <= 312))) return { dong: normDong(m[1]) };
  m = s.match(/^(\d{3,4})$/);
  if (m) return { ho: normHo(m[1]) };
  return { text: s };
}

export function shortDong(d) {
  const n = +d;
  return n > 300 && n < 400 ? `${n - 300}동` : `${d}동`;
}
export function hoLabel(ho) {
  return `${String(ho || "").replace(/^0+/, "")}호`;
}

/* ---------- 필터·그룹 ---------- */
/**
 * @param {object[]} rows
 * @param {object} f {dong, ho, preset, type, dts, st, trade, vendor, text}
 * @param {(r)=>string} st 상태 함수
 */
export function applyFilter(rows, f, st) {
  const pre = PRESETS[f.preset || "all"]?.test || (() => true);
  const text = (f.text || "").trim();
  return rows.filter((r) =>
    (!f.dong || r.dong === f.dong) &&
    (!f.ho || r.ho === f.ho) &&
    pre(r) &&
    (!f.type || r.type === f.type) &&
    (!f.dts || r.dts === f.dts) &&
    (!f.trade || r.trade === f.trade) &&
    (!f.vendor || r.vendor === f.vendor) &&
    (!f.st || (f.st === "open" ? st(r) !== "완료" : st(r) === f.st)) &&
    (!text || [r.text, r.loc, r.part, r.sub, r.trade, r.vendor, r.type].some((v) => (v || "").includes(text)))
  );
}

/** 공종 / 업체 그룹 → [{key, trade, vendor, rows}] 건수 내림차순 */
export function groupByTradeVendor(rows) {
  const m = new Map();
  for (const r of rows) {
    const key = `${r.trade || "(공종없음)"}|${r.vendor || "(업체없음)"}`;
    if (!m.has(key)) m.set(key, { key, trade: r.trade || "(공종없음)", vendor: r.vendor || "(업체없음)", rows: [] });
    m.get(key).rows.push(r);
  }
  return [...m.values()].sort((a, b) => b.rows.length - a.rows.length || a.trade.localeCompare(b.trade, "ko"));
}

export function countBy(rows, fn) {
  const m = new Map();
  for (const r of rows) {
    const k = fn(r) || "(없음)";
    m.set(k, (m.get(k) || 0) + 1);
  }
  return m;
}

export const sortUnit = (a, b) => (a.dong === b.dong ? (a.ho || "").localeCompare(b.ho || "") : (a.dong || "").localeCompare(b.dong || "")) || a.id - b.id;

/** 저장된 레코드의 동·호수 정규화(이전 버전 데이터 이전용). 바뀐 레코드만 반환 */
export function migrateRows(rows) {
  const changed = [];
  for (const r of rows) {
    const d = r.dong === undefined ? r.dong : normDong(r.dong);
    const h = r.ho === undefined ? r.ho : normHo(r.ho);
    if (d !== r.dong || h !== r.ho) { r.dong = d; r.ho = h; changed.push(r); }
  }
  return changed;
}

export function uniq(rows, key) {
  return [...new Set(rows.map((r) => r[key]).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
}
