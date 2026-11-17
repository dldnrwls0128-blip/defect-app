/* 엑셀 파싱 전용 워커 — 메인 화면이 멈추지 않게 29k행도 백그라운드에서 처리 */
import * as XLSX from "../vendor/xlsx/xlsx.mjs";
import { rowsToRecords } from "./data.js";

export function parseWorkbook(buf) {
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", dense: true, cellStyles: false, cellHTML: false, cellText: false, cellDates: false });
  let lastErr;
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
    const hi = rows.slice(0, 20).findIndex((r) => Array.isArray(r) && r.some((c) => String(c ?? "").trim() === "순번"));
    if (hi < 0) { lastErr = new Error("헤더(순번 열)를 찾지 못했습니다."); continue; }
    try {
      const records = rowsToRecords(rows[hi], rows.slice(hi + 1));
      return { sheet: name, records, keys: records.keys };
    } catch (e) { lastErr = e; }
  }
  throw lastErr || new Error("시트가 없습니다.");
}

if (typeof self !== "undefined" && typeof window === "undefined") {
  self.onmessage = (ev) => {
    const t0 = Date.now();
    try {
      const res = parseWorkbook(ev.data.buf);
      self.postMessage({ ok: true, ...res, ms: Date.now() - t0 });
    } catch (e) {
      self.postMessage({ ok: false, error: String(e?.message || e) });
    }
  };
}
