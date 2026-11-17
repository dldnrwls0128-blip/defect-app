/** 업체별/세대별 조치 리스트: 엑셀(xlsx)·CSV·카톡용 텍스트 */
import { hoLabel, shortDong, groupByTradeVendor, sortUnit } from "./data.js";

let _xlsx;
async function XLSX() {
  if (!_xlsx) _xlsx = await import("../vendor/xlsx/xlsx.mjs");
  return _xlsx;
}

export const EXPORT_HEAD = ["동", "호수", "위치", "부위", "공종", "세부공종", "유형", "시공업체명", "하자내용", "DTspace상태", "내 상태", "완료일", "메모", "사진", "순번", "접수일"];

export function exportRows(rows, ctx) {
  return [...rows].sort(sortUnit).map((r) => {
    const e = ctx.edit(r.id);
    return [r.dong, r.ho, r.loc || "", r.part || "", r.trade || "", r.sub || "", r.type || "", r.vendor || "", r.text || "",
      r.dts || "", ctx.status(r), e?.doneDate || "", e?.memo || "", ctx.photos(r.id) || "", r.id, r.rcv || ""];
  });
}

export async function toXlsx(rows, ctx, sheetName = "조치리스트") {
  const X = await XLSX();
  const aoa = [EXPORT_HEAD, ...exportRows(rows, ctx)];
  const ws = X.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [5, 6, 9, 10, 14, 14, 9, 18, 40, 9, 8, 11, 24, 5, 8, 11].map((w) => ({ wch: w }));
  ws["!autofilter"] = { ref: `A1:P${aoa.length}` };
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31).replace(/[\\/?*[\]:]/g, " "));
  const out = X.write(wb, { type: "array", bookType: "xlsx", compression: true });
  return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

export function toCsv(rows, ctx) {
  const q = (v) => {
    const s = String(v ?? "");
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [EXPORT_HEAD, ...exportRows(rows, ctx)].map((r) => r.map(q).join(","));
  return new Blob(["\uFEFF" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
}

const today = () => {
  const d = new Date();
  return `${d.getMonth() + 1}/${d.getDate()}`;
};
const line = (r, n, ctx) => {
  const st = ctx.status(r);
  const loc = [r.loc, r.sub].filter(Boolean).join(" ");
  return `${n}. ${loc} - ${(r.text || "").replace(/\s*\n\s*/g, " ")}${r.type ? ` [${r.type}]` : ""}${st !== "미처리" ? ` (${st})` : ""}`;
};

/** 세대 공유문: 공종/업체별 묶음 */
export function unitText(dong, ho, rows, ctx, label) {
  const groups = groupByTradeVendor(rows);
  const out = [`[${dong}동 ${hoLabel(ho)} ${label}] ${today()} 총 ${rows.length}건`];
  for (const g of groups) {
    out.push("", `■ ${g.trade} / ${g.vendor} (${g.rows.length})`);
    g.rows.forEach((r, i) => out.push(line(r, i + 1, ctx)));
  }
  return out.join("\n");
}

/** 업체 공유문: 동·호수별 묶음 */
export function vendorText(vendor, rows, ctx, label) {
  const by = new Map();
  for (const r of [...rows].sort(sortUnit)) {
    const k = `${r.dong}|${r.ho}`;
    if (!by.has(k)) by.set(k, []);
    by.get(k).push(r);
  }
  const out = [`[${vendor}] ${label} 조치요청 ${today()}`, `세대 ${by.size}곳 · 총 ${rows.length}건`];
  for (const [k, list] of by) {
    const [d, h] = k.split("|");
    out.push("", `■ ${d}동 ${hoLabel(h)} (${list.length})`);
    list.forEach((r, i) => out.push(line(r, i + 1, ctx)));
  }
  return out.join("\n");
}

export async function shareText(text, title = "") {
  try {
    if (navigator.share) {
      await navigator.share({ text, title });
      return "shared";
    }
  } catch (e) {
    if (e?.name === "AbortError") return "cancel";
  }
  try {
    await navigator.clipboard.writeText(text);
    return "copied";
  } catch {
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
    return "copied";
  }
}

export { shortDong };
