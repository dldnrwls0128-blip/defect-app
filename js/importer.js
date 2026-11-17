/** 엑셀 가져오기: 워커 파싱 → 순번 기준 병합(내 처리기록·사진은 그대로 유지) */
import * as db from "./db.js";
import { sameSource, SRC_KEYS } from "./data.js";

function parseInWorker(buf) {
  return new Promise((resolve, reject) => {
    let w;
    try {
      w = new Worker(new URL("./import-worker.js", import.meta.url), { type: "module" });
    } catch (e) {
      reject(e);
      return;
    }
    w.onmessage = (ev) => {
      w.terminate();
      ev.data.ok ? resolve(ev.data) : reject(new Error(ev.data.error));
    };
    w.onerror = (e) => {
      w.terminate();
      reject(new Error(e.message || "워커 오류"));
    };
    w.postMessage({ buf }, [buf]);
  });
}

async function parseMainThread(buf) {
  const mod = await import("./import-worker.js");
  return mod.parseWorkbook(buf);
}

/**
 * @param {File} file
 * @param {object[]} existing 현재 메모리의 하자 레코드
 * @param {(msg:string)=>void} onStep
 */
export async function importFile(file, existing, onStep = () => {}, confirmOld = async () => true) {
  const t0 = performance.now();
  onStep("엑셀 읽는 중…");
  const buf = await file.arrayBuffer();
  let parsed;
  try {
    parsed = await parseInWorker(buf.slice(0));
  } catch (e) {
    if (/형식|헤더|열/.test(String(e.message))) throw e;
    onStep("엑셀 분석 중…(호환 모드)");
    parsed = await parseMainThread(buf);
  }
  const tParse = performance.now();
  const recs = parsed.records;
  if (!recs.length) throw new Error("가져올 하자 행이 없습니다.");

  const keys = parsed.keys || [];
  const missingKeys = new Set(SRC_KEYS.filter((k) => !keys.includes(k)));
  // 오래된 엑셀로 최신 데이터를 덮어쓰지 않도록 확인
  const fileMax = recs.reduce((m, r) => (r.rcv && r.rcv > m ? r.rcv : m), "");
  const fileDongs = new Set(recs.map((r) => r.dong));
  const storedMax = existing.reduce((m, r) => (fileDongs.has(r.dong) && r.rcv && r.rcv > m ? r.rcv : m), "");
  if (fileMax && storedMax && fileMax < storedMax) {
    const ok = await confirmOld(fileMax, storedMax);
    if (!ok) throw new Error("가져오기를 취소했습니다.");
  }
  onStep(`${recs.length.toLocaleString()}건 병합 중…`);
  const importedAt = new Date().toISOString();
  const old = new Map(existing.map((r) => [r.id, r]));
  const seen = new Set();
  const dongsInFile = new Set(recs.map((r) => r.dong));
  const toPut = [];
  let added = 0, changed = 0, same = 0, gone = 0;
  for (const r of recs) {
    seen.add(r.id);
    const o = old.get(r.id);
    if (!o) { added++; r._first = importedAt; toPut.push(r); continue; }
    r._first = o._first || importedAt;
    for (const k of missingKeys) if (o[k] !== undefined) r[k] = o[k]; // 이번 엑셀에 없는 열은 기존 값 유지
    const same0 = sameSource(o, r, keys);
    if (same0 && !o._gone && o.days === r.days) { same++; continue; }
    if (!same0) { changed++; r._prevDts = o.dts !== r.dts ? o.dts || "" : o._prevDts; r._changed = importedAt; }
    else { if (o._prevDts !== undefined) r._prevDts = o._prevDts; if (o._changed) r._changed = o._changed; same++; }
    toPut.push(r);
  }
  // 부분 엑셀(필터 걸린 내보내기)로 오표시하지 않도록: 동별로 기존 건수의 80% 이상이 들어온 동만 '엑셀에 없음' 표시
  const inFile = new Map(), before = new Map();
  for (const r of recs) inFile.set(r.dong, (inFile.get(r.dong) || 0) + 1);
  for (const o of existing) if (!o._gone) before.set(o.dong, (before.get(o.dong) || 0) + 1);
  const fullDong = (d) => dongsInFile.has(d) && (inFile.get(d) || 0) >= 0.8 * (before.get(d) || 0);
  for (const o of existing) {
    if (!seen.has(o.id) && !o._gone && fullDong(o.dong)) {
      gone++;
      toPut.push({ ...o, _gone: importedAt });
    }
  }
  await db.putMany("defects", toPut, (n, t) => onStep(`저장 중… ${Math.round((n / t) * 100)}%`));
  const info = {
    fileName: file.name, importedAt, rows: recs.length, sheet: parsed.sheet,
    added, changed, same, gone,
    maxRcv: fileMax,
    parseMs: Math.round(tParse - t0), totalMs: Math.round(performance.now() - t0),
  };
  const hist = (await db.getMeta("imports")) || [];
  hist.unshift(info);
  await db.setMeta("imports", hist.slice(0, 20));
  return info;
}
