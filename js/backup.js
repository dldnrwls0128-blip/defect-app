/** 백업 zip: 내 처리기록(상태·메모·완료일) + 사진. 하자 원본은 엑셀 재가져오기로 복구(용량 절약). */
import * as db from "./db.js";

const J = () => {
  if (!globalThis.JSZip) throw new Error("JSZip을 불러오지 못했습니다.");
  return globalThis.JSZip;
};

export async function exportBackup() {
  const zip = new (J())();
  const edits = await db.getAll("edits");
  const photos = await db.getAll("photos");
  const meta = { app: "defect-app", version: 1, exportedAt: new Date().toISOString(), imports: (await db.getMeta("imports")) || [] };
  const pmeta = [];
  for (const p of photos) {
    const file = `photos/${p.defectId}_${p.kind}_${p.id}.jpg`;
    zip.file(file, p.blob);
    if (p.thumb) zip.file(`thumbs/${p.id}.jpg`, p.thumb);
    const { blob, thumb, ...rest } = p;
    pmeta.push({ ...rest, file, thumbFile: p.thumb ? `thumbs/${p.id}.jpg` : null });
  }
  zip.file("meta.json", JSON.stringify(meta, null, 1));
  zip.file("edits.json", JSON.stringify(edits, null, 1));
  zip.file("photos.json", JSON.stringify(pmeta, null, 1));
  return zip.generateAsync({ type: "blob", compression: "STORE" });
}

/** 병합 복원: 같은 순번 기록은 더 최근 수정본 유지, 사진은 id 기준 추가 */
export async function importBackup(file) {
  const zip = await J().loadAsync(file);
  const metaF = zip.file("meta.json");
  if (!metaF) throw new Error("하자관리 백업 파일이 아닙니다.");
  const meta = JSON.parse(await metaF.async("string"));
  if (meta.app !== "defect-app") throw new Error("하자관리 백업 파일이 아닙니다.");
  const edits = JSON.parse((await zip.file("edits.json")?.async("string")) || "[]");
  const pmeta = JSON.parse((await zip.file("photos.json")?.async("string")) || "[]");
  const cur = new Map((await db.getAll("edits")).map((e) => [e.id, e]));
  let e1 = 0;
  const putE = [];
  for (const e of edits) {
    const c = cur.get(e.id);
    if (!c || (e.updatedAt || "") > (c.updatedAt || "")) { putE.push(e); e1++; }
  }
  await db.putMany("edits", putE);
  const have = new Set((await db.getAll("photos")).map((p) => p.id));
  let p1 = 0;
  for (const p of pmeta) {
    if (have.has(p.id)) continue;
    const f = zip.file(p.file);
    if (!f) continue;
    const blob = new Blob([await f.async("arraybuffer")], { type: "image/jpeg" });
    const tf = p.thumbFile && zip.file(p.thumbFile);
    const thumb = tf ? new Blob([await tf.async("arraybuffer")], { type: "image/jpeg" }) : null;
    const { file: _f, thumbFile: _t, ...rest } = p;
    await db.put("photos", { ...rest, blob, thumb });
    p1++;
  }
  return { edits: e1, photos: p1, total: edits.length, totalPhotos: pmeta.length };
}
