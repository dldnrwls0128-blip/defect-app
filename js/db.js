/** IndexedDB — 하자 원본(defects), 내 처리기록(edits), 사진(photos), 메타(meta). 서버 전송 없음. */
export const DB_NAME = "haja-gwanri";
export const DB_VERSION = 1;

let _dbp;
export function openDb() {
  if (_dbp) return _dbp;
  _dbp = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains("defects")) db.createObjectStore("defects", { keyPath: "id" });
      if (!db.objectStoreNames.contains("edits")) db.createObjectStore("edits", { keyPath: "id" });
      if (!db.objectStoreNames.contains("photos")) {
        const s = db.createObjectStore("photos", { keyPath: "id" });
        s.createIndex("by_defect", "defectId", { unique: false });
      }
      if (!db.objectStoreNames.contains("meta")) db.createObjectStore("meta", { keyPath: "key" });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return _dbp;
}

export function done(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error("aborted"));
  });
}
const rq = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });

export async function getAll(store) {
  const db = await openDb();
  return rq(db.transaction(store).objectStore(store).getAll());
}
export async function get(store, key) {
  const db = await openDb();
  return rq(db.transaction(store).objectStore(store).get(key));
}
export async function put(store, value) {
  const db = await openDb();
  const tx = db.transaction(store, "readwrite");
  tx.objectStore(store).put(value);
  await done(tx);
  return value;
}
export async function del(store, key) {
  const db = await openDb();
  const tx = db.transaction(store, "readwrite");
  tx.objectStore(store).delete(key);
  await done(tx);
}
export async function putMany(store, values, onProgress) {
  const db = await openDb();
  const CH = 4000;
  for (let i = 0; i < values.length; i += CH) {
    const tx = db.transaction(store, "readwrite");
    const os = tx.objectStore(store);
    for (const v of values.slice(i, i + CH)) os.put(v);
    await done(tx);
    onProgress?.(Math.min(values.length, i + CH), values.length);
  }
}
export async function clear(store) {
  const db = await openDb();
  const tx = db.transaction(store, "readwrite");
  tx.objectStore(store).clear();
  await done(tx);
}
export async function photosOf(defectId) {
  const db = await openDb();
  return rq(db.transaction("photos").objectStore("photos").index("by_defect").getAll(defectId));
}
/** 사진 메타만(블롭 제외) 개수 집계용 */
export async function photoCounts() {
  const db = await openDb();
  const map = new Map();
  await new Promise((resolve, reject) => {
    const r = db.transaction("photos").objectStore("photos").index("by_defect").openKeyCursor();
    r.onsuccess = () => {
      const c = r.result;
      if (!c) return resolve();
      map.set(c.key, (map.get(c.key) || 0) + 1);
      c.continue();
    };
    r.onerror = () => reject(r.error);
  });
  return map;
}
export async function getMeta(key) {
  return (await get("meta", key))?.value;
}
export async function setMeta(key, value) {
  return put("meta", { key, value });
}
