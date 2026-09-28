// 브라우저 임시 보관 — 작업 내용을 이 브라우저(IndexedDB)에 잠가서 남겨 두고, 다시 열면 이어서 할 수 있게 합니다.
// 정식 저장은 여전히 JSON 파일입니다. 여기 남는 것은 항상 비밀번호로 암호화된 내용뿐이며, 비밀번호 자체는 저장하지 않습니다.
// localStorage 대신 IndexedDB를 쓰는 이유: 수강신청 파일이 여러 개면 5MB를 넘길 수 있습니다.

const DB_NAME = "curriculum-check";
const STORE = "keep";
const KEY = "snapshot";
export const KEEP_MAX_AGE_DAYS = 7;

function openDb() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") return reject(new Error("이 브라우저는 임시 보관을 지원하지 않습니다."));
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("브라우저 저장소를 열지 못했습니다."));
    req.onblocked = () => reject(new Error("브라우저 저장소가 다른 탭에서 쓰이고 있습니다."));
  });
}

function tx(db, mode, run) {
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE, mode);
    const req = run(t.objectStore(STORE));
    t.oncomplete = () => resolve(req && req.result);
    t.onerror = () => reject(t.error || new Error("브라우저 저장소에 쓰지 못했습니다."));
    t.onabort = () => reject(t.error || new Error("브라우저 저장소 작업이 중단되었습니다."));
  });
}

// { savedAt: ISO, summary: {…}, payload: {encrypted, salt, iv, data} } 또는 null. 너무 오래된 것은 지우고 null.
export async function keepRead() {
  const db = await openDb();
  try {
    const rec = await tx(db, "readonly", (s) => s.get(KEY));
    if (!rec || !rec.payload || !rec.savedAt) return null;
    const age = Date.now() - new Date(rec.savedAt).getTime();
    if (!(age >= 0) || age > KEEP_MAX_AGE_DAYS * 86400000) {
      await tx(db, "readwrite", (s) => s.delete(KEY));
      return null;
    }
    return rec;
  } finally {
    db.close();
  }
}

export async function keepWrite(rec) {
  const db = await openDb();
  try {
    await tx(db, "readwrite", (s) => s.put(rec, KEY));
  } finally {
    db.close();
  }
}

export async function keepClear() {
  let db;
  try {
    db = await openDb();
    await tx(db, "readwrite", (s) => s.delete(KEY));
  } catch (e) {
    // 지울 수 없는 환경이면 남는 것도 없습니다
  } finally {
    if (db) db.close();
  }
}
