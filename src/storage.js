// 기존 Claude 아티팩트가 쓰던 window.storage.get/set/delete(key, shared) 형태의 저장소를
// 그대로 흉내 내는 얇은 대체 계층(shim)이다. App.jsx의 나머지 코드(스토리지 I/O 섹션 제외)는
// 이 인터페이스가 어디에 실제로 저장되는지 전혀 몰라도 되도록, 원래와 똑같은 모양
// ({ value } 반환, key/value/shared 인자)을 유지한다.
//
//  - shared === true  → 여러 사람(학교 코드로 접속하는 모든 기기)이 함께 보는 데이터.
//                        저장 키에 따라 Firestore의 학교별 문서로 나눠 저장한다(아래 sharedRef).
//                          paps:config:{코드}    → schools/{코드}/data/config   (명단·등급 기준·설정, 비밀번호 없음)
//                          paps:records:{코드}   → schools/{코드}/data/records  (측정 기록. 신장·체중·BMI·체지방률 값은
//                                                    App.jsx가 보내기 전에 빼서 이 기기에만 둔다 — healthLocal.js)
//                          paps:auditlog:{코드}  → schools/{코드}/data/auditlog (변경 이력)
//                          paps:backuplog:{코드} → schools/{코드}/data/backuplog(백업 이력)
//                          paps:stats:*          → kv/{키}                      (방문 집계)
//  - shared === false → "이 브라우저(기기)에만" 남아야 하는 데이터(워크스페이스 코드 기억,
//                        비밀번호 시도 횟수, 화면 잠금 표시 등). 서버로 보내지 않고
//                        localStorage에만 저장한다.
//
// ───────────────────────────────────────────────────────────────
// 보안 구조(firestore.rules와 짝)
//  - 모든 기기는 Firebase 익명 로그인을 한다(이름·이메일 없음). 그 기기 고유번호(uid)로 권한을 구분한다.
//  - 비밀번호는 원문을 저장하지 않고 SHA-256 해시만 schools/{코드}/meta/secret에 둔다(아무도 읽을 수 없음).
//  - 들어올 때 claims/{uid}에 "내가 계산한 해시"를 먼저 쓰고, 그다음 members/{uid}(구성원 문서)를 쓴다.
//    보안 규칙이 해시가 맞는지 비교해서 맞을 때만 구성원 문서를 만들어 준다.
//  - 예전 방식(kv/paps:config:{코드}에 비밀번호 원문이 있던 코드)은 개설자가 새 버전에서 처음
//    로그인할 때 새 방식으로 옮긴다. 옮기기 전에도 예전 문서는 개설자 말고는 읽을 수 없다.
// ───────────────────────────────────────────────────────────────
import { db, ensureSignedIn } from "./firebase";
import {
  doc,
  getDoc,
  getDocs,
  setDoc,
  updateDoc,
  deleteDoc,
  collection,
  writeBatch,
  serverTimestamp,
} from "firebase/firestore";

const KV_COLLECTION = "kv";

// ---- 연습모드 ----
// 첫 화면의 "연습모드 체험"으로 들어가면 이 특별한 학교 코드를 쓴다. 코드에 "/"가 들어 있어서
// 선생님이 입력한 코드(sanitizeWorkspaceCode가 "/"를 "-"로 바꾼다)와는 절대 겹치지 않는다.
// 이 코드로 끝나는 저장 키는 shared 여부와 상관없이 서버(Firestore)로 보내지 않고, 이 기기의
// localStorage에만 따로 저장한다. 그래서 연습 중에 무엇을 하든 실제 학교 데이터와 섞이지 않는다.
export const PRACTICE_CODE = "연습모드/체험";
const PRACTICE_PREFIX = "paps-practice:";
function isPracticeKey(key) {
  return String(key).endsWith(":" + PRACTICE_CODE);
}
// 연습 데이터를 모두 지운다(연습 시작·초기화·종료 때 사용).
export function clearPracticeStore() {
  try {
    const ls = window.localStorage;
    const doomed = [];
    for (let i = 0; i < ls.length; i++) {
      const k = ls.key(i);
      if (k && (k.startsWith(PRACTICE_PREFIX) || isPracticeKey(k))) doomed.push(k);
    }
    doomed.forEach(k => ls.removeItem(k));
  } catch (e) {
    // localStorage를 쓸 수 없는 환경이면 지울 것도 없다.
  }
}

function sanitizeDocId(key) {
  // Firestore 문서 ID는 "/"를 쓸 수 없고 1500바이트 이하여야 한다. 원래 저장 키들은
  // "paps:xxx:yyy" 형태의 콜론 구분 문자열이라 대부분 그대로 써도 안전하지만, 혹시 모를
  // 경우를 대비해 방어적으로 다듬는다.
  return String(key).replace(/\//g, "_").slice(0, 400);
}

function safeLocalGet(key) {
  try {
    const v = window.localStorage.getItem(key);
    return v === null ? null : { value: v };
  } catch (e) {
    return null;
  }
}
function safeLocalSet(key, value) {
  try {
    window.localStorage.setItem(key, value);
    return true;
  } catch (e) {
    return false;
  }
}
function safeLocalDelete(key) {
  try {
    window.localStorage.removeItem(key);
    return true;
  } catch (e) {
    return false;
  }
}

const SCHOOLS = "schools";
const DATA_DOCS = [
  ["paps:config:", "config"],
  ["paps:records:", "records"],
  ["paps:auditlog:", "auditlog"],
  ["paps:backuplog:", "backuplog"],
];

// 저장 키 → Firestore 문서 위치. 위 목록에 없는 공유 키(예전 접근 목록 paps:access: 등)는
// 이제 서버에 쓰지 않는다(null).
function sharedRef(key) {
  const k = String(key);
  for (const [prefix, name] of DATA_DOCS) {
    if (k.startsWith(prefix)) {
      const code = k.slice(prefix.length);
      return isValidCode(code) ? doc(db, SCHOOLS, code, "data", name) : null;
    }
  }
  if (k.startsWith("paps:stats:")) return doc(db, KV_COLLECTION, sanitizeDocId(k));
  return null;
}

export const storage = {
  async get(key, shared) {
    if (isPracticeKey(key)) return safeLocalGet(PRACTICE_PREFIX + key);
    if (!shared) return safeLocalGet(key);
    try {
      const ref = sharedRef(key);
      if (!ref) return null;
      await ensureSignedIn();
      const snap = await getDoc(ref);
      if (!snap.exists()) return null;
      const data = snap.data();
      return typeof data.value === "string" ? { value: data.value } : null;
    } catch (e) {
      return null;
    }
  },
  async set(key, value, shared) {
    if (isPracticeKey(key)) return safeLocalSet(PRACTICE_PREFIX + key, String(value));
    if (!shared) return safeLocalSet(key, value);
    try {
      const ref = sharedRef(key);
      if (!ref) return false;
      await ensureSignedIn();
      await setDoc(ref, { value: String(value), updatedAt: Date.now() });
      return true;
    } catch (e) {
      return false;
    }
  },
  async delete(key, shared) {
    if (isPracticeKey(key)) return safeLocalDelete(PRACTICE_PREFIX + key);
    if (!shared) return safeLocalDelete(key);
    try {
      const ref = sharedRef(key);
      if (!ref) return false;
      await ensureSignedIn();
      await deleteDoc(ref);
      return true;
    } catch (e) {
      return false;
    }
  },
};

/* ============================== 접근 권한(보안) ============================== */

// Firestore 문서 이름으로 쓸 수 없는 코드("." ".." "__이름__")와 너무 긴 코드는 받지 않는다.
export function isValidCode(code) {
  const c = String(code || "");
  return c.length >= 1 && c.length <= 40 && c !== "." && c !== ".." && !/^__.*__$/.test(c) && !c.includes("/");
}

function infoRef(code) { return doc(db, SCHOOLS, code, "meta", "info"); }
function secretRef(code) { return doc(db, SCHOOLS, code, "meta", "secret"); }
function claimRef(code, uid) { return doc(db, SCHOOLS, code, "claims", uid); }
function memberRef(code, uid) { return doc(db, SCHOOLS, code, "members", uid); }
function dataRef(code, name) { return doc(db, SCHOOLS, code, "data", name); }
// 예전 방식 저장 위치(kv 컬렉션, 문서 값은 JSON 문자열)
const LEGACY_KINDS = ["config", "records", "access", "auditlog", "backuplog"];
function legacyRef(kind, code) { return doc(db, KV_COLLECTION, "paps:" + kind + ":" + code); }

async function sha256Hex(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
}
// 같은 비밀번호라도 코드·종류가 다르면 전혀 다른 해시가 나오도록 함께 섞는다.
export function hashPassword(code, kind, pw) {
  return sha256Hex("smart-paps-v1|" + code + "|" + kind + "|" + String(pw || "").trim());
}

export async function currentUid() {
  const user = await ensureSignedIn();
  return user.uid;
}

function isDenied(e) {
  return !!e && (e.code === "permission-denied" || e.code === "firestore/permission-denied");
}

// { state: "none" | "ready" | "legacy", viewerOn }
//  - none: 없는 코드 / ready: 새 방식 / legacy: 예전 방식(아직 옮기지 않음)
export async function probeSchool(code) {
  if (!isValidCode(code)) return { state: "none", viewerOn: false };
  await ensureSignedIn();
  const info = await getDoc(infoRef(code));
  if (info.exists()) return { state: "ready", viewerOn: !!info.data().viewerOn };
  try {
    const legacy = await getDoc(legacyRef("config", code));
    return { state: legacy.exists() ? "legacy" : "none", viewerOn: false };
  } catch (e) {
    // 예전 방식 설정 문서는 있으면 개설자 말고는 읽을 수 없게 막혀 있다 → 거절되면 "있다"는 뜻
    if (isDenied(e)) return { state: "legacy", viewerOn: false };
    throw e;
  }
}

// 내 구성원 문서(없거나 읽을 수 없으면 null). { id, type, status, name, ... }
export async function readMyMember(code) {
  if (!isValidCode(code)) return null;
  const uid = await currentUid();
  try {
    const snap = await getDoc(memberRef(code, uid));
    return snap.exists() ? { id: uid, ...snap.data() } : null;
  } catch (e) {
    if (isDenied(e)) return null;
    throw e;
  }
}

// 해시(또는 예전 비밀번호)를 먼저 따로 제출하고, 그다음 write()로 구성원 문서를 쓴다.
// 비밀번호가 틀리면 규칙이 거절한다. 제출 기록을 먼저 따로 저장해야 틀린 시도도 "5초에 한 번"
// 제한에 걸린다(한 묶음으로 쓰면 실패할 때 기록이 남지 않아 제한이 안 걸린다).
async function submitClaimThen(code, claim, write) {
  const uid = await currentUid();
  try {
    await setDoc(claimRef(code, uid), { ...claim, at: serverTimestamp() });
    await write(uid);
    return true;
  } catch (e) {
    if (isDenied(e)) return false;
    throw e;
  }
}

function founderEntry() {
  return { type: "founder", status: "approved", name: "개설자", submittedAt: Date.now() };
}

// 새 코드 만들기. 이미 있는 코드면 false.
export async function createSchool(code, viewerPassword, founderPassword) {
  if (!isValidCode(code) || !String(founderPassword || "").trim()) return false;
  const uid = await currentUid();
  const founderHash = await hashPassword(code, "founder", founderPassword);
  const viewerOn = !!String(viewerPassword || "").trim();
  const viewerHash = viewerOn ? await hashPassword(code, "viewer", viewerPassword) : "";
  const batch = writeBatch(db);
  batch.set(secretRef(code), { founder: founderHash, viewer: viewerHash });
  batch.set(infoRef(code), { createdAt: Date.now(), v: 2, viewerOn });
  batch.set(claimRef(code, uid), { h: founderHash, at: serverTimestamp() });
  batch.set(memberRef(code, uid), founderEntry());
  try {
    await batch.commit();
    return true;
  } catch (e) {
    if (isDenied(e)) return false;
    throw e;
  }
}

// JSON 문자열 안에 저장된 모양 그대로(따옴표·역슬래시 이스케이프) 맞춰 보내야 규칙이 비교할 수 있다.
function jsonEscaped(text) {
  return JSON.stringify(String(text)).slice(1, -1);
}

function parseJson(raw, fallback) {
  if (typeof raw !== "string") return fallback;
  try {
    const v = JSON.parse(raw);
    return v === null || v === undefined ? fallback : v;
  } catch (e) {
    return fallback;
  }
}

// 예전 방식 코드를 개설자가 처음 로그인할 때 새 방식으로 옮긴다(개설자가 된 뒤라 예전 문서를 읽을 수 있다).
async function migrateLegacy(code, founderPassword) {
  const uid = await currentUid();
  const raws = {};
  for (const kind of LEGACY_KINDS) {
    const snap = await getDoc(legacyRef(kind, code));
    raws[kind] = snap.exists() && typeof snap.data().value === "string" ? snap.data().value : null;
  }
  const cfg = parseJson(raws.config, {});
  const settings = { ...(cfg.settings || {}) };
  const viewerPw = String(settings.viewerPassword || "").trim();
  delete settings.viewerPassword;
  delete settings.founderPassword;
  const students = Array.isArray(cfg.students)
    ? cfg.students.map(({ id, grade, classNum, number, gender }) => ({ id, grade, classNum, number, gender }))
    : [];
  const nextCfg = { ...cfg, students, settings };
  const founderHash = await hashPassword(code, "founder", founderPassword);
  const viewerHash = viewerPw ? await hashPassword(code, "viewer", viewerPw) : "";

  const oldRequests = parseJson(raws.access, { requests: [] }).requests || [];
  const audit = parseJson(raws.auditlog, []);
  const auditList = Array.isArray(audit) ? audit : [];
  auditList.push({
    type: "access",
    ts: Date.now(),
    by: "개설자",
    message: "개설자님이 새 보안 방식(비밀번호 암호화·서버 권한 확인)으로 옮김"
      + (oldRequests.length > 0 ? ` — 예전 접근 신청 ${oldRequests.length}건은 정리되어, 동료 선생님은 다시 접근 신청해야 함` : ""),
  });

  const now = Date.now();
  const batch = writeBatch(db);
  batch.set(secretRef(code), { founder: founderHash, viewer: viewerHash });
  batch.set(infoRef(code), { createdAt: cfg.createdAt || now, v: 2, viewerOn: !!viewerPw });
  batch.set(dataRef(code, "config"), { value: JSON.stringify(nextCfg), updatedAt: now });
  if (raws.records !== null) batch.set(dataRef(code, "records"), { value: raws.records, updatedAt: now });
  batch.set(dataRef(code, "auditlog"), { value: JSON.stringify(auditList.slice(-300)), updatedAt: now });
  if (raws.backuplog !== null) batch.set(dataRef(code, "backuplog"), { value: raws.backuplog, updatedAt: now });
  batch.delete(claimRef(code, uid)); // 예전 비밀번호 원문을 담았던 제출 기록도 지운다
  for (const kind of LEGACY_KINDS) {
    if (raws[kind] !== null) batch.delete(legacyRef(kind, code)); // 비밀번호 원문이 든 예전 문서 삭제
  }
  await batch.commit();
}

// 결과: { ok: true, migrated? } | { ok: false, reason: "bad-code" | "bad-password" }
export async function founderLogin(code, founderPassword) {
  const pw = String(founderPassword || "").trim();
  if (!pw) return { ok: false, reason: "bad-password" };
  const { state } = await probeSchool(code);
  if (state === "none") return { ok: false, reason: "bad-code" };
  const write = (uid) => setDoc(memberRef(code, uid), founderEntry());
  if (state === "legacy") {
    const ok = await submitClaimThen(code, { h: "", legacy: jsonEscaped(pw) }, write);
    if (!ok) return { ok: false, reason: "bad-password" };
    await migrateLegacy(code, pw);
    return { ok: true, migrated: true };
  }
  const h = await hashPassword(code, "founder", pw);
  const ok = await submitClaimThen(code, { h }, write);
  return ok ? { ok: true } : { ok: false, reason: "bad-password" };
}

// 동료 교사의 접근 신청. 조회는 비밀번호가 맞으면 바로 승인, 수정 권한은 승인 대기로 들어간다.
// 결과: { ok: true, entry } | { ok: false, reason: "bad-code" | "legacy" | "no-password-set" | "bad-password" }
export async function requestAccessEntry(code, viewerPassword, name, wantsEdit) {
  const { state, viewerOn } = await probeSchool(code);
  if (state === "none") return { ok: false, reason: "bad-code" };
  if (state === "legacy") return { ok: false, reason: "legacy" };
  const existing = await readMyMember(code);
  if (existing && existing.status !== "denied") return { ok: true, entry: existing };
  if (!viewerOn) return { ok: false, reason: "no-password-set" };
  const entry = {
    type: wantsEdit ? "editor" : "viewer",
    status: wantsEdit ? "pending" : "approved",
    name: String(name || "").trim().slice(0, 60),
    submittedAt: Date.now(),
  };
  const h = await hashPassword(code, "viewer", viewerPassword);
  const ok = await submitClaimThen(code, { h }, (uid) => setDoc(memberRef(code, uid), entry));
  return ok ? { ok: true, entry: { id: await currentUid(), ...entry } } : { ok: false, reason: "bad-password" };
}

// 빔프로젝터 화면·자동 잠금 풀기: 접근 신청 비밀번호를 서버 규칙으로 확인한다.
export async function verifyViewer(code, viewerPassword) {
  if (!String(viewerPassword || "").trim()) return false;
  const h = await hashPassword(code, "viewer", viewerPassword);
  return submitClaimThen(code, { h }, (uid) => updateDoc(memberRef(code, uid), { verifiedAt: Date.now() }));
}

// 개설자: 비밀번호 바꾸기. 빈 칸은 그대로 두고, viewerOff가 true면 접근 신청을 끈다.
// 결과: { changed, viewerOn }
export async function changePasswords(code, { viewerPassword, founderPassword, viewerOff }) {
  const patch = {};
  if (String(founderPassword || "").trim()) patch.founder = await hashPassword(code, "founder", founderPassword);
  if (viewerOff) patch.viewer = "";
  else if (String(viewerPassword || "").trim()) patch.viewer = await hashPassword(code, "viewer", viewerPassword);
  if (Object.keys(patch).length === 0) return { changed: false };
  const batch = writeBatch(db);
  batch.update(secretRef(code), patch);
  if ("viewer" in patch) batch.update(infoRef(code), { viewerOn: patch.viewer !== "" });
  await batch.commit();
  const info = await getDoc(infoRef(code));
  return { changed: true, viewerOn: info.exists() ? !!info.data().viewerOn : false };
}

// 구성원 목록(개설자·승인된 수정 권한자만 읽을 수 있음)
export async function listMembers(code) {
  const q = await getDocs(collection(db, SCHOOLS, code, "members"));
  return q.docs.map(d => ({ id: d.id, ...d.data() }));
}
export async function updateMember(code, uid, patch) {
  await updateDoc(memberRef(code, uid), patch);
}
export async function deleteMember(code, uid) {
  await deleteDoc(memberRef(code, uid));
}
// 개설자 승계: 상대를 개설자로, 나는 승인된 수정 권한자로 한 번에 바꾼다.
export async function transferFounderRole(code, targetUid, myNewName) {
  const uid = await currentUid();
  const batch = writeBatch(db);
  batch.update(memberRef(code, targetUid), { type: "founder", status: "approved", decidedAt: Date.now() });
  batch.update(memberRef(code, uid), { type: "editor", status: "approved", name: String(myNewName || "이전 개설자").slice(0, 60), decidedAt: Date.now() });
  await batch.commit();
}

// 마감: 데이터·설정·비밀번호·구성원·제출 기록을 모두 지운다(개설자 전용).
// 내 구성원 문서를 지우면 개설자 권한이 사라지므로 맨 마지막에 지운다.
export async function deleteSchoolAll(code) {
  const uid = await currentUid();
  for (const name of ["records", "config", "auditlog", "backuplog"]) {
    await deleteDoc(dataRef(code, name)).catch(() => {});
  }
  const claims = await getDocs(collection(db, SCHOOLS, code, "claims")).catch(() => null);
  if (claims) {
    for (const d of claims.docs) if (d.id !== uid) await deleteDoc(d.ref).catch(() => {});
  }
  const members = await getDocs(collection(db, SCHOOLS, code, "members")).catch(() => null);
  if (members) {
    for (const d of members.docs) if (d.id !== uid) await deleteDoc(d.ref).catch(() => {});
  }
  await deleteDoc(infoRef(code)).catch(() => {});
  await deleteDoc(secretRef(code)).catch(() => {});
  await deleteDoc(claimRef(code, uid)).catch(() => {});
  await deleteDoc(memberRef(code, uid)).catch(() => {});
}
