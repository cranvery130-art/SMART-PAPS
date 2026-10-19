// 학생 건강 정보(신장·체중·BMI·체지방률)를 "이 기기(브라우저)에만" 두는 저장소.
//
// 신장·체중·BMI·체지방률은 건강에 관한 민감한 정보라서 외부 클라우드(Firestore)에 올리지 않는다
// (학교 개인정보 처리 안내: 민감정보는 클라우드에 올리지 않음). 측정 기록(records) 중 이 종목들만
// 서버로 보내기 직전에 빼내 이 기기의 localStorage에 따로 두고, 서버에는 "측정했음" 표시
// (값 없는 { localOnly: true, updatedAt, schoolGradeAtMeasure })만 남긴다. 화면에 보여줄 때는
// 이 기기에 있는 값을 다시 합친다. 그래서
//  - 값을 입력한 기기에서는 지금처럼 보이고(나이스 양식·엑셀·JSON 백업에도 들어감),
//  - 다른 기기(동료 선생님 등)에서는 "이 기기에 없음"으로 보인다.
// 비만도(마름·정상·과체중 …) 분류는 값에서 그때그때 계산하므로 따로 저장하지 않는다.
//
// 저장 위치: localStorage "paps:health:{학교 코드}" =
//   { v: 1, records: { "학생id__bmi__연도": 기록, ... }, audit: { 이력id: { prevValue, newValue, prevParts, ts } } }
// (학생 id·종목·연도로 구분. 학생 수백 명 × 2종목이라 수십 KB 수준이어서 localStorage로 충분하고,
//  이름표와 같은 방식으로 "바로(동기적으로)" 저장할 수 있다.)
// 연습모드 코드("연습모드/체험")의 키도 ":연습모드/체험"으로 끝나므로 clearPracticeStore가 함께 지운다.

export const HEALTH_EVENT_IDS = ["bmi", "bodyfat"];
const STORE_VERSION = 1;
const AUDIT_KEEP = 600; // 서버 이력(최근 300건)보다 넉넉하게

export function healthStoreKey(code) { return "paps:health:" + code; }

// 기록 키("학생id__종목__연도", 예전 형식 "학생id__종목")가 건강 종목인지
export function isHealthRecordKey(key) {
  const parts = String(key).split("__");
  return parts.length >= 2 && HEALTH_EVENT_IDS.includes(parts[1]);
}
export function isHealthEventId(eventId) {
  return HEALTH_EVENT_IDS.includes(eventId);
}

// 서버에 남기는 "측정했음" 표시(값 없음)
export function isHealthMarker(entry) {
  return !!entry && entry.localOnly === true;
}
function toMarker(entry) {
  return {
    localOnly: true,
    updatedAt: (entry && entry.updatedAt) || 0,
    schoolGradeAtMeasure: entry && entry.schoolGradeAtMeasure !== undefined ? entry.schoolGradeAtMeasure : null,
  };
}
// 값(또는 신장·체중)이 실제로 들어 있는 기록인지
function hasHealthValue(entry) {
  return !!entry && typeof entry === "object" && !isHealthMarker(entry);
}

function emptyStore() {
  return { v: STORE_VERSION, records: {}, audit: {} };
}

export function loadHealthStore(code) {
  try {
    const raw = window.localStorage.getItem(healthStoreKey(code));
    if (!raw) return emptyStore();
    const obj = JSON.parse(raw);
    return {
      v: STORE_VERSION,
      records: obj && typeof obj.records === "object" && obj.records ? obj.records : {},
      audit: obj && typeof obj.audit === "object" && obj.audit ? obj.audit : {},
    };
  } catch (e) {
    return emptyStore();
  }
}

// 저장한 뒤 다시 읽어 똑같이 들어갔는지 확인한다. 실패하면 false.
export function saveHealthStore(code, store) {
  try {
    const text = JSON.stringify({ v: STORE_VERSION, records: store.records || {}, audit: store.audit || {} });
    window.localStorage.setItem(healthStoreKey(code), text);
    return window.localStorage.getItem(healthStoreKey(code)) === text;
  } catch (e) {
    return false;
  }
}

export function clearHealthStore(code) {
  try { window.localStorage.removeItem(healthStoreKey(code)); } catch (e) { /* 없어도 무방 */ }
}

export function countHealthRecords(code) {
  return Object.keys(loadHealthStore(code).records).length;
}

// 서버에서 받은 기록에 이 기기의 건강 기록을 합친다(이 기기에 있으면 그 값이 우선).
export function mergeLocalHealth(code, serverRecords) {
  const store = loadHealthStore(code);
  const merged = { ...(serverRecords || {}) };
  Object.entries(store.records).forEach(([k, entry]) => {
    if (hasHealthValue(entry)) merged[k] = entry;
  });
  return merged;
}

// 서버로 보내기 직전: 건강 종목 값은 이 기기 저장소로 옮기고, 서버에는 "측정했음" 표시만 보낸다.
//  - 값이 있는 기록 → 이 기기에 저장, 서버에는 표시만
//  - 표시만 있는 기록(다른 기기 값) → 그대로 표시만, 이 기기에 값이 있으면 그 값은 지키기
//  - 목록에서 빠진 건강 기록(되돌리기로 지움, 백업 파일로 통째로 바꿈) → 이 기기에서도 지움
// 반환: { payload, localOk } — payload에는 건강 값이 절대 들어가지 않는다. localOk가 false면
// (이 기기에 저장 실패) 값을 잃지 않도록 부르는 쪽에서 서버 저장을 멈춰야 한다.
export function splitHealthForServer(code, records) {
  const store = loadHealthStore(code);
  const nextLocal = {};
  const payload = {};
  Object.entries(records || {}).forEach(([k, entry]) => {
    if (!isHealthRecordKey(k)) {
      payload[k] = entry;
      return;
    }
    if (hasHealthValue(entry)) {
      nextLocal[k] = entry;
      payload[k] = toMarker(entry);
    } else {
      if (hasHealthValue(store.records[k])) nextLocal[k] = store.records[k];
      payload[k] = toMarker(entry);
    }
  });
  const unchanged = JSON.stringify(nextLocal) === JSON.stringify(store.records);
  const localOk = unchanged || saveHealthStore(code, { ...store, records: nextLocal });
  return { payload, localOk };
}

// 서버 기록 안에 아직 남아 있는(예전 버전이 올린) 건강 값만 골라낸다.
export function findServerHealthValues(serverRecords) {
  const found = {};
  Object.entries(serverRecords || {}).forEach(([k, entry]) => {
    if (isHealthRecordKey(k) && hasHealthValue(entry)) found[k] = entry;
  });
  return found;
}

// 서버에서 찾은 건강 값을 이 기기 저장소에 합친다(이 기기 값이 더 최신이면 그대로 둠).
// 저장 후 다시 읽어서, 찾은 값이 모두(또는 더 최신 값으로) 들어 있는지 확인되면 true.
export function adoptHealthIntoLocal(code, found) {
  const store = loadHealthStore(code);
  const records = { ...store.records };
  Object.entries(found || {}).forEach(([k, entry]) => {
    const mine = records[k];
    if (!hasHealthValue(mine) || (mine.updatedAt || 0) < (entry.updatedAt || 0)) records[k] = entry;
  });
  if (!saveHealthStore(code, { ...store, records })) return false;
  const check = loadHealthStore(code).records;
  return Object.entries(found || {}).every(([k, entry]) => {
    const got = check[k];
    if (!hasHealthValue(got)) return false;
    if ((got.updatedAt || 0) > (entry.updatedAt || 0)) return true;
    return JSON.stringify(got) === JSON.stringify(entry);
  });
}

// 서버 기록에서 건강 값을 "측정했음" 표시로 바꾼 사본. 이 기기에 값이 확인된 것만 바꾸고,
// 확인되지 않은 값이 하나라도 있으면 null(서버를 고치지 않음).
export function stripServerHealth(code, serverRecords) {
  const local = loadHealthStore(code).records;
  const next = { ...(serverRecords || {}) };
  for (const [k, entry] of Object.entries(serverRecords || {})) {
    if (!isHealthRecordKey(k) || !hasHealthValue(entry)) continue;
    if (!hasHealthValue(local[k])) return null;
    next[k] = toMarker(local[k]);
  }
  return next;
}

/* ---------- 변경 이력(감사 로그) ---------- */
// 건강 종목의 변경 이력은 "누가 언제 어떤 학생의 BMI를 바꿨다"만 서버에 남기고, 바꾼 값
// (prevValue·newValue·prevParts)은 이 기기에 hid(이력 id)로 따로 둔다.

function newHid() {
  return "h_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function trimAudit(audit) {
  const ids = Object.keys(audit);
  if (ids.length <= AUDIT_KEEP) return audit;
  const keep = ids.sort((a, b) => (audit[a].ts || 0) - (audit[b].ts || 0)).slice(ids.length - AUDIT_KEEP);
  return Object.fromEntries(keep.map(id => [id, audit[id]]));
}
function auditHasValues(entry) {
  return !!entry && isHealthEventId(entry.eventId) && ("prevValue" in entry || "newValue" in entry || "prevParts" in entry);
}

// 서버로 보낼 변경 이력 한 건. 건강 종목이면 값을 이 기기에 두고 값 없는 사본을 돌려준다.
export function splitHealthAuditEntry(code, entry) {
  if (!auditHasValues(entry)) return entry;
  const { prevValue, newValue, prevParts, ...rest } = entry;
  const hid = entry.hid || newHid();
  const store = loadHealthStore(code);
  store.audit = trimAudit({ ...store.audit, [hid]: { prevValue: prevValue ?? null, newValue: newValue ?? null, prevParts: prevParts ?? null, ts: entry.ts || Date.now() } });
  saveHealthStore(code, store);
  return { ...rest, healthLocal: true, hid, cleared: newValue === null || newValue === undefined };
}

// 서버에서 받은 변경 이력에 이 기기에 있는 값을 다시 붙인다.
export function mergeHealthAudit(code, list) {
  const audit = loadHealthStore(code).audit;
  return (list || []).map(e => (e && e.healthLocal && e.hid && audit[e.hid])
    ? { ...e, prevValue: audit[e.hid].prevValue, newValue: audit[e.hid].newValue, prevParts: audit[e.hid].prevParts }
    : e);
}

// 서버 변경 이력에 아직 값이 남아 있는 건강 종목 항목이 있는지
export function auditHasServerHealth(list) {
  return (list || []).some(e => auditHasValues(e) && !e.healthLocal);
}
// 서버 변경 이력의 건강 값들을 이 기기로 옮긴 사본(저장·확인 실패 시 null)
export function stripHealthAudit(code, list) {
  const store = loadHealthStore(code);
  const audit = { ...store.audit };
  const next = (list || []).map(e => {
    if (!auditHasValues(e) || e.healthLocal) return e;
    const { prevValue, newValue, prevParts, ...rest } = e;
    const hid = newHid();
    audit[hid] = { prevValue: prevValue ?? null, newValue: newValue ?? null, prevParts: prevParts ?? null, ts: e.ts || 0 };
    return { ...rest, healthLocal: true, hid, cleared: newValue === null || newValue === undefined };
  });
  const trimmed = trimAudit(audit);
  if (!saveHealthStore(code, { ...store, audit: trimmed })) return null;
  const check = loadHealthStore(code).audit;
  const allThere = Object.keys(trimmed).every(hid => !!check[hid]);
  return allThere ? next : null;
}

/* ---------- 파일로 옮기기(다른 학교 PC로) ---------- */

export const HEALTH_FILE_KIND = "smart-paps-health";

export function buildHealthExport(code) {
  const store = loadHealthStore(code);
  return {
    kind: HEALTH_FILE_KIND,
    v: STORE_VERSION,
    code,
    exportedAt: Date.now(),
    records: store.records,
    audit: store.audit,
  };
}

// 파일 내용을 이 기기 저장소에 합친다(같은 기록이면 더 최신 값). 결과: { ok, added, updated, reason }
export function importHealthFile(code, payload) {
  if (!payload || payload.kind !== HEALTH_FILE_KIND || typeof payload.records !== "object" || !payload.records) {
    return { ok: false, reason: "bad-file" };
  }
  if (payload.code !== code) return { ok: false, reason: "other-code" };
  const store = loadHealthStore(code);
  const records = { ...store.records };
  let added = 0, updated = 0;
  Object.entries(payload.records).forEach(([k, entry]) => {
    if (!isHealthRecordKey(k) || !hasHealthValue(entry)) return;
    const mine = records[k];
    if (!hasHealthValue(mine)) { records[k] = entry; added++; }
    else if ((mine.updatedAt || 0) < (entry.updatedAt || 0)) { records[k] = entry; updated++; }
  });
  const audit = trimAudit({ ...(payload.audit && typeof payload.audit === "object" ? payload.audit : {}), ...store.audit });
  if (!saveHealthStore(code, { ...store, records, audit })) return { ok: false, reason: "save-failed" };
  return { ok: true, added, updated };
}
