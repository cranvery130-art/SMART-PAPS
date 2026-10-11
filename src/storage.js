// 기존 Claude 아티팩트가 쓰던 window.storage.get/set/delete(key, shared) 형태의 저장소를
// 그대로 흉내 내는 얇은 대체 계층(shim)이다. App.jsx의 나머지 코드(스토리지 I/O 섹션 제외)는
// 이 인터페이스가 어디에 실제로 저장되는지 전혀 몰라도 되도록, 원래와 똑같은 모양
// ({ value } 반환, key/value/shared 인자)을 유지한다.
//
//  - shared === true  → 여러 사람(학교 코드로 접속하는 모든 기기)이 함께 보는 데이터.
//                        Firestore의 "kv" 컬렉션에 문서 하나(문서 ID = 원래 저장 키)로 저장한다.
//  - shared === false → "이 브라우저(기기)에만" 남아야 하는 데이터(워크스페이스 코드 기억,
//                        기기별 접근 역할, 비밀번호 시도 횟수 등). 서버로 보내지 않고
//                        localStorage에만 저장한다.
import { db } from "./firebase";
import { doc, getDoc, setDoc, deleteDoc } from "firebase/firestore";

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

export const storage = {
  async get(key, shared) {
    if (isPracticeKey(key)) return safeLocalGet(PRACTICE_PREFIX + key);
    if (!shared) return safeLocalGet(key);
    try {
      const ref = doc(db, KV_COLLECTION, sanitizeDocId(key));
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
      const ref = doc(db, KV_COLLECTION, sanitizeDocId(key));
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
      const ref = doc(db, KV_COLLECTION, sanitizeDocId(key));
      await deleteDoc(ref);
      return true;
    } catch (e) {
      return false;
    }
  },
};
