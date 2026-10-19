// Firebase 초기화.
//
// 이 설정값(firebaseConfig)은 선생님이 Firebase 콘솔에서 만든 "smart-paps" 프로젝트의
// 웹 앱 설정값입니다. Firebase의 웹 SDK 설정값은 원래 브라우저 코드(누구나 페이지 소스로
// 볼 수 있는 곳)에 그대로 들어가도록 설계된 "공개" 값이라 노출 자체는 문제가 아닙니다.
// 실제 접근 통제는 Firebase 익명 로그인(아래 ensureSignedIn)과 Firestore "보안 규칙"
// (firestore.rules, 이 프로젝트에 함께 포함)이 함께 담당합니다 — 반드시 그 규칙도 Firebase 콘솔에 함께 배포해 주세요. 자세한 절차는
// README.md의 "Firebase 설정" 섹션을 참고하세요.
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth, signInAnonymously } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyDAEWEPOQJ2yp7vFGx4v9BUouRqJVWPQmQ",
  authDomain: "smart-paps.firebaseapp.com",
  projectId: "smart-paps",
  storageBucket: "smart-paps.firebasestorage.app",
  messagingSenderId: "17788001226",
  appId: "1:17788001226:web:87eb3637c8cae987cacd33",
  // measurementId(Google 애널리틱스)는 넣지 않는다 — 학생·교사 사용 기록을 외부 분석 서비스로 보내지 않기 위해서다.
};

export const firebaseApp = initializeApp(firebaseConfig);
export const db = getFirestore(firebaseApp);
export const auth = getAuth(firebaseApp);

// 익명 로그인: 이름·이메일 없이 이 기기(브라우저)에 고유번호(uid)만 붙인다.
// 보안 규칙이 이 uid로 개설자·수정 권한자·조회자를 구분한다. 같은 브라우저에서는 uid가
// 유지되고(빔프로젝터 새 창도 같은 uid), 브라우저 데이터를 지우면 새 uid가 생긴다.
let readyPromise = null;
export function ensureSignedIn() {
  if (!readyPromise) {
    readyPromise = (async () => {
      await auth.authStateReady(); // 저장돼 있던 로그인 상태를 먼저 불러온다(새 uid가 생기지 않게)
      if (auth.currentUser) return auth.currentUser;
      const cred = await signInAnonymously(auth);
      return cred.user;
    })().catch((e) => {
      readyPromise = null;
      throw e;
    });
  }
  return readyPromise;
}
