// Firebase 초기화.
//
// 이 설정값(firebaseConfig)은 선생님이 Firebase 콘솔에서 만든 "smart-paps" 프로젝트의
// 웹 앱 설정값입니다. Firebase의 웹 SDK 설정값은 원래 브라우저 코드(누구나 페이지 소스로
// 볼 수 있는 곳)에 그대로 들어가도록 설계된 "공개" 값이라 노출 자체는 문제가 아닙니다.
// 실제 접근 통제는 Firestore "보안 규칙"(firestore.rules, 이 프로젝트에 함께 포함)이
// 담당합니다 — 반드시 그 규칙도 Firebase 콘솔에 함께 배포해 주세요. 자세한 절차는
// README.md의 "Firebase 설정" 섹션을 참고하세요.
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDAEWEPOQJ2yp7vFGx4v9BUouRqJVWPQmQ",
  authDomain: "smart-paps.firebaseapp.com",
  projectId: "smart-paps",
  storageBucket: "smart-paps.firebasestorage.app",
  messagingSenderId: "17788001226",
  appId: "1:17788001226:web:87eb3637c8cae987cacd33",
  measurementId: "G-LHB4B661ZP",
};

export const firebaseApp = initializeApp(firebaseConfig);
export const db = getFirestore(firebaseApp);

// Analytics는 브라우저 환경에서만, 그리고 선택적으로만 초기화한다(측정 기능이 없어도
// 프로그램 동작에는 전혀 영향이 없으므로 실패해도 조용히 무시한다).
export async function initAnalyticsIfSupported() {
  try {
    const { getAnalytics, isSupported } = await import("firebase/analytics");
    if (await isSupported()) {
      getAnalytics(firebaseApp);
    }
  } catch (e) {
    // 분석 기능은 핵심 기능이 아니므로 실패해도 무시한다.
  }
}
