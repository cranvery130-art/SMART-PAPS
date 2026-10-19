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
  // measurementId(Google 애널리틱스)는 넣지 않는다 — 학생·교사 사용 기록을 외부 분석 서비스로 보내지 않기 위해서다.
};

export const firebaseApp = initializeApp(firebaseConfig);
export const db = getFirestore(firebaseApp);
