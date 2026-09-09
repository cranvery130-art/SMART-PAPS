import React from "react";
import ReactDOM from "react-dom/client";
import PapsApp from "./App.jsx";
import { initAnalyticsIfSupported } from "./firebase.js";

// 빔프로젝터 고정모드에서 연 "새 창"은 주소창에 ?present=1&code=학교코드 형태로 열린다.
// (App.jsx의 onEnterPresentation이 이 주소를 만들어 새 창을 연다.) 이 값을 읽어서
// PapsApp에 "시작하자마자 이 코드로, 전광판 모드로 바로 들어가라"고 알려준다.
const params = new URLSearchParams(window.location.search);
const presentParam = params.get("present") === "1";
const codeParam = params.get("code") || null;

initAnalyticsIfSupported();

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <PapsApp
      initialWorkspaceCode={presentParam ? codeParam : null}
      forcePresentation={presentParam}
    />
  </React.StrictMode>
);

// PWA 설치 조건(설치 가능한 앱)을 만족시키기 위한 서비스워커 등록.
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("/sw.js").catch(() => {
      // 서비스워커 등록 실패는 핵심 기능에 영향이 없으므로 조용히 무시한다.
    });
  });
}
