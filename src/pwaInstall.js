// 홈 화면(PWA) 설치를 위한 작은 헬퍼.
//
// 안드로이드(크롬·삼성 인터넷·엣지)는 브라우저가 "beforeinstallprompt" 이벤트를 보내주면 그걸 잡아뒀다가
// 버튼 클릭 시 바로 설치 창을 띄울 수 있다(한 번 눌러 승인). 아이폰·아이패드(Safari)는 애플 정책상
// 이 이벤트 자체가 없어서 버튼 설치가 불가능하므로, "공유 → 홈 화면에 추가" 순서를 안내한다.
// 카카오톡·네이버 앱 같은 앱 안 브라우저에서는 설치가 안 되므로 바깥 브라우저로 열라고 안내한다.
import { useEffect, useState } from "react";

let deferredPrompt = null;
let installedNow = false; // 이번 방문에 설치를 마쳤는지(appinstalled)
const listeners = new Set();
const notify = () => listeners.forEach((fn) => fn());

if (typeof window !== "undefined") {
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    deferredPrompt = e;
    notify();
  });
  window.addEventListener("appinstalled", () => {
    deferredPrompt = null;
    installedNow = true;
    notify();
  });
}

function ua() {
  return typeof navigator === "undefined" ? "" : navigator.userAgent || "";
}

function isIOSDevice() {
  if (typeof navigator === "undefined") return false;
  const u = ua();
  return /iPad|iPhone|iPod/.test(u) || (u.includes("Macintosh") && navigator.maxTouchPoints > 1);
}

function isAndroidDevice() {
  return /Android/i.test(ua());
}

// 카카오톡·네이버·인스타그램·페이스북·라인·밴드 등 앱 안에서 열린 브라우저
function isInAppBrowser() {
  return /KAKAOTALK|NAVER\(inapp|Instagram|FBAN|FBAV|Line\/|BAND\/|everytimeApp|DaumApps/i.test(ua());
}

// 아이폰에서 Safari가 아닌 브라우저(크롬·파이어폭스·엣지 등)
function isIOSNonSafari() {
  const u = ua();
  return isIOSDevice() && (/CriOS|FxiOS|EdgiOS|OPiOS|Whale/i.test(u) || !/Safari/.test(u));
}

function isStandalone() {
  if (typeof window === "undefined") return false;
  return (
    window.matchMedia?.("(display-mode: standalone)")?.matches ||
    window.matchMedia?.("(display-mode: fullscreen)")?.matches ||
    window.navigator.standalone === true
  );
}

function snapshot() {
  return { canInstall: !!deferredPrompt, installed: installedNow || isStandalone() };
}

export function usePwaInstall() {
  const [state, setState] = useState(snapshot);

  useEffect(() => {
    const fn = () => setState(snapshot());
    listeners.add(fn);
    fn(); // 마운트 전에 이벤트가 왔을 수도 있어요
    return () => listeners.delete(fn);
  }, []);

  async function promptInstall() {
    if (!deferredPrompt) return { outcome: "unavailable" };
    const p = deferredPrompt;
    p.prompt();
    const choice = await p.userChoice.catch(() => ({ outcome: "dismissed" }));
    // 설치 창은 한 번만 쓸 수 있어요. 취소하면 브라우저가 다음에 새 이벤트를 다시 보내요.
    deferredPrompt = null;
    if (choice?.outcome === "accepted") installedNow = true;
    notify();
    return choice;
  }

  return {
    canInstall: state.canInstall,
    installed: state.installed,
    promptInstall,
    isIOS: isIOSDevice(),
    isAndroid: isAndroidDevice(),
    isInApp: isInAppBrowser(),
    isIOSNonSafari: isIOSNonSafari(),
    // 예전 이름 그대로 쓰는 곳을 위해 남겨요.
    isStandalone: state.installed,
  };
}

// 브라우저별 "홈 화면에 추가" 순서(버튼 설치가 안 될 때 보여줘요). 각 줄은 짧은 문장.
export function installSteps({ isIOS, isAndroid, isInApp, isIOSNonSafari }) {
  if (isInApp) {
    return isIOS
      ? ["카카오톡·네이버 같은 앱 안에서는 설치가 안 돼요.", "오른쪽 아래(또는 위) ⋯ 메뉴에서 'Safari로 열기'를 눌러요.", "Safari에서 공유 버튼 ⬆︎ → '홈 화면에 추가' → '추가'를 눌러요."]
      : ["카카오톡·네이버 같은 앱 안에서는 설치가 안 돼요.", "오른쪽 위 ⋮ 메뉴에서 '다른 브라우저로 열기'(크롬·삼성 인터넷)를 눌러요.", "새로 열린 화면에서 '홈 화면에 설치'를 다시 눌러요."];
  }
  if (isIOSNonSafari) {
    return ["아이폰·아이패드는 Safari에서 설치할 수 있어요.", "이 주소를 Safari로 열어 주세요.", "공유 버튼 ⬆︎ → '홈 화면에 추가' → '추가'를 눌러요."];
  }
  if (isIOS) {
    return ["화면 아래(아이패드는 위) 공유 버튼 ⬆︎을 눌러요.", "목록을 내려 '홈 화면에 추가'를 눌러요.", "오른쪽 위 '추가'를 누르면 홈 화면에 SMART PAPS 아이콘이 생겨요."];
  }
  if (isAndroid) {
    return ["오른쪽 위 ⋮ 메뉴(삼성 인터넷은 아래 ≡ 메뉴)를 눌러요.", "'앱 설치' 또는 '홈 화면에 추가'를 눌러요.", "'설치'를 누르면 홈 화면에 SMART PAPS 아이콘이 생겨요."];
  }
  return ["크롬·엣지 주소창 오른쪽의 설치 아이콘(⊕ 또는 모니터 모양)을 눌러요.", "아이콘이 없으면 ⋮ 메뉴 → '앱 설치'(엣지는 '앱 → 이 사이트를 앱으로 설치')를 눌러요.", "바탕화면·시작 메뉴에서 SMART PAPS를 바로 열 수 있어요."];
}
