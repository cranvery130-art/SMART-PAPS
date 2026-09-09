// 홈 화면(PWA) 설치를 위한 작은 헬퍼.
//
// 안드로이드(크롬 계열)는 브라우저가 "beforeinstallprompt" 이벤트를 보내주면 그걸 잡아뒀다가
// 버튼 클릭 시 바로 설치 창을 띄울 수 있다. 아이폰(Safari)은 애플 정책상 이 이벤트 자체가
// 없어서 원클릭 설치가 불가능하므로, "공유 → 홈 화면에 추가"를 직접 누르라는 안내만 보여준다.
import { useEffect, useState } from "react";

let deferredPrompt = null;
const listeners = new Set();

if (typeof window !== "undefined") {
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    deferredPrompt = e;
    listeners.forEach((fn) => fn(true));
  });
  window.addEventListener("appinstalled", () => {
    deferredPrompt = null;
    listeners.forEach((fn) => fn(false));
  });
}

function isIOSDevice() {
  if (typeof navigator === "undefined") return false;
  const ua = navigator.userAgent || "";
  const isIOS = /iPad|iPhone|iPod/.test(ua) || (ua.includes("Macintosh") && navigator.maxTouchPoints > 1);
  return isIOS;
}

function isStandalone() {
  if (typeof window === "undefined") return false;
  return (
    window.matchMedia?.("(display-mode: standalone)")?.matches ||
    window.navigator.standalone === true
  );
}

export function usePwaInstall() {
  const [canInstall, setCanInstall] = useState(!!deferredPrompt);

  useEffect(() => {
    const fn = (v) => setCanInstall(v);
    listeners.add(fn);
    return () => listeners.delete(fn);
  }, []);

  async function promptInstall() {
    if (!deferredPrompt) return { outcome: "unavailable" };
    deferredPrompt.prompt();
    const choice = await deferredPrompt.userChoice;
    deferredPrompt = null;
    setCanInstall(false);
    return choice;
  }

  return {
    canInstall,
    promptInstall,
    isIOS: isIOSDevice(),
    isStandalone: isStandalone(),
  };
}
