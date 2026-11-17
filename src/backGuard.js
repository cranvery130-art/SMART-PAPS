// 휴대폰·태블릿 뒤로가기(갤럭시 뒤로 버튼·제스처, 아이폰 사파리 스와이프)를 앱 안에서 처리해요.
// 열린 창 닫기 → 바로 전 화면(탭) → 맨 처음 화면에서는 '한 번 더 누르면 나가요' 뒤 2초 안 두 번째에 나가요.
// 양쌤스튜디오 앱들과 같은 방식: 주소 기록에 표시 한 칸(__bk)을 얹어 두고, 뒤로가기로 그 칸이 빠지면(popstate)
// 맨 위 '층'의 할 일을 하고 다시 얹어요. 크롬은 사용자가 누르지 않은 상태에서 쌓은 기록을 건너뛰므로 첫 터치 때 얹어요.
// 쓰는 법: useBackLayer(열려 있는지, () => { 닫기; return true }) / useTabBack(tab, setTab)
import { useEffect, useRef, useState } from "react";

const layers = []; // 나중에 열린 것이 위
let started = false;

function startBackGuard() {
  if (started || typeof window === "undefined") return;
  started = true;
  let on = false, armedOnce = false, rootAt = 0, tEl = null, tT = 0;
  const arm = () => {
    if (on) return;
    try { history.pushState(Object.assign({}, history.state || {}, { __bk: 1 }), "", location.href); on = true; armedOnce = true; } catch (_) { /* 무시 */ }
  };
  const toast = (msg) => {
    if (!tEl) {
      tEl = document.createElement("div"); tEl.setAttribute("role", "status");
      tEl.style.cssText = "position:fixed;left:50%;bottom:calc(24px + env(safe-area-inset-bottom,0px));transform:translate(-50%,8px);z-index:2147483000;max-width:calc(100vw - 32px);padding:10px 16px;border-radius:999px;background:rgba(20,22,28,.92);color:#fff;font-weight:600;font-size:14px;line-height:1.3;font-family:inherit;box-shadow:0 6px 20px rgba(0,0,0,.25);opacity:0;pointer-events:none;transition:opacity .2s ease-out,transform .2s ease-out;white-space:nowrap";
      document.body.appendChild(tEl);
    }
    tEl.textContent = msg;
    requestAnimationFrame(() => { tEl.style.opacity = "1"; tEl.style.transform = "translate(-50%,0)"; });
    clearTimeout(tT); tT = setTimeout(() => { tEl.style.opacity = "0"; tEl.style.transform = "translate(-50%,8px)"; }, 1800);
  };
  ["pointerup", "touchend", "keydown", "click"].forEach((ev) => window.addEventListener(ev, arm, true));
  window.addEventListener("popstate", (e) => {
    if (e.state && e.state.__bk) { on = true; return; } // 앞으로가기로 표시 칸에 다시 올라온 경우
    on = false; if (!armedOnce) return;
    let done = false;
    for (let i = layers.length - 1; i >= 0 && !done; i--) {
      try { done = layers[i].fn.current() !== false; } catch (err) { console.error(err); }
    }
    if (done) { rootAt = 0; arm(); return; }
    const now = Date.now();
    if (now - rootAt < 2000) { rootAt = 0; history.back(); return; } // 2초 안에 한 번 더: 나가기
    rootAt = now; toast("한 번 더 누르면 나가요"); arm();
  });
}

// active인 동안 뒤로가기를 받아요. onBack이 false를 돌려주면 아래 층으로 넘겨요.
export function useBackLayer(active, onBack) {
  const fn = useRef(onBack);
  fn.current = onBack;
  useEffect(() => {
    startBackGuard();
    if (!active) return undefined;
    const layer = { fn };
    layers.push(layer);
    return () => { const i = layers.indexOf(layer); if (i >= 0) layers.splice(i, 1); };
  }, [active]);
}

// 탭을 옮겨 다닌 순서를 기억했다가 뒤로가기로 바로 전 탭에 돌아가요.
export function useTabBack(tab, setTab) {
  const hist = useRef([]);
  const prev = useRef(tab);
  const restoring = useRef(false);
  const [n, setN] = useState(0);
  useEffect(() => {
    if (prev.current === tab) return;
    if (!restoring.current) {
      const i = hist.current.lastIndexOf(tab);
      if (i >= 0) hist.current.splice(i); // 갔던 탭으로 돌아오면 그 뒤 기록은 지워요
      hist.current.push(prev.current);
      if (hist.current.length > 20) hist.current.shift();
    }
    restoring.current = false; prev.current = tab; setN(hist.current.length);
  }, [tab]);
  useBackLayer(n > 0, () => {
    const t = hist.current.pop(); setN(hist.current.length);
    if (t == null) return false;
    restoring.current = true; setTab(t);
    try { window.scrollTo(0, 0); } catch (_) { /* 무시 */ }
    return true;
  });
}

export { startBackGuard };
