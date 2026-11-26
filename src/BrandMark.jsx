import React from "react";

// AI:PAPS 로고. 'AI'는 소리 나는 대로 읽으면 '아이'라서, A는 달려 나가는 아이(벌린 다리 = A의 두 획,
// 흔드는 팔 = 가로획), i는 점 자리가 초시계인 모양으로 그렸어요. 두 글자는 PAPS 글자(Jua)의
// 대문자 높이·획 굵기에 맞춰 그려서 글자 줄에 튀지 않게 섞여요. 이 앱에 인공지능 기능은 없으니,
// 이름을 설명하는 문구에서도 AI 기능이 있다고 말하지 않아요.
//  - reading: 글자 아래 '아이'·'팝스' 읽는 법을 보여 줄지
//  - dropFrom: 숫자를 주면 글자가 위에서 하나씩 떨어지는 연출(시작 화면용, 초 단위 시작 시각)
// 그림 좌표는 글꼴 단위(1em = 1000, 대문자 높이 712)예요. 앱 아이콘(design/app-icon)도 같은 좌표를 써요.
export const APP_NAME = "AI:PAPS";
export const APP_READING = "아이팝스";

const ACCENT = "#E2622B";

export function AiMark({ className = "", style }) {
  return (
    <svg className={"aipaps-mark " + className} style={style} viewBox="0 0 1040 712" aria-hidden="true" focusable="false">
      {/* A: 달리는 아이(벌린 다리 = A의 두 획, 흔드는 팔 = 가로획, 꼭짓점 위 머리) */}
      <g fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round">
        <path d="M92 646 L372 254 L640 646" strokeWidth="136" />
        <path d="M196 482 L566 420" strokeWidth="104" />
      </g>
      <circle cx="404" cy="86" r="80" fill={ACCENT} />
      {/* i: 초시계 점 + 줄기 */}
      <path d="M885 418 V646" stroke="currentColor" strokeWidth="140" strokeLinecap="round" />
      <rect x="858" y="0" width="54" height="52" rx="14" fill="currentColor" />
      <circle cx="885" cy="178" r="116" fill="currentColor" />
      <path d="M885 178 L885 92 A86 86 0 0 1 962 216 Z" fill={ACCENT} />
    </svg>
  );
}

function Letters({ text, start, step }) {
  if (start == null) return text;
  return text.split("").map((ch, i) => (
    <span key={i} className={"aipaps-letter" + (ch === ":" ? " aipaps-colon" : "")} style={{ animationDelay: (start + i * step) + "s" }}>{ch}</span>
  ));
}

export default function BrandMark({ reading = false, dropFrom, className = "" }) {
  const step = 0.04;
  const drop = dropFrom != null;
  const readStyle = drop ? { animationDelay: (dropFrom + 0.4) + "s" } : undefined;
  return (
    <span className={"aipaps" + (reading ? " has-reading" : "") + (drop ? " is-drop" : "") + (className ? " " + className : "")}
      role="img" aria-label={APP_NAME + " (" + APP_READING + ")"}>
      <span className={"aipaps-ai" + (drop ? " aipaps-letter" : "")} style={drop ? { animationDelay: dropFrom + "s" } : undefined} aria-hidden="true">
        <AiMark />
      </span>
      <b className="aipaps-paps" aria-hidden="true">
        {drop ? <Letters text=":PAPS" start={dropFrom + step * 2} step={step} /> : <><span className="aipaps-colon">:</span>PAPS</>}
      </b>
      {reading && <small className="aipaps-read r1" style={readStyle} aria-hidden="true">아이</small>}
      {reading && <small className="aipaps-read r2" style={readStyle} aria-hidden="true">팝스</small>}
      <style>{BRAND_CSS}</style>
    </span>
  );
}

// 글자 줄: line-height 1이면 Jua의 기준선은 아래에서 0.2em 위라서, 그림 아래쪽을 0.2em 띄워 PAPS와 같은 기준선에 세워요.
const BRAND_CSS = `
.aipaps { display: inline-grid; grid-template-columns: auto auto; align-items: end; justify-items: center;
  font-family: 'Jua', 'IBM Plex Sans KR', 'Apple SD Gothic Neo', sans-serif; font-weight: 400; line-height: 1;
  letter-spacing: 0.02em; white-space: nowrap; vertical-align: middle; color: var(--text, #1D2747); }
.aipaps-ai { display: inline-flex; margin-bottom: 0.2em; }
.aipaps-mark { display: block; height: 0.712em; width: auto; overflow: visible; }
.aipaps-paps { font-weight: inherit; display: inline-flex; justify-self: start; }
.aipaps-colon { color: #E2622B; margin: 0 0.02em 0 0.04em; }
.aipaps-read { font-size: 0.3em; letter-spacing: 0.14em; margin-top: 0.3em; padding-left: 0.14em; opacity: 0.72; font-weight: 400; }
.aipaps-read.r1 { grid-column: 1; grid-row: 2; color: #E2622B; opacity: 0.9; }
.aipaps-read.r2 { grid-column: 2; grid-row: 2; justify-self: center; padding-left: 0.6em; }
.aipaps-letter { display: inline-block; opacity: 0; animation: aipapsDrop 0.4s cubic-bezier(0.22, 1, 0.36, 1) forwards; }
.aipaps.is-drop .aipaps-read { opacity: 0; animation: aipapsFade 0.35s cubic-bezier(0.22, 1, 0.36, 1) forwards; }
@keyframes aipapsDrop { 0% { opacity: 0; transform: translateY(-14px); } 100% { opacity: 1; transform: none; } }
@keyframes aipapsFade { 0% { opacity: 0; transform: translateY(4px); } 100% { opacity: 0.8; transform: none; } }
@media (prefers-reduced-motion: reduce) {
  .aipaps-letter, .aipaps.is-drop .aipaps-read { animation: none; opacity: 1; transform: none; }
}
`;
