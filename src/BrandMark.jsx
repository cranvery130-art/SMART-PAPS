import React from "react";

// AI:PAPS 글자 로고. 'AI'는 소리 나는 대로 읽으면 '아이'라서, 글자 아래에 작게 '아이 · 팝스'를
// 붙여 '아이(학생)들의 PAPS'로 읽히게 한다. 이 앱에 인공지능 기능은 없으니, 이름을 설명하는
// 문구에서도 AI 기능이 있다고 말하지 않는다.
//  - reading: 글자 아래 '아이'·'팝스' 읽는 법을 보여 줄지
//  - dropFrom: 숫자를 주면 글자가 위에서 하나씩 떨어지는 연출(시작 화면용, 초 단위 시작 시각)
export const APP_NAME = "AI:PAPS";
export const APP_READING = "아이팝스";

function Letters({ text, start, step }) {
  if (start == null) return text;
  return text.split("").map((ch, i) => (
    <span key={i} className="aipaps-letter" style={{ animationDelay: (start + i * step) + "s" }}>{ch}</span>
  ));
}

export default function BrandMark({ reading = false, dropFrom, className = "" }) {
  const step = 0.04;
  const at = (n) => (dropFrom == null ? null : dropFrom + n * step);
  const readStyle = dropFrom == null ? undefined : { animationDelay: (dropFrom + 0.4) + "s" };
  return (
    <span className={"aipaps" + (reading ? " has-reading" : "") + (dropFrom != null ? " is-drop" : "") + (className ? " " + className : "")}
      role="img" aria-label={APP_NAME + " (" + APP_READING + ")"}>
      <span className="aipaps-part aipaps-ai" aria-hidden="true">
        <b><Letters text="AI" start={at(0)} step={step} /></b>
        {reading && <small style={readStyle}>아이</small>}
      </span>
      <span className="aipaps-colon" aria-hidden="true">
        <span className={dropFrom != null ? "aipaps-letter" : undefined} style={dropFrom != null ? { animationDelay: at(2) + "s" } : undefined}>
          <i /><i />
        </span>
      </span>
      <span className="aipaps-part aipaps-paps" aria-hidden="true">
        <b><Letters text="PAPS" start={at(3)} step={step} /></b>
        {reading && <small style={readStyle}>팝스</small>}
      </span>
      <style>{BRAND_CSS}</style>
    </span>
  );
}

const BRAND_CSS = `
.aipaps { display: inline-flex; align-items: flex-start; font-family: 'Jua', 'IBM Plex Sans KR', 'Apple SD Gothic Neo', sans-serif;
  font-weight: 400; line-height: 1; letter-spacing: 0.02em; white-space: nowrap; vertical-align: middle; }
.aipaps-part { display: inline-flex; flex-direction: column; align-items: center; }
.aipaps-part b { font-weight: inherit; display: inline-flex; }
.aipaps-ai b { color: #E2622B; }
.aipaps-part small { font-size: 0.3em; letter-spacing: 0.14em; margin-top: 0.32em; padding-left: 0.14em; color: inherit; opacity: 0.72; font-weight: 400; }
.aipaps-ai small { color: #E2622B; opacity: 0.9; }
.aipaps-colon { display: inline-flex; align-items: center; height: 1em; margin: 0 0.07em 0 0.09em; }
.aipaps-colon > span { display: inline-flex; flex-direction: column; gap: 0.2em; padding-top: 0.06em; }
.aipaps-colon i { display: block; width: 0.16em; height: 0.16em; border-radius: 50%; background: #E2622B; }
.aipaps-letter { display: inline-block; opacity: 0; animation: aipapsDrop 0.4s cubic-bezier(0.22, 1, 0.36, 1) forwards; }
.aipaps.is-drop .aipaps-part small { opacity: 0; animation: aipapsFade 0.35s cubic-bezier(0.22, 1, 0.36, 1) forwards; }
@keyframes aipapsDrop { 0% { opacity: 0; transform: translateY(-14px); } 100% { opacity: 1; transform: none; } }
@keyframes aipapsFade { 0% { opacity: 0; transform: translateY(4px); } 100% { opacity: 0.8; transform: none; } }
@media (prefers-reduced-motion: reduce) {
  .aipaps-letter, .aipaps.is-drop .aipaps-part small { animation: none; opacity: 1; transform: none; }
}
`;
