import React, { useState } from "react";
import {
  BookOpen, X, ChevronLeft, ChevronRight, PlayCircle, ExternalLink,
  Users, Plus, Database, ShieldCheck, Trash2, Info,
} from "lucide-react";

// 사용 가이드: 예전 "사용설명서"·"기능설명"·"상세설명"을 하나로 합친 화면.
//  - 첫 화면(로그인 전)에서는 모달(GuideModal)로, 프로그램 안에서는 "사용 가이드" 탭(GuidePanel)으로 보인다.
//  - "따라하기"는 실제 화면을 본뜬 작은 데모 화면에 번호 표시(①②③)를 붙이고, 같은 번호로
//    해야 할 일을 짧게 적는다. 처음 쓰는 선생님은 물론 중학생도 순서대로 따라 할 수 있게 한다.
// 안내하는 동작이 바뀌면(탭 이름, 버튼 이름, 저장 방식 등) 이 파일의 문구도 함께 고쳐야 한다.

/* ---------------- 데모 화면 조각들 ---------------- */

// 번호 표시를 붙일 대상. hl이면 눈에 띄게 테두리가 깜박인다.
function T({ n, hl = true, children, block }) {
  return (
    <span className={"gd-target" + (hl ? " gd-hl" : "") + (block ? " gd-block" : "")}>
      {children}
      {n && <i className="gd-pin" aria-hidden="true">{n}</i>}
    </span>
  );
}
function MiniTabs({ active, n }) {
  const tabs = ["전광판", "학생관리", "기록관리", "등급표", "데이터 백업", "접근권한", "마감"];
  return (
    <div className="gd-tabs">
      {tabs.map(t => t === active
        ? <T key={t} n={n}><span className="gd-tab on">{t}</span></T>
        : <span key={t} className="gd-tab">{t}</span>)}
    </div>
  );
}
const Btn = ({ children, primary }) => <span className={"gd-btn" + (primary ? " primary" : "")}>{children}</span>;
const Field = ({ label, value }) => (
  <span className="gd-field"><small>{label}</small><span>{value}</span></span>
);
const Chip = ({ children, on }) => <span className={"gd-chip" + (on ? " on" : "")}>{children}</span>;

function Screen({ children, label }) {
  return (
    <div className="gd-screen" role="img" aria-label={label}>
      <div className="gd-screen-bar"><span /><span /><span /></div>
      <div className="gd-screen-body">{children}</div>
    </div>
  );
}

/* ---------------- 따라하기 단계 ---------------- */
// who: "all" 누구나 / "admin" 수정 권한 이상 / "founder" 개설자만. tab: "이 화면 열기"로 이동할 탭.
const STEPS = [
  {
    title: "학교 코드 만들기",
    who: "처음 한 번, 개설자",
    tab: null,
    actions: [
      "첫 화면에서 \"처음이신가요? 새 코드 만들기\"를 누르세요.",
      "코드 이름과 비밀번호 2개를 정하세요. 학교급(초·중·고)도 고르세요.",
      "\"만들기\"를 누르고, 바로 뜨는 메모 파일을 꼭 저장하세요.",
    ],
    tip: "코드에 학교 이름은 넣지 마세요. 예: 낭만체육123",
    more: [
      "개설자 전용 비밀번호: 나만 알아야 해요. 다른 기기에서 \"코드로 로그인\"할 때 써요.",
      "접근 신청 비밀번호: 동료 선생님이 \"접근 신청\"할 때 써요. 화면 잠금을 풀 때도 써요.",
      "비밀번호를 잊어버리면 되찾을 방법이 없어요. 메모 파일을 안전한 곳에 보관하세요.",
    ],
    Demo: () => (
      <Screen label="새 코드 만들기 화면">
        <div className="gd-center">
          <b className="gd-brand">SMART PAPS</b>
          <T n="1"><Btn>처음이신가요? 새 코드 만들기</Btn></T>
          <T n="2" block>
            <span className="gd-stack">
              <Field label="코드 이름" value="낭만체육123" />
              <Field label="개설자 전용 비밀번호" value="••••••" />
              <Field label="접근 신청 비밀번호" value="••••••" />
              <span className="gd-row"><Chip>초등</Chip><Chip on>중학교</Chip><Chip>고등</Chip></span>
            </span>
          </T>
          <T n="3"><Btn primary>만들기</Btn></T>
        </div>
      </Screen>
    ),
  },
  {
    title: "학생 명단 올리기",
    who: "개설자·수정 권한",
    tab: "roster",
    actions: [
      "\"학생관리\" 탭을 누르세요.",
      "나이스에서 받은 학생명렬 엑셀 파일을 점선 칸에 끌어다 놓으세요. 휴대폰은 칸을 눌러 파일을 고르면 돼요.",
      "아래 명단에서 성별이 맞게 들어갔는지 확인하세요.",
    ],
    tip: "같은 파일을 또 올려도 학생이 두 번 생기지 않아요.",
    more: [
      "학생 이름은 서버로 보내지 않고 이 기기에만 저장돼요. 다른 기기에서는 이름 대신 \"(이름 미확인 - 이 기기)\"로 보일 수 있어요. 그 기기에서 같은 명렬표를 다시 올리면 이름이 채워져요.",
      "성별 열이 없는 파일을 올리면 모두 \"여\"로 등록돼요. 꼭 확인하세요.",
      "전학생 한 명은 \"학생 추가\", 여러 명은 \"일괄 추가\"(한 줄에 1,3,12,홍길동,남)로 넣어요.",
    ],
    Demo: () => (
      <Screen label="학생관리 화면">
        <MiniTabs active="학생관리" n="1" />
        <T n="2" block><span className="gd-drop">엑셀 파일을 여기로 끌어다 놓거나 눌러서 선택</span></T>
        <T n="3" block>
          <span className="gd-list">
            <span>1-1-1 김하늘 <Chip on>남</Chip></span>
            <span>1-1-2 이바다 <Chip>여</Chip></span>
            <span>1-1-3 박가람 <Chip on>남</Chip></span>
          </span>
        </T>
      </Screen>
    ),
  },
  {
    title: "기록 입력하기",
    who: "개설자·수정 권한",
    tab: "records",
    actions: [
      "\"기록관리\" 탭에서 측정할 종목을 고르세요.",
      "학년·반 버튼을 누르세요. 숫자는 \"입력한 학생 / 전체 학생\"이에요.",
      "학생 옆 칸에 기록을 쓰고 Enter를 누르면 바로 저장돼요. 초록 체크가 뜨면 저장된 거예요.",
    ],
    tip: "여러 학생이 화면을 같이 볼 땐 \"이름 가림\"을 켜면 홍*동처럼 가려져요.",
    more: [
      "다른 기기(휴대폰 등)에서 입력해도 몇 초 안에 모든 화면에 함께 반영돼요.",
      "너무 큰 값처럼 오타로 보이는 기록을 넣으면 한 번 더 확인하라는 알림이 떠요.",
      "잘못 넣은 기록은 접근권한 탭의 \"최근 변경 이력\"에서 되돌릴 수 있어요(개설자).",
    ],
    Demo: () => (
      <Screen label="기록관리 화면">
        <MiniTabs active="기록관리" />
        <T n="1" block><span className="gd-row"><Chip on>왕복오래달리기</Chip><Chip>윗몸말아올리기</Chip><Chip>50m달리기</Chip></span></T>
        <T n="2" block><span className="gd-row"><Chip on>1학년 8/12</Chip><Chip>2학년 7/10</Chip><Chip on>1반</Chip><Chip>2반</Chip></span></T>
        <T n="3" block>
          <span className="gd-list">
            <span>1. 김하늘 <span className="gd-input">52</span> <em className="gd-ok">✓</em></span>
            <span>2. 이바다 <span className="gd-input caret">3</span></span>
          </span>
        </T>
      </Screen>
    ),
  },
  {
    title: "측정 도구 쓰기",
    who: "개설자·수정 권한",
    tab: "records",
    actions: [
      "기록관리에서 종목을 고르면 그 종목에 맞는 도구가 함께 나와요.",
      "왕복오래달리기·윗몸말아올리기는 \"재생\"을 누르세요. 횟수가 저절로 올라가요.",
      "학생이 멈추는 순간 그 학생 이름을 누르면, 그때 횟수가 바로 기록돼요.",
    ],
    tip: "50m달리기는 \"그룹 출발\" 후, 결승선에 들어오는 순서대로 그 번호의 \"기록\"을 누르세요.",
    more: [
      "스텝검사는 공식 음원(3분)이 끝나면 의자에 앉아 세 구간 각각 30초 동안 맥박을 세서 입력해요.",
      "오래달리기-걷기에는 운동장 코스 계산기, 반 전체 동시 출발 타이머가 있어요.",
      "종목마다 \"측정 방법 영상\" 버튼으로 교육부 안내 영상을 볼 수 있어요.",
    ],
    Demo: () => (
      <Screen label="측정 도구 화면">
        <T n="1" block><span className="gd-row"><Chip on>왕복오래달리기</Chip><Chip>1학년 1반</Chip></span></T>
        <span className="gd-big"><small>3단계</small>누적 24회</span>
        <T n="2"><Btn primary>▶ 신호음 재생</Btn></T>
        <T n="3" block>
          <span className="gd-row"><Chip>김하늘</Chip><Chip>이바다</Chip><Chip on>박가람</Chip><Chip>최슬기</Chip></span>
        </T>
      </Screen>
    ),
  },
  {
    title: "등급 확인하기",
    who: "누구나",
    tab: "grades",
    actions: [
      "\"등급표\" 탭을 누르세요.",
      "학년·반·성별을 고르면 학생별로 종목 등급이 색깔로 보여요.",
      "등급 기준이 궁금하면 아래 \"참고 등급 기준표\"를 보세요.",
    ],
    tip: "등급은 참고용이에요. 나이스에는 실제 기록 숫자가 들어가요.",
    more: [
      "BMI·체지방률은 등급 없이 마름·정상·과체중 같은 분류로만 보여줘요.",
      "심폐지구력·유연성처럼 비슷한 종목을 여러 개 측정해도 종합 순위에는 1개만 들어가요.",
    ],
    Demo: () => (
      <Screen label="등급표 화면">
        <MiniTabs active="등급표" n="1" />
        <T n="2" block>
          <span className="gd-table">
            <span className="h"><b>학생</b><b>왕복</b><b>윗몸</b><b>50m</b></span>
            <span><b>1-1-1</b><i className="g1">1</i><i className="g2">2</i><i className="g3">3</i></span>
            <span><b>1-1-2</b><i className="g2">2</i><i className="g4">4</i><i className="g2">2</i></span>
          </span>
        </T>
        <T n="3"><Btn>참고 등급 기준표</Btn></T>
      </Screen>
    ),
  },
  {
    title: "전광판 띄우기",
    who: "누구나",
    tab: "board",
    actions: [
      "\"전광판\" 탭에서 종목과 학년을 고르세요. 순위가 실시간으로 바뀌어요.",
      "\"빔프로젝터 고정모드\"를 누르면 잠긴 전광판이 새 창으로 열려요.",
      "그 창을 빔프로젝터 화면으로 옮기세요. 원래 창에서는 계속 다른 일을 할 수 있어요.",
    ],
    tip: "전광판에는 이름 대신 학년-반-번호만 나와요. 창을 닫으려면 접근 신청 비밀번호가 필요해요.",
    more: [
      "새 창이 열리지 않으면 브라우저 주소창 오른쪽의 팝업 차단 표시를 눌러 허용해 주세요.",
    ],
    Demo: () => (
      <Screen label="전광판 화면">
        <span className="gd-row gd-between">
          <T n="1"><span className="gd-row"><Chip on>왕복오래달리기</Chip><Chip>1학년</Chip></span></T>
          <T n="2"><Btn>⛶ 빔프로젝터 고정모드</Btn></T>
        </span>
        <span className="gd-podium">
          <span className="p2"><small>1-2-4</small>71회</span>
          <span className="p1"><small>1-1-3</small>84회</span>
          <span className="p3"><small>2-1-7</small>66회</span>
        </span>
      </Screen>
    ),
  },
  {
    title: "동료 선생님과 함께 쓰기",
    who: "개설자",
    tab: "access",
    actions: [
      "동료 선생님에게 학교 코드와 접근 신청 비밀번호를 알려주세요.",
      "동료는 첫 화면의 \"접근 신청\"으로 들어와요. 조회는 바로 돼요.",
      "기록 수정 권한을 신청하면 \"접근권한\" 탭에 뜨니 \"승인\"을 누르세요.",
    ],
    tip: "개설자 전용 비밀번호는 절대 알려주지 마세요.",
    more: [
      "승인한 선생님도 \"권한 취소\"로 언제든 막을 수 있어요.",
      "담당이 바뀌면 승인된 수정 권한자에게 \"개설자로 지정\"으로 자리를 넘길 수 있어요.",
      "스마트폰에서도 같은 코드로 들어와 운동장에서 바로 입력할 수 있어요. 홈 화면에 앱처럼 설치할 수도 있어요.",
    ],
    Demo: () => (
      <Screen label="접근권한 화면">
        <MiniTabs active="접근권한" />
        <T n="1" block><Field label="접근 신청 비밀번호" value="체육0925" /></T>
        <T n="2" hl={false}><span className="gd-sub">동료 화면: 첫 화면 → 접근 신청 → 이름·코드·비밀번호 입력</span></T>
        <span className="gd-sub">승인 대기 (1명)</span>
        <span className="gd-row gd-between gd-card">
          <span>1학년 2반 담임 박연습 <Chip on>수정 권한</Chip></span>
          <T n="3"><span className="gd-row"><Btn primary>승인</Btn><Btn>거절</Btn></span></T>
        </span>
      </Screen>
    ),
  },
  {
    title: "나이스에 올리기",
    who: "개설자·수정 권한",
    tab: "backup",
    actions: [
      "\"데이터 백업\" 탭을 누르세요.",
      "나이스에서 받은 양식 파일을 첨부하고 \"반영하기\"를 누르세요. 우리 기록이 자동으로 채워져요.",
      "미리보기를 확인하고 \"반영된 파일 다운로드\"를 눌러 나이스에 올리세요.",
    ],
    tip: "나이스 등록이 끝나면 내려받은 파일은 컴퓨터에서 바로 지워 주세요.",
    more: [
      "JSON 백업(내보내기·불러오기)은 개설자만 할 수 있어요.",
    ],
    Demo: () => (
      <Screen label="데이터 백업 화면">
        <MiniTabs active="데이터 백업" n="1" />
        <T n="2" block><span className="gd-drop">나이스양식.xlsx 첨부됨 <Btn primary>반영하기</Btn></span></T>
        <T n="3"><Btn>반영된 파일 다운로드</Btn></T>
      </Screen>
    ),
  },
  {
    title: "학기 마감하기",
    who: "개설자",
    tab: "closeout",
    actions: [
      "나이스 등록이 모두 끝났는지 확인하세요.",
      "\"마감\" 탭에서 \"지금 백업 파일 받기\"를 먼저 누르세요.",
      "\"마감하기\"를 누르면 기록·명단·학교 코드가 모두 지워져요.",
    ],
    tip: "되돌릴 수 없어요. 다음 학기에는 새 코드를 만들어 시작해요.",
    more: [
      "마감을 잊어도 코드를 만든 지 1년이 지나면 자동으로 지워져요. 30일 전부터 알림이 뜨고 연장할 수 있어요.",
    ],
    Demo: () => (
      <Screen label="마감 화면">
        <MiniTabs active="마감" />
        <T n="1" hl={false}><span className="gd-check">☑ 나이스 제출 완료</span></T>
        <T n="2"><Btn>지금 백업 파일 받기</Btn></T>
        <T n="3"><span className="gd-btn danger">마감하기</span></T>
      </Screen>
    ),
  },
];

const FEATURES = [
  { icon: Plus, title: "측정 도구", items: [
    "심박수·부위 점수를 넣으면 지수·합계가 바로 계산돼요.",
    "윗몸말아올리기·스텝검사는 교육부 공식 음원, 왕복오래달리기는 신호음이나 유튜브 영상으로 진행해요.",
    "50m 그룹 타이머, 반 전체 동시 출발 타이머, 운동장 코스 계산기가 있어요.",
  ] },
  { icon: Users, title: "함께 쓰기", items: [
    "조회는 비밀번호만 맞으면 바로, 수정은 개설자가 승인해야 돼요.",
    "누가 언제 어떤 기록을 바꿨는지 남고, 잘못 바꾼 기록은 되돌릴 수 있어요.",
    "휴대폰·노트북 어디서 입력해도 몇 초 안에 모두에게 반영돼요.",
  ] },
  { icon: ShieldCheck, title: "개인정보 보호", items: [
    "학교 코드와 비밀번호를 모두 알아야 들어올 수 있어요.",
    "학생 이름은 서버에 보내지 않고 이 기기에만 저장해요. 전광판엔 학년-반-번호만 나와요.",
    "12분 동안 아무것도 안 하면 화면이 자동으로 잠겨요(공용 PC 보호).",
    "비밀번호를 5번 틀리면 10분 동안 신청이 막혀요.",
  ] },
  { icon: Database, title: "나이스·백업", items: [
    "나이스 학생명렬 엑셀로 명단을 한 번에 등록해요.",
    "나이스 양식 파일에 우리 기록을 자동으로 채워 줘요.",
    "파일을 내려받을 때마다 다 쓰면 지우라는 안내가 떠요.",
  ] },
  { icon: Trash2, title: "보관 기간", items: [
    "기록은 \"마감\"을 누르기 전까지 보관돼요.",
    "마감하면 기록·명단·학교 코드가 모두 바로 지워져요.",
    "마감을 잊어도 1년이 지나면 자동으로 지워져요(30일 전부터 알림, 연장 가능).",
  ] },
];

const DEVICE_NOTES = [
  { title: "여러 기기에서 쓰기", body: "같은 학교 코드로 들어오면 휴대폰·노트북 모두 같은 기록을 봐요. 개설자는 \"코드로 로그인\"에 코드와 개설자 전용 비밀번호를, 동료는 \"접근 신청\"을 쓰면 돼요." },
  { title: "나눠 쓰는 예", body: "노트북은 교무실에서 등급표·나이스 작업, 휴대폰은 운동장에서 기록 입력처럼 나눠 쓰면 편해요." },
  { title: "다시 들어와야 할 때", body: "동료 선생님은 처음과 똑같은 이름, 똑같은 권한 종류로 다시 신청하면 기존 승인이 그대로 이어져요. 이름에 \"2학년 3반\"처럼 구분되는 말을 넣어 주세요." },
  { title: "조심할 점", body: "브라우저 데이터를 지우거나 시크릿 모드로 쓰면 승인 정보와 이 기기의 학생 이름표가 사라져요. 인터넷이 끊긴 동안 입력한 기록은 다시 연결돼야 다른 기기에 보여요. JSON 백업은 개설자 한 명만 맡아요." },
];

/* ---------------- 본문 ---------------- */

function canOpenTab(tab, role, isFounder) {
  if (!tab) return false;
  if (tab === "board" || tab === "grades") return true;
  if (role !== "admin") return false;
  if (tab === "closeout") return !!isFounder;
  return true;
}

export function GuideContent({ inApp = false, role, isFounder, isPractice, onGo, onStartPractice }) {
  const [section, setSection] = useState("steps");
  const [stepIdx, setStepIdx] = useState(0);
  const step = STEPS[stepIdx];
  const Demo = step.Demo;
  const openable = inApp && canOpenTab(step.tab, role, isFounder);

  return (
    <div className="gd">
      {isPractice ? (
        <div className="gd-practice-note">
          <PlayCircle size={16} />
          <span>
            지금은 <b>연습모드</b>예요. 가짜 학생으로 마음껏 눌러 보세요. 각 단계의 "이 화면 열기"를 누르면 그 화면으로 바로 가요.
          </span>
        </div>
      ) : onStartPractice ? (
        <div className="gd-practice-note">
          <PlayCircle size={16} />
          <span>읽기만 하는 것보다 직접 눌러 보는 게 빨라요. 가짜 학생이 들어 있는 연습모드로 실제 화면을 체험해 보세요.</span>
          <button type="button" className="btn btn-primary small" onClick={onStartPractice}>
            {inApp ? <><ExternalLink size={13} /> 연습모드 새 창으로 열기</> : <><PlayCircle size={13} /> 연습모드 체험</>}
          </button>
        </div>
      ) : null}

      <div className="gd-sections" role="tablist" aria-label="사용 가이드 구분">
        {[["steps", "따라하기"], ["features", "기능 한눈에"], ["devices", "여러 기기·주의사항"]].map(([id, label]) => (
          <button key={id} type="button" role="tab" aria-selected={section === id}
            className={"gd-section-btn" + (section === id ? " on" : "")} onClick={() => setSection(id)}>
            {label}
          </button>
        ))}
      </div>

      {section === "steps" && (
        <>
          <div className="gd-stepper" aria-label="단계 선택">
            {STEPS.map((s, i) => (
              <button key={i} type="button" className={"gd-step-dot" + (i === stepIdx ? " on" : "") + (i < stepIdx ? " done" : "")}
                onClick={() => setStepIdx(i)} aria-label={`${i + 1}단계 ${s.title}`} aria-current={i === stepIdx ? "step" : undefined}>
                <span>{i + 1}</span><em>{s.title}</em>
              </button>
            ))}
          </div>

          <div className="gd-card-step">
            <div className="gd-step-head">
              <span className="gd-step-count">{stepIdx + 1} / {STEPS.length}</span>
              <h4>{step.title}</h4>
              <span className="gd-who">{step.who}</span>
            </div>
            <div className="gd-step-grid">
              <div className="gd-demo"><Demo /></div>
              <div className="gd-explain">
                <ol className="gd-actions">
                  {step.actions.map((a, i) => <li key={i}><i className="gd-pin static">{i + 1}</i><span>{a}</span></li>)}
                </ol>
                {step.tip && <div className="gd-tip"><Info size={14} /><span>{step.tip}</span></div>}
                {step.more && step.more.length > 0 && (
                  <details className="gd-more">
                    <summary>더 알아보기</summary>
                    <ul>{step.more.map((m, i) => <li key={i}>{m}</li>)}</ul>
                  </details>
                )}
              </div>
            </div>
            <div className="gd-nav">
              <button type="button" className="btn btn-ghost" disabled={stepIdx === 0} onClick={() => setStepIdx(i => i - 1)}>
                <ChevronLeft size={15} /> 이전
              </button>
              {openable && (
                <button type="button" className="btn btn-secondary" onClick={() => onGo(step.tab)}>
                  이 화면 열기 <ChevronRight size={15} />
                </button>
              )}
              <button type="button" className="btn btn-primary" disabled={stepIdx === STEPS.length - 1} onClick={() => setStepIdx(i => i + 1)}>
                다음 <ChevronRight size={15} />
              </button>
            </div>
          </div>
        </>
      )}

      {section === "features" && (
        <div className="gd-feature-grid">
          {FEATURES.map(f => {
            const Icon = f.icon;
            return (
              <div key={f.title} className="gd-feature">
                <h4><Icon size={16} /> {f.title}</h4>
                <ul>{f.items.map((it, i) => <li key={i}>{it}</li>)}</ul>
              </div>
            );
          })}
          {inApp && role !== "admin" && (
            <div className="text-dim small-note">조회 전용 계정에서는 기록 입력, 학생 관리 등 일부 기능을 쓸 수 없어요.</div>
          )}
        </div>
      )}

      {section === "devices" && (
        <div className="gd-feature-grid">
          {DEVICE_NOTES.map(d => (
            <div key={d.title} className="gd-feature">
              <h4>{d.title}</h4>
              <p>{d.body}</p>
            </div>
          ))}
        </div>
      )}
      <style>{GUIDE_CSS}</style>
    </div>
  );
}

// 프로그램 안의 "사용 가이드" 탭
export function GuidePanel(props) {
  return (
    <div className="panel gd-panel">
      <h3 className="gd-title"><BookOpen size={18} color="var(--gold)" /> 사용 가이드</h3>
      <GuideContent inApp {...props} />
    </div>
  );
}

// 첫 화면(로그인 전)에서 여는 모달. 앱의 색 변수가 .paps-app 안에서만 정의되어 있으므로
// document.body로 포털하지 않고 첫 화면(.paps-app) 안에 그대로 그린다.
export function GuideModal({ onClose, onStartPractice }) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel gd-modal" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><BookOpen size={18} color="var(--gold)" /> 사용 가이드</h3>
          <button className="icon-btn" onClick={onClose} aria-label="닫기"><X size={16} /></button>
        </div>
        <div className="modal-body">
          <GuideContent onStartPractice={onStartPractice} />
        </div>
      </div>
    </div>
  );
}

const GUIDE_CSS = `
.gd { display: grid; gap: 14px; }
.gd-title { display: flex; align-items: center; gap: 8px; margin: 0 0 6px; }
.gd-modal { max-width: 860px; }
.gd-practice-note { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 10px; padding: 10px 12px; border-radius: 12px;
  background: var(--accent-soft); border: 1px solid var(--line-2); font-size: 13.5px; line-height: 1.55; }
.gd-practice-note > svg { flex-shrink: 0; color: var(--accent); }
.gd-practice-note > span { flex: 1 1 240px; min-width: 0; }
.gd-sections { display: flex; gap: 6px; flex-wrap: wrap; }
.gd-section-btn { border: 1px solid var(--line-2); background: transparent; color: var(--text-dim); border-radius: 999px;
  padding: 7px 14px; font: inherit; font-size: 13.5px; cursor: pointer; }
.gd-section-btn.on { background: var(--panel-2); color: var(--text); border-color: var(--accent); font-weight: 700; }
.gd-section-btn:focus-visible, .gd-step-dot:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }

.gd-stepper { display: flex; gap: 6px; overflow-x: auto; padding-bottom: 4px; scrollbar-width: thin; }
.gd-step-dot { flex: 0 0 auto; display: inline-flex; align-items: center; gap: 6px; border: 1px solid var(--line-2);
  background: transparent; color: var(--text-dim); border-radius: 999px; padding: 5px 11px 5px 5px; font: inherit; font-size: 12.5px; cursor: pointer; }
.gd-step-dot span { width: 22px; height: 22px; border-radius: 50%; display: grid; place-items: center; background: var(--panel-2);
  font-weight: 700; font-size: 12px; color: var(--text); }
.gd-step-dot em { font-style: normal; white-space: nowrap; }
.gd-step-dot.done span { background: var(--line-2); }
.gd-step-dot.on { border-color: var(--accent); color: var(--text); }
.gd-step-dot.on span { background: var(--accent); color: var(--on-accent); }

.gd-card-step { border: 1px solid var(--line); border-radius: 14px; padding: 16px; background: var(--panel-2); display: grid; gap: 14px; }
.gd-step-head { display: flex; flex-wrap: wrap; align-items: baseline; gap: 6px 10px; }
.gd-step-head h4 { margin: 0; font-size: 18px; font-family: var(--display); text-wrap: balance; }
.gd-step-count { font-size: 12px; color: var(--text-dim); font-variant-numeric: tabular-nums; }
.gd-who { margin-left: auto; font-size: 12px; color: var(--text-dim); border: 1px solid var(--line-2); border-radius: 999px; padding: 2px 9px; }
.gd-step-grid { display: grid; grid-template-columns: minmax(0, 1.1fr) minmax(0, 1fr); gap: 16px; align-items: start; }
@media (max-width: 720px) { .gd-step-grid { grid-template-columns: minmax(0, 1fr); } }
.gd-explain { display: grid; gap: 10px; min-width: 0; }
.gd-actions { list-style: none; margin: 0; padding: 0; display: grid; gap: 10px; }
.gd-actions li { display: flex; gap: 10px; align-items: flex-start; font-size: 15px; line-height: 1.55; }
.gd-actions li span { min-width: 0; word-break: keep-all; }
.gd-tip { display: flex; gap: 8px; align-items: flex-start; font-size: 13.5px; line-height: 1.55; color: var(--text);
  background: var(--panel); border: 1px dashed var(--line-2); border-radius: 10px; padding: 8px 10px; word-break: keep-all; }
.gd-tip svg { flex-shrink: 0; margin-top: 3px; color: var(--gold); }
.gd-more summary { cursor: pointer; font-size: 13px; color: var(--text-dim); }
.gd-more ul { margin: 8px 0 0; padding-left: 18px; display: grid; gap: 6px; font-size: 13px; line-height: 1.6; color: var(--text-dim); word-break: keep-all; }
.gd-nav { display: flex; gap: 8px; justify-content: space-between; flex-wrap: wrap; }
.gd-nav .btn { display: inline-flex; align-items: center; gap: 4px; }
.gd-nav .btn:disabled { opacity: 0.4; cursor: default; }
@media (max-width: 560px) {
  .gd-nav .btn-secondary { order: -1; flex-basis: 100%; justify-content: center; }
}

/* 번호 표시 */
.gd-pin { position: absolute; top: -9px; right: -9px; width: 20px; height: 20px; border-radius: 50%; display: grid; place-items: center;
  background: var(--gold); color: #1a1405; font-style: normal; font-weight: 800; font-size: 11.5px; box-shadow: 0 0 0 2px var(--panel-2); z-index: 1; }
.gd-pin.static { position: static; flex-shrink: 0; width: 24px; height: 24px; font-size: 12.5px; margin-top: 1px; box-shadow: none; }
.gd-target { position: relative; display: inline-flex; border-radius: 9px; }
.gd-target.gd-block { display: flex; }
.gd-hl { outline: 2px solid var(--gold); outline-offset: 3px; animation: gdPulse 1.8s ease-in-out infinite; }
@keyframes gdPulse { 0%, 100% { outline-color: var(--gold); } 50% { outline-color: transparent; } }
@media (prefers-reduced-motion: reduce) { .gd-hl { animation: none; } }

/* 데모 화면 */
.gd-screen { border: 1px solid var(--line-2); border-radius: 12px; background: var(--bg); overflow: hidden; }
.gd-screen-bar { display: flex; gap: 5px; padding: 7px 10px; border-bottom: 1px solid var(--line); background: var(--panel); }
.gd-screen-bar span { width: 8px; height: 8px; border-radius: 50%; background: var(--line-2); }
.gd-screen-body { padding: 16px 14px 18px; display: grid; gap: 14px; font-size: 12.5px; color: var(--text); }
.gd-center { display: grid; justify-items: center; gap: 14px; }
.gd-brand { font-family: var(--display); color: var(--gold); letter-spacing: 1px; }
.gd-tabs { display: flex; gap: 4px; flex-wrap: wrap; }
.gd-tab { padding: 3px 7px; border-radius: 6px; color: var(--text-dim); font-size: 11px; }
.gd-tab.on { background: var(--panel-2); color: var(--text); font-weight: 700; }
.gd-btn { display: inline-flex; align-items: center; gap: 4px; padding: 6px 11px; border-radius: 8px; border: 1px solid var(--line-2);
  background: var(--panel); font-size: 12px; white-space: nowrap; }
.gd-btn.primary { background: var(--grad, var(--accent)); color: var(--on-accent); border-color: transparent; font-weight: 700; }
.gd-btn.danger { background: #E85D5D; color: #fff; border-color: transparent; font-weight: 700; }
.gd-field { display: grid; gap: 2px; min-width: 0; }
.gd-field small { color: var(--text-dim); font-size: 10.5px; }
.gd-field > span { border: 1px solid var(--line-2); border-radius: 7px; padding: 5px 8px; background: var(--panel); }
.gd-stack { display: grid; gap: 8px; width: 100%; max-width: 260px; }
.gd-row { display: flex; flex-wrap: wrap; gap: 5px; align-items: center; }
.gd-between { justify-content: space-between; }
.gd-chip { padding: 3px 9px; border-radius: 999px; border: 1px solid var(--line-2); font-size: 11.5px; white-space: nowrap; }
.gd-chip.on { background: var(--grad, var(--accent)); border-color: transparent; color: var(--on-accent); font-weight: 700; }
.gd-drop { display: flex; flex-wrap: wrap; align-items: center; justify-content: center; gap: 8px; width: 100%; padding: 14px 10px;
  border: 1.5px dashed var(--line-2); border-radius: 10px; color: var(--text-dim); text-align: center; }
.gd-list { display: grid; gap: 6px; width: 100%; }
.gd-list > span { display: flex; align-items: center; gap: 8px; padding: 5px 8px; border-radius: 7px; background: var(--panel); }
.gd-input { min-width: 46px; padding: 3px 8px; border: 1px solid var(--line-2); border-radius: 6px; background: var(--bg); font-variant-numeric: tabular-nums; }
.gd-input.caret::after { content: ""; display: inline-block; width: 1px; height: 12px; margin-left: 2px; background: var(--text); vertical-align: -2px; animation: gdCaret 1s steps(1) infinite; }
@keyframes gdCaret { 50% { opacity: 0; } }
.gd-ok { color: #7FD98A; font-style: normal; font-weight: 700; }
.gd-big { display: grid; justify-items: center; font-family: var(--display); font-size: 22px; font-weight: 800; }
.gd-big small { font-size: 11px; color: var(--text-dim); font-weight: 600; }
.gd-table { display: grid; gap: 4px; width: 100%; }
.gd-table > span { display: grid; grid-template-columns: 1.3fr repeat(3, 1fr); gap: 4px; align-items: center; }
.gd-table .h b { color: var(--text-dim); font-weight: 600; font-size: 11px; }
.gd-table b { font-weight: 600; font-variant-numeric: tabular-nums; }
.gd-table i { font-style: normal; font-weight: 800; text-align: center; border-radius: 6px; padding: 2px 0; color: #1a1405; }
.gd-table .g1 { background: #FFC93C; } .gd-table .g2 { background: #7FD98A; } .gd-table .g3 { background: #4EA8DE; } .gd-table .g4 { background: #F2994A; }
.gd-podium { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; align-items: end; }
.gd-podium > span { display: grid; justify-items: center; gap: 2px; border-radius: 8px 8px 0 0; padding: 6px 4px; font-weight: 800; color: #1a1405; font-variant-numeric: tabular-nums; }
.gd-podium small { font-weight: 600; font-size: 10.5px; }
.gd-podium .p1 { background: #FFC93C; height: 74px; } .gd-podium .p2 { background: #C9D3DC; height: 58px; } .gd-podium .p3 { background: #CD8B4B; height: 46px; }
.gd-sub { color: var(--text-dim); font-size: 11.5px; }
.gd-card { padding: 7px 9px; border-radius: 8px; background: var(--panel); }
.gd-check { color: var(--text-dim); }

/* 기능 한눈에 · 여러 기기 */
.gd-feature-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 12px; }
.gd-feature { border: 1px solid var(--line); border-radius: 12px; padding: 14px; background: var(--panel-2); min-width: 0; }
.gd-feature h4 { display: flex; align-items: center; gap: 7px; margin: 0 0 8px; font-size: 15px; }
.gd-feature h4 svg { color: var(--gold); }
.gd-feature ul { margin: 0; padding-left: 18px; display: grid; gap: 6px; font-size: 13.5px; line-height: 1.6; word-break: keep-all; }
.gd-feature p { margin: 0; font-size: 13.5px; line-height: 1.65; color: var(--text-dim); word-break: keep-all; }
`;
