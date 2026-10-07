// 연습모드에 쓰는 가짜 학생 명단과 예시 기록.
// 실존 인물과 겹치지 않도록 흔한 성과 이름 글자를 조합해 만들고, 매번 같은 결과가 나오도록
// 고정된 씨앗값으로 난수를 만든다(연습할 때마다 화면이 똑같이 시작되게).

export const PRACTICE_SCHOOL_NAME = "연습중학교";
// 연습모드 안에서 화면 잠금(빔프로젝터·자동 잠금)과 개설자 로그인을 체험할 때 쓰는 비밀번호.
export const PRACTICE_VIEWER_PASSWORD = "1234";
export const PRACTICE_FOUNDER_PASSWORD = "0000";

const SURNAMES = ["김", "이", "박", "최", "정", "강", "조", "윤", "장", "임", "한", "오", "서", "신"];
const GIVEN = ["하늘", "바다", "가람", "다온", "라온", "보람", "새봄", "슬기", "아름", "나래", "누리", "한결", "이든", "시원", "도담", "은솔", "하랑", "푸름"];

function seededRandom(seed) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 4294967296;
  };
}

// 1학년 1반·2반, 2학년 1반. 1학년 2반은 기록을 비워 두어 직접 입력해 볼 수 있게 한다.
const CLASSES = [
  { grade: 1, classNum: 1, size: 12 },
  { grade: 1, classNum: 2, size: 12 },
  { grade: 2, classNum: 1, size: 10 },
];

export function makePracticeStudents() {
  const rand = seededRandom(20260301);
  const students = [];
  let n = 0;
  CLASSES.forEach(({ grade, classNum, size }) => {
    for (let number = 1; number <= size; number++) {
      const name = SURNAMES[Math.floor(rand() * SURNAMES.length)] + GIVEN[n % GIVEN.length];
      students.push({
        id: "practice_" + grade + "_" + classNum + "_" + number,
        grade, classNum, number, name,
        gender: number % 2 === 1 ? "M" : "F",
      });
      n++;
    }
  });
  return students;
}

// 종목별 [남 최소, 남 최대, 여 최소, 여 최대, 소수 자릿수]
const RANGES = {
  shuttlerun: [25, 85, 14, 55, 0],
  situp: [18, 60, 10, 45, 0],
  fifty_m: [7.3, 9.8, 8.4, 11.2, 2],
  longjump: [155, 235, 125, 190, 1],
  sitreach: [-2, 16, 3, 21, 1],
};

// 1학년 1반은 다섯 종목을 대부분 채우고(몇 명은 일부러 비움), 2학년 1반은 세 종목만 채운다.
const FILL_PLAN = [
  { grade: 1, classNum: 1, events: ["shuttlerun", "situp", "fifty_m", "longjump", "sitreach"], skipEvery: 5 },
  { grade: 2, classNum: 1, events: ["shuttlerun", "situp", "fifty_m"], skipEvery: 4 },
];

export function makePracticeRecords(students, year, recKey) {
  const rand = seededRandom(7);
  const records = {};
  const now = Date.now();
  FILL_PLAN.forEach(({ grade, classNum, events, skipEvery }) => {
    students
      .filter(s => s.grade === grade && s.classNum === classNum)
      .forEach(s => {
        events.forEach((eventId, ei) => {
          if ((s.number + ei) % skipEvery === 0) return; // 아직 측정 안 한 학생처럼 비워 둔다
          const [mMin, mMax, fMin, fMax, digits] = RANGES[eventId];
          const [lo, hi] = s.gender === "M" ? [mMin, mMax] : [fMin, fMax];
          const factor = Math.pow(10, digits);
          const value = Math.round((lo + rand() * (hi - lo)) * factor) / factor;
          records[recKey(s.id, eventId, year)] = { value, schoolGradeAtMeasure: s.grade, updatedAt: now };
        });
      });
  });
  return records;
}

// 접근권한 탭에서 승인·거절을 연습해 볼 수 있도록 넣어 두는 예시 신청.
export function makePracticeAccessList() {
  const now = Date.now();
  return {
    requests: [
      { id: "practice_req_editor", name: "1학년 2반 담임 박연습", submittedAt: now - 10 * 60 * 1000, status: "pending", type: "editor" },
      { id: "practice_req_viewer", name: "보건교사 이예시", submittedAt: now - 2 * 24 * 60 * 60 * 1000, status: "approved", type: "viewer" },
    ],
  };
}
