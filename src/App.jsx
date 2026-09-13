import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import {
  Trophy, Medal, RefreshCw, Plus, Trash2, Users,
  ClipboardList, Monitor, X, Maximize2, Minimize2, AlertTriangle,
  CheckCircle2, XCircle, Info, Play,
  RotateCcw, Square, Smartphone, Copy,
  FileSpreadsheet, Check, ShieldCheck, Database, Award, Eye, EyeOff
} from "lucide-react";
import * as XLSX from "xlsx";
import { storage } from "./storage.js";
import { usePwaInstall } from "./pwaInstall.js";

// 학기말 마감 화면의 "문의하기" 버튼이 여는 구글 폼 링크.
const INQUIRY_FORM_URL = "https://forms.gle/hEwyNGdcMKjtMTWu8";

/* ============================== 상수 정의 ============================== */

const EVENTS = [
  { id: "shuttlerun",    name: "왕복오래달리기",       category: "심폐지구력", unit: "회", better: "high", altGroup: "cardio" },
  { id: "run_walk",      name: "오래달리기-걷기",       category: "심폐지구력", unit: "초", better: "low",  altGroup: "cardio" },
  { id: "step_test",     name: "스텝검사",             category: "심폐지구력", unit: "PEI", better: "high", altGroup: "cardio" },
  { id: "situp",         name: "윗몸말아올리기",       category: "근력·근지구력", unit: "회", better: "high" },
  { id: "pushup",        name: "팔굽혀펴기",           category: "근력·근지구력", unit: "회", better: "high" },
  { id: "sitreach",      name: "앉아윗몸앞으로굽히기", category: "유연성",     unit: "cm", better: "high", altGroup: "flex" },
  { id: "flex_total",    name: "종합유연성",           category: "유연성",     unit: "점", better: "high", altGroup: "flex" },
  { id: "longjump",      name: "제자리멀리뛰기",       category: "순발력",     unit: "cm", better: "high" },
  { id: "fifty_m",       name: "50m달리기",            category: "순발력",     unit: "초", better: "low"  },
  { id: "gripstrength",  name: "악력",                 category: "근력·근지구력", unit: "kg", better: "high" },
];
const BMI_EVENT = { id: "bmi", name: "BMI", category: "신체구성", unit: "", better: "low" };
const BODYFAT_EVENT = { id: "bodyfat", name: "체지방률", category: "신체구성", unit: "%", better: "low" };
// BMI·체지방률은 등급을 매기지 않고 종합등급에도 반영되지 않는 성장 확인용 참고 지표라서,
// 등급이 매겨지는 EVENTS 배열에는 넣지 않고 ALL_EVENTS에만 별도로 합친다.
const NO_GRADE_EVENT_IDS = ["bmi", "bodyfat"];
const ALL_EVENTS = [...EVENTS, BMI_EVENT, BODYFAT_EVENT];
const EVENT_MAP = Object.fromEntries(ALL_EVENTS.map(e => [e.id, e]));

const GENDERS = [{ id: "M", label: "남" }, { id: "F", label: "여" }];
const SCHOOL_GRADES = [1, 2, 3];

const GRADE_COLORS = { 1: "#FFC93C", 2: "#7FD98A", 3: "#4EA8DE", 4: "#F2994A", 5: "#E85D5D" };
const GRADE_LABELS = { 1: "1등급", 2: "2등급", 3: "3등급", 4: "4등급", 5: "5등급" };

// 여러 학교/선생님이 같은 화면(아티팩트)을 함께 사용하더라도 학생 기록이
// 서로 섞이지 않도록, 데이터 저장 키에 "워크스페이스 코드"를 포함시킨다.
// 코드가 다르면 완전히 분리된 저장 공간을 쓰게 된다.
const WORKSPACE_CODE_KEY = "paps:workspace-code";
function configKey(code) { return "paps:config:" + code; }
function recordsKey(code) { return "paps:records:" + code; }

function sanitizeWorkspaceCode(raw) {
  return String(raw || "").trim().replace(/[\s/\\'"]+/g, "-").slice(0, 40);
}

/* ============================== 유틸 함수 ============================== */

function uid(prefix = "id") {
  return prefix + "_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// 터무니없는 값(오타 등)이 그대로 저장되는 걸 막기 위한 최소한의 상식적 범위.
// 실제로 나올 수 있는 값보다 넉넉하게 잡아서, 정상적인 기록은 절대 막히지 않게 한다.
const EVENT_PLAUSIBLE_RANGE = {
  shuttlerun: [0, 300], run_walk: [0, 3600], step_test: [0, 200],
  situp: [0, 200], pushup: [0, 200], gripstrength: [0, 100],
  sitreach: [-30, 60], flex_total: [0, 8], longjump: [0, 400], fifty_m: [2, 30],
  bodyfat: [0, 60], bmi: [8, 60],
};
function isPlausibleValue(eventId, value) {
  const range = EVENT_PLAUSIBLE_RANGE[eventId];
  if (!range || value === null || value === undefined || Number.isNaN(value)) return true;
  return value >= range[0] && value <= range[1];
}
// BMI는 계산값(value)뿐 아니라, 오타가 가장 흔한 원재료인 신장·체중 자체도 함께 확인한다.
function isPlausibleBmiParts(parts) {
  if (!parts) return true;
  const h = parts.height, w = parts.weight;
  if (h !== undefined && h !== null && !Number.isNaN(h) && (h < 100 || h > 220)) return false;
  if (w !== undefined && w !== null && !Number.isNaN(w) && (w < 15 || w > 200)) return false;
  return true;
}

function recKey(studentId, eventId, year) {
  return studentId + "__" + eventId + "__" + year;
}

// 4초마다 도는 동기화 폴링이, 이 기기에서 방금 막 저장한(아직 서버 응답에 반영되지
// 않은) 최신 기록을 낡은 스냅샷으로 덮어써버리는 경합을 막기 위한 병합 함수.
// 각 기록의 updatedAt을 비교해서, 로컬이 폴링 결과보다 더 최신이면 로컬 값을 지킨다.
function mergeRecords(local, polled) {
  const merged = { ...polled };
  Object.keys(local || {}).forEach(key => {
    const localEntry = local[key];
    const polledEntry = polled[key];
    if (!polledEntry || (localEntry.updatedAt || 0) > (polledEntry.updatedAt || 0)) {
      merged[key] = localEntry;
    }
  });
  return merged;
}

function thisYear() {
  return new Date().getFullYear();
}

// 실시간 측정 중 프로젝터·대형화면에 학생 이름이 그대로 노출되는 것을 막기 위한 이름
// 마스킹("이름 마스킹 모드" 토글에서 사용). 2글자면 뒷글자, 3글자 이상이면 가운데 글자(들)를
// 가린다(예: "김민" → "김*", "홍길동" → "홍*동", "황보영수" → "황**수"). 1글자 이름은 그대로 둔다.
function maskStudentName(name) {
  if (!name) return name;
  const chars = Array.from(String(name));
  if (chars.length <= 1) return name;
  if (chars.length === 2) return chars[0] + "*";
  return chars[0] + "*".repeat(chars.length - 2) + chars[chars.length - 1];
}

function fmtValue(v, unit) {
  if (v === null || v === undefined || Number.isNaN(v)) return "-";
  const rounded = Math.round(v * 100) / 100;
  return rounded + (unit ? " " + unit : "");
}

// ---- 공식 기준: 학교건강검사규칙 [별표 4] 신체능력검사(필수평가) 기준표 ----
// 선생님이 올려주신 공식 hwp 문서에서 그대로 옮긴 실제 수치(각 등급 구간의 최소~최대값
// 그대로). 왕복오래달리기·윗몸말아올리기·악력·앉아윗몸앞으로굽히기·제자리멀리뛰기 5개
// 종목을 초4~고3까지 학년별·성별로 담고 있다. 5등급/1등급의 극단값은 "그 이상/이하는
// 각각 최소·최대 등급"이라는 문서의 규칙에 따라 -9999/9999(범위 없음)로 열어둔다.
// 문서에 없는 팔굽혀펴기·50m달리기·BMI는 기존 예시값을 그대로 둔다(아래 ESTIMATED_BANDS).
const OFFICIAL_LEVEL_BANDS = {
  elem: {
    4: {
      M: { shuttlerun: [[-9999,25],[26,44],[45,68],[69,95],[96,9999]], situp: [[-9999,6],[7,21],[22,39],[40,79],[80,9999]], gripstrength: [[-9999,11.4],[11.5,14.9],[15.0,18.4],[18.5,30.9],[31.0,9999]], sitreach: null, longjump: [[-9999,100],[100.1,130],[130.1,149],[149.1,170],[170.1,9999]], pushup: null, fifty_m: [[13.21,9999],[10.51,13.20],[9.71,10.50],[8.81,9.70],[-9999,8.80]] },
      F: { shuttlerun: [[-9999,20],[21,39],[40,56],[57,76],[77,9999]], situp: [[-9999,5],[6,17],[18,28],[29,59],[60,9999]], gripstrength: [[-9999,10.4],[10.5,13.4],[13.5,17.9],[18.0,28.9],[29.0,9999]], sitreach: null, longjump: [[-9999,97],[97.1,119],[119.1,135],[135.1,161],[161.1,9999]], pushup: null, fifty_m: [[13.31,9999],[11.01,13.30],[10.41,11.00],[9.41,10.40],[-9999,9.40]] },
    },
    5: {
      M: { shuttlerun: [[-9999,28],[29,49],[50,72],[73,99],[100,9999]], situp: [[-9999,9],[10,21],[22,39],[40,79],[80,9999]], gripstrength: [[-9999,12.4],[12.5,16.9],[17.0,22.9],[23.0,30.9],[31.0,9999]], sitreach: [[-9999,-4.1],[-4.0,0.9],[1.0,4.9],[5.0,7.9],[8.0,9999]], longjump: [[-9999,111],[111.1,141],[141.1,159],[159.1,180],[180.1,9999]], pushup: null, fifty_m: [[13.21,9999],[10.21,13.20],[9.41,10.20],[8.51,9.40],[-9999,8.50]] },
      F: { shuttlerun: [[-9999,22],[23,44],[45,62],[63,84],[85,9999]], situp: [[-9999,6],[7,22],[23,35],[36,59],[60,9999]], gripstrength: [[-9999,11.9],[12.0,15.4],[15.5,18.9],[19.0,28.9],[29.0,9999]], sitreach: [[-9999,0.9],[1.0,4.9],[5.0,6.9],[7.0,9.9],[10.0,9999]], longjump: [[-9999,100],[100.1,123],[123.1,139],[139.1,170],[170.1,9999]], pushup: null, fifty_m: [[13.31,9999],[10.71,13.30],[9.91,10.70],[8.91,9.90],[-9999,8.90]] },
    },
    6: {
      M: { shuttlerun: [[-9999,31],[32,53],[54,77],[78,103],[104,9999]], situp: [[-9999,9],[10,21],[22,39],[40,79],[80,9999]], gripstrength: [[-9999,14.9],[15.0,18.9],[19.0,26.4],[26.5,34.9],[35.0,9999]], sitreach: [[-9999,-4.1],[-4.0,0.9],[1.0,4.9],[5.0,7.9],[8.0,9999]], longjump: [[-9999,122],[122.1,148],[148.1,167],[167.1,200],[200.1,9999]], pushup: null, fifty_m: [[12.51,9999],[10.01,12.50],[9.11,10.00],[8.11,9.10],[-9999,8.10]] },
      F: { shuttlerun: [[-9999,24],[25,49],[50,68],[69,92],[93,9999]], situp: [[-9999,6],[7,22],[23,42],[43,59],[60,9999]], gripstrength: [[-9999,13.9],[14.0,18.9],[19.0,21.9],[22.0,32.9],[33.0,9999]], sitreach: [[-9999,1.9],[2.0,4.9],[5.0,9.9],[10.0,13.9],[14.0,9999]], longjump: [[-9999,100],[100.1,127],[127.1,144],[144.1,175],[175.1,9999]], pushup: null, fifty_m: [[12.91,9999],[10.71,12.90],[9.81,10.70],[8.91,9.80],[-9999,8.90]] },
    },
  },
  middle: {
    1: {
      M: { shuttlerun: [[-9999,19],[20,35],[36,49],[50,63],[64,9999]], situp: [[-9999,13],[14,32],[33,54],[55,89],[90,9999]], gripstrength: [[-9999,16.4],[16.5,22.4],[22.5,29.9],[30.0,41.9],[42.0,9999]], sitreach: [[-9999,-4.1],[-4.0,1.9],[2.0,5.9],[6.0,9.9],[10.0,9999]], longjump: [[-9999,131],[131.1,159],[159.1,177],[177.1,211],[211.1,9999]], pushup: [[-9999,3],[4,11],[12,24],[25,33],[34,9999]], fifty_m: [[11.51,9999],[9.31,11.50],[8.41,9.30],[7.51,8.40],[-9999,7.50]] },
      F: { shuttlerun: [[-9999,13],[14,18],[19,24],[25,34],[35,9999]], situp: [[-9999,6],[7,21],[22,42],[43,57],[58,9999]], gripstrength: [[-9999,13.9],[14.0,18.9],[19.0,22.9],[23.0,35.9],[36.0,9999]], sitreach: [[-9999,1.9],[2.0,7.9],[8.0,10.9],[11.0,14.9],[15.0,9999]], longjump: [[-9999,100],[100.1,127],[127.1,144],[144.1,175],[175.1,9999]], pushup: [[-9999,5],[6,13],[14,23],[24,44],[45,9999]], fifty_m: [[12.21,9999],[10.51,12.20],[9.81,10.50],[8.81,9.80],[-9999,8.80]] },
    },
    2: {
      M: { shuttlerun: [[-9999,21],[22,37],[38,51],[52,65],[66,9999]], situp: [[-9999,13],[14,32],[33,54],[55,89],[90,9999]], gripstrength: [[-9999,21.9],[22.0,28.4],[28.5,36.9],[37.0,44.4],[44.5,9999]], sitreach: [[-9999,-4.1],[-4.0,1.9],[2.0,6.9],[7.0,9.9],[10.0,9999]], longjump: [[-9999,136],[136.1,169],[169.1,187],[187.1,218],[218.1,9999]], pushup: [[-9999,3],[4,11],[12,24],[25,33],[34,9999]], fifty_m: [[11.51,9999],[9.01,11.50],[8.21,9.00],[7.31,8.20],[-9999,7.30]] },
      F: { shuttlerun: [[-9999,14],[15,20],[21,28],[29,39],[40,9999]], situp: [[-9999,6],[7,18],[19,38],[39,57],[58,9999]], gripstrength: [[-9999,13.9],[14.0,19.4],[19.5,25.4],[25.5,35.9],[36.0,9999]], sitreach: [[-9999,1.9],[2.0,7.9],[8.0,10.9],[11.0,14.9],[15.0,9999]], longjump: [[-9999,100],[100.1,127],[127.1,145],[145.1,183],[183.1,9999]], pushup: [[-9999,5],[6,13],[14,23],[24,39],[40,9999]], fifty_m: [[12.21,9999],[10.51,12.20],[9.81,10.50],[8.81,9.80],[-9999,8.80]] },
    },
    3: {
      M: { shuttlerun: [[-9999,23],[24,39],[40,53],[54,67],[68,9999]], situp: [[-9999,13],[14,32],[33,54],[55,89],[90,9999]], gripstrength: [[-9999,24.9],[25.0,32.9],[33.0,40.4],[40.5,48.4],[48.5,9999]], sitreach: [[-9999,-3.1],[-3.0,2.5],[2.6,6.9],[7.0,9.9],[10.0,9999]], longjump: [[-9999,145],[145.1,180],[180.1,201],[201.1,238],[238.1,9999]], pushup: [[-9999,3],[4,13],[14,24],[25,33],[34,9999]], fifty_m: [[11.01,9999],[8.51,11.00],[7.81,8.50],[7.01,7.80],[-9999,7.00]] },
      F: { shuttlerun: [[-9999,15],[16,22],[23,32],[33,44],[45,9999]], situp: [[-9999,5],[6,16],[17,33],[34,51],[52,9999]], gripstrength: [[-9999,15.9],[16.0,19.4],[19.5,27.4],[27.5,35.9],[36.0,9999]], sitreach: [[-9999,1.9],[2.0,7.9],[8.0,10.9],[11.0,15.9],[16.0,9999]], longjump: [[-9999,100],[100.1,127],[127.1,145],[145.1,183],[183.1,9999]], pushup: [[-9999,5],[6,13],[14,23],[24,39],[40,9999]], fifty_m: [[12.21,9999],[10.51,12.20],[9.81,10.50],[8.81,9.80],[-9999,8.80]] },
    },
  },
  high: {
    1: {
      M: { shuttlerun: [[-9999,25],[26,41],[42,55],[56,69],[70,9999]], situp: [[-9999,14],[15,34],[35,59],[60,89],[90,9999]], gripstrength: [[-9999,28.9],[29.0,35.4],[35.5,42.4],[42.5,60.9],[61.0,9999]], sitreach: [[-9999,-2.1],[-2.0,3.9],[4.0,8.9],[9.0,12.9],[13.0,9999]], longjump: [[-9999,160],[160.1,195],[195.1,216],[216.1,255],[255.1,9999]], pushup: [[-9999,6],[7,15],[16,29],[30,45],[46,9999]], fifty_m: [[10.01,9999],[8.11,10.00],[7.61,8.10],[7.01,7.60],[-9999,7.00]] },
      F: { shuttlerun: [[-9999,16],[17,24],[25,36],[37,49],[50,9999]], situp: [[-9999,3],[4,12],[13,29],[30,39],[40,9999]], gripstrength: [[-9999,16.4],[16.5,22.9],[23.0,28.9],[29.0,35.9],[36.0,9999]], sitreach: [[-9999,1.9],[2.0,7.9],[8.0,10.9],[11.0,15.9],[16.0,9999]], longjump: [[-9999,100],[100.1,139],[139.1,159],[159.1,186],[186.1,9999]], pushup: [[-9999,5],[6,13],[14,23],[24,39],[40,9999]], fifty_m: [[12.21,9999],[10.51,12.20],[9.81,10.50],[8.81,9.80],[-9999,8.80]] },
    },
    2: {
      M: { shuttlerun: [[-9999,27],[28,43],[44,57],[58,71],[72,9999]], situp: [[-9999,16],[17,34],[35,59],[60,89],[90,9999]], gripstrength: [[-9999,30.9],[31.0,38.9],[39.0,45.9],[46.0,60.9],[61.0,9999]], sitreach: [[-9999,0.0],[0.1,4.9],[5.0,10.9],[11.0,15.9],[16.0,9999]], longjump: [[-9999,177],[177.1,212],[212.1,228],[228.1,258],[258.1,9999]], pushup: [[-9999,10],[11,24],[25,41],[42,49],[50,9999]], fifty_m: [[9.51,9999],[7.91,9.50],[7.51,7.90],[6.71,7.50],[-9999,6.70]] },
      F: { shuttlerun: [[-9999,17],[18,26],[27,40],[41,54],[55,9999]], situp: [[-9999,3],[4,12],[13,29],[30,39],[40,9999]], gripstrength: [[-9999,17.9],[18.0,24.9],[25.0,29.4],[29.5,37.4],[37.5,9999]], sitreach: [[-9999,4.9],[5.0,8.9],[9.0,11.9],[12.0,16.9],[17.0,9999]], longjump: [[-9999,100],[100.1,139],[139.1,159],[159.1,186],[186.1,9999]], pushup: [[-9999,8],[9,17],[18,29],[30,39],[40,9999]], fifty_m: [[12.21,9999],[10.51,12.20],[9.51,10.50],[8.81,9.50],[-9999,8.80]] },
    },
    3: {
      M: { shuttlerun: [[-9999,29],[30,45],[46,59],[60,73],[74,9999]], situp: [[-9999,16],[17,34],[35,59],[60,89],[90,9999]], gripstrength: [[-9999,30.9],[31.0,38.9],[39.0,45.9],[46.0,63.4],[63.5,9999]], sitreach: [[-9999,0.0],[0.1,5.9],[6.0,10.9],[11.0,15.9],[16.0,9999]], longjump: [[-9999,185],[185.1,221],[221.1,243],[243.1,264],[264.1,9999]], pushup: [[-9999,16],[17,29],[30,45],[46,55],[56,9999]], fifty_m: [[8.71,9999],[7.91,8.70],[7.51,7.90],[6.71,7.50],[-9999,6.70]] },
      F: { shuttlerun: [[-9999,17],[18,26],[27,40],[41,54],[55,9999]], situp: [[-9999,3],[4,12],[13,29],[30,39],[40,9999]], gripstrength: [[-9999,17.9],[18.0,24.9],[25.0,29.4],[29.5,37.4],[37.5,9999]], sitreach: [[-9999,4.9],[5.0,8.9],[9.0,11.9],[12.0,16.9],[17.0,9999]], longjump: [[-9999,100],[100.1,139],[139.1,159],[159.1,186],[186.1,9999]], pushup: [[-9999,8],[9,17],[18,29],[30,39],[40,9999]], fifty_m: [[12.21,9999],[10.51,12.20],[9.51,10.50],[8.81,9.50],[-9999,8.80]] },
    },
  },
};
const OFFICIAL_EVENT_IDS = ["shuttlerun", "situp", "gripstrength", "sitreach", "longjump", "pushup", "fifty_m", "run_walk", "step_test", "flex_total"];

// 심폐지구력의 대체 종목(오래달리기-걷기)은 초4는 실시하지 않고, 학년마다 기준이 다르다.
const RUN_WALK_BANDS = {
  elem: {
    4: { M: null, F: null },
    5: { M: [[480,9999],[410,479],[325,409],[282,324],[-9999,281]], F: [[502,9999],[442,501],[360,441],[300,359],[-9999,299]] },
    6: { M: [[450,9999],[380,449],[315,379],[251,314],[-9999,250]], F: [[480,9999],[430,479],[354,429],[300,353],[-9999,299]] },
  },
  middle: {
    1: { M: [[700,9999],[600,699],[503,599],[426,502],[-9999,425]], F: [[609,9999],[518,608],[443,517],[380,442],[-9999,379]] },
    2: { M: [[680,9999],[584,679],[488,583],[417,487],[-9999,416]], F: [[609,9999],[518,608],[443,517],[380,442],[-9999,379]] },
    3: { M: [[660,9999],[568,659],[473,567],[408,472],[-9999,407]], F: [[609,9999],[518,608],[443,517],[380,442],[-9999,379]] },
  },
  high: {
    1: { M: [[640,9999],[552,639],[458,551],[399,457],[-9999,398]], F: [[609,9999],[518,608],[443,517],[380,442],[-9999,379]] },
    2: { M: [[620,9999],[536,619],[443,535],[390,442],[-9999,389]], F: [[609,9999],[518,608],[443,517],[380,442],[-9999,379]] },
    3: { M: [[600,9999],[520,599],[428,519],[381,427],[-9999,380]], F: [[609,9999],[518,608],[443,517],[380,442],[-9999,379]] },
  },
};
// 스텝검사(PEI)와 종합유연성(점수)은 학년·성별 구분 없이 "전 학년 공통" 기준 하나만 쓴다.
const STEP_TEST_BANDS = [[-9999,46.9],[47.0,51.9],[52.0,61.9],[62.0,75.9],[76.0,9999]];
const FLEX_TOTAL_BANDS = [[-9999,4],[5,5],[6,6],[7,7],[8,9999]];
Object.keys(OFFICIAL_LEVEL_BANDS).forEach(levelId => {
  Object.keys(OFFICIAL_LEVEL_BANDS[levelId]).forEach(grade => {
    ["M", "F"].forEach(g => {
      OFFICIAL_LEVEL_BANDS[levelId][grade][g].run_walk = RUN_WALK_BANDS[levelId][grade][g];
      OFFICIAL_LEVEL_BANDS[levelId][grade][g].step_test = STEP_TEST_BANDS;
      OFFICIAL_LEVEL_BANDS[levelId][grade][g].flex_total = FLEX_TOTAL_BANDS;
    });
  });
});

// ---- 같은 체력요소의 "대체 종목" 처리 ----
// 심폐지구력(왕복오래달리기/오래달리기-걷기/스텝검사), 유연성(앉아윗몸앞으로굽히기/종합유연성)처럼
// 한 요소를 여러 종목 중 하나로 측정하는 경우, 종합등급 계산에서 그 요소가 중복으로 반영되면 안 된다.
// "생략"으로 명시적으로 뺀 것은 그대로 두고, 같은 대체군(altGroup) 안에서 생략되지 않은 종목이
// 여러 개 남아있다면 그중 EVENTS 배열 순서상 첫 번째만 남기고 나머지는 종합등급 계산에서만
// 자동으로 제외한다(기록입력·등급표 열 표시 자체는 그대로 유지).
function computeEffectiveSkip(skippedEvents) {
  const skipSet = new Set(skippedEvents || []);
  const groups = {};
  EVENTS.forEach(e => {
    if (e.altGroup) (groups[e.altGroup] = groups[e.altGroup] || []).push(e.id);
  });
  Object.values(groups).forEach(ids => {
    const active = ids.filter(id => !skipSet.has(id));
    if (active.length > 1) active.slice(1).forEach(id => skipSet.add(id));
  });
  return skipSet;
}

function toBandObjects(raw) {
  if (!raw) return null;
  return raw.map((range, idx) => ({ grade: idx + 1, min: range[0], max: range[1] }));
}

// BMI는 등급(1~5등급)으로 나누는 대신, 교육부 학생건강정보센터 자료의 "마름/정상/
// 과체중/경도비만/고도비만" 5단계 공식 분류를 그대로 사용한다(아래 BMI_OFFICIAL_CATEGORIES).
// 그래서 여기 ESTIMATED_BANDS에는 더 이상 bmi를 넣지 않는다. bmi 자체는 등급·종합
// 등급에 반영되지 않으므로, 아래 값은 실제로는 쓰이지 않고 과거 데이터 호환용으로만 남긴다.
const ESTIMATED_BANDS = {
  bmi: { M: [[-9999, 23], [23.1, 25], [25.1, 27], [27.1, 30], [30.1, 9999]],
         F: [[-9999, 22], [22.1, 24], [24.1, 26], [26.1, 29], [29.1, 9999]] },
};

// ---- BMI 공식 분류(교육부 학생건강정보센터, 마름/정상/과체중/경도비만/고도비만) ----
// 각 배열은 [마름 상한, 정상 상한, 과체중 상한, 경도비만 상한]이며, 이보다 크면 고도비만이다.
// (BMI는 "낮을수록 좋다"는 등급 개념이 아니라 마름·정상·과체중·비만이 서로 다른 방향의
// 상태라서, 다른 종목처럼 1~5등급으로 매기지 않고 이 5단계 분류만 참고용으로 보여준다.)
const BMI_OFFICIAL_CATEGORIES = {
  elem: {
    4: { M: [14.2, 20.7, 23.2, 29.9], F: [13.9, 19.8, 22.0, 29.9] },
    5: { M: [14.5, 21.6, 24.4, 29.9], F: [14.2, 20.6, 23.0, 29.9] },
    6: { M: [14.8, 22.5, 24.9, 29.9], F: [14.6, 21.4, 23.9, 29.9] },
  },
  middle: {
    1: { M: [15.3, 23.2, 24.9, 29.9], F: [15.1, 22.1, 24.7, 29.9] },
    2: { M: [15.7, 23.8, 24.9, 29.9], F: [15.6, 22.7, 24.9, 29.9] },
    3: { M: [16.2, 24.3, 24.9, 29.9], F: [16.2, 23.2, 24.9, 29.9] },
  },
  high: {
    1: { M: [16.7, 24.6, 24.9, 29.9], F: [16.7, 23.6, 24.9, 29.9] },
    2: { M: [17.2, 24.9, 24.9, 29.9], F: [17.2, 23.8, 24.9, 29.9] },
    3: { M: [17.6, 24.9, 24.9, 29.9], F: [17.6, 23.9, 24.9, 29.9] },
  },
};
const BMI_CATEGORY_LABELS = ["마름", "정상", "과체중", "경도비만", "고도비만"];
function classifyBmi(bmi, grade, gender, schoolLevel) {
  if (bmi === null || bmi === undefined) return null;
  const boundaries = BMI_OFFICIAL_CATEGORIES[schoolLevel || "middle"]?.[grade]?.[gender];
  if (!boundaries) return null;
  const [b1, b2, b3, b4] = boundaries;
  if (bmi <= b1) return BMI_CATEGORY_LABELS[0];
  if (bmi <= b2) return BMI_CATEGORY_LABELS[1];
  if (bmi <= b3) return BMI_CATEGORY_LABELS[2];
  if (bmi <= b4) return BMI_CATEGORY_LABELS[3];
  return BMI_CATEGORY_LABELS[4];
}

// 체지방률(%Fat) 공식 기준표 — 전 학년 공통(성별로만 구분).
const BODYFAT_OFFICIAL_CATEGORIES = { M: [11.9, 14.9, 24.9, 32.9], F: [14.9, 26.9, 31.9, 39.9] };
function classifyBodyfat(pct, gender) {
  if (pct === null || pct === undefined) return null;
  const boundaries = BODYFAT_OFFICIAL_CATEGORIES[gender];
  if (!boundaries) return null;
  const [b1, b2, b3, b4] = boundaries;
  if (pct <= b1) return BMI_CATEGORY_LABELS[0];
  if (pct <= b2) return BMI_CATEGORY_LABELS[1];
  if (pct <= b3) return BMI_CATEGORY_LABELS[2];
  if (pct <= b4) return BMI_CATEGORY_LABELS[3];
  return BMI_CATEGORY_LABELS[4];
}

function bandsForLevelGrade(levelId, grade, gender, eventId) {
  // BMI·체지방률은 등급(1~5등급)을 매기지 않는 성장 확인용 지표라서 밴드 자체가 없다.
  if (NO_GRADE_EVENT_IDS.includes(eventId)) return null;
  if (OFFICIAL_EVENT_IDS.includes(eventId)) {
    const raw = OFFICIAL_LEVEL_BANDS[levelId]?.[grade]?.[gender]?.[eventId];
    if (!raw) return null;
    // 공식 문서는 "아주낮음(5등급) → 아주높음(1등급)" 순서로 되어 있어, 배열 순서를
    // 그대로 idx+1로 등급을 매기면 1등급과 5등급이 뒤바뀐다. 순서를 뒤집어서 바로잡는다.
    return toBandObjects([...raw].reverse());
  }
  return toBandObjects(ESTIMATED_BANDS[eventId][gender]);
}

// ---- 학교급별(초/중/고) 참고 등급 기준표 ----
// 왕복오래달리기·윗몸말아올리기·악력·앉아윗몸앞으로굽히기·제자리멀리뛰기는 학교건강검사규칙
// [별표 4] 공식 수치를 그대로 담았고, 팔굽혀펴기·50m달리기·BMI는 그 문서에 없어 예시값을
// 그대로 사용한다(학교급별 차이는 반영하지 않음). 화면에서도 이 구분을 함께 안내한다.
const SCHOOL_LEVELS = [
  { id: "elem", label: "초등학교", grades: [4, 5, 6] },
  { id: "middle", label: "중학교", grades: [1, 2, 3] },
  { id: "high", label: "고등학교", grades: [1, 2, 3] },
];
function gradesForLevel(schoolLevel) {
  return (SCHOOL_LEVELS.find(l => l.id === schoolLevel) || SCHOOL_LEVELS[1]).grades;
}

function buildDefaultCriteria(schoolLevel) {
  const lvl = schoolLevel || "middle";
  const grades = gradesForLevel(lvl);
  const criteria = {};
  ALL_EVENTS.forEach(ev => {
    criteria[ev.id] = {};
    ["M", "F"].forEach(g => {
      criteria[ev.id][g] = {};
      grades.forEach(sg => {
        criteria[ev.id][g][sg] = bandsForLevelGrade(lvl, sg, g, ev.id) || [];
      });
    });
  });
  return criteria;
}

function buildReferenceBands() {
  const ref = {};
  SCHOOL_LEVELS.forEach(lvl => {
    ref[lvl.id] = {};
    ALL_EVENTS.forEach(ev => {
      ref[lvl.id][ev.id] = {};
      ["M", "F"].forEach(g => {
        ref[lvl.id][ev.id][g] = {};
        lvl.grades.forEach(grade => {
          ref[lvl.id][ev.id][g][grade] = bandsForLevelGrade(lvl.id, grade, g, ev.id);
        });
      });
    });
  });
  return ref;
}
const REFERENCE_BANDS = buildReferenceBands();

function buildDefaultConfig(schoolLevel, viewerPassword, founderPassword) {
  const lvl = schoolLevel || "middle";
  return {
    students: [],
    criteria: buildDefaultCriteria(lvl),
    settings: { schoolName: "", schoolLevel: lvl, currentYear: thisYear(), skippedEvents: [], theme: "default", viewerPassword: viewerPassword || "", founderPassword: founderPassword || "" },
    createdAt: Date.now(),
  };
}

// 예전 버전(연도 구분 없음)에서 저장된 기록을 "studentId__eventId" 형식으로
// 저장했다면 이를 올해 기록으로 변환해 새 형식("studentId__eventId__year")과
// 호환되게 만든다.
function migrateLegacyRecords(records, students, year) {
  let migrated = false;
  const next = {};
  Object.entries(records || {}).forEach(([k, v]) => {
    const parts = k.split("__");
    if (parts.length === 2) {
      migrated = true;
      const [sid, eid] = parts;
      const student = students.find(s => s.id === sid);
      next[recKey(sid, eid, year)] = { ...v, schoolGradeAtMeasure: student ? student.grade : null };
    } else {
      next[k] = v;
    }
  });
  return { records: next, migrated };
}

// 왕복오래달리기(20m 셔틀런) 표준 진행 방식.
// 국제표준으로 통용되는 Léger 20m Shuttle Run Test(FitnessGram PACER) 프로토콜을
// 그대로 적용한 고정값으로, 중학생 대상 학교 현장 측정에 널리 쓰이는 방식이다.
const SHUTTLE_PRESET = { startSpeed: 8.0, increment: 0.5, distance: 20, levelDuration: 61, maxLevels: 23 };

// 선생님이 직접 녹음한 왕복오래달리기 시작 안내 음성(길이·무음 구간을 다듬어 재생용으로 편집).
// 선생님이 업로드하신 영상 속 공식 안내 음성에서 신호음 이전의 안내 멘트 구간만 추출해 편집.

// 셔틀런(20m 왕복오래달리기) 진행 스케줄 계산 함수 (범용).
function computeShuttleSchedule({ startSpeed, increment, distance, levelDuration, maxLevels }) {
  const schedule = [];
  let t = 0;
  for (let level = 1; level <= maxLevels; level++) {
    const speed = startSpeed + increment * (level - 1);
    const lapTime = (distance * 3.6) / speed; // seconds per 20m at this speed
    const laps = Math.max(1, Math.round(levelDuration / lapTime));
    for (let i = 0; i < laps; i++) {
      t += lapTime;
      schedule.push({ level, lapInLevel: i + 1, time: t, levelStart: i === 0 });
    }
  }
  return schedule;
}

function computeGradeFromBands(value, bands) {
  if (value === null || value === undefined || Number.isNaN(value) || !bands) return null;
  for (const b of bands) {
    if (value >= b.min && value <= b.max) return b.grade;
  }
  return null;
}

/* ============================== 스토리지 I/O ============================== */

// ---- 학생 실명 익명화 ----
// Firebase(Firestore, 외부 상업 클라우드)에는 학생 이름을 절대 저장하지 않는다. 저장되는
// 학생 필드는 id(익명 고유키)·grade·classNum·number·gender뿐이고, 실명은 오직 "이 브라우저
// (교사 기기)"의 localStorage에만 코드별로 남아 화면에 표시할 때만 참고용으로 합쳐진다.
// 그래서 처음 보는 기기(다른 선생님의 컴퓨터 등)에서는 그 기기에 아직 이름표가 없는
// 학생은 "(이름 미확인 - 이 기기)"로 보이는데, 실명 포함 백업 파일을 그 기기에서 한 번
// 불러오면(데이터 백업 탭) 그 기기에도 이름표가 채워진다.
function nameMapKey(code) { return "paps:namemap:" + code; }
function loadLocalNameMap(code) {
  try {
    const raw = window.localStorage.getItem(nameMapKey(code));
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    return {};
  }
}
function saveLocalNameMap(code, map) {
  try {
    window.localStorage.setItem(nameMapKey(code), JSON.stringify(map));
  } catch (e) {
    // 이름표 저장 실패는 서버 저장 자체를 막을 정도로 중요하지 않으므로 조용히 무시한다.
  }
}
const ANONYMOUS_NAME_PLACEHOLDER = "(이름 미확인 - 이 기기)";
// 화면이 알고 있는 학생 이름을 이 기기의 로컬 이름표에 즉시(동기적으로) 반영한다.
// 이름을 입력/추가한 시점에 곧바로 호출해야 한다 — 네트워크 저장(Firestore 왕복)이 끝나길
// 기다렸다가 반영하면, 그 사이에 탭을 닫아버릴 경우 이름이 이 기기에 전혀 남지 않게 된다
// (localStorage.setItem 자체는 동기 함수라 즉시 호출하면 탭이 곧바로 닫혀도 안전하다).
function syncLocalNameMap(code, students) {
  if (!code || !Array.isArray(students)) return;
  const nameMap = loadLocalNameMap(code);
  let changed = false;
  students.forEach(s => {
    if (s.name && s.name !== ANONYMOUS_NAME_PLACEHOLDER && nameMap[s.id] !== s.name) {
      nameMap[s.id] = s.name;
      changed = true;
    }
  });
  if (changed) saveLocalNameMap(code, nameMap);
}

async function loadConfig(code) {
  try {
    const r = await storage.get(configKey(code), true);
    if (!r) return null;
    const obj = JSON.parse(r.value);
    if (obj && Array.isArray(obj.students)) {
      const nameMap = loadLocalNameMap(code);
      obj.students = obj.students.map(s => ({ ...s, name: nameMap[s.id] || s.name || ANONYMOUS_NAME_PLACEHOLDER }));
    }
    return obj;
  } catch (e) {
    return null;
  }
}
async function saveConfigRemote(code, obj) {
  try {
    let payload = obj;
    if (obj && Array.isArray(obj.students)) {
      // 저장 직전에, 이 화면이 알고 있는 이름을 이 기기의 이름표(로컬 저장소)에 먼저
      // 반영해 둔다 — 그래야 다음에 이 기기에서 다시 불러올 때도 이름이 유지된다.
      // (보통은 persistConfig에서 이미 더 일찍 동기적으로 반영해두지만, saveConfigRemote를
      // 직접 호출하는 다른 경로 — 백업 파일 불러오기 등 — 를 위해 여기서도 한 번 더 반영한다.)
      syncLocalNameMap(code, obj.students);
      // 서버(Firestore)로 나가는 값에는 이름 필드 자체를 아예 넣지 않는다.
      const anonStudents = obj.students.map(({ id, grade, classNum, number, gender }) => ({ id, grade, classNum, number, gender }));
      payload = { ...obj, students: anonStudents };
    }
    await storage.set(configKey(code), JSON.stringify(payload), true);
    return true;
  } catch (e) {
    return false;
  }
}
async function loadRecordsRemote(code) {
  try {
    const r = await storage.get(recordsKey(code), true);
    return r ? JSON.parse(r.value) : {};
  } catch (e) {
    return null;
  }
}
async function saveRecordsRemote(code, obj) {
  try {
    await storage.set(recordsKey(code), JSON.stringify(obj), true);
    return true;
  } catch (e) {
    return false;
  }
}

// 방문자 집계: 학교 코드와 무관하게, 이 프로그램을 여는 모든 사람의 방문(=화면이 켜진 횟수)을
// 전역 저장소 하나에 모아 센다. "고유 방문자"를 정확히 구분할 방법은 없어서(로그인 시스템이
// 없으므로), 화면이 열릴 때마다 1회로 집계하는 단순한 방식이다.
function todayKey() {
  const d = new Date();
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}
async function recordVisit() {
  try {
    const totalKey = "paps:stats:total";
    const dailyKey = "paps:stats:daily:" + todayKey();
    let total = 0, daily = 0;
    try { const r = await storage.get(totalKey, true); if (r) total = JSON.parse(r.value); } catch (e) { /* 없으면 0부터 */ }
    try { const r = await storage.get(dailyKey, true); if (r) daily = JSON.parse(r.value); } catch (e) { /* 없으면 0부터 */ }
    await storage.set(totalKey, JSON.stringify(total + 1), true);
    await storage.set(dailyKey, JSON.stringify(daily + 1), true);
  } catch (e) {
    // 통계 저장 실패는 프로그램 사용 자체를 막을 정도로 중요하지 않으므로 조용히 무시한다.
  }
}
async function loadVisitStats() {
  const totalKey = "paps:stats:total";
  const dailyKey = "paps:stats:daily:" + todayKey();
  let total = 0, daily = 0;
  try { const r = await storage.get(totalKey, true); if (r) total = JSON.parse(r.value); } catch (e) { /* 0 */ }
  try { const r = await storage.get(dailyKey, true); if (r) daily = JSON.parse(r.value); } catch (e) { /* 0 */ }
  return { total, daily };
}

// 변경 이력(감사 로그): 누가 어떤 학생의 어떤 기록을 언제 바꿨는지, 학생 추가/삭제나
// 접근권한 결정 같은 관리 행위도 함께 최근 300건까지 남긴다. 성적 이의제기 등 나중에
// "누가 언제 무엇을 했는지" 확인이 필요할 때를 위한 최소한의 기록이다.
function auditLogKey(code) { return "paps:auditlog:" + code; }
async function appendAuditLog(code, entry) {
  try {
    let list = [];
    try {
      const r = await storage.get(auditLogKey(code), true);
      if (r) list = JSON.parse(r.value);
    } catch (e) { /* 아직 로그가 없으면 새로 시작 */ }
    list.push(entry);
    if (list.length > 300) list = list.slice(list.length - 300);
    await storage.set(auditLogKey(code), JSON.stringify(list), true);
  } catch (e) {
    // 감사 로그 저장 실패는 실제 작업을 막을 정도로 중요하지 않으므로 조용히 무시한다.
  }
}
async function loadAuditLog(code) {
  try {
    const r = await storage.get(auditLogKey(code), true);
    return r ? JSON.parse(r.value) : [];
  } catch (e) {
    return [];
  }
}

// 백업(내보내기) 이력: JSON으로 언제, 누가 내보냈는지 최근 30건까지 남긴다. 학교(코드) 단위로
// 공유되어, 여러 선생님 중 누가 마지막으로 백업해뒀는지 서로 확인할 수 있다.
function backupLogKey(code) { return "paps:backuplog:" + code; }
async function appendBackupLog(code, entry) {
  try {
    let list = [];
    try {
      const r = await storage.get(backupLogKey(code), true);
      if (r) list = JSON.parse(r.value);
    } catch (e) { /* 아직 이력이 없으면 새로 시작 */ }
    list.push(entry);
    if (list.length > 30) list = list.slice(list.length - 30);
    await storage.set(backupLogKey(code), JSON.stringify(list), true);
  } catch (e) {
    // 백업 이력 저장 실패가 실제 백업 자체를 막을 정도로 중요하지는 않으므로 조용히 무시한다.
  }
}
async function loadBackupLog(code) {
  try {
    const r = await storage.get(backupLogKey(code), true);
    return r ? JSON.parse(r.value) : [];
  } catch (e) {
    return [];
  }
}

// 워크스페이스 코드는 "개인" 저장소(shared:false)에 저장한다 — 즉 이 값은
// 지금 로그인된 계정에만 연결되어, 같은 계정으로 다른 기기에서 열어도
// 자동으로 이어서 쓸 수 있다. 반면 학생 기록 자체(configKey/recordsKey)는
// "공유" 저장소를 쓰되 코드별로 키가 나뉘어 있어, 다른 코드를 쓰는 다른
// 학교/선생님과는 데이터가 섞이지 않는다.
async function loadSavedWorkspaceCode() {
  try {
    const r = await storage.get(WORKSPACE_CODE_KEY, false);
    return r ? r.value : null;
  } catch (e) {
    return null;
  }
}
async function saveWorkspaceCodeRemote(code) {
  try {
    await storage.set(WORKSPACE_CODE_KEY, code, false);
    return true;
  } catch (e) {
    return false;
  }
}
// 마감(학교 코드 완전 삭제) 시, 이 기기가 "마지막으로 쓰던 코드"로 기억해 둔 값도 함께
// 지운다. 이걸 지우지 않으면, 마감 직후에는 첫 화면으로 잘 돌아가더라도 앱을 나갔다가
// 다시 열 때(특히 모바일에서 브라우저/PWA를 새로 열 때) 이 기기가 방금 지운 코드를
// 자동으로 다시 불러와 그 코드로 재접속을 시도하게 된다.
async function clearSavedWorkspaceCode() {
  try {
    await storage.delete(WORKSPACE_CODE_KEY, false);
    return true;
  } catch (e) {
    return false;
  }
}

// ---- 접근 요청(열람 신청) 관련 저장소 ----
// 열람 신청자는 관리자에게 안내받은 "학교 코드"와 "열람 비밀번호"를 직접
// 입력해야 하므로, 별도의 숨겨진 포인터 없이 입력받은 코드로 바로 해당
// 워크스페이스의 설정(비밀번호)을 확인한다.
function accessKey(code) { return "paps:access:" + code; }
function deviceKey(code) { return "paps:device:" + code; }

async function loadAccessList(code) {
  try {
    const r = await storage.get(accessKey(code), true);
    return r ? JSON.parse(r.value) : { requests: [] };
  } catch (e) {
    return { requests: [] };
  }
}
async function saveAccessList(code, list) {
  try {
    await storage.set(accessKey(code), JSON.stringify(list), true);
    return true;
  } catch (e) {
    return false;
  }
}
// 기기(계정)별 접근 역할은 "개인" 저장소에 저장한다 — 본인 계정에만 연결되어
// 다른 사람의 승인 상태에 영향을 주지 않는다.
async function loadDeviceRole(code) {
  try {
    const r = await storage.get(deviceKey(code), false);
    return r ? JSON.parse(r.value) : null;
  } catch (e) {
    return null;
  }
}
async function saveDeviceRole(code, obj) {
  try {
    await storage.set(deviceKey(code), JSON.stringify(obj), false);
    return true;
  } catch (e) {
    return false;
  }
}
async function clearDeviceRole(code) {
  try {
    await storage.delete(deviceKey(code), false);
    return true;
  } catch (e) {
    return false;
  }
}

/* ============================== 메인 앱 ============================== */

export default function PapsApp({ initialWorkspaceCode = null, forcePresentation = false } = {}) {
  const [workspaceCode, setWorkspaceCode] = useState(initialWorkspaceCode);
  const [gateInitialMode, setGateInitialMode] = useState("code");
  const pendingSchoolLevelRef = useRef("middle"); // 새 코드를 만들 때 고른 학교급(기존 코드면 무시됨)
  const pendingViewerPasswordRef = useRef(""); // 새 코드를 만들 때 함께 정한 접근 신청 비밀번호
  const pendingFounderPasswordRef = useRef(""); // 새 코드를 만들 때 정한(또는 기존 코드 재접속 시 입력한) 개설자 전용 비밀번호
  // 방금 제출된 코드가 "새 코드 만들기"로 들어온 건지("create"), "코드로 로그인"으로
  // 들어온 건지("login")를 기억해 둔다. 아직 아무도 만든 적 없는(또는 마감으로 방금
  // 삭제된) 코드일 때, "새 코드 만들기"로 들어온 경우에만 지금 이 사람을 새 개설자로
  // 만들고, "코드로 로그인"으로 들어왔거나(사람이 직접 입력하지 않고) 이 기기가 예전에
  // 기억해 둔 코드를 자동으로 다시 불러온 경우에는 절대 새로 만들지 않는다 — 이걸 구분하지
  // 않으면, 마감으로 지워진 코드를 다시 열었을 때(특히 이 기기가 예전 코드를 기억하고
  // 있다가 자동으로 재접속을 시도할 때) 마치 아무 일도 없었다는 듯 새 빈 코드가 조용히
  // 다시 만들어지는 문제가 있었다.
  const pendingIntentRef = useRef("login");
  const [workspaceChecking, setWorkspaceChecking] = useState(true);
  const [role, setRole] = useState(null); // 'admin' | 'viewer' | 'pending' | 'blocked'
  const [isFounder, setIsFounder] = useState(false);
  const [myDisplayName, setMyDisplayName] = useState("개설자");
  const [myDeviceId, setMyDeviceId] = useState(null);
  const [roleChecking, setRoleChecking] = useState(false);
  const [blockReason, setBlockReason] = useState(null); // 'denied' | 'revoked'
  const [myRequest, setMyRequest] = useState(null); // { id, name, type } while pending
  const [accessList, setAccessList] = useState({ requests: [] });
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(false);
  const [students, setStudents] = useState([]);
  const [criteria, setCriteria] = useState(null);
  const [settings, setSettings] = useState({ schoolName: "", currentYear: thisYear() });
  const [records, setRecords] = useState({});
  // 나이스 반영(NeisTemplateFiller)에서 첨부한 파일·인식결과·반영결과. 데이터백업 탭을
  // 벗어났다가 다시 돌아오면 그 탭 내부 컴포넌트가 다시 마운트되면서 로컬 state가 초기화돼
  // 방금 첨부·반영한 내용이 사라져 버리는 문제가 있었다. 세션 내내 유지되는 여기(최상위
  // 컴포넌트)에 상태를 두어, 탭을 오가도 그대로 남아있게 한다.
  const [neisFillState, setNeisFillState] = useState({
    fileName: "", headerRow: null, dataRows: null, mapping: [], resultRows: null, sheetName: "Sheet1",
  });
  const [view, setView] = useState("board");
  const [presentation, setPresentation] = useState(false);
  const [lastSync, setLastSync] = useState(null);
  const [toast, setToast] = useState(null);
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const [activeYear, setActiveYear] = useState(thisYear());
  const [codeCreatedAt, setCodeCreatedAt] = useState(null);
  const [justCreatedNotice, setJustCreatedNotice] = useState(false);
  const [shareModalOpen, setShareModalOpen] = useState(false);
  const [presentationUnlocked, setPresentationUnlocked] = useState(false);
  const [lockPromptOpen, setLockPromptOpen] = useState(false);
  const [lockPromptError, setLockPromptError] = useState("");
  // 공용 PC 자동 잠금(세션 타임아웃): 일정 시간 조작이 없으면 화면을 잠근다.
  const [idleLocked, setIdleLocked] = useState(false);
  const [idleLockError, setIdleLockError] = useState("");
  const lastActivityRef = useRef(Date.now());

  const prevTopRef = useRef({});
  const configVersionRef = useRef(0);

  // 방문 집계: 이 화면이 열릴 때마다 한 번, 학교 코드와 무관하게 전역 통계에 더한다.
  // (개발자용 통계 화면에서만 확인 가능 — 아래 WorkspaceGate의 숨겨진 접근점 참고)
  useEffect(() => {
    recordVisit();
  }, []);

  // 빔프로젝터에 띄우는 순간부터 자동으로 잠긴다(별도 켜기/끄기 설정 없음). 나가려면 접근 신청
  // 비밀번호를 입력해야 한다. 비밀번호가 아예 설정되어 있지 않다면 잠글 방법이 없으니 그냥 나간다.
  function requestExitPresentation() {
    if (settings.viewerPassword && !presentationUnlocked) {
      setLockPromptError("");
      setLockPromptOpen(true);
      return;
    }
    setPresentation(false);
  }

  function handleUnlockAttempt(pw) {
    if (pw && settings.viewerPassword && pw === settings.viewerPassword) {
      setPresentationUnlocked(true);
      setLockPromptOpen(false);
      setLockPromptError("");
      setPresentation(false);
    } else {
      setLockPromptError("비밀번호가 올바르지 않습니다.");
    }
  }

  // 공용 PC(체육관·교무실 공용 컴퓨터 등)에 로그인된 채로 방치되어 다른 사람이 학생
  // 개인정보에 접근하는 것을 막기 위한 자동 잠금. 로그인(코드 확인)이 끝난 뒤부터, 화면
  // 조작(마우스·키보드·터치·스크롤)이 일정 시간(12분) 없으면 자동으로 잠긴다. 전광판
  // 모드는 애초에 계속 켜두는 용도이고 이미 자체 잠금이 있으므로 이 타이머 대상에서
  // 제외한다. 접근 신청 비밀번호가 설정되어 있지 않으면(=잠글 방법이 없으면) 아예
  // 타이머를 켜지 않는다(그렇지 않으면 아무도 못 푸는 상태로 잠겨버릴 수 있다).
  const IDLE_TIMEOUT_MS = 12 * 60 * 1000;
  useEffect(() => {
    if (presentation) return;
    if (role !== "admin" && role !== "viewer") return;
    if (!settings.viewerPassword) return;

    lastActivityRef.current = Date.now();
    const markActive = () => { lastActivityRef.current = Date.now(); };
    const activityEvents = ["mousemove", "mousedown", "keydown", "touchstart", "scroll", "wheel"];
    activityEvents.forEach(ev => window.addEventListener(ev, markActive, { passive: true }));

    const timer = setInterval(() => {
      if (Date.now() - lastActivityRef.current >= IDLE_TIMEOUT_MS) {
        setIdleLocked(true);
      }
    }, 15000);

    return () => {
      activityEvents.forEach(ev => window.removeEventListener(ev, markActive));
      clearInterval(timer);
    };
  }, [role, presentation, settings.viewerPassword]);

  function handleIdleUnlockAttempt(pw) {
    if (pw && settings.viewerPassword && pw === settings.viewerPassword) {
      setIdleLocked(false);
      setIdleLockError("");
      lastActivityRef.current = Date.now();
    } else {
      setIdleLockError("비밀번호가 올바르지 않습니다.");
    }
  }

  /* ---------- 워크스페이스 코드 확인 (계정별 개인 저장소) ---------- */
  useEffect(() => {
    (async () => {
      if (typeof window === "undefined") {
        setWorkspaceChecking(false);
        setLoadError(true);
        setLoading(false);
        return;
      }
      // 빔프로젝터 새 창(?present=1&code=...)으로 열린 경우에는 저장된(다를 수 있는) 코드를
      // 덮어쓰지 않고, 주소창에 실려온 코드를 그대로 쓴다.
      if (initialWorkspaceCode) {
        setWorkspaceChecking(false);
        return;
      }
      const saved = await loadSavedWorkspaceCode();
      if (saved) setWorkspaceCode(saved);
      setWorkspaceChecking(false);
    })();
  }, []);

  function submitWorkspaceCode(raw, schoolLevel, viewerPassword, founderPassword, intent) {
    const code = sanitizeWorkspaceCode(raw);
    if (!code) return;
    pendingSchoolLevelRef.current = schoolLevel || "middle";
    pendingViewerPasswordRef.current = (viewerPassword || "").trim();
    pendingFounderPasswordRef.current = (founderPassword || "").trim();
    pendingIntentRef.current = intent === "create" ? "create" : "login";
    saveWorkspaceCodeRemote(code);
    setWorkspaceCode(code);
  }

  function changeWorkspaceCode() {
    setWorkspaceCode(null);
    setRole(null);
    setLoading(true);
  }

  /* ---------- 접근 권한(역할) 확인 ---------- */
  useEffect(() => {
    if (!workspaceCode) return;
    let cancelled = false;
    setRoleChecking(true);
    (async () => {
      const device = await loadDeviceRole(workspaceCode);
      if (cancelled) return;
      if (device) setMyDeviceId(device.id);

      // 이 코드가 실제로 아직 존재하는지 먼저 확인한다. 마감으로 방금 삭제된 코드처럼, 이
      // 기기에 예전 역할이 캐시되어 있거나(예전에 개설자·조회자·수정권한자였던 기기) 이
      // 기기가 예전에 쓰던 코드를 자동으로 다시 불러온 경우, 설정(cfg)이 이미 사라졌다면
      // 그 캐시된 역할을 그대로 믿어서는 안 된다. 이걸 확인하지 않으면, 마감으로 지워진
      // 코드를 다시 열었을 때(특히 모바일에서 앱을 나갔다가 다시 열어 이 기기가 예전 코드를
      // 자동으로 재접속 시도할 때) 아무 확인 절차 없이 조용히 새 빈 코드가 다시 만들어지거나
      // (개설자였던 기기) 예전 권한이 유령처럼 되살아나 버리는(조회자·수정권한자였던 기기)
      // 문제가 있었다.
      const cfg = await loadConfig(workspaceCode);
      if (cancelled) return;

      if (!cfg) {
        if (!device && pendingIntentRef.current === "create") {
          // "새 코드 만들기"로 직접 제출한 경우에만, 지금 이 사람을 새 개설자로 만든다.
          const newId = uid("dev");
          await saveDeviceRole(workspaceCode, { id: newId, role: "admin" });
          setMyDeviceId(newId);
          setIsFounder(true);
          setMyDisplayName("개설자");
          setRole("admin");
          setRoleChecking(false);
          return;
        }
        // 그 외의 경우(존재하지 않는 코드로 "코드로 로그인"을 시도했거나, 이 기기가 예전에
        // 쓰던 코드를 자동으로 다시 불러왔거나, 이 기기가 예전에 이 코드로 뭔가 권한을
        // 가졌었지만 지금은 코드 자체가 없는 경우)에는 조용히 새로 만들지 않는다. 이 기기에
        // 남은 캐시(역할·마지막으로 쓰던 코드로 기억해 둔 값)를 정리해, 다음에 또 자동으로
        // 이 코드를 불러오는 일이 없게 한다.
        if (device) await clearDeviceRole(workspaceCode);
        await clearSavedWorkspaceCode();
        if (cancelled) return;
        setRole("blocked");
        setBlockReason("no-such-code");
        setRoleChecking(false);
        return;
      }

      if (device && device.role === "admin") {
        // 이 기기가 "최초 개설자"인지 "승인받은 수정 권한자"인지 구분한다. 접근권한 목록에
        // 내 항목이 있으면(=신청해서 승인받은 사람) 개설자가 아니라 제한된 관리자다.
        const list = await loadAccessList(workspaceCode);
        const mine = list.requests.find(r => r.id === device.id);
        setIsFounder(!mine);
        setMyDisplayName(mine ? mine.name : "개설자");
        setAccessList(list);
        setRole("admin");
        setRoleChecking(false);
        return;
      }

      if (device && (device.role === "viewer" || device.role === "pending")) {
        const list = await loadAccessList(workspaceCode);
        const mine = list.requests.find(r => r.id === device.id);
        if (mine && mine.status === "approved") {
          const grantedRole = mine.type === "editor" ? "admin" : "viewer";
          await saveDeviceRole(workspaceCode, { id: device.id, role: grantedRole });
          setIsFounder(false);
          setMyDisplayName(mine.name);
          setRole(grantedRole);
          setAccessList(list);
        } else if (mine && mine.status === "denied") {
          setRole("blocked");
          setBlockReason("denied");
        } else if (mine && mine.status === "pending") {
          setRole("pending");
          setMyRequest({ id: mine.id, name: mine.name, type: mine.type });
        } else {
          // 요청 기록을 찾을 수 없음(관리자가 삭제 등) — 취소된 것으로 처리
          await clearDeviceRole(workspaceCode);
          setRole("blocked");
          setBlockReason("revoked");
        }
        setRoleChecking(false);
        return;
      }

      // 이 기기(계정)에 저장된 역할은 없지만, 코드 자체는 이미 존재하는 경우(위에서 cfg
      // 확인으로 "존재하지 않는 코드"는 이미 걸러졌다). 단, 이번에 함께 입력한 값이
      // "개설자 전용 비밀번호"와 정확히 일치하면(개설자 본인이 다른 기기로 넘어온 경우),
      // 별도 승인 절차 없이 곧바로 개설자로 인정한다. 그 외에는 "접근 신청" 절차를 거쳐야 한다.
      const founderPw = cfg.settings?.founderPassword;
      if (founderPw && pendingFounderPasswordRef.current && pendingFounderPasswordRef.current === founderPw) {
        const newId = uid("dev");
        await saveDeviceRole(workspaceCode, { id: newId, role: "admin" });
        setMyDeviceId(newId);
        setIsFounder(true);
        setMyDisplayName("개설자");
        setRole("admin");
        setRoleChecking(false);
        return;
      }

      setRole("blocked");
      setBlockReason("need-request");
      setRoleChecking(false);
    })();
    return () => { cancelled = true; };
  }, [workspaceCode]);

  /* ---------- 열람 신청: 학교 코드 + 열람 비밀번호로 즉시 확인 ---------- */
  // 비밀번호를 무작위로 계속 시도하는 것을 막기 위해, 이 기기에서 같은 학교 코드로 틀린
  // 비밀번호를 반복 입력하면 잠시 잠긴다(완전한 서버 차단은 아니지만 최소한의 안전장치).
  async function checkPasswordAttempt(code) {
    const key = "paps:pwattempts:" + code;
    const WINDOW_MS = 10 * 60 * 1000; // 10분
    const MAX_ATTEMPTS = 5;
    let data = { count: 0, windowStart: Date.now() };
    try {
      const r = await storage.get(key, false);
      if (r) data = JSON.parse(r.value);
    } catch (e) { /* 기록 없음 */ }
    if (Date.now() - data.windowStart > WINDOW_MS) data = { count: 0, windowStart: Date.now() };
    if (data.count >= MAX_ATTEMPTS) return { blocked: true, key, data };
    return { blocked: false, key, data };
  }
  async function recordPasswordFailure(key, data) {
    const next = { count: data.count + 1, windowStart: data.windowStart };
    await storage.set(key, JSON.stringify(next), false);
  }
  async function clearPasswordAttempts(key) {
    try { await storage.delete(key, false); } catch (e) { /* 없어도 무방 */ }
  }

  async function requestAccess({ name, code, password, wantsEdit }) {
    const sanitizedCode = sanitizeWorkspaceCode(code);
    if (!sanitizedCode) return { ok: false, reason: "bad-code" };
    const attempt = await checkPasswordAttempt(sanitizedCode);
    if (attempt.blocked) return { ok: false, reason: "too-many-attempts" };
    const cfg = await loadConfig(sanitizedCode);
    if (!cfg) return { ok: false, reason: "bad-code" };
    const setPassword = cfg.settings?.viewerPassword;
    if (!setPassword) return { ok: false, reason: "no-password-set" };
    if (password !== setPassword) {
      await recordPasswordFailure(attempt.key, attempt.data);
      return { ok: false, reason: "bad-password" };
    }
    await clearPasswordAttempts(attempt.key);

    const list = await loadAccessList(sanitizedCode);
    const trimmedName = name.trim();
    const type = wantsEdit ? "editor" : "viewer";

    // 이미 같은 이름 + 같은 권한 종류로 승인받은 계정이 있다면(브라우저 데이터가 지워졌거나
    // 새 기기로 바꾼 경우일 가능성이 높음), 다시 승인을 기다리지 않고 그 기존 승인을 이
    // 기기로 그대로 이어받는다. 이름과 비밀번호가 둘 다 맞아야 하므로 위험은 낮다.
    const existing = list.requests.find(r => r.status === "approved" && r.name === trimmedName && r.type === type);
    if (existing) {
      await saveDeviceRole(sanitizedCode, { id: existing.id, role: type === "editor" ? "admin" : "viewer" });
      await saveWorkspaceCodeRemote(sanitizedCode);
      setWorkspaceCode(sanitizedCode);
      return { ok: true, reclaimed: true };
    }

    const id = uid("req");
    // 조회는 비밀번호 확인만으로 즉시 승인되지만, 데이터 수정 권한은 최초 개설자의
    // 별도 승인이 있어야 부여된다.
    const status = wantsEdit ? "pending" : "approved";
    const entry = { id, name: trimmedName, submittedAt: Date.now(), status, type };
    await saveAccessList(sanitizedCode, { requests: [...list.requests, entry] });
    await saveDeviceRole(sanitizedCode, { id, role: wantsEdit ? "pending" : "viewer" });
    await saveWorkspaceCodeRemote(sanitizedCode);
    setWorkspaceCode(sanitizedCode);
    return { ok: true };
  }

  function retryAfterBlock(nextMode) {
    setRole(null);
    setBlockReason(null);
    setMyRequest(null);
    setWorkspaceCode(null);
    setGateInitialMode(nextMode || "code");
  }

  /* ---------- 관리자: 접근 요청 관리 ---------- */
  const refreshAccessList = useCallback(async () => {
    if (!workspaceCode) return;
    const list = await loadAccessList(workspaceCode);
    setAccessList(list);
  }, [workspaceCode]);

  async function decideRequest(id, status, actionLabel) {
    const list = await loadAccessList(workspaceCode);
    const target = list.requests.find(r => r.id === id);
    const nextList = { requests: list.requests.map(r => r.id === id ? { ...r, status } : r) };
    await saveAccessList(workspaceCode, nextList);
    setAccessList(nextList);
    if (target) {
      appendAuditLog(workspaceCode, {
        type: "access",
        ts: Date.now(),
        by: myDisplayName,
        message: `${myDisplayName}님이 ${target.name}님의 접근 신청을 ${actionLabel}`,
      });
    }
  }
  const approveRequest = (id) => decideRequest(id, "approved", "승인함");
  const denyRequest = (id) => decideRequest(id, "denied", "거절함");
  const revokeApproval = (id) => decideRequest(id, "denied", "취소함(권한 회수)");
  const restoreRequest = (id) => decideRequest(id, "approved", "복구함(다시 승인)");

  // 거절/취소된 신청을 목록에서 아예 지운다(실수로 거절한 신청이 목록에 계속 남아있는 게
  // 번거롭다는 요청 반영). 복구와 달리 삭제는 되돌릴 수 없다.
  async function deleteRequestEntry(id) {
    const list = await loadAccessList(workspaceCode);
    const target = list.requests.find(r => r.id === id);
    const nextList = { requests: list.requests.filter(r => r.id !== id) };
    await saveAccessList(workspaceCode, nextList);
    setAccessList(nextList);
    if (target) {
      appendAuditLog(workspaceCode, {
        type: "access",
        ts: Date.now(),
        by: myDisplayName,
        message: `${myDisplayName}님이 ${target.name}님의 거절/취소 기록을 삭제함`,
      });
    }
  }

  // 개설자 승계: 예를 들어 담당 선생님이 전근을 가는 경우, 승인된 편집자 중 한 명에게
  // "개설자" 자리를 넘겨준다. 넘겨준 뒤에는 나(원래 개설자)도 평범한 승인 편집자가 되어,
  // 승인/거절/취소·비밀번호 변경·마감 같은 민감한 작업은 더 이상 할 수 없게 된다.
  async function transferFounder(targetRequestId) {
    if (!isFounder || !myDeviceId) return { ok: false };
    const list = await loadAccessList(workspaceCode);
    const target = list.requests.find(r => r.id === targetRequestId);
    if (!target || target.status !== "approved") return { ok: false };
    // 넘겨받을 사람의 신청 항목은 제거한다(항목이 없어지면 다음 확인 때 그 사람이 개설자가 된다).
    const withoutTarget = list.requests.filter(r => r.id !== targetRequestId);
    // 나 자신은 이제 평범한 승인 편집자로 등록해 둔다.
    const meAsEditor = { id: myDeviceId, name: myDisplayName === "개설자" ? "이전 개설자" : myDisplayName, submittedAt: Date.now(), status: "approved", type: "editor" };
    const nextList = { requests: [...withoutTarget, meAsEditor] };
    await saveAccessList(workspaceCode, nextList);
    setAccessList(nextList);
    setIsFounder(false);
    setMyDisplayName(meAsEditor.name);
    return { ok: true, newFounderName: target.name };
  }

  /* ---------- 초기 로드 (권한이 확인된 뒤) ---------- */
  useEffect(() => {
    if (!workspaceCode || (role !== "admin" && role !== "viewer")) return;
    setLoading(true);
    (async () => {
      // loadConfig/loadRecordsRemote return null both when the key doesn't
      // exist yet (first-ever use) and when the read genuinely failed —
      // either way the correct move is to fall back to empty/default data.
      const cfg = await loadConfig(workspaceCode);
      const recs = await loadRecordsRemote(workspaceCode);
      const finalCfg = cfg || buildDefaultConfig(pendingSchoolLevelRef.current, pendingViewerPasswordRef.current, pendingFounderPasswordRef.current);
      const backfilledCreatedAt = !finalCfg.createdAt;
      if (backfilledCreatedAt) finalCfg.createdAt = Date.now(); // 이전 버전에 만들어진 코드는 지금부터 1년을 새로 센다

      // 1년 넘게 방치된 코드는 마감을 깜빡 잊은 것으로 보고 자동으로 마감(전체 삭제)
      // 처리한다. 기록이 쓸데없이 계속 쌓여 프로그램이 느려지는 것을 막기 위한 안전망이다.
      const ONE_YEAR_MS = 365 * 24 * 60 * 60 * 1000;
      if (role === "admin" && cfg && Date.now() - finalCfg.createdAt > ONE_YEAR_MS) {
        await storage.delete(recordsKey(workspaceCode), true).catch(() => {});
        await storage.delete(configKey(workspaceCode), true).catch(() => {});
        await storage.delete(accessKey(workspaceCode), true).catch(() => {});
        await storage.delete(auditLogKey(workspaceCode), true).catch(() => {});
        await storage.delete(backupLogKey(workspaceCode), true).catch(() => {});
        await storage.delete(deviceKey(workspaceCode), false).catch(() => {});
        try { window.localStorage.removeItem(nameMapKey(workspaceCode)); } catch (e) {}
        showToast("이 코드가 개설된 지 1년이 지나 자동으로 마감(전체 삭제)되었습니다.", "warn");
        setWorkspaceCode(null);
        setRole(null);
        setLoading(true);
        return;
      }
      if (!finalCfg.settings) finalCfg.settings = { schoolName: "" };
      if (!finalCfg.settings.currentYear) finalCfg.settings.currentYear = thisYear();
      const year = finalCfg.settings.currentYear;

      // 이전 버전(연도 구분 없는) 기록이 있다면 올해 기록으로 변환한다.
      const { records: migratedRecords, migrated } = migrateLegacyRecords(recs || {}, finalCfg.students || [], year);

      if ((!cfg || migrated || backfilledCreatedAt) && role === "admin") await saveConfigRemote(workspaceCode, finalCfg);
      if (migrated && role === "admin") await saveRecordsRemote(workspaceCode, migratedRecords);
      if (!cfg && role === "admin") setJustCreatedNotice(true);

      setStudents(finalCfg.students || []);
      setCriteria(finalCfg.criteria || buildDefaultCriteria(finalCfg.settings?.schoolLevel));
      setSettings(finalCfg.settings);
      setCodeCreatedAt(finalCfg.createdAt);
      setRecords(migratedRecords);
      // 연도 선택은 저장된 값이 아니라 "지금 실제 연도"로 항상 자동 설정한다. 저장된 값을
      // 그대로 쓰면, 작년에 만든 워크스페이스를 올해 다시 열었을 때 여전히 작년으로 남아있게 된다.
      setActiveYear(thisYear());
      setLastSync(Date.now());
      setLoading(false);
      if (role === "admin") refreshAccessList();
    })();
  }, [workspaceCode, role, refreshAccessList]);

  /* ---------- 빔프로젝터 새 창: 데이터가 준비되면 바로 전광판(잠금) 모드로 진입 ---------- */
  // "빔프로젝터 고정모드" 버튼은 지금 창을 그대로 전체화면으로 바꾸는 대신, 이 화면을
  // ?present=1&code=... 로 새 창에 띄운다. 그 새 창에서는(=forcePresentation) 데이터
  // 로딩이 끝나자마자 자동으로 전광판 화면으로 들어가고 자동으로 잠긴다. 원래 창은 이
  // 효과와 무관하므로 그대로 다른 작업을 계속할 수 있다.
  useEffect(() => {
    if (!forcePresentation) return;
    if (loading || roleChecking) return;
    if (role !== "admin" && role !== "viewer") return;
    setView("board");
    setPresentation(true);
  }, [forcePresentation, loading, roleChecking, role]);

  /* ---------- 폴링 (실시간 동기화 + 권한 재확인) ---------- */
  useEffect(() => {
    if (!workspaceCode || (role !== "admin" && role !== "viewer")) return;
    const interval = setInterval(async () => {
      const recs = await loadRecordsRemote(workspaceCode);
      if (recs !== null) {
        setRecords(prevLocal => mergeRecords(prevLocal, recs));
        setLastSync(Date.now());
      }
      // 편집 화면이 아닐 때만 설정/명단을 통째로 동기화 (편집 중 덮어쓰기 방지)
      if (view !== "roster" && view !== "access" && view !== "grades" && view !== "records") {
        const cfg = await loadConfig(workspaceCode);
        if (cfg) {
          setStudents(cfg.students || []);
          setCriteria(cfg.criteria || buildDefaultCriteria(cfg.settings?.schoolLevel));
          setSettings(cfg.settings || { schoolName: "", currentYear: thisYear() });
        }
      } else {
        // 이 화면들에 머무는 동안에도, 다른 선생님이 그 사이에 새로 등록한 학생만큼은 놓치지
        // 않도록 "새로 생긴 학생"만 추가로 반영한다. 이미 알고 있는 학생 정보는 손대지 않아,
        // 지금 편집 중인 내용을 덮어쓰지 않는다.
        const cfg = await loadConfig(workspaceCode);
        if (cfg && cfg.students) {
          setStudents(prevStudents => {
            const knownIds = new Set(prevStudents.map(s => s.id));
            const newOnes = cfg.students.filter(s => !knownIds.has(s.id));
            return newOnes.length > 0 ? [...prevStudents, ...newOnes] : prevStudents;
          });
        }
      }
      if (role === "admin") {
        refreshAccessList();
        // 이 기기가 "승인받아 수정 권한을 받은 동료 교사"(최초 개설자 본인이 아님)라면,
        // 그 승인이 취소되지는 않았는지 주기적으로 확인한다. 또한 개설자 승계가 일어나서
        // 내 항목이 목록에서 사라졌다면(=내가 새 개설자가 됐다면) isFounder도 갱신한다.
        const device = await loadDeviceRole(workspaceCode);
        if (device) {
          const list = await loadAccessList(workspaceCode);
          const mine = list.requests.find(r => r.id === device.id);
          if (mine && mine.status !== "approved") {
            await clearDeviceRole(workspaceCode);
            setRole("blocked");
            setBlockReason("revoked");
          } else {
            setIsFounder(!mine);
          }
        }
      } else if (role === "viewer") {
        // 조회 권한 보유자는 관리자가 승인을 취소했는지 주기적으로 확인한다.
        const device = await loadDeviceRole(workspaceCode);
        if (device) {
          const list = await loadAccessList(workspaceCode);
          const mine = list.requests.find(r => r.id === device.id);
          if (mine && mine.status !== "approved") {
            await clearDeviceRole(workspaceCode);
            setRole("blocked");
            setBlockReason("revoked");
          }
        }
      }
    }, 10000);
    return () => clearInterval(interval);
  }, [view, workspaceCode, role, refreshAccessList]);

  /* ---------- 대기 화면 폴링 ---------- */
  useEffect(() => {
    if (role !== "pending" || !workspaceCode) return;
    const interval = setInterval(async () => {
      const list = await loadAccessList(workspaceCode);
      const mine = list.requests.find(r => r.id === myRequest?.id);
      if (mine && mine.status === "approved") {
        // 승인된 것이 "수정 권한(editor)" 요청이었다면 관리자로, 아니면 조회 전용으로 부여한다.
        // (이 부분에서 type을 확인하지 않아 수정 권한 승인자가 조회 전용으로 잘못 들어가던 버그를 수정함)
        const grantedRole = mine.type === "editor" ? "admin" : "viewer";
        await saveDeviceRole(workspaceCode, { id: mine.id, role: grantedRole });
        setRole(grantedRole);
      } else if (mine && mine.status === "denied") {
        setRole("blocked");
        setBlockReason("denied");
      }
    }, 4000);
    return () => clearInterval(interval);
  }, [role, workspaceCode, myRequest]);

  /* ---------- 저장 헬퍼 ---------- */
  // 조회 전용(viewer) 권한은 어떤 경우에도 데이터를 쓰지 않도록 이중으로 막는다
  // (화면 자체에서도 입력/관리 탭을 보여주지 않지만, 안전장치를 하나 더 둔다).
  const persistConfig = useCallback((nextStudents, nextCriteria, nextSettings) => {
    if (role !== "admin") return Promise.resolve();
    // 이름을 입력/추가하자마자, 서버 저장(아래 큐)을 기다리지 않고 이 기기의 로컬 이름표에
    // 곧바로(동기적으로) 반영해 둔다. 그렇지 않으면 아직 네트워크 왕복(loadConfig→
    // saveConfigRemote) 중일 때 탭을 닫아버릴 경우 방금 입력한 이름이 이 기기에 전혀
    // 저장되지 않아, 다시 열었을 때 "(이름 미확인 - 이 기기)"로 보이는 문제가 생긴다.
    if (nextStudents !== undefined) syncLocalNameMap(workspaceCode, nextStudents);
    // 저장 직전에 서버의 최신 값을 한 번 더 받아와서, 지금 바꾸는 항목이 아닌 나머지는
    // 내 화면에 캐시된(어쩌면 오래된) 값이 아니라 최신 값을 그대로 유지한다. 엑셀 파일을
    // 연달아 여러 번 드래그하는 등 이 저장이 겹쳐 호출될 수 있으므로, 기록 저장과 같은
    // 큐를 통해 반드시 순서대로(하나씩) 진행되도록 한다 — 그렇지 않으면 나중에 시작했지만
    // 먼저 끝난 저장이 최신 저장을 덮어써서, 방금 추가한 학생이 사라지는 문제가 생긴다.
    return enqueueSave(async () => {
      const latest = await loadConfig(workspaceCode);
      const cfg = {
        ...(latest || {}),
        students: nextStudents !== undefined ? nextStudents : (latest?.students ?? students),
        criteria: nextCriteria !== undefined ? nextCriteria : (latest?.criteria ?? criteria),
        settings: nextSettings !== undefined ? nextSettings : (latest?.settings ?? settings),
      };
      await saveConfigRemote(workspaceCode, cfg);
    });
  }, [students, criteria, settings, workspaceCode, role]);

  const persistRecords = useCallback(async (nextRecords) => {
    if (role !== "admin") return;
    await saveRecordsRemote(workspaceCode, nextRecords);
  }, [workspaceCode, role]);

  // 저장 큐: 여러 학생 기록을 빠르게 연속으로 입력하면, 각각의 "저장소에서 읽어와 → 값
  // 얹어서 → 다시 저장" 과정이 서로 겹쳐(비동기 경합) 나중에 끝난 저장이 먼저 끝난 저장을
  // 덮어써 버릴 수 있었다(실제로 이 문제로 일부 기록이 사라지는 현상이 있었다). 이를 막기
  // 위해 이 기기 안에서의 모든 기록 저장은 반드시 한 번에 하나씩, 순서대로만 진행되도록
  // 큐에 연결한다 — 앞 저장이 완전히 끝난 뒤에야 다음 저장이 시작된다.
  const saveQueueRef = useRef(Promise.resolve());
  function enqueueSave(task) {
    const chained = saveQueueRef.current.then(task, task);
    saveQueueRef.current = chained.catch(() => {});
    return chained;
  }

  const saveNewRecord = useCallback((studentId, eventId, value, year, schoolGradeAtMeasure, parts) => {
    if (role !== "admin") return Promise.resolve();
    return enqueueSave(async () => {
      // 오타로 보이는 극단적인 값은 저장 자체를 막지는 않되(진짜 특이한 기록일 수도 있으니),
      // 한 번 더 확인해보시라고 즉시 알려준다.
      const plausible = isPlausibleValue(eventId, value) && (eventId !== "bmi" || isPlausibleBmiParts(parts));
      if (!plausible) {
        setToast({ text: "입력하신 값이 일반적인 범위를 크게 벗어났어요. 오타는 아닌지 한 번 더 확인해 주세요.", kind: "warn", id: uid("toast") });
        setTimeout(() => setToast(t => (t && t.id) ? null : t), 4000);
      }
      const entry = { value, schoolGradeAtMeasure, updatedAt: Date.now() };
      // 큐 덕분에 이 시점에는 같은 기기의 이전 저장이 이미 완전히 끝나 있으므로, 저장소의
      // 최신 값을 그대로 기준으로 삼아도 안전하다(다른 기기의 거의 동시 저장까지 대비해
      // 화면에 캐시된 값도 함께 참고한다).
      const latest = (await loadRecordsRemote(workspaceCode)) || {};
      const base = mergeRecords(records, latest);
      const key = recKey(studentId, eventId, year);
      const prevValue = base[key]?.value;
      if (parts) {
        // 제자리멀리뛰기·앉아윗몸앞으로굽히기(1차/2차), 악력(1차/2차×좌우), BMI(신장/체중)처럼
        // 한 종목 기록이 여러 하위 값으로 이뤄진 경우, 이 화면(컴포넌트)이 열려 있는 동안
        // 다른 기기에서 그중 다른 하위 값만 먼저 저장했을 수 있다. 이때 이 화면에 아직 반영 안
        // 된(값이 비어있는) 하위 항목까지 통째로 덮어써버리면, 방금 다른 기기가 저장한 값이
        // 사라져버린다("1차는 있는데 2차가 어느 순간 없어짐" 같은 증상의 원인). 그래서 방금
        // 서버에서 새로 받아온 이전 값을 바탕으로, 이번에 실제로 입력된(비어있지 않은) 하위
        // 값만 덮어씌운다.
        const prevParts = base[key]?.parts || {};
        const mergedParts = { ...prevParts };
        Object.keys(parts).forEach(k => {
          if (parts[k] !== null && parts[k] !== undefined && !Number.isNaN(parts[k])) mergedParts[k] = parts[k];
        });
        entry.parts = mergedParts;
        // 제자리멀리뛰기·앉아윗몸앞으로굽히기·악력처럼 "여러 번 측정 중 최고기록"을 대표값으로
        // 쓰는 종목은, 방금 병합된 하위 값들을 기준으로 대표값도 다시 계산한다. 그렇지 않으면
        // 이 기기가 자신이 입력한 값만으로 최고기록을 계산해 저장하기 때문에, 다른 기기가 먼저
        // 저장해둔 더 좋은 하위 기록이 있어도(위에서 parts는 병합됐지만) 대표값(등급·순위·제출
        // 양식 반영에 쓰이는 값)은 그보다 낮게 저장되어버리는 경우가 생길 수 있다.
        if (eventId === "longjump" || eventId === "sitreach" || eventId === "gripstrength") {
          const mergedNums = Object.values(mergedParts).filter(v => typeof v === "number" && !Number.isNaN(v));
          if (mergedNums.length > 0) entry.value = Math.max(...mergedNums);
        }
      }
      const next = { ...base, [key]: entry };
      setRecords(next);
      setLastSync(Date.now());
      await persistRecords(next);
      const student = students.find(s => s.id === studentId);
      appendAuditLog(workspaceCode, {
        type: "record",
        ts: Date.now(),
        by: myDisplayName,
        studentName: student ? student.name : "(삭제된 학생)",
        eventName: EVENT_MAP[eventId]?.name || eventId,
        prevValue: prevValue ?? null,
        newValue: value,
        studentId,
        eventId,
        year,
        schoolGradeAtMeasure,
        prevParts: base[key]?.parts ?? null,
      });
    });
  }, [records, persistRecords, role, workspaceCode, students, myDisplayName]);

  // 변경 이력에서 "되돌리기"를 눌렀을 때: 그 기록이 원래 없던 값이었다면(prevValue가 null)
  // 그 기록 자체를 삭제하고, 값이 있었다면 그 이전 값으로 다시 저장한다(되돌리기 자체도
  // 새로운 변경 이력으로 남아, 누가 언제 되돌렸는지도 함께 추적된다). 이것도 같은 큐를 타서
  // 다른 저장과 순서가 뒤섞이지 않는다.
  const undoRecordChange = useCallback((entry) => {
    if (role !== "admin") return Promise.resolve();
    if (entry.prevValue === null || entry.prevValue === undefined) {
      return enqueueSave(async () => {
        const latest = (await loadRecordsRemote(workspaceCode)) || {};
        const base = mergeRecords(records, latest);
        const key = recKey(entry.studentId, entry.eventId, entry.year);
        const next = { ...base };
        delete next[key];
        setRecords(next);
        setLastSync(Date.now());
        await persistRecords(next);
        appendAuditLog(workspaceCode, {
          type: "record", ts: Date.now(), by: myDisplayName, studentName: entry.studentName,
          eventName: entry.eventName, prevValue: entry.newValue, newValue: null,
          studentId: entry.studentId, eventId: entry.eventId, year: entry.year,
          schoolGradeAtMeasure: entry.schoolGradeAtMeasure, prevParts: null,
        });
      });
    }
    return saveNewRecord(entry.studentId, entry.eventId, entry.prevValue, entry.year, entry.schoolGradeAtMeasure, entry.prevParts);
  }, [role, workspaceCode, records, persistRecords, myDisplayName, saveNewRecord]);

  const showToast = useCallback((text, kind = "info") => {
    setToast({ text, kind, id: uid("toast") });
    setTimeout(() => setToast(t => (t && t.text === text ? null : t)), 4000);
  }, []);

  // 무료 요금제 등 자동 저장이 안 되는 환경을 위한 수동 백업/복원.
  const handleImportBackup = useCallback(async (payload) => {
    if (role !== "admin" || !isFounder) return;
    const nextStudents = payload.students || [];
    const nextSettings = payload.settings || { schoolName: "", currentYear: thisYear() };
    const nextCriteria = payload.criteria || buildDefaultCriteria(nextSettings.schoolLevel);
    const nextRecords = payload.records || {};
    setStudents(nextStudents);
    setCriteria(nextCriteria);
    setSettings(nextSettings);
    setRecords(nextRecords);
    setActiveYear(nextSettings.currentYear || thisYear());
    await saveConfigRemote(workspaceCode, { students: nextStudents, criteria: nextCriteria, settings: nextSettings });
    await saveRecordsRemote(workspaceCode, nextRecords);
    setLastSync(Date.now());
    showToast("백업 파일을 불러왔습니다.", "ok");
  }, [role, isFounder, workspaceCode, showToast]);

  // 학기 마감: 나이스 제출 등 사용 목적을 다한 뒤, 학생 개인정보를 계속 저장해둘 필요가
  // 없도록 기록·학생 명단뿐 아니라 이 학교 코드 자체(설정·접근권한·이력 등 전부)를 완전히
  // 지운다. 마감 후에는 이 코드가 다시 "아무도 만든 적 없는 코드"가 되어, 화면은 첫
  // 화면(코드 설정)으로 돌아간다. 되돌릴 수 없다.
  const performSemesterCloseout = useCallback(async () => {
    if (role !== "admin") return;
    await storage.delete(recordsKey(workspaceCode), true).catch(() => {});
    await storage.delete(configKey(workspaceCode), true).catch(() => {});
    await storage.delete(accessKey(workspaceCode), true).catch(() => {});
    await storage.delete(auditLogKey(workspaceCode), true).catch(() => {});
    await storage.delete(backupLogKey(workspaceCode), true).catch(() => {});
    await storage.delete(deviceKey(workspaceCode), false).catch(() => {});
    // 이 기기가 "마지막으로 쓰던 코드"로 기억해 둔 값도 지운다. 그렇지 않으면 마감 직후
    // 화면은 첫 화면으로 돌아가더라도, 앱을 나갔다가 다시 열 때(특히 모바일) 이 기기가
    // 방금 지운 코드를 자동으로 다시 불러와 재접속을 시도하는 문제가 있었다.
    await clearSavedWorkspaceCode().catch(() => {});
    // 이 기기에 남아있던 학생 이름표(익명화를 위해 로컬에만 저장해뒀던 실명 매핑)도 함께
    // 지운다 — 서버 데이터가 사라진 뒤에도 이 브라우저에만 실명이 남아있지 않도록 한다.
    try { window.localStorage.removeItem(nameMapKey(workspaceCode)); } catch (e) {}
    showToast("마감이 완료되었습니다. 이 학교 코드의 모든 데이터가 삭제되었습니다.", "ok");
    setWorkspaceCode(null);
    setRole(null);
    setLoading(true);
  }, [role, workspaceCode, showToast]);

  // "계속 쓸게요" — 자동 마감 경고가 뜬 코드를 아직 쓰고 있다면, 개설 시각을 지금으로
  // 다시 잡아 1년을 새로 센다(수동 마감 없이 그대로 사용을 이어간다).
  const extendCodeLifetime = useCallback(async () => {
    if (role !== "admin") return;
    const cfg = await loadConfig(workspaceCode);
    const nextCfg = { ...(cfg || {}), students, criteria, settings, createdAt: Date.now() };
    await saveConfigRemote(workspaceCode, nextCfg);
    setCodeCreatedAt(nextCfg.createdAt);
    showToast("계속 사용합니다. 자동 마감 기한이 1년 뒤로 다시 설정되었습니다.", "ok");
  }, [role, workspaceCode, students, criteria, settings, showToast]);

  /* ---------- 계산 함수 (연도별) ---------- */
  const getBands = useCallback((eventId, gender, schoolGrade) => {
    if (!criteria) return null;
    return criteria?.[eventId]?.[gender]?.[schoolGrade] || null;
  }, [criteria]);

  // 학교급(초/중/고)에 따라 실제 존재하는 학년 범위가 다르므로, 저장된 학교급 설정에서 파생한다.
  const schoolGrades = gradesForLevel(settings.schoolLevel || "middle");

  const studentValue = useCallback((studentId, eventId, year) => {
    const r = records[recKey(studentId, eventId, year)];
    return r ? r.value : null;
  }, [records]);

  const studentParts = useCallback((studentId, eventId, year) => {
    const r = records[recKey(studentId, eventId, year)];
    return r && r.parts ? r.parts : null;
  }, [records]);

  const studentGrade = useCallback((student, eventId, year) => {
    const r = records[recKey(student.id, eventId, year)];
    if (!r) return null;
    const gradeAtMeasure = r.schoolGradeAtMeasure || student.grade;
    const bands = getBands(eventId, student.gender, gradeAtMeasure);
    return computeGradeFromBands(r.value, bands);
  }, [records, getBands]);

  /* ---------- 창 닫기/새로고침 전 경고 ---------- */
  // 무료 요금제 등에서 입력 중이던 내용을 깜빡하고 창을 닫아 잃어버리는 걸 막기 위해,
  // 관리자로 화면을 쓰는 동안에는 브라우저가 기본 제공하는 "정말 나가시겠습니까" 확인창을
  // 띄운다. 브라우저 보안 정책상 문구 자체는 각 브라우저가 정한 문구로 고정되며 직접
  // 바꿀 수는 없다.
  useEffect(() => {
    if (role !== "admin") return;
    function handleBeforeUnload(e) {
      e.preventDefault();
      e.returnValue = "";
      return "";
    }
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, [role]);

  /* ---------- 신기록 알림 ---------- */
  useEffect(() => {
    const genderList = GENDERS; // [{id:'M',label:'남'},{id:'F',label:'여'}]
    const gradeList = Array.from(new Set(students.map(s => s.grade))).sort((a, b) => a - b);
    const classList = Array.from(new Set(students.map(s => s.grade + "-" + s.classNum)));
    const messages = [];

    EVENTS.forEach(ev => {
      // 이 종목의 1위를 검사할 범위(전체/학년별/성별/반별) 목록을 구성한다.
      const scopes = [{ key: "overall", label: "전체", filter: () => true }];
      gradeList.forEach(g => scopes.push({ key: "grade-" + g, label: g + "학년", filter: s => s.grade === g }));
      genderList.forEach(g => scopes.push({ key: "gender-" + g.id, label: g.label, filter: s => s.gender === g.id }));
      classList.forEach(gc => {
        const [g, c] = gc.split("-").map(Number);
        scopes.push({ key: "class-" + gc, label: g + "학년 " + c + "반", filter: s => s.grade === g && s.classNum === c });
      });

      const changed = []; // { label, student, value }
      scopes.forEach(scope => {
        const list = students
          .filter(scope.filter)
          .map(s => ({ s, v: studentValue(s.id, ev.id, activeYear) }))
          .filter(x => x.v !== null);
        if (list.length === 0) return;
        list.sort((a, b) => (ev.better === "high" ? b.v - a.v : a.v - b.v));
        const top = list[0];
        const prevKey = ev.id + "|" + scope.key;
        const prev = prevTopRef.current[prevKey];
        if (prev && (prev.id !== top.s.id || prev.v !== top.v)) {
          changed.push({ label: scope.label, student: top.s, value: top.v });
        }
        prevTopRef.current[prevKey] = { id: top.s.id, v: top.v };
      });

      if (changed.length > 0) {
        // 같은 학생이 같은 기록으로 여러 범위에서 동시에 1위가 됐다면 알림 하나로 묶는다.
        const groups = new Map();
        changed.forEach(c => {
          const gk = c.student.id + "_" + c.value;
          if (!groups.has(gk)) groups.set(gk, { student: c.student, value: c.value, labels: [] });
          groups.get(gk).labels.push(c.label);
        });
        groups.forEach(g => {
          const labelText = g.labels.join(" · ") + " 1위";
          const maskedId = g.student.grade + "학년 " + g.student.classNum + "반 " + g.student.number + "번";
          messages.push(`🎉 ${ev.name} 신기록! ${maskedId} 학생 (${fmtValue(g.value, ev.unit)}) - ${labelText}`);
        });
      }
    });

    // 한 번에 여러 알림이 발생하면 서로 덮어쓰지 않도록 순서대로 하나씩 띄운다.
    messages.forEach((msg, i) => {
      setTimeout(() => showToast(msg, "record"), i * 4500);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [records, students, activeYear]);

  // 조회 전용 권한이면 입력/관리용 탭에 머물러 있지 않도록 안전하게 되돌린다.
  // (모든 조건부 return보다 앞에 위치해야 훅 호출 순서가 매 렌더마다 동일하게 유지된다.)
  useEffect(() => {
    if (role === "viewer" && ["records", "roster", "access", "backup"].includes(view)) {
      setView("board");
    }
  }, [role, view]);

  if (workspaceChecking) {
    return (
      <PapsStyles>
        <div className="paps-app paps-loading">
          <RefreshCw className="spin" size={28} />
          <div>확인 중...</div>
        </div>
      </PapsStyles>
    );
  }

  if (loadError) {
    return (
      <PapsStyles>
        <div className="paps-app paps-loading">
          <AlertTriangle size={28} color="#E85D5D" />
          <div>데이터를 불러오지 못했습니다. 새로고침 해주세요.</div>
        </div>
      </PapsStyles>
    );
  }

  if (!workspaceCode) {
    return (
      <PapsStyles>
        <WorkspaceGate onSubmit={submitWorkspaceCode} onRequestAccess={requestAccess} initialMode={gateInitialMode} />
      </PapsStyles>
    );
  }

  if (roleChecking || role === null) {
    return (
      <PapsStyles>
        <div className="paps-app paps-loading">
          <RefreshCw className="spin" size={28} />
          <div>권한 확인 중...</div>
        </div>
      </PapsStyles>
    );
  }

  if (role === "pending") {
    return (
      <PapsStyles>
        <PendingApprovalScreen name={myRequest?.name} type={myRequest?.type} onCancel={retryAfterBlock} />
      </PapsStyles>
    );
  }

  if (role === "blocked") {
    return (
      <PapsStyles>
        <BlockedScreen reason={blockReason} onRetry={retryAfterBlock} />
      </PapsStyles>
    );
  }

  if (loading) {
    return (
      <PapsStyles>
        <div className="paps-app paps-loading">
          <RefreshCw className="spin" size={28} />
          <div>데이터 불러오는 중...</div>
        </div>
      </PapsStyles>
    );
  }

  return (
    <PapsStyles>
      <div className={"paps-app" + (presentation ? " presentation" : "") + (settings.theme === "champion" ? " champion" : "")}>
        {!presentation && (
          <TopNav
            view={view}
            setView={setView}
            role={role}
            isFounder={isFounder}
            pendingCount={accessList.requests.filter(r => r.status === "pending").length}
            schoolName={settings.schoolName}
            lastSync={lastSync}
            activeYear={activeYear}
            setActiveYear={setActiveYear}
            workspaceCode={workspaceCode}
            onChangeWorkspace={changeWorkspaceCode}
            onEnterPresentation={() => {
              // 지금 창은 그대로 두고, 전광판(빔프로젝터) 전용 화면을 새 창으로 띄운다.
              // 새 창은 같은 브라우저(=같은 기기)에서 열리므로 접근 권한을 다시 확인할
              // 필요 없이 곧바로 전광판 모드로 들어간다. 두 창은 같은 Firestore 데이터를
              // 보므로 자연스럽게 동기화된다.
              const url = new URL(window.location.href);
              url.search = "";
              url.searchParams.set("present", "1");
              url.searchParams.set("code", workspaceCode);
              window.open(url.toString(), "_blank", "noopener");
            }}
            theme={settings.theme}
            onToggleTheme={role === "admin" ? () => {
              const next = { ...settings, theme: settings.theme === "champion" ? "default" : "champion" };
              setSettings(next);
              persistConfig(undefined, undefined, next);
            } : undefined}
          />
        )}

        {shareModalOpen && <ShareGuideModal workspaceCode={workspaceCode} onClose={() => setShareModalOpen(false)} />}
        {justCreatedNotice && (
          <CodeCreatedNoticeModal
            workspaceCode={workspaceCode}
            viewerPassword={settings.viewerPassword}
            founderPassword={settings.founderPassword}
            onClose={() => setJustCreatedNotice(false)}
          />
        )}
        {lockPromptOpen && (
          <BoardLockPrompt
            error={lockPromptError}
            onUnlock={handleUnlockAttempt}
            onCancel={() => { setLockPromptOpen(false); setLockPromptError(""); }}
          />
        )}
        {idleLocked && (
          <IdleLockScreen
            error={idleLockError}
            onUnlock={handleIdleUnlockAttempt}
          />
        )}

        {!bannerDismissed && !presentation && (
          <div className="info-banner">
            <Info size={16} />
            <span>
              {role === "viewer"
                ? "지도 목적 확인을 위한 조회 전용 화면입니다. 학생 개인정보인 체력 기록은 지도 목적 외로 저장·공유하지 마세요."
                : "이 기록판의 데이터는 이 화면 링크를 여는 모든 기기에서 함께 보이고 수정됩니다."}
            </span>
            <button className="icon-btn" onClick={() => setBannerDismissed(true)}><X size={14} /></button>
          </div>
        )}

        {isFounder && !presentation && codeCreatedAt && (() => {
          const ONE_YEAR_MS = 365 * 24 * 60 * 60 * 1000;
          const daysLeft = Math.ceil((codeCreatedAt + ONE_YEAR_MS - Date.now()) / (24 * 60 * 60 * 1000));
          if (daysLeft > 30) return null;
          return (
            <div className="info-banner warn-banner">
              <AlertTriangle size={16} />
              <span>
                이 코드가 개설된 지 곧 1년이 됩니다. <b>{Math.max(daysLeft, 0)}일 후 자동으로
                마감(기록·명단·설정 전체 삭제)</b>됩니다. 계속 쓰신다면 아래 버튼을 눌러
                기한을 늘려주세요.
              </span>
              <button className="btn btn-secondary small" onClick={extendCodeLifetime}>계속 사용(1년 연장)</button>
            </div>
          );
        })()}

        <div className="paps-body">
          {view === "board" && (
            <ScoreBoard
              students={students}
              records={records}
              activeYear={activeYear}
              studentValue={studentValue}
              studentGrade={studentGrade}
              skippedEvents={settings.skippedEvents || []}
              schoolGrades={schoolGrades}
              presentation={presentation}
              setPresentation={setPresentation}
              onExitPresentation={requestExitPresentation}
              lastSync={lastSync}
            />
          )}
          {view === "records" && role === "admin" && (
            <RecordManagementView
              students={students}
              records={records}
              activeYear={activeYear}
              onSave={saveNewRecord}
              studentValue={studentValue}
              studentParts={studentParts}
              settings={settings}
              setSettings={(next) => { setSettings(next); persistConfig(undefined, undefined, next); }}
              showToast={showToast}
            />
          )}
          {view === "grades" && (
            <GradeTable
              students={students}
              activeYear={activeYear}
              studentValue={studentValue}
              studentGrade={studentGrade}
              settings={settings}
              setSettings={(next) => { setSettings(next); persistConfig(undefined, undefined, next); }}
              isAdmin={role === "admin"}
            />
          )}
          {view === "roster" && role === "admin" && (
            <RosterManager
              students={students}
              setStudents={(next) => {
                // 학생이 추가되거나 삭제된 경우 감사 로그에도 남긴다(누가 언제 몇 명을 바꿨는지).
                const prevIds = new Set(students.map(s => s.id));
                const nextIds = new Set(next.map(s => s.id));
                const added = next.filter(s => !prevIds.has(s.id));
                const removed = students.filter(s => !nextIds.has(s.id));
                if (added.length > 0) {
                  appendAuditLog(workspaceCode, { type: "access", ts: Date.now(), by: myDisplayName, message: `${myDisplayName}님이 학생 ${added.map(s => s.name).join(", ")}을(를) 추가함` });
                }
                if (removed.length > 0) {
                  appendAuditLog(workspaceCode, { type: "access", ts: Date.now(), by: myDisplayName, message: `${myDisplayName}님이 학생 ${removed.map(s => s.name).join(", ")}을(를) 삭제함` });
                }
                setStudents(next);
                persistConfig(next, undefined, undefined);
              }}
              showToast={showToast}
              schoolGrades={schoolGrades}
            />
          )}
          {view === "backup" && role === "admin" && (
            <DataBackupPanel
              students={students}
              records={records}
              criteria={criteria}
              settings={settings}
              activeYear={activeYear}
              onImportBackup={handleImportBackup}
              showToast={showToast}
              workspaceCode={workspaceCode}
              myDisplayName={myDisplayName}
              isFounder={isFounder}
              neisFillState={neisFillState}
              setNeisFillState={setNeisFillState}
            />
          )}
          {view === "closeout" && role === "admin" && isFounder && (
            <SemesterCloseoutPanel
              students={students}
              records={records}
              criteria={criteria}
              settings={settings}
              onCloseout={performSemesterCloseout}
              workspaceCode={workspaceCode}
              activeYear={activeYear}
            />
          )}
          {view === "access" && role === "admin" && (
            <AccessRequestsPanel
              accessList={accessList}
              onApprove={approveRequest}
              onDeny={denyRequest}
              onRevoke={revokeApproval}
              onRestore={restoreRequest}
              onDeleteEntry={deleteRequestEntry}
              onTransferFounder={transferFounder}
              onUndoRecord={undoRecordChange}
              settings={settings}
              setSettings={(next) => { setSettings(next); persistConfig(undefined, undefined, next); }}
              onOpenShare={() => setShareModalOpen(true)}
              isFounder={isFounder}
              workspaceCode={workspaceCode}
              showToast={showToast}
            />
          )}
        </div>

        {toast && (
          <div className={"toast" + (toast.kind === "record" ? " toast-record" : "")}>
            {toast.text}
          </div>
        )}
      </div>
    </PapsStyles>
  );
}

/* ============================== 워크스페이스 코드 설정 ============================== */

function WorkspaceGate({ onSubmit, onRequestAccess, initialMode }) {
  const [value, setValue] = useState("");
  const [mode, setMode] = useState(initialMode || "code"); // 'code' | 'notice' | 'form'
  const [schoolLevel, setSchoolLevel] = useState("middle");
  const [initialPassword, setInitialPassword] = useState("");
  const [showInitialPassword, setShowInitialPassword] = useState(false);
  const [founderPassword, setFounderPassword] = useState("");
  const [createOpen, setCreateOpen] = useState(false);
  const [showFounderPassword, setShowFounderPassword] = useState(false);
  const [agree, setAgree] = useState(false);
  const [name, setName] = useState("");
  const [reqCode, setReqCode] = useState("");
  const [reqPassword, setReqPassword] = useState("");
  const [wantsEdit, setWantsEdit] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [formError, setFormError] = useState("");
  const [visitStats, setVisitStats] = useState(null);
  const [manualOpen, setManualOpen] = useState(false);

  // 방문자 통계는 누구나 볼 수 있게 첫 화면에 그냥 표시한다.
  useEffect(() => {
    loadVisitStats().then(setVisitStats);
  }, []);

  async function handleRequestSubmit() {
    if (!name.trim() || !reqCode.trim() || !reqPassword.trim()) return;
    setSubmitting(true);
    setFormError("");
    const result = await onRequestAccess({ name, code: reqCode, password: reqPassword, wantsEdit });
    setSubmitting(false);
    if (!result.ok) {
      if (result.reason === "bad-code") setFormError("존재하지 않는 학교 코드입니다. 관리자에게 다시 확인해 주세요.");
      else if (result.reason === "no-password-set") setFormError("이 학교는 아직 열람 비밀번호가 설정되어 있지 않습니다. 관리자에게 문의해 주세요.");
      else if (result.reason === "bad-password") setFormError("열람 비밀번호가 올바르지 않습니다.");
      else if (result.reason === "too-many-attempts") setFormError("비밀번호를 너무 여러 번 틀렸습니다. 10분 뒤 다시 시도해 주세요.");
      else setFormError("신청을 처리하지 못했습니다. 다시 시도해 주세요.");
    }
  }

  function submitCode() {
    if (!value.trim()) return;
    onSubmit(value, schoolLevel, initialPassword, founderPassword, createOpen ? "create" : "login");
  }

  if (mode === "code") {
    return (
      <div className="paps-app gate-screen">
        <div className="gate-card">
          <div className="gate-brand">
            <Trophy size={26} color="var(--gold)" />
            <span className="gate-brand-name">SMART PAPS</span>
            <Trophy size={26} color="var(--gold)" />
          </div>
          <p className="gate-tagline">
            일일이 출력, 수기 기록, 재입력 하던 업무가<br />
            <b>첨부, 모바일 기록, 마감</b>으로 끝.
          </p>
          <button className="manual-btn" onClick={() => setManualOpen(true)}>
            <ClipboardList size={13} /> 사용설명서
          </button>
          <div className="gate-divider" />

          {/* "새 코드 만들기"를 제목 바로 아래(첫 화면에서 가장 먼저 보이는 위치)로 옮겨
              처음 오는 선생님이 로그인 화면을 지나칠 필요 없이 바로 시작할 수 있게 한다. */}
          {!createOpen && (
            <>
              <button className="btn btn-secondary big-btn gate-create-emphasis" onClick={() => setCreateOpen(true)}>
                <Plus size={15} /> 처음이신가요? 새 코드 만들기
              </button>
              <div className="gate-divider" />
            </>
          )}

          {createOpen && (
            <>
              <h2>새 코드 만들기</h2>
              <p className="gate-desc">
                우리 학교만의 코드를 새로 만드세요. 학교 이름이 들어가지 않은 코드를 추천합니다.
              </p>
              <input
                className="input big-input gate-input"
                placeholder="새로 만들 코드 이름, 예: 낭만체육123"
                value={value}
                onChange={e => setValue(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter" && value.trim()) submitCode(); }}
              />
              <div className="text-dim small-note gate-code-warn">
                학교 이름이 그대로 들어간 코드는 피해주세요. 학년·반·번호와 학교 이름이 함께
                알려지면 학생이 누구인지 유추될 수 있습니다. "낭만체육123"처럼 학교와 무관한
                이름을 추천합니다.
              </div>
              <div className="gate-level-row">
                <label>학교급</label>
                <div className="chip-row gate-level-chips">
                  {SCHOOL_LEVELS.map(l => (
                    <button
                      key={l.id}
                      type="button"
                      className={"chip" + (schoolLevel === l.id ? " active" : "")}
                      onClick={() => setSchoolLevel(l.id)}
                    >
                      {l.label}
                    </button>
                  ))}
                </div>
              </div>
              <div className="gate-pw-row">
                <label>비밀번호 설정</label>
                <div className="pw-row">
                  <input
                    className="input"
                    type={showInitialPassword ? "text" : "password"}
                    value={initialPassword}
                    onChange={e => setInitialPassword(e.target.value)}
                    onKeyDown={e => { if (e.key === "Enter" && value.trim()) submitCode(); }}
                    placeholder="예: 체육0925"
                  />
                  <button type="button" className="btn btn-ghost small" onClick={() => setShowInitialPassword(v => !v)}>{showInitialPassword ? "숨기기" : "보기"}</button>
                </div>
                <div className="text-dim small-note gate-pw-hint">
                  동료 교사가 접근 신청 시 입력할 비밀번호입니다. 비워두면, 다른 선생님이 신청해도 아무도 들어올 수 없어요.
                  연도나 "1111" 같은 숫자만으로는 짐작되기 쉬우니, 영문+숫자를 섞어 6자 이상으로 정해주세요.
                </div>
              </div>
              <div className="gate-pw-row">
                <label>개설자 전용 비밀번호</label>
                <div className="pw-row">
                  <input
                    className="input"
                    type={showFounderPassword ? "text" : "password"}
                    value={founderPassword}
                    onChange={e => setFounderPassword(e.target.value)}
                    onKeyDown={e => { if (e.key === "Enter" && value.trim()) submitCode(); }}
                    placeholder="예: 원장선생님0925"
                  />
                  <button type="button" className="btn btn-ghost small" onClick={() => setShowFounderPassword(v => !v)}>{showFounderPassword ? "숨기기" : "보기"}</button>
                </div>
                <div className="text-dim small-note gate-pw-hint">
                  선생님(개설자) 본인만 알아야 하는 비밀번호입니다. 나중에 다른 기기(휴대폰↔컴퓨터 등)에서
                  같은 코드와 이 비밀번호를 "코드로 로그인" 화면에 입력하면, 승인 절차 없이 곧바로
                  개설자로 다시 들어올 수 있어요. 위 "비밀번호 설정"과는 다른 값으로 정해주세요
                  (동료 교사에게는 절대 알려주지 마세요).
                </div>
              </div>
              <button className="btn btn-primary big-btn" disabled={!value.trim()} onClick={submitCode}>
                만들기
              </button>
              <div className="gate-note">
                <Info size={14} />
                <span>
                  이미 등록되어 있는 학교코드라면, 개설자 전용 비밀번호가 맞을 때만 개설자로
                  들어가지고, 그 외에는 반영되지 않습니다.
                </span>
              </div>
              <button className="btn btn-ghost gate-back-toggle" onClick={() => setCreateOpen(false)}>
                ← 코드로 로그인 화면으로 돌아가기
              </button>
              <div className="gate-divider" />
            </>
          )}

          {!createOpen && (
            <>
              <h2>코드로 로그인</h2>
              <p className="gate-desc">
                이미 만들어 둔 학교 코드가 있으신가요? 코드와 개설자 전용 비밀번호를 입력하면
                바로 들어갈 수 있습니다.
              </p>
              <input
                className="input big-input gate-input"
                placeholder="예: 낭만체육123"
                value={value}
                onChange={e => setValue(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter" && value.trim()) submitCode(); }}
              />
              <div className="gate-pw-row">
                <label>개설자 전용 비밀번호</label>
                <div className="pw-row">
                  <input
                    className="input"
                    type={showFounderPassword ? "text" : "password"}
                    value={founderPassword}
                    onChange={e => setFounderPassword(e.target.value)}
                    onKeyDown={e => { if (e.key === "Enter" && value.trim()) submitCode(); }}
                    placeholder="예: 원장선생님0925"
                  />
                  <button type="button" className="btn btn-ghost small" onClick={() => setShowFounderPassword(v => !v)}>{showFounderPassword ? "숨기기" : "보기"}</button>
                </div>
              </div>
              <button className="btn btn-primary big-btn" disabled={!value.trim() || !founderPassword.trim()} onClick={submitCode}>
                로그인
              </button>

              <div className="gate-divider" />

              <button className="gate-link-btn" onClick={() => setMode("notice")}>
                이미 학교 코드가 있으신가요? <span className="gate-link-cta">접근 신청 →</span>
              </button>
              <div className="gate-input-hint">
                동료 교사가 신청하면, 조회는 바로 이용할 수 있고 수정 권한은 개설자가 승인해야 사용할 수 있어요.
              </div>
            </>
          )}

          {visitStats && (
            <div className="visit-stats-row">
              <span>오늘 방문 <b>{visitStats.daily}</b></span>
              <span>누적 방문 <b>{visitStats.total}</b></span>
            </div>
          )}
        </div>
        {manualOpen && <UserManualModal onClose={() => setManualOpen(false)} />}
      </div>
    );
  }

  if (mode === "notice") {
    return (
      <div className="paps-app gate-screen">
        <div className="gate-card">
          <AlertTriangle size={32} color="var(--gold)" />
          <h2>접근 신청 안내</h2>
          <ul className="gate-notice-list">
            <li><b>해당 학교 소속 선생님만</b> 그 학교의 학생 기록에 접근할 수 있습니다. 다른 학교 기록은 볼 수 없습니다.</li>
            <li>신청하려면 관리자(체육교사)에게 <b>학교 코드</b>와 <b>열람 비밀번호</b>를 미리 안내받아야 합니다. 둘 다 정확히 입력해야 합니다.</li>
            <li>이 신청은 <b>학생은 이용할 수 없습니다.</b> 학생 체력 기록 확인·입력이 필요한 교사만 신청해 주세요.</li>
            <li>학생의 체력 측정 기록은 민감한 개인정보입니다. 확인한 뒤에도 지도 목적 외 용도로 저장·촬영·공유하지 않아야 합니다.</li>
            <li>신청 시 입력한 이름은 관리자가 접속자를 파악하는 용도로만 쓰이며, 관리자는 언제든 개별적으로 접근을 취소할 수 있습니다.</li>
            <li>기본으로 부여되는 권한은 <b>조회(확인) 전용</b>입니다. 학생 기록을 입력·수정해야 한다면 다음 화면에서 <b>수정 권한도 함께 신청</b>할 수 있으며, 이 경우 관리자의 별도 승인이 필요합니다.</li>
          </ul>
          <label className="gate-agree">
            <input type="checkbox" checked={agree} onChange={e => setAgree(e.target.checked)} />
            위 내용을 확인했으며, 본인은 학생이 아닌 지도 목적 확인이 필요한 교사임을 확인합니다.
          </label>
          <div className="gate-btn-row">
            <button className="btn btn-ghost" onClick={() => setMode("code")}>돌아가기</button>
            <button className="btn btn-primary" disabled={!agree} onClick={() => setMode("form")}>다음</button>
          </div>
        </div>
      </div>
    );
  }

  if (mode === "form") {
    return (
      <div className="paps-app gate-screen">
        <div className="gate-card">
          <Users size={32} color="var(--gold)" />
          <h2>접근 신청</h2>
          <p className="gate-desc">관리자에게 안내받은 학교 코드와 열람 비밀번호를 입력해 주세요.</p>
          <input
            className="input big-input gate-input"
            placeholder="이름 (예: 2학년 3반 담임 김민준)"
            value={name}
            onChange={e => setName(e.target.value)}
          />
          <div className="text-dim small-note gate-name-hint">
            동명이인이 있으면 같은 신청으로 헷갈릴 수 있으니, "김민준"보다 "2학년 3반 김민준"처럼
            소속이나 담당 학급을 함께 적어 다른 선생님과 구분되게 해주세요.
          </div>
          <input
            className="input big-input gate-input"
            placeholder="학교 코드"
            value={reqCode}
            onChange={e => setReqCode(e.target.value)}
          />
          <input
            className="input big-input gate-input"
            type="password"
            placeholder="열람 비밀번호"
            value={reqPassword}
            onChange={e => setReqPassword(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter" && name.trim() && reqCode.trim() && reqPassword.trim()) handleRequestSubmit(); }}
          />
          <label className="gate-agree">
            <input type="checkbox" checked={wantsEdit} onChange={e => setWantsEdit(e.target.checked)} />
            학생 기록을 입력·수정할 권한도 필요합니다 (같은 학교 동료 체육교사만 해당 · 관리자 승인 필요)
          </label>
          {formError && <div className="gate-error">{formError}</div>}
          <div className="gate-btn-row">
            <button className="btn btn-ghost" onClick={() => setMode("notice")}>이전</button>
            <button
              className="btn btn-primary"
              disabled={!name.trim() || !reqCode.trim() || !reqPassword.trim() || submitting}
              onClick={handleRequestSubmit}
            >
              {submitting ? "확인 중..." : (wantsEdit ? "신청하기" : "확인하기")}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return null;
}

function UserManualModal({ onClose }) {
  const [codeNoteOpen, setCodeNoteOpen] = useState(false);
  const [mobileNoteOpen, setMobileNoteOpen] = useState(false);
  const [nameNoteOpen, setNameNoteOpen] = useState(false);
  const steps = [
    { title: "시작하기", body: "학교급(초/중/고)을 고르고 우리 학교만의 코드를 만드세요. 학교 이름이 들어가지 않은 코드를 추천해요(예: 낭만체육123). 이때 비밀번호도 함께 정해두면, 나중에 따로 설정할 필요가 없어요." },
    { title: "학생 등록", body: "\"학생관리\" 탭에서 명단을 등록하세요. 한 명씩 직접 입력하거나, 엑셀 파일을 끌어다 놓으면 한 번에 등록됩니다." },
    { title: "기록 측정·입력", body: "\"기록관리\" 탭에서 종목을 고르고, 학년·반을 선택해 기록을 입력하세요. 종목별로 음원 재생·타이머·자동 계산 같은 도구가 함께 제공됩니다. 체육관 등에서 화면을 여러 학생이 함께 보는 상황이라면, 화면 위쪽의 \"이름 가림\" 버튼을 눌러 이름을 \"홍*동\" 형태로 가리고 번호로 확인하며 입력할 수 있습니다." },
    { title: "등급 확인", body: "\"등급표\" 탭에서 학생별 종목별 등급을 참고용으로 확인할 수 있습니다." },
    { title: "전광판으로 공유 가능(선택)", body: "\"전광판\" 탭에서 실시간 순위를 보여주세요. 빔프로젝터 고정모드를 누르면 화면이 자동으로 잠겨, 학생이 함부로 조작할 수 없습니다. 개인정보보호법에 따라 전광판에는 학생 이름이 표시되지 않습니다." },
    { title: "나이스 제출", body: "\"데이터 백업\" 탭에서 나이스 엑셀양식 파일을 올리면, 우리 기록을 자동으로 채워줍니다. 학교 시스템 제출용 양식이므로 이 파일에는 학생 이름이 포함되어 만들어집니다. 다운로드하면 삭제 안내 팝업이 함께 뜨니, 나이스 등록을 마쳤다면 컴퓨터에서 바로 지워주세요." },
    { title: "학기 마감", body: "측정이 모두 끝나면 \"마감\" 탭에서 백업을 받은 뒤 기록을 정리하세요. 학생 개인정보를 필요 이상 보관하지 않기 위한 절차입니다. 필수인 JSON 백업 외에, 나중에 참고가 필요할 수도 있는 경우를 대비해 나이스 제출양식과 비슷한 형태의 엑셀로 전체 기록을 받아둘 수도 있습니다(선택). 이때 받는 백업 파일들은 나이스 등록이 끝난 뒤에는 컴퓨터에서 삭제해 주세요 — 앱 안의 기록은 마감으로 지워져도, 한 번 내려받아 다운로드 폴더에 남은 파일은 이 프로그램이 대신 지울 수 없습니다." },
  ];
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><ClipboardList size={18} color="var(--gold)" /> 사용설명서</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <div className="text-dim small-note">처음 쓰시는 분도 이 순서만 따라 하시면 됩니다.</div>
          <div className="text-dim small-note">
            <b>공용 PC 자동 잠금</b>: 접근 신청 비밀번호를 설정해 두면, 로그인 후 약 12분간
            마우스·키보드 조작이 없을 때 화면이 자동으로 잠기고 비밀번호를 다시 입력해야
            계속 쓸 수 있습니다. 체육관·교무실처럼 여러 사람이 함께 쓰는 컴퓨터에서 자리를
            비웠을 때 학생 정보가 그대로 노출되는 것을 막기 위한 기능입니다.
          </div>
          {steps.map((s, i) => (
            <div className="share-step" key={i}>
              <span className="share-step-num">{i + 1}</span>
              <div>
                <div className="share-step-title">{s.title}</div>
                <div className="share-step-body">{s.body}</div>
                {i === 0 && (
                  <>
                    <button type="button" className="manual-note-btn" onClick={() => setCodeNoteOpen(v => !v)}>
                      참고사항 {codeNoteOpen ? "▲" : "▼"}
                    </button>
                    {codeNoteOpen && (
                      <div className="manual-note-box">
                        학생 기록에는 학년·반·번호가 들어갑니다. 여기에 학교 이름까지 코드로
                        알려지면, 그 학교 사정을 아는 사람은 "몇 학년 몇 반 몇 번이 누구인지"를
                        비교적 쉽게 유추할 수 있습니다. 코드를 학교 이름과 무관하게 정하면, 이
                        코드만으로는 어느 학교인지 알 수 없어 이런 위험을 줄일 수 있습니다.
                      </div>
                    )}
                  </>
                )}
                {i === 1 && (
                  <>
                    <button type="button" className="manual-note-btn" onClick={() => setNameNoteOpen(v => !v)}>
                      참고사항 {nameNoteOpen ? "▲" : "▼"}
                    </button>
                    {nameNoteOpen && (
                      <div className="manual-note-box">
                        <b>지금 이 이름은 어디에 저장되나요?</b><br />
                        이 사이트는 학생 이름을 서버(Firestore)에 아예 보내지 않습니다. 서버에는
                        학년·반·번호·성별처럼 학생을 구분하는 정보만 저장되고, 실제 이름은 지금
                        입력하고 있는 <b>이 기기(브라우저)에만</b> 남습니다.<br /><br />
                        그래서 같은 학교 코드로 다른 기기(동료 선생님 컴퓨터, 새로 바꾼 휴대폰 등)에
                        처음 접속하면, 그 기기엔 아직 이름표가 없어서 이름 대신 "(이름 미확인 - 이
                        기기)"처럼 보일 수 있어요. 그 상태에서 실명이 포함된 백업 파일을 불러오면
                        그 기기에도 이름이 채워집니다.<br /><br />
                        <b>전광판 화면</b>은 어느 기기에서 보든 이름 대신 [학년-반-번호] 형태로만
                        표시됩니다(개인정보 보호를 위해 항상 가림).<br /><br />
                        <b>이름이 그대로 들어가는 곳</b>: 데이터 백업(JSON) 파일, 나이스 제출용
                        엑셀 파일 — 이 둘은 이 기기에 저장된 이름표를 이용해 실명을 채워 넣으며,
                        실명이 필요한 목적이라 의도적으로 포함시킵니다. JSON 백업은 여러 명이
                        각자 백업하면 혼선이 생길 수 있어 개설자만 내보내고 불러올 수 있습니다.
                      </div>
                    )}
                  </>
                )}
                {i === 2 && (
                  <>
                    <button type="button" className="manual-note-btn" onClick={() => setMobileNoteOpen(v => !v)}>
                      참고사항 {mobileNoteOpen ? "▲" : "▼"}
                    </button>
                    {mobileNoteOpen && (
                      <div className="manual-note-box">
                        <b>스마트폰으로 기록하기</b>: 이 화면의 링크를 스마트폰 브라우저(사파리·크롬 등)로
                        열고 학교 코드를 입력하면 노트북과 똑같이 기록을 입력할 수 있습니다.<br /><br />
                        <b>홈 화면에 아이콘처럼 추가하기</b>: 브라우저의 공유 버튼 → "홈 화면에 추가"를
                        누르면, 매번 링크를 찾지 않아도 앱처럼 아이콘을 눌러 바로 열립니다.<br /><br />
                        다만 이건 <b>바로가기 아이콘</b>이라, 잠금화면이나 홈 화면에 실시간 순위 같은
                        정보가 그대로 표시되는 &quot;위젯&quot;까지는 만들어지지 않습니다. 열면 화면이 뜨는
                        정도로 이해해 주세요.
                      </div>
                    )}
                  </>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function PendingApprovalScreen({ name, type, onCancel }) {
  const isEditor = type === "editor";
  return (
    <div className="paps-app gate-screen">
      <div className="gate-card">
        <RefreshCw className="spin" size={28} color="var(--gold)" />
        <h2>관리자 승인 대기 중</h2>
        <p className="gate-desc">
          <b>{name}</b>님의 {isEditor ? "수정 권한 요청" : "열람 신청"}이 접수되었습니다. 최초 개설자(관리자)가
          승인하면 자동으로 화면이 전환됩니다.
          {isEditor && " 승인 전까지는 기록을 입력·수정할 수 없습니다."}
        </p>
        <button className="btn btn-ghost" onClick={onCancel}>신청 취소</button>
      </div>
    </div>
  );
}

function BlockedScreen({ reason, onRetry }) {
  if (reason === "no-such-code") {
    return (
      <div className="paps-app gate-screen">
        <div className="gate-card">
          <Info size={32} color="var(--gold)" />
          <h2>존재하지 않는 코드입니다</h2>
          <p className="gate-desc">
            입력하신 코드를 찾을 수 없습니다. 아직 만들어진 적이 없거나, 이미 "마감"으로
            삭제된 코드일 수 있습니다. 코드를 다시 확인해 주시거나, 새로 만들어 주세요.
          </p>
          <button className="btn btn-primary" onClick={() => onRetry("code")}>처음으로</button>
        </div>
      </div>
    );
  }
  if (reason === "need-request") {
    return (
      <div className="paps-app gate-screen">
        <div className="gate-card">
          <Info size={32} color="var(--gold)" />
          <h2>접근 신청이 필요합니다</h2>
          <p className="gate-desc">
            이미 다른 선생님이 만들어 둔 코드입니다. 코드를 입력하는 것만으로는 더 이상 권한이
            주어지지 않아요. 다음 화면에서 "접근 신청"으로 조회 권한 또는 수정 권한을 신청해 주세요.
          </p>
          <button className="btn btn-primary" onClick={() => onRetry("form")}>접근 신청 하러 가기</button>
        </div>
      </div>
    );
  }
  const message = reason === "denied"
    ? "관리자가 이 신청을 승인하지 않았습니다."
    : "이 권한이 취소되었습니다. 필요하다면 관리자에게 다시 문의해 주세요.";
  return (
    <div className="paps-app gate-screen">
      <div className="gate-card">
        <XCircle size={32} color="#E85D5D" />
        <h2>접근할 수 없습니다</h2>
        <p className="gate-desc">{message}</p>
        <button className="btn btn-primary" onClick={() => onRetry("code")}>처음으로</button>
      </div>
    </div>
  );
}

/* ============================== 접근 요청 관리(관리자) ============================== */

function AccessRequestsPanel({ accessList, onApprove, onDeny, onRevoke, onRestore, onDeleteEntry, onTransferFounder, onUndoRecord, settings, setSettings, onOpenShare, isFounder, workspaceCode, showToast }) {
  const [showPw, setShowPw] = useState(false);
  const [showFounderPw, setShowFounderPw] = useState(false);
  const [auditLog, setAuditLog] = useState(null);
  const [auditLoading, setAuditLoading] = useState(false);
  const [transferTarget, setTransferTarget] = useState(null);
  const requests = accessList.requests || [];
  const pending = requests.filter(r => r.status === "pending").sort((a, b) => a.submittedAt - b.submittedAt);
  const approved = requests.filter(r => r.status === "approved").sort((a, b) => b.submittedAt - a.submittedAt);
  const denied = requests.filter(r => r.status === "denied").sort((a, b) => b.submittedAt - a.submittedAt);

  async function loadAuditLogClick() {
    setAuditLoading(true);
    const log = await loadAuditLog(workspaceCode);
    setAuditLog(log.slice().reverse().slice(0, 30));
    setAuditLoading(false);
  }

  async function confirmTransfer() {
    if (!transferTarget) return;
    const result = await onTransferFounder(transferTarget.id);
    setTransferTarget(null);
    if (result.ok) showToast(result.newFounderName + "님에게 개설자 권한을 넘겼습니다. 이제 이 계정은 일반 편집자입니다.", "ok");
  }

  return (
    <div className="entry-layout">
      <div className="panel">
        <h3>모바일로 함께 보기</h3>
        <div className="text-dim small-note">
          노트북 없이 본인 스마트폰으로 실시간 기록을 확인·입력하거나, 동료 교사도 함께 쓰게 하려면 아래에서 안내를 확인하세요.
        </div>
        <button className="btn btn-ghost" onClick={onOpenShare}>
          <Smartphone size={14} /> 모바일 공유 안내 보기
        </button>

        <div className="divider" />

        <h3>선생님 접근 신청 설정</h3>
        {!isFounder && (
          <div className="warn-note">
            <ShieldCheck size={16} />
            <span>비밀번호 변경, 승인·거절·권한 취소, 개설자 승계는 <b>최초 개설자만</b> 할 수 있습니다. 필요하면 개설자 선생님께 요청해 주세요.</span>
          </div>
        )}
        <div className="text-dim small-note">
          코드를 처음 만들 때 이미 정하셨다면 여기서 확인·변경할 수 있습니다. 아직 안 정하셨거나
          바꾸고 싶다면 아래에서 설정하세요. 조회는 비밀번호만 맞으면 바로 승인되고, 수정 권한은
          선생님이 직접 승인해야 합니다. 비워두면 접근 신청 자체가 꺼집니다.
        </div>
        <div className="form-row">
          <label>접근 신청 비밀번호</label>
          <div className="pw-row">
            <input
              className="input"
              type={showPw ? "text" : "password"}
              value={settings.viewerPassword || ""}
              onChange={e => setSettings({ ...settings, viewerPassword: e.target.value })}
              placeholder="예: 체육0925"
              disabled={!isFounder}
            />
            <button className="btn btn-ghost small" onClick={() => setShowPw(v => !v)}>{showPw ? "숨기기" : "보기"}</button>
          </div>
        </div>
        <div className="text-dim small-note">
          비밀번호를 바꾸면 예전 비밀번호로는 신청할 수 없습니다. 이미 승인된 선생님을 막으려면
          아래 목록에서 따로 취소해야 해요. 이 비밀번호는 아래 "빔프로젝터 화면 잠금"에도 쓰입니다.
          연도나 "1111" 같은 숫자만으로는 짐작되기 쉬우니, 영문+숫자를 섞어 6자 이상을 권장합니다.
        </div>

        {isFounder && (
          <>
            <div className="divider" />
            <h3>개설자 전용 비밀번호</h3>
            <div className="text-dim small-note">
              다른 기기(휴대폰↔컴퓨터 등)에서 이 코드와 이 비밀번호를 함께 입력하면, 승인 절차
              없이 곧바로 개설자로 다시 들어올 수 있습니다. <b>동료 교사에게는 절대 알려주지
              마세요</b> — 이걸 아는 사람은 누구나 개설자 전권을 갖게 됩니다.
            </div>
            <div className="form-row">
              <label>개설자 전용 비밀번호</label>
              <div className="pw-row">
                <input
                  className="input"
                  type={showFounderPw ? "text" : "password"}
                  value={settings.founderPassword || ""}
                  onChange={e => setSettings({ ...settings, founderPassword: e.target.value })}
                  placeholder="예: 원장선생님0925"
                />
                <button className="btn btn-ghost small" onClick={() => setShowFounderPw(v => !v)}>{showFounderPw ? "숨기기" : "보기"}</button>
              </div>
            </div>
          </>
        )}

        <div className="divider" />

        <h3>빔프로젝터 화면 잠금</h3>
        <div className="text-dim small-note">
          "빔프로젝터 고정모드"로 전광판을 띄우는 순간부터 자동으로 잠깁니다. 나가려면 위
          비밀번호를 입력해야 해서, 자리를 비운 사이 학생 등이 함부로 다른 화면을 열 수 없습니다.
          <b> 이 잠금은 지금 이 화면(계정)에만 적용되고, 동료 선생님 화면에는 영향을 주지 않습니다.</b>
          {!settings.viewerPassword && " 비밀번호를 설정하지 않으면 잠기지 않고 바로 나가집니다."}
        </div>
      </div>

      <div className="panel">
        <h3>승인 대기 ({pending.length}명)</h3>
        <div className="text-dim small-note">"수정 권한"은 접근 신청 시 "수정 권한도 필요"를 체크한 경우이고, "열람 전용"은 체크하지 않은 경우입니다.</div>
        {pending.length === 0 && <div className="text-dim">대기 중인 신청이 없습니다.</div>}
        <div className="access-list">
          {pending.map(r => (
            <div className="access-row" key={r.id}>
              <div>
                <div className="access-name">{r.name} <span className={"access-type-badge" + (r.type === "editor" ? " editor" : "")}>{r.type === "editor" ? "수정 권한" : "열람 전용"}</span></div>
                <div className="access-time text-dim">{new Date(r.submittedAt).toLocaleString("ko-KR")}</div>
              </div>
              {isFounder ? (
                <div className="access-actions">
                  <button className="btn btn-primary small" onClick={() => onApprove(r.id)}>승인</button>
                  <button className="btn btn-ghost small danger-btn" onClick={() => onDeny(r.id)}>거절</button>
                </div>
              ) : (
                <span className="text-dim small-note">개설자 승인 대기 중</span>
              )}
            </div>
          ))}
        </div>

        <div className="divider" />

        <h3>승인된 계정 ({approved.length}명)</h3>
        {approved.length === 0 && <div className="text-dim">아직 승인된 계정이 없습니다.</div>}
        <div className="access-list">
          {approved.map(r => (
            <div className="access-row" key={r.id}>
              <div>
                <div className="access-name">{r.name} <span className={"access-type-badge" + (r.type === "editor" ? " editor" : "")}>{r.type === "editor" ? "수정 권한" : "열람 전용"}</span></div>
                <div className="access-time text-dim">승인됨</div>
              </div>
              {isFounder && (
                <div className="access-actions">
                  {r.type === "editor" && (
                    <button className="btn btn-ghost small" onClick={() => setTransferTarget(r)}>개설자로 지정</button>
                  )}
                  <button className="btn btn-ghost small danger-btn" onClick={() => onRevoke(r.id)}>권한 취소</button>
                </div>
              )}
            </div>
          ))}
        </div>
        {denied.length > 0 && (
          <>
            <div className="divider" />
            <h3>거절/취소된 신청</h3>
            <div className="access-list">
              {denied.map(r => (
                <div className="access-row" key={r.id}>
                  <div>
                    <div className="access-name text-dim">{r.name}</div>
                  </div>
                  {isFounder ? (
                    <div className="access-actions">
                      <button className="btn btn-ghost small" onClick={() => onRestore(r.id)}>복구</button>
                      <button className="btn btn-ghost small danger-btn" onClick={() => onDeleteEntry(r.id)}>삭제</button>
                    </div>
                  ) : (
                    <span className="text-dim small-note">거절됨</span>
                  )}
                </div>
              ))}
            </div>
          </>
        )}
      </div>

      {isFounder && (
        <div className="panel">
          <h3>최근 변경 이력</h3>
          <div className="text-dim small-note">
            누가 언제 어떤 학생의 기록을 바꿨는지 최근 항목을 확인할 수 있습니다(최대 300건 보관).
            성적 관련 문의가 들어왔을 때 참고하세요.
          </div>
          <button className="btn btn-ghost" onClick={loadAuditLogClick} disabled={auditLoading}>
            <RefreshCw size={14} /> {auditLoading ? "불러오는 중..." : "최근 변경 이력 보기"}
          </button>
          {auditLog && (
            auditLog.length === 0 ? (
              <div className="text-dim small-note">아직 기록된 변경 이력이 없습니다.</div>
            ) : (
              <div className="audit-log-list">
                {auditLog.map((a, i) => (
                  <div className="audit-log-row" key={i}>
                    <span className="audit-log-time">{new Date(a.ts).toLocaleString("ko-KR")}</span>
                    {a.type === "access" ? (
                      <span>{a.message}</span>
                    ) : (
                      <div className="audit-log-body">
                        <span>
                          <b>{a.by}</b>님이 <b>{a.studentName}</b>의 {a.eventName} 기록을{" "}
                          {a.prevValue === null || a.prevValue === undefined ? "새로" : a.prevValue + " →"}{" "}
                          {a.newValue === null ? "삭제" : a.newValue + "로 변경"}
                        </span>
                        {a.newValue !== null && (
                          <button className="btn btn-ghost small" onClick={() => onUndoRecord(a)}>
                            <RotateCcw size={12} /> 되돌리기
                          </button>
                        )}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )
          )}
        </div>
      )}

      {transferTarget && (
        <ConfirmModal
          title="개설자 권한 넘기기"
          message={`"${transferTarget.name}" 선생님에게 개설자 권한을 넘기시겠습니까? 넘기고 나면 이 계정(지금 화면)은 일반 편집자가 되어, 승인·거절·비밀번호 변경·마감 같은 작업을 더 이상 할 수 없습니다.`}
          confirmLabel="넘기기"
          danger
          onConfirm={confirmTransfer}
          onCancel={() => setTransferTarget(null)}
        />
      )}
    </div>
  );
}

function TopNav({ view, setView, role, isFounder, pendingCount, schoolName, lastSync, activeYear, setActiveYear, workspaceCode, onChangeWorkspace, onEnterPresentation, theme, onToggleTheme }) {
  const isAdmin = role === "admin";
  const [updatesOpen, setUpdatesOpen] = useState(false);
  const [deviceGuideOpen, setDeviceGuideOpen] = useState(false);
  const allTabs = [
    { id: "board", label: "전광판", icon: Monitor },
    { id: "roster", label: "학생관리", icon: Users, adminOnly: true },
    { id: "records", label: "기록관리", icon: Plus, adminOnly: true },
    { id: "grades", label: "등급표", icon: ClipboardList },
    { id: "backup", label: "데이터 백업", icon: Database, adminOnly: true },
    { id: "access", label: "접근권한", icon: ShieldCheck, adminOnly: true, badge: pendingCount },
    { id: "closeout", label: "마감", icon: Trash2, founderOnly: true, danger: true },
  ];
  const tabs = allTabs.filter(t => (!t.adminOnly || isAdmin) && (!t.founderOnly || (isAdmin && isFounder)));
  const secAgo = lastSync ? Math.max(0, Math.round((Date.now() - lastSync) / 1000)) : null;

  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);
  const pad = n => String(n).padStart(2, "0");
  const nowText = `${pad(now.getMonth() + 1)}월 ${pad(now.getDate())}일 ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

  return (
    <div className="topnav">
      <div className="brand">
        <span className="sync-indicator">
          <RefreshCw size={12} className="spin-slow" />
          {secAgo !== null ? secAgo + "초 전 동기화" : "동기화 대기"}
        </span>
        <Trophy size={20} color="var(--gold)" />
        <span className="brand-text">{schoolName ? schoolName + " " : ""}SMART PAPS</span>
        <button className="feature-updates-btn" onClick={() => setUpdatesOpen(true)} title="이 프로그램이 할 수 있는 일 모아보기">
          <Award size={12} /> 기능설명
        </button>
        <button className="feature-updates-btn" onClick={() => setDeviceGuideOpen(true)} title="휴대폰·노트북 함께 쓰는 법 자세히 보기">
          <Info size={12} /> 상세설명
        </button>
        {isAdmin ? (
          <button className="workspace-badge" onClick={onChangeWorkspace} title="워크스페이스 코드 변경">
            코드: {workspaceCode}
          </button>
        ) : (
          <span className="workspace-badge viewer-badge" title="조회 전용 계정">
            조회 전용
          </span>
        )}
      </div>
      <div className="tabs">
        {tabs.map(t => {
          const Icon = t.icon;
          return (
            <button
              key={t.id}
              className={"tab" + (view === t.id ? " active" : "") + (t.danger ? " tab-danger" : "")}
              onClick={() => setView(t.id)}
            >
              <Icon size={16} />
              <span>{t.label}</span>
              {!!t.badge && <span className="tab-badge">{t.badge}</span>}
            </button>
          );
        })}
      </div>
      <div className="nav-right">
        <div className="datetime-row">
          <span className="year-badge" title="측정 연도(자동으로 올해로 설정됩니다)">{activeYear}년</span>
          <span className="now-text">{nowText}</span>
        </div>
        {onToggleTheme && (
          <button className={"btn btn-ghost theme-toggle-btn" + (theme === "champion" ? " on" : "")} onClick={onToggleTheme} title="화면 디자인 모드 전환">
            <Award size={14} /> {theme === "champion" ? "챔피언십 모드" : "기본 모드"}
          </button>
        )}
        <button className="btn btn-ghost" onClick={onEnterPresentation} title="빔프로젝터로 학생들에게 크게 보여주는 동안, 노트북에서는 계속 다른 작업을 이어갈 수 있습니다">
          <Maximize2 size={14} /> 빔프로젝터 고정모드
        </button>
      </div>
      {updatesOpen && <FeatureUpdatesModal isAdmin={isAdmin} onClose={() => setUpdatesOpen(false)} />}
      {deviceGuideOpen && <DeviceSyncGuideModal onClose={() => setDeviceGuideOpen(false)} />}
    </div>
  );
}

// 참고: 원래 아티팩트 버전엔 "개설자 전용 비밀번호로 다른 기기에서 즉시 재접속"과
// "동료 교사가 같은 이름으로 재신청하면 기존 승인을 자동으로 이어받는" 기능까지 함께
// 안내하는 내용이 있었지만, 이 웹사이트 버전에는 그 두 기능 자체가 아직 반영되어 있지
// 않아 여기서는 뺐다(두 기능을 나중에 이 웹사이트에도 반영하면 그때 안내를 다시 채우면 됨).
function DeviceSyncGuideModal({ onClose }) {
  const sections = [
    {
      title: "기본 원리",
      body: "휴대폰·노트북 어디서 접속하든 \"학교 코드\"만 같으면 같은 데이터를 봅니다. 한 기기에서 기록을 입력하면, 몇 초 안에 다른 기기 화면에도 자동으로 반영돼요(따로 저장·새로고침 누를 필요 없음).",
    },
    {
      title: "여러 기기를 어떻게 나눠 쓰면 좋은가",
      body: "예: 노트북은 교무실 책상에 두고 등급표·백업 등 정리 작업을, 휴대폰은 운동장에 들고 나가 실측 기록 입력을 담당하는 식으로 나눠 쓰면 편합니다. 두 기기 모두 같은 코드로 로그인하면 됩니다.",
    },
    {
      title: "개설자(관리자) 권한을 여러 기기에서 쓰려면",
      body: "코드를 처음 만들 때 정한 \"개설자 전용 비밀번호\"를 기억해두세요. 다른 기기의 \"코드로 로그인\" 화면에 코드와 이 비밀번호를 입력하면, 별도 승인 없이 바로 개설자 권한으로 들어갈 수 있습니다.",
    },
    {
      title: "동료 교사가 다른 기기에서 다시 들어와야 할 때",
      body: "이미 승인받은 것과 똑같은 이름 + 똑같은 권한 종류(수정 권한/조회)로 접근 신청을 다시 하면, 처음부터 다시 승인을 기다리지 않고 기존 승인을 그대로 이어받습니다. 이름을 정확히 똑같이 입력하는 게 중요해요.",
    },
    {
      title: "1년 지난 코드는 자동으로 마감됩니다",
      body: "마감(전체 데이터 삭제)을 깜빡 잊고 넘어가는 경우를 대비해, 코드를 개설한 지 1년이 지나면 자동으로 마감 처리되어 기록·명단·설정이 모두 삭제되고 첫 화면으로 돌아갑니다. 만료 30일 전부터 개설자에게 경고 배너가 뜨고, 계속 쓰실 거라면 \"계속 사용(1년 연장)\" 버튼으로 기한을 늘릴 수 있습니다.",
    },
  ];
  const cautions = [
    "브라우저의 \"사이트 데이터 지우기\"나 시크릿(비공개) 모드로 접속하면, 이 기기가 승인받았다는 정보가 사라져 다시 접근 절차를 밟아야 할 수 있습니다.",
    "같은 이름을 쓰는 동료 교사가 두 명 이상이면, 위 \"기존 승인 이어받기\" 기능 때문에 서로 같은 자리를 나눠 쓰게 될 수 있어요. 이름에 학년·반처럼 구분되는 정보를 꼭 포함해 주세요.",
    "개설자 전용 비밀번호는 동료 교사에게 알려주지 마세요 — 이걸 아는 사람은 승인 절차 없이 곧바로 전체 권한을 갖게 됩니다.",
    "JSON 백업(내보내기·불러오기)은 개설자 기기에서만 할 수 있습니다. 여러 기기에서 각자 백업·복원하면 서로 다른 시점의 기록이 뒤섞일 수 있어, 백업은 개설자 한 명이 맡는 것을 권장합니다.",
    "인터넷 연결이 끊긴 상태에서 입력한 기록은 연결이 복구되어야 다른 기기에 반영됩니다.",
    "1년 자동 마감은 되돌릴 수 없습니다. 계속 쓰실 코드라면 경고 배너가 뜰 때 꼭 \"계속 사용(1년 연장)\"을 눌러주세요.",
  ];
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><Info size={18} color="var(--gold)" /> 상세설명 — 여러 기기 함께 쓰기</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">
          {sections.map((s, i) => (
            <div className="feature-update-group" key={i}>
              <h4>{s.title}</h4>
              <div className="text-dim small-note">{s.body}</div>
            </div>
          ))}
          <div className="feature-update-group">
            <h4>주의·유의사항</h4>
            <ul className="device-guide-caution-list">
              {cautions.map((c, i) => <li key={i} className="text-dim small-note">{c}</li>)}
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}

function FeatureUpdatesModal({ isAdmin, onClose }) {
  const groups = [
    {
      title: "측정 도구",
      items: [
        "종목별 실시간 지수·합계 자동계산 — 심박수·부위 점수 등을 입력하면 지수·합계가 즉시 계산됨",
        "종목별 측정 편의성 향상 — 반 전체 동시 측정 타이머, 운동장 코스 계산기 등 현장에서 바로 쓸 수 있는 세팅 제공",
        "일부 종목 공식음원 탑재 — 왕복오래달리기(음원 또는 영상)·윗몸말아올리기·스텝검사를 화면에서 바로 재생",
      ],
    },
    {
      title: "함께 쓰기",
      items: [
        "조회는 즉시, 수정은 개설자 승인 후 — 권한별 안전한 분리",
        "승인·비밀번호·마감은 개설자만 — 오조작 위험 최소화",
        "변경 이력·되돌리기로 잘못된 입력도 바로 복구",
      ],
    },
    {
      title: "데이터 보안",
      items: [
        "학교 코드를 모르면 애초에 접근 자체가 불가능",
        "비밀번호 5회 오답 시 10분간 자동 잠금 — 무작위 대입 시도 차단",
        "누가 언제 무엇을 바꿨는지 전부 기록 — 문제 발생 시 추적 가능",
        "승인·비밀번호 변경·마감 같은 민감한 조작은 개설자 1인만 — 통제된 접근 구조",
        "빔프로젝터 고정모드를 켜면 그 순간부터 자동으로 화면 잠김 — 교사가 자리를 비운 사이 학생의 임의 조작·확인 방지",
        "공용 PC 자동 잠금(세션 타임아웃) — 로그인 후 약 12분간 조작이 없으면 화면이 자동으로 잠기고, 계속 쓰려면 비밀번호를 다시 입력해야 함",
        "실시간 측정 중 이름 가림 모드 — 기록관리 화면에서 버튼 하나로 학생 이름을 \"홍*동\" 형태로 가리고 번호로 확인하며 입력 가능",
        "다운로드한 파일 삭제 안내 — 나이스 반영·백업(엑셀/JSON) 파일을 내려받을 때마다, 등록을 마쳤다면 컴퓨터에서 삭제해 달라는 안내가 뜸",
      ],
    },
    {
      title: "데이터 보관 기간",
      items: [
        "측정 기록은 \"마감\"을 누르기 전까지 계속 보관됨 — 학기 중에는 자동으로 사라지지 않음",
        "마감 시 기록과 학생 명단이 함께 즉시 삭제됨",
        "변경 이력은 최근 300건까지만 보관 — 그 이상은 오래된 순으로 자동 정리",
        "백업 파일은 내려받는 선생님의 개인 컴퓨터에만 저장 — 프로그램이 별도로 영구 보관하지 않음",
      ],
    },
    {
      title: "데이터 관리",
      items: [
        "백업 파일로 데이터 손실 위험 최소화 — 여러 명이 각자 백업·복원하면 최신 기록이 뒤섞일 수 있어 개설자만 가능",
        "나이스 측정명단 양식 엑셀 파일 첨부로 명단 반영 — 학생관리에서 파일만 올리면 학년·반·번호·이름을 자동으로 채워줌",
        "나이스 '자료올리기'용 엑셀 형식 지원 — 프로그램 내 기록을 토대로 나이스 업로드 양식에 맞춰 채워줌",
        "마감 시 백업 필수화로 학생 개인정보 최소 보관",
        "1년 지난 코드는 자동 마감 — 마감을 깜빡 잊어도 개설 1년 후 자동으로 전체 삭제되어 기록이 쌓이지 않음(만료 30일 전부터 경고, 연장 가능)",
      ],
    },
    {
      title: "학교 상황에 맞추기",
      items: [
        "학교급 선택만으로 학년·등급 기준 자동 적용",
        "빔프로젝터 자동 잠금 — 학생 정보 노출 방지",
        "모바일 지원 — 현장 어디서든 입력 가능",
      ],
    },
  ];
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><Award size={18} color="var(--gold)" /> 기능설명</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">
          {groups.map(g => (
            <div className="feature-update-group" key={g.title}>
              <h4>{g.title}</h4>
              <ul>
                {g.items.map((it, i) => {
                  const parts = it.split(" — ");
                  return (
                    <li key={i}>
                      <span className="feature-update-headline">{parts[0]}</span>
                      {parts[1] && <span className="feature-update-desc">{parts[1]}</span>}
                    </li>
                  );
                })}
              </ul>
            </div>
          ))}
          {!isAdmin && (
            <div className="text-dim small-note">조회 전용 계정에서는 이 중 일부(기록 입력, 학생 관리 등)는 사용할 수 없습니다.</div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================== 모바일 공유 안내 ============================== */

function BoardLockPrompt({ onUnlock, onCancel, error }) {
  const [pw, setPw] = useState("");
  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><ShieldCheck size={18} /> 빔프로젝터 화면 나가기</h3>
          <button className="icon-btn" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <div className="text-dim small-note">전광판 화면을 나가려면 접근 신청 비밀번호를 입력해 주세요.</div>
          <input
            className="input"
            type="password"
            value={pw}
            onChange={e => setPw(e.target.value)}
            placeholder="접근 신청 비밀번호"
            onKeyDown={e => { if (e.key === "Enter") onUnlock(pw); }}
            autoFocus
          />
          {error && <div className="gate-error">{error}</div>}
          <div className="confirm-actions">
            <button className="btn btn-ghost" onClick={onCancel}>취소</button>
            <button className="btn btn-primary" onClick={() => onUnlock(pw)}>확인</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// 공용 PC(체육관·교무실 공용 컴퓨터 등)에 마감 처리 전 상태로 화면이 켜진 채 방치되는 것을
// 막기 위한 자동 잠금. 한동안(10~15분) 마우스·키보드 조작이 없으면 화면을 잠그고, 접근 신청
// 비밀번호를 다시 입력해야 계속 쓸 수 있게 한다. 뒤로가기/닫기로 우회할 수 없도록 "취소"
// 버튼을 두지 않는다(단, 브라우저를 새로고침하면 다시 열릴 수 있음 — 이 프로그램 전체가
// "코드를 아는 사람은 접근 가능" 수준의 보안이라는 점과 같은 한계).
function IdleLockScreen({ onUnlock, error }) {
  const [pw, setPw] = useState("");
  return (
    <div className="modal-backdrop idle-lock-backdrop">
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><ShieldCheck size={18} /> 자동 잠금</h3>
        </div>
        <div className="modal-body">
          <div className="text-dim small-note">
            일정 시간 조작이 없어 화면이 자동으로 잠겼습니다(공용 PC에서 학생 개인정보 노출을
            막기 위한 기능). 계속하려면 접근 신청 비밀번호를 입력해 주세요.
          </div>
          <input
            className="input"
            type="password"
            value={pw}
            onChange={e => setPw(e.target.value)}
            placeholder="접근 신청 비밀번호"
            onKeyDown={e => { if (e.key === "Enter") onUnlock(pw); }}
            autoFocus
          />
          {error && <div className="gate-error">{error}</div>}
          <div className="confirm-actions">
            <button className="btn btn-primary" onClick={() => onUnlock(pw)}>잠금 해제</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function CodeCreatedNoticeModal({ workspaceCode, viewerPassword, founderPassword, onClose }) {
  function downloadMemo() {
    const lines = [
      "SMART PAPS 학교 코드 메모",
      "",
      "학교 코드: " + workspaceCode,
      "접근 신청 비밀번호(동료 교사용): " + (viewerPassword || "(설정 안 함)"),
      "개설자 전용 비밀번호(본인만): " + (founderPassword || "(설정 안 함)"),
      "",
      "이 파일을 잃어버리면 비밀번호를 되찾을 방법이 없습니다. 안전한 곳에 보관하세요.",
    ];
    const blob = new Blob([lines.join("\n")], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "smart-paps-" + workspaceCode + "-메모.txt";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><CheckCircle2 size={18} color="var(--gold)" /> 코드가 만들어졌습니다</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <div className="text-dim small-note">
            이 코드와 비밀번호는 <b>이메일이나 계정이 없어 잊어버리면 되찾을 방법이 없습니다.</b>
            아래 버튼으로 메모 파일을 받아 꼭 저장해 두세요(개인 노트, 클라우드 드라이브 등).
          </div>
          <div className="closeout-summary">
            학교 코드: <b>{workspaceCode}</b><br />
            접근 신청 비밀번호: <b>{viewerPassword || "설정 안 함"}</b><br />
            개설자 전용 비밀번호: <b>{founderPassword || "설정 안 함"}</b>
          </div>
          <button className="btn btn-primary big-btn" onClick={downloadMemo}>
            <Copy size={14} /> 메모 파일 다운로드
          </button>
          <button className="btn btn-ghost" onClick={onClose} style={{ width: "100%", justifyContent: "center", marginTop: 8 }}>
            나중에 하기
          </button>
        </div>
      </div>
    </div>
  );
}

function ShareGuideModal({ workspaceCode, onClose }) {
  const [codeCopied, setCodeCopied] = useState(false);
  const [linkCopied, setLinkCopied] = useState(false);
  const { canInstall, promptInstall, isIOS, isStandalone } = usePwaInstall();

  async function copyCode() {
    try {
      await navigator.clipboard.writeText(workspaceCode || "");
      setCodeCopied(true);
      setTimeout(() => setCodeCopied(false), 2000);
    } catch (e) {
      setCodeCopied(false);
    }
  }

  async function copyLink() {
    try {
      await navigator.clipboard.writeText(window.location.origin);
      setLinkCopied(true);
      setTimeout(() => setLinkCopied(false), 2000);
    } catch (e) {
      setLinkCopied(false);
    }
  }

  const steps = [
    { title: "프로그램 링크 열기", body: "스마트폰 브라우저에서 이 프로그램 주소를 직접 입력하거나, 아래 링크 복사 버튼으로 복사해 문자·메신저로 보내서 열어주세요." },
    { title: "학교 코드 입력", body: `학교 코드("${workspaceCode}")를 입력하면 바로 같은 데이터로 접속됩니다.` },
    { title: "동료 교사에게는 '접근 신청' 안내", body: "코드를 직접 알려주기보다, 접근권한 탭에서 비밀번호를 정해두고 동료 선생님이 '접근 신청'(이름+코드+비밀번호)으로 들어오게 하는 걸 권장해요." },
    { title: "홈 화면에 앱처럼 설치(선택)", body: "아래 '홈 화면에 설치' 버튼을 쓰면 앱처럼 아이콘이 생겨 더 빠르게 열 수 있어요." },
  ];

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><Smartphone size={18} /> 모바일로 함께 보기</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <div className="text-dim small-note">
            노트북이 없어도 스마트폰으로 실시간 기록을 확인·입력할 수 있게 해주는 기능입니다.
          </div>
          <div className="warn-note">
            <AlertTriangle size={16} />
            <span><b>링크</b>(프로그램을 여는 열쇠)와 <b>코드</b>(우리 학교 데이터로 들어가는 열쇠)는 서로 다릅니다. 둘 다 있어야 접속돼요. 코드는 본인 기기 외에는 신뢰하는 동료 교사에게만 알려주세요.</span>
          </div>
          {steps.map((s, i) => (
            <div className="share-step" key={i}>
              <span className="share-step-num">{i + 1}</span>
              <div>
                <div className="share-step-title">{s.title}</div>
                <div className="share-step-body">{s.body}</div>
              </div>
            </div>
          ))}
          <div className="share-link-row">
            <input className="input" readOnly value={workspaceCode || ""} onFocus={e => e.target.select()} />
            <button className="btn btn-secondary" onClick={copyCode}>
              {codeCopied ? <><Check size={14} /> 복사됨</> : <><Copy size={14} /> 코드 복사</>}
            </button>
          </div>
          <div className="share-link-row">
            <input className="input" readOnly value={typeof window !== "undefined" ? window.location.origin : ""} onFocus={e => e.target.select()} />
            <button className="btn btn-secondary" onClick={copyLink}>
              {linkCopied ? <><Check size={14} /> 복사됨</> : <><Copy size={14} /> 링크 복사</>}
            </button>
          </div>

          <div className="divider" />
          <h3>홈 화면에 앱처럼 설치</h3>
          {isStandalone ? (
            <div className="text-dim small-note">이미 홈 화면 앱으로 열려 있습니다.</div>
          ) : isIOS ? (
            <div className="text-dim small-note">
              아이폰(사파리)에서는 자동 설치 버튼이 없어요. 하단의 <b>공유</b> 버튼을 누른 뒤{" "}
              <b>"홈 화면에 추가"</b>를 직접 눌러주세요.
            </div>
          ) : canInstall ? (
            <button className="btn btn-primary" onClick={promptInstall}>
              <Smartphone size={14} /> 홈 화면에 설치
            </button>
          ) : (
            <div className="text-dim small-note">
              브라우저 메뉴에서 "홈 화면에 추가" 또는 "앱 설치"를 눌러도 같은 효과가 있어요.
              (이미 설치되어 있거나, 이 브라우저가 자동 설치를 지원하지 않을 수 있습니다.)
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function ScoreBoard({ students, records, activeYear, studentValue, studentGrade, skippedEvents, schoolGrades, presentation, setPresentation, onExitPresentation, lastSync }) {
  const [boardMode, setBoardMode] = useState("single"); // 'single' | 'all'
  const [eventId, setEventId] = useState(EVENTS[0].id);
  const [unit, setUnit] = useState("student"); // 'student' | 'class'
  const [gender, setGender] = useState("ALL");
  const [schoolGrade, setSchoolGrade] = useState("ALL");
  const [classNum, setClassNum] = useState("ALL");

  const activeEvents = useMemo(() => {
    const skipped = computeEffectiveSkip(skippedEvents);
    return EVENTS.filter(e => !skipped.has(e.id));
  }, [skippedEvents]);

  const classOptions = useMemo(() => {
    const pool = schoolGrade === "ALL" ? students : students.filter(s => s.grade === Number(schoolGrade));
    const set = new Set(pool.map(s => s.classNum));
    return Array.from(set).sort((a, b) => a - b);
  }, [students, schoolGrade]);

  const filtered = useMemo(() => {
    return students.filter(s => {
      if (gender !== "ALL" && s.gender !== gender) return false;
      if (schoolGrade !== "ALL" && s.grade !== Number(schoolGrade)) return false;
      if (classNum !== "ALL" && s.classNum !== Number(classNum)) return false;
      return true;
    });
  }, [students, gender, schoolGrade, classNum]);

  const isOverall = eventId === "overall";
  const ev = isOverall ? null : EVENT_MAP[eventId];

  const studentAvgGrade = useCallback((s) => {
    const grades = activeEvents.map(e => studentGrade(s, e.id, activeYear)).filter(g => g !== null);
    if (grades.length === 0) return null;
    return grades.reduce((a, b) => a + b, 0) / grades.length;
  }, [studentGrade, activeYear, activeEvents]);

  const ranking = useMemo(() => {
    if (unit === "class") {
      // 반별 랭킹: 학생 개인 기록이 아니라, 그 반 학생들이 받은 등급의 평균으로 반 사이의 순위를 매긴다.
      // (반 필터는 랭킹 대상 자체를 좁히는 게 아니라 "어떤 반이 있는지" 보는 용도라 여기서는 학년/성별만 반영)
      const pool = students.filter(s => {
        if (gender !== "ALL" && s.gender !== gender) return false;
        if (schoolGrade !== "ALL" && s.grade !== Number(schoolGrade)) return false;
        return true;
      });
      const groups = {};
      pool.forEach(s => {
        const key = s.grade + "-" + s.classNum;
        if (!groups[key]) groups[key] = { grade: s.grade, classNum: s.classNum, students: [] };
        groups[key].students.push(s);
      });
      const list = Object.values(groups).map(g => {
        const grades = g.students
          .map(s => (isOverall ? studentAvgGrade(s) : studentGrade(s, eventId, activeYear)))
          .filter(x => x !== null);
        if (grades.length === 0) return null;
        const avg = grades.reduce((a, b) => a + b, 0) / grades.length;
        return {
          id: g.grade + "-" + g.classNum,
          name: g.grade + "학년 " + g.classNum + "반",
          meta: grades.length + "명 측정",
          display: "평균 " + (Math.round(avg * 100) / 100) + "등급",
          grade: Math.round(avg),
          avg,
        };
      }).filter(Boolean);
      list.sort((a, b) => a.avg - b.avg);
      return list;
    }
    if (isOverall) {
      // 종합 랭킹은 체력 종목(EVENTS)만 반영하고 BMI는 제외한다.
      // BMI는 순위 경쟁으로 노출되면 학생, 특히 여학생에게 민감할 수 있어
      // 개인별 등급표(종목별 등급 참고) 용도로만 별도로 보여준다.
      const list = filtered
        .map(s => {
          const avg = studentAvgGrade(s);
          if (avg === null) return null;
          return { id: s.id, name: s.grade + "학년 " + s.classNum + "반 " + s.number + "번", meta: "", avg, grade: Math.round(avg) };
        })
        .filter(Boolean);
      list.sort((a, b) => a.avg - b.avg);
      return list.map(x => ({ ...x, display: "평균 " + (Math.round(x.avg * 100) / 100) + "등급" }));
    }
    const list = filtered
      .map(s => ({ s, v: studentValue(s.id, eventId, activeYear) }))
      .filter(x => x.v !== null);
    list.sort((a, b) => (ev.better === "high" ? b.v - a.v : a.v - b.v));
    return list.map(x => ({
      id: x.s.id,
      name: x.s.grade + "학년 " + x.s.classNum + "반 " + x.s.number + "번",
      meta: "",
      display: fmtValue(x.v, ev.unit),
      grade: studentGrade(x.s, eventId, activeYear),
    }));
  }, [filtered, students, unit, gender, schoolGrade, eventId, isOverall, studentValue, studentGrade, studentAvgGrade, activeYear, ev]);

  // 전체 화면 랭킹모드: 활성화된 모든 종목의 상위 3명을 한 화면에 카드로 모아 보여준다.
  const overviewData = useMemo(() => {
    if (boardMode !== "all") return [];
    const allEvs = activeEvents;
    return allEvs.map(e2 => {
      const list = filtered
        .map(s => {
          if (e2.id === "overall") {
            const avg = studentAvgGrade(s);
            if (avg === null) return null;
            return { id: s.id, name: s.grade + "학년 " + s.classNum + "반 " + s.number + "번", display: "평균 " + (Math.round(avg * 100) / 100) + "등급", sortVal: avg };
          }
          const v = studentValue(s.id, e2.id, activeYear);
          if (v === null) return null;
          return { id: s.id, name: s.grade + "학년 " + s.classNum + "반 " + s.number + "번", display: fmtValue(v, e2.unit), sortVal: v };
        })
        .filter(Boolean);
      list.sort((a, b) => (e2.better === "high" ? b.sortVal - a.sortVal : a.sortVal - b.sortVal));
      return { event: e2, top: list.slice(0, 3), total: list.length };
    });
  }, [boardMode, activeEvents, filtered, studentValue, studentAvgGrade, activeYear]);

  const isGradeAvgDisplay = unit === "class" || isOverall;
  const top3 = ranking.slice(0, 3);
  const rest = ranking.slice(3, 12);

  return (
    <div className="board">
      <div className="board-head">
        <div className="board-title">
          <span className="board-event">
            {isOverall ? "종합 랭킹(참고용 평균)" : ev.name}
            {!isOverall && <span className="board-cat"> ({ev.category})</span>}
          </span>
        </div>
        {presentation && (
          <button className="btn btn-ghost small" onClick={onExitPresentation}>
            <Minimize2 size={14} /> 나가기
          </button>
        )}
      </div>

      <div className="board-filters">
        <FilterChips
          label="화면"
          value={boardMode}
          onChange={setBoardMode}
          options={[{ id: "single", label: "개별종목 랭킹" }, { id: "all", label: "전체 화면 랭킹" }]}
        />
        {boardMode === "single" && (
          <>
            <FilterChips
              label="단위"
              value={unit}
              onChange={setUnit}
              options={[{ id: "student", label: "개인별" }, { id: "class", label: "반별" }]}
            />
            <FilterChips
              label="종목"
              value={eventId}
              onChange={setEventId}
              options={EVENTS.map(e => ({ id: e.id, label: e.name }))}
            />
          </>
        )}
        <FilterChips
          label="성별"
          value={gender}
          onChange={setGender}
          options={[{ id: "ALL", label: "전체" }, ...GENDERS.map(g => ({ id: g.id, label: g.label }))]}
        />
        <FilterChips
          label="학년"
          value={schoolGrade}
          onChange={(v) => { setSchoolGrade(v); setClassNum("ALL"); }}
          options={[{ id: "ALL", label: "전체" }, ...schoolGrades.map(g => ({ id: String(g), label: g + "학년" }))]}
        />
        {boardMode === "single" && unit === "student" && (
          <FilterChips
            label="반"
            value={classNum}
            onChange={setClassNum}
            disabled={schoolGrade === "ALL"}
            options={[{ id: "ALL", label: "전체" }, ...classOptions.map(c => ({ id: String(c), label: c + "반" }))]}
          />
        )}
        {boardMode === "all" && (
          <FilterChips
            label="반"
            value={classNum}
            onChange={setClassNum}
            disabled={schoolGrade === "ALL"}
            options={[{ id: "ALL", label: "전체" }, ...classOptions.map(c => ({ id: String(c), label: c + "반" }))]}
          />
        )}
      </div>

      {boardMode === "all" ? (
        <div className="overview-grid">
          {overviewData.map(({ event, top, total }) => (
            <div className="overview-card" key={event.id}>
              <div className="overview-card-head">{event.name}</div>
              {top.length === 0 ? (
                <div className="overview-empty">기록 없음</div>
              ) : (
                <div className="overview-list">
                  {top.map((item, i) => (
                    <div className="overview-row" key={item.id}>
                      <span className={"overview-rank r" + (i + 1)}>{i + 1}</span>
                      <span className="overview-name">{item.name}</span>
                      <span className="overview-value">{item.display}</span>
                    </div>
                  ))}
                </div>
              )}
              <div className="overview-foot">{total}명 측정</div>
            </div>
          ))}
        </div>
      ) : (
        <>
      {isGradeAvgDisplay && ranking.length > 0 && (
        <div className="rank-note">※ 등급 평균이 낮을수록(1등급에 가까울수록) 좋은 성적입니다. 숫자가 크다고 좋은 게 아니에요.</div>
      )}

      {ranking.length === 0 ? (
        <div className="empty-state">
          <Trophy size={40} color="var(--text-dim)" />
          <div>아직 이 조건에 해당하는 기록이 없습니다.</div>
          <div className="empty-sub">기록입력 화면에서 학생 기록을 입력해 주세요.</div>
        </div>
      ) : (
        <>
          <div className="podium">
            {[1, 0, 2].map(pos => {
              const item = top3[pos];
              if (!item) return <div key={pos} className="podium-slot empty" />;
              const rank = pos + 1;
              return <PodiumCard key={item.id} item={item} rank={rank} />;
            })}
          </div>

          {rest.length > 0 && (
            <div className="rank-list">
              {rest.map((item, i) => (
                <div className="rank-row" key={item.id}>
                  <span className="rank-num">{i + 4}</span>
                  <span className="rank-name">{item.name}</span>
                  <span className="rank-meta">{item.meta}</span>
                  {item.grade && (
                    <span className="grade-dot" style={{ background: GRADE_COLORS[item.grade] }}>{item.grade}</span>
                  )}
                  <span className="rank-value">{item.display}</span>
                </div>
              ))}
            </div>
          )}
          <div className="board-foot">
            총 {ranking.length}{unit === "class" ? "개 반" : "명"} 측정 · {rest.length > 0 ? "상위 " + Math.min(ranking.length, 12) + (unit === "class" ? "개 반 표시" : "명 표시") : ""}
          </div>
        </>
      )}
        </>
      )}
    </div>
  );
}

function PodiumCard({ item, rank }) {
  const tierClass = rank === 1 ? "tier-1" : rank === 2 ? "tier-2" : "tier-3";
  return (
    <div className={"podium-card " + tierClass}>
      {rank === 1 && <div className="podium-shine" />}
      {rank === 1 && <div className="podium-burst" />}
      <div className="podium-rank">
        {rank === 1 ? <Trophy size={40} className="podium-trophy-icon" /> : <Medal size={24} />}
        <span>{rank}</span>
      </div>
      <div className="podium-name">{item.name}</div>
      <div className="podium-meta">{item.meta}</div>
      <div className="podium-value">{item.display}</div>
      {item.grade && (
        <div className="grade-dot big" style={{ background: GRADE_COLORS[item.grade] }}>{item.grade}등급</div>
      )}
    </div>
  );
}

function FilterChips({ label, value, onChange, options, disabled }) {
  return (
    <div className={"filter-group" + (disabled ? " disabled" : "")}>
      <span className="filter-label">{label}</span>
      <div className="chip-row">
        {options.map(opt => (
          <button
            key={opt.id}
            className={"chip" + (value === opt.id ? " active" : "")}
            onClick={() => !disabled && onChange(opt.id)}
            disabled={disabled}
          >
            {opt.label}
          </button>
        ))}
      </div>
    </div>
  );
}

// 샌드박스 환경에서는 브라우저 기본 window.confirm()이 동작하지 않을 수 있어
// (아무 반응 없이 무시됨) 화면 안에서 직접 확인을 받는 모달을 사용한다.
function ConfirmModal({ title, message, confirmLabel = "확인", danger, onConfirm, onCancel }) {
  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal-panel confirm-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><AlertTriangle size={18} /> {title}</h3>
          <button className="icon-btn" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <div className="confirm-message">{message}</div>
          <div className="confirm-actions">
            <button className="btn btn-ghost" onClick={onCancel}>취소</button>
            <button className={"btn " + (danger ? "btn-primary" : "btn-secondary")} onClick={onConfirm}>{confirmLabel}</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== 기록 입력 ============================== */

function inferSchoolGradeAtYear(student, year) {
  const diff = thisYear() - year;
  const g = student.grade - diff;
  if (g >= 1 && g <= 6) return g;
  return student.grade;
}

// BMI는 교육부 학생건강정보센터의 공식 5단계 분류(마름/정상/과체중/경도비만/고도비만)를
// 그대로 안내에 사용한다. 등급이나 종합점수에는 전혀 반영되지 않는, 참고용 표시다.
function describeBmiCategory(bmi, student, schoolLevel) {
  const cat = classifyBmi(bmi, student.grade, student.gender, schoolLevel);
  return cat ? `또래 기준 ${cat}` : "";
}

/* ============================== 기록관리 (기록입력 + 타이머 통합) ============================== */

function groupEventsByCategory() {
  const order = [];
  const map = {};
  ALL_EVENTS.forEach(ev => {
    if (!map[ev.category]) { map[ev.category] = []; order.push(ev.category); }
    map[ev.category].push(ev);
  });
  return order.map(c => ({ category: c, events: map[c] }));
}
const EVENT_CATEGORY_GROUPS = groupEventsByCategory();
// 체력측정요소(카테고리)별 색상 — 종목 선택 화면에서 "측정요소"와 "종목"을 시각적으로
// 구분하기 위한 용도.
const CATEGORY_COLORS = {
  "심폐지구력": "#4EA8DE",
  "근력·근지구력": "#7FD98A",
  "유연성": "#B983FF",
  "순발력": "#FFC93C",
  "신체구성": "#E85D5D",
};

// 교육부 학생건강정보센터 PAPS 자료(측정방법 안내 + 참고 영상)를 종목별로 정리한 것.
// 이 환경(아티팩트)은 보안상 외부 영상을 화면 안에 직접 끼워 넣을 수 없어서(유튜브 임베드 때와
// 동일한 제약), QR코드 이미지 대신 "새 탭에서 열기" 버튼으로 안내한다.
// 교육부 학생건강정보센터 PAPS 공식 음원(측정 시 재생용). 원본은 각각 3분(스텝검사, 고정 길이 그대로 사용)과
// 10분 이상(윗몸말아올리기, 3초 간격 신호음이 끝까지 반복되는 구조라 대표 구간 12초만 추출해 반복 재생)이다.
const STEP_TEST_AUDIO_B64 = "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hx9+GFE0aNBAMlCymBha3qpVq9dff12iUDJDJMrhWn6tVasWMqxXr55UyiX8S8uRGB0XOWN0evfuXalSJeVgqZNRYq8pH01lsPgzZEH0XsrkmGsZC0aWYmk/ApR6xViIqZIvlUpjpDSxODa9bTsuaqOilkiaZJt7scvCbVd0IZB/kYzIwTZkjJGUw0+cJ79tmMS7SH9lZMXT2FSXXjO+PXr0yJcvn4Rd0kjG1GlQKlasOHjw4BIlSkiNlO/k/98RTgKLjYcAgwYNeuCBB9atW7dixYp//etfffr06du3LxnQpzvvvPPw4cMrV6688cYbGXJItWbNmquvvlpZkrUjDVtAhQsXvvbaazdv3swlH374Ido2bty4Q4cONWvWTOzFgw8+yE8o38yZM7/++utVq1Z99dVXSBnunT9/funSpbt37164cCGl0RIyINa33377xIkTu3btMgxj6tSpw4YN46Br167Iulq1aj/88MPevXs5j8Jx/qmnntqyZUvdunWnTJnCJTt27Dh16hQsYqg++eQTGg835s2bd+DAgX379h0/fnz79u0DBw6kRrKVK1euTZs2NIlsy5cvX7x4ccuWLb/88ssvvvgCgtlyo6iOHTs+9thj33777datW7/77juUYMyYMQgEAwGfMVhPPPEEPf3ss8/IsGfPHiSMtDGRVEqTJk2aRDm33HILXVaWLdu4cSOdJf8333xz11130Ufa07NnT7tS1PHFF19ctmwZB8pSWYbp4MGDYiXtwKFx48aijgwWBgLBqmgYYvthoYTt0LBNFEt3+JPLkR4ly2hKHtvXiZ2ScmznLxolntPO6bOW9ATkxyQJqaQB/CnlCF1tLZIDOCk5beY74YwI7CBLeiQn7VCoSZMmO3fuHDJkiJge6hKTYccL/IuJ/+WXX9A9pCp1OSPKvxfs2OYCJwy1Pvjgg99++w2iQoCjR4/SJTwDPw0YMAAlfvLJJxcsWID2kBNdR79x13I5Eq9RowZssQvEX3366aeUBks3bNiAg8J97d+/f9SoURISv/zyy/xapEiRm266iQxr165977338EsdOnRAF3FZQkJh5q+//kqZ8Kp58+a4CH4aO3YspELdccJYCphACxctWkQjMSuwEZJAy8qVKzdt2nTWrFn4cCwCJMT0HDlyhKpxzmj2Tz/9BKngGBn4CYvDYOPiKAQScgazAvdw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";
const SITUP_LOOP_AUDIO_B64 = "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";

const EVENT_METHOD_INFO = {
  shuttlerun: { tip: "신호음에 맞춰 출발점과 반대편을 왕복합니다. 다음 신호음이 울리기 전 도착하지 못하면 측정을 종료합니다.", videoUrl: "https://www.youtube.com/watch?v=FcFI01Edx6E" },
  run_walk: { tip: "정해진 거리를 편평한 코스에서 완주하는 데 걸린 시간을 측정합니다. 거리 기준: 초등학생 1000m, 여자 중·고등학생 1200m, 남자 중·고등학생 1600m.", videoUrl: "https://www.youtube.com/watch?v=pY6PWQtSD0I" },
  step_test: { tip: "스텝박스를 3분간 오르내린 뒤 의자에 앉아, 1분·2분·3분 시점의 심박수를 각각 측정해 입력하면 심폐효율지수(PEI)가 자동으로 계산됩니다.", videoUrl: "https://www.youtube.com/watch?v=mr7qrjq3j0Q" },
  sitreach: { tip: "다리를 곧게 뻗고 앉아 상체와 손을 최대한 앞으로 굽혀 뻗은 거리를 측정합니다.", videoUrl: "https://www.youtube.com/watch?v=g0Q9zWuAFUg" },
  flex_total: { tip: "어깨·몸통·옆구리·하체 4부분을 각각 좌우 한 번씩 검사합니다. 양쪽 모두 성공 2점, 한쪽만 성공 1점, 모두 실패 0점이며, 4부분 점수를 합산(0~8점)해 기록합니다.", videoUrl: "https://www.youtube.com/watch?v=r0nVnsqWYpk" },
  pushup: { tip: "머리부터 발끝까지 일직선을 유지하며, 가슴이 봉(또는 바닥)에 가까이 닿을 때까지 팔굽혀펴기를 반복합니다(여학생은 무릎대고 팔굽혀펴기).", videoUrl: "https://www.youtube.com/watch?v=BMRmVDVeUtg" },
  situp: { tip: "무릎을 세우고 누운 자세에서, 3초 간격 신호음에 맞춰 상체를 말아올려 손으로 무릎을 감쌉니다.", videoUrl: "https://www.youtube.com/watch?v=2GVpH3X9k68" },
  gripstrength: { tip: "직립 자세에서 악력계를 손에 맞게 폭을 조절한 뒤, 좌우 각각 2회씩 측정합니다.", videoUrl: "https://www.youtube.com/watch?v=k5r-uNie16c" },
  fifty_m: { tip: "스탠딩스타트 자세로 출발선에서 출발해 50m를 최대한 빠르게 달립니다.", videoUrl: "https://www.youtube.com/watch?v=aYR_B-jyUnM" },
  longjump: { tip: "제자리에서 한 번만 굴러 최대한 멀리 뛰며, 2회 실시해 좋은 기록을 0.1cm 단위로 채택합니다.", videoUrl: "https://www.youtube.com/watch?v=wPm-zaUI8G0" },
};

function RecordManagementView({ students, records, activeYear, onSave, studentValue, studentParts, settings, setSettings, showToast }) {
  const [eventId, setEventId] = useState(EVENTS[0].id);
  const [grade, setGrade] = useState(null);
  const [classNum, setClassNum] = useState(null);
  const ev = EVENT_MAP[eventId];
  const skippedEvents = settings.skippedEvents || [];
  const schoolGrades = gradesForLevel(settings.schoolLevel || "middle");
  // 체육관 등에서 프로젝터·대형화면에 이 화면을 띄운 채 측정할 때, 학생 이름이 그대로
  // 노출되지 않도록 "김*동" 형태로 가리는 모드. 기본은 꺼짐(평소 입력 편의를 위해 실명
  // 표시), 필요할 때만 교사가 켠다. 화면을 새로고침하면 다시 꺼짐 상태로 돌아온다(측정
  // 세션마다 매번 의식적으로 켜도록 하기 위함).
  const [maskNames, setMaskNames] = useState(false);
  // BMI 입력 화면에서 학생마다 하나씩 "직접 입력" 체크박스를 누르지 않아도, 버튼 하나로
  // 현재 보이는 반 전체를 한 번에 전체선택/전체해제할 수 있게 하기 위한 신호(epoch가 바뀔
  // 때마다 BmiRow들이 자신의 직접입력 모드를 forceValue로 맞춘다). 그 이후 개별 학생이
  // 다시 체크박스를 직접 누르면 그 학생만 원래대로 개별 조정 가능하다.
  const [bmiDirectEpoch, setBmiDirectEpoch] = useState(0);
  const [bmiDirectForceValue, setBmiDirectForceValue] = useState(true);
  function applyBmiDirectAll(value) {
    setBmiDirectForceValue(value);
    setBmiDirectEpoch(e => e + 1);
  }

  const gradeStats = useMemo(() => {
    const stats = {};
    schoolGrades.forEach(g => {
      const pool = students.filter(s => s.grade === g);
      const done = pool.filter(s => studentValue(s.id, eventId, activeYear) !== null).length;
      stats[g] = { total: pool.length, done };
    });
    return stats;
  }, [students, eventId, activeYear, studentValue]);

  const classOptions = useMemo(() => {
    if (grade === null) return [];
    return Array.from(new Set(students.filter(s => s.grade === grade).map(s => s.classNum))).sort((a, b) => a - b);
  }, [students, grade]);

  const classStats = useMemo(() => {
    const stats = {};
    classOptions.forEach(c => {
      const pool = students.filter(s => s.grade === grade && s.classNum === c);
      const done = pool.filter(s => studentValue(s.id, eventId, activeYear) !== null).length;
      stats[c] = { total: pool.length, done };
    });
    return stats;
  }, [students, grade, classOptions, eventId, activeYear, studentValue]);

  const classStudents = useMemo(() => {
    if (grade === null || classNum === null) return [];
    return students.filter(s => s.grade === grade && s.classNum === classNum).sort((a, b) => a.number - b.number);
  }, [students, grade, classNum]);

  const eventStatsForClass = useMemo(() => {
    const stats = {};
    ALL_EVENTS.forEach(e2 => {
      const done = classStudents.filter(s => studentValue(s.id, e2.id, activeYear) !== null).length;
      stats[e2.id] = { done, total: classStudents.length };
    });
    return stats;
  }, [classStudents, activeYear, studentValue]);

  function selectEvent(id) {
    setEventId(id);
  }

  function toggleSkip(id, e) {
    e.stopPropagation();
    const next = skippedEvents.includes(id) ? skippedEvents.filter(x => x !== id) : [...skippedEvents, id];
    setSettings({ ...settings, skippedEvents: next });
  }

  const drillDown = (
    <div className="panel drill-panel">
      <div className="drill-panel-head">
        <h3>학년 · 반별 기록 입력 <span className="board-cat">({ev.name}{ev.unit ? ", " + ev.unit : ""})</span></h3>
        <button
          type="button"
          className={"btn btn-secondary small mask-toggle-btn" + (maskNames ? " active" : "")}
          onClick={() => setMaskNames(v => !v)}
          title="체육관 등 화면을 여러 사람이 함께 보는 상황에서, 학생 이름을 가려서 입력할 수 있어요."
        >
          {maskNames ? <EyeOff size={14} /> : <Eye size={14} />}
          {maskNames ? "이름 가림: 켜짐" : "이름 가림: 꺼짐"}
        </button>
      </div>
      <div className="text-dim small-note">아래 학년·반 버튼을 눌러 학생을 고르세요. 숫자는 "입력 완료 인원 / 전체 인원"을 뜻합니다.</div>
      {maskNames && <div className="text-dim small-note">이름을 "홍*동" 형태로 가리는 중입니다. 학번(번호) 순서는 그대로 보이니, 번호로 학생을 확인해 주세요.</div>}
      <div className="drill-row">
        {schoolGrades.map(g => (
          <button key={g} className={"drill-chip" + (grade === g ? " active" : "")} onClick={() => { setGrade(g); setClassNum(null); }}>
            {g}학년 <span className="drill-count">{gradeStats[g].done}/{gradeStats[g].total}</span>
          </button>
        ))}
      </div>
      {grade !== null && (
        <div className="drill-row">
          {classOptions.length === 0 && <span className="text-dim">해당 학년 학생이 없습니다.</span>}
          {classOptions.map(c => (
            <button key={c} className={"drill-chip" + (classNum === c ? " active" : "")} onClick={() => setClassNum(c)}>
              {c}반 <span className="drill-count">{classStats[c].done}/{classStats[c].total}</span>
            </button>
          ))}
        </div>
      )}
      {grade !== null && classNum !== null && (
        <div className="drill-row">
          {ALL_EVENTS.filter(e2 => NO_GRADE_EVENT_IDS.includes(e2.id) || !skippedEvents.includes(e2.id)).map(e2 => (
            <button key={e2.id} className={"drill-chip" + (eventId === e2.id ? " active" : "")} onClick={() => setEventId(e2.id)}>
              {e2.name} <span className="drill-count">{eventStatsForClass[e2.id].done}/{eventStatsForClass[e2.id].total}</span>
            </button>
          ))}
        </div>
      )}
      {grade !== null && classNum !== null && eventId === "bmi" && classStudents.length > 0 && (
        <div className="bmi-bulk-toggle-row">
          <span className="text-dim small-note">직접 입력 일괄 설정:</span>
          <button type="button" className="btn btn-secondary small" onClick={() => applyBmiDirectAll(true)}>전체 선택(직접 입력)</button>
          <button type="button" className="btn btn-secondary small" onClick={() => applyBmiDirectAll(false)}>전체 해제(신장·체중 입력)</button>
        </div>
      )}
      {grade !== null && classNum !== null && (
        <div className="drill-students drill-scroll">
          {classStudents.length === 0 && <div className="text-dim">학생이 없습니다.</div>}
          {classStudents.map(s => (
            <EventInputRow key={s.id} student={s} eventId={eventId} activeYear={activeYear} studentValue={studentValue} studentParts={studentParts} onSave={onSave} schoolLevel={settings.schoolLevel || "middle"}
              forceDirectMode={eventId === "bmi" ? { epoch: bmiDirectEpoch, value: bmiDirectForceValue } : undefined} maskNames={maskNames} />
          ))}
        </div>
      )}
    </div>
  );

  const bigToolEvents = ["shuttlerun", "situp", "step_test"];
  const isBigTool = bigToolEvents.includes(eventId);

  return (
    <div className="record-mgmt">
      <div className="panel">
        <h3>종목 선택</h3>
        <div className="text-dim small-note">체크 해제 = 이 종목은 측정 안 함(등급표에서도 숨김 처리됨). BMI·체지방률은 항상 자동으로 빠집니다.</div>
        <div className="text-dim small-note">심폐지구력·유연성처럼 비슷한 종목이 여러 개면, 다 측정해도 등급엔 1개만 반영됩니다(맨 앞 종목 우선). 특정 종목 하나만 쓰고 싶으면 나머지는 체크를 해제하세요.</div>
        <div className="category-groups">
          {EVENT_CATEGORY_GROUPS.map(g => {
            const color = CATEGORY_COLORS[g.category] || "var(--text-dim)";
            return (
              <div className="category-group" key={g.category} style={{ "--cat-color": color }}>
                <div className="category-label" style={{ background: color }}>{g.category}</div>
                <div className="chip-row">
                  {g.events.map(e => (
                    <button key={e.id} className={"chip event-chip" + (eventId === e.id ? " active" : "")} onClick={() => selectEvent(e.id)}>
                      {e.name}
                      {!NO_GRADE_EVENT_IDS.includes(e.id) && (
                        <span className="chip-skip" onClick={(ev2) => toggleSkip(e.id, ev2)} title="사용(체크 해제 시 등급표에서 제외)">
                          <span className={"mini-check" + (!skippedEvents.includes(e.id) ? " on" : "")}>{!skippedEvents.includes(e.id) && <Check size={9} />}</span>
                        </span>
                      )}
                    </button>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {isBigTool ? (
        <div className="tool-drill-row">
          <div className="panel tool-col">
            <h3>{ev.name} <span className="board-cat">({ev.category})</span></h3>
            {eventId === "shuttlerun" && <ShuttleRunPlayer settings={settings} setSettings={setSettings} />}
            {eventId === "situp" && <SitupAudioPlayer />}
            {eventId === "step_test" && <StepTestPlayer />}
          </div>
          {drillDown}
        </div>
      ) : (
        <>
          {eventId === "run_walk" && <TrackCourseCalculator students={students} schoolGrades={schoolGrades} schoolLevel={settings.schoolLevel || "middle"} />}
          <div className="panel">
            <h3>{ev.name} <span className="board-cat">({ev.category})</span></h3>
            {eventId === "fifty_m" ? (
              <FiftyMGroupTimer eventId={eventId} students={students} activeYear={activeYear} onSave={onSave} showToast={showToast} schoolGrades={schoolGrades} />
            ) : eventId === "run_walk" ? (
              <ClassRunTimer eventId={eventId} students={students} activeYear={activeYear} onSave={onSave} showToast={showToast} maskNames={maskNames} />
            ) : (
              <div className="text-dim small-note">이 종목은 별도 보조 도구가 없습니다. 아래에서 학년·반을 골라 바로 기록을 입력하세요.</div>
            )}
          </div>
          {drillDown}
        </>
      )}

      {EVENT_METHOD_INFO[eventId] && (
        <div className="panel method-ref-panel">
          <h3>측정 방법 참고 <span className="board-cat">({ev.name})</span></h3>
          <div className="method-ref-row">
            <div className="text-dim small-note method-ref-tip">{EVENT_METHOD_INFO[eventId].tip}</div>
            <a
              className="btn btn-secondary method-ref-btn"
              href={EVENT_METHOD_INFO[eventId].videoUrl}
              target="_blank"
              rel="noopener noreferrer"
            >
              <Play size={14} /> 새 탭에서 측정 방법 영상 보기
            </a>
          </div>
          <div className="text-dim small-note">
            교육부 학생건강정보센터 PAPS 안내자료 기준입니다. 버튼을 누르면 새 탭(또는 유튜브 앱)에서 열립니다. 기기에 유튜브 앱이 깔려 있으면 그 앱으로 넘어가려다 실패해서 버튼이 안 눌리는 것처럼 보일 수 있는데, 그럴 땐 아래 주소를 길게 눌러 복사한 뒤 브라우저에 직접 붙여넣어 열어주세요.
          </div>
          <div className="method-ref-url-row">
            <input className="input method-ref-url" readOnly value={EVENT_METHOD_INFO[eventId].videoUrl} onFocus={e => e.target.select()} />
          </div>
        </div>
      )}
    </div>
  );
}

function EventInputRow(props) {
  const { eventId } = props;
  if (eventId === "sitreach" || eventId === "longjump") return <TwoTrialRow {...props} />;
  if (eventId === "gripstrength") return <GripRow {...props} />;
  if (eventId === "bmi") return <BmiRow {...props} />;
  if (eventId === "bodyfat") return <BodyFatRow {...props} />;
  if (eventId === "flex_total") return <FlexTotalRow {...props} />;
  if (eventId === "step_test") return <StepTestRow {...props} />;
  return <SimpleRow {...props} />;
}

// 스텝검사: 매뉴얼 실시방법 그대로, 스텝운동 종료 후 1분·2분·3분 시점의 심박수 3회를
// 입력하면 심폐효율지수(PEI = 운동지속시간(초)×100 / (2×맥박수 합), 0.01단위에서 올림해 0.1단위로
// 기록)를 자동 계산해 저장한다. 운동지속시간은 매뉴얼 규정대로 3분(180초) 완주를 기준으로 한다.
function StepTestRow({ student, activeYear, studentValue, studentParts, onSave, maskNames }) {
  const eventId = "step_test";
  const existing = studentValue(student.id, eventId, activeYear);
  const parts = studentParts(student.id, eventId, activeYear);
  const [hr1, setHr1] = useState(parts ? String(parts.hr1 ?? "") : "");
  const [hr2, setHr2] = useState(parts ? String(parts.hr2 ?? "") : "");
  const [hr3, setHr3] = useState(parts ? String(parts.hr3 ?? "") : "");

  useEffect(() => {
    const p = parts;
    setHr1(p ? String(p.hr1 ?? "") : "");
    setHr2(p ? String(p.hr2 ?? "") : "");
    setHr3(p ? String(p.hr3 ?? "") : "");
  }, [existing]); // eslint-disable-line react-hooks/exhaustive-deps

  async function commit(n1, n2, n3) {
    const H1 = n1 === "" ? null : Number(n1);
    const H2 = n2 === "" ? null : Number(n2);
    const H3 = n3 === "" ? null : Number(n3);
    if ([H1, H2, H3].some(v => v === null || Number.isNaN(v))) return; // 3회 모두 입력해야 계산·저장
    const prev = parts || {};
    if (H1 === (prev.hr1 ?? null) && H2 === (prev.hr2 ?? null) && H3 === (prev.hr3 ?? null)) return;
    // 심폐효율지수(PEI) = 운동지속시간(초) × 100 / (2 × (1·2·3회 맥박수 합), 3분(180초) 완주 기준
    const pei = Math.ceil((180 * 100 / (2 * (H1 + H2 + H3))) * 10) / 10;
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, pei, activeYear, gradeAtMeasure, { hr1: H1, hr2: H2, hr3: H3 });
  }

  return (
    <div className="drill-student-row step-test-row">
      <span className="drill-student-name">{student.number}. {maskNames ? maskStudentName(student.name) : student.name}</span>
      <input className="input drill-student-input" type="number" step="1" value={hr1} placeholder="1분 심박수"
        onChange={e => setHr1(e.target.value)} onBlur={() => commit(hr1, hr2, hr3)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <input className="input drill-student-input" type="number" step="1" value={hr2} placeholder="2분 심박수"
        onChange={e => setHr2(e.target.value)} onBlur={() => commit(hr1, hr2, hr3)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <input className="input drill-student-input" type="number" step="1" value={hr3} placeholder="3분 심박수"
        onChange={e => setHr3(e.target.value)} onBlur={() => commit(hr1, hr2, hr3)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <span className={"drill-best" + (existing !== null ? " has-value" : "")}>{existing !== null ? "PEI " + existing : "-"}</span>
    </div>
  );
}

// 종합유연성: 학생건강체력평가 매뉴얼의 실시방법 그대로, 어깨·몸통·옆구리·하체 4부위를
// 각각 "오른쪽/왼쪽 한 번씩 시행 → 성공 개수(0~2점)"로 판정해 입력하면 합계(0~8점)가
// 자동으로 계산되어 기록·등급에 반영된다.
const FLEX_TOTAL_PARTS = [
  { key: "shoulder", label: "어깨" },
  { key: "torso", label: "몸통" },
  { key: "side", label: "옆구리" },
  { key: "lower", label: "하체" },
];

function FlexTotalRow({ student, activeYear, studentValue, studentParts, onSave, maskNames }) {
  const eventId = "flex_total";
  const existing = studentValue(student.id, eventId, activeYear);
  const parts = studentParts(student.id, eventId, activeYear);
  const [scores, setScores] = useState({
    shoulder: parts?.shoulder ?? null,
    torso: parts?.torso ?? null,
    side: parts?.side ?? null,
    lower: parts?.lower ?? null,
  });

  useEffect(() => {
    setScores({
      shoulder: parts?.shoulder ?? null,
      torso: parts?.torso ?? null,
      side: parts?.side ?? null,
      lower: parts?.lower ?? null,
    });
  }, [existing]); // eslint-disable-line react-hooks/exhaustive-deps

  async function commit(next) {
    const vals = FLEX_TOTAL_PARTS.map(p => next[p.key]);
    if (vals.some(v => v === null || v === undefined)) return; // 4부위 모두 입력해야 저장
    const prev = parts || {};
    const unchanged = FLEX_TOTAL_PARTS.every(p => next[p.key] === (prev[p.key] ?? null));
    if (unchanged) return;
    const total = vals.reduce((a, b) => a + b, 0);
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, total, activeYear, gradeAtMeasure, next);
  }

  function setPart(key, val) {
    const next = { ...scores, [key]: val };
    setScores(next);
    commit(next);
  }

  return (
    <div className="drill-student-row flex-total-row">
      <span className="drill-student-name">{student.number}. {maskNames ? maskStudentName(student.name) : student.name}</span>
      {FLEX_TOTAL_PARTS.map(p => (
        <div className="flex-part-group" key={p.key}>
          <span className="flex-part-label">{p.label}</span>
          <div className="flex-part-btns">
            {[0, 1, 2].map(v => (
              <button
                key={v}
                className={"flex-score-btn" + (scores[p.key] === v ? " active" : "")}
                onClick={() => setPart(p.key, v)}
                title={v === 2 ? "양쪽 모두 성공" : v === 1 ? "한쪽만 성공" : "모두 실패"}
              >
                {v}
              </button>
            ))}
          </div>
        </div>
      ))}
      <span className={"drill-best" + (existing !== null ? " has-value" : "")}>{existing !== null ? "합계 " + existing + "점" : "-"}</span>
    </div>
  );
}
function SimpleRow({ student, eventId, activeYear, studentValue, onSave, maskNames }) {
  const ev = EVENT_MAP[eventId];
  const existing = studentValue(student.id, eventId, activeYear);
  const [value, setValue] = useState(existing === null ? "" : String(existing));

  useEffect(() => {
    setValue(existing === null ? "" : String(existing));
  }, [existing]);

  async function handleBlur() {
    if (value === "") return;
    const num = Number(value);
    if (Number.isNaN(num) || existing === num) return;
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, num, activeYear, gradeAtMeasure);
  }

  return (
    <div className="drill-student-row">
      <span className="drill-student-name">{student.number}. {maskNames ? maskStudentName(student.name) : student.name}</span>
      <input
        className="input drill-student-input"
        type="number"
        step="0.1"
        value={value}
        placeholder={ev.unit}
        onChange={e => setValue(e.target.value)}
        onBlur={handleBlur}
        onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }}
      />
      {existing !== null && <CheckCircle2 size={14} color="#7FD98A" />}
    </div>
  );
}

// 앉아윗몸앞으로굽히기 · 제자리멀리뛰기: 2회 측정해서 더 좋은 값을 기록에 반영한다.
function TwoTrialRow({ student, eventId, activeYear, studentValue, studentParts, onSave, maskNames }) {
  const ev = EVENT_MAP[eventId];
  const existing = studentValue(student.id, eventId, activeYear);
  const parts = studentParts(student.id, eventId, activeYear);
  const [t1, setT1] = useState(parts ? String(parts.trial1 ?? "") : "");
  const [t2, setT2] = useState(parts ? String(parts.trial2 ?? "") : "");

  useEffect(() => {
    const p = parts;
    setT1(p ? String(p.trial1 ?? "") : "");
    setT2(p ? String(p.trial2 ?? "") : "");
  }, [existing]); // eslint-disable-line react-hooks/exhaustive-deps

  async function commit(nt1, nt2) {
    const n1 = nt1 === "" ? null : Number(nt1);
    const n2 = nt2 === "" ? null : Number(nt2);
    const vals = [n1, n2].filter(v => v !== null && !Number.isNaN(v));
    if (vals.length === 0) return;
    const best = Math.max(...vals);
    // 최고기록(best)만 비교하면, 최고기록에 영향을 주지 않는 다른 차수 값을 고쳐도
    // "달라진 게 없다"고 판단해 저장을 건너뛰어 방금 고친 값이 사라진 것처럼 보이는
    // 문제가 있었다. 두 차수 값을 직접 비교해야 실제 변경 여부를 정확히 안다.
    const prevN1 = parts ? (parts.trial1 ?? null) : null;
    const prevN2 = parts ? (parts.trial2 ?? null) : null;
    if (n1 === prevN1 && n2 === prevN2) return;
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, best, activeYear, gradeAtMeasure, { trial1: n1, trial2: n2 });
  }

  return (
    <div className="drill-student-row two-trial">
      <span className="drill-student-name">{student.number}. {maskNames ? maskStudentName(student.name) : student.name}</span>
      <input className="input drill-student-input" type="number" step="0.1" value={t1} placeholder={"1차" + (ev.unit ? "(" + ev.unit + ")" : "")}
        onChange={e => setT1(e.target.value)} onBlur={() => commit(t1, t2)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <input className="input drill-student-input" type="number" step="0.1" value={t2} placeholder={"2차" + (ev.unit ? "(" + ev.unit + ")" : "")}
        onChange={e => setT2(e.target.value)} onBlur={() => commit(t1, t2)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <span className={"drill-best" + (existing !== null ? " has-value" : "")}>{existing !== null ? "최고 " + fmtValue(existing, ev.unit) : "-"}</span>
      {existing !== null && <CheckCircle2 size={14} color="#7FD98A" />}
    </div>
  );
}

// 악력: 좌/우 각 2회씩 측정 → 각 손의 최고값을 구해 평균한 값을 기록에 반영한다.
function GripRow({ student, activeYear, studentValue, studentParts, onSave, maskNames }) {
  const eventId = "gripstrength";
  const existing = studentValue(student.id, eventId, activeYear);
  const parts = studentParts(student.id, eventId, activeYear);
  const [l1, setL1] = useState(parts ? String(parts.left1 ?? "") : "");
  const [r1, setR1] = useState(parts ? String(parts.right1 ?? "") : "");
  const [l2, setL2] = useState(parts ? String(parts.left2 ?? "") : "");
  const [r2, setR2] = useState(parts ? String(parts.right2 ?? "") : "");

  useEffect(() => {
    const p = parts;
    setL1(p ? String(p.left1 ?? "") : ""); setR1(p ? String(p.right1 ?? "") : "");
    setL2(p ? String(p.left2 ?? "") : ""); setR2(p ? String(p.right2 ?? "") : "");
  }, [existing]); // eslint-disable-line react-hooks/exhaustive-deps

  async function commit(nl1, nr1, nl2, nr2) {
    const toNum = v => (v === "" ? null : Number(v));
    const L1 = toNum(nl1), R1 = toNum(nr1), L2 = toNum(nl2), R2 = toNum(nr2);
    const allVals = [L1, R1, L2, R2].filter(v => v !== null && !Number.isNaN(v));
    if (allVals.length === 0) return;
    // 기록에 실제로 반영되는 값은 4칸 중 최고기록이다(다른 2차시기 종목과 동일한 방식).
    // 평균(좌우 각각의 최고기록 평균)은 참고용으로만 별도 표시한다.
    const finalVal = Math.max(...allVals);
    // 값 비교는 4개 입력값 자체로 해야, 최고기록에 영향 없는 칸을 고쳐도 정확히 저장된다.
    const prev = parts || {};
    if (L1 === (prev.left1 ?? null) && R1 === (prev.right1 ?? null) && L2 === (prev.left2 ?? null) && R2 === (prev.right2 ?? null)) return;
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, finalVal, activeYear, gradeAtMeasure, { left1: L1, right1: R1, left2: L2, right2: R2 });
  }

  // 참고용 평균(좌우 각각의 최고기록을 구해 평균) — 화면 표시 전용, 기록·등급에는 반영되지 않는다.
  const leftVals = [parts?.left1, parts?.left2].filter(v => v !== null && v !== undefined && !Number.isNaN(v));
  const rightVals = [parts?.right1, parts?.right2].filter(v => v !== null && v !== undefined && !Number.isNaN(v));
  const bothBest = [leftVals.length ? Math.max(...leftVals) : null, rightVals.length ? Math.max(...rightVals) : null].filter(v => v !== null);
  const avgRef = bothBest.length ? Math.round((bothBest.reduce((a, b) => a + b, 0) / bothBest.length) * 10) / 10 : null;

  return (
    <div className="drill-student-row grip-row">
      <span className="drill-student-name">{student.number}. {maskNames ? maskStudentName(student.name) : student.name}</span>
      <input className="input drill-student-input" type="number" step="0.1" value={l1} placeholder="1차 좌"
        onChange={e => setL1(e.target.value)} onBlur={() => commit(l1, r1, l2, r2)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <input className="input drill-student-input" type="number" step="0.1" value={r1} placeholder="1차 우"
        onChange={e => setR1(e.target.value)} onBlur={() => commit(l1, r1, l2, r2)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <input className="input drill-student-input" type="number" step="0.1" value={l2} placeholder="2차 좌"
        onChange={e => setL2(e.target.value)} onBlur={() => commit(l1, r1, l2, r2)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <input className="input drill-student-input" type="number" step="0.1" value={r2} placeholder="2차 우"
        onChange={e => setR2(e.target.value)} onBlur={() => commit(l1, r1, l2, r2)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <div className="grip-result-col">
        <span className={"drill-best" + (existing !== null ? " has-value" : "")}>{existing !== null ? "최고 " + existing + "kg" : "-"}</span>
        {avgRef !== null && <span className="drill-ref-note">참고 평균 {avgRef}kg</span>}
      </div>
      {existing !== null && <CheckCircle2 size={14} color="#7FD98A" />}
    </div>
  );
}

// BMI: 신장·체중만 입력하면 자동 계산. 등급을 매기지 않고 또래 평균 대비 위치만 안내한다.
function BmiRow({ student, activeYear, studentValue, studentParts, onSave, schoolLevel, forceDirectMode, maskNames }) {
  const eventId = "bmi";
  const existing = studentValue(student.id, eventId, activeYear);
  const parts = studentParts(student.id, eventId, activeYear);
  const [height, setHeight] = useState(parts ? String(parts.height ?? "") : "");
  const [weight, setWeight] = useState(parts ? String(parts.weight ?? "") : "");
  // 신장·체중 대신 BMI 수치 자체를 이미 알고 있어 바로 입력하고 싶은 경우를 위한 직접입력
  // 모드. 이 모드로 저장하면 신장·체중(parts) 없이 BMI 값만 저장된다.
  const [directMode, setDirectMode] = useState(!parts && existing !== null);
  const [directValue, setDirectValue] = useState(!parts && existing !== null ? String(existing) : "");

  useEffect(() => {
    const p = parts;
    setHeight(p ? String(p.height ?? "") : "");
    setWeight(p ? String(p.weight ?? "") : "");
    if (!p && existing !== null) {
      setDirectMode(true);
      setDirectValue(String(existing));
    }
  }, [existing]); // eslint-disable-line react-hooks/exhaustive-deps

  // 상위(RecordManagementView)에서 "전체 선택/전체 해제" 버튼을 눌러 epoch가 바뀌면,
  // 이 학생 행도 그 값으로 직접입력 모드를 일괄 전환한다(그 뒤엔 다시 개별로 바꿀 수 있음).
  useEffect(() => {
    if (!forceDirectMode || !forceDirectMode.epoch) return;
    setDirectMode(forceDirectMode.value);
  }, [forceDirectMode?.epoch]); // eslint-disable-line react-hooks/exhaustive-deps

  async function commit(nh, nw) {
    const h = nh === "" ? null : Number(nh);
    const w = nw === "" ? null : Number(nw);
    if (!h || !w || Number.isNaN(h) || Number.isNaN(w)) return;
    const bmi = Math.round((w / ((h / 100) ** 2)) * 10) / 10;
    const prev = parts || {};
    if (h === (prev.height ?? null) && w === (prev.weight ?? null)) return;
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, bmi, activeYear, gradeAtMeasure, { height: h, weight: w });
  }

  async function commitDirect(nv) {
    const v = nv === "" ? null : Number(nv);
    if (v === null || Number.isNaN(v)) return;
    if (v === existing && !parts) return;
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, v, activeYear, gradeAtMeasure, null);
  }

  const refText = describeBmiCategory(existing, student, schoolLevel);

  return (
    <div className="drill-student-row bmi-row">
      <span className="drill-student-name">{student.number}. {maskNames ? maskStudentName(student.name) : student.name}</span>
      {directMode ? (
        <input className="input drill-student-input" type="number" step="0.1" value={directValue} placeholder="BMI 수치"
          onChange={e => setDirectValue(e.target.value)} onBlur={() => commitDirect(directValue)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      ) : (
        <>
          <input className="input drill-student-input" type="number" step="0.1" value={height} placeholder="신장(cm)"
            onChange={e => setHeight(e.target.value)} onBlur={() => commit(height, weight)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
          <input className="input drill-student-input" type="number" step="0.1" value={weight} placeholder="체중(kg)"
            onChange={e => setWeight(e.target.value)} onBlur={() => commit(height, weight)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
        </>
      )}
      <label className="bmi-direct-toggle">
        <input type="checkbox" checked={directMode} onChange={e => setDirectMode(e.target.checked)} /> 직접 입력
      </label>
      <span className={"drill-best" + (existing !== null ? " has-value" : "")}>{existing !== null ? "BMI " + existing : "-"}</span>
      {existing !== null && <span className="bmi-ref-tag">{refText}</span>}
    </div>
  );
}

// 체지방률: 값 하나만 입력. BMI와 동일하게 등급을 매기지 않고, 종합등급에도 반영되지 않는
// 성장 확인용 참고 지표다(공식 학년·성별 기준표가 없어 분류 문구는 따로 붙이지 않는다).
function BodyFatRow({ student, activeYear, studentValue, onSave, maskNames }) {
  const eventId = "bodyfat";
  const existing = studentValue(student.id, eventId, activeYear);
  const [value, setValue] = useState(existing === null ? "" : String(existing));

  useEffect(() => {
    setValue(existing === null ? "" : String(existing));
  }, [existing]);

  async function commit(nv) {
    const v = nv === "" ? null : Number(nv);
    if (v === null || Number.isNaN(v)) return;
    if (v === existing) return;
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, v, activeYear, gradeAtMeasure);
  }

  const cat = existing !== null ? classifyBodyfat(existing, student.gender) : null;

  return (
    <div className="drill-student-row bodyfat-row">
      <span className="drill-student-name">{student.number}. {maskNames ? maskStudentName(student.name) : student.name}</span>
      <input className="input drill-student-input" type="number" step="0.1" value={value} placeholder="체지방률(%)"
        onChange={e => setValue(e.target.value)} onBlur={() => commit(value)} onKeyDown={e => { if (e.key === "Enter") e.target.blur(); }} />
      <span className={"drill-best" + (existing !== null ? " has-value" : "")}>{existing !== null ? existing + "%" : "-"}</span>
      {cat && <span className="bmi-ref-tag">{cat}</span>}
    </div>
  );
}

// 50m달리기: 최대 3명을 한 조로 묶어 동시에 출발시키고, 각자 결승선을 통과하는 순간
// 해당 번호의 "기록" 버튼을 눌러 시간을 잡는다. 마지막에 한 번에 저장한다.
function StudentPicker({ students, value, onChange, schoolGrades }) {
  const current = students.find(s => s.id === value);
  const [grade, setGrade] = useState(current ? String(current.grade) : "");
  const [classNum, setClassNum] = useState(current ? String(current.classNum) : "");

  const classOptions = useMemo(() => {
    if (grade === "") return [];
    return Array.from(new Set(students.filter(s => s.grade === Number(grade)).map(s => s.classNum))).sort((a, b) => a - b);
  }, [students, grade]);

  const numberOptions = useMemo(() => {
    if (grade === "" || classNum === "") return [];
    return students.filter(s => s.grade === Number(grade) && s.classNum === Number(classNum)).sort((a, b) => a.number - b.number);
  }, [students, grade, classNum]);

  return (
    <div className="fiftym-picker">
      <select className="select" value={grade} onChange={e => { setGrade(e.target.value); setClassNum(""); onChange(""); }}>
        <option value="">학년</option>
        {(schoolGrades || SCHOOL_GRADES).map(g => <option key={g} value={g}>{g}학년</option>)}
      </select>
      <select className="select" value={classNum} onChange={e => { setClassNum(e.target.value); onChange(""); }} disabled={grade === ""}>
        <option value="">반</option>
        {classOptions.map(c => <option key={c} value={c}>{c}반</option>)}
      </select>
      <select className="select" value={value} onChange={e => onChange(e.target.value)} disabled={classNum === ""}>
        <option value="">번호(이름)</option>
        {numberOptions.map(s => <option key={s.id} value={s.id}>{s.number}번 {s.name}</option>)}
      </select>
    </div>
  );
}

function FiftyMGroupTimer({ eventId, students, activeYear, onSave, showToast, schoolGrades }) {
  const [grade, setGrade] = useState("");
  const [classNum, setClassNum] = useState("");
  const [slots, setSlots] = useState([
    { studentId: "", time: null },
    { studentId: "", time: null },
    { studentId: "", time: null },
  ]);
  const [running, setRunning] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const startRef = useRef(0);
  const rafRef = useRef(null);
  const MIN_SLOTS = 1;
  const MAX_SLOTS = 10;

  const classOptions = useMemo(() => {
    if (grade === "") return [];
    return Array.from(new Set(students.filter(s => s.grade === Number(grade)).map(s => s.classNum))).sort((a, b) => a - b);
  }, [students, grade]);

  // 학년·반을 한 번만 고르면, 같은 반 안에서 여러 명을 잴 때마다 매번 학년·반을 다시
  // 고를 필요 없이 번호만 골라 배정하면 된다(실제로 다른 반과 섞어서 재는 경우는 거의 없으므로).
  const classRoster = useMemo(() => {
    if (grade === "" || classNum === "") return [];
    return students.filter(s => s.grade === Number(grade) && s.classNum === Number(classNum)).sort((a, b) => a.number - b.number);
  }, [students, grade, classNum]);

  function tick() {
    setElapsed((Date.now() - startRef.current) / 1000);
    rafRef.current = requestAnimationFrame(tick);
  }

  function start() {
    startRef.current = Date.now();
    setSlots(s => s.map(x => ({ ...x, time: null })));
    setRunning(true);
    rafRef.current = requestAnimationFrame(tick);
  }
  function stopAll() {
    cancelAnimationFrame(rafRef.current);
    setRunning(false);
  }
  function reset() {
    cancelAnimationFrame(rafRef.current);
    setRunning(false);
    setElapsed(0);
    setSlots(s => s.map(x => ({ ...x, time: null })));
  }
  function captureSlot(i) {
    if (!running) return;
    const t = (Date.now() - startRef.current) / 1000;
    setSlots(s => s.map((x, idx) => (idx === i ? { ...x, time: Math.round(t * 100) / 100 } : x)));
  }
  function setSlotStudent(i, id) {
    setSlots(s => s.map((x, idx) => (idx === i ? { ...x, studentId: id } : x)));
  }
  function editSlotTime(i, val) {
    const num = val === "" ? null : Number(val);
    setSlots(s => s.map((x, idx) => (idx === i ? { ...x, time: Number.isNaN(num) ? x.time : num } : x)));
  }
  function addSlot() {
    if (running) return;
    setSlots(s => (s.length >= MAX_SLOTS ? s : [...s, { studentId: "", time: null }]));
  }
  function removeSlot() {
    if (running) return;
    setSlots(s => (s.length <= MIN_SLOTS ? s : s.slice(0, -1)));
  }

  useEffect(() => () => cancelAnimationFrame(rafRef.current), []);
  useEffect(() => {
    // 반이 바뀌면 이전 반 학생이 슬롯에 남아있지 않도록 배정을 초기화한다.
    setSlots(s => s.map(x => ({ ...x, studentId: "" })));
  }, [grade, classNum]);


  async function saveAll() {
    let count = 0;
    for (const slot of slots) {
      if (slot.studentId && slot.time !== null) {
        const student = students.find(s => s.id === slot.studentId);
        if (student) {
          const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
          await onSave(student.id, eventId, slot.time, activeYear, gradeAtMeasure);
          count++;
        }
      }
    }
    if (count > 0) {
      showToast(count + "명의 " + EVENT_MAP[eventId].name + " 기록을 저장했습니다.", "ok");
      reset();
    }
  }

  const canSave = slots.some(s => s.studentId && s.time !== null);

  return (
    <div className="fiftym-group">
      <div className="board-filters compact fiftym-class-row">
        <FilterChips label="학년" value={grade === "" ? "" : String(grade)} onChange={v => { setGrade(Number(v)); setClassNum(""); }}
          options={(schoolGrades || SCHOOL_GRADES).map(g => ({ id: String(g), label: g + "학년" }))} />
        {grade !== "" && (
          <FilterChips label="반" value={classNum === "" ? "" : String(classNum)} onChange={v => setClassNum(Number(v))}
            options={classOptions.map(c => ({ id: String(c), label: c + "반" }))} />
        )}
      </div>
      {classRoster.length === 0 ? (
        <div className="text-dim small-note">먼저 학년·반을 골라주세요. 그 반 학생 전체를 대상으로 번호만 골라 배정하면 됩니다.</div>
      ) : (
        <>
          <div className="fiftym-clock">{elapsed.toFixed(2)}초</div>
          <div className="fiftym-controls">
            {!running ? (
              <button className="btn btn-primary" onClick={start}><Play size={16} /> 그룹 출발</button>
            ) : (
              <button className="btn btn-ghost" onClick={stopAll}><Square size={16} /> 정지</button>
            )}
            <button className="btn btn-ghost" onClick={reset}><RotateCcw size={16} /> 초기화</button>
          </div>
          <div className="fiftym-count-row">
            <span className="fiftym-count-label">측정 인원</span>
            <button className="fiftym-count-btn" disabled={running || slots.length <= MIN_SLOTS} onClick={removeSlot} aria-label="인원 줄이기">−</button>
            <span className="fiftym-count-value">{slots.length}명</span>
            <button className="fiftym-count-btn" disabled={running || slots.length >= MAX_SLOTS} onClick={addSlot} aria-label="인원 늘리기">+</button>
          </div>
          <div className="fiftym-slots">
            {slots.map((slot, i) => (
              <div className="fiftym-slot" key={i}>
                <span className="fiftym-slot-num">{i + 1}번</span>
                <select className="select" value={slot.studentId} onChange={e => setSlotStudent(i, e.target.value)}>
                  <option value="">학생 선택</option>
                  {classRoster.map(s => <option key={s.id} value={s.id}>{s.number}. {s.name}</option>)}
                </select>
                <button className="btn btn-secondary small" disabled={!running || !slot.studentId} onClick={() => captureSlot(i)}>기록</button>
                <input
                  className="input fiftym-slot-time-input"
                  type="number"
                  step="0.01"
                  placeholder="-"
                  value={slot.time !== null ? slot.time : ""}
                  onChange={e => editSlotTime(i, e.target.value)}
                />
              </div>
            ))}
          </div>
          <button className="btn btn-primary big-btn" disabled={!canSave} onClick={saveAll}>기록 저장</button>
          <div className="text-dim small-note">위 "측정 인원"에서 인원 수를 조정한 뒤 그룹으로 묶어 동시에 출발시키고, 각자 결승선을 통과하는 순간 해당 번호의 "기록" 버튼을 누르세요. 잘못 잡혔다면 오른쪽 기록 칸을 눌러 직접 고칠 수 있습니다. 끝나면 "기록 저장"으로 한 번에 반영됩니다.</div>
        </>
      )}
    </div>
  );
}

// 오래달리기-걷기: 50m달리기와 달리 반 전체(최대 30명 안팎)가 한 번에 함께 출발해서
// 각자 완주하는 순간 기록하는 방식이라, 그룹(레인)별 타이머 대신 "학년·반을 고르면
// 그 반 학생 전체가 명단으로 뜨고, 완주할 때마다 이름을 눌러 시간을 기록"하는 방식으로
// 구성한다. 시계 하나로 반 전체를 동시에 재고, 개별 학생을 미리 슬롯에 배정할 필요가 없다.
// 오래달리기-걷기 코스 계산기: 운동장을 자로 재기 어려운 경우, 측정자의 신장으로 보폭을
// 대략 추정하고("걸음 폭 ≈ 신장 × 0.45"는 흔히 쓰이는 근사치), 운동장을 가로·세로로
// 걸었을 때의 걸음 수를 입력하면 운동장 둘레와 목표 거리(여학생 1200m·남학생 1600m)에
// 맞춰 몇 바퀴를 돌아야 하는지, 도착점이 대략 어디쯤인지 간단한 그림으로 보여준다.
// 어디까지나 참고용 추정치이며, 실제로는 줄자 등으로 정확히 재는 것이 가장 정확하다.
function TrackCourseCalculator({ students, schoolGrades, schoolLevel }) {
  const [shape, setShape] = useState("rect"); // 'rect' | 'oval'
  const [grade, setGrade] = useState("ALL");
  const [classNum, setClassNum] = useState("ALL");
  const [height, setHeight] = useState("170");
  const [hSteps, setHSteps] = useState("");
  const [vSteps, setVSteps] = useState("");

  const classOptions = useMemo(() => {
    if (grade === "ALL") return [];
    return Array.from(new Set(students.filter(s => s.grade === Number(grade)).map(s => s.classNum))).sort((a, b) => a - b);
  }, [students, grade]);

  // 지금 측정하려는 반을 고르면, 그 반에 실제로 있는 성별만 걸러서 목표 거리를 보여준다.
  const classGenders = useMemo(() => {
    if (grade === "ALL" || classNum === "ALL") return { M: true, F: true };
    const list = students.filter(s => s.grade === Number(grade) && s.classNum === Number(classNum));
    return { M: list.some(s => s.gender === "M"), F: list.some(s => s.gender === "F") };
  }, [students, grade, classNum]);

  const h = Number(height) || 0;
  const hs = Number(hSteps) || 0;
  const vs = Number(vSteps) || 0;
  const strideM = (h * 0.45) / 100;
  const horizM = strideM * hs;
  const vertM = strideM * vs;
  const radiusM = vertM / 2;
  // 네모형은 직사각형 둘레, 트랙형은 직선구간(가로) 2개 + 반원(세로를 지름으로 하는) 2개로 계산한다.
  const perimeter = shape === "rect" ? 2 * (horizM + vertM) : 2 * horizM + Math.PI * vertM;
  const ready = h > 0 && hs > 0 && vs > 0 && perimeter > 0;

  // 초등학생은 성별 구분 없이 1000m 하나, 중·고등학생은 성별에 따라 1200m(여)/1600m(남)이다.
  const allTargets = schoolLevel === "elem"
    ? [{ label: "초등학생 목표(1000m)", dist: 1000, gender: "ALL" }]
    : [
        { label: "여학생 목표(1200m)", dist: 1200, gender: "F" },
        { label: "남학생 목표(1600m)", dist: 1600, gender: "M" },
      ];
  const targets = allTargets.filter(t => t.gender === "ALL" || classGenders[t.gender]);

  function courseInfo(dist) {
    if (!ready) return null;
    const lapsExact = dist / perimeter;
    const lapsFull = Math.floor(lapsExact);
    const remainder = Math.round((dist - lapsFull * perimeter) * 10) / 10;
    const lapsNeeded = Math.ceil(lapsExact * 10) / 10;
    return { lapsFull, remainder, lapsNeeded };
  }

  // 네모형: 출발점(왼쪽아래) → 아래변(가로) → 오른쪽변(세로) → 윗변(가로) → 왼쪽변(세로) 순서로
  // remainder 거리만큼 이동했을 때의 좌표를 구한다.
  function pointAtRemainderRect(remainder, W, H) {
    let d = remainder;
    if (d <= horizM) return { x: d / (horizM || 1) * W, y: H };
    d -= horizM;
    if (d <= vertM) return { x: W, y: H - (d / (vertM || 1)) * H };
    d -= vertM;
    if (d <= horizM) return { x: W - (d / (horizM || 1)) * W, y: 0 };
    d -= horizM;
    return { x: 0, y: (d / (vertM || 1)) * H };
  }

  // 트랙형: 왼쪽 반원 맨아래(출발) → 아래 직선 → 오른쪽 반원 → 위 직선 → 왼쪽 반원 순서.
  // scale은 실제 m 좌표를 그림 픽셀로 바꾸는 배율이다.
  function pointAtRemainderOval(remainder, scale) {
    const r = radiusM * scale;
    const straight = horizM * scale;
    let d = remainder * scale;
    if (d <= straight) return { x: r + d, y: 2 * r };
    d -= straight;
    const rightArc = Math.PI * r;
    if (d <= rightArc) {
      const theta = Math.PI - (d / r); // r+straight 지점(아래)에서 시계반대로 반원을 돌아 위쪽으로
      return { x: r + straight + r * Math.sin(theta), y: r - r * Math.cos(theta) };
    }
    d -= rightArc;
    if (d <= straight) return { x: r + straight - d, y: 0 };
    d -= straight;
    const theta2 = Math.PI - (d / r);
    return { x: r - r * Math.sin(theta2), y: r + r * Math.cos(theta2) };
  }

  const DW = 240, DH = 150;
  let boxW = DW, boxH = DH, scale = 1;
  if (ready && horizM > 0 && vertM > 0) {
    if (shape === "rect") {
      const ratio = horizM / vertM;
      if (ratio >= DW / DH) { boxW = DW; boxH = DW / ratio; } else { boxH = DH; boxW = DH * ratio; }
    } else {
      const totalW = horizM + vertM; // 반원 반지름 두 개 + 직선
      const totalH = vertM;
      const ratio = totalW / totalH;
      if (ratio >= DW / DH) { boxW = DW; boxH = DW / ratio; } else { boxH = DH; boxW = DH * ratio; }
      scale = boxH / vertM;
    }
  }

  return (
    <div className="panel track-calc-panel">
      <h3>운동장 코스 계산기 <span className="board-cat">(오래달리기-걷기)</span></h3>
      <div className="text-dim small-note">
        줄자 없이도 대략적인 운동장 크기를 가늠할 수 있도록, 측정자의 신장으로 보폭을 추정해
        계산하는 참고용 도구입니다. 실제 정확한 거리는 가능하면 줄자나 트랙 표시로 확인해 주세요.
      </div>
      <div className="board-filters compact">
        <FilterChips label="학년" value={grade} onChange={v => { setGrade(v); setClassNum("ALL"); }}
          options={[{ id: "ALL", label: "전체" }, ...(schoolGrades || SCHOOL_GRADES).map(g => ({ id: String(g), label: g + "학년" }))]} />
        {grade !== "ALL" && (
          <FilterChips label="반" value={classNum} onChange={setClassNum}
            options={[{ id: "ALL", label: "전체" }, ...classOptions.map(c => ({ id: String(c), label: c + "반" }))]} />
        )}
      </div>
      {grade !== "ALL" && classNum !== "ALL" && (
        <div className="text-dim small-note">지금 측정하려는 <b>{grade}학년 {classNum}반</b> 기준으로 목표 거리를 보여드립니다.</div>
      )}
      <div className="track-calc-inputs">
        <div className="form-row">
          <label>측정자 신장(cm)</label>
          <input className="input" type="number" value={height} onChange={e => setHeight(e.target.value)} placeholder="예: 170" />
        </div>
        <div className="form-row">
          <label>{shape === "rect" ? "운동장 가로 걸음 수" : "직선 구간 걸음 수"}</label>
          <input className="input" type="number" value={hSteps} onChange={e => setHSteps(e.target.value)} placeholder="예: 80" />
        </div>
        <div className="form-row">
          <label>{shape === "rect" ? "운동장 세로 걸음 수" : "트랙 폭(곡선 지름) 걸음 수"}</label>
          <input className="input" type="number" value={vSteps} onChange={e => setVSteps(e.target.value)} placeholder="예: 40" />
        </div>
      </div>

      {ready && (
        <>
          <div className="track-calc-summary">
            <div>추정 보폭 <b>{(strideM * 100).toFixed(0)}cm</b></div>
            {shape === "rect" ? (
              <div>가로 약 <b>{horizM.toFixed(1)}m</b> · 세로 약 <b>{vertM.toFixed(1)}m</b></div>
            ) : (
              <div>직선 약 <b>{horizM.toFixed(1)}m</b> · 트랙 폭 약 <b>{vertM.toFixed(1)}m</b>(반지름 {radiusM.toFixed(1)}m)</div>
            )}
            <div>운동장 둘레(한 바퀴) 약 <b>{perimeter.toFixed(1)}m</b></div>
          </div>

          <div className="board-filters compact">
            <FilterChips label="코스 모양" value={shape} onChange={setShape}
              options={[{ id: "rect", label: "네모형(사각 운동장)" }, { id: "oval", label: "트랙형(육상 트랙)" }]} />
          </div>

          <div className="track-calc-diagram-wrap">
            <svg viewBox="0 0 280 190" className="track-calc-svg">
              <g transform={`translate(${(280 - boxW) / 2}, ${(190 - boxH) / 2})`}>
                {shape === "rect" ? (
                  <rect x="0" y="0" width={boxW} height={boxH} fill="none" stroke="var(--text-dim)" strokeWidth="2" strokeDasharray="6 4" />
                ) : (
                  <path
                    d={`M ${boxH / 2},0 L ${boxW - boxH / 2},0 A ${boxH / 2},${boxH / 2} 0 0 1 ${boxW - boxH / 2},${boxH} L ${boxH / 2},${boxH} A ${boxH / 2},${boxH / 2} 0 0 1 ${boxH / 2},0 Z`}
                    fill="none" stroke="var(--text-dim)" strokeWidth="2" strokeDasharray="6 4"
                  />
                )}
                {shape === "rect" ? (
                  <>
                    <circle cx="0" cy={boxH} r="6" fill="#7FD98A" />
                    <text x="10" y={boxH + 4} fontSize="11" fill="#7FD98A">출발/도착 기준점</text>
                    <polygon points={`${boxW * 0.5 - 5},${boxH - 6} ${boxW * 0.5 + 5},${boxH - 6} ${boxW * 0.5},${boxH - 16}`} fill="var(--text-dim)" />
                  </>
                ) : (
                  <>
                    <circle cx={boxH / 2} cy={boxH} r="6" fill="#7FD98A" />
                    <text x={boxH / 2 + 10} y={boxH + 4} fontSize="11" fill="#7FD98A">출발/도착 기준점</text>
                    <polygon points={`${boxH / 2 + 15},${boxH - 6} ${boxH / 2 + 25},${boxH - 6} ${boxH / 2 + 20},${boxH - 16}`} fill="var(--text-dim)" />
                  </>
                )}
                {targets.map((t) => {
                  const info = courseInfo(t.dist);
                  if (!info || info.remainder <= 0.5) return null;
                  const p = shape === "rect" ? pointAtRemainderRect(info.remainder, boxW, boxH) : pointAtRemainderOval(info.remainder, scale);
                  const color = t.gender === "F" ? "#FFD54A" : t.gender === "M" ? "#E85D5D" : "#4EA8DE";
                  const gLabel = t.gender === "F" ? "여" : t.gender === "M" ? "남" : "";
                  return (
                    <g key={t.label}>
                      <circle cx={p.x} cy={p.y} r="5" fill={color} />
                      <text x={p.x + 8} y={p.y + 4} fontSize="10" fill={color}>{gLabel} 도착점</text>
                    </g>
                  );
                })}
              </g>
            </svg>
          </div>

          <div className="track-calc-results">
            {targets.length === 0 && (
              <div className="text-dim small-note">이 반에는 등록된 학생이 없어 목표 거리를 표시할 수 없습니다.</div>
            )}
            {targets.map(t => {
              const info = courseInfo(t.dist);
              return (
                <div className="track-calc-result-row" key={t.label}>
                  <span className="track-calc-result-label">{t.label}</span>
                  <span className="track-calc-result-value">
                    {info.lapsFull}바퀴 + {info.remainder}m
                    {info.remainder <= 0.5 ? " (거의 정확히 나누어떨어짐)" : ""}
                  </span>
                </div>
              );
            })}
          </div>
          <div className="text-dim small-note">
            "N바퀴 + 나머지 Xm"은 마지막 바퀴를 다 돌지 않고 그림의 표시된 지점에서 도착점을 그으면
            된다는 뜻입니다. 매번 도착점을 표시하기 번거롭다면, 출발선을 그 나머지 거리(Xm)만큼
            앞당겨 그어서 "정확히 N+1바퀴"로 완주하도록 코스를 조정하는 것도 좋은 방법입니다.
          </div>
        </>
      )}
    </div>
  );
}

function ClassRunTimer({ eventId, students, activeYear, onSave, showToast, maskNames }) {
  const [grade, setGrade] = useState(null);
  const [classNum, setClassNum] = useState(null);
  const [running, setRunning] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const [finishTimes, setFinishTimes] = useState({}); // studentId -> time
  const startRef = useRef(0);
  const rafRef = useRef(null);

  const gradeOptions = useMemo(() => Array.from(new Set(students.map(s => s.grade))).sort((a, b) => a - b), [students]);
  const classOptions = useMemo(() => {
    if (grade === null) return [];
    return Array.from(new Set(students.filter(s => s.grade === grade).map(s => s.classNum))).sort((a, b) => a - b);
  }, [students, grade]);

  const roster = useMemo(() => {
    if (grade === null || classNum === null) return [];
    return students.filter(s => s.grade === grade && s.classNum === classNum).sort((a, b) => a.number - b.number);
  }, [students, grade, classNum]);

  function tick() {
    setElapsed((Date.now() - startRef.current) / 1000);
    rafRef.current = requestAnimationFrame(tick);
  }
  function start() {
    startRef.current = Date.now();
    setFinishTimes({});
    setRunning(true);
    rafRef.current = requestAnimationFrame(tick);
  }
  function stopClock() {
    cancelAnimationFrame(rafRef.current);
    setRunning(false);
  }
  function reset() {
    cancelAnimationFrame(rafRef.current);
    setRunning(false);
    setElapsed(0);
    setFinishTimes({});
  }

  useEffect(() => () => cancelAnimationFrame(rafRef.current), []);
  useEffect(() => { reset(); }, [grade, classNum]); // eslint-disable-line react-hooks/exhaustive-deps

  async function markFinish(student) {
    if (!running) return;
    const t = Math.round(((Date.now() - startRef.current) / 1000) * 100) / 100;
    setFinishTimes(prev => ({ ...prev, [student.id]: t }));
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, t, activeYear, gradeAtMeasure);
  }

  async function editFinish(student, val) {
    const num = val === "" ? null : Number(val);
    if (num === null || Number.isNaN(num)) return;
    setFinishTimes(prev => ({ ...prev, [student.id]: num }));
    const gradeAtMeasure = inferSchoolGradeAtYear(student, activeYear);
    await onSave(student.id, eventId, num, activeYear, gradeAtMeasure);
  }

  function undoFinish(student) {
    setFinishTimes(prev => {
      const next = { ...prev };
      delete next[student.id];
      return next;
    });
  }

  const remaining = roster.filter(s => finishTimes[s.id] === undefined);
  const finished = roster.filter(s => finishTimes[s.id] !== undefined).sort((a, b) => finishTimes[a.id] - finishTimes[b.id]);

  return (
    <div className="class-run-timer">
      <div className="board-filters compact">
        <FilterChips label="학년" value={grade === null ? "" : String(grade)} onChange={v => { setGrade(Number(v)); setClassNum(null); }}
          options={gradeOptions.map(g => ({ id: String(g), label: g + "학년" }))} />
        {grade !== null && (
          <FilterChips label="반" value={classNum === null ? "" : String(classNum)} onChange={v => setClassNum(Number(v))}
            options={classOptions.map(c => ({ id: String(c), label: c + "반" }))} />
        )}
      </div>

      {roster.length === 0 ? (
        <div className="text-dim small-note">학년·반을 선택하면 그 반 학생 전체(최대 30명 안팎)가 아래 명단에 나타납니다.</div>
      ) : (
        <>
          <div className="class-run-sticky-header">
            <div className="class-run-clock">{elapsed.toFixed(2)}초</div>
            <div className="class-run-controls">
              {!running ? (
                <button className="btn btn-primary big-btn" onClick={start}><Play size={16} /> 전체 출발</button>
              ) : (
                <button className="btn btn-ghost big-btn" onClick={stopClock}><Square size={16} /> 시계 정지</button>
              )}
              <button className="btn btn-ghost" onClick={reset}><RotateCcw size={16} /> 초기화</button>
            </div>
            <div className="class-run-progress">완주 {finished.length} / {roster.length}명</div>
          </div>

          <div className="class-run-lists">
            <div className="class-run-col">
              <h4>달리는 중 ({remaining.length}명)</h4>
              <div className="class-run-grid">
                {remaining.map(s => (
                  <button key={s.id} className="class-run-btn" disabled={!running} onClick={() => markFinish(s)}>
                    {s.number}. {maskNames ? maskStudentName(s.name) : s.name}
                  </button>
                ))}
                {remaining.length === 0 && <div className="text-dim small-note">전원 완주했습니다.</div>}
              </div>
            </div>
            <div className="class-run-col">
              <h4>완주 ({finished.length}명)</h4>
              <div className="class-run-finished-list">
                {finished.map(s => (
                  <div className="class-run-finished-row" key={s.id}>
                    <span className="class-run-finished-name">{s.number}. {maskNames ? maskStudentName(s.name) : s.name}</span>
                    <input
                      className="input class-run-time-input"
                      type="number"
                      step="0.01"
                      value={finishTimes[s.id]}
                      onChange={e => editFinish(s, e.target.value)}
                    />
                    <button className="icon-btn" onClick={() => undoFinish(s)} title="완주 취소(다시 달리는 중으로)">
                      <RotateCcw size={13} />
                    </button>
                  </div>
                ))}
                {finished.length === 0 && <div className="text-dim small-note">아직 완주한 학생이 없습니다.</div>}
              </div>
            </div>
          </div>
          <div className="text-dim small-note">
            "전체 출발"을 누르면 반 전체가 동시에 뛰는 것으로 시계가 시작됩니다. 학생이 결승선을 통과하는
            순간 왼쪽 "달리는 중" 명단에서 그 학생 이름을 누르면 바로 시간이 기록되며 저장되고, 오른쪽
            "완주" 목록으로 옮겨갑니다. 잘못 눌렀다면 완주 목록에서 되돌리기 버튼으로 취소하거나
            시간을 직접 고칠 수 있습니다.
          </div>
        </>
      )}
    </div>
  );
}

/* ============================== 등급표 확인 ============================== */

function GradeTable({ students, activeYear, studentValue, studentGrade, settings, setSettings, isAdmin }) {
  const [mode, setMode] = useState("students"); // 'students' | 'reference'
  const [schoolGrade, setSchoolGrade] = useState("ALL");
  const [classNum, setClassNum] = useState("ALL");
  const [genderFilter, setGenderFilter] = useState("ALL"); // ALL | M | F
  const [schoolNameDraft, setSchoolNameDraft] = useState(settings.schoolName || "");
  const [hiddenAllByEvent, setHiddenAllByEvent] = useState({ bmi: true, bodyfat: true });
  const [revealedIdsByEvent, setRevealedIdsByEvent] = useState({ bmi: new Set(), bodyfat: new Set() });
  const schoolGrades = gradesForLevel(settings.schoolLevel || "middle");

  function toggleHiddenAll(eventId, checked) {
    setHiddenAllByEvent(prev => ({ ...prev, [eventId]: checked }));
    if (checked) setRevealedIdsByEvent(prev => ({ ...prev, [eventId]: new Set() })); // 다시 가릴 때는 개별로 풀어뒀던 것도 함께 초기화
  }

  function toggleStudentReveal(eventId, studentId) {
    setRevealedIdsByEvent(prev => {
      const next = new Set(prev[eventId] || []);
      if (next.has(studentId)) next.delete(studentId); else next.add(studentId);
      return { ...prev, [eventId]: next };
    });
  }


  useEffect(() => {
    setSchoolNameDraft(settings.schoolName || "");
  }, [settings.schoolName]);

  function saveSchoolName() {
    setSettings({ ...settings, schoolName: schoolNameDraft });
  }

  const visibleEvents = useMemo(() => {
    const skipped = computeEffectiveSkip(settings.skippedEvents);
    return ALL_EVENTS.filter(ev => NO_GRADE_EVENT_IDS.includes(ev.id) || !skipped.has(ev.id));
  }, [settings.skippedEvents]);

  const classOptions = useMemo(() => {
    const pool = schoolGrade === "ALL" ? students : students.filter(s => s.grade === Number(schoolGrade));
    return Array.from(new Set(pool.map(s => s.classNum))).sort((a, b) => a - b);
  }, [students, schoolGrade]);

  const genderCounts = useMemo(() => {
    const pool = students.filter(s => (schoolGrade === "ALL" || s.grade === Number(schoolGrade)) && (classNum === "ALL" || s.classNum === Number(classNum)));
    return {
      total: pool.length,
      M: pool.filter(s => s.gender === "M").length,
      F: pool.filter(s => s.gender === "F").length,
    };
  }, [students, schoolGrade, classNum]);

  const finalList = useMemo(() => {
    return students
      .filter(s => (schoolGrade === "ALL" || s.grade === Number(schoolGrade)) && (classNum === "ALL" || s.classNum === Number(classNum)) && (genderFilter === "ALL" || s.gender === genderFilter))
      .sort((a, b) => a.grade - b.grade || a.classNum - b.classNum || a.number - b.number);
  }, [students, schoolGrade, classNum, genderFilter]);

  return (
    <div>
      <div className="board-filters compact">
        <FilterChips label="보기" value={mode} onChange={setMode}
          options={[{ id: "students", label: "우리 학생 기록" }, { id: "reference", label: "학년별 참고기준표" }]} />
      </div>
      <div className="text-dim small-note">
        이 표의 <b>등급</b>은 종목별 기록을 교육부 공식 등급 기준표와 비교한 참고용 정보입니다. 나이스 제출은
        이 프로그램에 기록한 종목별 실측값을 그대로 사용하며, 등급 표기는 프로그램 내부 참고용일 뿐 제출·합불 판정과는 무관합니다.
      </div>

      {mode === "students" && isAdmin && (
        <div className="panel">
          <h3>학교 설정</h3>
          <div className="form-row">
            <label>학교명 (전광판 상단 표시, 선택)</label>
            <div className="pw-row">
              <input className="input" value={schoolNameDraft} onChange={e => setSchoolNameDraft(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") saveSchoolName(); }} placeholder="예: OO중학교" />
              <button
                className={"btn btn-primary school-save-btn" + (schoolNameDraft !== (settings.schoolName || "") ? " pending" : "")}
                disabled={schoolNameDraft === (settings.schoolName || "")}
                onClick={saveSchoolName}
              >
                <Check size={16} /> 저장
              </button>
            </div>
          </div>
        </div>
      )}

      {mode === "reference" ? (
        <ReferenceGradeTable defaultLevel={settings.schoolLevel || "middle"} />
      ) : students.length === 0 ? (
        <div className="panel empty-state">
          <ClipboardList size={40} color="var(--text-dim)" />
          <div>등록된 학생이 없습니다.</div>
        </div>
      ) : (
        <div className="panel">
          <h3>건강체력등급표 · 학생별 종목별 등급 확인 <span className="text-dim entry-year-tag">{activeYear}년</span></h3>
          <div className="board-filters compact">
            <FilterChips label="학년" value={schoolGrade} onChange={(v) => { setSchoolGrade(v); setClassNum("ALL"); }}
              options={[{ id: "ALL", label: "전체" }, ...schoolGrades.map(g => ({ id: String(g), label: g + "학년" }))]} />
            <FilterChips label="반" value={classNum} onChange={setClassNum} disabled={schoolGrade === "ALL"}
              options={[{ id: "ALL", label: "전체" }, ...classOptions.map(c => ({ id: String(c), label: c + "반" }))]} />
            <FilterChips label="성별" value={genderFilter} onChange={setGenderFilter}
              options={[
                { id: "ALL", label: "전체 (" + genderCounts.total + "명)" },
                { id: "M", label: "남 (" + genderCounts.M + "명)" },
                { id: "F", label: "여 (" + genderCounts.F + "명)" },
              ]} />
          </div>
          <div className="table-wrap">
            <table className="grade-table">
              <thead>
                <tr>
                  <th>학생</th>
                  {visibleEvents.map(ev => (
                    <th key={ev.id}>
                      {ev.name}
                      {NO_GRADE_EVENT_IDS.includes(ev.id) && (
                        <label className="bmi-hide-toggle">
                          <input type="checkbox" checked={hiddenAllByEvent[ev.id]} onChange={e => toggleHiddenAll(ev.id, e.target.checked)} />
                          가리기
                        </label>
                      )}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {finalList.map(s => {
                  return (
                    <tr key={s.id}>
                      <td className="student-cell">
                        <div className="student-cell-name">{s.name}</div>
                        <div className="student-cell-meta">{s.grade}-{s.classNum}-{s.number} · {s.gender === "M" ? "남" : "여"}</div>
                      </td>
                      {visibleEvents.map(ev => {
                        const v = studentValue(s.id, ev.id, activeYear);
                        if (NO_GRADE_EVENT_IDS.includes(ev.id)) {
                          const hiddenAll = hiddenAllByEvent[ev.id];
                          const isVisible = !hiddenAll || (revealedIdsByEvent[ev.id] || new Set()).has(s.id);
                          const labelText = ev.id === "bmi"
                            ? fmtValue(v, ev.unit) + (classifyBmi(v, s.grade, s.gender, settings.schoolLevel) ? " (" + classifyBmi(v, s.grade, s.gender, settings.schoolLevel) + ")" : "")
                            : ev.id === "bodyfat"
                            ? fmtValue(v, ev.unit) + (classifyBodyfat(v, s.gender) ? " (" + classifyBodyfat(v, s.gender) + ")" : "")
                            : fmtValue(v, ev.unit);
                          return (
                            <td key={ev.id}>
                              {v === null ? (
                                <span className="text-dim">미측정</span>
                              ) : (
                                <label className="bmi-cell">
                                  <input
                                    type="checkbox"
                                    checked={!isVisible}
                                    disabled={!hiddenAll}
                                    onChange={() => toggleStudentReveal(ev.id, s.id)}
                                  />
                                  <span className="cell-value">{isVisible ? labelText : "가려짐"}</span>
                                </label>
                              )}
                            </td>
                          );
                        }
                        const g = studentGrade(s, ev.id, activeYear);
                        return (
                          <td key={ev.id}>
                            {v === null ? (
                              <span className="text-dim">미측정</span>
                            ) : (
                              <div className="cell-grade">
                                <span className="grade-dot small" style={{ background: g ? GRADE_COLORS[g] : "#555" }}>{g || "-"}</span>
                                <span className="cell-value">{fmtValue(v, ev.unit)}</span>
                              </div>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
                {finalList.length === 0 && (
                  <tr><td colSpan={visibleEvents.length + 1} className="text-dim">해당 조건의 학생이 없습니다.</td></tr>
                )}
              </tbody>
            </table>
          </div>
          <div className="table-foot text-dim">등급은 종목별 참고 정보이며, BMI는 성장 확인용 참고 지표로 등급 산정에 포함되지 않습니다.</div>
        </div>
      )}
    </div>
  );
}

function ReferenceGradeTable({ defaultLevel }) {
  const levelId = defaultLevel || "middle";
  const level = SCHOOL_LEVELS.find(l => l.id === levelId);
  const [grade, setGrade] = useState(level.grades[0]);
  const [gender, setGender] = useState("M");
  const [eventId, setEventId] = useState(EVENTS[0].id);

  const ev = EVENT_MAP[eventId];
  const isBmi = eventId === "bmi";
  const isBodyfat = eventId === "bodyfat";
  const isOfficial = OFFICIAL_EVENT_IDS.includes(eventId);
  const bands = REFERENCE_BANDS[levelId]?.[eventId]?.[gender]?.[grade] || null;
  const notConducted = isOfficial && bands === null;
  const bmiBoundaries = isBmi ? BMI_OFFICIAL_CATEGORIES[levelId]?.[grade]?.[gender] : null;

  return (
    <div className="panel">
      <h3>참고 등급 기준표 <span className="board-cat">({level.label})</span></h3>
      {isBmi ? (
        <div className="warn-note official-note">
          <CheckCircle2 size={16} />
          <span>
            BMI는 교육부 학생건강정보센터 자료의 <b>마름/정상/과체중/경도비만/고도비만</b> 공식
            분류를 그대로 사용합니다. 다른 종목과 달리 등급(1~5등급)으로 나누지 않으며, 등급 산정에도
            포함되지 않는 성장 확인용 참고 지표입니다.
          </span>
        </div>
      ) : isBodyfat ? (
        <div className="warn-note official-note">
          <CheckCircle2 size={16} />
          <span>
            체지방률은 교육부 학생건강정보센터 자료의 <b>마름/정상/과체중/경도비만/고도비만</b> 공식
            분류를 그대로 사용합니다(전 학년 공통, 성별로만 구분). 다른 종목과 달리 등급(1~5등급)으로
            나누지 않으며, 등급 산정에도 포함되지 않는 성장 확인용 참고 지표입니다.
          </span>
        </div>
      ) : isOfficial ? (
        <div className="warn-note official-note">
          <CheckCircle2 size={16} />
          <span>
            이 종목은 <b>교육부 학생건강정보센터(PAPS) 공식 기준표</b>의 실제 수치입니다.
            연도·개정에 따라 바뀔 수 있으니 최신 공식 문서와 한 번씩 대조해 보시는 걸 권장합니다.
          </span>
        </div>
      ) : (
        <div className="warn-note">
          <AlertTriangle size={16} />
          <span>
            이 종목은 위 공식 문서에 포함되어 있지 않아, 중학교 기준값을 나이대에 맞춰 늘리거나 줄인
            <b> 참고용 예시</b>입니다. 실제 평가·학생 채점에는 학교 공식 매뉴얼 수치를 확인해 사용하세요.
          </span>
        </div>
      )}
      <div className="board-filters compact">
        <FilterChips label="학년" value={grade} onChange={setGrade}
          options={level.grades.map(g => ({ id: g, label: g + "학년" }))} />
        <FilterChips label="성별" value={gender} onChange={setGender}
          options={GENDERS.map(g => ({ id: g.id, label: g.label }))} />
        <div className="filter-group">
          <span className="filter-label">종목</span>
          <select className="select" value={eventId} onChange={e => setEventId(e.target.value)}>
            {ALL_EVENTS.map(e2 => <option key={e2.id} value={e2.id}>{e2.name}</option>)}
          </select>
        </div>
      </div>
      {isBmi ? (
        !bmiBoundaries ? (
          <div className="empty-state">
            <Info size={40} color="var(--text-dim)" />
            <div>자료가 없습니다.</div>
          </div>
        ) : (
          <div className="band-table">
            <div className="band-head-row">
              <span>분류</span><span>기준(BMI)</span><span></span>
            </div>
            {(() => {
              const [b1, b2, b3, b4] = bmiBoundaries;
              const rows = [
                { label: "마름", lo: "-", hi: b1 + " 이하" },
                { label: "정상", lo: (Math.round((b1 + 0.1) * 10) / 10), hi: b2 },
              ];
              if (b3 > b2) rows.push({ label: "과체중", lo: Math.round((b2 + 0.1) * 10) / 10, hi: b3 });
              rows.push({ label: "경도비만", lo: Math.round(((b3 > b2 ? b3 : b2) + 0.1) * 10) / 10, hi: b4 });
              rows.push({ label: "고도비만", lo: Math.round((b4 + 0.1) * 10) / 10, hi: "이상" });
              return rows.map(r => (
                <div className="band-row" key={r.label}>
                  <span className="grade-dot small bmi-cat-dot">{r.label}</span>
                  <span className="band-static">{r.lo}</span>
                  <span className="band-static">{r.hi}</span>
                </div>
              ));
            })()}
          </div>
        )
      ) : notConducted ? (
        <div className="empty-state">
          <Info size={40} color="var(--text-dim)" />
          <div>{level.label} {grade}학년은 이 종목을 실시하지 않습니다.</div>
        </div>
      ) : isBodyfat ? (
        <div className="band-table">
          <div className="band-head-row">
            <span>분류</span><span>기준(%)</span><span></span>
          </div>
          {(() => {
            const [b1, b2, b3, b4] = BODYFAT_OFFICIAL_CATEGORIES[gender];
            const rows = [
              { label: "마름", lo: "-", hi: b1 + " 이하" },
              { label: "정상", lo: Math.round((b1 + 0.1) * 10) / 10, hi: b2 },
              { label: "과체중", lo: Math.round((b2 + 0.1) * 10) / 10, hi: b3 },
              { label: "경도비만", lo: Math.round((b3 + 0.1) * 10) / 10, hi: b4 },
              { label: "고도비만", lo: Math.round((b4 + 0.1) * 10) / 10, hi: "이상" },
            ];
            return rows.map(r => (
              <div className="band-row" key={r.label}>
                <span className="grade-dot small bmi-cat-dot">{r.label}</span>
                <span className="band-static">{r.lo}</span>
                <span className="band-static">{r.hi}</span>
              </div>
            ));
          })()}
        </div>
      ) : (
        <div className="band-table">
          <div className="band-head-row">
            <span>등급</span><span>최소값{ev.unit ? " (" + ev.unit + ")" : ""}</span><span>최대값{ev.unit ? " (" + ev.unit + ")" : ""}</span>
          </div>
          {(bands || []).map(b => (
            <div className="band-row" key={b.grade}>
              <span className="grade-dot small" style={{ background: GRADE_COLORS[b.grade] }}>{b.grade}</span>
              <span className="band-static">{b.min}{eventId === "run_walk" && b.min > -9000 && b.min < 9000 ? " (" + fmtTime(b.min) + ")" : ""}</span>
              <span className="band-static">{b.max}{eventId === "run_walk" && b.max > -9000 && b.max < 9000 ? " (" + fmtTime(b.max) + ")" : ""}</span>
            </div>
          ))}
        </div>
      )}
      <div className="text-dim small-note">
        {level.label} {grade}학년 · {gender === "M" ? "남" : "여"} · {ev.name}{ev.unit ? " (" + ev.unit + ")" : ""} 기준표입니다.
      </div>
    </div>
  );
}

/* ============================== 학생 관리 ============================== */

// 성별 표기를 정규화한다. 엑셀/CSV에 "남"/"여", "M"/"F" 외에
// 숫자 코드(1=남, 2=여)로 입력된 경우도 함께 인식한다.
function normalizeGender(raw) {
  const s = String(raw ?? "").trim();
  if (s === "1") return "M";
  if (s === "2") return "F";
  if (s.startsWith("남") || s.toUpperCase().startsWith("M")) return "M";
  if (s.startsWith("여") || s.toUpperCase().startsWith("F")) return "F";
  return "F";
}

function parseRowsToStudents(rows) {
  if (!rows || rows.length === 0) return [];
  const HEADER_WORDS = { grade: ["학년"], classNum: ["반"], number: ["번호", "출석번호"], name: ["이름", "성명"], gender: ["성별"] };
  const headerCells = (rows[0] || []).map(c => String(c ?? "").trim());
  const detected = {};
  headerCells.forEach((cell, idx) => {
    Object.entries(HEADER_WORDS).forEach(([key, words]) => {
      if (words.some(w => cell.includes(w))) detected[key] = idx;
    });
  });
  const hasHeader = detected.grade !== undefined && detected.name !== undefined;
  const colMap = hasHeader ? detected : { grade: 0, classNum: 1, number: 2, name: 3, gender: 4 };
  const startIdx = hasHeader ? 1 : (isNaN(Number(headerCells[0])) ? 1 : 0);

  const result = [];
  for (let i = startIdx; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const nameRaw = row[colMap.name];
    if (nameRaw === undefined || String(nameRaw).trim() === "") continue;
    const grade = Number(row[colMap.grade]);
    const classNum = Number(row[colMap.classNum]);
    const number = Number(row[colMap.number]);
    if (!grade || !classNum || !number) continue;
    const gender = normalizeGender(row[colMap.gender]);
    result.push({ grade, classNum, number, name: String(nameRaw).trim(), gender });
  }
  return result;
}

// 학년+반+번호가 같으면 "같은 학생"으로 보고 새로 추가하는 대신 이름(과 성별)만 최신
// 값으로 갱신한다. 같은 엑셀 파일을 실수로 다시 올리거나(예: 이름이 안 보여서 재업로드해
// 봤더니 또 하나 늘어나는 문제), 다음 학년도에 번호가 그대로인 학생 명단을 다시 올릴 때
// 등, 학년·반·번호가 겹치는데도 매번 새 학생으로 중복 추가되는 것을 막기 위함이다.
function mergeParsedStudents(existingStudents, parsedRows) {
  const next = existingStudents.slice();
  const added = [];
  const updated = [];
  parsedRows.forEach(row => {
    const idx = next.findIndex(s => s.grade === row.grade && s.classNum === row.classNum && s.number === row.number);
    if (idx >= 0) {
      const prev = next[idx];
      if (prev.name !== row.name || prev.gender !== row.gender) {
        const merged = { ...prev, name: row.name, gender: row.gender };
        next[idx] = merged;
        updated.push(merged);
      }
    } else {
      const created = { id: uid("stu"), ...row };
      next.push(created);
      added.push(created);
    }
  });
  return { next, added, updated };
}
function describeBulkMergeResult(addedCount, updatedCount) {
  if (addedCount > 0 && updatedCount > 0) return `${addedCount}명 추가, ${updatedCount}명 정보를 갱신했습니다(같은 학년·반·번호의 기존 학생으로 인식).`;
  if (addedCount > 0) return `${addedCount}명을 일괄 추가했습니다.`;
  if (updatedCount > 0) return `${updatedCount}명의 정보를 갱신했습니다(같은 학년·반·번호의 기존 학생으로 인식해 새로 추가하지 않았습니다).`;
  return "변경 사항이 없습니다(이미 동일한 정보였습니다).";
}

function RosterManager({ students, setStudents, showToast, schoolGrades }) {
  const [form, setForm] = useState({ grade: 1, classNum: 1, number: "", name: "", gender: "M" });
  const [bulkText, setBulkText] = useState("");
  const [filterGrade, setFilterGrade] = useState("ALL");
  const [filterClass, setFilterClass] = useState("ALL");
  const fileInputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  const [selected, setSelected] = useState(() => new Set());
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const [lastBulkAddedIds, setLastBulkAddedIds] = useState(null);
  // 엑셀 파일을 연달아 여러 번 드래그하면, 앞선 파일을 아직 읽는 중(비동기)일 때 뒤이은
  // 드래그가 시작될 수 있다. 이때 각각의 처리 함수가 자신이 호출된 시점의(어쩌면 오래된)
  // students prop만 보고 "기존 목록 + 이번에 추가된 명단"을 계산하면, 나중에 끝난 쪽이
  // 앞서 추가된 학생들을 빼먹은 채로 덮어써 버릴 수 있다. 이를 막기 위해 매 렌더마다 최신
  // students 값을 ref에 동기화해두고, 일괄 추가 시에는 이 ref를 기준으로 이어붙인 뒤 ref도
  // 즉시 갱신해서 다음 호출이 항상 "지금까지 추가된 모든 학생"을 보게 한다.
  const studentsRef = useRef(students);
  studentsRef.current = students;

  function toggleSelect(id) {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  function toggleSelectAll(list) {
    setSelected(prev => {
      const allSelected = list.length > 0 && list.every(s => prev.has(s.id));
      const next = new Set(prev);
      if (allSelected) list.forEach(s => next.delete(s.id));
      else list.forEach(s => next.add(s.id));
      return next;
    });
  }

  function bulkSetGender(gender) {
    if (selected.size === 0) return;
    const next = students.map(s => (selected.has(s.id) ? { ...s, gender } : s));
    setStudents(next);
    showToast(selected.size + "명의 성별을 " + (gender === "M" ? "남" : "여") + "로 변경했습니다.", "ok");
    setSelected(new Set());
  }

  function bulkDelete() {
    if (selected.size === 0) return;
    setConfirmDeleteOpen(true);
  }

  function confirmBulkDelete() {
    const count = selected.size;
    setStudents(students.filter(s => !selected.has(s.id)));
    showToast(count + "명을 삭제했습니다.", "ok");
    setSelected(new Set());
    setConfirmDeleteOpen(false);
  }

  function addStudent() {
    if (!form.name || !form.number) { showToast("번호와 이름을 입력해 주세요.", "warn"); return; }
    const next = [...students, {
      id: uid("stu"), grade: Number(form.grade), classNum: Number(form.classNum),
      number: Number(form.number), name: form.name.trim(), gender: form.gender,
    }];
    setStudents(next);
    setForm(f => ({ ...f, number: "", name: "" }));
    showToast(form.name + " 학생을 추가했습니다.", "ok");
  }

  function removeStudent(id) {
    setStudents(students.filter(s => s.id !== id));
  }

  function undoLastBulkAdd() {
    if (!lastBulkAddedIds) return;
    const idSet = new Set(lastBulkAddedIds);
    const next = studentsRef.current.filter(s => !idSet.has(s.id));
    setStudents(next);
    studentsRef.current = next;
    setLastBulkAddedIds(null);
    showToast("방금 추가한 명단을 되돌렸습니다.", "ok");
  }

  function bulkAdd() {
    const lines = bulkText.split("\n").map(l => l.trim()).filter(Boolean);
    const parsed = [];
    lines.forEach(line => {
      const parts = line.split(/[,\t]+/).map(p => p.trim()).filter(Boolean);
      if (parts.length < 5) return;
      const [g, c, n, name, gender] = parts;
      const genderNorm = normalizeGender(gender);
      parsed.push({ grade: Number(g), classNum: Number(c), number: Number(n), name, gender: genderNorm });
    });
    if (parsed.length > 0) {
      // 학년+반+번호가 같은 기존 학생이 있으면 이름/성별만 갱신하고, 없을 때만 새로 추가한다
      // (같은 명단을 실수로 두 번 붙여넣어도 중복 학생이 생기지 않게).
      const { next, added, updated } = mergeParsedStudents(studentsRef.current, parsed);
      setStudents(next);
      studentsRef.current = next;
      setLastBulkAddedIds(added.length > 0 ? added.map(s => s.id) : null);
      setBulkText("");
      showToast(describeBulkMergeResult(added.length, updated.length), "ok");
    } else {
      showToast("형식을 확인해 주세요. 예) 1,3,12,홍길동,남", "warn");
    }
  }

  function readFileAsStudents(file) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = new Uint8Array(evt.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
          resolve({ name: file.name, added: parseRowsToStudents(rows), ok: true });
        } catch (err) {
          resolve({ name: file.name, added: [], ok: false });
        }
      };
      reader.onerror = () => resolve({ name: file.name, added: [], ok: false });
      reader.readAsArrayBuffer(file);
    });
  }

  async function processExcelFiles(fileList) {
    const files = Array.from(fileList || []);
    if (files.length === 0) return;
    const results = await Promise.all(files.map(readFileAsStudents));
    const allParsed = results.flatMap(r => r.added);
    const problemFiles = results.filter(r => !r.ok || r.added.length === 0).map(r => r.name);

    let added = [], updated = [];
    if (allParsed.length > 0) {
      // 항상 studentsRef(최신 값)를 기준으로 병합한다 — 파일을 읽는 동안(await) 다른
      // 드래그가 먼저 끝나 학생을 추가했을 수도 있으므로, 이 함수가 시작될 때 캡처했던
      // students 클로저가 아니라 지금 가장 최신인 목록을 기준으로 삼는다. 학년+반+번호가
      // 같은 기존 학생이 있으면 이름/성별만 갱신하고, 없을 때만 새로 추가한다 — 같은 엑셀
      // 파일을 실수로(또는 이름이 안 보여서) 두 번 올려도 중복 학생이 생기지 않고, 오히려
      // 그 재업로드로 누락된 이름을 다시 채워 넣을 수 있다.
      const merged = mergeParsedStudents(studentsRef.current, allParsed);
      added = merged.added;
      updated = merged.updated;
      setStudents(merged.next);
      studentsRef.current = merged.next;
      setLastBulkAddedIds(added.length > 0 ? added.map(s => s.id) : null);
    }
    if (allParsed.length > 0 && problemFiles.length === 0) {
      showToast(
        (files.length > 1 ? `파일 ${files.length}개 — ` : "") + describeBulkMergeResult(added.length, updated.length),
        "ok"
      );
    } else if (allParsed.length > 0 && problemFiles.length > 0) {
      showToast(`${describeBulkMergeResult(added.length, updated.length)} (인식 실패: ${problemFiles.join(", ")})`, "warn");
    } else {
      showToast("엑셀 내용을 인식하지 못했습니다. 학년·반·번호·이름·성별 열을 확인해 주세요.", "warn");
    }
  }

  function handleExcelFile(e) {
    processExcelFiles(e.target.files);
    e.target.value = "";
  }

  function handleDrop(e) {
    e.preventDefault();
    setDragOver(false);
    processExcelFiles(e.dataTransfer.files);
  }

  const filterClassOptions = useMemo(() => {
    const pool = filterGrade === "ALL" ? students : students.filter(s => s.grade === Number(filterGrade));
    return Array.from(new Set(pool.map(s => s.classNum))).sort((a, b) => a - b);
  }, [students, filterGrade]);

  const filtered = useMemo(() => {
    return students
      .filter(s => (filterGrade === "ALL" || s.grade === Number(filterGrade)) && (filterClass === "ALL" || s.classNum === Number(filterClass)))
      .sort((a, b) => a.grade - b.grade || a.classNum - b.classNum || a.number - b.number);
  }, [students, filterGrade, filterClass]);

  return (
    <div className="entry-layout">
      <div className="panel">
        <h3>엑셀 파일로 추가</h3>
        <div className="text-dim small-note">여러 학생을 한 번에 등록할 때(예: 새 학년도 전체 명단 등록) 활용하세요.</div>
        <div className="info-banner">
          <AlertTriangle size={16} />
          <span>
            <b>나이스 "학생명렬 내려받기"로 받은 엑셀 파일을 올려주세요.</b> 그래야 성별이 정확히 반영됩니다.
            성별 열이 없거나 인식되지 않는 파일을 올리면 학생 전원이 자동으로 "여"로 등록되니,
            올린 뒤에는 아래 명단에서 성별이 맞게 들어갔는지 꼭 확인해 주세요.
          </span>
        </div>
        <label
          htmlFor="roster-excel-file-input"
          className={"dropzone" + (dragOver ? " drag-over" : "")}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleDrop}
        >
          <FileSpreadsheet size={26} />
          <div className="dropzone-text">엑셀/CSV 파일을 여기로 끌어다 놓거나 눌러서 선택하세요 (여러 개 한 번에 가능)</div>
          <div className="text-dim small-note">열 순서: 학년, 반, 번호, 이름, 성별 (.xlsx, .xls, .csv · 제목 줄 자동 인식 · 성별은 남/여, M/F, 1(남)/2(여) 모두 인식)</div>
        </label>
        <input
          id="roster-excel-file-input"
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv,text/comma-separated-values,text/plain"
          multiple
          className="visually-hidden-input"
          onChange={handleExcelFile}
        />
        {lastBulkAddedIds && (
          <div className="bulk-undo-row">
            <span className="text-dim small-note">방금 {lastBulkAddedIds.length}명을 추가했습니다.</span>
            <button className="btn btn-ghost small" onClick={undoLastBulkAdd}><RotateCcw size={13} /> 되돌리기</button>
          </div>
        )}

        <div className="divider" />

        <h3>학생 추가</h3>
        <div className="text-dim small-note">새 학생이 들어왔을 때(예: 전학생) 한 명만 빠르게 추가하세요.</div>
        <div className="form-row">
          <label>학년</label>
          <div className="chip-row">
            {(schoolGrades || SCHOOL_GRADES).map(g => (
              <button key={g} className={"chip" + (form.grade === g ? " active" : "")} onClick={() => setForm(f => ({ ...f, grade: g }))}>{g}학년</button>
            ))}
          </div>
        </div>
        <div className="form-row-inline">
          <div>
            <label>반</label>
            <input className="input" type="number" min="1" value={form.classNum} onChange={e => setForm(f => ({ ...f, classNum: e.target.value }))} />
          </div>
          <div>
            <label>번호</label>
            <input className="input" type="number" min="1" value={form.number} onChange={e => setForm(f => ({ ...f, number: e.target.value }))} />
          </div>
          <div>
            <label>성별</label>
            <div className="chip-row">
              {GENDERS.map(g => (
                <button key={g.id} className={"chip" + (form.gender === g.id ? " active" : "")} onClick={() => setForm(f => ({ ...f, gender: g.id }))}>{g.label}</button>
              ))}
            </div>
          </div>
        </div>
        <div className="form-row">
          <label>이름</label>
          <input className="input" value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))}
            onKeyDown={e => { if (e.key === "Enter") addStudent(); }} />
        </div>
        <button className="btn btn-primary" onClick={addStudent}><Plus size={16} /> 추가</button>

        <div className="divider" />

        <h3>일괄 추가</h3>
        <div className="text-dim small-note">전학생 등 갑작스러운 신입생이 여러 명 한꺼번에 생겼을 때, 한 줄에 한 명씩 적어 빠르게 추가할 때 활용하세요.</div>
        <div className="text-dim small-note">한 줄에 한 명씩: 학년,반,번호,이름,성별 (예: 1,3,12,홍길동,남 · 성별은 남/여, M/F, 1(남)/2(여) 모두 인식)</div>
        <textarea className="textarea" rows={6} value={bulkText} onChange={e => setBulkText(e.target.value)}
          placeholder={"1,1,1,김민준,남\n1,1,2,이서연,여"} />
        <button className="btn btn-secondary" onClick={bulkAdd}>일괄 추가</button>
      </div>

      <div className="panel">
        <div className="panel-head-row">
          <h3>학생 명단 ({students.length}명)</h3>
        </div>
        <div className="board-filters compact">
          <FilterChips label="학년" value={filterGrade} onChange={(v) => { setFilterGrade(v); setFilterClass("ALL"); }}
            options={[{ id: "ALL", label: "전체" }, ...(schoolGrades || SCHOOL_GRADES).map(g => ({ id: String(g), label: g + "학년" }))]} />
          <FilterChips label="반" value={filterClass} onChange={setFilterClass} disabled={filterGrade === "ALL"}
            options={[{ id: "ALL", label: "전체" }, ...filterClassOptions.map(c => ({ id: String(c), label: c + "반" }))]} />
        </div>

        <div className="bulk-gender-bar">
          <label className="select-all-label">
            <input
              type="checkbox"
              checked={filtered.length > 0 && filtered.every(s => selected.has(s.id))}
              onChange={() => toggleSelectAll(filtered)}
            />
            전체선택
          </label>
          <span className="text-dim small-note" style={{ margin: 0 }}>{selected.size}명 선택됨</span>
          <div className="bulk-gender-actions">
            <button className="btn btn-ghost small" disabled={selected.size === 0} onClick={() => bulkSetGender("M")}>남학생으로 변경</button>
            <button className="btn btn-ghost small" disabled={selected.size === 0} onClick={() => bulkSetGender("F")}>여학생으로 변경</button>
            <button className="btn btn-ghost small danger-btn" disabled={selected.size === 0} onClick={bulkDelete}>
              <Trash2 size={13} /> 선택 삭제
            </button>
          </div>
        </div>

        <div className="roster-list">
          {filtered.length === 0 && <div className="text-dim">학생이 없습니다.</div>}
          {filtered.map(s => (
            <div className="roster-row" key={s.id}>
              <input type="checkbox" checked={selected.has(s.id)} onChange={() => toggleSelect(s.id)} />
              <span className="roster-meta">{s.grade}-{s.classNum}-{s.number}</span>
              <span className="roster-name">{s.name}</span>
              <div className="chip-row">
                {GENDERS.map(g => (
                  <button
                    key={g.id}
                    className={"chip small" + (s.gender === g.id ? " active" : "")}
                    onClick={() => setStudents(students.map(st => st.id === s.id ? { ...st, gender: g.id } : st))}
                  >
                    {g.label}
                  </button>
                ))}
              </div>
              <button className="icon-btn danger" onClick={() => removeStudent(s.id)}><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      </div>

      {confirmDeleteOpen && (
        <ConfirmModal
          title="선택한 학생 삭제"
          message={`${selected.size}명을 명단에서 삭제할까요? 되돌릴 수 없습니다.`}
          confirmLabel="삭제"
          danger
          onConfirm={confirmBulkDelete}
          onCancel={() => setConfirmDeleteOpen(false)}
        />
      )}
    </div>
  );
}

/* ============================== 등급 기준 설정 ============================== */

// 나이스 등 학교 시스템 엑셀 양식에 기록을 채워 넣기 위한 열 인식 사전.
// 더 구체적인 항목을 먼저 검사해야 헷갈리지 않는다(예: "왕복오래달리기"를 "오래달리기"보다 먼저 검사).
const NEIS_FIELD_RULES = [
  { field: "student_grade", label: "학년", test: h => /학년/.test(h) && !/학년도/.test(h) },
  { field: "student_class", label: "반", test: h => /반명|반코드|학급/.test(h) },
  { field: "student_number", label: "번호", test: h => /번호/.test(h) },
  { field: "student_name", label: "성명/이름", test: h => /성명|이름/.test(h) },
  { field: "shuttlerun", label: "왕복오래달리기", test: h => /왕복/.test(h) },
  { field: "step_test", label: "스텝검사", test: h => /스텝|PEI/i.test(h) },
  { field: "run_walk", label: "오래달리기-걷기", test: h => /오래달리기|걷기/.test(h) && !/왕복/.test(h) },
  { field: "flex_total", label: "종합유연성", test: h => /종합\s*유연성/.test(h) },
  { field: "sitreach", label: "앉아윗몸앞으로굽히기", test: h => /앉아|유연성/.test(h) && !/종합/.test(h) },
  { field: "gripstrength", label: "악력", test: h => /악력/.test(h) },
  { field: "situp", label: "윗몸말아올리기", test: h => /윗몸\s*말아올리기/.test(h) },
  { field: "pushup", label: "팔굽혀펴기", test: h => /팔굽혀펴기/.test(h) },
  { field: "longjump", label: "제자리멀리뛰기", test: h => /제자리|멀리뛰기/.test(h) },
  { field: "fifty_m", label: "50m달리기", test: h => /50\s*m|50\s*미터/i.test(h) },
  { field: "bodyfat", label: "체지방률", test: h => /체지방/.test(h) },
  // "BMI"/"체질량지수" 열엔 계산된 BMI 수치를, "신장"/"체중" 열엔 기록입력 화면에서
  // 입력한 신장·체중 실측값을 각각 그대로 채운다(둘은 서로 다른 열).
  { field: "bmi_value", label: "BMI", test: h => /BMI/i.test(h) || /체질량\s*지수/.test(h) },
  { field: "bmi_height", label: "신장", test: h => /신장|키\(/.test(h) || /^키$/.test(h) },
  { field: "bmi_weight", label: "체중", test: h => /체중|몸무게/.test(h) },
];
function guessNeisField(header) {
  const h = String(header || "");
  const hasTrial2 = /2차/.test(h);
  const hasLeft = /왼쪽|좌/.test(h);
  for (const rule of NEIS_FIELD_RULES) {
    if (rule.test(h)) {
      if (["sitreach", "longjump"].includes(rule.field)) return rule.field + (hasTrial2 ? "_2" : "_1");
      if (rule.field === "gripstrength") return "gripstrength_" + (hasTrial2 ? "2" : "1") + "_" + (hasLeft ? "left" : "right");
      return rule.field;
    }
  }
  return "";
}

// 학생 1명에 대해, 인식된 필드 하나의 실제 값을 구한다.
function neisFieldValue(student, field, records, activeYear) {
  const v = x => (x === undefined || x === null ? "" : String(x));
  if (field === "student_grade") return v(student.grade);
  if (field === "student_class") return v(student.classNum);
  if (field === "student_number") return v(student.number);
  if (field === "student_name") return student.name;
  if (field === "bmi_value") return v(records[recKey(student.id, "bmi", activeYear)]?.value);
  if (field === "bmi_height") return v(records[recKey(student.id, "bmi", activeYear)]?.parts?.height);
  if (field === "bmi_weight") return v(records[recKey(student.id, "bmi", activeYear)]?.parts?.weight);
  if (field.startsWith("sitreach_")) return v(records[recKey(student.id, "sitreach", activeYear)]?.parts?.["trial" + field.slice(-1)]);
  if (field.startsWith("longjump_")) return v(records[recKey(student.id, "longjump", activeYear)]?.parts?.["trial" + field.slice(-1)]);
  if (field.startsWith("gripstrength_")) {
    const [, trial, side] = field.split("_"); // gripstrength_1_right 형태
    return v(records[recKey(student.id, "gripstrength", activeYear)]?.parts?.[side + trial]);
  }
  // 나머지는 값 하나짜리 종목(왕복오래달리기, 윗몸말아올리기, 팔굽혀펴기, 50m달리기, 스텝검사,
  // 오래달리기-걷기, 종합유연성, 체지방률 등)이다.
  return v(records[recKey(student.id, field, activeYear)]?.value);
}

function NeisTemplateFiller({ students, records, activeYear, showToast, fillState, setFillState }) {
  const fileInputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false); // 드래그 중인지 여부는 탭을 오가며 유지할 필요가 없는 순간 UI 상태라 그대로 로컬로 둔다.
  const [showDeleteReminder, setShowDeleteReminder] = useState(false); // 다운로드 직후 파기 안내 팝업
  // 첨부파일명·인식된 헤더·반영결과 등은 상위(PapsApp)에서 내려주는 fillState에 둬서,
  // 데이터백업 탭을 벗어났다가 다시 돌아와도 사라지지 않는다.
  const { fileName, headerRow, dataRows, mapping, resultRows, sheetName } = fillState;
  const setFileName = v => setFillState(prev => ({ ...prev, fileName: v }));
  const setHeaderRow = v => setFillState(prev => ({ ...prev, headerRow: v }));
  const setDataRows = v => setFillState(prev => ({ ...prev, dataRows: v }));
  const setMapping = v => setFillState(prev => ({ ...prev, mapping: v }));
  const setResultRows = v => setFillState(prev => ({ ...prev, resultRows: v }));
  const setSheetName = v => setFillState(prev => ({ ...prev, sheetName: v }));

  function processFile(file) {
    if (!file) return;
    setFileName(file.name);
    setResultRows(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const wsName = wb.SheetNames[0];
        setSheetName(wsName);
        const ws = wb.Sheets[wsName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (!rows.length) {
          showToast("빈 파일이거나 읽을 수 없는 파일입니다.", "warn");
          return;
        }
        const headers = rows[0].map(h => String(h || "").trim());
        setHeaderRow(headers);
        setDataRows(rows.slice(1).filter(r => r.some(c => String(c || "").trim() !== "")));
        setMapping(headers.map(h => guessNeisField(h)));
      } catch (err) {
        showToast("파일을 읽는 중 문제가 발생했습니다. 엑셀(.xlsx) 파일인지 확인해 주세요.", "warn");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleFile(e) {
    processFile(e.target.files?.[0]);
  }
  function handleDrop(e) {
    e.preventDefault();
    setDragOver(false);
    processFile(e.dataTransfer.files?.[0]);
  }

  function applyFill() {
    const gradeCol = mapping.indexOf("student_grade");
    const classCol = mapping.indexOf("student_class");
    const numberCol = mapping.indexOf("student_number");
    const nameCol = mapping.indexOf("student_name");

    // 템플릿에 이미 학생 행이 있으면 그 순서를 그대로 따르고, 학년/반/번호(있으면 이름까지)로
    // 우리 학생 명단과 매칭한다. 템플릿이 비어 있으면(헤더만 있으면) 우리 학생 명단 순서로 새로 만든다.
    const templateHasRows = dataRows && dataRows.length > 0;
    const baseList = templateHasRows
      ? dataRows.map(r => {
          const g = gradeCol >= 0 ? Number(r[gradeCol]) : null;
          const c = classCol >= 0 ? Number(r[classCol]) : null;
          const n = numberCol >= 0 ? Number(r[numberCol]) : null;
          const nm = nameCol >= 0 ? String(r[nameCol] || "").trim() : null;
          let match = null;
          if (g && c && n) match = students.find(s => s.grade === g && s.classNum === c && s.number === n);
          if (!match && nm) match = students.find(s => s.name === nm && (!g || s.grade === g));
          return { student: match, templateRow: r };
        })
      : [...students].sort((a, b) => a.grade - b.grade || a.classNum - b.classNum || a.number - b.number).map(s => ({ student: s, templateRow: null }));

    const filled = baseList.map(({ student, templateRow }) => {
      return headerRow.map((h, idx) => {
        const field = mapping[idx];
        if (!student) return templateRow ? (templateRow[idx] ?? "") : "";
        if (!field) return templateRow ? (templateRow[idx] ?? "") : "";
        return neisFieldValue(student, field, records, activeYear);
      });
    });

    const unmatchedCount = baseList.filter(b => !b.student).length;
    if (unmatchedCount > 0) {
      showToast(templateHasRows
        ? unmatchedCount + "개 행은 우리 학생 명단과 매칭되지 않아 원본 값을 그대로 두었습니다."
        : "일부 학생 매칭에 실패했습니다.", "warn");
    }
    setResultRows(filled);
    showToast("반영이 끝났습니다. 아래에서 확인 후 다운로드해 주세요.", "ok");
  }

  function downloadResult() {
    if (!resultRows) return;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headerRow, ...resultRows]);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    XLSX.writeFile(wb, (fileName.replace(/\.xlsx?$/i, "") || "나이스양식") + "_반영_" + dateStr + ".xlsx");
    showToast("파일을 다운로드했습니다.", "ok");
    setShowDeleteReminder(true);
  }

  return (
    <div className="panel">
      <h3>학교 시스템(나이스) 양식에 직접 반영하기</h3>
      <div className="text-dim small-note">
        나이스에서 받은 진짜 양식 파일을 여기에 첨부한 뒤 "반영하기"를 누르면, 이 프로그램의
        기록을 열 이름에 맞춰 자동으로 채워줍니다. "신장"·"체중" 열은 기록입력 화면에서
        입력한 신장·체중 실측값이, "BMI" 열은 그 값으로 계산된 BMI 수치가 채워집니다.
        <b> 실제로 제출하시기 전에 아래 미리보기에서 값이 정확히 채워졌는지 꼭 확인해 주세요.</b>
      </div>

      <input id="neis-template-input" ref={fileInputRef} type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" className="visually-hidden-input" onChange={handleFile} />
      <label
        htmlFor="neis-template-input"
        className={"dropzone" + (dragOver ? " drag-over" : "")}
        onDragOver={e => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={handleDrop}
      >
        <FileSpreadsheet size={22} />
        <div className="dropzone-text">나이스 양식 파일을 여기로 끌어다 놓거나 눌러서 선택하세요</div>
      </label>
      {fileName && <div className="text-dim small-note">첨부한 파일: {fileName}{dataRows && dataRows.length > 0 ? ` (기존 학생 행 ${dataRows.length}개 발견)` : " (빈 양식으로 인식, 우리 학생 명단 순서로 새로 채웁니다)"}</div>}

      {headerRow && (
        <>
          <div className="divider" />
          <button className="btn btn-primary" onClick={applyFill}><Check size={14} /> 반영하기</button>

          {resultRows && (
            <>
              <div className="divider" />
              <div className="text-dim small-note">
                아래는 반영된 결과의 앞부분 미리보기입니다. 실제로 제출하시기 전에 반드시 값을 확인해 주세요.
              </div>
              <div className="table-wrap neis-preview-wrap">
                <table className="grade-table">
                  <thead><tr>{headerRow.map((h, i) => <th key={i}>{h}</th>)}</tr></thead>
                  <tbody>
                    {resultRows.slice(0, 5).map((row, ri) => (
                      <tr key={ri}>{row.map((c, ci) => <td key={ci}>{String(c)}</td>)}</tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <button className="btn btn-secondary" onClick={downloadResult}><Copy size={14} /> 반영된 파일 다운로드</button>
            </>
          )}
        </>
      )}
      {showDeleteReminder && <DownloadDeleteReminderModal onClose={() => setShowDeleteReminder(false)} />}
    </div>
  );
}

// 학생 개인정보(이름·기록 등)가 담긴 파일을 방금 내려받은 직후 한 번 보여주는 파기 안내.
// 나이스 반영, 백업(JSON), 마감 전 백업 세 곳에서 공통으로 쓴다. 앱 내부 데이터는 "마감"으로
// 지울 수 있어도, 한 번 내려받아 교사 PC 다운로드 폴더에 남은 파일은 이 프로그램이 지울 수
// 없어 별도의 유출 통로가 될 수 있기 때문에, 다운로드 시점마다 삭제를 상기시킨다.
function DownloadDeleteReminderModal({ onClose }) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><ShieldCheck size={18} /> 다운로드한 파일 삭제 안내</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <div className="text-dim small-note">
            방금 내려받은 파일에는 학생 이름 등 개인정보가 들어 있습니다. 나이스 등록(또는
            필요한 처리)을 마쳤다면, 컴퓨터에 남겨두지 말고 <b>다운로드 폴더에서 바로
            삭제(가능하면 휴지통을 거치지 않는 Shift+Delete로 완전 삭제)</b>해 주세요. 특히
            여러 사람이 함께 쓰는 컴퓨터라면 다운로드 폴더에 방치된 파일이 개인정보 유출
            통로가 될 수 있습니다.
          </div>
          <div className="confirm-actions">
            <button className="btn btn-primary" onClick={onClose}>확인</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function DataBackupPanel({ students, records, criteria, settings, activeYear, onImportBackup, showToast, workspaceCode, myDisplayName, isFounder, neisFillState, setNeisFillState }) {
  const [pendingImport, setPendingImport] = useState(null);
  const importInputRef = useRef(null);
  const [backupLog, setBackupLog] = useState(null);
  const [showDeleteReminder, setShowDeleteReminder] = useState(false); // 다운로드 직후 파기 안내 팝업

  useEffect(() => {
    loadBackupLog(workspaceCode).then(list => setBackupLog(list.slice().reverse()));
  }, [workspaceCode]);

  function exportBackup() {
    const payload = { exportedAt: Date.now(), students, criteria, settings, records };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const dateStr = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = "paps-backup-" + (settings.schoolName || "data") + "-" + dateStr + ".json";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showToast("백업 파일을 다운로드했습니다.", "ok");
    const entry = { ts: Date.now(), by: myDisplayName };
    appendBackupLog(workspaceCode, entry);
    setBackupLog(prev => [entry, ...(prev || [])]);
    setShowDeleteReminder(true);
  }

  function handleImportFile(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const payload = JSON.parse(evt.target.result);
        if (!payload || !Array.isArray(payload.students)) {
          showToast("올바른 백업 파일이 아닙니다.", "warn");
          return;
        }
        setPendingImport(payload);
      } catch (err) {
        showToast("파일을 읽는 중 오류가 발생했습니다.", "warn");
      }
    };
    reader.readAsText(file);
  }

  function confirmImport() {
    onImportBackup(pendingImport);
    setPendingImport(null);
  }

  return (
    <div className="entry-layout">
      <div className="panel">
        <h3>데이터 백업이란?</h3>
        <div className="text-dim small-note">
          이 프로그램의 기록은 평소 자동으로 저장됩니다. 혹시 모를 상황에 대비해 파일로
          따로 저장해두고 싶을 때 쓰는 기능입니다.
        </div>
      </div>

      <div className="panel">
        <h3>백업 파일(JSON)</h3>
        {!isFounder ? (
          <div className="text-dim small-note">
            JSON 백업(내보내기·불러오기)은 개설자만 할 수 있습니다. 여러 명이 각자 따로
            백업하고 불러오면 서로 다른 시점의 기록이 뒤섞여 최신 데이터가 덮어써질 수
            있어서, 혼선을 막기 위해 개설자 한 명으로만 창구를 좁혀두었습니다.
          </div>
        ) : (
        <div className="backup-steps">
          <div className="backup-step">
            <span className="backup-step-num">1</span>
            <div>
              <div className="backup-step-title">저장하기</div>
              <div className="text-dim small-note">아래 버튼을 누르면 파일이 저장됩니다.</div>
              <button className="btn btn-secondary" onClick={exportBackup}><Copy size={14} /> JSON으로 내보내기 (교사 보관용, 실명 포함)</button>
              {backupLog && (
                backupLog.length === 0 ? (
                  <div className="text-dim small-note backup-log-empty">아직 내보낸 기록이 없습니다.</div>
                ) : (
                  <div className="backup-log-latest">
                    가장 최근 내보내기: <b>{new Date(backupLog[0].ts).toLocaleString("ko-KR")}</b> ({backupLog[0].by})
                  </div>
                )
              )}
              {backupLog && backupLog.length > 1 && (
                <details className="backup-log-more">
                  <summary>이전 내보내기 이력 더 보기 ({backupLog.length - 1}건)</summary>
                  <div className="backup-log-list">
                    {backupLog.slice(1).map((b, i) => (
                      <div className="backup-log-row" key={i}>{new Date(b.ts).toLocaleString("ko-KR")} — {b.by}</div>
                    ))}
                  </div>
                </details>
              )}
            </div>
          </div>
          <div className="backup-step">
            <span className="backup-step-num">2</span>
            <div>
              <div className="backup-step-title">나중에 다시 불러오기</div>
              <div className="text-dim small-note">저장해둔 파일을 선택하면, 지금 화면의 기록이 그 파일 내용으로 <b>완전히 바뀝니다</b>(되돌릴 수 없어요).</div>
              <input id="json-restore-file-input" ref={importInputRef} type="file" accept=".json,application/json" className="visually-hidden-input" onChange={handleImportFile} />
              <label htmlFor="json-restore-file-input" className="btn btn-ghost file-label-btn">JSON 백업 파일 불러오기</label>
            </div>
          </div>
        </div>
        )}
        {isFounder && (
        <div className="warn-note">
          <AlertTriangle size={16} />
          <span>학생 이름·기록이 담긴 파일입니다. 개인 기기 등 안전한 곳에만 보관하고, 다 쓴 옛 파일은 삭제해 주세요.</span>
        </div>
        )}
      </div>

      <NeisTemplateFiller students={students} records={records} activeYear={activeYear} showToast={showToast} fillState={neisFillState} setFillState={setNeisFillState} />

      {showDeleteReminder && <DownloadDeleteReminderModal onClose={() => setShowDeleteReminder(false)} />}

      {pendingImport && isFounder && (
        <ConfirmModal
          title="백업 파일 불러오기"
          message={`이 파일로 불러오면 현재 학생 ${students.length}명의 데이터가 백업 파일 속 학생 ${pendingImport.students.length}명 데이터로 완전히 대체됩니다. 계속할까요?`}
          confirmLabel="불러오기"
          danger
          onConfirm={confirmImport}
          onCancel={() => setPendingImport(null)}
        />
      )}
    </div>
  );
}

/* ============================== 학기 마감 ============================== */

// 파일명에 다운로드 시점(날짜+시각)이 바로 보이도록 "YYYYMMDD_HHmm" 형태로 만든다. 같은 날
// 여러 번 백업을 받아도 파일명만 보고 어느 게 최신인지 구분할 수 있게 하기 위함.
function fileTimestamp() {
  const now = new Date();
  const pad = n => String(n).padStart(2, "0");
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
}

// 마감 전, JSON 백업과는 별도로 "혹시 나중에 필요할 수도 있는" 나이스 제출양식 형태의 엑셀로
// 전체 구성원의 모든 종목 기록을 내려받을 수 있게 한다. 특정 학교의 실제 나이스 업로드
// 양식(열 구성)은 학교/연도마다 다를 수 있어 그 양식에 정확히 맞추기보다는, 나이스가 흔히
// 요구하는 열 이름(guessNeisField가 인식하는 헤더와 같은 표기)으로 모든 종목의 세부 측정값을
// 빠짐없이 한 장에 담는 것을 목표로 한다.
const CLOSEOUT_NEIS_EXPORT_COLUMNS = [
  { header: "학년", field: "student_grade" },
  { header: "반", field: "student_class" },
  { header: "번호", field: "student_number" },
  { header: "성명", field: "student_name" },
  { header: "왕복오래달리기(회)", field: "shuttlerun" },
  { header: "오래달리기-걷기(초)", field: "run_walk" },
  { header: "스텝검사(PEI)", field: "step_test" },
  { header: "윗몸말아올리기(회)", field: "situp" },
  { header: "팔굽혀펴기(회)", field: "pushup" },
  { header: "앉아윗몸앞으로굽히기 1차(cm)", field: "sitreach_1" },
  { header: "앉아윗몸앞으로굽히기 2차(cm)", field: "sitreach_2" },
  { header: "종합유연성(점)", field: "flex_total" },
  { header: "제자리멀리뛰기 1차(cm)", field: "longjump_1" },
  { header: "제자리멀리뛰기 2차(cm)", field: "longjump_2" },
  { header: "50m달리기(초)", field: "fifty_m" },
  { header: "악력 1차 왼쪽(kg)", field: "gripstrength_1_left" },
  { header: "악력 1차 오른쪽(kg)", field: "gripstrength_1_right" },
  { header: "악력 2차 왼쪽(kg)", field: "gripstrength_2_left" },
  { header: "악력 2차 오른쪽(kg)", field: "gripstrength_2_right" },
  { header: "신장(cm)", field: "bmi_height" },
  { header: "체중(kg)", field: "bmi_weight" },
  { header: "BMI", field: "bmi_value" },
  { header: "체지방률(%)", field: "bodyfat" },
];

function SemesterCloseoutPanel({ students, records, criteria, settings, onCloseout, workspaceCode, activeYear }) {
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [backedUp, setBackedUp] = useState(false);
  const [showDeleteReminder, setShowDeleteReminder] = useState(false); // 다운로드 직후 파기 안내 팝업

  const recordCount = Object.keys(records || {}).length;
  const studentCount = students.length;
  const isEmpty = studentCount === 0 && recordCount === 0;

  function downloadBackupNow() {
    const payload = { exportedAt: Date.now(), students, criteria, settings, records };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const dateStr = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = "paps-backup-" + (settings.schoolName || "data") + "-마감전-" + dateStr + ".json";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setBackedUp(true);
    setShowDeleteReminder(true);
  }

  // 위 JSON 백업(프로그램이 스스로 복원하는 용도)과 별개로, 사람이 열어보거나 나이스에 참고할
  // 수 있는 엑셀 형태 전체 백업. 마감 필수 조건(backedUp)에는 영향을 주지 않는 선택 사항이다.
  function downloadNeisFullExcelNow() {
    const sorted = [...students].sort((a, b) => a.grade - b.grade || a.classNum - b.classNum || a.number - b.number);
    const header = CLOSEOUT_NEIS_EXPORT_COLUMNS.map(c => c.header);
    const rows = sorted.map(s => CLOSEOUT_NEIS_EXPORT_COLUMNS.map(c => neisFieldValue(s, c.field, records, activeYear)));
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws["!cols"] = header.map(() => ({ wch: 14 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "전체기록");
    XLSX.writeFile(wb, (settings.schoolName || "학교") + "_나이스제출양식_전체기록_" + fileTimestamp() + ".xlsx");
    setShowDeleteReminder(true);
  }

  return (
    <div>
      {showDeleteReminder && <DownloadDeleteReminderModal onClose={() => setShowDeleteReminder(false)} />}
      <div className="panel closeout-hero">
        <Trash2 size={36} color="#E85D5D" />
        <h2 className="closeout-title">마감</h2>
        <p className="closeout-lead">
          나이스 제출 등 이번 학기 측정 목적을 모두 마쳤다면, 학생 개인정보(체력 기록)를 계속
          저장해 둘 필요가 없습니다. 이 화면에서 <b>기록을 한 번에 정리</b>해 개인정보 노출
          기간을 줄일 수 있습니다.
        </p>
      </div>

      <div className="panel">
        <h3>마감 전 꼭 확인하세요</h3>
        <ul className="closeout-checklist">
          <li><b>나이스(NEIS) 제출을 마쳤는지</b> 다시 한 번 확인해 주세요 — 데이터 백업 탭에서 나이스 양식에 기록을 반영하셨나요?</li>
          <li><b>이 작업은 되돌릴 수 없습니다.</b> 마감하면 기록·학생 명단뿐 아니라 <b>이 학교 코드 자체(설정, 접근권한, 변경 이력 등 전부)</b>가 삭제되고, 화면이 첫 화면으로 돌아갑니다.</li>
        </ul>

        <div className="closeout-backup-step">
          {isEmpty ? (
            <div className="text-dim small-note">
              현재 학생 명단과 기록이 없어서, 백업 없이 바로 마감할 수 있습니다.
            </div>
          ) : (
            <>
              <div className="backup-step-title">① 백업 파일부터 받으세요 (필수)</div>
              <div className="text-dim small-note">
                아래 버튼으로 지금 상태를 백업해두지 않으면 "마감하기" 버튼이 눌리지 않습니다.
                만약을 대비한 최소한의 안전장치입니다.
              </div>
              <button className="btn btn-secondary" onClick={downloadBackupNow}>
                <Copy size={14} /> {backedUp ? "백업 파일 다시 받기" : "지금 백업 파일 받기"}
              </button>
              {backedUp && <div className="closeout-backup-done"><CheckCircle2 size={14} /> 백업을 받았습니다. 이제 마감할 수 있습니다.</div>}

              <div className="text-dim small-note closeout-neis-export-hint">
                (선택) 위 백업과는 별도로, 전체 학생의 모든 측정 기록을 나이스 제출양식과 비슷한
                형태의 엑셀 파일로도 받아둘 수 있습니다. 마감에 필수는 아니며, 나중에 참고가
                필요할 때를 대비한 것입니다.
              </div>
              <button className="btn btn-ghost" onClick={downloadNeisFullExcelNow}>
                <FileSpreadsheet size={14} /> 나이스 제출양식 엑셀로 전체 기록 받기
              </button>
            </>
          )}
        </div>

        <div className="closeout-summary">
          현재 <b>학생 {studentCount}명</b>, <b>기록 {recordCount}건</b>이 저장되어 있습니다.
          마감하면 <b>학생 명단과 기록이 모두 삭제</b>됩니다.
        </div>
        <div className="closeout-step-title">{isEmpty ? "마감하기" : "② 마감하기"}</div>
        <button className="btn btn-primary closeout-btn" onClick={() => setConfirmOpen(true)} disabled={!isEmpty && !backedUp}>
          <Trash2 size={16} /> 마감하기
        </button>
        {!isEmpty && !backedUp && <div className="text-dim small-note">먼저 위에서 백업 파일을 받아야 눌러집니다.</div>}

        <div className="closeout-footer">
          <p className="closeout-footer-text">
            이 프로그램은 수시로 여유가 생길 때마다 업데이트할 예정입니다.<br />
            PAPS 측정 업무 간 많은 도움 되셨길 바랍니다.
          </p>
          <button className="btn btn-ghost" onClick={() => window.open(INQUIRY_FORM_URL, "_blank", "noopener,noreferrer")}>
            <Info size={14} /> 문의하기
          </button>
        </div>
      </div>

      {confirmOpen && (
        <SemesterCloseoutConfirm
          studentCount={studentCount}
          recordCount={recordCount}
          onConfirm={() => { onCloseout(); setConfirmOpen(false); }}
          onCancel={() => setConfirmOpen(false)}
        />
      )}
    </div>
  );
}

function SemesterCloseoutConfirm({ studentCount, recordCount, onConfirm, onCancel }) {
  const [understood, setUnderstood] = useState(false);
  return (
    <div className="modal-backdrop" onClick={onCancel}>
      <div className="modal-panel confirm-panel" onClick={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3><AlertTriangle size={18} color="#E85D5D" /> 정말 마감하시겠습니까?</h3>
          <button className="icon-btn" onClick={onCancel}><X size={16} /></button>
        </div>
        <div className="modal-body">
          <div className="closeout-confirm-big">기록 {recordCount}건이 영구히 삭제됩니다</div>
          <div className="closeout-confirm-big">학생 명단 {studentCount}명도 함께 삭제됩니다</div>
          <div className="closeout-confirm-big">이 학교 코드 자체도 삭제되어, 첫 화면으로 돌아갑니다</div>
          <div className="closeout-confirm-warn">이 작업은 되돌릴 수 없습니다.</div>
          <label className="closeout-option closeout-confirm-check">
            <input type="checkbox" checked={understood} onChange={e => setUnderstood(e.target.checked)} />
            <span>백업을 마쳤으며, 삭제 후 복구할 수 없다는 것을 이해했습니다.</span>
          </label>
          <div className="confirm-actions">
            <button className="btn btn-ghost" onClick={onCancel}>취소</button>
            <button className="btn btn-primary danger-confirm-btn" disabled={!understood} onClick={onConfirm}>네, 마감합니다</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== 타이머 & 셔틀런 음원 ============================== */

function beepAt(ctx, time, freq, dur, gain) {
  const osc = ctx.createOscillator();
  const g = ctx.createGain();
  osc.type = "sine";
  osc.frequency.value = freq;
  g.gain.value = gain;
  osc.connect(g);
  g.connect(ctx.destination);
  osc.start(time);
  osc.stop(time + dur);
}

// 브라우저의 자동재생 제한 때문에 새로 만든 AudioContext가 "suspended" 상태로
// 멈춰 있는 경우가 흔하다(특히 이 화면처럼 iframe 안에서 실행될 때). 이 상태에서는
// 오류 없이 그냥 아무 소리도 나지 않으므로, 실제로 재생 중(running) 상태가 된 뒤에
// 신호음을 예약해야 한다. 또한 정지된 동안에는 ctx.currentTime이 멈춰 있어, 재개
// "전"에 계산해둔 시각으로 예약하면 타이밍이 어긋날 수 있어 반드시 재개 이후에
// 시각을 계산해서 예약해야 한다.
function ensureRunningContext(ctx, onReady) {
  let done = false;
  const proceed = () => { if (done) return; done = true; onReady(); };
  if (ctx.state === "running") { proceed(); return; }
  ctx.resume().then(proceed).catch(proceed);
  // 일부 환경에서는 resume()의 Promise가 끝내 응답하지 않는 경우가 있어,
  // 최후의 안전장치로 300ms 뒤에도 강제로 한 번 더 진행시킨다.
  setTimeout(() => { if (ctx.state !== "closed") proceed(); }, 300);
}

function ShuttleRunPlayer({ settings, setSettings }) {
  const { startSpeed, increment, distance, levelDuration, maxLevels } = SHUTTLE_PRESET;
  const [playing, setPlaying] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const [level, setLevel] = useState(0);
  const [lap, setLap] = useState(0);
  const [totalLaps, setTotalLaps] = useState(0);

  const ctxRef = useRef(null);
  const scheduleRef = useRef([]);
  const startTimeRef = useRef(0);
  const rafRef = useRef(null);

  const shuttleSource = settings.shuttleSource || "youtube";
  const youtubeUrl = settings.shuttleYoutubeUrl || "";
  const youtubeId = extractYoutubeId(youtubeUrl);

  function setShuttleSource(src) {
    setSettings({ ...settings, shuttleSource: src });
  }
  function setYoutubeUrl(url) {
    setSettings({ ...settings, shuttleYoutubeUrl: url });
  }

  const currentSpeed = level > 0 ? Math.round((startSpeed + increment * (level - 1)) * 10) / 10 : startSpeed;

  function tick() {
    const ctx = ctxRef.current;
    if (!ctx) return;
    const rel = ctx.currentTime - startTimeRef.current;
    setElapsed(Math.max(0, rel));
    let idx = 0;
    while (idx < scheduleRef.current.length && scheduleRef.current[idx].time <= rel) idx++;
    const cur = scheduleRef.current[idx - 1];
    if (cur) {
      setLevel(cur.level);
      setLap(cur.lapInLevel);
      setTotalLaps(idx);
    }
    rafRef.current = requestAnimationFrame(tick);
  }

  function start() {
    const AC = window.AudioContext || window.webkitAudioContext;
    if (!AC) return;
    const ctx = new AC();
    ctxRef.current = ctx;

    ensureRunningContext(ctx, () => {
      const schedule = computeShuttleSchedule(SHUTTLE_PRESET);
      scheduleRef.current = schedule;
      const now = ctx.currentTime + 0.15;
      // 공인 진행 멘트("다섯, 넷, 셋, 둘, 하나, 시작")와 동일하게 5초 카운트다운
      [0, 1, 2, 3, 4].forEach(i => beepAt(ctx, now + i, 440, 0.15, 0.35));
      const goTime = now + 5;
      beepAt(ctx, goTime, 1046, 0.25, 0.5);
      schedule.forEach(item => {
        const freq = item.levelStart ? 1200 : 880;
        beepAt(ctx, goTime + item.time, freq, item.levelStart ? 0.2 : 0.12, item.levelStart ? 0.5 : 0.4);
      });
      startTimeRef.current = goTime;
      setPlaying(true);
      setLevel(0); setLap(0); setTotalLaps(0); setElapsed(-5);
      rafRef.current = requestAnimationFrame(tick);
    });
  }

  function stop() {
    cancelAnimationFrame(rafRef.current);
    if (ctxRef.current) { ctxRef.current.close(); ctxRef.current = null; }
    setPlaying(false);
    setElapsed(0); setLevel(0); setLap(0); setTotalLaps(0);
  }

  useEffect(() => () => {
    cancelAnimationFrame(rafRef.current);
    if (ctxRef.current) ctxRef.current.close();
  }, []);

  return (
    <div className="tool-stack">
      <div className="board-filters compact">
        <FilterChips label="음원 방식" value={shuttleSource} onChange={setShuttleSource}
          options={[{ id: "beep", label: "자체 신호음" }, { id: "youtube", label: "유튜브 영상" }]} />
      </div>

      {shuttleSource === "youtube" ? (
        <div className="panel">
          <h3>유튜브 영상으로 진행</h3>
          <div className="text-dim small-note">
            유튜브에 "왕복오래달리기 신호음" 등으로 검색해 원하는 영상의 링크를 붙여넣으세요. 버튼을
            누르면 별도 탭/앱에서 유튜브가 열립니다. 그 화면에서 재생하시고, 이 프로그램은 옆에 그대로
            띄워두고 학년·반별 기록 입력에 사용해 주세요.
          </div>
          <div className="form-row">
            <label>유튜브 영상 링크</label>
            <input
              className="input"
              value={youtubeUrl}
              placeholder="https://www.youtube.com/watch?v=..."
              onChange={e => setYoutubeUrl(e.target.value)}
            />
          </div>
          {youtubeId ? (
            <a className="btn btn-primary big-btn youtube-open-btn" href={"https://www.youtube.com/watch?v=" + youtubeId} target="_blank" rel="noopener noreferrer">
              <Play size={16} /> 새 탭에서 유튜브 영상 열기
            </a>
          ) : youtubeUrl ? (
            <div className="warn-note">
              <AlertTriangle size={16} />
              <span>영상 링크를 다시 확인해 주세요. youtube.com 또는 youtu.be 주소여야 합니다.</span>
            </div>
          ) : (
            <div className="empty-state">
              <Info size={32} color="var(--text-dim)" />
              <div>링크를 입력하면 여기에 열기 버튼이 나타납니다.</div>
            </div>
          )}
          <div className="warn-note">
            <AlertTriangle size={16} />
            <span>
              이 영상은 유튜브에 있는 원본 그대로 재생되는 것으로, 저작권은 영상 제작자에게 있습니다.
              단계·회수는 영상 자체의 안내(화면 표시 또는 음성)를 참고해 주세요.
            </span>
          </div>
        </div>
      ) : (
        <>
      <div className="panel timer-display shuttle big-display">
        {playing && elapsed < 0 ? (
          <>
            <div className="timer-big-label">곧 출발</div>
            <div className="timer-big-number huge">{Math.ceil(-elapsed)}</div>
          </>
        ) : (
          <div className="shuttle-stats">
            <div className="shuttle-stat">
              <div className="timer-big-label">단계</div>
              <div className="timer-big-number huge">{level || "-"}</div>
            </div>
            <div className="shuttle-stat">
              <div className="timer-big-label">누적 횟수</div>
              <div className="timer-big-number huge accent">{totalLaps}</div>
            </div>
          </div>
        )}
        {playing && elapsed >= 0 && (
          <div className="timer-sub big-sub">이번 단계 {lap}회째 · 시속 {currentSpeed}km/h</div>
        )}
        <div className="timer-elapsed big-elapsed">{fmtTime(Math.max(0, elapsed))}</div>
        <div className="timer-controls">
          {!playing ? (
            <button className="btn btn-primary big-btn" onClick={start}><Play size={16} /> 신호음 재생</button>
          ) : (
            <button className="btn btn-ghost big-btn" onClick={stop}><Square size={16} /> 정지</button>
          )}
        </div>
      </div>
      <div className="panel">
        <h3>중학생 표준 셔틀런 신호음</h3>
        <div className="preset-info">
          <div className="preset-row"><span>기준 방식</span><span>20m 셔틀런(Léger·PACER 국제 표준 프로토콜)</span></div>
          <div className="preset-row"><span>왕복 거리</span><span>{distance}m</span></div>
          <div className="preset-row"><span>1단계 속도</span><span>시속 {startSpeed}km</span></div>
          <div className="preset-row"><span>단계별 증가</span><span>매 단계 시속 +{increment}km</span></div>
          <div className="preset-row"><span>단계 길이</span><span>약 {levelDuration}초</span></div>
        </div>
        <div className="warn-note">
          <AlertTriangle size={16} />
          <span>중학생 현장 측정에 널리 쓰이는 국제 표준 진행 방식을 그대로 적용해 별도 설정 없이 바로 사용하도록 맞춰두었습니다. 학교마다 사용하는 공식 CD/음원과 초 단위까지 완전히 동일하지는 않을 수 있으니, 정식 기록으로 인정되는 측정에는 학교 지정 음원도 함께 확인해 주세요.</span>
        </div>
      </div>
        </>
      )}
    </div>
  );
}

function extractYoutubeId(url) {
  if (!url) return null;
  const m = url.match(/(?:youtube\.com\/(?:watch\?v=|embed\/|shorts\/)|youtu\.be\/)([a-zA-Z0-9_-]{6,})/);
  return m ? m[1] : null;
}

function fmtTime(sec) {
  const m = Math.floor(sec / 60);
  const s = (sec % 60).toFixed(1);
  return (m > 0 ? m + "분 " : "") + s + "초";
}


function SitupAudioPlayer() {
  const [playing, setPlaying] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(false);
  const [count, setCount] = useState(0);
  const ctxRef = useRef(null);
  const bufferRef = useRef(null);
  const sourceRef = useRef(null);
  const countIntervalRef = useRef(null);

  // 셔틀런 신호음과 동일하게 <audio> 태그 대신 Web Audio API로 직접 디코딩·재생한다.
  // (이 방식이 이 환경에서 이미 안정적으로 동작함이 확인된 방식이다.)
  function ensureCtx() {
    if (!ctxRef.current) {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (AC) ctxRef.current = new AC();
    }
    if (ctxRef.current && ctxRef.current.state === "suspended") {
      ctxRef.current.resume().catch(() => {});
    }
    return ctxRef.current;
  }

  async function getBuffer(ctx) {
    if (bufferRef.current) return bufferRef.current;
    const binary = atob(SITUP_LOOP_AUDIO_B64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const buf = await ctx.decodeAudioData(bytes.buffer.slice(0));
    bufferRef.current = buf;
    return buf;
  }

  function stopPlayback() {
    if (sourceRef.current) {
      try { sourceRef.current.stop(); } catch (e) { /* 이미 멈춘 경우 무시 */ }
      sourceRef.current.disconnect();
      sourceRef.current = null;
    }
    clearInterval(countIntervalRef.current);
    setPlaying(false);
  }

  async function toggle() {
    if (playing) {
      stopPlayback();
      return;
    }
    setError(false);
    const ctx = ensureCtx();
    if (!ctx) { setError(true); return; }
    try {
      setLoading(true);
      const buf = await getBuffer(ctx);
      setLoading(false);
      const src = ctx.createBufferSource();
      src.buffer = buf;
      src.loop = true;
      src.connect(ctx.destination);
      src.start(0);
      sourceRef.current = src;
      setPlaying(true);
      setCount(0);
      // 공식 음원은 3초 간격 신호음이므로, 같은 주기로 실시간 개수를 함께 세어 보여준다.
      countIntervalRef.current = setInterval(() => setCount(c => c + 1), 3000);
    } catch (e) {
      setLoading(false);
      setError(true);
    }
  }

  useEffect(() => () => {
    clearInterval(countIntervalRef.current);
    if (sourceRef.current) { try { sourceRef.current.stop(); } catch (e) { /* noop */ } }
    if (ctxRef.current) ctxRef.current.close();
  }, []);

  return (
    <div className="tool-stack">
      <div className="panel situp-audio-panel">
        <Play size={40} className="situp-audio-icon" />
        <h3 className="situp-audio-title">윗몸말아올리기 공식 음원</h3>
        <div className="text-dim small-note situp-audio-desc">
          교육부 공식 음원의 3초 간격 신호음을 재생합니다. 재생 버튼을 누르면 신호음에 맞춰
          실시간 개수도 함께 세어 보여드립니다.
        </div>
        <div className="situp-count-display">
          <div className="timer-big-label">실시간 개수</div>
          <div className="timer-big-number huge accent">{count}</div>
        </div>
        <button className={"btn btn-primary big-btn situp-audio-btn" + (playing ? " playing" : "")} onClick={toggle} disabled={loading}>
          {loading ? "불러오는 중..." : playing ? <><Square size={18} /> 음원 정지</> : <><Play size={18} /> 공식 음원 재생</>}
        </button>
        {error && (
          <div className="warn-note situp-audio-error">
            <AlertTriangle size={16} />
            <span>음원 재생에 실패했습니다. 화면을 한 번 터치한 뒤 다시 눌러보시거나, 기기의 무음 모드를
            해제하고 다시 시도해 주세요. 계속 안 되면 브라우저를 새로고침해 주세요.</span>
          </div>
        )}
      </div>
    </div>
  );
}

function StepTestPlayer() {
  const [playing, setPlaying] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const [totalDur, setTotalDur] = useState(180);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(false);
  const ctxRef = useRef(null);
  const bufferRef = useRef(null);
  const sourceRef = useRef(null);
  const startTimeRef = useRef(0);
  const rafRef = useRef(null);

  // 윗몸말아올리기 공식 음원과 동일하게, <audio> 태그 대신 Web Audio API로 직접
  // 디코딩·재생한다(이 환경에서 안정적으로 동작함이 확인된 방식).
  function ensureCtx() {
    if (!ctxRef.current) {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (AC) ctxRef.current = new AC();
    }
    if (ctxRef.current && ctxRef.current.state === "suspended") {
      ctxRef.current.resume().catch(() => {});
    }
    return ctxRef.current;
  }

  async function getBuffer(ctx) {
    if (bufferRef.current) return bufferRef.current;
    const binary = atob(STEP_TEST_AUDIO_B64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const buf = await ctx.decodeAudioData(bytes.buffer.slice(0));
    bufferRef.current = buf;
    setTotalDur(buf.duration);
    return buf;
  }

  function tick() {
    const ctx = ctxRef.current;
    if (!ctx) return;
    const el = ctx.currentTime - startTimeRef.current;
    setElapsed(el);
    if (bufferRef.current && el >= bufferRef.current.duration) {
      stop();
      return;
    }
    rafRef.current = requestAnimationFrame(tick);
  }

  function stop() {
    cancelAnimationFrame(rafRef.current);
    if (sourceRef.current) {
      try { sourceRef.current.stop(); } catch (e) { /* 이미 멈춘 경우 무시 */ }
      sourceRef.current.disconnect();
      sourceRef.current = null;
    }
    setPlaying(false);
    setElapsed(0);
  }

  async function start() {
    setError(false);
    const ctx = ensureCtx();
    if (!ctx) { setError(true); return; }
    try {
      setLoading(true);
      const buf = await getBuffer(ctx);
      setLoading(false);
      const src = ctx.createBufferSource();
      src.buffer = buf;
      src.connect(ctx.destination);
      src.onended = () => { setPlaying(false); setElapsed(0); };
      startTimeRef.current = ctx.currentTime;
      src.start(0);
      sourceRef.current = src;
      setElapsed(0);
      setPlaying(true);
      rafRef.current = requestAnimationFrame(tick);
    } catch (e) {
      setLoading(false);
      setError(true);
    }
  }

  useEffect(() => () => {
    cancelAnimationFrame(rafRef.current);
    if (sourceRef.current) { try { sourceRef.current.stop(); } catch (e) { /* noop */ } }
    if (ctxRef.current) ctxRef.current.close();
  }, []);

  return (
    <div className="tool-stack">
      <div className="panel timer-display big-display">
        <div className="timer-big-label">스텝검사 진행</div>
        <div className="timer-big-number huge">{fmtTime(Math.max(0, totalDur - elapsed))}</div>
        <div className="timer-sub big-sub">전체 {fmtTime(totalDur)} 중 {fmtTime(elapsed)} 경과</div>
        <div className="timer-controls">
          {!playing ? (
            <button className="btn btn-primary big-btn" onClick={start} disabled={loading}>
              {loading ? "불러오는 중..." : <><Play size={16} /> 공식 음원 재생</>}
            </button>
          ) : (
            <button className="btn btn-ghost big-btn" onClick={stop}><Square size={16} /> 정지</button>
          )}
        </div>
        {error && (
          <div className="warn-note">
            <AlertTriangle size={16} />
            <span>음원 재생에 실패했습니다. 화면을 한 번 터치한 뒤 다시 눌러보시거나, 기기의 무음 모드를
            해제하고 다시 시도해 주세요. 계속 안 되면 브라우저를 새로고침해 주세요.</span>
          </div>
        )}
      </div>
      <div className="panel">
        <h3>스텝검사 안내</h3>
        <div className="preset-info">
          <div className="preset-row"><span>진행 방식</span><span>3분간 스텝박스 오르내리기(교육부 공식 음원 재생 시간에 맞춤)</span></div>
          <div className="preset-row"><span>스텝박스 높이</span><span>초등 5~6학년 20.3cm · 중(남·여)·고(여) 45.7cm · 고(남) 50.8cm</span></div>
          <div className="preset-row"><span>측정</span><span>스텝운동 종료 후 의자에 앉아 안정을 취하며 심박수 측정</span></div>
        </div>
        <div className="warn-note">
          <AlertTriangle size={16} />
          <span>교육부 학생건강정보센터의 공식 스텝검사 음원을 그대로 재생합니다. 재생이 끝나면(3분) 바로 의자에 앉아 안정 심박수 측정을 시작해 주세요.</span>
        </div>
      </div>
    </div>
  );
}


/* ============================== 스타일 ============================== */

function PapsStyles({ children }) {
  return (
    <div>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap');

        .paps-app {
          --ink: #0B1229;
          --ink-2: #131B3A;
          --panel: #1A234A;
          --panel-2: #212C57;
          --line: rgba(255,255,255,0.09);
          /* 기본 강조색을 각진 경고톤 빨강 대신, 장년층 사용자도 편안하게 느끼는 차분한 파란색으로.
             변수 이름은 기존 코드 곳곳에서 쓰여서 그대로 두고 값만 바꿨다(실제 "위험/삭제" 표시는
             여전히 별도의 빨간색을 그대로 사용한다). */
          --track-red: #3B7DD8;
          --gold: #FFC93C;
          --silver: #C9D3DC;
          --bronze: #CD8B4B;
          --teal: #2EC4B6;
          --text: #F3F5FB;
          --text-dim: #9AA5C7;

          font-family: 'Inter', -apple-system, sans-serif;
          font-size: 15px;
          background: radial-gradient(ellipse at top, var(--ink-2) 0%, var(--ink) 60%);
          color: var(--text);
          min-height: 100vh;
          width: 100%;
          box-sizing: border-box;
        }
        .paps-app * { box-sizing: border-box; }

        .paps-loading {
          display: flex; flex-direction: column; align-items: center; justify-content: center;
          gap: 12px; height: 60vh; color: var(--text-dim); font-size: 14px;
        }
        .spin { animation: spin 1s linear infinite; }
        .spin-slow { animation: spin 2.5s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }

        .gate-screen { display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 20px; }
        .gate-card {
          max-width: 420px; width: 100%; background: var(--panel); border: 1px solid var(--line);
          border-top: 2px solid var(--gold);
          border-radius: 14px; padding: 36px 30px; text-align: center; display: flex; flex-direction: column; align-items: center; gap: 10px;
          box-shadow: 0 20px 50px rgba(0,0,0,0.35);
        }
        .gate-card h2 { font-family: 'Oswald', sans-serif; font-size: 19px; margin: 6px 0 0; font-weight: 600; letter-spacing: 0.2px; }
        .gate-desc { font-size: 13px; color: var(--text-dim); line-height: 1.6; margin: 0 0 8px; }
        .gate-brand { display: flex; align-items: center; justify-content: center; gap: 12px; margin-bottom: 4px; }
        .gate-brand-name { font-family: 'Oswald', sans-serif; font-size: 23px; font-weight: 700; letter-spacing: 1px; color: var(--gold); }
        .gate-tagline { font-size: 13px; color: var(--text-dim); text-align: center; margin: 0 0 8px; line-height: 1.6; }
        .gate-tagline b { color: var(--gold); font-weight: 600; }
        .gate-level-row { width: 100%; margin-bottom: 14px; text-align: center; }
        .gate-level-row label { display: block; font-size: 12px; color: var(--text-dim); margin-bottom: 6px; text-align: center; }
        .gate-level-chips, .gate-level-row .chip-row { justify-content: center; }
        .gate-pw-row { width: 100%; margin: 12px 0; text-align: left; }
        .gate-pw-row > label { display: block; font-size: 12px; color: var(--text-dim); margin-bottom: 6px; }
        .gate-pw-hint { margin-top: 6px; }
        .gate-code-warn { text-align: left; margin: 6px 0 12px; }
        .gate-name-hint { text-align: left; margin: -6px 0 12px; }
        .gate-back-toggle { width: 100%; justify-content: center; margin-top: 10px; font-size: 12px; }
        .gate-divider { width: 48px; height: 2px; background: var(--line); border-radius: 999px; margin: 6px 0 16px; }
        .manual-btn {
          display: inline-flex; align-items: center; gap: 5px; background: rgba(255,201,60,0.12);
          border: 1px solid rgba(255,201,60,0.4); color: var(--gold); border-radius: 999px;
          padding: 6px 14px; font-size: 12px; font-weight: 700; cursor: pointer; margin-bottom: 4px;
        }
        .manual-btn:hover { background: rgba(255,201,60,0.22); }
        .gate-input { width: 100%; text-align: center; margin-bottom: 4px; }
        .gate-input-hint { font-size: 13px; color: var(--text-dim); text-align: center; margin-bottom: 14px; }
        .gate-note { display: flex; gap: 8px; text-align: left; font-size: 12px; color: var(--text-dim); line-height: 1.6; margin-top: 14px; }
        .visit-stats-row {
          display: flex; justify-content: center; gap: 18px; margin-top: 14px; padding-top: 14px;
          border-top: 1px solid var(--line); font-size: 11px; color: var(--text-dim); width: 100%;
        }
        .visit-stats-row b { font-family: 'Oswald', sans-serif; color: var(--gold); font-size: 13px; }
        .gate-link { background: none; border: none; color: var(--text-dim); font-size: 12px; text-decoration: underline; cursor: pointer; margin-top: 10px; }
        .gate-link:hover { color: var(--text); }
        .gate-link-btn {
          display: block; width: 100%; margin-top: 16px; padding: 13px 16px;
          background: rgba(255,255,255,0.03); border: 1px solid var(--line); border-radius: 10px;
          color: var(--text-dim); font-size: 13px; cursor: pointer; text-align: center; line-height: 1.6;
          transition: background 0.15s, border-color 0.15s;
        }
        .gate-link-btn:hover { background: rgba(255,255,255,0.07); border-color: var(--text-dim); }
        .gate-link-cta { display: inline-block; margin-top: 2px; color: var(--gold); font-weight: 700; font-size: 14px; }
        @media (max-width: 420px) {
          .gate-input { font-size: 17px; }
        }
        .gate-notice-list { text-align: left; font-size: 12.5px; line-height: 1.7; color: var(--text); padding-left: 18px; margin: 6px 0 14px; }
        .gate-notice-list li { margin-bottom: 6px; }
        .gate-agree { display: flex; align-items: flex-start; gap: 8px; text-align: left; font-size: 12px; color: var(--text-dim); margin-bottom: 16px; cursor: pointer; }
        .gate-agree input { margin-top: 2px; width: 16px; height: 16px; accent-color: var(--track-red); flex-shrink: 0; }
        .gate-btn-row { display: flex; gap: 10px; width: 100%; justify-content: center; }
        .gate-error { color: #E85D5D; font-size: 12px; margin: -6px 0 12px; text-align: center; }

        .viewer-badge { cursor: default; }
        .tab-badge {
          background: var(--track-red); color: #fff; font-size: 10px; font-weight: 700;
          border-radius: 999px; padding: 1px 6px; margin-left: 2px;
        }
        .access-list { display: flex; flex-direction: column; gap: 8px; }
        .audit-log-list { display: flex; flex-direction: column; gap: 6px; margin-top: 12px; max-height: 360px; overflow-y: auto; }
        .audit-log-row { font-size: 12px; padding: 8px 10px; background: rgba(255,255,255,0.03); border-radius: 6px; line-height: 1.5; }
        .audit-log-body { display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-wrap: wrap; }
        .audit-log-time { display: block; color: var(--text-dim); font-size: 11px; margin-bottom: 2px; }
        .access-row { display: flex; align-items: center; justify-content: space-between; padding: 10px 12px; background: rgba(255,255,255,0.03); border-radius: 10px; border: 1px solid var(--line); }
        .access-name { font-weight: 600; font-size: 13px; }
        .access-type-badge { font-size: 10px; font-weight: 600; padding: 2px 7px; border-radius: 999px; background: rgba(255,255,255,0.08); color: var(--text-dim); margin-left: 6px; vertical-align: middle; }
        .access-type-badge.editor { background: rgba(230,57,70,0.18); color: #FF9AA5; }
        .access-time { font-size: 11px; }
        .access-actions { display: flex; gap: 8px; }

        .workspace-badge {
          background: rgba(255,255,255,0.06); border: 1px solid var(--line); color: var(--text-dim);
          border-radius: 999px; padding: 4px 10px; font-size: 11px; font-family: 'Inter', sans-serif;
          font-weight: 500; cursor: pointer; letter-spacing: 0;
        }
        .workspace-badge:hover { color: var(--text); border-color: var(--text-dim); }

        .feature-updates-btn {
          display: inline-flex; align-items: center; gap: 5px; background: rgba(255,201,60,0.1);
          border: 1px solid rgba(255,201,60,0.35); color: var(--gold); border-radius: 999px;
          padding: 4px 10px; font-size: 11px; font-weight: 600; cursor: pointer;
        }
        .feature-updates-btn:hover { background: rgba(255,201,60,0.2); }
        .feature-update-group { margin-bottom: 16px; }
        .device-guide-caution-list { margin: 6px 0 0; padding-left: 18px; display: flex; flex-direction: column; gap: 6px; }
        .device-guide-caution-list li { list-style: disc; }
        .feature-update-group h4 { font-family: 'Oswald', sans-serif; font-size: 14px; margin: 0 0 8px; color: var(--gold); }
        .feature-update-group ul { margin: 0; padding-left: 18px; display: flex; flex-direction: column; gap: 10px; }
        .feature-update-group li { font-size: 13px; line-height: 1.5; display: flex; flex-direction: column; gap: 2px; }
        .feature-update-headline { font-weight: 700; color: var(--text); }
        .feature-update-desc { font-size: 12px; color: var(--text-dim); }

        .topnav {
          display: flex; align-items: center; gap: 20px;
          padding: 12px 20px; border-bottom: 1px solid var(--line);
          background: rgba(255,255,255,0.02); flex-wrap: wrap;
        }
        .brand { display: flex; align-items: center; gap: 8px; font-family: 'Oswald', sans-serif; font-weight: 600; letter-spacing: 0.3px; }
        .brand-text { font-size: 16px; }
        .tabs { display: flex; gap: 4px; flex: 1; flex-wrap: wrap; }
        .tab {
          display: flex; align-items: center; gap: 6px; padding: 8px 12px;
          background: transparent; border: none; color: var(--text-dim);
          border-radius: 8px; cursor: pointer; font-size: 13px; font-weight: 500;
          transition: background 0.15s, color 0.15s;
        }
        .tab:hover { background: rgba(255,255,255,0.06); color: var(--text); }
        .tab.active { background: var(--panel-2); color: var(--text); }
        .tab-danger { color: #E85D5D; margin-left: 6px; border-left: 1px solid var(--line); padding-left: 14px; border-radius: 0 8px 8px 0; }
        .tab-danger:hover { background: rgba(232,93,93,0.14); color: #FF8080; }
        .tab-danger.active { background: rgba(232,93,93,0.22); color: #FF8080; }
        .nav-right { display: flex; align-items: center; gap: 12px; margin-left: auto; flex-wrap: wrap; }
        .sync-indicator { display: flex; align-items: center; gap: 5px; font-size: 11px; color: var(--text-dim); }

        .info-banner {
          display: flex; align-items: center; gap: 8px; padding: 8px 20px;
          background: rgba(46,196,182,0.1); color: #9FE6DE; font-size: 12px; border-bottom: 1px solid var(--line);
        }
        .info-banner .icon-btn { margin-left: auto; }
        .warn-banner { background: rgba(232,93,93,0.12); color: #F2A5A5; }
        .warn-banner button { margin-left: auto; flex-shrink: 0; }

        .btn {
          display: inline-flex; align-items: center; gap: 6px; padding: 10px 17px;
          border-radius: 8px; border: none; cursor: pointer; font-size: 14px; font-weight: 600;
          transition: transform 0.1s, opacity 0.15s; font-family: 'Inter', sans-serif;
        }
        .btn:active { transform: scale(0.97); }
        .btn:disabled { opacity: 0.4; cursor: not-allowed; }
        .btn-primary { background: var(--track-red); color: white; }
        .btn-primary:hover:not(:disabled) { opacity: 0.9; }
        .btn-secondary { background: var(--panel-2); color: var(--text); }
        .btn-ghost { background: transparent; color: var(--text-dim); border: 1px solid var(--line); }
        .btn-ghost:hover { color: var(--text); border-color: var(--text-dim); }
        .btn.small { padding: 6px 10px; font-size: 12px; }
        .big-btn { width: 100%; justify-content: center; padding: 13px; font-size: 15px; margin-top: 6px; }
        .gate-create-emphasis { border: 2px solid var(--gold); box-shadow: 0 0 0 3px rgba(255,201,60,0.15); }
        .icon-btn { background: transparent; border: none; color: var(--text-dim); cursor: pointer; padding: 4px; border-radius: 6px; }
        .icon-btn:hover { color: var(--text); background: rgba(255,255,255,0.08); }
        .icon-btn.danger:hover { color: var(--track-red); }
        .bmi-bulk-toggle-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin: 4px 0 10px; }

        .paps-body { padding: 20px; }
        .presentation .paps-body { padding: 24px 32px; }

        /* ---------- 전광판 ---------- */
        .board-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
        .board-title { display: flex; align-items: baseline; gap: 14px; font-family: 'Oswald', sans-serif; }
        .board-event { font-size: 32px; font-weight: 600; letter-spacing: 0.3px; }
        .board-cat { font-size: 16px; color: var(--text-dim); font-weight: 400; }

        .board-filters { display: flex; gap: 22px; flex-wrap: wrap; margin-bottom: 22px; padding: 14px 16px; background: rgba(255,255,255,0.03); border-radius: 12px; border: 1px solid var(--line); }
        .board-filters.compact { padding: 10px 12px; margin-bottom: 16px; }
        .filter-group { display: flex; flex-direction: column; gap: 6px; }
        .filter-group.disabled { opacity: 0.4; }
        .filter-label { font-size: 11px; color: var(--text-dim); font-weight: 600; }
        .chip-row { display: flex; gap: 6px; flex-wrap: wrap; }
        .rank-note { font-size: 11px; color: var(--text-dim); margin: -8px 0 16px; text-align: center; }
        .chip {
          padding: 7px 13px; border-radius: 999px; border: 1px solid var(--line);
          background: transparent; color: var(--text-dim); font-size: 13px; cursor: pointer;
          font-weight: 500; white-space: nowrap; transition: all 0.12s;
        }
        .chip:hover { border-color: var(--text-dim); color: var(--text); }
        .chip.active { background: var(--track-red); border-color: var(--track-red); color: white; }
        .chip:disabled { cursor: not-allowed; }

        .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 8px; padding: 60px 20px; color: var(--text-dim); text-align: center; }
        .youtube-open-btn { text-decoration: none; display: flex; align-items: center; justify-content: center; gap: 8px; }
        .empty-sub { font-size: 12px; }

        .podium { display: flex; align-items: flex-end; justify-content: center; gap: 18px; margin-bottom: 26px; }
        @media (max-width: 700px) {
          .podium { flex-direction: column; align-items: stretch; }
          .podium-card { width: 100% !important; }
          .podium-card.tier-1 { order: 1; min-height: 160px; }
          .podium-card.tier-2 { order: 2; min-height: 130px; }
          .podium-card.tier-3 { order: 3; min-height: 130px; }
          .podium-slot.empty { display: none; }
        }
        .podium-slot.empty { width: 200px; }
        .podium-card {
          width: 200px; border-radius: 16px; padding: 20px 16px; text-align: center;
          background: var(--panel); border: 1px solid var(--line); position: relative;
          overflow: hidden;
        }
        .podium-card.tier-1 {
          min-height: 250px;
          background: linear-gradient(180deg, rgba(255,213,74,0.24), var(--panel) 65%);
          border: 2px solid rgba(255,213,74,0.7);
          box-shadow: 0 0 40px rgba(255,201,60,0.3), 0 0 0 1px rgba(255,213,74,0.25) inset;
          animation: trophyGlow 2.2s ease-in-out infinite;
        }
        .podium-card.tier-2 {
          min-height: 190px;
          background: linear-gradient(180deg, rgba(220,228,235,0.2), var(--panel) 65%);
          border-color: rgba(220,228,235,0.5);
          box-shadow: 0 0 16px rgba(200,210,220,0.15);
        }
        .podium-card.tier-3 {
          min-height: 170px;
          background: linear-gradient(180deg, rgba(220,150,90,0.2), var(--panel) 65%);
          border-color: rgba(220,150,90,0.5);
          box-shadow: 0 0 16px rgba(205,139,75,0.15);
        }
        @keyframes trophyGlow {
          0%, 100% { box-shadow: 0 0 32px rgba(255,201,60,0.28), 0 0 0 1px rgba(255,213,74,0.25) inset; }
          50% { box-shadow: 0 0 58px rgba(255,201,60,0.55), 0 0 0 1px rgba(255,213,74,0.5) inset; }
        }
        .podium-shine {
          position: absolute; top: -50%; left: -60%; width: 40%; height: 220%;
          background: linear-gradient(120deg, transparent, rgba(255,255,255,0.6), transparent);
          transform: rotate(20deg); pointer-events: none;
          animation: shineSweep 3s ease-in-out infinite;
        }
        @keyframes shineSweep { 0% { left: -60%; } 45% { left: 130%; } 100% { left: 130%; } }
        .podium-burst {
          position: absolute; top: 4px; left: 50%; transform: translateX(-50%);
          width: 130px; height: 130px; border-radius: 50%;
          background: radial-gradient(circle, rgba(255,213,74,0.5) 0%, transparent 70%);
          pointer-events: none; animation: burstPulse 2.2s ease-in-out infinite;
        }
        @keyframes burstPulse {
          0%, 100% { opacity: 0.5; transform: translateX(-50%) scale(1); }
          50% { opacity: 0.95; transform: translateX(-50%) scale(1.18); }
        }
        .podium-trophy-icon { color: #FFD54A; filter: drop-shadow(0 0 10px rgba(255,213,74,0.85)); position: relative; z-index: 1; }
        .podium-rank { display: flex; flex-direction: column; align-items: center; gap: 2px; margin-bottom: 10px; position: relative; z-index: 1; }
        .tier-1 .podium-rank { color: var(--gold); }
        .tier-2 .podium-rank { color: var(--silver); }
        .tier-3 .podium-rank { color: var(--bronze); }
        .podium-rank span { font-family: 'Oswald', sans-serif; font-size: 13px; font-weight: 600; }
        .podium-name { font-family: 'Oswald', sans-serif; font-size: 22px; font-weight: 600; margin-bottom: 4px; position: relative; z-index: 1; }
        .tier-1 .podium-name {
          font-size: 28px; font-weight: 800; letter-spacing: 0.3px;
          background: linear-gradient(180deg, #FFF9E5, #FFD54A 60%, #B8860B);
          -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
        }
        .podium-meta { font-size: 11px; color: var(--text-dim); margin-bottom: 12px; position: relative; z-index: 1; }
        .podium-value { font-family: 'Oswald', sans-serif; font-size: 24px; font-weight: 600; font-variant-numeric: tabular-nums; position: relative; z-index: 1; }
        .tier-1 .podium-value { font-size: 34px; font-weight: 800; color: #FFD54A; text-shadow: 0 0 18px rgba(255,213,74,0.55); }

        .rank-list { display: flex; flex-direction: column; gap: 6px; max-width: 760px; margin: 0 auto; }

        /* 전체 화면 랭킹모드 */
        .overview-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 14px; }
        .overview-card { background: var(--panel); border: 1px solid var(--line); border-radius: 12px; padding: 14px 16px; }
        .overview-card-head { font-family: 'Oswald', sans-serif; font-size: 15px; font-weight: 700; margin-bottom: 10px; padding-bottom: 8px; border-bottom: 1px solid var(--line); }
        .overview-list { display: flex; flex-direction: column; gap: 6px; }
        .overview-row { display: grid; grid-template-columns: 20px 1fr auto; align-items: center; gap: 8px; }
        .overview-rank { font-family: 'Oswald', sans-serif; font-weight: 800; font-size: 13px; color: var(--text-dim); text-align: center; }
        .overview-rank.r1 { color: var(--gold); font-size: 15px; }
        .overview-rank.r2 { color: var(--silver); }
        .overview-rank.r3 { color: var(--bronze); }
        .overview-name { font-size: 13px; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .overview-value { font-family: 'Oswald', sans-serif; font-size: 13px; color: var(--text-dim); white-space: nowrap; }
        .overview-row:has(.overview-rank.r1) .overview-value { color: var(--gold); font-weight: 700; }
        .overview-empty { font-size: 12px; color: var(--text-dim); padding: 8px 0; }
        .overview-foot { margin-top: 10px; padding-top: 8px; border-top: 1px solid var(--line); font-size: 11px; color: var(--text-dim); text-align: right; }
        .rank-row {
          display: grid; grid-template-columns: 28px 1fr auto auto 90px; align-items: center; gap: 10px;
          padding: 8px 14px; background: rgba(255,255,255,0.03); border-radius: 8px; border: 1px solid var(--line);
        }
        .rank-num { font-family: 'Oswald', sans-serif; color: var(--text-dim); font-size: 14px; text-align: center; }
        .rank-name { font-weight: 600; font-size: 14px; }
        .rank-meta { font-size: 11px; color: var(--text-dim); }
        .rank-value { font-family: 'Oswald', sans-serif; font-size: 16px; text-align: right; font-variant-numeric: tabular-nums; }
        .board-foot { text-align: center; margin-top: 14px; font-size: 11px; color: var(--text-dim); }

        .grade-dot {
          display: inline-flex; align-items: center; justify-content: center;
          width: 26px; height: 26px; border-radius: 50%; color: #10162E; font-weight: 700; font-size: 13px;
          font-family: 'Oswald', sans-serif;
        }
        .grade-dot.small { width: 20px; height: 20px; font-size: 11px; }
        .bmi-cat-dot { width: auto !important; height: auto !important; border-radius: 999px !important; padding: 3px 10px; background: rgba(255,255,255,0.08); color: var(--text) !important; font-weight: 600; }
        .grade-dot.big { width: auto; height: auto; border-radius: 999px; padding: 4px 10px; margin-top: 10px; font-size: 12px; }

        /* ---------- 공통 패널/폼 ---------- */
        .panel { background: var(--panel); border: 1px solid var(--line); border-radius: 14px; padding: 20px; box-shadow: 0 8px 24px rgba(0,0,0,0.18); }
        .panel h3 { margin: 0 0 14px 0; padding-left: 10px; border-left: 3px solid var(--gold); font-family: 'Oswald', sans-serif; font-size: 16px; font-weight: 600; }
        .panel-head-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
        .entry-layout { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
        @media (max-width: 860px) { .entry-layout { grid-template-columns: 1fr; } }

        .form-row { margin-bottom: 14px; }
        .form-row label { display: block; font-size: 12px; color: var(--text-dim); margin-bottom: 6px; font-weight: 500; }
        .form-row-inline { display: flex; gap: 14px; margin-bottom: 14px; }
        .form-row-inline > div { flex: 1; }
        .form-row-inline label { display: block; font-size: 12px; color: var(--text-dim); margin-bottom: 6px; }

        .input, .select, .textarea {
          width: 100%; padding: 10px 12px; border-radius: 8px; border: 1px solid var(--line);
          background: var(--ink-2); color: var(--text); font-size: 14px; font-family: inherit;
        }
        .input:focus, .select:focus, .textarea:focus { outline: 2px solid var(--track-red); outline-offset: 1px; }
        .big-input { font-size: 22px; padding: 14px; text-align: center; font-family: 'Oswald', sans-serif; }
        .textarea { resize: vertical; font-family: 'Menlo', monospace; font-size: 12px; margin-bottom: 10px; }

        .existing-note { font-size: 12px; color: var(--gold); margin-bottom: 10px; }
        .current-student-card { margin-top: 16px; padding: 12px 14px; background: rgba(255,255,255,0.04); border-radius: 10px; text-align: center; }
        .cs-name { font-family: 'Oswald', sans-serif; font-size: 18px; font-weight: 600; }
        .cs-meta { font-size: 12px; color: var(--text-dim); }

        .recent-list { display: flex; flex-direction: column; gap: 6px; }
        .recent-row { display: grid; grid-template-columns: 1fr auto auto; gap: 10px; padding: 8px 10px; background: rgba(255,255,255,0.03); border-radius: 8px; font-size: 13px; }
        .recent-name { font-weight: 600; }
        .recent-event { color: var(--text-dim); }
        .recent-value { font-family: 'Oswald', sans-serif; font-variant-numeric: tabular-nums; }

        .divider { height: 1px; background: var(--line); margin: 18px 0; }
        .small-note { font-size: 13px; line-height: 1.6; margin-bottom: 8px; }
        .text-dim { color: var(--text-dim); }
        .warn-note { display: flex; gap: 8px; align-items: flex-start; font-size: 12px; color: #FFD98A; background: rgba(255,201,60,0.08); padding: 10px 12px; border-radius: 8px; margin-bottom: 14px; }

        /* 학기 마감 */
        .closeout-hero { display: flex; flex-direction: column; align-items: center; text-align: center; gap: 10px; padding: 32px 24px; }
        .closeout-title { font-size: 26px; font-weight: 800; color: #E85D5D; margin: 0; }
        .closeout-lead { font-size: 14px; color: var(--text); line-height: 1.7; max-width: 560px; }
        .closeout-checklist { list-style: none; padding: 0; margin: 0 0 16px; display: flex; flex-direction: column; gap: 10px; }
        .closeout-checklist li { font-size: 13px; line-height: 1.6; padding: 10px 14px; background: rgba(255,255,255,0.03); border-radius: 8px; border-left: 3px solid #E85D5D; }
        .closeout-option { display: flex; align-items: center; gap: 8px; font-size: 13px; cursor: pointer; }
        .closeout-option input { width: 16px; height: 16px; accent-color: #E85D5D; cursor: pointer; flex-shrink: 0; }
        .closeout-summary { font-size: 13px; color: var(--text-dim); margin: 14px 0; }
        .closeout-btn { background: #E85D5D !important; border-color: #E85D5D !important; font-size: 15px; padding: 14px 20px; width: 100%; justify-content: center; }
        .closeout-backup-step { background: rgba(255,201,60,0.08); border: 1px solid rgba(255,201,60,0.3); border-radius: 10px; padding: 14px; margin: 14px 0; }
        .closeout-backup-done { display: flex; align-items: center; gap: 6px; color: #7FD98A; font-size: 13px; margin-top: 8px; font-weight: 600; }
        .closeout-neis-export-hint { margin-top: 14px; }
        .closeout-step-title { font-weight: 700; font-size: 14px; margin: 16px 0 8px; }
        .closeout-footer { display: flex; flex-direction: column; align-items: center; gap: 10px; margin-top: 22px; padding-top: 18px; border-top: 1px solid var(--line); text-align: center; }
        .closeout-footer-text { font-size: 12px; color: var(--text-dim); line-height: 1.7; margin: 0; }
        .closeout-confirm-big { font-size: 20px; font-weight: 800; color: #E85D5D; text-align: center; line-height: 1.5; margin-bottom: 6px; }
        .closeout-confirm-warn { font-size: 14px; font-weight: 700; color: #FFD98A; text-align: center; margin-bottom: 18px; }
        .closeout-confirm-check { margin-bottom: 16px; padding: 10px 12px; background: rgba(232,93,93,0.1); border-radius: 8px; }
        .danger-confirm-btn { background: #E85D5D !important; border-color: #E85D5D !important; }
        .danger-confirm-btn:disabled { opacity: 0.4; }
        .warn-note.official-note { color: #7FD98A; background: rgba(127,217,138,0.08); }
        .export-option { padding: 12px 0; }
        .export-option-head { display: flex; align-items: center; gap: 7px; font-weight: 600; font-size: 13px; margin-bottom: 6px; }
        .export-option .small-note { margin-bottom: 10px; }
        

        .roster-list { display: flex; flex-direction: column; gap: 6px; max-height: 480px; overflow-y: auto; }
        .roster-row { display: grid; grid-template-columns: 20px 70px 1fr auto 30px; align-items: center; gap: 8px; padding: 8px 10px; background: rgba(255,255,255,0.03); border-radius: 8px; font-size: 13px; }
        .roster-row input[type="checkbox"] { width: 16px; height: 16px; accent-color: var(--track-red); cursor: pointer; }
        .roster-meta { color: var(--text-dim); font-variant-numeric: tabular-nums; font-size: 12px; }
        .roster-gender { color: var(--text-dim); font-size: 12px; }

        .bulk-gender-bar {
          display: flex; align-items: center; gap: 14px; flex-wrap: wrap;
          padding: 10px 12px; background: rgba(255,255,255,0.03); border-radius: 10px; border: 1px solid var(--line);
          margin-bottom: 10px;
        }
        .select-all-label { display: flex; align-items: center; gap: 6px; font-size: 12px; cursor: pointer; }
        .select-all-label input[type="checkbox"] { width: 16px; height: 16px; accent-color: var(--track-red); cursor: pointer; }
        .bulk-gender-actions { display: flex; gap: 8px; margin-left: auto; }
        .danger-btn:hover:not(:disabled) { color: var(--track-red); border-color: var(--track-red); }
        .chip.small { padding: 4px 9px; font-size: 11px; }

        .bulk-undo-row { display: flex; align-items: center; gap: 10px; margin-top: 8px; }
        .dropzone {
          border: 2px dashed var(--line); border-radius: 12px; padding: 22px 14px; text-align: center;
          cursor: pointer; display: flex; flex-direction: column; align-items: center; gap: 6px;
          color: var(--text-dim); transition: border-color 0.15s, background 0.15s; margin-bottom: 4px;
        }
        .dropzone:hover { border-color: var(--text-dim); }
        .dropzone.drag-over { border-color: var(--track-red); background: rgba(230,57,70,0.08); color: var(--text); }
        .dropzone-text { font-size: 13px; font-weight: 500; }
        .visually-hidden-input {
          position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px;
          overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0;
        }
        .file-label-btn { user-select: none; }

        .table-wrap { overflow-x: auto; }

        .checklist-section { margin-bottom: 16px; }
        .checklist-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; gap: 10px; }
        .checklist-title { font-size: 12px; font-weight: 600; color: var(--text-dim); }
        .check-chip-row { display: flex; gap: 6px; flex-wrap: wrap; }
        .check-chip {
          display: flex; align-items: center; gap: 6px; padding: 6px 11px; border-radius: 999px;
          border: 1px solid var(--line); background: transparent; color: var(--text-dim); font-size: 12px; cursor: pointer;
          transition: all 0.12s;
        }
        .check-chip input { width: 13px; height: 13px; accent-color: var(--track-red); cursor: pointer; }
        .check-chip.active { border-color: var(--track-red); color: var(--text); background: rgba(230,57,70,0.1); }
        .bulk-entry .table-wrap { border: 1px solid var(--line); border-radius: 10px; }
        .bulk-table { border-collapse: collapse; font-size: 13px; width: 100%; }
        .bulk-table th, .bulk-table td { padding: 7px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }
        .bulk-table th { color: var(--text-dim); font-weight: 600; font-size: 11px; text-align: left; background: rgba(255,255,255,0.03); }
        .bulk-table tr:last-child td { border-bottom: none; }
        .sticky-col { position: sticky; left: 0; background: var(--panel); z-index: 1; }
        .bulk-table th.sticky-col { z-index: 2; }
        .bulk-cell-input { width: 78px; padding: 6px 8px; text-align: center; font-family: 'Oswald', sans-serif; }
        .bulk-hint { margin-top: 10px; }
        .grade-table { width: 100%; border-collapse: collapse; font-size: 15px; }
        .grade-table th { text-align: left; padding: 10px 12px; color: var(--text-dim); font-weight: 600; font-size: 12px; border-bottom: 1px solid var(--line); white-space: nowrap; }
        .bmi-hide-toggle { display: inline-flex; align-items: center; gap: 4px; margin-left: 8px; font-weight: 500; font-size: 11px; color: var(--text-dim); cursor: pointer; text-transform: none; }
        .bmi-hide-toggle input { width: 12px; height: 12px; accent-color: var(--track-red); cursor: pointer; }
        .bmi-cell { display: flex; align-items: center; gap: 6px; cursor: pointer; }
        .bmi-cell input { width: 13px; height: 13px; accent-color: var(--track-red); cursor: pointer; flex-shrink: 0; }
        .grade-table td { padding: 10px 12px; border-bottom: 1px solid var(--line); white-space: nowrap; }
        .student-cell-name { font-weight: 700; font-size: 16px; }
        .student-cell-meta { font-size: 12px; color: var(--text-dim); }
        .cell-grade { display: flex; align-items: center; gap: 8px; }
        .cell-value { font-size: 14px; color: var(--text-dim); }
        .grade-table .grade-dot { width: 30px; height: 30px; font-size: 15px; }
        .grade-table .grade-dot.small { width: 24px; height: 24px; font-size: 13px; }
        .table-foot { margin-top: 10px; font-size: 11px; }


        .band-table { margin-bottom: 14px; }
        .band-head-row, .band-row { display: grid; grid-template-columns: 50px 1fr 1fr; gap: 10px; align-items: center; margin-bottom: 8px; }
        .band-head-row { font-size: 11px; color: var(--text-dim); font-weight: 600; }
        .band-static { font-size: 13px; padding: 8px 10px; background: rgba(255,255,255,0.03); border-radius: 6px; text-align: center; font-family: 'Oswald', sans-serif; }
        .pw-row { display: flex; gap: 8px; }
        .school-save-btn { flex-shrink: 0; display: flex; align-items: center; gap: 6px; font-weight: 700; white-space: nowrap; }
        .school-save-btn.pending { box-shadow: 0 0 0 3px rgba(230,57,70,0.35); animation: saveBtnPulse 1.4s ease-in-out infinite; }
        .school-save-btn:disabled { opacity: 0.4; box-shadow: none; animation: none; }
        @keyframes saveBtnPulse {
          0%, 100% { box-shadow: 0 0 0 3px rgba(230,57,70,0.35); }
          50% { box-shadow: 0 0 0 6px rgba(230,57,70,0.15); }
        }
        .pw-row .input { flex: 1; }

        .toast {
          position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
          background: var(--panel-2); border: 1px solid var(--line); padding: 12px 22px;
          border-radius: 999px; font-size: 13px; font-weight: 600; box-shadow: 0 8px 24px rgba(0,0,0,0.4);
          z-index: 50; animation: toast-in 0.25s ease-out;
        }
        .toast-record {
          top: 0; bottom: auto; left: 0; right: 0; transform: none; width: 100%;
          border-radius: 0; border: none; border-bottom: 3px solid #FFF3C4;
          padding: 20px 24px; font-size: clamp(16px, 2.4vw, 26px); font-weight: 800; text-align: center;
          color: #241a00; letter-spacing: 0.3px;
          background: linear-gradient(90deg, #B8860B, #FFD54A 30%, #FFF3C4 50%, #FFD54A 70%, #B8860B);
          background-size: 250% 100%;
          box-shadow: 0 6px 28px rgba(0,0,0,0.5);
          z-index: 100;
          animation: recordBannerIn 0.4s ease-out, recordBannerShine 3.5s linear infinite;
        }
        @keyframes toast-in { from { opacity: 0; transform: translateX(-50%) translateY(10px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }
        @keyframes recordBannerIn { from { opacity: 0; transform: translateY(-100%); } to { opacity: 1; transform: translateY(0); } }
        @keyframes recordBannerShine { 0% { background-position: 0% 0; } 100% { background-position: 250% 0; } }

        .year-badge { padding: 6px 10px; font-size: 12px; font-weight: 600; color: var(--text-dim); background: rgba(255,255,255,0.04); border: 1px solid var(--line); border-radius: 8px; }
        .datetime-row { display: flex; align-items: center; gap: 8px; }
        .now-text { font-size: 12px; color: var(--text-dim); font-variant-numeric: tabular-nums; white-space: nowrap; }
        .entry-year-tag { font-size: 12px; font-weight: 500; margin-left: 8px; }

        /* ---------- 타이머 ---------- */
        .timer-body { margin-top: 16px; }

        /* ---------- 기록관리 ---------- */
        .record-mgmt { display: flex; flex-direction: column; gap: 18px; }
        .category-groups { display: flex; flex-direction: row; flex-wrap: wrap; gap: 10px; align-items: flex-start; }
        .method-ref-row { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
        .method-ref-tip { flex: 1; min-width: 220px; margin: 0; }
        .method-ref-btn { flex-shrink: 0; white-space: nowrap; display: flex; align-items: center; gap: 6px; }
        .method-ref-url-row { margin-top: 10px; }
        .method-ref-url { width: 100%; font-size: 12px; color: var(--text-dim); }
        .official-audio-row { display: flex; flex-direction: column; align-items: center; gap: 8px; margin-top: 14px; padding-top: 14px; border-top: 1px solid var(--line); width: 100%; }

        /* 윗몸말아올리기 공식 음원 전용 화면 */
        .situp-audio-panel {
          display: flex; flex-direction: column; align-items: center; text-align: center; gap: 10px;
          padding: 36px 24px; background: linear-gradient(180deg, rgba(255,201,60,0.1), var(--panel) 70%);
          border: 1px solid rgba(255,201,60,0.3);
        }
        .situp-audio-icon { color: var(--gold); filter: drop-shadow(0 0 10px rgba(255,213,74,0.6)); }
        .situp-audio-title { font-size: 20px; margin: 0; }
        .situp-audio-desc { max-width: 420px; }
        .situp-audio-btn { font-size: 17px; padding: 16px 32px; margin-top: 6px; }
        .situp-audio-btn.playing { background: var(--track-red); animation: situpPulse 1.4s ease-in-out infinite; }
        @keyframes situpPulse {
          0%, 100% { box-shadow: 0 0 0 0 rgba(230,57,70,0.5); }
          50% { box-shadow: 0 0 0 10px rgba(230,57,70,0); }
        }
        .situp-audio-error { width: 100%; max-width: 420px; margin-top: 6px; text-align: left; }
        .situp-count-display { display: flex; flex-direction: column; align-items: center; margin: 6px 0; }
        .category-group {
          display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
          padding: 8px 12px; border-radius: 10px; background: rgba(255,255,255,0.03);
          border: 1px solid var(--line); border-left: 3px solid var(--cat-color, var(--text-dim));
        }
        .category-label {
          font-size: 11px; font-weight: 800; letter-spacing: 0.3px; color: #10162E;
          padding: 4px 10px; border-radius: 999px; white-space: nowrap; flex-shrink: 0;
        }
        .drill-row { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 10px; }
        .drill-chip {
          display: flex; align-items: center; gap: 8px; padding: 9px 16px; border-radius: 10px;
          border: 1px solid var(--line); background: rgba(255,255,255,0.05); color: var(--text-dim);
          font-size: 13px; font-weight: 700; cursor: pointer; transition: all 0.12s;
          box-shadow: 0 2px 0 rgba(0,0,0,0.25), 0 1px 3px rgba(0,0,0,0.2);
        }
        .drill-chip:hover { border-color: var(--text-dim); color: var(--text); background: rgba(255,255,255,0.09); transform: translateY(-1px); box-shadow: 0 3px 0 rgba(0,0,0,0.25), 0 2px 5px rgba(0,0,0,0.25); }
        .drill-chip:active { transform: translateY(1px); box-shadow: 0 0 0 rgba(0,0,0,0.25); }
        .drill-chip.active { border-color: var(--track-red); background: rgba(230,57,70,0.16); color: var(--text); box-shadow: 0 2px 0 rgba(150,30,40,0.4), 0 1px 3px rgba(0,0,0,0.2); }
        .drill-count { font-size: 11px; font-family: 'Oswald', sans-serif; color: var(--text-dim); background: rgba(255,255,255,0.06); padding: 1px 6px; border-radius: 999px; }
        .drill-chip.active .drill-count { color: var(--text); }
        .drill-students { display: flex; flex-direction: column; gap: 6px; }
        .drill-student-row { display: grid; grid-template-columns: 1fr 90px 20px; align-items: center; gap: 10px; padding: 8px 12px; background: rgba(255,255,255,0.03); border-radius: 8px; }
        .drill-student-row.two-trial { grid-template-columns: 1fr 90px 90px 120px 20px; }
        .drill-student-row.grip-row { grid-template-columns: 1fr 62px 62px 62px 62px 150px 20px; }
        .drill-student-row.bmi-row { grid-template-columns: 1fr 90px 90px 90px 180px; }
        .drill-student-row.bodyfat-row { grid-template-columns: 1fr 90px 130px 140px; }
        .drill-student-row.flex-total-row { grid-template-columns: 1fr repeat(4, 96px) 110px; align-items: center; }
        .drill-student-row.step-test-row { grid-template-columns: 1fr 90px 90px 90px 100px; }
        .flex-part-group { display: flex; flex-direction: column; align-items: center; gap: 4px; }
        .flex-part-label { font-size: 10px; color: var(--text-dim); }
        .flex-part-btns { display: flex; gap: 3px; }
        .flex-score-btn {
          width: 24px; height: 24px; border-radius: 6px; border: 1px solid var(--line);
          background: rgba(255,255,255,0.04); color: var(--text-dim); font-size: 12px; font-weight: 700;
          cursor: pointer; font-family: 'Oswald', sans-serif;
        }
        .flex-score-btn.active { background: var(--track-red); border-color: var(--track-red); color: #fff; }
        .grip-result-col { display: flex; flex-direction: column; align-items: flex-end; gap: 3px; }
        .drill-ref-note { font-size: 10px; color: var(--text-dim); white-space: nowrap; }
        .drill-best {
          font-size: 12px; color: var(--text-dim); white-space: nowrap; text-align: right;
          font-family: 'Oswald', sans-serif;
        }
        .drill-best.has-value {
          font-size: 14px; font-weight: 800; color: var(--gold); letter-spacing: 0.2px;
          background: rgba(255,201,60,0.12); padding: 4px 10px; border-radius: 999px;
          border: 1px solid rgba(255,201,60,0.3);
        }
        .bmi-ref-tag { font-size: 11px; color: var(--text-dim); white-space: nowrap; text-align: right; }
        .bmi-direct-toggle { display: flex; align-items: center; gap: 4px; font-size: 11px; color: var(--text-dim); white-space: nowrap; cursor: pointer; }
        @media (max-width: 760px) {
          .drill-student-row.two-trial, .drill-student-row.grip-row, .drill-student-row.bmi-row, .drill-student-row.bodyfat-row {
            grid-template-columns: 1fr 1fr; row-gap: 6px;
          }
          .drill-student-row.two-trial .drill-student-name,
          .drill-student-row.grip-row .drill-student-name,
          .drill-student-row.bmi-row .drill-student-name,
          .drill-student-row.bodyfat-row .drill-student-name { grid-column: 1 / -1; }
          .drill-student-row.two-trial .drill-best,
          .drill-student-row.grip-row .grip-result-col,
          .drill-student-row.bmi-row .drill-best,
          .drill-student-row.bmi-row .bmi-ref-tag,
          .drill-student-row.bodyfat-row .drill-best,
          .drill-student-row.bodyfat-row .bmi-ref-tag { grid-column: 1 / -1; text-align: left; }
          .drill-student-row.grip-row .grip-result-col { align-items: flex-start; }
          .drill-student-row.flex-total-row {
            grid-template-columns: repeat(4, 1fr); row-gap: 8px;
          }
          .drill-student-row.flex-total-row .drill-student-name,
          .drill-student-row.flex-total-row .drill-best { grid-column: 1 / -1; text-align: left; }
          .drill-student-row.step-test-row {
            grid-template-columns: 1fr 1fr; row-gap: 6px;
          }
          .drill-student-row.step-test-row .drill-student-name,
          .drill-student-row.step-test-row .drill-best { grid-column: 1 / -1; text-align: left; }
        }
        .drill-student-name { font-size: 13px; font-weight: 600; }
        .drill-student-input { text-align: center; font-family: 'Oswald', sans-serif; padding: 6px 8px; }
        .drill-panel-head { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 8px; }
        .drill-panel-head h3 { margin: 0; }
        .mask-toggle-btn { display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }

        /* 종목 선택 칩 + 생략 체크 */
        .event-chip { display: flex; align-items: center; gap: 8px; }
        .chip-skip { display: flex; align-items: center; padding-left: 8px; margin-left: 2px; border-left: 1px solid var(--line); }
        .mini-check { width: 13px; height: 13px; border-radius: 3px; border: 1px solid var(--text-dim); display: flex; align-items: center; justify-content: center; }
        .mini-check.on { background: var(--track-red); border-color: var(--track-red); color: #fff; }

        /* 종목별 도구(왕복오래달리기·윗몸말아올리기·팔굽혀펴기): 좌측 큰 화면 + 우측 스크롤 기록입력 */
        .tool-drill-row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }
        @media (max-width: 900px) {
          .tool-drill-row { grid-template-columns: 1fr; }
          /* 화면이 좁아지면 종목 도구보다 학년·반별 기록 입력이 먼저(위로) 오도록 순서를 바꾼다 */
          .tool-drill-row .drill-panel { order: -1; }
        }
        .drill-panel .drill-scroll { max-height: 420px; overflow-y: auto; padding-right: 4px; }
        .big-display { display: flex; flex-direction: column; align-items: center; }
        .timer-big-number.huge { font-size: 96px; line-height: 1; }
        .shuttle-stats .timer-big-number.huge { font-size: 76px; }
        .big-sub { font-size: 18px; margin-top: 8px; }
        .timer-elapsed.big-elapsed { font-size: 22px; }
        .timer-unit-label { font-size: 14px; color: var(--text-dim); margin-top: -6px; }
        .tool-stack.square .timer-display.square-display { aspect-ratio: 1 / 1; display: flex; flex-direction: column; align-items: center; justify-content: center; max-width: 420px; margin: 0 auto; }
        .tool-stack.square .timer-display.square-display .timer-big-number.huge { font-size: 110px; }


        .fiftym-group { display: flex; flex-direction: column; align-items: center; gap: 14px; }
        .fiftym-clock { font-family: 'Oswald', sans-serif; font-size: 48px; font-weight: 700; font-variant-numeric: tabular-nums; }
        .fiftym-controls { display: flex; gap: 10px; }

        /* 오래달리기-걷기 반 전체 타이머 */
        .class-run-timer { display: flex; flex-direction: column; align-items: center; gap: 12px; }

        /* 오래달리기-걷기 운동장 코스 계산기 */
        .track-calc-inputs { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin: 10px 0; }
        .track-calc-summary { display: flex; flex-wrap: wrap; gap: 16px; font-size: 13px; color: var(--text-dim); margin: 10px 0; padding: 10px 12px; background: rgba(255,255,255,0.03); border-radius: 8px; }
        .track-calc-summary b { color: var(--text); font-family: 'Oswald', sans-serif; }
        .track-calc-diagram-wrap { display: flex; justify-content: center; margin: 14px 0; }
        .track-calc-svg { width: 100%; max-width: 340px; height: auto; }
        .track-calc-results { display: flex; flex-direction: column; gap: 6px; margin: 10px 0; }
        .track-calc-result-row { display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; background: rgba(255,255,255,0.03); border-radius: 8px; font-size: 13px; }
        .track-calc-result-label { color: var(--text-dim); }
        .track-calc-result-value { font-weight: 700; font-family: 'Oswald', sans-serif; }
        @media (max-width: 600px) {
          .track-calc-inputs { grid-template-columns: 1fr; }
        }
        .class-run-sticky-header {
          position: sticky; top: 0; z-index: 5; width: 100%;
          display: flex; flex-direction: column; align-items: center; gap: 10px;
          background: var(--panel); padding: 10px 0 14px; border-bottom: 1px solid var(--line);
        }
        .class-run-clock { font-family: 'Oswald', sans-serif; font-size: 48px; font-weight: 700; font-variant-numeric: tabular-nums; }
        .class-run-controls { display: flex; gap: 10px; }
        .class-run-progress { font-size: 13px; color: var(--text-dim); }
        .class-run-lists { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; width: 100%; align-items: start; }
        .class-run-col h4 { font-size: 13px; margin: 0 0 8px; color: var(--text-dim); }
        .class-run-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(90px, 1fr)); gap: 6px; max-height: 360px; overflow-y: auto; padding-right: 4px; }
        .class-run-btn {
          padding: 10px 8px; border-radius: 8px; border: 1px solid var(--line);
          background: rgba(255,255,255,0.05); color: var(--text); font-size: 12px; font-weight: 600;
          cursor: pointer; transition: all 0.1s;
        }
        .class-run-btn:hover:not(:disabled) { background: rgba(230,57,70,0.2); border-color: var(--track-red); transform: translateY(-1px); }
        .class-run-btn:active:not(:disabled) { transform: scale(0.96); }
        .class-run-btn:disabled { opacity: 0.3; cursor: not-allowed; }
        .class-run-finished-list { display: flex; flex-direction: column; gap: 4px; max-height: 360px; overflow-y: auto; padding-right: 4px; }
        .class-run-finished-row {
          display: grid; grid-template-columns: 1fr 70px 26px; align-items: center; gap: 6px;
          padding: 6px 8px; background: rgba(127,217,138,0.08); border-radius: 6px;
        }
        .class-run-finished-name { font-size: 12px; font-weight: 600; }
        .class-run-time-input { text-align: center; font-family: 'Oswald', sans-serif; font-size: 12px; padding: 4px; }
        @media (max-width: 700px) {
          .class-run-lists { grid-template-columns: 1fr; }
        }
        .fiftym-count-row { display: flex; align-items: center; gap: 10px; }
        .fiftym-count-label { font-size: 12px; color: var(--text-dim); }
        .fiftym-count-value { font-family: 'Oswald', sans-serif; font-size: 15px; min-width: 32px; text-align: center; }
        .fiftym-count-btn {
          width: 34px; height: 34px; border-radius: 8px; border: 1px solid var(--line);
          background: rgba(255,255,255,0.06); color: var(--text); font-size: 20px; font-weight: 700;
          line-height: 1; cursor: pointer; display: flex; align-items: center; justify-content: center;
        }
        .fiftym-count-btn:hover:not(:disabled) { background: rgba(230,57,70,0.18); border-color: var(--track-red); }
        .fiftym-count-btn:disabled { opacity: 0.35; cursor: not-allowed; }
        .fiftym-slots { width: 100%; display: flex; flex-direction: column; gap: 8px; }
        .fiftym-slot { display: grid; grid-template-columns: 40px 1fr 70px 70px; align-items: center; gap: 8px; padding: 8px 10px; background: rgba(255,255,255,0.03); border-radius: 8px; }
        .fiftym-picker { display: flex; gap: 4px; min-width: 0; }
        .fiftym-picker .select { min-width: 0; flex: 1; padding: 6px 4px; font-size: 12px; }
        .fiftym-slot-time-input { text-align: center; font-family: 'Oswald', sans-serif; }
        @media (max-width: 700px) {
          .fiftym-slot { grid-template-columns: 1fr; gap: 6px; }
          .fiftym-picker { flex-wrap: wrap; }
        }
        .fiftym-slot-num { font-family: 'Oswald', sans-serif; font-size: 13px; color: var(--text-dim); }
        .fiftym-slot-time { font-family: 'Oswald', sans-serif; font-size: 15px; text-align: right; font-variant-numeric: tabular-nums; }
        .timer-grid { display: grid; grid-template-columns: 1.1fr 1fr; gap: 18px; }
        @media (max-width: 860px) { .timer-grid { grid-template-columns: 1fr; } }
        .timer-display { display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; gap: 6px; padding: 36px 20px; }
        .timer-display.shuttle { background: linear-gradient(180deg, rgba(255,201,60,0.08), var(--panel) 60%); }
        .timer-big-label { font-size: 13px; color: var(--text-dim); font-weight: 600; letter-spacing: 0.5px; }
        .timer-big-number { font-family: 'Oswald', sans-serif; font-size: 72px; font-weight: 700; line-height: 1; margin: 6px 0; font-variant-numeric: tabular-nums; }
        .timer-big-number.accent { color: var(--gold); }
        .timer-big-number.small-fit { font-size: 40px; }
        .shuttle-stats { display: flex; gap: 28px; justify-content: center; align-items: flex-end; }
        .shuttle-stat { display: flex; flex-direction: column; align-items: center; }
        .shuttle-stat .timer-big-number { font-size: 56px; margin: 4px 0; }
        .shuttle-stat .timer-big-label { font-size: 12px; }
        .timer-big-number.mono { font-size: 56px; }
        .timer-sub { font-size: 13px; color: var(--text-dim); margin-bottom: 4px; }
        .timer-elapsed { font-size: 14px; color: var(--text-dim); font-variant-numeric: tabular-nums; margin-bottom: 14px; }
        .timer-controls { display: flex; gap: 10px; justify-content: center; margin-top: 8px; }

        .preset-info { display: flex; flex-direction: column; gap: 6px; margin-bottom: 14px; }
        .preset-row { display: flex; justify-content: space-between; font-size: 13px; padding: 7px 10px; background: rgba(255,255,255,0.03); border-radius: 8px; }
        .preset-row span:first-child { color: var(--text-dim); }
        .preset-row span:last-child { font-weight: 600; }
        .voice-toggle {
          display: flex; align-items: center; gap: 8px; font-size: 12px; color: var(--text-dim);
          margin-bottom: 12px; cursor: pointer; background: none; border: none; padding: 4px 0;
          font-family: inherit; text-align: left;
        }
        .voice-toggle:disabled { cursor: not-allowed; opacity: 0.6; }
        .voice-toggle-box {
          flex-shrink: 0; width: 17px; height: 17px; border-radius: 4px; border: 1.5px solid var(--text-dim);
          display: flex; align-items: center; justify-content: center; color: #10162E; transition: all 0.12s;
        }
        .voice-toggle.on .voice-toggle-box { background: var(--track-red); border-color: var(--track-red); color: #fff; }

        .modal-backdrop {
          position: fixed; inset: 0; background: rgba(5,8,20,0.72); z-index: 100;
          display: flex; align-items: center; justify-content: center; padding: 20px;
        }
        .idle-lock-backdrop { background: rgba(5,8,20,0.97); z-index: 200; }
        .mask-toggle-btn.active { background: var(--track-red); border-color: var(--track-red); color: #fff; }
        .modal-panel { background: var(--panel); border: 1px solid var(--line); border-radius: 16px; max-width: 480px; width: 100%; max-height: 85vh; overflow-y: auto; padding: 20px; }
        .modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
        .modal-head h3 { display: flex; align-items: center; gap: 8px; margin: 0; font-family: 'Oswald', sans-serif; font-size: 17px; }
        .share-step { display: flex; gap: 12px; margin-bottom: 14px; }
        .share-step-num {
          flex-shrink: 0; width: 24px; height: 24px; border-radius: 50%; background: var(--track-red);
          display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; font-family: 'Oswald', sans-serif;
        }
        .share-step-title { font-weight: 600; font-size: 13px; margin-bottom: 2px; }
        .share-step-body { font-size: 12px; color: var(--text-dim); line-height: 1.5; }
        .manual-note-btn {
          background: none; border: none; color: var(--gold); font-size: 11px; font-weight: 600;
          cursor: pointer; padding: 4px 0; margin-top: 2px;
        }
        .manual-note-box {
          font-size: 12px; color: var(--text-dim); line-height: 1.6; margin-top: 6px;
          padding: 10px 12px; background: rgba(255,201,60,0.06); border: 1px solid rgba(255,201,60,0.25);
          border-radius: 8px;
        }
        .share-link-row { display: flex; gap: 8px; margin: 14px 0 8px; }
        .share-link-row .input { font-size: 12px; }
        .confirm-message { font-size: 13px; line-height: 1.6; color: var(--text); margin-bottom: 18px; }
        .confirm-actions { display: flex; justify-content: flex-end; gap: 8px; }

        .backup-steps { display: flex; flex-direction: column; gap: 16px; margin: 12px 0; }
        .backup-step { display: flex; gap: 12px; }
        .backup-step-num {
          flex-shrink: 0; width: 26px; height: 26px; border-radius: 50%; background: var(--track-red);
          display: flex; align-items: center; justify-content: center; font-size: 13px; font-weight: 700; font-family: 'Oswald', sans-serif;
        }
        .backup-step-title { font-weight: 700; font-size: 14px; margin-bottom: 4px; }
        .backup-step .small-note { margin-bottom: 8px; }
        .backup-log-latest { font-size: 12px; color: var(--text-dim); margin-top: 10px; }
        .backup-log-latest b { color: var(--gold); font-weight: 700; }
        .backup-log-empty { margin-top: 10px; }
        .backup-log-more { margin-top: 8px; font-size: 12px; color: var(--text-dim); }
        .backup-log-more summary { cursor: pointer; }
        .backup-log-list { display: flex; flex-direction: column; gap: 4px; margin-top: 8px; padding-left: 4px; }
        .backup-log-row { font-size: 12px; color: var(--text-dim); }
        .backup-copy-box {
          background: rgba(255,255,255,0.05); border: 1px dashed var(--line); border-radius: 8px;
          padding: 10px 12px; font-size: 13px; color: var(--text); line-height: 1.6;
        }
        .neis-preview-wrap { margin: 12px 0; }

        /* ================= 챔피언십 모드(테마) ================= */
        .theme-toggle-btn.on { color: var(--gold); border-color: var(--gold); }

        .paps-app.champion {
          --ink: #05060f;
          --ink-2: #100a26;
          --panel: #14101f;
          --panel-2: #1c1730;
          --line: rgba(255, 213, 74, 0.22);
          --track-red: #FF3B5C;
          --gold: #FFD54A;
          --silver: #E8EDF5;
          --bronze: #E0954B;
          --teal: #2EC4B6;
          --text: #FFF8E7;
          --text-dim: #C9B98A;
          background: radial-gradient(ellipse at top, #241a3d 0%, #05060f 55%, #000 100%);
        }
        .paps-app.champion .panel {
          background: linear-gradient(180deg, rgba(255,213,74,0.07), rgba(255,255,255,0.02));
          border: 1px solid rgba(255,213,74,0.22);
          box-shadow: 0 8px 28px rgba(0,0,0,0.45);
        }
        .paps-app.champion .board-event {
          text-transform: uppercase;
          letter-spacing: 2px;
          background: linear-gradient(180deg, #FFF3C4, #FFD54A 55%, #B8860B);
          -webkit-background-clip: text;
          background-clip: text;
          -webkit-text-fill-color: transparent;
          text-shadow: 0 2px 22px rgba(255,213,74,0.4);
        }
        .paps-app.champion .board-cat { color: var(--text-dim); -webkit-text-fill-color: var(--text-dim); }
        .paps-app.champion .podium-card {
          background: linear-gradient(180deg, rgba(255,213,74,0.09), rgba(255,255,255,0.02));
          border: 1px solid rgba(255,213,74,0.28);
        }
        .paps-app.champion .podium-card.tier-1 {
          animation: champGlow 2.4s ease-in-out infinite;
        }
        @keyframes champGlow {
          0%, 100% { box-shadow: 0 0 28px rgba(255,213,74,0.35), 0 0 0 1px rgba(255,213,74,0.5) inset; }
          50% { box-shadow: 0 0 52px rgba(255,213,74,0.6), 0 0 0 1px rgba(255,213,74,0.85) inset; }
        }
        .paps-app.champion .podium-rank { color: var(--gold); }
        .paps-app.champion .podium-name { text-shadow: 0 0 12px rgba(255,213,74,0.2); }
        .paps-app.champion .grade-dot { box-shadow: 0 2px 10px rgba(0,0,0,0.5), 0 0 10px rgba(255,213,74,0.25); }
        .paps-app.champion .tab.active {
          background: linear-gradient(180deg, rgba(255,213,74,0.18), rgba(255,213,74,0.03));
          color: var(--gold);
          box-shadow: inset 0 -2px 0 var(--gold);
        }
        .paps-app.champion .rank-num { color: var(--gold); font-family: 'Oswald', sans-serif; font-weight: 700; }
        .paps-app.champion .btn-primary {
          background: linear-gradient(180deg, #FFE38A, #FFC93C 60%, #D89B12);
          color: #241a00; font-weight: 700; border: none;
        }
        .paps-app.champion .brand-text { text-shadow: 0 0 16px rgba(255,213,74,0.3); }
        .paps-app.champion .timer-big-number.accent { color: var(--gold); text-shadow: 0 0 20px rgba(255,213,74,0.4); }

        /* ================= 모바일 화면 전용 정리 ================= */
        @media (max-width: 680px) {
          .paps-body { padding: 14px 12px; }
          .presentation .paps-body { padding: 16px 14px; }

          /* 상단 바: 제목을 위쪽 가운데에, 탭은 한 줄로 스와이프, 나머지는 그 아래 깔끔히 정리 */
          .topnav { flex-direction: column; align-items: stretch; gap: 10px; padding: 12px; }
          .brand { justify-content: center; text-align: center; flex-wrap: wrap; gap: 6px 10px; }
          .brand-text { font-size: 15px; }
          .tabs { flex-wrap: wrap; justify-content: center; overflow-x: visible; margin: 0; padding: 0; }
          .tab { padding: 8px 10px; }
          .tab span { font-size: 12px; }
          .nav-right { width: 100%; margin-left: 0; justify-content: center; flex-wrap: wrap; }
          .datetime-row { width: 100%; justify-content: center; }
          .nav-right .btn { flex: 1; justify-content: center; }

          .info-banner { padding: 10px 14px; font-size: 12px; line-height: 1.55; }

          /* 전광판: 종목명·카테고리를 세로로 가운데 정렬해서 한눈에 들어오게 */
          .board-head { flex-direction: column; gap: 10px; text-align: center; }
          .board-title { flex-direction: column; gap: 3px; align-items: center; }
          .board-event { font-size: 22px; }
          .board-cat { display: block; font-size: 13px; }

          .board-filters { gap: 14px; padding: 12px; }
          .filter-group { width: 100%; }
          .chip-row { justify-content: flex-start; }

          .gate-card { padding: 26px 20px; max-width: 100%; }
          .gate-card h2 { font-size: 18px; }

          .panel { padding: 16px; }
          .panel h3 { font-size: 15px; }

          .podium-card { width: 100%; }

          /* 가로로 긴 표는 스크롤 힌트를 붙여 표시 */
          .table-wrap { position: relative; }
          .table-wrap::after {
            content: "← 좌우로 스크롤 →"; display: block; text-align: center; font-size: 10px;
            color: var(--text-dim); padding-top: 6px;
          }

          .modal-panel { padding: 16px; }
          .timer-big-number { font-size: 56px; }
        }
      `}</style>
      {children}
    </div>
  );
}
