// ============================================================
// SPLIT & BUILD — v2.1 (단일) → 세션1 + 세션2 (분리) + 박은희/한용진 실명 케이스 통합
// 실행: node split_and_build.js
// 출력: 오가렌_세일즈가이드_세션1.pptx, 오가렌_세일즈가이드_세션2.pptx
// ============================================================
const fs = require("fs");
const path = require("path");

const SRC = fs.readFileSync(path.join(__dirname, "build_deck_v2.js"), "utf-8");
const PART6_NEW = fs.readFileSync(path.join(__dirname, "part6_revised.js"), "utf-8");

// build_deck_v2.js의 슬라이드 마커 위치 (라인 번호, 1-indexed)
const SLIDE_MARKERS = {
  HEADER_END: 69,         // 라인 1~69: 헤더 (require, 색상 팔레트, 유틸 함수)
  SLIDE_1: 70,            // 표지
  SLIDE_2: 114,           // 사용법
  SLIDE_3: 158,           // 목차
  PART1_START: 203,       // SLIDE 4 디바이더
  PART2_START: 638,       // SLIDE 12 디바이더
  PART3_START: 1082,      // SLIDE 20 디바이더 (+16 from slide 18 인사이트 추가)
  PART4_START: 1304,      // SLIDE 25 디바이더
  PART5_START: 1643,      // SLIDE 32 디바이더 (+10 from slide 30 12분 룰 추가)
  PART6_START: 1930,      // SLIDE 41 디바이더
  PART7_START: 2193,      // SLIDE 48 디바이더
  PART8_START: 2405,      // SLIDE 53 디바이더
  PART9_START: 3108,      // SLIDE 66 디바이더 (+1 from slide 54 rowH 추가)
  APPENDIX_S1_START: 3365, // 제품 라인업 어펜딕스 (세션 1 전용)
  APPENDIX_S2_START: 3392, // 경쟁사 비교 어펜딕스 (세션 2 전용)
  FINAL_START: 3422,       // SLIDE 70 (마무리)
  EOF: -1,
};

// Part 6 내부 분할 (가상 케이스 3개 → 박은희/한용진/5단계 + 매장 벤치마크 + 월간 업데이트로 교체)
const PART6_INTERNAL = {
  DIVIDER: 1930,           // SLIDE 41 디바이더
  BEHAVIORS_7: 1936,       // SLIDE 42: TOP 7가지 행동 패턴 (유지)
  CASE_1: 1990,            // SLIDE 43: 가상 케이스 1 (교체)
  CASE_2: 2069,            // SLIDE 44: 가상 케이스 2 (교체)
  CASE_3: 2087,            // SLIDE 45: 가상 케이스 3 (교체)
  BEST_LINES: 2105,        // SLIDE 46: 베스트 멘트 (유지)
  CHECKLIST: 2150,         // SLIDE 47: 자기 진단 체크리스트 (유지)
  PART6_END: 2193,         // Part 7 시작
};

const lines = SRC.split("\n");
const get = (start, end) => lines.slice(start - 1, end - 1).join("\n") + "\n";

// --- 헤더 (require + 색상 + 유틸 함수) ---
const HEADER = get(1, SLIDE_MARKERS.SLIDE_1);

// --- 헬퍼 함수 정의 (Part 외부에 있는데 세션 분리 시 누락되는 함수들) ---
// makeDivider: SLIDE 4 (Part 1 디바이더) 직전 라인 205-232. 모든 세션에 필요.
// makeObjectionSlide: SLIDE 34 (Part 5) 직전 라인 1670-1729. 세션 1엔 Part 5가 없지만 Part 8에 호출.
const MAKE_DIVIDER_FN = get(205, 233);
const MAKE_OBJECTION_SLIDE_FN = get(1696, 1756);
const HELPER_FNS = MAKE_DIVIDER_FN + "\n" + MAKE_OBJECTION_SLIDE_FN + "\n";

// --- 세션 1 슬라이드 ---
const S1_PART1 = get(SLIDE_MARKERS.PART1_START, SLIDE_MARKERS.PART2_START);
const S1_PART2 = get(SLIDE_MARKERS.PART2_START, SLIDE_MARKERS.PART3_START);
const S1_PART3 = get(SLIDE_MARKERS.PART3_START, SLIDE_MARKERS.PART4_START);
const S1_PART8 = get(SLIDE_MARKERS.PART8_START, SLIDE_MARKERS.PART9_START);

// --- 세션 2 슬라이드 ---
const S2_PART4 = get(SLIDE_MARKERS.PART4_START, SLIDE_MARKERS.PART5_START);
const S2_PART5 = get(SLIDE_MARKERS.PART5_START, SLIDE_MARKERS.PART6_START);
// Part 6: 디바이더 + 7가지 행동 (유지) + part6_revised.js의 박은희/한용진/5단계 + 베스트 멘트 + 체크리스트 + 매장 벤치마크 + 월간 업데이트
const S2_PART6_HEAD = get(PART6_INTERNAL.DIVIDER, PART6_INTERNAL.CASE_1);  // 디바이더 + 7가지 행동
const S2_PART6_TAIL = get(PART6_INTERNAL.BEST_LINES, PART6_INTERNAL.PART6_END);  // 베스트 멘트 + 체크리스트
const S2_PART7 = get(SLIDE_MARKERS.PART7_START, SLIDE_MARKERS.PART8_START);
const S2_PART9 = get(SLIDE_MARKERS.PART9_START, SLIDE_MARKERS.APPENDIX_S1_START);
const APPENDIX_S1 = get(SLIDE_MARKERS.APPENDIX_S1_START, SLIDE_MARKERS.APPENDIX_S2_START);
const APPENDIX_S2 = get(SLIDE_MARKERS.APPENDIX_S2_START, SLIDE_MARKERS.FINAL_START);

// --- 마무리 슬라이드 (writeFile 라인은 제외하고, addSlide 부분만) ---
const FINAL_BLOCK_FULL = lines.slice(SLIDE_MARKERS.FINAL_START - 1).join("\n");
// writeFile 라인 제외
const FINAL_BLOCK = FINAL_BLOCK_FULL.replace(/\n\/\/ ===+\n\/\/ 저장\n\/\/ ===+[\s\S]*$/, "");

// --- part6_revised.js에서 5개 슬라이드 추출 ---
// 박은희 + 한용진 + 5단계 (CASE 3개 자리), 매장 벤치마크 (체크리스트 뒤), 월간 업데이트 (Part 6 마지막)
const PART6_NEW_CASES = (() => {
  // 단순히 part6_revised.js 전체를 사용. 마지막 2개 슬라이드는 별도 분리.
  // 박은희(케이스#1) / 한용진(케이스#2) / 시그니처 5단계 → CASE 3개 교체
  // 매장 벤치마크 → 체크리스트 직후
  // 월간 업데이트 → Part 6 마지막
  return PART6_NEW;  // 전체 5개 슬라이드. 순서대로 들어감.
})();

// 헬퍼: 페이지 번호 자동 매김 (글로벌 카운터)
const AUTO_PAGENUM_PATCH = `
// === [v2.2] 페이지 번호 자동 매김 ===
let __pageCounter = 0;
const __origAddFooter = addFooter;
addFooter = function(slide, _ignoredNum, sectionLabel) {
  __pageCounter += 1;
  __origAddFooter(slide, __pageCounter, sectionLabel);
};
// ==================================
`;

// 헬퍼: 표지 부제목 변경
function patchCover(code, sessionLabel, subtitle, partLabel) {
  // "Edition 2.1 ..." 라인 변경
  code = code.replace(
    /"Edition 2\.[0-9]+[^"]*"/,
    `"${sessionLabel}  ·  ${subtitle}"`
  );
  // "매장 매니저\\n교육 · 세일즈 가이드" → 세션 라벨로
  // (선택적: 표지 메인 텍스트도 변경 가능)
  return code;
}

// 헬퍼: 목차 변경
function patchToc(code, isSession1) {
  if (isSession1) {
    const newToc = `  const toc = [
    ["Part 1", "오가렌과 3대 브랜드", "04"],
    ["Part 2", "매트리스 마스터하기", "12"],
    ["Part 3", "침대 프레임 마스터하기", "20"],
    ["Part 4", "인증 & 안전성", "25"],
    ["부록", "제품 라인업 가격표 (영업 기밀)", "38"],
  ];`;
    code = code.replace(
      /const toc = \[[\s\S]*?\];/,
      newToc
    );
  } else {
    const newToc = `  const toc = [
    ["Part 1", "고객 응대 프로세스", "04"],
    ["Part 2", "세일즈 스크립트 & 거절 처리", "11"],
    ["Part 3", "TOP 매니저 케이스 (실명)", "20"],
    ["Part 4", "운영 · KPI · 부록", "29"],
    ["Part 5", "배송 실무", "34"],
    // ["부록", "경쟁사 비교 자료"] → 세션 3 (쇼룸·백화점 교육자료) 로 이전 (2026-05-22)
  ];`;
    code = code.replace(
      /const toc = \[[\s\S]*?\];/,
      newToc
    );
  }
  return code;
}

// 헬퍼: Part 디바이더의 번호 재매김 (Part 1~9 → Part 1~5)
function renumberDividers(code, mapping) {
  // makeDivider("08", ...) 같은 호출의 번호를 mapping에 따라 변경
  for (const [oldNum, newNum] of Object.entries(mapping)) {
    const oldPattern = new RegExp(`makeDivider\\("${oldNum}",`, "g");
    code = code.replace(oldPattern, `makeDivider("${newNum}",`);
  }
  return code;
}

// 헬퍼: 푸터의 "Part N" 라벨 재매김 (zero-padding 없는 형태 변환)
function renumberFooterLabels(code, mapping) {
  for (const [oldNum, newNum] of Object.entries(mapping)) {
    // mapping은 "06" → "03" 형태이므로, footer는 "6" → "3"로 변환
    const oldUnpadded = parseInt(oldNum, 10);
    const newUnpadded = parseInt(newNum, 10);
    const oldPattern = new RegExp(`"Part ${oldUnpadded} ·`, "g");
    code = code.replace(oldPattern, `"Part ${newUnpadded} ·`);
  }
  return code;
}

// ============================================================
// 세션 1 빌드
// ============================================================
const cover1 = patchCover(get(SLIDE_MARKERS.SLIDE_1, SLIDE_MARKERS.SLIDE_2),
  "세션 1", "우리는 잘 잔다");
const usage = get(SLIDE_MARKERS.SLIDE_2, SLIDE_MARKERS.SLIDE_3);
const toc1 = patchToc(get(SLIDE_MARKERS.SLIDE_3, SLIDE_MARKERS.PART1_START), true);

const S1_MAPPING = { "01": "01", "02": "02", "03": "03", "08": "04" };
let session1 = HEADER
  + HELPER_FNS  // ← makeDivider + makeObjectionSlide 함수 정의
  + AUTO_PAGENUM_PATCH
  + cover1
  + usage
  + toc1
  + S1_PART1
  + S1_PART2
  + S1_PART3
  + S1_PART8
  + APPENDIX_S1
  + FINAL_BLOCK;
session1 = renumberDividers(session1, S1_MAPPING);
session1 = renumberFooterLabels(session1, S1_MAPPING);
session1 = session1.replace(
  /pres\.writeFile\(\{ fileName: "[^"]*" \}\)/,
  `pres.writeFile({ fileName: "오가렌_세일즈가이드_세션1.pptx" })`
);
// 만약 writeFile이 위에서 잘렸으면 직접 추가
if (!session1.includes("pres.writeFile")) {
  session1 += `\n\npres.writeFile({ fileName: "오가렌_세일즈가이드_세션1.pptx" })\n  .then(fn => console.log("DONE:", fn))\n  .catch(err => { console.error("ERROR:", err); process.exit(1); });\n`;
}

fs.writeFileSync(path.join(__dirname, "build_session1.js"), session1, "utf-8");
console.log("✓ build_session1.js 생성 (", session1.split("\n").length, "lines)");

// ============================================================
// 세션 2 빌드
// ============================================================
const cover2 = patchCover(get(SLIDE_MARKERS.SLIDE_1, SLIDE_MARKERS.SLIDE_2),
  "세션 2", "그래서 잘 안내한다");
const toc2 = patchToc(get(SLIDE_MARKERS.SLIDE_3, SLIDE_MARKERS.PART1_START), false);

const S2_MAPPING = { "04": "01", "05": "02", "06": "03", "07": "04", "09": "05" };
let session2 = HEADER
  + HELPER_FNS  // ← makeDivider + makeObjectionSlide 함수 정의
  + AUTO_PAGENUM_PATCH
  + cover2
  + usage
  + toc2
  + S2_PART4
  + S2_PART5
  // Part 6: 디바이더 + 7가지 행동 + 박은희/한용진/5단계 (NEW) + 베스트 멘트 + 체크리스트
  + S2_PART6_HEAD
  + PART6_NEW_CASES  // ← part6_revised.js 전체 (박은희, 한용진, 5단계, 매장 벤치마크, 월간 업데이트)
  + S2_PART6_TAIL
  + S2_PART7
  + S2_PART9
  + APPENDIX_S2
  + FINAL_BLOCK;
session2 = renumberDividers(session2, S2_MAPPING);
session2 = renumberFooterLabels(session2, S2_MAPPING);
session2 = session2.replace(
  /pres\.writeFile\(\{ fileName: "[^"]*" \}\)/,
  `pres.writeFile({ fileName: "오가렌_세일즈가이드_세션2.pptx" })`
);
if (!session2.includes("pres.writeFile")) {
  session2 += `\n\npres.writeFile({ fileName: "오가렌_세일즈가이드_세션2.pptx" })\n  .then(fn => console.log("DONE:", fn))\n  .catch(err => { console.error("ERROR:", err); process.exit(1); });\n`;
}

fs.writeFileSync(path.join(__dirname, "build_session2.js"), session2, "utf-8");
console.log("✓ build_session2.js 생성 (", session2.split("\n").length, "lines)");

// ============================================================
// 실제 PPT 빌드 실행
// ============================================================
console.log("\nPPT build start");
const { execSync } = require("child_process");
try {
  execSync("node build_session1.js", { cwd: __dirname, stdio: "inherit" });
  execSync("node build_session2.js", { cwd: __dirname, stdio: "inherit" });
  console.log("All builds done");
} catch (e) {
  console.error("Build failed:", e.message);
}
