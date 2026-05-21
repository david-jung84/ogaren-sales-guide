// 오가렌 매니저 교육 및 세일즈 가이드 - PPT 생성 스크립트
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.author = "오가렌 경영기획본부";
pres.company = "주식회사 오가렌";
pres.title = "오가렌 매장 매니저 교육·세일즈 가이드 v2.0";

// ============================================================
// 컬러 팔레트 (프리미엄 침대 브랜드 톤)
// ============================================================
const C = {
  ink: "2D2926",           // 진한 차콜 (제목/본문)
  inkSoft: "5B544C",       // 부드러운 차콜 (보조 텍스트)
  muted: "8E867C",         // 흐린 회색 (캡션)
  gold: "B8956A",          // 따뜻한 골드 (포인트)
  goldDeep: "8E6E47",      // 진한 골드 (강조)
  cream: "F4EFE6",         // 크림 (섹션 배경)
  creamSoft: "FAF6EF",     // 더 옅은 크림
  white: "FFFFFF",
  line: "DAD2C4",          // 라인 / 구분선
  alert: "A8413E",         // 강조 (경고/포인트)
  blue: "3B5F7D",          // 보조 컬러
};

const FONT_H = "Malgun Gothic";  // 헤더용 - Windows 한글 기본
const FONT_B = "Malgun Gothic";  // 본문용

// ============================================================
// 유틸리티 - 슬라이드 헬퍼
// ============================================================
function addFooter(slide, pageNum, sectionLabel) {
  // 페이지 번호
  slide.addText(String(pageNum), {
    x: 9.3, y: 5.32, w: 0.4, h: 0.25,
    fontSize: 9, color: C.muted, fontFace: FONT_B,
    align: "right", margin: 0
  });
  // 섹션 라벨
  if (sectionLabel) {
    slide.addText(sectionLabel, {
      x: 0.5, y: 5.32, w: 6, h: 0.25,
      fontSize: 9, color: C.muted, fontFace: FONT_B,
      align: "left", margin: 0
    });
  }
  // 브랜드
  slide.addText("OGAREN  ·  SLEEPER  ·  nooer  ·  Toddles", {
    x: 3.5, y: 5.32, w: 5.7, h: 0.25,
    fontSize: 8, color: C.muted, fontFace: FONT_B,
    align: "right", margin: 0, charSpacing: 2
  });
}

function addSectionTitle(slide, kicker, title) {
  slide.addText(kicker, {
    x: 0.5, y: 0.32, w: 9, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_B,
    bold: true, charSpacing: 3, margin: 0
  });
  slide.addText(title, {
    x: 0.5, y: 0.62, w: 9, h: 0.6,
    fontSize: 26, color: C.ink, fontFace: FONT_H,
    bold: true, margin: 0
  });
}

// ============================================================
function makeDivider(num, partTitle, subtitle) {
  const s = pres.addSlide();
  s.background = { color: C.ink };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.4, w: 0.6, h: 0.06,
    fill: { color: C.gold }, line: { color: C.gold }
  });

  s.addText(`PART ${num}`, {
    x: 0.5, y: 1.9, w: 9, h: 0.4,
    fontSize: 14, color: C.gold, fontFace: FONT_B,
    bold: true, charSpacing: 5, margin: 0
  });

  s.addText(partTitle, {
    x: 0.5, y: 2.6, w: 9, h: 1.0,
    fontSize: 42, color: C.white, fontFace: FONT_H,
    bold: true, margin: 0
  });

  s.addText(subtitle, {
    x: 0.5, y: 3.7, w: 9, h: 0.5,
    fontSize: 15, color: C.line, fontFace: FONT_B,
    italic: true, margin: 0
  });
  return s;
}

function makeObjectionSlide(num, page, objection, kicker, badAns, goodAns, why) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, `OBJECTION #${num}`, `"${objection}"`);

  s.addText(kicker, {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 13, color: C.gold, fontFace: FONT_H,
    italic: true, bold: true, margin: 0
  });

  // 나쁜 답 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.0, w: 4.3, h: 1.5,
    fill: { color: C.creamSoft }, line: { color: C.alert, width: 1 }
  });
  s.addText("NG  ·  이렇게는 NO", {
    x: 0.65, y: 2.05, w: 4.0, h: 0.3,
    fontSize: 11, color: C.alert, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(badAns, {
    x: 0.65, y: 2.4, w: 4.0, h: 1.05,
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B,
    italic: true, margin: 0
  });

  // 좋은 답 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 2.0, w: 4.3, h: 1.5,
    fill: { color: C.cream }, line: { color: C.gold, width: 1 }
  });
  s.addText("OK  ·  이렇게 답하세요", {
    x: 5.35, y: 2.05, w: 4.0, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(goodAns, {
    x: 5.35, y: 2.4, w: 4.0, h: 1.05,
    fontSize: 11, color: C.ink, fontFace: FONT_B,
    italic: true, margin: 0
  });

  // 왜?
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.75, w: 9, h: 1.35,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("WHY  ·  이 답이 통하는 이유", {
    x: 0.7, y: 3.85, w: 8.5, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(why, {
    x: 0.7, y: 4.2, w: 8.5, h: 0.9,
    fontSize: 11, color: C.white, fontFace: FONT_B, margin: 0
  });

  addFooter(s, page, `Part 5 · 거절 #${num}`);
}


// === [v2.2] 페이지 번호 자동 매김 ===
let __pageCounter = 0;
const __origAddFooter = addFooter;
addFooter = function(slide, _ignoredNum, sectionLabel) {
  __pageCounter += 1;
  __origAddFooter(slide, __pageCounter, sectionLabel);
};
// ==================================
// SLIDE 1 — 표지
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.ink };

  // 골드 액센트 사각형 (왼쪽 상단)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.5, w: 0.5, h: 0.06,
    fill: { color: C.gold }, line: { color: C.gold }
  });

  s.addText("OGAREN  EDUCATION  PROGRAM", {
    x: 0.5, y: 0.7, w: 9, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_B,
    bold: true, charSpacing: 4, margin: 0
  });

  s.addText("매장 매니저\n교육 · 세일즈 가이드", {
    x: 0.5, y: 1.7, w: 9, h: 2.0,
    fontSize: 54, color: C.white, fontFace: FONT_H,
    bold: true, margin: 0
  });

  s.addText("백화점 · 쇼룸 매니저를 위한 통합 플레이북", {
    x: 0.5, y: 3.7, w: 9, h: 0.5,
    fontSize: 18, color: C.line, fontFace: FONT_B,
    margin: 0
  });

  s.addText("SLEEPER  ·  nooer  ·  Toddles", {
    x: 0.5, y: 4.6, w: 9, h: 0.3,
    fontSize: 13, color: C.gold, fontFace: FONT_B,
    bold: true, charSpacing: 5, margin: 0
  });

  s.addText("세션 2  ·  그래서 잘 판다", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: FONT_B,
    margin: 0
  });
}

// ============================================================
// SLIDE 2 — 본 가이드 사용법
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "HOW TO USE", "이 가이드, 이렇게 활용하세요");

  // 카드 4개 (2x2)
  const cards = [
    { num: "01", title: "신입 온보딩 3주차 교재", desc: "Part 1~3을 기본 교재로 활용.\n주차별 학습 → 시험 → 단독 응대 자격" },
    { num: "02", title: "매장 즉시 사용 무기", desc: "Part 4~5는 응대·거절 처리 스크립트.\n출근 후 10분, 응대 직전 빠른 복습" },
    { num: "03", title: "월간 워크샵 자료", desc: "Part 6 TOP 매니저 케이스로\n월 1회 롤플레잉 세션 진행" },
    { num: "04", title: "정기 업데이트", desc: "분기마다 신제품·시즌·\n베스트 사례 보강 후 재배포" },
  ];
  cards.forEach((c, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6, y = 1.55 + row * 1.7;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 1.5,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addText(c.num, {
      x: x + 0.25, y: y + 0.18, w: 0.8, h: 0.4,
      fontSize: 22, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(c.title, {
      x: x + 1.0, y: y + 0.2, w: 3.1, h: 0.4,
      fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(c.desc, {
      x: x + 0.25, y: y + 0.7, w: 3.85, h: 0.7,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("Tip — 가이드는 \"공부할 책\"이 아니라 \"꺼내 쓸 도구\"입니다. 외우려 하지 말고, 필요할 때 찾으세요.", {
    x: 0.5, y: 4.95, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 2, "사용 안내");
}

// ============================================================
// SLIDE 3 — 목차
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "CONTENTS", "목차");

    const toc = [
    ["Part 1", "고객 응대 프로세스", "04"],
    ["Part 2", "세일즈 스크립트 & 거절 처리", "11"],
    ["Part 3", "TOP 매니저 케이스 (실명)", "20"],
    ["Part 4", "운영 · KPI · 부록", "29"],
    ["Part 5", "배송 실무", "34"],
  ];
  toc.forEach((row, i) => {
    const y = 1.5 + i * 0.45;
    s.addText(row[0], {
      x: 0.7, y, w: 1.2, h: 0.4,
      fontSize: 12, color: C.gold, fontFace: FONT_B, bold: true,
      charSpacing: 2, margin: 0
    });
    s.addText(row[1], {
      x: 2.0, y, w: 6.5, h: 0.4,
      fontSize: 17, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(row[2], {
      x: 8.5, y, w: 0.8, h: 0.4,
      fontSize: 13, color: C.muted, fontFace: FONT_B,
      align: "right", margin: 0
    });
    // 구분선
    s.addShape(pres.shapes.LINE, {
      x: 0.7, y: y + 0.4, w: 8.6, h: 0,
      line: { color: C.line, width: 0.5 }
    });
  });

  addFooter(s, 3, "목차");
}

// ============================================================
// SLIDE 25 — Part 4 디바이더
// ============================================================
makeDivider("01", "고객 응대 프로세스",
  "입장부터 클로징까지 — 5단계로 흐름을 잡으세요.");

// ============================================================
// SLIDE 26 — 침대 매장 고객 5가지 페르소나
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "5 PERSONAS", "매장에 들어오는 5가지 고객 유형");

  const ps = [
    { tag: "A", title: "결혼 준비형", sign: "\"신혼 침대 보러 왔어요\"", who: "20대 후반~30대\n부부 동반 90%", play: "스토리텔링 · 패브릭 침대 + 시그니처 매트리스 풀세트 제안" },
    { tag: "B", title: "이사 / 신축형", sign: "\"방에 맞는 게 있을까요?\"", who: "다양 연령 · 공간 명확", play: "방 사이즈·방향 먼저 묻기. 1:1 맞춤 강점 강조" },
    { tag: "C", title: "건강·통증형", sign: "\"허리가 안 좋아서…\"", who: "40대 이상 · 단독 방문 多", play: "전문가 모드 ON. 체압 분산·라텍스·고밀도 메모리폼 설명" },
    { tag: "D", title: "비교 쇼핑형", sign: "\"시몬스도 보고 왔는데요\"", who: "정보 무장, 까다로움", play: "차이점 솔직 비교 · 1:1 맞춤 / 30일 트라이얼 USP 강조" },
    { tag: "E", title: "둘러보기형", sign: "\"그냥 한 번 보러…\"", who: "구매 의향 낮아 보이지만 잠재", play: "압박 NO. 체험만 권유 · 명함·카탈로그 챙겨드리기" },
  ];

  ps.forEach((p, i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.45 + Math.floor(i / 3) * 1.85;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 1.7,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    // 태그 원
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.15, y: y + 0.15, w: 0.55, h: 0.55,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(p.tag, {
      x: x + 0.15, y: y + 0.15, w: 0.55, h: 0.55,
      fontSize: 16, color: C.gold, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(p.title, {
      x: x + 0.85, y: y + 0.15, w: 1.95, h: 0.3,
      fontSize: 13, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(p.sign, {
      x: x + 0.85, y: y + 0.45, w: 1.95, h: 0.3,
      fontSize: 10, color: C.gold, fontFace: FONT_B, italic: true, margin: 0
    });
    s.addText("WHO  " + p.who, {
      x: x + 0.15, y: y + 0.8, w: 2.6, h: 0.35,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
    s.addText("PLAY  " + p.play, {
      x: x + 0.15, y: y + 1.15, w: 2.6, h: 0.5,
      fontSize: 9.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
    });
  });

  s.addText("Tip — 5가지 유형 중 \"E. 둘러보기형\"이 사실 가장 큰 기회. 압박 없는 응대 + 명함은 재방문률을 가장 크게 높입니다.", {
    x: 0.5, y: 5.1, w: 9, h: 0.3,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 26, "Part 1 · 페르소나");
}

// ============================================================
// SLIDE 27 — 응대 5단계 프로세스
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "5-STEP PROCESS", "고객 응대 5단계 — 매장 표준 흐름");

  const steps = [
    { n: "1", t: "환영", d: "3초 미소 · 3분 자율 · 가벼운 한 마디", goal: "편안한 첫인상" },
    { n: "2", t: "니즈 파악", d: "공간 · 인원 · 자세 · 기간 4질문", goal: "고객을 이해" },
    { n: "3", t: "제안", d: "매트리스 + 프레임 묶음 추천", goal: "선택지 좁히기" },
    { n: "4", t: "체험", d: "직접 누워보기 · 토퍼 비교", goal: "감각으로 확신" },
    { n: "5", t: "클로징", d: "옵션 확정 · 결제 · A/S 안내", goal: "재방문·추천 약속" },
  ];

  steps.forEach((st, i) => {
    const x = 0.5 + i * 1.85, y = 1.55;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 1.75, h: 2.9,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.6, y: y + 0.2, w: 0.55, h: 0.55,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(st.n, {
      x: x + 0.6, y: y + 0.2, w: 0.55, h: 0.55,
      fontSize: 18, color: C.white, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(st.t, {
      x: x + 0.1, y: y + 0.95, w: 1.55, h: 0.35,
      fontSize: 14, color: C.ink, fontFace: FONT_H,
      bold: true, align: "center", margin: 0
    });
    s.addText(st.d, {
      x: x + 0.1, y: y + 1.35, w: 1.55, h: 1.0,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B,
      align: "center", margin: 0
    });
    s.addText(st.goal, {
      x: x + 0.1, y: y + 2.4, w: 1.55, h: 0.4,
      fontSize: 10, color: C.goldDeep, fontFace: FONT_B,
      bold: true, italic: true, align: "center", margin: 0
    });
  });

  s.addText("핵심 — 5단계는 \"순서\"이지만 정해진 시간이 없습니다. 단계 사이에 머무르는 시간이 신뢰를 결정합니다.", {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 27, "Part 1 · 5단계 응대");
}

// ============================================================
// SLIDE 28 — 입장 3분 룰 (좋은 첫 멘트 5종)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "FIRST 3 MINUTES", "입장 3분 룰 — 첫 멘트 5종 세트");

  s.addText("입장 즉시 달려가지 마세요. 3분간 자율 관람 → 매니저는 자연스럽게 동선 안에서 등장.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const lines = [
    { good: "\"천천히 둘러보세요. 마음 가시는 제품 있으면 한 번 누워보셔도 됩니다.\"", note: "체험 권유 + 압박 제로" },
    { good: "\"오늘 어떤 제품 보러 오셨어요? 신혼이세요, 이사세요?\"", note: "페르소나 1차 파악" },
    { good: "\"날씨가 더운데 들어오시느라 고생하셨어요. 차 한 잔 드릴까요?\"", note: "관계 형성 우선" },
    { good: "\"부부 함께 오셨으니까 두 분 같이 누워보실 수 있는 모델 위주로 안내드릴게요.\"", note: "동반자 인지 시 효과적" },
    { good: "\"혹시 지금 쓰시는 침대에서 가장 불편하신 점 있으세요?\"", note: "니즈 파악 직진" },
  ];

  lines.forEach((l, i) => {
    const y = 2.0 + i * 0.55;
    s.addText(`${i + 1}`, {
      x: 0.5, y, w: 0.4, h: 0.4,
      fontSize: 18, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(l.good, {
      x: 0.95, y, w: 5.9, h: 0.4,
      fontSize: 12, color: C.ink, fontFace: FONT_H, italic: true, margin: 0
    });
    s.addText(l.note, {
      x: 6.95, y, w: 2.5, h: 0.4,
      fontSize: 10, color: C.goldDeep, fontFace: FONT_B, margin: 0
    });
  });

  // 하단 NO 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fill: { color: C.alert, transparency: 85 }, line: { color: C.alert, width: 0.5 }
  });
  s.addText("NO  ·  \"찾으시는 거 있으세요?\" / \"이게 잘 나가요\" / \"오늘만 할인이에요\" — 첫 멘트로 가장 피해야 할 3대 NG", {
    x: 0.7, y: 4.85, w: 8.6, h: 0.4,
    fontSize: 10.5, color: C.alert, fontFace: FONT_B, bold: true, valign: "middle", margin: 0
  });

  addFooter(s, 28, "Part 1 · 입장 3분");
}

// ============================================================
// SLIDE 29 — 니즈 파악 4질문 프레임
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "NEEDS DISCOVERY", "니즈 파악 4질문 — 외워두세요");

  s.addText("이 4가지만 자연스럽게 물어보면, 추천이 거의 정해집니다.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 13, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const qs = [
    {
      n: "Q1", t: "공간",
      ask: "\"침실 방 크기는 어느 정도 되세요?\"",
      learn: "사이즈(Q/K/LK) 결정\n수납 침대 여부"
    },
    {
      n: "Q2", t: "동침",
      ask: "\"혼자 쓰세요, 같이 쓰세요?\"",
      learn: "1인 vs 부부 모델 분기\n모션 분리 중요도"
    },
    {
      n: "Q3", t: "자세 · 불편",
      ask: "\"보통 어떤 자세로 주무세요? 어디 불편한 곳 있으세요?\"",
      learn: "매트리스 단단함\n메모리폼/라텍스 분기"
    },
    {
      n: "Q4", t: "사용 주기",
      ask: "\"지금 침대는 몇 년 정도 쓰셨어요?\"",
      learn: "교체 시급도\n예산 감각 자연스럽게"
    },
  ];

  qs.forEach((q, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 2.0 + row * 1.5;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 1.35,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addText(q.n, {
      x: x + 0.2, y: y + 0.1, w: 0.6, h: 0.4,
      fontSize: 16, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(q.t, {
      x: x + 0.8, y: y + 0.13, w: 3.4, h: 0.35,
      fontSize: 13, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(q.ask, {
      x: x + 0.2, y: y + 0.5, w: 4.0, h: 0.4,
      fontSize: 11.5, color: C.goldDeep, fontFace: FONT_H, italic: true, margin: 0
    });
    s.addText("→ " + q.learn, {
      x: x + 0.2, y: y + 0.92, w: 4.0, h: 0.4,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 29, "Part 1 · 4질문");
}

// ============================================================
// SLIDE 30 — 체험 유도 멘트
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "TRY IT", "체험(누워보기)을 부드럽게 유도하는 법");

  s.addText("침대 매장의 진실 — 누워본 고객의 구매 전환율이 안 누워본 고객의 3배.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 14, color: C.goldDeep, fontFace: FONT_H, bold: true, italic: true, margin: 0
  });

  const mentions = [
    { stage: "초기 진입", line: "\"먼저 한 번 누워보세요. 누워봐야 진짜 차이가 느껴지거든요.\"" },
    { stage: "비교 중", line: "\"두 모델 같은 자세로 30초씩 누워보세요. 몸이 알아요.\"" },
    { stage: "부부 동반", line: "\"두 분 같이 누워보세요. 모션 분리도 직접 느끼실 수 있어요.\" (코인 데모)" },
    { stage: "주저할 때", line: "\"신발만 벗어두시면 돼요. 베개도 가져다 드릴게요.\"" },
    { stage: "체험 후", line: "\"어느 쪽이 더 편하셨어요? 어떤 느낌이 더 좋으셨어요?\"" },
  ];

  mentions.forEach((m, i) => {
    const y = 1.95 + i * 0.55;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 1.85, h: 0.45,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(m.stage, {
      x: 0.5, y, w: 1.85, h: 0.45,
      fontSize: 11, color: C.white, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 2.45, y, w: 7.05, h: 0.45,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addText(m.line, {
      x: 2.6, y, w: 6.85, h: 0.45,
      fontSize: 11.5, color: C.ink, fontFace: FONT_B,
      italic: true, valign: "middle", margin: 0
    });
  });

  s.addText("Tip — 체험 중에는 가급적 말을 줄이세요. 고객이 자기 몸의 신호를 듣는 시간입니다.", {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 30, "Part 1 · 체험 유도");
}

// ============================================================
// SLIDE 31 — 부부 동반 시 응대 전략
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "COUPLE SCENE", "부부 동반 — 실전 응대 전략");

  s.addText("두 사람을 한 명처럼 응대하면 실패. 한 명씩 \"보이게\" 응대하면 성공.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 13, color: C.goldDeep, fontFace: FONT_H, italic: true, bold: true, margin: 0
  });

  const rules = [
    { r: "두 사람 모두에게 눈맞춤", d: "한 사람만 보지 말기. 번갈아가며 자연스럽게 시선·질문." },
    { r: "의사결정자 파악", d: "초반 5분이면 누가 \"실제 결정자\"인지 보입니다. 그 사람에게 핵심 설명을 더." },
    { r: "두 사람 다 누워보게", d: "한 명만 누워서 결정 NO. 반드시 두 명 다 체험 → 자기 의견 말하게." },
    { r: "서로의 의견 들어주기", d: "\"아내 분은 어떠셨어요?\", \"남편 분이 좋아하시는 단단함은요?\" 균형 잡힌 질문." },
    { r: "결정 압박 NO", d: "\"오늘 결정 안 하셔도 돼요. 두 분 더 얘기 나누고 오세요.\" 재방문률 높이는 안전한 멘트." },
  ];

  rules.forEach((rl, i) => {
    const y = 1.95 + i * 0.55;
    s.addText(`${i + 1}`, {
      x: 0.5, y, w: 0.4, h: 0.5,
      fontSize: 20, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(rl.r, {
      x: 0.95, y, w: 3.0, h: 0.3,
      fontSize: 12, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(rl.d, {
      x: 4.0, y, w: 5.45, h: 0.5,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, valign: "top", margin: 0
    });
    s.addShape(pres.shapes.LINE, {
      x: 0.5, y: y + 0.5, w: 9.0, h: 0,
      line: { color: C.line, width: 0.5 }
    });
  });

  addFooter(s, 31, "Part 1 · 부부 응대");
}

// ============================================================
// SLIDE 32 — Part 5 디바이더
// ============================================================
makeDivider("02", "세일즈 스크립트 & 거절 처리",
  "고객이 망설이는 모든 순간을 위한 답이 여기 있습니다.");

// ============================================================
// SLIDE 33 — 거절 처리 TOP 7 (인덱스)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "TOP 7 OBJECTIONS", "고객이 자주 하는 거절 7가지");

  const objs = [
    "비싸요 / 예산이 그 정도는…",
    "좀 더 알아보고 올게요",
    "온라인이 더 싸요",
    "남편/아내랑 상의해볼게요",
    "다른 브랜드(시몬스, 에이스 등)도 보고 싶어요",
    "이사할 때 다시 올게요",
    "지금 쓰는 거 아직 멀쩡해요",
  ];

  objs.forEach((o, i) => {
    const y = 1.45 + i * 0.5;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 0.42,
      fill: { color: i % 2 === 0 ? C.creamSoft : C.white },
      line: { color: C.line, width: 0.5 }
    });
    s.addText(`#${i + 1}`, {
      x: 0.7, y, w: 0.6, h: 0.42,
      fontSize: 14, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
    s.addText(`"${o}"`, {
      x: 1.3, y, w: 7.0, h: 0.42,
      fontSize: 13, color: C.ink, fontFace: FONT_B,
      italic: true, valign: "middle", margin: 0
    });
    s.addText("→ 다음 페이지", {
      x: 8.3, y, w: 1.1, h: 0.42,
      fontSize: 9, color: C.muted, fontFace: FONT_B,
      align: "right", valign: "middle", margin: 0
    });
  });

  addFooter(s, 33, "Part 2 · 거절 TOP 7");
}

// ============================================================
// SLIDE 34 — 거절 #1 "비싸요"
// ============================================================
function makeObjectionSlide(num, page, objection, kicker, badAns, goodAns, why) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, `OBJECTION #${num}`, `"${objection}"`);

  s.addText(kicker, {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 13, color: C.gold, fontFace: FONT_H,
    italic: true, bold: true, margin: 0
  });

  // 나쁜 답 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.0, w: 4.3, h: 1.5,
    fill: { color: C.creamSoft }, line: { color: C.alert, width: 1 }
  });
  s.addText("NG  ·  이렇게는 NO", {
    x: 0.65, y: 2.05, w: 4.0, h: 0.3,
    fontSize: 11, color: C.alert, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(badAns, {
    x: 0.65, y: 2.4, w: 4.0, h: 1.05,
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B,
    italic: true, margin: 0
  });

  // 좋은 답 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 2.0, w: 4.3, h: 1.5,
    fill: { color: C.cream }, line: { color: C.gold, width: 1 }
  });
  s.addText("OK  ·  이렇게 답하세요", {
    x: 5.35, y: 2.05, w: 4.0, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(goodAns, {
    x: 5.35, y: 2.4, w: 4.0, h: 1.05,
    fontSize: 11, color: C.ink, fontFace: FONT_B,
    italic: true, margin: 0
  });

  // 왜?
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.75, w: 9, h: 1.35,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("WHY  ·  이 답이 통하는 이유", {
    x: 0.7, y: 3.85, w: 8.5, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(why, {
    x: 0.7, y: 4.2, w: 8.5, h: 0.9,
    fontSize: 11, color: C.white, fontFace: FONT_B, margin: 0
  });

  addFooter(s, page, `Part 5 · 거절 #${num}`);
}

makeObjectionSlide(
  1, 34, "비싸요 / 예산이 그 정도는…",
  "가격을 깎지 마세요. 가치의 단위를 바꾸세요.",
  "\"많이 비싸진 않아요\"\n\"이번 달 할인이라…\"\n→ 가격 방어는 가격 인정. 더 깎아야 한다는 메시지만 줍니다.",
  "\"하루로 나누면 약 X백 원이세요. 8~10년 매일 8시간 쓰시니까요.\"\n또는\n\"이 침대는 8년 보증입니다. 매일 밤 좋은 잠을 보장받는 비용이라고 생각하시면 어떠세요?\"",
  "고가 제품은 \"비교 단위\"를 일 단위·시간 단위로 나누면 인식이 바뀝니다.\n또한 \"수면 = 건강 = 일·관계의 질\"로 가치의 차원을 옮기면, \"가격\"이 아니라 \"투자\"로 보입니다.\n절대 깎으려 하지 말고, 가치를 다시 설명하세요."
);

// ============================================================
// SLIDE 35 — 거절 #2 "더 알아볼게요"
// ============================================================
makeObjectionSlide(
  2, 35, "좀 더 알아보고 올게요",
  "막지 마세요. 돌아올 이유를 만들어주세요.",
  "\"오늘만 가격이에요\"\n\"지금 안 사시면…\"\n→ 압박은 도망갈 명분만 줍니다. 거의 안 돌아옵니다.",
  "\"천천히 보세요. 그런데 오늘 누워본 모델 두 가지만 적어드릴게요. 나중에 비교하실 때 쉬워요.\"\n+ 명함 + 매장 카탈로그 챙겨드리기",
  "고객이 \"알아본다\"는 것은 \"아직 확신이 없다\"는 신호입니다.\n오늘 정보를 정리해주면, 다른 매장 가서도 \"오가렌이 더 친절했네\" 라는 기억이 남습니다.\n실제 TOP 매니저들의 재방문 클로징 비율은 30~40%. 떠나는 고객이 진짜 고객입니다."
);

// ============================================================
// SLIDE 36 — 거절 #3 "온라인이 더 싸요"
// ============================================================
makeObjectionSlide(
  3, 36, "온라인이 더 싸요",
  "가격 싸움을 하지 마세요. \"체험의 가치\"로 전환하세요.",
  "\"저희도 온라인 가격 맞춰드려요\"\n→ 가격으로 시작한 거래는 가격으로 끝납니다. AS·재구매 안 옵니다.",
  "\"맞아요, 온라인이 저렴할 수도 있어요. 그런데 매트리스는 직접 누워보지 않으면 모르거든요.\n오늘 누워보신 게 진짜 맞는 모델이라면, 저희 매장에서 사시는 게 가장 안전해요. A/S도 가까이서 받으시고요.\"",
  "온라인 vs 오프라인은 가격 비교가 아니라 \"리스크 비교\"입니다.\n8~10년 쓸 제품을 못 누워보고 사는 것의 리스크를 인지시키세요.\n그리고 \"내가 직접 봐드린다\"는 인간적 신뢰의 가치를 강조하세요."
);

// ============================================================
// SLIDE 37 — 거절 #4 "상의해볼게요"
// ============================================================
makeObjectionSlide(
  4, 37, "남편/아내랑 상의해볼게요",
  "상의를 막지 말고, 상의가 잘되게 도와주세요.",
  "\"전화로 같이 통화해보시죠\"\n\"오늘 같이 결정하시는 게 어때요\"\n→ 결혼 생활을 모르는 것처럼 보입니다.",
  "\"네, 두 분이 같이 누워보셔야죠. 같이 한 번 더 오실 때까지 이 모델 빼두지 않을게요.\n그리고 이거 한 장 챙겨가세요 — 두 모델 비교표예요. 상의하실 때 도움 되실 거예요.\"",
  "고객이 \"상의\"를 말할 때는 결정권을 \"같이\" 행사하고 싶다는 신호입니다.\n그걸 존중해 주는 매니저가 결국 두 번째 방문을 받습니다.\n비교표·카탈로그를 챙겨주면, 집에서 상의할 때 우리 제품이 \"기준\"이 됩니다."
);

// ============================================================
// SLIDE 38 — 거절 #5 "다른 브랜드도 봐야"
// ============================================================
makeObjectionSlide(
  5, 38, "다른 브랜드(시몬스 등)도 보고 싶어요",
  "경쟁사를 깎아내리지 마세요. 우리의 본질을 보여주세요.",
  "\"거기는 스프링이 별로예요\"\n\"브랜드 마케팅만 비싸요\"\n→ 경쟁사 비방은 자신감 부족으로 보입니다.",
  "\"좋은 생각이세요. 시몬스도 좋은 브랜드죠. 다만 저희 차이는 \"1:1 맞춤\"이에요. 시몬스는 기성 사이즈, 저희는 고객님 침실에 맞춰 만들어 드립니다. 그리고 누어 매트리스는 30일 트라이얼이 있어서 실패 부담이 없어요.\"",
  "고객이 비교하겠다는 건 건강한 신호입니다. 막을 수 없고, 막을 필요도 없습니다.\n경쟁사 인정 → 우리 USP 한 줄(1:1 맞춤 / 30일 트라이얼 / 사후 케어) → \"비교해 보고 결정하세요\" 의 흐름이 가장 강력합니다.\n자신감 있는 매니저만이 이렇게 답할 수 있습니다."
);

// ============================================================
// SLIDE 39 — 거절 #6 / #7
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "OBJECTIONS #6, #7", "이사 / 멀쩡함 — 거절 처리");

  // #6 이사할 때
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 9, h: 1.8,
    fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
  });
  s.addText("#6  \"이사할 때 다시 올게요\"", {
    x: 0.7, y: 1.5, w: 8.6, h: 0.35,
    fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText(
    "→ \"이사 일정이 언제세요? 침대는 보통 이사 2~4주 전에 주문하셔야 입주일에 받으실 수 있어요.\n오늘 결정 안 하셔도 되는데, 일정만 알려주시면 제가 챙겨서 연락드릴게요.\"",
    {
      x: 0.7, y: 1.9, w: 8.6, h: 0.7,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
    }
  );
  s.addText("WHY — \"이사할 때 다시\"는 90% 안 옵니다. 연락처를 얻고 능동적으로 추적하세요. CRM 입력 필수.", {
    x: 0.7, y: 2.65, w: 8.6, h: 0.5,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, italic: true, margin: 0
  });

  // #7 멀쩡함
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.4, w: 9, h: 1.8,
    fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
  });
  s.addText("#7  \"지금 쓰는 거 아직 멀쩡해요\"", {
    x: 0.7, y: 3.5, w: 8.6, h: 0.35,
    fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText(
    "→ \"몇 년 쓰셨어요? … 7년이면 거의 교체 시점이세요. 매트리스는 안 꺼져도 위생적으로 진드기·먼지가 누적되거든요.\n저희 누어 든든 토퍼만 얹어도 완전히 다른 잠을 느끼세요. 매트리스 교체보다 부담 적고요.\"",
    {
      x: 0.7, y: 3.9, w: 8.6, h: 0.7,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
    }
  );
  s.addText("WHY — \"멀쩡함\"은 \"바꿀 이유 없음\"이 아니라 \"바꿀 동기 부족\". 위생·토퍼 업셀로 부드럽게 진입.", {
    x: 0.7, y: 4.65, w: 8.6, h: 0.5,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, italic: true, margin: 0
  });

  addFooter(s, 39, "Part 2 · 거절 #6 #7");
}

// ============================================================
// SLIDE 40 — 클로징 멘트 3종
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "CLOSING", "클로징 — 마지막 한 마디 3종");

  const closes = [
    {
      type: "확신 클로징",
      when: "고객이 거의 결정한 상태",
      line: "\"두 모델 다 좋으셨는데, 부부 두 분에게는 시그니처가 더 맞을 것 같아요. 이걸로 주문 도와드릴까요?\"",
      note: "선택지 좁히기 + 행동 유도"
    },
    {
      type: "옵션 클로징",
      when: "결제 직전 마지막 제안",
      line: "\"매트리스만 가져가실래요, 토퍼·베개까지 침실 풀세트로 같이 가져가실래요?\"",
      note: "Yes/No가 아닌 A/B 선택"
    },
    {
      type: "안전 클로징",
      when: "결정 망설일 때",
      line: "\"오늘 결정 안 하셔도 돼요. 다만 30일 트라이얼이 있으니, 일주일 자보시고 안 맞으면 교환·환불 가능합니다.\"",
      note: "리스크 제로 명시 → 결정 부담 ↓"
    },
  ];

  closes.forEach((c, i) => {
    const y = 1.5 + i * 1.15;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 1.05,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 0.08, h: 1.05,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(c.type, {
      x: 0.7, y: y + 0.1, w: 2.0, h: 0.3,
      fontSize: 13, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText("WHEN  " + c.when, {
      x: 2.7, y: y + 0.13, w: 5.5, h: 0.25,
      fontSize: 10, color: C.gold, fontFace: FONT_B, charSpacing: 1, margin: 0
    });
    s.addText(c.note, {
      x: 8.0, y: y + 0.13, w: 1.4, h: 0.25,
      fontSize: 9, color: C.muted, fontFace: FONT_B,
      align: "right", margin: 0
    });
    s.addText(c.line, {
      x: 0.7, y: y + 0.45, w: 8.6, h: 0.55,
      fontSize: 12, color: C.goldDeep, fontFace: FONT_H,
      italic: true, margin: 0
    });
  });

  s.addText("Tip — 클로징은 \"강요\"가 아니라 \"정리\"입니다. 고객 머릿속에 흩어진 선택지를 매니저가 한 번에 묶어주는 일.", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 40, "Part 2 · 클로징");
}

// ============================================================
// SLIDE 41 — Part 6 디바이더
// ============================================================
makeDivider("03", "TOP 매니저 케이스 & 모범 사례",
  "데이터가 아닌 사람에게서 배우는 영업 — 우리 매장의 영웅들.");

// ============================================================
// SLIDE 42 — TOP 매니저 7가지 행동 패턴
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "7 BEHAVIORS", "TOP 매니저 7가지 행동 패턴");

  const behaviors = [
    { num: "1", t: "먼저 듣는다", d: "팔기보다 묻는다. 첫 5분은 질문만." },
    { num: "2", t: "공간을 묻는다", d: "예산 아닌 \"방 크기\"가 첫 질문." },
    { num: "3", t: "체험을 만든다", d: "누워보지 않은 고객은 안 산다." },
    { num: "4", t: "부부를 분리한다", d: "두 사람 따로 의견을 듣는다." },
    { num: "5", t: "기록한다", d: "고객·관심 모델·연락처를 즉시 메모." },
    { num: "6", t: "기다린다", d: "압박 NO. 두 번째 방문을 만든다." },
    { num: "7", t: "사후를 약속한다", d: "배송·A/S·재방문 약속을 지킨다." },
  ];

  behaviors.forEach((b, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.5 + col * 2.3, y = 1.5 + row * 1.6;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.1, h: 1.45,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.7, y: y + 0.2, w: 0.55, h: 0.55,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(b.num, {
      x: x + 0.7, y: y + 0.2, w: 0.55, h: 0.55,
      fontSize: 18, color: C.white, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(b.t, {
      x: x + 0.1, y: y + 0.85, w: 1.9, h: 0.3,
      fontSize: 12, color: C.ink, fontFace: FONT_H,
      bold: true, align: "center", margin: 0
    });
    s.addText(b.d, {
      x: x + 0.1, y: y + 1.15, w: 1.9, h: 0.3,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B,
      align: "center", margin: 0
    });
  });

  s.addText("핵심 — 7가지 모두 한 줄로 묶으면 \"고객을 존중한다\"입니다. 기술이 아니라 태도가 본질입니다.", {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 42, "Part 3 · 7가지 행동");
}

// ============================================================
// ============================================================
// PART 6 — TOP 매니저 케이스 REVISED [v2.2]
// 데일리 SV 세일즈톡 기반 실명·실 매장 데이터 반영
// 기존 SLIDE 43,44,45 (가상 케이스 3개) → 박은희·한용진·시그니처 5단계로 교체
// 기존 SLIDE 47 (자기 진단) 직후 → 매장별 판매 현황 + 데일리 세일즈톡 운영 안내 추가
// ============================================================

// ============================================================
// NEW SLIDE — 박은희 매니저 (롯데광주) 케이스 ⭐ 전국 1위
// 기존 SLIDE 43 (케이스 1) 교체
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "TOP CASE #1  ·  전국 1위",
    "박은희 매니저 (롯데광주) — 시그니처 매출 비중 80%");

  // 좌측: 실적 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 3.0, h: 3.7,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("3월 실적 (3/1~3/10)", {
    x: 0.65, y: 1.5, w: 2.7, h: 0.3,
    fontSize: 10, color: C.gold, fontFace: FONT_B, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("80%", {
    x: 0.65, y: 1.85, w: 2.7, h: 0.9,
    fontSize: 60, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("시그니처 매출 비중", {
    x: 0.65, y: 2.75, w: 2.7, h: 0.3,
    fontSize: 11, color: C.white, fontFace: FONT_B, margin: 0
  });
  s.addText("총 매출  35,135,000원", {
    x: 0.65, y: 3.25, w: 2.7, h: 0.3,
    fontSize: 11, color: C.line, fontFace: FONT_B, margin: 0
  });
  s.addText("시그니처  28,080,000원", {
    x: 0.65, y: 3.55, w: 2.7, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_B, bold: true, margin: 0
  });
  s.addText("오더 9건 중 8건이 시그니처\n시그니처 10조 판매", {
    x: 0.65, y: 4.05, w: 2.7, h: 0.9,
    fontSize: 10.5, color: C.white, fontFace: FONT_B, italic: true, margin: 0
  });

  // 우측: 판매 전략
  s.addText("판매 전략", {
    x: 3.7, y: 1.4, w: 5.8, h: 0.35,
    fontSize: 14, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
  });

  const tactics = [
    "첫 설명은 무조건 시그니처부터",
    "시그니처 특장점 집중 상담 진행",
    "순차적으로 하위 제품 설명",
    "객단가는 곧 매장 매출과 연결",
    "특정 제품 보고 온 고객 → 그 제품 설명 후, 다른 제품 언급 없이 시그니처 안내",
    "체험형 방문 고객 → 시그니처부터 체험 유도 + 특장점 설명",
    "고객님과 체험 전 니즈 파악 및 공감대 형성",
  ];
  tactics.forEach((t, i) => {
    const y = 1.85 + i * 0.45;
    s.addShape(pres.shapes.OVAL, {
      x: 3.7, y: y + 0.12, w: 0.18, h: 0.18,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(t, {
      x: 4.0, y, w: 5.5, h: 0.4,
      fontSize: 10.5, color: C.ink, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 43, "Part 3 · 박은희 (롯데광주)");
}

// ============================================================
// NEW SLIDE — 한용진 매니저 (광주쇼룸) 케이스 🥈 전국 2위
// 기존 SLIDE 44 (케이스 2) 교체
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "TOP CASE #2  ·  전국 2위",
    "한용진 매니저 (광주쇼룸) — 시그니처 매출 비중 48%");

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 3.0, h: 3.7,
    fill: { color: C.goldDeep }, line: { color: C.goldDeep }
  });
  s.addText("3월 실적 (3/1~3/10)", {
    x: 0.65, y: 1.5, w: 2.7, h: 0.3,
    fontSize: 10, color: C.creamSoft, fontFace: FONT_B, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("48%", {
    x: 0.65, y: 1.85, w: 2.7, h: 0.9,
    fontSize: 60, color: C.white, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("시그니처 매출 비중", {
    x: 0.65, y: 2.75, w: 2.7, h: 0.3,
    fontSize: 11, color: C.creamSoft, fontFace: FONT_B, margin: 0
  });
  s.addText("총 매출  50,155,000원", {
    x: 0.65, y: 3.25, w: 2.7, h: 0.3,
    fontSize: 11, color: C.creamSoft, fontFace: FONT_B, margin: 0
  });
  s.addText("시그니처  24,240,000원", {
    x: 0.65, y: 3.55, w: 2.7, h: 0.3,
    fontSize: 11, color: C.white, fontFace: FONT_B, bold: true, margin: 0
  });
  s.addText("오더 16건 중 9건이 시그니처\n시그니처 10조 판매", {
    x: 0.65, y: 4.05, w: 2.7, h: 0.9,
    fontSize: 10.5, color: C.creamSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  s.addText("판매 전략", {
    x: 3.7, y: 1.4, w: 5.8, h: 0.35,
    fontSize: 14, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
  });

  const tactics = [
    "입점 고객 질문 중심 상담 START",
    "현재 사용 매트리스의 불편 포인트 파악 필수",
    "원하는 착와감 · 수면 스타일 확인 필수",
    "시그니처 스프링 구조 차이 설명 강력 필수",
    "고객님 1순위·2순위 제품 재체험을 통한 확신감 부여",
    "시그니처 모델부터 체험 → 상위 모델 장점 안내",
    "지지력과 안락함은 별개라는 점 강조",
  ];
  tactics.forEach((t, i) => {
    const y = 1.85 + i * 0.45;
    s.addShape(pres.shapes.OVAL, {
      x: 3.7, y: y + 0.12, w: 0.18, h: 0.18,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(t, {
      x: 4.0, y, w: 5.5, h: 0.4,
      fontSize: 10.5, color: C.ink, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 44, "Part 3 · 한용진 (광주쇼룸)");
}

// ============================================================
// NEW SLIDE — 시그니처 매트리스 업셀링 5단계 프로세스
// 기존 SLIDE 45 (케이스 3) 교체
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "UP-SELLING PROCESS",
    "시그니처 매트리스 업셀링 — 본사 표준 5단계");

  s.addText("일방적 권유 NO. 고객이 \"체험을 통해 스스로 선택하게\" 하는 5단계 표준 프로세스.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const steps = [
    {
      n: "1", t: "고객 니즈 파악 & 라포 형성",
      d: "체험 전, 수면 습관 · 매트리스 교체 이유 · 수면 시 불편 사항 파악 → 공감대 형성 → 자연스러운 수면 컨설팅 환경 조성"
    },
    {
      n: "2", t: "상위 모델 중심 체험 진행",
      d: "상위 매트리스 모델부터 안내 → 고객이 편안함과 지지력을 먼저 경험 → 이후 비교 체험으로 매트리스별 차이 자연 인지"
    },
    {
      n: "3", t: "체험 중심 비교 설명",
      d: "직접 누워보고 체감하도록 유도 → 매트리스별 지지력·쿠션감·체형 분산 차이를 경험 → 설명이 아닌 체험 기반 선택"
    },
    {
      n: "4", t: "수면 체감 포인트 설명",
      d: "허리 지지력 · 어깨 체형 분산 · 뒤척임 감소 등 수면 시 체감 포인트 중심 설명 → 매일 · 장기 사용 강조하여 가치 이해 높임"
    },
    {
      n: "5", t: "베스트 모델 안내",
      d: "매장에서 가장 많이 판매되는 베스트 모델 안내 → 신뢰도 향상 → 자연스러운 상위 모델 선택 유도"
    },
  ];

  steps.forEach((st, i) => {
    const y = 1.9 + i * 0.65;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 0.55,
      fill: { color: i % 2 === 0 ? C.creamSoft : C.white },
      line: { color: C.line, width: 0.4 }
    });
    s.addShape(pres.shapes.OVAL, {
      x: 0.65, y: y + 0.12, w: 0.32, h: 0.32,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(st.n, {
      x: 0.65, y: y + 0.12, w: 0.32, h: 0.32,
      fontSize: 14, color: C.white, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(st.t, {
      x: 1.1, y: y + 0.05, w: 3.3, h: 0.25,
      fontSize: 11.5, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(st.d, {
      x: 4.45, y: y + 0.05, w: 5.0, h: 0.5,
      fontSize: 9.5, color: C.ink, fontFace: FONT_B, valign: "top", margin: 0
    });
  });

  s.addText("출처 — 본사 데일리 SV 세일즈톡 표준 프로세스 (3월 업셀링편)", {
    x: 0.5, y: 5.25, w: 9, h: 0.25,
    fontSize: 9, color: C.muted, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 45, "Part 3 · 시그니처 5단계");
}

// ============================================================
// NEW SLIDE — 3월 매장별 시그니처 판매 현황 (TOP 매장 자극용)
// 기존 SLIDE 47 (체크리스트) 직후 삽입
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "SALES BENCHMARK",
    "3월 매장별 시그니처 판매 — 너의 위치는?");

  // 좌측: 종합 수치
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 3.0, h: 3.7,
    fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
  });
  s.addText("3월 전체 (3/1~3/10)", {
    x: 0.65, y: 1.5, w: 2.7, h: 0.3,
    fontSize: 10, color: C.goldDeep, fontFace: FONT_B, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("총 매출", {
    x: 0.65, y: 1.85, w: 2.7, h: 0.25,
    fontSize: 10, color: C.muted, fontFace: FONT_B, margin: 0
  });
  s.addText("648,438,608원", {
    x: 0.65, y: 2.1, w: 2.7, h: 0.4,
    fontSize: 15, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("시그니처 매출", {
    x: 0.65, y: 2.55, w: 2.7, h: 0.25,
    fontSize: 10, color: C.muted, fontFace: FONT_B, margin: 0
  });
  s.addText("190,598,500원", {
    x: 0.65, y: 2.8, w: 2.7, h: 0.4,
    fontSize: 15, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("시그니처 비중", {
    x: 0.65, y: 3.3, w: 2.7, h: 0.25,
    fontSize: 10, color: C.muted, fontFace: FONT_B, margin: 0
  });
  s.addText("29%", {
    x: 0.65, y: 3.5, w: 2.7, h: 0.7,
    fontSize: 38, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("INSIGHT — 시그니처 체험 중심 매장은\n비중 평균 70%+. 프리미엄 기준 체험으로\n구매 전환율 상승.", {
    x: 0.65, y: 4.3, w: 2.7, h: 0.75,
    fontSize: 9, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  // 우측: TOP 매장 순위 표
  s.addText("매장별 시그니처 비중 TOP 10", {
    x: 3.7, y: 1.4, w: 5.8, h: 0.3,
    fontSize: 12, color: C.goldDeep, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });

  const top = [
    { rank: "1", name: "롯데광주", pct: "80%", color: C.gold },
    { rank: "2", name: "AK수원", pct: "79%", color: C.gold },
    { rank: "3", name: "AK분당", pct: "57%", color: C.gold },
    { rank: "4", name: "광주쇼룸", pct: "48%", color: C.goldDeep },
    { rank: "5", name: "울산", pct: "43%", color: C.goldDeep },
    { rank: "6", name: "롯데광복", pct: "43%", color: C.goldDeep },
    { rank: "7", name: "부산", pct: "43%", color: C.goldDeep },
    { rank: "8", name: "롯데동탄", pct: "41%", color: C.inkSoft },
    { rank: "9", name: "신세계타임", pct: "39%", color: C.inkSoft },
    { rank: "10", name: "롯데대구", pct: "36%", color: C.inkSoft },
  ];

  top.forEach((t, i) => {
    const y = 1.8 + i * 0.33;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 3.7, y, w: 5.8, h: 0.3,
      fill: { color: i % 2 === 0 ? C.creamSoft : C.white },
      line: { color: C.line, width: 0.3 }
    });
    s.addText(t.rank, {
      x: 3.8, y, w: 0.4, h: 0.3,
      fontSize: 10, color: C.muted, fontFace: FONT_H, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(t.name, {
      x: 4.3, y, w: 3.0, h: 0.3,
      fontSize: 10.5, color: C.ink, fontFace: FONT_B,
      valign: "middle", margin: 0
    });
    s.addText(t.pct, {
      x: 8.5, y, w: 0.9, h: 0.3,
      fontSize: 12, color: t.color, fontFace: FONT_H, bold: true,
      align: "right", valign: "middle", margin: 0
    });
  });

  s.addText("출처 — 본사 데일리 SV 세일즈톡 (3월호) · 매장별 데이터 매월 공유", {
    x: 0.5, y: 5.25, w: 9, h: 0.25,
    fontSize: 9, color: C.muted, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 46, "Part 3 · 매장별 벤치마크");
}

// ============================================================
// NEW SLIDE — 데일리 SV 세일즈톡 운영 안내
// 마지막에 추가 (가이드와 본사 월간 콘텐츠의 관계 정립)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "MONTHLY UPDATE",
    "데일리 SV 세일즈톡 — 본사 월간 업데이트 활용법");

  s.addText("이 가이드는 \"상시 매뉴얼\". 본사가 매월 발행하는 \"데일리 SV 세일즈톡\"은 \"이번 달 캠페인\". 둘이 한 세트입니다.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 11.5, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  // 비교 표
  const headers = ["", "이 세일즈 가이드 (상시)", "데일리 SV 세일즈톡 (월간)"];
  const rows = [
    ["주기", "분기 1회 개정", "매월 발행"],
    ["내용", "기본기 · 응대 표준 · 거절 처리\n인증 · 배송 · TOP 케이스 표준화", "그달의 매장별 실 데이터\n실명 TOP 사례 · 월간 캠페인"],
    ["분량", "약 35장", "3~5장"],
    ["용도", "신입 온보딩 + 매일 참조용", "월 1회 회독 + 캠페인 드라이브"],
    ["보관", "라미네이트 · 매장 비치", "사내 메신저 · 캘린더 알림"],
  ];

  const colW = [1.5, 3.8, 3.7];
  const colX = [0.5, 2.05, 5.9];
  let y = 1.95;

  headers.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y, w: colW[ci], h: 0.4,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(h, {
      x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.4,
      fontSize: 11, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", charSpacing: 1, margin: 0
    });
  });
  y += 0.4;
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? C.creamSoft : C.white;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: 0.5,
        fill: { color: fill }, line: { color: C.line, width: 0.4 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.5,
        fontSize: ci === 0 ? 11 : 10,
        color: ci === 0 ? C.goldDeep : C.ink,
        fontFace: ci === 0 ? FONT_H : FONT_B,
        bold: ci === 0,
        valign: "middle", margin: 0
      });
    });
    y += 0.5;
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.85, w: 9, h: 0.45,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("매니저 행동 — 매월 첫 영업일 데일리 SV 세일즈톡 받으면 5분 안에 읽고, TOP 매장 사례 1개 자기 매장에 적용해보세요.", {
    x: 0.65, y: 4.85, w: 8.7, h: 0.45,
    fontSize: 10.5, color: C.gold, fontFace: FONT_B,
    italic: true, valign: "middle", margin: 0
  });

  addFooter(s, 49, "Part 3 · 월간 업데이트");
}

// SLIDE 46 — 베스트 멘트 모음 (현장 즉시 사용)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "BEST LINES", "TOP 매니저들이 실제로 쓰는 멘트");

  const lines = [
    "\"먼저 한 번 누워보세요. 누워봐야 진짜 차이가 느껴지거든요.\"",
    "\"하루로 나누면 약 X백 원이세요. 8년 동안 매일 좋은 잠 사시는 거예요.\"",
    "\"두 분 같이 누워보세요. 모션 분리도 느끼실 수 있어요.\"",
    "\"이사 일정만 알려주시면 제가 챙겨서 연락드릴게요.\"",
    "\"매트리스만 가져가실래요, 침실 풀세트로 같이 가져가실래요?\"",
    "\"오늘 결정 안 하셔도 돼요. 30일 트라이얼 있어서 한 번 자보시고 결정하세요.\"",
    "\"제가 케어 가이드 한 장 정리해서 같이 챙겨드릴게요.\"",
    "\"두 분이 같이 다시 한 번 누워보고 가세요. 같은 모델 빼두지 않을게요.\"",
  ];

  lines.forEach((l, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 1.5 + row * 0.85;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 0.75,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: 0.75,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(l, {
      x: x + 0.2, y, w: 4.05, h: 0.75,
      fontSize: 11, color: C.ink, fontFace: FONT_B,
      italic: true, valign: "middle", margin: 0
    });
  });

  s.addText("→ 이 8개 멘트를 자기 말로 다시 적어 외워두세요. 매장 즉시 사용 가능한 \"현장 무기\"입니다.", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, bold: true, margin: 0
  });

  addFooter(s, 46, "Part 3 · 베스트 멘트");
}

// ============================================================
// SLIDE 47 — 매니저 자기 진단 체크리스트
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "SELF-CHECK", "매니저 자기 진단 체크리스트 — 매일 30초");

  s.addText("매일 마감 후, 이 10가지를 자신에게 물어보세요. 7개 이상 YES면 좋은 하루입니다.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const checks = [
    "오늘 고객의 \"공간\"을 물었는가?",
    "오늘 부부 두 사람 모두에게 눈을 맞췄는가?",
    "오늘 누워보게 한 고객이 절반 이상인가?",
    "오늘 명함·카탈로그를 챙겨준 고객이 있는가?",
    "오늘 가격 깎으려 하지 않았는가?",
    "오늘 거절한 고객에게 부드럽게 답했는가?",
    "오늘 토퍼·베개·침구 업셀을 시도했는가?",
    "오늘 고객 이름·연락처를 CRM에 입력했는가?",
    "오늘 케어·A/S 안내를 빠뜨리지 않았는가?",
    "오늘 \"좋은 잠 보내세요\" 인사로 마무리했는가?",
  ];

  checks.forEach((c, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 2.0 + row * 0.55;
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.05, y: y + 0.1, w: 0.25, h: 0.25,
      fill: { color: C.white }, line: { color: C.gold, width: 1 }
    });
    s.addText(`${String(i + 1).padStart(2, "0")}.  ${c}`, {
      x: x + 0.45, y, w: 3.85, h: 0.45,
      fontSize: 11, color: C.ink, fontFace: FONT_B,
      valign: "middle", margin: 0
    });
  });

  addFooter(s, 47, "Part 3 · 자기 진단");
}

// ============================================================
// SLIDE 48 — Part 7 디바이더
// ============================================================
makeDivider("04", "운영 · KPI · 부록",
  "팔고 끝이 아닙니다 — 사후가 진짜 시작입니다.");

// ============================================================
// SLIDE 49 — 사후 관리 & CRM
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "POST-SALE", "구매 이후 — 진짜 관계의 시작");

  const steps = [
    { d: "D+1", t: "감사 인사 카톡", desc: "\"구매 감사드려요. 배송 일정 다시 안내드릴게요\"" },
    { d: "D+7", t: "배송 전 안내", desc: "배송 일자 확인 + 사전 준비사항 안내 (공간 정리, 기존 침대 처리)" },
    { d: "D+14", t: "사용 점검", desc: "\"매트리스 잘 맞으세요? 처음 일주일이 가장 적응 중요해요\"" },
    { d: "D+30", t: "30일 트라이얼 체크", desc: "누어 매트리스의 경우 만족도 확인, 필요시 교환 안내" },
    { d: "D+180", t: "리뷰 요청", desc: "오늘의집·네이버 리뷰 부드럽게 요청 + 케어 팁 재안내" },
    { d: "D+365", t: "1주년 안부", desc: "\"1년 잘 쓰셨어요? 토퍼 교체 시점도 챙겨드릴게요\"" },
  ];

  steps.forEach((st, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.45 + row * 1.65;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 1.5,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 0.45,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(st.d, {
      x: x + 0.15, y, w: 1, h: 0.45,
      fontSize: 13, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", charSpacing: 1, margin: 0
    });
    s.addText(st.t, {
      x: x + 1.05, y, w: 1.75, h: 0.45,
      fontSize: 12, color: C.white, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
    s.addText(st.desc, {
      x: x + 0.2, y: y + 0.55, w: 2.55, h: 0.85,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("재구매·추천 고객은 신규 고객 획득 비용의 1/5. CRM은 매장 매니저의 \"개인 자산\"입니다.", {
    x: 0.5, y: 4.95, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, bold: true, margin: 0
  });

  addFooter(s, 49, "Part 4 · 사후 관리");
}

// ============================================================
// SLIDE 50 — KPI 자기 점검표
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "MANAGER KPI", "매니저 자기 KPI — 월간 자기 점검");

  const headers = ["KPI", "정의", "목표 (예시)", "왜 중요한가"];
  const rows = [
    ["방문 → 체험률", "매장 방문자 중 누워본 비율", "70%+", "체험 = 구매 결정의 핵심 트리거"],
    ["체험 → 견적률", "체험 고객 중 견적 받은 비율", "50%+", "심도 있는 컨설팅 지표"],
    ["견적 → 구매율", "견적 고객 중 실제 구매 비율", "30%+", "클로징·재방문 관리 지표"],
    ["객단가", "1건당 평균 매출", "본사 목표 +α", "업셀·풀세트 능력"],
    ["옵션 부착률", "토퍼·베개 등 동시 구매율", "40%+", "단순 매트리스 아닌 \"침실\" 제안"],
    ["재방문률", "1개월 내 재방문 비율", "20%+", "압박 NO, 신뢰 응대"],
    ["고객 리뷰 점수", "구매 후 리뷰 평점/건수", "4.7/5+", "장기 매장 평판"],
  ];

  const colX = [0.5, 2.4, 4.7, 6.4];
  const colW = [1.9, 2.3, 1.7, 3.1];
  let y = 1.45;
  headers.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y, w: colW[ci], h: 0.45,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(h, {
      x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.45,
      fontSize: 11, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", charSpacing: 1, margin: 0
    });
  });
  y += 0.45;
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? C.creamSoft : C.white;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: 0.42,
        fill: { color: fill }, line: { color: C.line, width: 0.5 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.42,
        fontSize: 10.5,
        color: ci === 0 ? C.goldDeep : C.ink,
        fontFace: ci === 0 ? FONT_H : FONT_B,
        bold: ci === 0,
        valign: "middle", margin: 0
      });
    });
    y += 0.42;
  });

  s.addText("운영 — 본사는 분기마다 매니저별 KPI 대시보드 제공. 1:1 코칭은 분기 1회 정례.", {
    x: 0.5, y: 4.95, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 50, "Part 4 · KPI");
}

// ============================================================
// SLIDE 51 — 고객 FAQ TOP 10
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "CUSTOMER FAQ", "고객이 가장 많이 묻는 질문 TOP 10");

  const faqs = [
    "Q. 매트리스 단단함은 어떻게 골라요?",
    "Q. 메모리폼이 여름에 덥지 않나요?",
    "Q. 사이즈가 방에 들어갈까요?",
    "Q. 배송·설치는 얼마나 걸려요?",
    "Q. 기존 매트리스 처분도 해주나요?",
    "Q. 보증 기간이 어떻게 되나요?",
    "Q. 토퍼는 꼭 필요한가요?",
    "Q. 1:1 맞춤은 추가 비용이 있나요?",
    "Q. 30일 트라이얼 환불 절차는?",
    "Q. 매트리스 회전·세탁은 어떻게?",
  ];

  faqs.forEach((q, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 1.5 + row * 0.7;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 0.55,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addText(`${String(i + 1).padStart(2, "0")}`, {
      x: x + 0.15, y, w: 0.4, h: 0.55,
      fontSize: 14, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
    s.addText(q, {
      x: x + 0.6, y, w: 3.65, h: 0.55,
      fontSize: 11.5, color: C.ink, fontFace: FONT_B,
      valign: "middle", margin: 0
    });
  });

  s.addText("→ 본사 매니저 채널에서 \"FAQ Top 30 상세 답변집\" 별도 배포. 응대 중 모르는 질문은 절대 답하지 말고 \"확인 후 안내\"하세요.", {
    x: 0.5, y: 5.05, w: 9, h: 0.3,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 51, "Part 4 · FAQ");
}

// ============================================================
// SLIDE 52 — 본사 지원 채널
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "HQ SUPPORT", "본사 지원 채널 — 막힐 때 바로 연락");

  const channels = [
    { t: "고객센터", c: "1644-4766", d: "평일 10:00~18:00\n(점심 12:30~14:00)" },
    { t: "A/S 신청", c: "구글폼", d: "보증·A/S 신청서\n링크 별도 제공" },
    { t: "CS 메일", c: "ogaren_cs@ogaren.com", d: "응대 관련 문의\n케이스 공유" },
    { t: "제휴·기업", c: "contact@ogaren.com", d: "B2B·기업 거래\n쇼룸 대량 문의" },
    { t: "마케팅 자료", c: "본사 매니저 카톡 채널", d: "신상품·프로모션\n월간 업데이트" },
    { t: "교육·코칭", c: "분기 정기 워크샵", d: "롤플레잉·신제품 교육\n1:1 코칭" },
  ];

  channels.forEach((ch, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.55 + row * 1.65;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 1.5,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addText(ch.t, {
      x: x + 0.2, y: y + 0.15, w: 2.55, h: 0.3,
      fontSize: 12, color: C.goldDeep, fontFace: FONT_H,
      bold: true, charSpacing: 1, margin: 0
    });
    s.addText(ch.c, {
      x: x + 0.2, y: y + 0.5, w: 2.55, h: 0.35,
      fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(ch.d, {
      x: x + 0.2, y: y + 0.9, w: 2.55, h: 0.6,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 52, "Part 4 · 본사 지원");
}

// ============================================================
// PART 8 — 인증 & 안전성 [NEW v2.0]
// ============================================================

// SLIDE 66 — Part 9 디바이더
makeDivider("05", "배송 실무 가이드",
  "\"이거 우리 집에 들어와요?\" — 매니저가 즉시 답하는 배송 체크리스트.");

// ============================================================
// SLIDE 67 — 엘리베이터 사이즈별 매트리스 반입 가이드
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "ELEVATOR GUIDE", "엘리베이터 사이즈 × 매트리스 반입 기준");

  s.addText("고객이 매장에서 가장 자주 묻는 질문 — \"우리 집에 들어와요?\" 엘리베이터 인승을 먼저 확인하세요.", {
    x: 0.5, y: 1.35, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  // 엘리베이터 표
  const headers = ["엘리베이터 인승", "최대 반입 사이즈", "비고"];
  const rows = [
    ["8인승", "SS (슈퍼싱글)", "1인용 표준"],
    ["11인승", "Q · K (퀸 · 킹)", "부부 표준 — 가장 보편적"],
    ["16인승", "LK (라지킹)", "대형 평수 · 프리미엄"],
    ["18인승", "EK (이스턴킹)", "특수 사이즈 · 사전 확인 필수"],
  ];
  const colW = [2.6, 3.4, 3.0];
  const colX = [0.5, 3.15, 6.6];
  let y = 1.9;
  headers.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y, w: colW[ci], h: 0.45,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(h, {
      x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.45,
      fontSize: 11, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", charSpacing: 1, margin: 0
    });
  });
  y += 0.45;
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? C.creamSoft : C.white;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: 0.5,
        fill: { color: fill }, line: { color: C.line, width: 0.4 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.5,
        fontSize: ci === 0 ? 13 : 11,
        color: ci === 0 ? C.goldDeep : C.ink,
        fontFace: ci === 0 ? FONT_H : FONT_B,
        bold: ci === 0,
        valign: "middle", margin: 0
      });
    });
    y += 0.5;
  });

  // 주의 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.4, w: 9, h: 0.9,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("⚠  주의 — 인승 같아도 실제 규격 상이", {
    x: 0.7, y: 4.45, w: 8.6, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("엘리베이터 인증이 동일하더라도 내부 크기·출입문 너비 등 실제 규격에 따라 반입 가능 여부가 다를 수 있습니다.\n사전 확인 안내: \"고객님, 엘리베이터 내부 가로·세로·출입문 너비를 알려주시면 더 정확히 안내드릴 수 있어요.\"", {
    x: 0.7, y: 4.78, w: 8.6, h: 0.55,
    fontSize: 10, color: C.white, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 67, "Part 5 · 엘리베이터");
}

// ============================================================
// SLIDE 68 — 계단·사다리차·창문탈거 배송 옵션
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "STAIRS · LIFT · WINDOW", "엘리베이터 안 되면? — 계단 · 사다리차 · 창문탈거");

  // 3 박스: 계단, 사다리차, 창문 탈거
  const opts = [
    {
      title: "계단 배송",
      icon: "STAIRS",
      lines: [
        "2층 2만원 · 추가 층당 1만원",
        "복층 시 추가 작업비 2만원",
        "제품 규격 이상의 가로폭 확보 필요",
        "협소 시 천장 높이도 확보 필요",
        "직선형만 가능 · ㄱ자/ㄴ자 제한 있음",
        "최대 5층까지 가능",
      ],
    },
    {
      title: "사다리차",
      icon: "LADDER",
      lines: [
        "고객 직접 호출 (사용료 고객 부담)",
        "자사 호출 시 콜 비용 추가 (서울 기준)",
        "엘리베이터·계단 모두 불가 시 필수",
        "복도 진입 불가 시에도 필요",
        "현장 거부로 회수 시 왕복비 고객 부담",
      ],
    },
    {
      title: "창문 탈거",
      icon: "WINDOW",
      lines: [
        "사다리차로 발코니·거실 창 반입 시",
        "창문 폭이 제품 규격보다 좁을 때",
        "이중창·안정창·통창은 탈거 없이 어려움",
        "사전 협의 필요 (별도 비용 발생 가능)",
      ],
    },
  ];

  opts.forEach((o, i) => {
    const x = 0.5 + i * 3.05, y = 1.4;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 3.5,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 0.55,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(o.title, {
      x: x + 0.2, y, w: 2.6, h: 0.55,
      fontSize: 14, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", charSpacing: 1, margin: 0
    });
    s.addText(o.icon, {
      x: x + 2.0, y, w: 0.8, h: 0.55,
      fontSize: 9, color: C.muted, fontFace: FONT_B,
      bold: true, charSpacing: 1, align: "right", valign: "middle", margin: 0
    });
    // 라인 리스트
    o.lines.forEach((line, j) => {
      const ly = y + 0.7 + j * 0.42;
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.2, y: ly + 0.12, w: 0.1, h: 0.1,
        fill: { color: C.gold }, line: { color: C.gold }
      });
      s.addText(line, {
        x: x + 0.4, y: ly, w: 2.4, h: 0.38,
        fontSize: 10, color: C.ink, fontFace: FONT_B,
        valign: "top", margin: 0
      });
    });
  });

  // 하단 유의사항
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.0, w: 9, h: 0.35,
    fill: { color: C.alert, transparency: 85 }, line: { color: C.alert, width: 0.5 }
  });
  s.addText("최종 판단 — 현장 구조·건물 규정·주차 환경에 따라 가능 범위 상이. 배송 기사 현장 판단이 최종입니다.", {
    x: 0.65, y: 5.0, w: 8.7, h: 0.35,
    fontSize: 10, color: C.alert, fontFace: FONT_B, bold: true, valign: "middle", margin: 0
  });

  addFooter(s, 68, "Part 5 · 계단·사다리차·탈거");
}

// ============================================================
// SLIDE 69 — 오늘의집 배송 조회 가이드
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "OHOUSE TRACKING", "오늘의집 배송 조회 — 송장상태 7가지 해독");

  s.addText("고객이 \"제 침대 어디까지 왔어요?\" 물으면, 매니저는 즉시 송장상태로 답해야 합니다.", {
    x: 0.5, y: 1.4, w: 9, h: 0.35,
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  // 좌측: 조회 경로
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.85, w: 2.6, h: 3.2,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("조회 방법", {
    x: 0.65, y: 1.95, w: 2.4, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });
  const steps = [
    "1. 주문 → 주문조회",
    "2. 필터값 설정",
    "3. 유사검색/일치검색 선택",
    "4. 송장상태 확인",
    "",
    "배송완료 사진/서명",
    "→ 주문상세 → 상세보기",
    "→ 상단 사진보기",
  ];
  steps.forEach((st, i) => {
    s.addText(st, {
      x: 0.65, y: 2.35 + i * 0.32, w: 2.4, h: 0.3,
      fontSize: 10, color: i >= 5 ? C.gold : C.white, fontFace: FONT_B,
      bold: i === 5, margin: 0
    });
  });

  // 우측: 송장상태 7가지
  s.addText("송장상태 — 의미 한눈에", {
    x: 3.3, y: 1.85, w: 6, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });

  const states = [
    { t: "주소오류", c: C.alert, d: "카카오맵 미조회 — 주소 수정 필요" },
    { t: "처리중", c: C.muted, d: "전산 처리 중 (5분 이상이면 문의)" },
    { t: "재고없음", c: C.alert, d: "재고 미할당 — 가용재고 확인" },
    { t: "접수대기", c: C.gold, d: "정상 등록 · 담당자 처리 예정" },
    { t: "접수완료/스케줄확정", c: C.gold, d: "재고 할당 완료 · 지정일 배송 예정" },
    { t: "배송중", c: C.goldDeep, d: "기사님 배송 중" },
    { t: "배송완료", c: C.goldDeep, d: "완료 · 사진/고객서명 확인 가능" },
  ];
  states.forEach((st, i) => {
    const y = 2.2 + i * 0.4;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 3.3, y, w: 6.2, h: 0.34,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.3 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 3.3, y, w: 0.08, h: 0.34,
      fill: { color: st.c }, line: { color: st.c }
    });
    s.addText(st.t, {
      x: 3.5, y, w: 2.3, h: 0.34,
      fontSize: 10.5, color: C.ink, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
    s.addText(st.d, {
      x: 5.85, y, w: 3.55, h: 0.34,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B,
      valign: "middle", margin: 0
    });
  });

  // 하단 Tip
  s.addText("Tip — 주문승인은 지정일 -1일 오전 11시 이전. 배송일정상담 의뢰 건은 임시로 지정일 \"2099년\" 표시, 해피콜 후 변경됩니다.", {
    x: 0.5, y: 5.1, w: 9, h: 0.3,
    fontSize: 10, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 69, "Part 5 · 오늘의집 조회");
}


// ============================================================
// SLIDE 70 — 마무리 (우리의 약속)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.ink };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.5, w: 0.5, h: 0.06,
    fill: { color: C.gold }, line: { color: C.gold }
  });

  s.addText("CLOSING  ·  우리의 약속", {
    x: 0.5, y: 0.7, w: 9, h: 0.3,
    fontSize: 12, color: C.gold, fontFace: FONT_B,
    bold: true, charSpacing: 4, margin: 0
  });

  s.addText("좋은 잠은\n좋은 매니저에서 시작됩니다.", {
    x: 0.5, y: 1.6, w: 9, h: 1.8,
    fontSize: 44, color: C.white, fontFace: FONT_H,
    bold: true, margin: 0
  });

  s.addText(
    "고객은 침대를 사러 오는 것이 아니라, 인생의 3분의 1을 결정하러 옵니다.\n" +
    "그 결정의 가장 가까이 서 있는 사람이 바로 우리입니다.",
    {
      x: 0.5, y: 3.6, w: 9, h: 0.9,
      fontSize: 15, color: C.line, fontFace: FONT_B,
      italic: true, margin: 0
    }
  );

  s.addText("SLEEPER  ·  nooer  ·  Toddles  ·  OGAREN", {
    x: 0.5, y: 4.8, w: 9, h: 0.3,
    fontSize: 12, color: C.gold, fontFace: FONT_B,
    bold: true, charSpacing: 5, margin: 0
  });

  s.addText("Edition 2.1  ·  인증·안전성·배송 실무 보강판  ·  오가렌 경영기획본부  ·  www.sleeper.co.kr", {
    x: 0.5, y: 5.2, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: FONT_B, margin: 0
  });
}


pres.writeFile({ fileName: "오가렌_세일즈가이드_세션2.pptx" })
  .then(fn => console.log("DONE:", fn))
  .catch(err => { console.error("ERROR:", err); process.exit(1); });
