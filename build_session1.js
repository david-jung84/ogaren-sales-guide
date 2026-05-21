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

  s.addText("세션 1  ·  우리는 그래서 잘 잔다", {
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
    ["Part 1", "오가렌과 3대 브랜드", "04"],
    ["Part 2", "매트리스 마스터하기", "12"],
    ["Part 3", "침대 프레임 마스터하기", "20"],
    ["Part 4", "인증 & 안전성", "25"],
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
// SLIDE 4 — Part 1 디바이더: 오가렌과 3대 브랜드
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
makeDivider("01", "오가렌과 3대 브랜드",
  "우리는 누구이며, 무엇을 파는가 — 매니저의 자신감은 여기서 시작됩니다.");

// ============================================================
// SLIDE 5 — 오가렌 정체성
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "WHO WE ARE", "우리는 침대 회사가 아닙니다");

  s.addText("우리는 고객의 \"하루의 1/3\"을 디자인합니다.", {
    x: 0.5, y: 1.45, w: 9, h: 0.5,
    fontSize: 22, color: C.goldDeep, fontFace: FONT_H, bold: true, italic: true, margin: 0
  });

  s.addText(
    "침대는 평균 8~10년을 사용하는, 일생에 손꼽히는 고관여 구매입니다.\n" +
    "고객은 \"가구\"를 사러 온 것이 아니라 \"앞으로 3,000일의 밤\"을 결정하러 옵니다.\n" +
    "그 결정의 가장 가까이 서 있는 사람이 매장 매니저, 바로 우리입니다.",
    {
      x: 0.5, y: 2.15, w: 9, h: 1.2,
      fontSize: 14, color: C.inkSoft, fontFace: FONT_B,
      margin: 0, paraSpaceAfter: 4
    }
  );

  // 우리의 약속 3가지
  const promises = [
    { num: "1", title: "1:1 맞춤", desc: "고객의 체형·취향·공간에\n맞춘 침대를 만듭니다" },
    { num: "2", title: "정직한 컨설팅", desc: "비싼 것이 아닌\n맞는 것을 권합니다" },
    { num: "3", title: "사후가 진짜 시작", desc: "구매 이후가\n관계의 시작입니다" },
  ];
  promises.forEach((p, i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 3.7, w: 2.9, h: 1.35,
      fill: { color: C.cream }, line: { color: C.line, width: 0.5 }
    });
    s.addText(p.num, {
      x: x + 0.2, y: 3.8, w: 0.4, h: 0.4,
      fontSize: 26, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(p.title, {
      x: x + 0.7, y: 3.85, w: 2.1, h: 0.35,
      fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(p.desc, {
      x: x + 0.2, y: 4.3, w: 2.6, h: 0.7,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 5, "Part 1 · 오가렌 정체성");
}

// ============================================================
// SLIDE 6 — 3대 브랜드 한 페이지 요약
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "THREE BRANDS", "한 회사, 세 가지 가치");

  const brands = [
    {
      name: "SLEEPER",
      kr: "슬립퍼",
      tagline: "더 깊이 있게",
      target: "프리미엄 / 1:1 맞춤",
      desc: "퍼스널 패브릭 · 호텔식 · 시그니처 매트리스. 자신만의 침실을 완성하고 싶은 고객을 위한 프리미엄 라인.",
      color: C.ink,
      accent: C.gold,
    },
    {
      name: "nooer",
      kr: "누어",
      tagline: "더 똑똑하게",
      target: "합리적 / 모던 / 트렌디",
      desc: "룬드 시리즈, 100% 국내생산 매트리스. 합리적 가격으로 모던한 라이프스타일을 완성하는 스마트 초이스.",
      color: C.goldDeep,
      accent: C.cream,
    },
    {
      name: "Toddles",
      kr: "토들즈",
      tagline: "더 사랑스럽게",
      target: "유아 · 아동 / 가족",
      desc: "아이의 첫 침대부터 청소년기 데이베드까지. 안전과 성장을 함께 디자인하는 어린이 라인.",
      color: C.blue,
      accent: C.creamSoft,
    },
  ];

  brands.forEach((b, i) => {
    const x = 0.5 + i * 3.1, y = 1.4;
    // 상단 컬러 블록
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.9, h: 1.0,
      fill: { color: b.color }, line: { color: b.color }
    });
    s.addText(b.name, {
      x: x + 0.2, y: y + 0.15, w: 2.5, h: 0.4,
      fontSize: 18, color: C.white, fontFace: FONT_H,
      bold: true, charSpacing: 3, margin: 0
    });
    s.addText(b.kr, {
      x: x + 0.2, y: y + 0.55, w: 2.5, h: 0.35,
      fontSize: 13, color: C.gold, fontFace: FONT_B, margin: 0
    });
    // 하단 흰 블록
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: y + 1.0, w: 2.9, h: 2.7,
      fill: { color: C.white }, line: { color: C.line, width: 0.5 }
    });
    s.addText(`"${b.tagline}"`, {
      x: x + 0.2, y: y + 1.1, w: 2.5, h: 0.4,
      fontSize: 16, color: b.color, fontFace: FONT_H,
      bold: true, italic: true, margin: 0
    });
    s.addText(b.target, {
      x: x + 0.2, y: y + 1.55, w: 2.5, h: 0.3,
      fontSize: 10, color: C.muted, fontFace: FONT_B,
      bold: true, charSpacing: 2, margin: 0
    });
    s.addText(b.desc, {
      x: x + 0.2, y: y + 2.0, w: 2.5, h: 1.6,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("우리 회사가 강한 이유 — 세 브랜드가 가격대 · 라이프스타일 · 가족 구성을 모두 커버합니다.", {
    x: 0.5, y: 4.7, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 6, "Part 1 · 3대 브랜드");
}

// ============================================================
// SLIDE 7 — 슬립퍼 라인업
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "BRAND 01  ·  SLEEPER", "슬립퍼 — 프리미엄 패브릭 침대 5선");

  const items = [
    { img: "images/bloom.jpg", name: "블룸", tag: "싱글헤드 패브릭", note: "여성 1인 가구 · 모던 침실" },
    { img: "images/hush.jpg", name: "허쉬", tag: "호텔식 침대", note: "호텔 라이크 · 클래식" },
    { img: "images/tulia.jpg", name: "튤리아", tag: "패브릭 디자인", note: "감각적 디자인 선호" },
    { img: "images/tira.jpg", name: "티라", tag: "스테디 베스트", note: "베이직 · 안정감" },
    { img: "images/chic.jpg", name: "시크", tag: "세련된 디자인", note: "모던 시크 · 도시형" },
  ];

  items.forEach((it, i) => {
    const x = 0.5 + i * 1.85, y = 1.4;
    s.addImage({ path: it.img, x, y, w: 1.75, h: 1.4, sizing: { type: "cover", w: 1.75, h: 1.4 } });
    s.addText(it.name, {
      x, y: y + 1.5, w: 1.75, h: 0.35,
      fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(it.tag, {
      x, y: y + 1.85, w: 1.75, h: 0.3,
      fontSize: 10, color: C.gold, fontFace: FONT_B, charSpacing: 1, margin: 0
    });
    s.addText(it.note, {
      x, y: y + 2.18, w: 1.75, h: 0.6,
      fontSize: 9, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  // 차별 포인트
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.35, w: 9, h: 0.85,
    fill: { color: C.cream }, line: { color: C.line, width: 0.5 }
  });
  s.addText("슬립퍼의 차별 포인트", {
    x: 0.7, y: 4.42, w: 3, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("· 1:1 맞춤 사이즈(S/SS/Q/K/LK)   · 패브릭 컬러·디자인 선택   · 헤드보드 디자인 다양   · 침실 인테리어 매칭 컨설팅",
    { x: 0.7, y: 4.72, w: 8.7, h: 0.45, fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0 }
  );

  addFooter(s, 7, "Part 1 · 슬립퍼 라인업");
}

// ============================================================
// SLIDE 8 — 누어 라인업 (룬드 시리즈)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "BRAND 02  ·  nooer", "누어 — 룬드 시리즈 (SNS 대란템)");

  const items = [
    { img: "images/lundvik_low.jpg", name: "룬드빅 로우", tag: "저상형 호텔형", note: "낮고 모던한 라인" },
    { img: "images/lundvik_storage.jpg", name: "룬드빅 수납", tag: "호텔형 + 수납", note: "좁은 침실 · 정리" },
    { img: "images/lund_arc.jpg", name: "룬드 아크", tag: "반원헤드", note: "부드러운 곡선 디자인" },
    { img: "images/lund_head.jpg", name: "룬드 헤드", tag: "기본 헤드", note: "베이직 · 안정감" },
    { img: "images/lund_flat.jpg", name: "룬드 무헤드", tag: "평상형 · 무헤드", note: "미니멀 · 자유 배치" },
  ];

  items.forEach((it, i) => {
    const x = 0.5 + i * 1.85, y = 1.4;
    s.addImage({ path: it.img, x, y, w: 1.75, h: 1.4, sizing: { type: "cover", w: 1.75, h: 1.4 } });
    s.addText(it.name, {
      x, y: y + 1.5, w: 1.75, h: 0.35,
      fontSize: 13, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(it.tag, {
      x, y: y + 1.85, w: 1.75, h: 0.3,
      fontSize: 10, color: C.gold, fontFace: FONT_B, charSpacing: 1, margin: 0
    });
    s.addText(it.note, {
      x, y: y + 2.18, w: 1.75, h: 0.6,
      fontSize: 9, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.35, w: 9, h: 0.85,
    fill: { color: C.cream }, line: { color: C.line, width: 0.5 }
  });
  s.addText("누어의 차별 포인트", {
    x: 0.7, y: 4.42, w: 3, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("· 100% 국내 생산   · 30일 트라이얼(매트리스)   · 합리적 가격 · 트렌디 디자인   · 2030 1인가구·신혼 부부 호응",
    { x: 0.7, y: 4.72, w: 8.7, h: 0.45, fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0 }
  );

  addFooter(s, 8, "Part 1 · 누어 라인업");
}

// ============================================================
// SLIDE 9 — 브랜드별 타겟 고객 페르소나
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "TARGET CUSTOMER", "브랜드별 타겟 고객 — 한눈에");

  const rows = [
    ["", "SLEEPER", "nooer", "Toddles"],
    ["주 타겟 연령", "30~50대 부부 · 신중년", "20대 후반~30대\n1인가구·신혼", "30~40대 가족\n(자녀 3~12세)"],
    ["라이프스타일", "프리미엄 · 인테리어 중시\n호텔 라이크 침실 추구", "트렌디 · 가성비\n오늘의집·SNS 영향", "안전 · 성장\n가족 시간 중시"],
    ["주요 의사결정 요인", "디자인 · 맞춤 · 품질", "디자인 · 가격 · 후기", "안전 · 내구성 · 디자인"],
    ["가격대 (Q 매트 기준)", "프리미엄대", "합리적 가격대", "어린이 라인 별도"],
  ];

  const colW = [1.7, 2.5, 2.5, 2.5];
  const colX = [0.5, 2.2, 4.7, 7.2];
  let y = 1.45;
  rows.forEach((row, ri) => {
    const h = ri === 0 ? 0.45 : 0.65;
    const fill = ri === 0 ? C.ink : (ri % 2 === 1 ? C.creamSoft : C.white);
    const textColor = ri === 0 ? C.white : C.ink;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h,
        fill: { color: fill }, line: { color: C.line, width: 0.5 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h,
        fontSize: ri === 0 ? 12 : 11,
        color: ri === 0 ? C.gold : textColor,
        fontFace: ri === 0 ? FONT_H : FONT_B,
        bold: ri === 0 || ci === 0, valign: "middle",
        margin: 0, charSpacing: ri === 0 ? 2 : 0
      });
    });
    y += h;
  });

  s.addText("Tip — 고객의 옷차림·동반자·첫 질문만으로 1차 페르소나는 보입니다. 그 다음은 \"가격대\"가 아니라 \"라이프스타일\"부터 물으세요.", {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 9, "Part 1 · 타겟 페르소나");
}

// ============================================================
// SLIDE 10 — 가격대 포지셔닝 맵
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "POSITIONING", "가격대 × 디자인 포지셔닝 맵");

  // 차트 영역 박스
  const cx = 1.0, cy = 1.4, cw = 8.0, ch = 3.3;
  s.addShape(pres.shapes.RECTANGLE, {
    x: cx, y: cy, w: cw, h: ch,
    fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
  });
  // X / Y 축 라벨
  s.addText("← 합리적 가격", {
    x: cx + 0.2, y: cy + ch - 0.3, w: 2, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: FONT_B, bold: true, margin: 0
  });
  s.addText("프리미엄 →", {
    x: cx + cw - 2.2, y: cy + ch - 0.3, w: 2, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: FONT_B,
    bold: true, align: "right", margin: 0
  });
  s.addText("↑ 디자인 · 맞춤", {
    x: cx + 0.15, y: cy + 0.1, w: 2.5, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: FONT_B, bold: true, margin: 0
  });
  s.addText("기본 · 표준 ↓", {
    x: cx + cw - 2.65, y: cy + 0.1, w: 2.5, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: FONT_B,
    bold: true, align: "right", margin: 0
  });

  // 브랜드 마커 (원)
  const markers = [
    { x: cx + 5.8, y: cy + 0.9, w: 1.6, label: "SLEEPER", sub: "프리미엄 · 맞춤", color: C.ink },
    { x: cx + 2.6, y: cy + 1.4, w: 1.4, label: "nooer", sub: "합리적 · 모던", color: C.goldDeep },
    { x: cx + 1.5, y: cy + 2.3, w: 1.1, label: "Toddles", sub: "어린이 · 안전", color: C.blue },
    { x: cx + 4.3, y: cy + 2.5, w: 1.0, label: "(경쟁사)", sub: "스프링 중심", color: C.muted },
  ];
  markers.forEach(m => {
    s.addShape(pres.shapes.OVAL, {
      x: m.x, y: m.y, w: m.w, h: m.w,
      fill: { color: m.color, transparency: 15 },
      line: { color: m.color, width: 1 }
    });
    s.addText(m.label, {
      x: m.x, y: m.y + m.w / 2 - 0.2, w: m.w, h: 0.3,
      fontSize: m.label === "(경쟁사)" ? 9 : 11,
      color: C.white, fontFace: FONT_H,
      bold: true, align: "center", charSpacing: 1.5, margin: 0
    });
    s.addText(m.sub, {
      x: m.x - 0.2, y: m.y + m.w + 0.02, w: m.w + 0.4, h: 0.3,
      fontSize: 9, color: C.inkSoft, fontFace: FONT_B,
      align: "center", margin: 0
    });
  });

  // 인사이트
  s.addText("매니저 핵심 메시지 — 우리 3개 브랜드가 합리에서 프리미엄까지, 1인가구부터 가족까지 모두 커버합니다.\n경쟁사 한 곳에서 모두 살 수 없는 이유, 바로 \"라인업의 깊이\"입니다.", {
    x: 0.5, y: 4.85, w: 9, h: 0.5,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 10, "Part 1 · 포지셔닝");
}

// ============================================================
// SLIDE 11 — 시장 인사이트 (침대 산업)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "MARKET INSIGHT", "침대를 사는 고객은 이렇게 생각합니다");

  const stats = [
    { num: "8~10년", label: "평균 침대 교체 주기", desc: "고관여 · 신중한 구매" },
    { num: "60%↑", label: "부부 함께 방문 비율", desc: "동시 설득이 핵심" },
    { num: "1/3", label: "인생에서 누워있는 시간", desc: "수면 = 건강 = 가치" },
    { num: "30분+", label: "평균 매장 체류 시간", desc: "체험과 컨설팅이 결정" },
  ];

  stats.forEach((st, i) => {
    const x = 0.5 + i * 2.3, y = 1.5;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.1, h: 2.3,
      fill: { color: C.cream }, line: { color: C.gold, width: 0.5 }
    });
    s.addText(st.num, {
      x, y: y + 0.25, w: 2.1, h: 0.85,
      fontSize: 32, color: C.goldDeep, fontFace: FONT_H,
      bold: true, align: "center", margin: 0
    });
    s.addText(st.label, {
      x: x + 0.1, y: y + 1.2, w: 1.9, h: 0.4,
      fontSize: 12, color: C.ink, fontFace: FONT_H,
      bold: true, align: "center", margin: 0
    });
    s.addText(st.desc, {
      x: x + 0.1, y: y + 1.65, w: 1.9, h: 0.55,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B,
      align: "center", margin: 0
    });
  });

  s.addText(
    "그래서 매니저는 \"파는 사람\"이 아니라 \"수면 컨설턴트\"여야 합니다.\n" +
    "고객이 30분을 투자한다면, 우리도 그 30분을 \"건강·관계·일상\"의 대화로 만들어야 합니다.",
    {
      x: 0.5, y: 4.2, w: 9, h: 0.8,
      fontSize: 12, color: C.inkSoft, fontFace: FONT_B,
      italic: true, margin: 0
    }
  );

  addFooter(s, 11, "Part 1 · 시장 인사이트");
}

// ============================================================
// SLIDE 12 — Part 2 디바이더
// ============================================================
makeDivider("02", "매트리스 마스터하기",
  "매니저의 진짜 무기는 \"매트리스를 설명할 줄 아는 능력\"입니다.");

// ============================================================
// SLIDE 13 — 매트리스 기본 구조 4종
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "MATTRESS 101", "매트리스 4대 기술 — 한 번에 정리");

  const techs = [
    {
      name: "포켓 스프링",
      pros: "독립 지지 · 모션 분리\n부부 동침에 최적",
      cons: "통기성 보통 · 무거움",
      who: "부부, 체중 차이 있는 분"
    },
    {
      name: "메모리폼",
      pros: "체압 분산 우수\n부드러운 감싸는 느낌",
      cons: "통기성 약함\n여름에 더울 수 있음",
      who: "옆으로 자는 분, 관절 부담"
    },
    {
      name: "라텍스",
      pros: "탄력 · 항균 · 통기성\n내구성 길다",
      cons: "무게 무거움\n가격대 있음",
      who: "알레르기·아토피, 단단함 선호"
    },
    {
      name: "하이브리드",
      pros: "스프링+폼 장점 결합\n범용성 가장 높음",
      cons: "제품별 편차 큼\n설명이 핵심",
      who: "취향 모르겠는 첫 구매자"
    },
  ];

  techs.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 1.45 + row * 1.85;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 1.7,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    // 좌측 컬러바
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: 1.7,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(t.name, {
      x: x + 0.25, y: y + 0.1, w: 4, h: 0.4,
      fontSize: 16, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText("장점  ", {
      x: x + 0.25, y: y + 0.5, w: 0.7, h: 0.3,
      fontSize: 10, color: C.gold, fontFace: FONT_B, bold: true, margin: 0
    });
    s.addText(t.pros, {
      x: x + 0.85, y: y + 0.5, w: 3.3, h: 0.5,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
    s.addText("주의  ", {
      x: x + 0.25, y: y + 1.0, w: 0.7, h: 0.3,
      fontSize: 10, color: C.alert, fontFace: FONT_B, bold: true, margin: 0
    });
    s.addText(t.cons, {
      x: x + 0.85, y: y + 1.0, w: 3.3, h: 0.4,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
    s.addText("추천 고객 — " + t.who, {
      x: x + 0.25, y: y + 1.4, w: 4, h: 0.3,
      fontSize: 10, color: C.goldDeep, fontFace: FONT_B, italic: true, bold: true, margin: 0
    });
  });

  addFooter(s, 13, "Part 2 · 매트리스 기술");
}

// ============================================================
// SLIDE 14 — 슬립퍼 매트리스 라인업
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "MATTRESS · SLEEPER", "슬립퍼 매트리스 4종");

  const mats = [
    {
      img: "images/signature.jpg",
      name: "시그니처 커스텀 매트리스 & 토퍼 세트",
      sub: "1:1 맞춤 · 슬립퍼의 시그니처",
      key: "체형·취향 맞춤 옵션 · 토퍼 세트",
      target: "프리미엄 · 부부 동침 · 까다로운 고객"
    },
    {
      img: "images/hybrid.jpg",
      name: "하이브리드 메모리폼 매트리스",
      sub: "스프링 + 메모리폼 결합",
      key: "체압 분산 + 모션 분리 · 범용성 우수",
      target: "첫 구매자 · 취향 모를 때 안전한 추천"
    },
    {
      img: "images/stable_latex.jpg",
      name: "스테이블 라텍스 매트리스",
      sub: "라텍스 단일 구조",
      key: "항균·통기성·내구성 · 단단한 지지",
      target: "알레르기·아토피, 단단함 선호"
    },
    {
      img: "images/kr_memory.jpg",
      name: "한국형 고밀도 메모리폼 매트리스",
      sub: "한국 체형 맞춤 설계",
      key: "지지력 강화 · 한국 체형에 맞춘 밀도",
      target: "허리 부담, 단단한 메모리폼 원하는 분"
    },
  ];

  mats.forEach((m, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 1.45 + row * 1.75;
    s.addImage({ path: m.img, x, y, w: 1.5, h: 1.6, sizing: { type: "cover", w: 1.5, h: 1.6 } });
    s.addText(m.name, {
      x: x + 1.65, y: y + 0.05, w: 2.65, h: 0.55,
      fontSize: 12, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(m.sub, {
      x: x + 1.65, y: y + 0.6, w: 2.65, h: 0.3,
      fontSize: 10, color: C.gold, fontFace: FONT_B, charSpacing: 1, margin: 0
    });
    s.addText("KEY  ", {
      x: x + 1.65, y: y + 0.95, w: 0.5, h: 0.25,
      fontSize: 9, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
    });
    s.addText(m.key, {
      x: x + 2.1, y: y + 0.95, w: 2.2, h: 0.45,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
    s.addText("WHO  ", {
      x: x + 1.65, y: y + 1.35, w: 0.5, h: 0.25,
      fontSize: 9, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
    });
    s.addText(m.target, {
      x: x + 2.1, y: y + 1.35, w: 2.2, h: 0.3,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 14, "Part 2 · 슬립퍼 매트리스");
}

// ============================================================
// SLIDE 15 — 누어 매트리스 3종 + 토퍼
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "MATTRESS · nooer", "누어 매트리스 — 편/포/탄 그리고 든든");

  const items = [
    { img: "images/pyeonan.jpg", name: "편안히 누어", sub: "오가닉 코튼 스프링", who: "친환경, 자연 소재\n중간 지지력" },
    { img: "images/pogeun.jpg", name: "포근히 누어", sub: "간단 설치 · 압축포장", who: "1인가구, 빠른 배송\n간편함 선호" },
    { img: "images/tantan.jpg", name: "탄탄히 누어", sub: "단단한 지지력", who: "허리 부담 있는 분\n단단함 선호" },
    { img: "images/topper.jpg", name: "든든 토퍼", sub: "프리미엄 린넨 토퍼", who: "기존 매트리스 업그레이드\n5cm 안락 추가" },
  ];

  items.forEach((it, i) => {
    const x = 0.5 + i * 2.3, y = 1.5;
    s.addImage({ path: it.img, x, y, w: 2.1, h: 1.5, sizing: { type: "cover", w: 2.1, h: 1.5 } });
    s.addText(it.name, {
      x, y: y + 1.6, w: 2.1, h: 0.35,
      fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(it.sub, {
      x, y: y + 1.95, w: 2.1, h: 0.3,
      fontSize: 10, color: C.gold, fontFace: FONT_B, charSpacing: 1, margin: 0
    });
    s.addText(it.who, {
      x, y: y + 2.25, w: 2.1, h: 0.7,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  // 30일 트라이얼 강조 박스
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.55, w: 9, h: 0.7,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("매니저 핵심 무기  ·  누어 매트리스 30일 트라이얼", {
    x: 0.7, y: 4.6, w: 6, h: 0.3,
    fontSize: 13, color: C.gold, fontFace: FONT_H,
    bold: true, charSpacing: 2, margin: 0
  });
  s.addText("\"매트리스는 일주일 자봐야 알아요. 30일 써보시고 안 맞으면 교환·환불됩니다\" — 망설이는 고객에 결정적 한 마디", {
    x: 0.7, y: 4.92, w: 8.6, h: 0.3,
    fontSize: 10, color: C.white, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 15, "Part 2 · 누어 매트리스");
}

// ============================================================
// SLIDE 16 — 매트리스 추천 매칭표 (체형×자세)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "RECOMMEND MATRIX", "체형 × 수면 자세별 추천 매트릭스");

  const headers = ["", "옆으로 잠 (Side)", "바로 누움 (Back)", "엎드림 (Stomach)"];
  const body = [
    ["가벼운 체형\n(~60kg)", "하이브리드 메모리폼\n포근히 누어", "편안히 누어\n시그니처 커스텀", "포근히 누어 (얕은 매트)"],
    ["보통 체형\n(60~80kg)", "하이브리드 메모리폼\n한국형 고밀도", "시그니처 커스텀\n편안히 누어", "탄탄히 누어\n스테이블 라텍스"],
    ["무거운 체형\n(80kg~)", "한국형 고밀도 메모리폼\n시그니처 커스텀", "탄탄히 누어\n스테이블 라텍스", "스테이블 라텍스\n탄탄히 누어"],
  ];

  const colX = [0.5, 2.5, 4.85, 7.2];
  const colW = [2.0, 2.35, 2.35, 2.3];
  let y = 1.45;
  // 헤더
  headers.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y, w: colW[ci], h: 0.5,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(h, {
      x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.5,
      fontSize: 11, color: C.gold, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
  });
  y += 0.5;
  // 본문
  body.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? C.creamSoft : C.white;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: 0.85,
        fill: { color: ci === 0 ? C.cream : fill }, line: { color: C.line, width: 0.5 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.85,
        fontSize: ci === 0 ? 11 : 10,
        color: ci === 0 ? C.goldDeep : C.ink,
        fontFace: ci === 0 ? FONT_H : FONT_B,
        bold: ci === 0,
        align: "center", valign: "middle", margin: 0
      });
    });
    y += 0.85;
  });

  s.addText("Tip — 이 표는 \"시작점\"입니다. 반드시 고객을 실제로 누워보게 해서 본인 감각으로 최종 결정하게 하세요.", {
    x: 0.5, y: 4.6, w: 9, h: 0.4,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 16, "Part 2 · 추천 매트릭스");
}

// ============================================================
// SLIDE 17 — 부부 동침 컨설팅
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "COUPLE CONSULTING", "부부 동침 — 가장 어려운, 가장 중요한 컨설팅");

  s.addText("부부 매장 방문 60% 이상. 두 사람의 \"체중·자세·취향\"이 다르다는 사실에서 출발하세요.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  // 3가지 시나리오
  const scenarios = [
    {
      title: "체중 차이 큼",
      problem: "한쪽이 푹 꺼지거나 다른 쪽이 너무 단단함",
      solution: "독립 스프링 강화 모델 추천\n시그니처 커스텀 — 양면 다른 강도 옵션\n또는 토퍼로 한쪽만 조정"
    },
    {
      title: "한 명이 뒤척임 심함",
      problem: "상대방의 움직임이 그대로 전달",
      solution: "모션 분리 우수 모델 — 하이브리드, 시그니처\n\"동전 떨어뜨리기\" 데모 활용\n매트리스 큰 사이즈(K, LK) 추가 제안"
    },
    {
      title: "자세가 다름 (옆/바로)",
      problem: "한 매트리스로 두 자세 만족 어려움",
      solution: "중간 단단함 + 토퍼로 보정\n또는 시그니처 양면 매트리스로 분리\n토퍼 업셀 자연스러운 기회"
    },
  ];

  scenarios.forEach((sc, i) => {
    const x = 0.5 + i * 3.1, y = 1.95;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.9, h: 2.9,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.9, h: 0.5,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(`상황 ${i + 1}  ·  ${sc.title}`, {
      x: x + 0.15, y, w: 2.7, h: 0.5,
      fontSize: 11, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", charSpacing: 1, margin: 0
    });
    s.addText("문제", {
      x: x + 0.2, y: y + 0.6, w: 2.6, h: 0.25,
      fontSize: 10, color: C.alert, fontFace: FONT_B, bold: true, margin: 0
    });
    s.addText(sc.problem, {
      x: x + 0.2, y: y + 0.85, w: 2.6, h: 0.6,
      fontSize: 10.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
    s.addText("해법", {
      x: x + 0.2, y: y + 1.55, w: 2.6, h: 0.25,
      fontSize: 10, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
    });
    s.addText(sc.solution, {
      x: x + 0.2, y: y + 1.8, w: 2.6, h: 1.05,
      fontSize: 10.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 17, "Part 2 · 부부 컨설팅");
}

// ============================================================
// SLIDE 18 — 부부 객단가 상승 비법: 토퍼·베개 업셀
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "UP-SELL", "객단가를 자연스럽게 올리는 법");

  // 좌측: 든든 토퍼 강조
  s.addImage({ path: "images/topper.jpg", x: 0.5, y: 1.45, w: 3.5, h: 2.8,
    sizing: { type: "cover", w: 3.5, h: 2.8 } });
  s.addText("누어 든든 토퍼", {
    x: 0.5, y: 4.3, w: 3.5, h: 0.35,
    fontSize: 14, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("프리미엄 린넨 · 5cm 안락 추가", {
    x: 0.5, y: 4.65, w: 3.5, h: 0.3,
    fontSize: 10, color: C.gold, fontFace: FONT_B, margin: 0
  });

  // 우측: 4가지 업셀 시나리오
  s.addText("자연스러운 업셀 4단계", {
    x: 4.3, y: 1.4, w: 5.2, h: 0.4,
    fontSize: 15, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
  });

  const ups = [
    { num: "1", title: "체험 중 한 마디", desc: "\"이 매트리스에 든든 토퍼 한 번 얹어드릴까요? 차이 바로 느끼실 거예요.\"" },
    { num: "2", title: "허리·관절 언급 시", desc: "\"허리 부담이 있으시면 토퍼가 체압 분산해서 훨씬 편하세요.\"" },
    { num: "3", title: "기존 침대 있는 분", desc: "\"매트리스 새로 사기 부담스러우면 토퍼만으로도 완전히 달라져요.\"" },
    { num: "4", title: "결제 직전 마지막", desc: "\"베개·토퍼·침구 같이 가져가시면 침실이 한 번에 완성됩니다.\"" },
  ];
  ups.forEach((u, i) => {
    const y = 1.95 + i * 0.7;
    s.addText(u.num, {
      x: 4.3, y, w: 0.4, h: 0.55,
      fontSize: 22, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(u.title, {
      x: 4.8, y, w: 4.6, h: 0.3,
      fontSize: 12, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(u.desc, {
      x: 4.8, y: y + 0.28, w: 4.6, h: 0.4,
      fontSize: 10.5, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
    });
  });

  addFooter(s, 18, "Part 2 · 업셀");
}

// ============================================================
// SLIDE 19 — 매트리스 케어 & A/S
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "CARE & WARRANTY", "고객에게 꼭 안내해야 할 케어 · A/S");

  const items = [
    { t: "회전·회전", d: "3개월마다 매트리스 180° 회전\n양면일 경우 위아래도 뒤집기 권장" },
    { t: "통기", d: "주 1회 침구 걷고 30분 통풍\n수분 누적 방지로 수명 연장" },
    { t: "직사광선 NO", d: "직사광선 노출 피하기\n변색·소재 변형 원인" },
    { t: "흘림 즉시 처리", d: "물·음료 흘렸을 때 마른 수건으로\n바로 흡수 후 통풍" },
    { t: "토퍼 교체 주기", d: "토퍼는 1.5~2년 주기로 교체 권장\n매트리스 수명 연장 효과" },
    { t: "보증 기간", d: "구조 결함 보증 — 제품별 상이\n구매 시 보증서 발급, 본사 1644-4766" },
  ];

  items.forEach((it, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.45 + row * 1.6;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 1.45,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: 1.45,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(it.t, {
      x: x + 0.2, y: y + 0.15, w: 2.6, h: 0.35,
      fontSize: 13, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(it.d, {
      x: x + 0.2, y: y + 0.55, w: 2.6, h: 0.85,
      fontSize: 10.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("매니저 멘트 — \"제가 케어 방법 한 장 정리해서 같이 챙겨드릴게요.\" — 케어 가이드 제공은 \"진짜 전문가\" 인상을 결정합니다.", {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 19, "Part 2 · 케어 & A/S");
}

// ============================================================
// SLIDE 20 — Part 3 디바이더
// ============================================================
makeDivider("03", "침대 프레임 마스터하기",
  "디자인은 첫인상, 사이즈는 진짜 결정 요인.");

// ============================================================
// SLIDE 21 — 프레임 카테고리 4종
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "FRAME CATEGORY", "프레임 4대 카테고리");

  const cats = [
    { name: "패브릭 침대", desc: "감각적·따뜻한 분위기\n컬러·소재 선택 폭 넓음", who: "디자인 우선, 인테리어 중시" },
    { name: "원목 침대", desc: "내추럴·고급스러움\n오래 쓸수록 멋이 깊어짐", who: "친환경, 클래식 선호" },
    { name: "수납 침대", desc: "침대 하부 수납 공간\n좁은 침실 솔루션", who: "1인가구·신혼, 수납 부족" },
    { name: "호텔 침대", desc: "낮은 베이스·격자 헤드\n호텔식 격조", who: "프리미엄 무드, 럭셔리" },
  ];
  cats.forEach((c, i) => {
    const x = 0.5 + i * 2.3, y = 1.5;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.1, h: 3.0,
      fill: { color: C.cream }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.1, h: 0.5,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(c.name, {
      x: x + 0.15, y, w: 1.9, h: 0.5,
      fontSize: 13, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
    s.addText(c.desc, {
      x: x + 0.15, y: y + 0.65, w: 1.85, h: 1.2,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
    s.addText("WHO", {
      x: x + 0.15, y: y + 2.0, w: 1.85, h: 0.25,
      fontSize: 9, color: C.gold, fontFace: FONT_B, bold: true, charSpacing: 1, margin: 0
    });
    s.addText(c.who, {
      x: x + 0.15, y: y + 2.25, w: 1.85, h: 0.7,
      fontSize: 10, color: C.ink, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 21, "Part 3 · 프레임 카테고리");
}

// ============================================================
// SLIDE 22 — 사이즈 가이드
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "SIZE GUIDE", "사이즈별 추천 공간 · 사용자");

  const sizes = [
    { s: "S", n: "싱글", w: "100×200cm", who: "1인 / 자녀 (초중)", room: "방 2.5평~" },
    { s: "SS", n: "슈퍼싱글", w: "110×200cm", who: "1인 여유 / 청소년", room: "방 3평~" },
    { s: "Q", n: "퀸", w: "150×200cm", who: "부부 표준 / 신혼", room: "방 3.5평~ (가장 보편)" },
    { s: "K", n: "킹", w: "160×200cm", who: "부부 + 자녀 1", room: "방 4평~" },
    { s: "LK", n: "라지킹", w: "180×200cm", who: "가족 / 프리미엄", room: "방 4.5평~" },
  ];

  // 헤더
  const colW = [0.6, 1.6, 1.6, 3.0, 2.7];
  const colX = [0.5, 1.1, 2.7, 4.3, 7.3];
  const headers = ["", "사이즈", "치수 (W×L)", "주 사용자", "최소 추천 공간"];
  let y = 1.45;
  headers.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y, w: colW[ci], h: 0.45,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(h, {
      x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.45,
      fontSize: 11, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
  });
  y += 0.45;
  sizes.forEach((sz, ri) => {
    const fill = ri % 2 === 0 ? C.creamSoft : C.white;
    const row = ["", sz.n, sz.w, sz.who, sz.room];
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: 0.55,
        fill: { color: fill }, line: { color: C.line, width: 0.5 }
      });
      if (ci === 0) {
        s.addText(sz.s, {
          x: colX[ci] + 0.05, y, w: colW[ci] - 0.1, h: 0.55,
          fontSize: 14, color: C.gold, fontFace: FONT_H,
          bold: true, align: "center", valign: "middle", margin: 0
        });
      } else {
        s.addText(cell, {
          x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.55,
          fontSize: 11, color: C.ink, fontFace: FONT_B,
          valign: "middle", margin: 0
        });
      }
    });
    y += 0.55;
  });

  s.addText("매니저 핵심 질문 — \"침실 방 크기 어느 정도세요?\" → 사이즈 컨설팅의 출발점.\n공간 작은데 큰 침대를 권하지 않는 정직함이 신뢰를 만듭니다.", {
    x: 0.5, y: 4.75, w: 9, h: 0.5,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 22, "Part 3 · 사이즈");
}

// ============================================================
// SLIDE 23 — 1:1 맞춤 옵션 (USP)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "OUR USP", "오가렌의 무기 — 1:1 맞춤 제작");

  s.addText("\"우리는 만들어둔 침대를 파는 게 아니라, 고객님 침실에 맞춘 침대를 만들어 드립니다.\"", {
    x: 0.5, y: 1.4, w: 9, h: 0.5,
    fontSize: 16, color: C.goldDeep, fontFace: FONT_H, bold: true, italic: true, margin: 0
  });

  const options = [
    { num: "01", t: "사이즈", d: "S · SS · Q · K · LK 전 사이즈\n특수 사이즈 별도 문의" },
    { num: "02", t: "패브릭 · 컬러", d: "다양한 패브릭 소재 선택\n10여 가지 컬러 옵션" },
    { num: "03", t: "헤드보드", d: "디자인 · 높이 · 형태\n공간과 인테리어에 맞춤" },
    { num: "04", t: "기능 옵션", d: "수납 추가 · 모션 베이스\n저상형 / 평상형 선택" },
    { num: "05", t: "매트리스 강도", d: "단단함 · 부드러움\n시그니처 양면 강도" },
    { num: "06", t: "토퍼 · 침구", d: "전용 토퍼 매칭\n패브릭과 어울리는 침구" },
  ];

  options.forEach((o, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 2.1 + row * 1.4;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 1.25,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addText(o.num, {
      x: x + 0.2, y: y + 0.15, w: 0.7, h: 0.4,
      fontSize: 18, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(o.t, {
      x: x + 0.85, y: y + 0.18, w: 1.9, h: 0.35,
      fontSize: 13, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(o.d, {
      x: x + 0.2, y: y + 0.6, w: 2.5, h: 0.65,
      fontSize: 10.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("매니저 무기 — 경쟁사가 \"기성 사이즈만 됩니다\" 할 때, 우리는 \"고객님 방에 맞춰 만들어 드립니다\" 라고 할 수 있습니다.", {
    x: 0.5, y: 4.95, w: 9, h: 0.3,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 23, "Part 3 · 1:1 맞춤");
}

// ============================================================
// SLIDE 24 — 인테리어 매칭 컨설팅
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "INTERIOR MATCH", "침실 인테리어 매칭 — 패브릭 컬러 가이드");

  const tones = [
    { name: "그레이 톤", desc: "모던 · 도시형\n어떤 인테리어와도 잘 매치", best: "20~30대 1인가구, 모던 침실" },
    { name: "베이지 · 아이보리", desc: "내추럴 · 따뜻함\n베스트셀러 컬러", best: "신혼·가족, 우드 가구와 매치" },
    { name: "딥 그린 · 네이비", desc: "고급스러움 · 깊이감\n포인트 컬러 침실", best: "프리미엄 무드, 4050대 부부" },
    { name: "라이트 핑크 · 텍스처", desc: "감각적 · 트렌디\n오늘의집 SNS 인기", best: "여성 1인, 감성 인테리어" },
  ];

  tones.forEach((t, i) => {
    const x = 0.5 + i * 2.3, y = 1.5;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.1, h: 2.95,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    // 컬러 샘플 영역
    const sample = [C.muted, "C9AC8E", "3D5240", "E8C3C3"][i];
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.15, y: y + 0.15, w: 1.8, h: 0.8,
      fill: { color: sample }, line: { color: sample }
    });
    s.addText(t.name, {
      x: x + 0.15, y: y + 1.05, w: 1.8, h: 0.4,
      fontSize: 13, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(t.desc, {
      x: x + 0.15, y: y + 1.5, w: 1.8, h: 0.7,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
    s.addText("적합 고객", {
      x: x + 0.15, y: y + 2.2, w: 1.8, h: 0.25,
      fontSize: 9, color: C.gold, fontFace: FONT_B, bold: true, charSpacing: 1, margin: 0
    });
    s.addText(t.best, {
      x: x + 0.15, y: y + 2.45, w: 1.8, h: 0.5,
      fontSize: 9.5, color: C.ink, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("Tip — 고객의 핸드폰 배경, 인스타그램, 자녀 사진 등 \"공간 단서\"를 자연스럽게 물어보세요. 컬러 추천이 정확해집니다.", {
    x: 0.5, y: 4.65, w: 9, h: 0.4,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 24, "Part 3 · 인테리어 매칭");
}

// ============================================================
// SLIDE 53 — Part 8 디바이더
makeDivider("04", "인증 & 안전성",
  "우리는 \"좋다\"고 말하지 않습니다. 인증서로 증명합니다.");

// ============================================================
// SLIDE 54 — 우리의 인증 종합표
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "CERTIFICATION MAP", "오가렌 인증 자산 — 한 페이지 종합");

  const headers = ["카테고리", "주요 인증", "대상 제품"];
  const rows = [
    ["방사선 안전", "라돈 시험성적서 (한국)", "전 매트리스·베개·토퍼"],
    ["유해물질", "OEKO-TEX Standard 100, 6대 중금속, PFOS·PFOA", "원단·솜·토들즈"],
    ["내구성", "내구성·수직하중 시험성적서", "시그니처·하이브리드 전 등급"],
    ["천연 라텍스", "ISPA · GOLS · LGA · OEKO", "슬립퍼 스테이블"],
    ["메모리폼", "밀도 시험 (100K/60K/50K) · CertiPUR", "전 메모리폼·베개"],
    ["접착제", "GREENGUARD Gold · ECOPASSPORT · LGA", "전 매트리스 접착"],
    ["원료 원산지", "호주 양모·캐시미어 원산지 증명", "시그니처 토퍼"],
    ["공장 운영", "ISO 9001 · ISO 14001 · KS", "이앤이 트레이딩 (제조 공장)"],
    ["어린이·유아", "KC · 6대중금속 · OEKO · fiti", "토들즈 + 방수 매트커버"],
    ["원단 (패브릭)", "SGS 세균억제 · PCFfree · 마모강도 30-50K rubs", "FABB·NEA·RAVEN·갤럭시 스웨이드 등"],
  ];

  const colW = [2.0, 4.0, 3.0];
  const colX = [0.5, 2.55, 6.6];
  let y = 1.35;

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
        x: colX[ci], y, w: colW[ci], h: 0.35,
        fill: { color: fill }, line: { color: C.line, width: 0.4 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.35,
        fontSize: 9.5, color: ci === 0 ? C.goldDeep : C.ink,
        fontFace: ci === 0 ? FONT_H : FONT_B,
        bold: ci === 0, valign: "middle", margin: 0
      });
    });
    y += 0.35;
  });

  s.addText("매니저 핵심 — 이 표만 외워두면, 고객이 어떤 안전 질문을 해도 답할 수 있습니다.", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 54, "Part 4 · 인증 종합");
}

// ============================================================
// SLIDE 55 — 라돈 시험: 전 라인업 통과
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "RADON SAFETY", "라돈 시험 — 전 라인업 통과");

  s.addText("최근 침대 시장에서 가장 민감한 이슈는 \"라돈\". 우리는 모든 매트리스에 대해 한국 공인기관 시험을 받았습니다.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const lines = [
    "슬립퍼 시그니처 (프레스티지 / 프라임 / 익스클루시브)",
    "슬립퍼 하이브리드 (프라임 / 익스클루시브)",
    "슬립퍼 스테이블 (프라임 / 익스클루시브)",
    "슬립퍼 밸런스드 (프레스티지 / 익스클루시브)",
    "슬립퍼 한국형 고밀도 메모리폼",
    "슬립퍼 원헌드레드 토퍼 (7cm / 11cm)",
    "슬립퍼 경추베개",
    "누어 매트리스 (편안히 / 포근히 / 탄탄히)",
    "누어 든든 토퍼",
    "토들즈 버터 매트리스 (마일드 / 멜티 / 퓨어)",
  ];

  lines.forEach((l, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 2.0 + row * 0.42;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 0.36,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.4 }
    });
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.15, y: y + 0.08, w: 0.2, h: 0.2,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText("OK", {
      x: x + 0.15, y: y + 0.05, w: 0.2, h: 0.26,
      fontSize: 7, color: C.white, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(l, {
      x: x + 0.5, y, w: 3.7, h: 0.36,
      fontSize: 10.5, color: C.ink, fontFace: FONT_B,
      valign: "middle", margin: 0
    });
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.5, w: 9, h: 0.7,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("매니저 응대 멘트", {
    x: 0.7, y: 4.55, w: 4, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("\"저희는 모든 매트리스에 대해 한국 공인기관에서 라돈 시험을 받았습니다. 시험성적서 보여드릴 수 있어요.\"", {
    x: 0.7, y: 4.85, w: 8.6, h: 0.3,
    fontSize: 11, color: C.white, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 55, "Part 4 · 라돈");
}

// ============================================================
// SLIDE 56 — 메모리폼 밀도 + 라텍스 인증
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "FOAM & LATEX", "메모리폼 밀도 · 라텍스 4중 인증");

  s.addText("MEMORY FOAM  ·  밀도가 곧 품질", {
    x: 0.5, y: 1.4, w: 4.5, h: 0.35,
    fontSize: 12, color: C.gold, fontFace: FONT_H, bold: true, charSpacing: 2, margin: 0
  });

  const foams = [
    { d: "100kg", n: "초고밀도", p: "원헌드레드 토퍼 (7cm/11cm)" },
    { d: "100K", n: "시그니처 표준", p: "슬립퍼 시그니처" },
    { d: "60K", n: "범용 고밀도", p: "슬립퍼 하이브리드" },
    { d: "50K", n: "한국형 표준", p: "슬립퍼 한고메매" },
  ];

  foams.forEach((f, i) => {
    const y = 1.85 + i * 0.65;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 4.3, h: 0.55,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.4 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 0.06, h: 0.55,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(f.d, {
      x: 0.65, y, w: 1.2, h: 0.55,
      fontSize: 16, color: C.goldDeep, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
    s.addText(f.n, {
      x: 1.85, y: y + 0.05, w: 2.8, h: 0.25,
      fontSize: 11, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(f.p, {
      x: 1.85, y: y + 0.27, w: 2.8, h: 0.25,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("LATEX  ·  슬립퍼 스테이블 4중 인증", {
    x: 5.1, y: 1.4, w: 4.5, h: 0.35,
    fontSize: 12, color: C.gold, fontFace: FONT_H, bold: true, charSpacing: 2, margin: 0
  });

  const latex = [
    { t: "ISPA", d: "국제 천연 라텍스\n협회 인증 (2023)" },
    { t: "GOLS", d: "유기농 라텍스\n글로벌 표준 (2022)" },
    { t: "LGA", d: "독일 LGA\n품질 인증 (2017)" },
    { t: "OEKO", d: "유해물질 미검출\nCert. 70158" },
  ];
  latex.forEach((l, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 5.1 + col * 2.25, y = 1.85 + row * 1.3;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.1, h: 1.2,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(l.t, {
      x: x + 0.1, y: y + 0.15, w: 1.9, h: 0.5,
      fontSize: 22, color: C.gold, fontFace: FONT_H,
      bold: true, align: "center", margin: 0, charSpacing: 2
    });
    s.addText(l.d, {
      x: x + 0.1, y: y + 0.65, w: 1.9, h: 0.5,
      fontSize: 9.5, color: C.white, fontFace: FONT_B,
      align: "center", margin: 0
    });
  });

  s.addText("Tip — 메모리폼은 \"밀도\"가 핵심 품질 지표. 우리 100K = 시중 50K의 2배 밀도, 그만큼 오래갑니다.", {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 56, "Part 4 · 폼 · 라텍스");
}

// ============================================================
// SLIDE 57 — 친환경 접착제 & 공장 인증
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "FACTORY & GLUE", "공장 · 접착제까지 — 보이지 않는 곳의 인증");

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 4.3, h: 3.6,
    fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
  });
  s.addText("매트리스 접착제  ·  SIMALFA", {
    x: 0.7, y: 1.55, w: 4, h: 0.35,
    fontSize: 13, color: C.goldDeep, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("\"매트리스는 접착제만 안전하면 됩니다.\"", {
    x: 0.7, y: 1.95, w: 4, h: 0.35,
    fontSize: 11, color: C.gold, fontFace: FONT_B, italic: true, margin: 0
  });

  const glues = [
    { t: "GREENGUARD Gold", d: "어린이·학교용 환경 인증 (2023)\n실내 공기질 기준 통과" },
    { t: "ECOPASSPORT", d: "OEKO-TEX 가공 화학물질 인증\n매트리스 산업 표준" },
    { t: "LGA 성적서", d: "독일 품질 시험 (2023)\n장기 내구성 검증" },
  ];
  glues.forEach((g, i) => {
    const y = 2.4 + i * 0.82;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y, w: 3.9, h: 0.7,
      fill: { color: C.white }, line: { color: C.gold, width: 0.5 }
    });
    s.addText(g.t, {
      x: 0.8, y: y + 0.05, w: 3.7, h: 0.3,
      fontSize: 11, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(g.d, {
      x: 0.8, y: y + 0.35, w: 3.7, h: 0.35,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.3, h: 3.6,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("제조 공장 인증  ·  이앤이 트레이딩", {
    x: 5.4, y: 1.55, w: 4, h: 0.35,
    fontSize: 13, color: C.gold, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("\"좋은 매트리스는 좋은 공장에서 나옵니다.\"", {
    x: 5.4, y: 1.95, w: 4, h: 0.35,
    fontSize: 11, color: C.line, fontFace: FONT_B, italic: true, margin: 0
  });

  const factory = [
    { t: "KS K ISO 9001", d: "품질 경영 시스템 (2021)" },
    { t: "KS K ISO 14001", d: "환경 경영 시스템 (2021)" },
    { t: "KS 인증 (율암리 공장)", d: "한국 산업 표준 인증" },
  ];
  factory.forEach((f, i) => {
    const y = 2.4 + i * 0.82;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.4, y, w: 3.9, h: 0.7,
      fill: { color: C.ink }, line: { color: C.gold, width: 0.5 }
    });
    s.addText(f.t, {
      x: 5.5, y: y + 0.05, w: 3.7, h: 0.3,
      fontSize: 11, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(f.d, {
      x: 5.5, y: y + 0.35, w: 3.7, h: 0.35,
      fontSize: 9.5, color: C.white, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 57, "Part 4 · 공장·접착제");
}

// ============================================================
// SLIDE 58 — 누어의 인증 자산
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "nooer CERTIFIED", "누어 — 100% 국내 + 인증으로 완성");

  s.addText("누어는 \"합리적\"이지만 \"덜 안전한\" 매트리스가 아닙니다. 본사가 자신 있게 30일 트라이얼을 제공하는 이유.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const certs = [
    { t: "100% 국내 생산", d: "전 라인업 국내 공장 생산\n품질 직접 컨트롤" },
    { t: "라돈 시험성적서", d: "한국 공인기관 통과\n전 라인업 (편/포/탄)" },
    { t: "원단 OEKO-TEX", d: "유해물질 미검출\n피부 직접 닿는 부위 안전" },
    { t: "솜 OEKO-TEX", d: "이탈리아 인증 (2021-22)\nCertificato OEKO-TEX" },
    { t: "에어프로파일폼 밀도", d: "통기성 + 지지력 시험성적서\n편안히·포근히 코어 소재" },
    { t: "패딩 형광 fiti", d: "형광물질 미검출 (2025)\n알레르기·아토피 안심" },
  ];

  certs.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 2.0 + row * 1.45;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 1.3,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.08, h: 1.3,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(c.t, {
      x: x + 0.2, y: y + 0.15, w: 2.55, h: 0.35,
      fontSize: 12, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(c.d, {
      x: x + 0.2, y: y + 0.55, w: 2.55, h: 0.75,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("매니저 무기 — \"누어가 왜 합리적인 가격인지 아세요? 국내 직접 생산 + 마케팅 비용 줄여서, 인증은 그대로입니다.\"", {
    x: 0.5, y: 4.95, w: 9, h: 0.3,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 58, "Part 4 · 누어 인증");
}

// ============================================================
// SLIDE 59 — 토들즈 안전성 (어린이 라인)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "Toddles SAFETY", "토들즈 — 어린이가 매일 자는 곳, 가장 까다롭게");

  s.addText("자녀를 위한 침대는 \"좋아 보이는 것\"이 아니라 \"증명된 것\"이어야 합니다. 토들즈 버터 매트리스 3종(마일드·멜티·퓨어)은 다음 인증을 모두 통과했습니다.", {
    x: 0.5, y: 1.35, w: 9, h: 0.6,
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const safety = [
    { t: "유아 KC 안전 인증", d: "방수 매트커버 별도 KC\n어린이 공급자적합성" },
    { t: "6대 중금속 시험", d: "납·카드뮴·수은·크롬·셀레늄·비소\n전 항목 미검출" },
    { t: "라돈 시험성적서", d: "마일드·멜티·퓨어\n3종 각각 별도 시험" },
    { t: "원단 OEKO-TEX", d: "유해물질 미검출\n피부 직접 닿는 부위 안전" },
    { t: "원단 fiti 인증", d: "형광물질 미검출\n알레르기 안심" },
    { t: "방수포 OEKO + 진드기 방수", d: "어린이 침대의 핵심 위생\n진드기 차단 + 친환경" },
    { t: "패딩솜 KOTITI", d: "솜 안전성 시험 (2021)\n국내 공인 검사" },
    { t: "냉감성 COOD 시험", d: "여름 통기·냉감성 검증\n어린이 땀 안전" },
    { t: "접착제 OEKO-TEX", d: "어린이 매트리스 내장재\n접착제까지 안전 검증" },
  ];

  safety.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 2.05 + row * 1.05;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 2.85, h: 0.92,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.4 }
    });
    s.addText(c.t, {
      x: x + 0.15, y: y + 0.1, w: 2.6, h: 0.3,
      fontSize: 11, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(c.d, {
      x: x + 0.15, y: y + 0.42, w: 2.6, h: 0.5,
      fontSize: 9, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 59, "Part 4 · 토들즈 안전");
}

// ============================================================
// SLIDE 60 — 토퍼·베개 디테일 안전성
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "TOPPER · PILLOW", "토퍼 · 베개 — 원료 원산지까지 추적");

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 4.3, h: 3.7,
    fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
  });
  s.addText("슬립퍼 토퍼 라인", {
    x: 0.7, y: 1.55, w: 4, h: 0.35,
    fontSize: 13, color: C.goldDeep, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("· 시그니처 토퍼 — 호주 양모 + 캐시미어 (원산지 증명)\n· 양모 혼용률 시험성적서\n· 원헌드레드 토퍼 (7cm/11cm)\n  - 100kg 초고밀도 메모리폼\n  - 라돈 시험 통과\n  - 스퀘어 텐셀 원단", {
    x: 0.7, y: 1.95, w: 4, h: 1.6,
    fontSize: 10.5, color: C.ink, fontFace: FONT_B, margin: 0
  });
  s.addText("누어 든든 토퍼", {
    x: 0.7, y: 3.7, w: 4, h: 0.3,
    fontSize: 12, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
  });
  s.addText("· 통기고탄성폼 30kg 시험성적서\n· HR폼 비건 인증\n· 린넨삼중직 시험 (유니밋)\n· 라돈 시험 통과 (2025)", {
    x: 0.7, y: 4.0, w: 4, h: 1.0,
    fontSize: 10.5, color: C.ink, fontFace: FONT_B, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.3, h: 3.7,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("슬립퍼 경추베개", {
    x: 5.4, y: 1.55, w: 4, h: 0.35,
    fontSize: 13, color: C.gold, fontFace: FONT_H, bold: true, charSpacing: 1, margin: 0
  });
  s.addText("\"베개도 인증이 있다는 사실, 매니저만 압니다.\"", {
    x: 5.4, y: 1.95, w: 4, h: 0.35,
    fontSize: 11, color: C.line, fontFace: FONT_B, italic: true, margin: 0
  });

  const pillow = [
    { t: "CertiPUR (진양폴리)", d: "폼 안전성 국제 표준\n유해 화학물질 미검출" },
    { t: "폼 시험성적서", d: "물성 시험 (M2702201927)\n경도·복원력 검증" },
    { t: "라돈 시험", d: "한국 공인 (2024)\n베개도 라돈 안심" },
    { t: "원단 시험 (그레이/아이보리)", d: "2025년 시험성적서\n피부 안전성" },
  ];
  pillow.forEach((p, i) => {
    const y = 2.45 + i * 0.65;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.4, y, w: 3.9, h: 0.55,
      fill: { color: C.ink }, line: { color: C.gold, width: 0.4 }
    });
    s.addText(p.t, {
      x: 5.5, y: y + 0.05, w: 3.7, h: 0.25,
      fontSize: 10.5, color: C.gold, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(p.d, {
      x: 5.5, y: y + 0.3, w: 3.7, h: 0.22,
      fontSize: 9, color: C.white, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 60, "Part 4 · 토퍼·베개");
}

// ============================================================
// SLIDE 61 — 슬립퍼 패브릭 원단 인증
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "FABRIC CERTIFIED", "슬립퍼 패브릭 — 만지는 모든 것의 인증");

  s.addText("슬립퍼 패브릭 침대는 \"디자인이 예쁘다\"로 끝나면 절반만 설명한 것. 원단까지 인증서로 채웠습니다.", {
    x: 0.5, y: 1.4, w: 9, h: 0.4,
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const fabric = [
    { t: "OEKO-TEX Standard 100", d: "FABB · NEA · RAVEN · FROCA · 갤럭시 스웨이드\n전 원단 유해물질 미검출 인증" },
    { t: "SGS 세균억제 처리", d: "FABB-RAVEN 원단 / 위생 안심" },
    { t: "PCFfree", d: "FABB-NEA · FABB-RAVEN\n과불소화합물(PFC) 무첨가" },
    { t: "마모강도 30,000-50,000 rubs", d: "ZAKUDO 린넨 (PROCA)\n수십 년 사용에도 닳지 않음" },
    { t: "H2OH! 발수 원단", d: "FABIO (FROCA) 라인\n물·음료 흘려도 안심" },
    { t: "PFOS · PFOA 시험", d: "갤럭시 스웨이드 · ZAKUDO 린넨\n과불소화합물 검출 안 됨" },
    { t: "알러지 시험", d: "갤럭시 스웨이드 (크림 컬러)\n피부 알러지 안전성" },
    { t: "발수방오 처리", d: "갤럭시 스웨이드 추가 옵션\n생활 오염 강력 차단" },
  ];

  fabric.forEach((c, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.55, y = 2.0 + row * 0.78;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.3, h: 0.68,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.4 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: 0.68,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(c.t, {
      x: x + 0.2, y: y + 0.05, w: 4.0, h: 0.25,
      fontSize: 11, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(c.d, {
      x: x + 0.2, y: y + 0.3, w: 4.0, h: 0.38,
      fontSize: 9.5, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 61, "Part 4 · 패브릭 인증");
}

// ============================================================
// SLIDE 62 — 거절 처리 #8: 안전성 우려
// ============================================================
makeObjectionSlide(
  8, 62, "라돈 · 메모리폼 화학물질 괜찮나요?",
  "이런 질문이 나오는 게 가장 좋은 신호. 자신감 있게 인증서로 답하세요.",
  "\"걱정 마세요\"\n\"저희는 안전해요\"\n→ 추상적 답변. 신뢰가 안 됩니다.",
  "\"좋은 질문이세요. 저희는 모든 매트리스에 대해 한국 공인기관 라돈 시험을 받았고, 시험성적서를 보유하고 있어요. 매트리스 접착제도 GREENGUARD Gold — 어린이 학교용 기준 통과한 것 쓰고, 메모리폼은 OEKO-TEX 인증입니다. 시험성적서 제가 한 장 보여드릴게요.\"",
  "라돈·화학물질 불안은 \"논리\"가 아니라 \"증명\"으로 답해야 합니다.\n인증 카드를 항상 매장에 비치하고, 즉시 보여드리는 것이 가장 강력합니다.\n이 질문 받았을 때 매니저가 자료를 꺼낼 줄 알면, 그 자체로 \"전문가\" 인상이 됩니다."
);

// ============================================================
// SLIDE 63 — 매트리스 등급 체계
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "GRADE SYSTEM", "슬립퍼 매트리스 등급 체계 — 1라인 안의 3선택지");

  s.addText("슬립퍼는 \"한 라인 안에 등급\"이 있어요. 같은 시그니처라도 예산·니즈에 따라 프레스티지·프라임·익스클루시브를 권할 수 있습니다.", {
    x: 0.5, y: 1.35, w: 9, h: 0.5,
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  const grades = [
    { line: "시그니처", grades: "프레스티지 · 프라임 · 익스클루시브", spec: "1:1 맞춤 · 100K 메모리폼 · 토퍼 세트 옵션" },
    { line: "하이브리드", grades: "프라임 · 익스클루시브", spec: "60K 메모리폼 + 포켓 스프링 · 면사틴 60수 원단" },
    { line: "스테이블 (라텍스)", grades: "M · MH · 프라임 · 익스클루시브", spec: "천연 라텍스 · ISPA·GOLS·LGA · KC 인증" },
    { line: "밸런스드  [NEW]", grades: "프레스티지 · 익스클루시브", spec: "균형감 있는 지지력 · KC · 라돈 통과" },
    { line: "한고메매", grades: "한국형 고밀도 메모리폼", spec: "50K 밀도 · 한국 체형 맞춤 · 라돈 통과" },
  ];

  const headers = ["라인", "등급 옵션", "핵심 사양"];
  const colW = [2.0, 3.0, 4.0];
  const colX = [0.5, 2.55, 5.6];
  let y = 2.0;
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
  grades.forEach((g, ri) => {
    const fill = ri % 2 === 0 ? C.creamSoft : C.white;
    const row = [g.line, g.grades, g.spec];
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: 0.45,
        fill: { color: fill }, line: { color: C.line, width: 0.4 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.45,
        fontSize: 10.5,
        color: ci === 0 ? C.goldDeep : C.ink,
        fontFace: ci === 0 ? FONT_H : FONT_B,
        bold: ci === 0,
        valign: "middle", margin: 0
      });
    });
    y += 0.45;
  });

  s.addText("매니저 클로징 무기 — 고객이 \"프레스티지가 너무 비싸요\" → \"같은 시그니처 라인의 프라임도 똑같은 인증·1:1 맞춤이세요\"", {
    x: 0.5, y: 4.95, w: 9, h: 0.3,
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, italic: true, margin: 0
  });

  addFooter(s, 63, "Part 4 · 등급 체계");
}

// ============================================================
// SLIDE 64 — 매니저 휴대용 인증 카드
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "POCKET CARD", "매니저 휴대용 인증 카드 — 매장 즉시 사용");

  s.addText("이 슬라이드를 인쇄·라미네이트해서 매장에 비치하세요. 고객 응대 중 즉시 꺼내 보여드릴 한 장.", {
    x: 0.5, y: 1.35, w: 9, h: 0.4,
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, italic: true, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 1.85, w: 8.6, h: 3.4,
    fill: { color: C.ink }, line: { color: C.gold, width: 2 }
  });

  s.addText("OGAREN  ·  CERTIFIED SAFE", {
    x: 0.9, y: 2.0, w: 8.2, h: 0.4,
    fontSize: 14, color: C.gold, fontFace: FONT_H,
    bold: true, charSpacing: 4, margin: 0
  });

  s.addText("우리 매트리스가 통과한 인증", {
    x: 0.9, y: 2.4, w: 8.2, h: 0.35,
    fontSize: 18, color: C.white, fontFace: FONT_H, bold: true, margin: 0
  });

  const card = [
    { t: "라돈", d: "전 매트리스 한국 공인기관 시험 통과" },
    { t: "메모리폼", d: "100K·60K·50K 밀도 시험성적서" },
    { t: "라텍스 (스테이블)", d: "ISPA · GOLS · LGA · OEKO 4중 인증" },
    { t: "접착제", d: "GREENGUARD Gold · ECOPASSPORT" },
    { t: "원단 (전 라인)", d: "OEKO-TEX Standard 100" },
    { t: "공장", d: "ISO 9001 · ISO 14001 · KS" },
    { t: "토들즈 (어린이)", d: "유아 KC · 6대중금속 · OEKO" },
    { t: "토퍼 원산지", d: "호주 양모 · 캐시미어 원산지 증명" },
  ];
  card.forEach((c, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.9 + col * 4.15, y = 2.95 + row * 0.53;
    s.addShape(pres.shapes.OVAL, {
      x, y: y + 0.1, w: 0.2, h: 0.2,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(c.t, {
      x: x + 0.3, y, w: 1.6, h: 0.45,
      fontSize: 12, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", margin: 0
    });
    s.addText(c.d, {
      x: x + 1.9, y, w: 2.2, h: 0.45,
      fontSize: 9.5, color: C.white, fontFace: FONT_B,
      valign: "middle", margin: 0
    });
  });

  s.addText("\"고객님 안전 인증서, 보여드릴게요.\"  ·  1644-4766", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B,
    italic: true, bold: true, align: "center", margin: 0
  });

  addFooter(s, 64, "Part 4 · 휴대 카드");
}

// ============================================================
// SLIDE 65 — 인증 활용 매니저 행동 가이드
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, "HOW TO USE", "인증, 이렇게 활용하세요 — 5가지 행동 원칙");

  const actions = [
    { n: "1", t: "묻지 않아도 한 번은 언급", d: "응대 중반쯤 \"저희는 라돈 시험 모든 매트리스에 받았어요\" 한 번 자연스럽게." },
    { n: "2", t: "질문 받으면 자료 꺼내기", d: "추상적 답변 NO. 휴대 카드·카탈로그·시험성적서 PDF를 즉시 보여드리기." },
    { n: "3", t: "어린이 고객은 토들즈 안전성 먼저", d: "자녀와 함께 온 고객에게는 6대중금속·KC·진드기 방수포부터 설명." },
    { n: "4", t: "라텍스는 ISPA·GOLS·LGA 4중 인증으로 차별화", d: "라텍스 비교 고객에게는 우리 4중 인증이 \"진짜 천연\"임을 강조." },
    { n: "5", t: "패브릭 침대는 원단 인증을 디자인과 함께", d: "\"디자인뿐 아니라 마모강도 30,000회 견디는 원단\"이 럭셔리 본질." },
  ];

  actions.forEach((a, i) => {
    const y = 1.5 + i * 0.7;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9, h: 0.6,
      fill: { color: C.creamSoft }, line: { color: C.line, width: 0.4 }
    });
    s.addShape(pres.shapes.OVAL, {
      x: 0.65, y: y + 0.13, w: 0.35, h: 0.35,
      fill: { color: C.gold }, line: { color: C.gold }
    });
    s.addText(a.n, {
      x: 0.65, y: y + 0.13, w: 0.35, h: 0.35,
      fontSize: 14, color: C.white, fontFace: FONT_H,
      bold: true, align: "center", valign: "middle", margin: 0
    });
    s.addText(a.t, {
      x: 1.15, y: y + 0.05, w: 8.0, h: 0.3,
      fontSize: 12, color: C.ink, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText(a.d, {
      x: 1.15, y: y + 0.32, w: 8.0, h: 0.25,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("핵심 — 인증은 \"방어\"가 아니라 \"공격\". 고객이 묻기 전에 우리가 먼저 보여주면, 프리미엄 가격이 정당화됩니다.", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, italic: true, bold: true, margin: 0
  });

  addFooter(s, 65, "Part 4 · 활용 원칙");
}


// ============================================================
// PART 9 — 배송 실무 가이드 [NEW v2.1]
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


pres.writeFile({ fileName: "오가렌_세일즈가이드_세션1.pptx" })
  .then(fn => console.log("DONE:", fn))
  .catch(err => { console.error("ERROR:", err); process.exit(1); });
