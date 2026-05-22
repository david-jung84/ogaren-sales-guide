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

  s.addText("Edition 2.1  ·  인증·안전성·배송 실무 보강판  ·  오가렌 경영기획본부", {
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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
    ["Part 4", "고객 응대 프로세스", "25"],
    ["Part 5", "세일즈 스크립트 & 거절 처리", "32"],
    ["Part 6", "TOP 매니저 케이스 & 모범 사례", "41"],
    ["Part 7", "운영 · KPI · 부록", "48"],
    ["Part 8", "인증 & 안전성", "53"],
    ["Part 9", "배송 실무 가이드  [NEW v2.1]", "66"],
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
    bold: true, margin: 0
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
    fontSize: 22, color: C.goldDeep, fontFace: FONT_H, bold: true, bold: true, margin: 0
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
      x, y: y + 1.0, w: 2.9, h: 2.3,
      fill: { color: C.white }, line: { color: C.line, width: 0.5 }
    });
    s.addText(`"${b.tagline}"`, {
      x: x + 0.2, y: y + 1.1, w: 2.5, h: 0.4,
      fontSize: 16, color: b.color, fontFace: FONT_H,
      bold: true, margin: 0
    });
    s.addText(b.target, {
      x: x + 0.2, y: y + 1.55, w: 2.5, h: 0.3,
      fontSize: 10, color: C.muted, fontFace: FONT_B,
      bold: true, charSpacing: 2, margin: 0
    });
    s.addText(b.desc, {
      x: x + 0.2, y: y + 1.95, w: 2.5, h: 1.3,
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  s.addText("우리 회사가 강한 이유 — 세 브랜드가 가격대 · 라이프스타일 · 가족 구성을 모두 커버합니다.", {
    x: 0.5, y: 4.9, w: 9, h: 0.35,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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

  // 브랜드 마커 (원) — 차트 박스(y=1.4~4.7) 안에 마커+sub까지 모두 들어오도록 배치
  const markers = [
    { x: cx + 5.8, y: cy + 0.3, w: 1.5, label: "SLEEPER", sub: "프리미엄 · 맞춤", color: C.ink },
    { x: cx + 2.7, y: cy + 0.6, w: 1.3, label: "nooer", sub: "합리적 · 모던", color: C.goldDeep },
    { x: cx + 1.4, y: cy + 1.5, w: 1.0, label: "Toddles", sub: "어린이 · 안전", color: C.blue },
    { x: cx + 4.4, y: cy + 1.7, w: 0.95, label: "(경쟁사)", sub: "스프링 중심", color: C.muted },
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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
      bold: true, margin: 0
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
      fontSize: 10, color: C.goldDeep, fontFace: FONT_B, bold: true, bold: true, margin: 0
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
    fontSize: 10, color: C.white, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
      fontSize: 10.5, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
    });
  });

  // 국제 표준 인사이트: 베개 = 수면 불만의 절반
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.95, w: 9, h: 0.45,
    fill: { color: C.cream }, line: { color: C.gold, width: 0.5 }
  });
  s.addText("국제 수면학회 인사이트", {
    x: 0.7, y: 4.95, w: 3.0, h: 0.45,
    fontSize: 10, color: C.goldDeep, fontFace: FONT_H,
    bold: true, charSpacing: 1.5, valign: "middle", margin: 0
  });
  s.addText("수면 불만의 50%는 매트리스가 아닌 \"베개\"가 원인. 매트리스만큼 베개를 정성껏 권하세요. 새 매트리스 + 묵은 베개 = 새 차에 펑크 난 타이어.", {
    x: 3.5, y: 4.95, w: 5.9, h: 0.45,
    fontSize: 9.5, color: C.ink, fontFace: FONT_B,
    bold: true, valign: "middle", margin: 0
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
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 16, color: C.goldDeep, fontFace: FONT_H, bold: true, bold: true, margin: 0
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
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 24, "Part 3 · 인테리어 매칭");
}

// ============================================================
// SLIDE 25 — Part 4 디바이더
// ============================================================
makeDivider("04", "고객 응대 프로세스",
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
      fontSize: 10, color: C.gold, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 26, "Part 4 · 페르소나");
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
      bold: true, bold: true, align: "center", margin: 0
    });
  });

  s.addText("핵심 — 5단계는 \"순서\"이지만 정해진 시간이 없습니다. 단계 사이에 머무르는 시간이 신뢰를 결정합니다.", {
    x: 0.5, y: 4.85, w: 9, h: 0.4,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 27, "Part 4 · 5단계 응대");
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
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
      x: 0.95, y, w: 6.5, h: 0.5,
      fontSize: 11.5, color: C.ink, fontFace: FONT_H, bold: true, valign: "middle", margin: 0
    });
    s.addText(l.note, {
      x: 7.55, y, w: 1.9, h: 0.5,
      fontSize: 9.5, color: C.goldDeep, fontFace: FONT_B, valign: "middle", margin: 0
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

  addFooter(s, 28, "Part 4 · 입장 3분");
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
    fontSize: 13, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
      x: x + 0.15, y: y + 0.5, w: 4.05, h: 0.4,
      fontSize: 10.5, color: C.goldDeep, fontFace: FONT_H, bold: true, margin: 0
    });
    s.addText("→ " + q.learn, {
      x: x + 0.2, y: y + 0.92, w: 4.0, h: 0.4,
      fontSize: 10, color: C.inkSoft, fontFace: FONT_B, margin: 0
    });
  });

  addFooter(s, 29, "Part 4 · 4질문");
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
    fontSize: 14, color: C.goldDeep, fontFace: FONT_H, bold: true, bold: true, margin: 0
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
      bold: true, valign: "middle", margin: 0
    });
  });

  // 글로벌 표준: 12분 룰 + 두 번 누워보기
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.7, w: 9, h: 0.65,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("글로벌 표준  ·  12분 룰 + 두 번 누워보기", {
    x: 0.7, y: 4.75, w: 5.5, h: 0.28,
    fontSize: 11, color: C.gold, fontFace: FONT_H,
    bold: true, charSpacing: 1.5, margin: 0
  });
  s.addText("매트리스는 최소 12분은 누워봐야 진짜 느낌이 옵니다. 처음에는 \"편하다\", 다음에는 \"어디가 어떻게 편하다\"가 보여요. 후보 2개는 반드시 두 번씩 누워보게 — 첫 체험은 인상, 두 번째 체험이 결정.", {
    x: 0.7, y: 5.02, w: 8.6, h: 0.3,
    fontSize: 9.5, color: C.white, fontFace: FONT_B, margin: 0
  });

  addFooter(s, 30, "Part 4 · 체험 유도");
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
    fontSize: 13, color: C.goldDeep, fontFace: FONT_H, bold: true, bold: true, margin: 0
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

  addFooter(s, 31, "Part 4 · 부부 응대");
}

// ============================================================
// SLIDE 32 — Part 5 디바이더
// ============================================================
makeDivider("05", "세일즈 스크립트 & 거절 처리",
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
      bold: true, valign: "middle", margin: 0
    });
    s.addText("→ 다음 페이지", {
      x: 8.3, y, w: 1.1, h: 0.42,
      fontSize: 9, color: C.muted, fontFace: FONT_B,
      align: "right", valign: "middle", margin: 0
    });
  });

  addFooter(s, 33, "Part 5 · 거절 TOP 7");
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
    bold: true, bold: true, margin: 0
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
    bold: true, margin: 0
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
    bold: true, margin: 0
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
  "가격을 깎지 마세요. 시간 단위로 환산하세요.  ·  8시간 × 365일 × 8년 = 23,360시간",
  "\"많이 비싸진 않아요\"\n\"이번 달 할인이라…\"\n→ 가격 방어는 가격 인정. 더 깎아야 한다는 메시지만 줍니다.",
  "\"8시간 × 365일 × 8년 = 23,360시간을 쓰시는 거예요. 시간당 100원이 안 돼요.\"\n또는\n\"하루 약 X백 원이세요. 매일 밤 좋은 잠을 보장받는 8년치 비용이라고 보시면 됩니다.\"",
  "Tempur·Sleep Number·Saatva 등 글로벌 매장 표준 전법 — 매트리스만큼 자주·오래 쓰는 가구는 없습니다. \"가격\"을 \"시간\"으로 환산하면 \"비싸다\"가 \"가성비 최강\"으로 바뀝니다.\n또한 \"수면 = 건강 = 일·관계의 질\". 가격이 아니라 투자로 보이게 만드세요. 절대 깎지 마세요."
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
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
    }
  );
  s.addText("WHY — \"이사할 때 다시\"는 90% 안 옵니다. 연락처를 얻고 능동적으로 추적하세요. CRM 입력 필수.", {
    x: 0.7, y: 2.65, w: 8.6, h: 0.5,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, bold: true, margin: 0
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
      fontSize: 11, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
    }
  );
  s.addText("WHY — \"멀쩡함\"은 \"바꿀 이유 없음\"이 아니라 \"바꿀 동기 부족\". 위생·토퍼 업셀로 부드럽게 진입.", {
    x: 0.7, y: 4.65, w: 8.6, h: 0.5,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, bold: true, margin: 0
  });

  addFooter(s, 39, "Part 5 · 거절 #6 #7");
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
      x: 2.7, y: y + 0.13, w: 3.5, h: 0.25,
      fontSize: 10, color: C.gold, fontFace: FONT_B, charSpacing: 1, margin: 0
    });
    s.addText(c.note, {
      x: 6.3, y: y + 0.13, w: 3.0, h: 0.25,
      fontSize: 9, color: C.muted, fontFace: FONT_B,
      align: "right", margin: 0
    });
    s.addText(c.line, {
      x: 0.7, y: y + 0.45, w: 8.6, h: 0.55,
      fontSize: 12, color: C.goldDeep, fontFace: FONT_H,
      bold: true, margin: 0
    });
  });

  s.addText("Tip — 클로징은 \"강요\"가 아니라 \"정리\"입니다. 고객 머릿속에 흩어진 선택지를 매니저가 한 번에 묶어주는 일.", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 40, "Part 5 · 클로징");
}

// ============================================================
// SLIDE 41 — Part 6 디바이더
// ============================================================
makeDivider("06", "TOP 매니저 케이스 & 모범 사례",
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 42, "Part 6 · 7가지 행동");
}

// ============================================================
// SLIDE 43 — 케이스 1: 부부 객단가 200% 사례
// ============================================================
function makeCaseSlide(num, page, title, situation, action, result, lesson) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  addSectionTitle(s, `CASE STUDY #${num}`, title);

  // 좌측: 상황 + 행동
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 5.3, h: 3.7,
    fill: { color: C.creamSoft }, line: { color: C.line, width: 0.5 }
  });
  s.addText("상황", {
    x: 0.7, y: 1.5, w: 5.0, h: 0.3,
    fontSize: 12, color: C.goldDeep, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(situation, {
    x: 0.7, y: 1.8, w: 5.0, h: 1.4,
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, margin: 0
  });

  s.addText("매니저의 행동", {
    x: 0.7, y: 3.2, w: 5.0, h: 0.3,
    fontSize: 12, color: C.goldDeep, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(action, {
    x: 0.7, y: 3.5, w: 5.0, h: 1.55,
    fontSize: 11, color: C.ink, fontFace: FONT_B, bold: true, margin: 0
  });

  // 우측: 결과 + 교훈
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.0, y: 1.4, w: 3.5, h: 1.7,
    fill: { color: C.ink }, line: { color: C.ink }
  });
  s.addText("결과", {
    x: 6.2, y: 1.5, w: 3.2, h: 0.3,
    fontSize: 12, color: C.gold, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(result, {
    x: 6.2, y: 1.85, w: 3.2, h: 1.2,
    fontSize: 11, color: C.white, fontFace: FONT_B, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.0, y: 3.25, w: 3.5, h: 1.85,
    fill: { color: C.cream }, line: { color: C.gold, width: 1 }
  });
  s.addText("우리가 배울 것", {
    x: 6.2, y: 3.35, w: 3.2, h: 0.3,
    fontSize: 12, color: C.goldDeep, fontFace: FONT_H,
    bold: true, charSpacing: 1, margin: 0
  });
  s.addText(lesson, {
    x: 6.2, y: 3.7, w: 3.2, h: 1.3,
    fontSize: 11, color: C.ink, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, page, `Part 6 · 케이스 #${num}`);
}

makeCaseSlide(
  1, 43, "부부 객단가 +180% — 시그니처 풀세트 성공",
  "30대 후반 신혼 부부. 처음에는 \"퀸 매트리스만 보러 왔어요\"라고 함. " +
  "예산도 정해두지 않았다고 강조. 둘러보기형 + 결혼 준비형 혼재.",
  "1) 예산 묻지 않고 \"신혼 첫 집이 어디세요?\" 부터 질문\n" +
  "2) 신축 아파트라는 답에 \"방에 맞춰 만드는 1:1 맞춤\" 강조\n" +
  "3) 두 사람 모두 누워보게 한 뒤 \"두 분 체중·자세가 달라서 시그니처 양면이 맞을 것\" 설명\n" +
  "4) 토퍼·베개까지 \"침실 완성\" 패키지 제안",
  "퀸 매트리스만 사려던 부부가 시그니처 커스텀(K) + 토퍼 + 베개 + 패브릭 프레임 풀세트 구매. " +
  "객단가 약 180% 상승. 6개월 후 부모님 침대까지 추천 구매.",
  "\"매트리스만 사러 왔어요\"는 \"매트리스부터 보여달라\"가 아닙니다.\n" +
  "신혼·이사 같은 라이프 이벤트는 \"풀세트의 기회\"."
);

// ============================================================
// SLIDE 44 — 케이스 2: 거절 후 재방문 클로징
// ============================================================
makeCaseSlide(
  2, 44, "\"더 알아볼게요\" → 2주 뒤 클로징",
  "50대 부부, 허리 통증 호소. 매장 방문 30분 뒤 \"좀 더 알아보고 올게요\". " +
  "타 브랜드 매장 비교 중인 게 명확. 평범한 매니저라면 \"네…\" 끝.",
  "1) 막지 않고 \"천천히 보세요\"\n" +
  "2) 비교표 1장 + 명함 + 본인 카톡 ID 챙겨드림\n" +
  "3) \"오늘 누워보신 한국형 고밀도 메모리폼 — 이거 한 번 더 와보시고 결정하세요\"\n" +
  "4) 2일 뒤 카톡으로 \"허리 어떠세요? 비교는 잘 되셨어요?\" 가벼운 안부\n" +
  "5) 1주일 뒤 \"매트리스 비교 가이드\" 추가 자료 발송",
  "2주 뒤 재방문 → 한국형 고밀도 메모리폼 K 사이즈 + 토퍼 구매. " +
  "고객 인터뷰에서 \"비교했지만 가장 친절했어서 다시 왔다\"고 답변.",
  "거절을 받아들이고 \"잔잔하게 따라가는\" 사후 응대가 재방문률을 결정.\n" +
  "압박 한 번이 잃은 고객을 \"기억해주는 매니저\"가 데려옵니다."
);

// ============================================================
// SLIDE 45 — 케이스 3: 둘러보기형 → 단골
// ============================================================
makeCaseSlide(
  3, 45, "둘러보기형 → 3개월 뒤 가족 침대 + 추천 고객 3명",
  "40대 여성, 자녀 데리고 매장 방문. \"그냥 한 번 보러…\" 시그널. " +
  "5분 만에 나갈 분위기. 다수 매니저는 \"네, 천천히 보세요\"로 끝.",
  "1) 자녀에게 토들즈 어린이 침대 직접 누워보게 권유\n" +
  "2) 어머니에게는 \"자녀 침대만이라도 카탈로그 챙겨드릴게요\"\n" +
  "3) 명함 + 토들즈 + 누어 카탈로그 3종 패키지\n" +
  "4) 1주 뒤 \"아이는 잘 자나요?\" 카톡 안부 (구매 압박 NO)\n" +
  "5) 본인 인스타그램 매장 스타일링 사진 공유",
  "3개월 뒤 토들즈 어린이 침대 구매. 6개월 뒤 본인 부부 침대 누어 룬드빅 추가. " +
  "이후 친구 3명 매장 방문 → 그중 2명 구매. 약 1년간 매니저 \"개인 고객\".",
  "\"둘러보기형\"이 가장 큰 기회 → 안 살 고객도 \"단골 + 추천원\"이 됩니다.\n" +
  "당장의 구매 아닌 \"기억\"을 파는 것."
);

// ============================================================
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
      bold: true, valign: "middle", margin: 0
    });
  });

  s.addText("→ 이 8개 멘트를 자기 말로 다시 적어 외워두세요. 매장 즉시 사용 가능한 \"현장 무기\"입니다.", {
    x: 0.5, y: 5.0, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, bold: true, margin: 0
  });

  addFooter(s, 46, "Part 6 · 베스트 멘트");
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
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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

  addFooter(s, 47, "Part 6 · 자기 진단");
}

// ============================================================
// SLIDE 48 — Part 7 디바이더
// ============================================================
makeDivider("07", "운영 · KPI · 부록",
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, bold: true, margin: 0
  });

  addFooter(s, 49, "Part 7 · 사후 관리");
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 50, "Part 7 · KPI");
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
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 51, "Part 7 · FAQ");
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

  addFooter(s, 52, "Part 7 · 본사 지원");
}

// ============================================================
// PART 8 — 인증 & 안전성 [NEW v2.0]
// ============================================================

// SLIDE 53 — Part 8 디바이더
makeDivider("08", "인증 & 안전성",
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
  let y = 1.25;
  const rowH = 0.31;

  headers.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y, w: colW[ci], h: 0.36,
      fill: { color: C.ink }, line: { color: C.ink }
    });
    s.addText(h, {
      x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: 0.36,
      fontSize: 11, color: C.gold, fontFace: FONT_H,
      bold: true, valign: "middle", charSpacing: 1, margin: 0
    });
  });
  y += 0.36;
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? C.creamSoft : C.white;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: rowH,
        fill: { color: fill }, line: { color: C.line, width: 0.4 }
      });
      s.addText(cell, {
        x: colX[ci] + 0.1, y, w: colW[ci] - 0.2, h: rowH,
        fontSize: 9.5, color: ci === 0 ? C.goldDeep : C.ink,
        fontFace: ci === 0 ? FONT_H : FONT_B,
        bold: ci === 0, valign: "middle", margin: 0
      });
    });
    y += rowH;
  });

  s.addText("매니저 핵심 — 이 표만 외워두면, 고객이 어떤 안전 질문을 해도 답할 수 있습니다.", {
    x: 0.5, y: 4.95, w: 9, h: 0.3,
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 54, "Part 8 · 인증 종합");
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
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 11, color: C.white, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 55, "Part 8 · 라돈");
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
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 56, "Part 8 · 폼 · 라텍스");
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
    fontSize: 11, color: C.gold, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 11, color: C.line, fontFace: FONT_B, bold: true, margin: 0
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

  addFooter(s, 57, "Part 8 · 공장·접착제");
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
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 58, "Part 8 · 누어 인증");
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
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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

  addFooter(s, 59, "Part 8 · 토들즈 안전");
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
    fontSize: 11, color: C.line, fontFace: FONT_B, bold: true, margin: 0
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

  addFooter(s, 60, "Part 8 · 토퍼·베개");
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
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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

  addFooter(s, 61, "Part 8 · 패브릭 인증");
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
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 10.5, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 63, "Part 8 · 등급 체계");
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
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
    bold: true, bold: true, align: "center", margin: 0
  });

  addFooter(s, 64, "Part 8 · 휴대 카드");
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
    fontSize: 11, color: C.goldDeep, fontFace: FONT_B, bold: true, bold: true, margin: 0
  });

  addFooter(s, 65, "Part 8 · 활용 원칙");
}


// ============================================================
// PART 9 — 배송 실무 가이드 [NEW v2.1]
// ============================================================

// SLIDE 66 — Part 9 디바이더
makeDivider("09", "배송 실무 가이드",
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
    fontSize: 12, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 10, color: C.white, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 67, "Part 9 · 엘리베이터");
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

  addFooter(s, 68, "Part 9 · 계단·사다리차·탈거");
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
    fontSize: 11, color: C.inkSoft, fontFace: FONT_B, bold: true, margin: 0
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
    fontSize: 10, color: C.goldDeep, fontFace: FONT_B, bold: true, margin: 0
  });

  addFooter(s, 69, "Part 9 · 오늘의집 조회");
}


// ============================================================
// APPENDIX — 경쟁사 비교 자료 (세션 2 전용 부록 / 영업 기밀)
// ============================================================
const appendixPages = [
  { src: "appendix_pdf_extract/page_01.png", label: "표지" },
  { src: "appendix_pdf_extract/page_02.png", label: "12종 시리즈별 특장 + 경도 정리" },
  { src: "appendix_pdf_extract/page_03.png", label: "시그니처 vs 시몬스 · 씰리" },
  { src: "appendix_pdf_extract/page_04.png", label: "하이브리드 vs 코웨이 · 씰리 · 베스트슬립" },
  { src: "appendix_pdf_extract/page_05.png", label: "스테이블 vs 씰리 · 코웨이 · 에이스 · 한셀 · 일룸" },
  { src: "appendix_pdf_extract/page_06.png", label: "밸런스드 vs 에이스 · 코웨이 · 일룸" },
];

appendixPages.forEach((p, i) => {
  const s = pres.addSlide();
  s.background = { color: C.white };
  // 상단 어펜딕스 라벨
  s.addText(`APPENDIX  ·  ${i+1} / ${appendixPages.length}  ·  ${p.label}`, {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 11, color: C.gold, fontFace: FONT_H,
    bold: true, charSpacing: 3, margin: 0
  });
  // PDF 페이지 풀 임베드 (16:9 contain → 8.18 x 4.6 centered)
  s.addImage({
    path: p.src,
    x: 0.3, y: 0.8, w: 9.4, h: 4.5,
    sizing: { type: "contain", w: 9.4, h: 4.5 }
  });
  addFooter(s, 0, "Appendix · 경쟁사 비교 (영업 기밀)");
});

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
      bold: true, margin: 0
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

// ============================================================
// 저장
// ============================================================
pres.writeFile({ fileName: "오가렌_매니저_교육_세일즈_가이드_v2.1.pptx" })
  .then(fn => console.log("DONE:", fn))
  .catch(err => { console.error("ERROR:", err); process.exit(1); });
