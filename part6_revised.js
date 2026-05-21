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

  addFooter(s, 43, "Part 6 · 박은희 (롯데광주)");
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

  addFooter(s, 44, "Part 6 · 한용진 (광주쇼룸)");
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

  addFooter(s, 45, "Part 6 · 시그니처 5단계");
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

  addFooter(s, 46, "Part 6 · 매장별 벤치마크");
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

  addFooter(s, 49, "Part 6 · 월간 업데이트");
}

