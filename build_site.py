"""
오가렌 세일즈가이드 정적 사이트 빌더
- PPT → PNG 슬라이드 변환 (PowerPoint COM)
- 슬라이드 제목 추출 (python-pptx)
- index.html / session1.html / session2.html 생성

매월 갱신 워크플로우:
  1) part6_revised.js 수정 → node split_and_build.js
  2) python build_site.py
  3) git add . && git commit && git push
"""
from __future__ import annotations
import os
import re
import json
import shutil
import time
from pathlib import Path

import pythoncom
import win32com.client
from pptx import Presentation

ROOT = Path(__file__).parent
SITE = ROOT  # GitHub Pages는 repo root에서 서빙
ASSETS_S1 = SITE / "assets" / "session1"
ASSETS_S2 = SITE / "assets" / "session2"
ASSETS_DL = SITE / "assets" / "download"

S1_PPT = ROOT / "오가렌_세일즈가이드_세션1.pptx"
S2_PPT = ROOT / "오가렌_세일즈가이드_세션2.pptx"
FULL_PPT = ROOT / "오가렌_매니저_교육_세일즈_가이드_v2.1.pptx"

# 세션별 TOC 정의 (제목 + 슬라이드 범위)
# 슬라이드 번호는 1-indexed, 세션 PPT 내 기준
SESSION1_TOC = [
    {"id": "cover",  "label": "표지",                 "from": 1, "to": 1},
    {"id": "usage",  "label": "본 가이드 사용법",      "from": 2, "to": 2},
    {"id": "toc",    "label": "목차",                 "from": 3, "to": 3},
    {"id": "part1",  "label": "Part 1 · 오가렌과 3대 브랜드", "from": 4,  "to": 11},
    {"id": "part2",  "label": "Part 2 · 매트리스 마스터하기", "from": 12, "to": 19},
    {"id": "part3",  "label": "Part 3 · 침대 프레임 마스터하기", "from": 20, "to": 24},
    {"id": "part4",  "label": "Part 4 · 인증 & 안전성", "from": 25, "to": 37},
    {"id": "end",    "label": "마무리",               "from": 38, "to": 99},
]

SESSION2_TOC = [
    {"id": "cover",  "label": "표지",                 "from": 1, "to": 1},
    {"id": "usage",  "label": "본 가이드 사용법",      "from": 2, "to": 2},
    {"id": "toc",    "label": "목차",                 "from": 3, "to": 3},
    {"id": "part1",  "label": "Part 1 · 고객 응대 프로세스", "from": 4,  "to": 10},
    {"id": "part2",  "label": "Part 2 · 세일즈 스크립트 & 거절 처리", "from": 11, "to": 19},
    {"id": "part3",  "label": "Part 3 · TOP 매니저 케이스 & 모범 사례", "from": 20, "to": 28},
    {"id": "part4",  "label": "Part 4 · 운영 · KPI · 부록", "from": 29, "to": 33},
    {"id": "part5",  "label": "Part 5 · 배송 실무 가이드", "from": 34, "to": 37},
    {"id": "end",    "label": "마무리",               "from": 38, "to": 99},
]


def ppt_export_pngs(ppt_path: Path, out_dir: Path, width: int = 1600) -> int:
    """PowerPoint COM을 통해 슬라이드를 PNG로 export. 반환: 슬라이드 개수."""
    out_dir.mkdir(parents=True, exist_ok=True)
    for f in out_dir.glob("*.png"):
        f.unlink()

    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = 1  # 일부 버전에서 헤드리스 모드가 거부됨
        try:
            deck = app.Presentations.Open(str(ppt_path), WithWindow=False)
        except Exception:
            deck = app.Presentations.Open(str(ppt_path))
        n = deck.Slides.Count
        height = int(width * 9 / 16)
        deck.SaveCopyAs(str(out_dir / "_export"), 18, EmbedTrueTypeFonts=False)  # 17=PNG, 18=GIF? — use Export
        # 위 SaveCopyAs는 ppt별 차이가 있어 Slide.Export 직접 호출이 더 안정적
        # 정리 후 Slide.Export 사용
        for tmp in out_dir.glob("_export*"):
            if tmp.is_dir():
                shutil.rmtree(tmp, ignore_errors=True)
            else:
                tmp.unlink(missing_ok=True)
        for i in range(1, n + 1):
            slide = deck.Slides(i)
            target = out_dir / f"slide_{i:02d}.png"
            slide.Export(str(target), "PNG", width, height)
        deck.Close()
        return n
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def extract_titles(ppt_path: Path) -> list[str]:
    """python-pptx로 각 슬라이드의 제목 텍스트 추출. 없으면 빈 문자열."""
    titles: list[str] = []
    pres = Presentation(str(ppt_path))
    for slide in pres.slides:
        title = ""
        # 1순위: title placeholder
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            title = (slide.shapes.title.text_frame.text or "").strip()
        if not title:
            # 2순위: 가장 위쪽의 큰 폰트 텍스트
            best = None
            for shp in slide.shapes:
                if not shp.has_text_frame:
                    continue
                txt = (shp.text_frame.text or "").strip()
                if not txt:
                    continue
                # 첫 줄만
                first_line = txt.splitlines()[0].strip()
                # 너무 긴 거(>40자) 제외
                if len(first_line) > 60:
                    continue
                key = (shp.top or 0, -(shp.width or 0))
                if best is None or key < best[0]:
                    best = (key, first_line)
            if best:
                title = best[1]
        # 정리
        title = re.sub(r"\s+", " ", title).strip()
        titles.append(title)
    return titles


# ============================================================
# HTML 생성
# ============================================================

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --brand: #1F4E79;
  --brand2: #2E75B6;
  --accent: #C5504B;
  --bg: #F7F9FC;
  --card: #FFFFFF;
  --text: #1B2737;
  --muted: #5B6B80;
  --border: #E1E8F0;
  --shadow: 0 4px 16px rgba(31,78,121,0.08);
  --shadow-lg: 0 12px 32px rgba(31,78,121,0.18);
}
html, body { font-family: 'Malgun Gothic','Apple SD Gothic Neo','Pretendard',system-ui,sans-serif; color: var(--text); background: var(--bg); line-height: 1.55; -webkit-font-smoothing: antialiased; }
a { color: var(--brand2); text-decoration: none; }
a:hover { text-decoration: underline; }
.container { max-width: 1240px; margin: 0 auto; padding: 0 24px; }

/* ===== Index ===== */
.hero { background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%); color: #fff; padding: 72px 0 64px; }
.hero .container { display: flex; align-items: flex-end; justify-content: space-between; gap: 32px; flex-wrap: wrap; }
.hero h1 { font-size: 38px; font-weight: 700; letter-spacing: -0.5px; margin-bottom: 8px; }
.hero h1 .accent { color: #FFD966; }
.hero .sub { font-size: 16px; opacity: 0.85; }
.hero .meta { font-size: 13px; opacity: 0.75; text-align: right; }
.hero .meta b { font-weight: 600; }

.sessions { padding: 56px 0; }
.sessions h2 { font-size: 24px; font-weight: 700; margin-bottom: 8px; color: var(--brand); }
.sessions .lead { color: var(--muted); margin-bottom: 32px; font-size: 15px; }
.session-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
@media (max-width: 800px) { .session-grid { grid-template-columns: 1fr; } }
.session-card { background: var(--card); border: 1px solid var(--border); border-radius: 16px; padding: 32px; box-shadow: var(--shadow); transition: transform .15s, box-shadow .15s; cursor: pointer; display: block; }
.session-card:hover { transform: translateY(-4px); box-shadow: var(--shadow-lg); text-decoration: none; }
.session-card .num { font-size: 13px; font-weight: 600; color: var(--brand2); letter-spacing: 1px; text-transform: uppercase; }
.session-card h3 { font-size: 26px; font-weight: 700; color: var(--brand); margin: 8px 0 12px; }
.session-card .tag { display: inline-block; background: #F0F4FA; color: var(--brand); font-size: 12px; font-weight: 600; padding: 4px 10px; border-radius: 999px; margin-bottom: 16px; }
.session-card .parts { list-style: none; margin-top: 12px; }
.session-card .parts li { font-size: 14px; color: var(--muted); padding: 4px 0; }
.session-card .parts li::before { content: "·"; color: var(--brand2); font-weight: 700; margin-right: 8px; }
.session-card .footer { margin-top: 20px; display: flex; justify-content: space-between; align-items: center; }
.session-card .count { font-size: 13px; color: var(--muted); }
.session-card .go { font-size: 14px; font-weight: 600; color: var(--brand2); }

.downloads { padding: 24px 0 80px; }
.downloads .panel { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 24px; box-shadow: var(--shadow); }
.downloads h3 { font-size: 17px; color: var(--brand); margin-bottom: 12px; }
.downloads ul { list-style: none; }
.downloads li { padding: 6px 0; font-size: 14px; }
.downloads li::before { content: "↓ "; color: var(--brand2); font-weight: 700; }

.footer { padding: 24px 0; text-align: center; font-size: 12px; color: var(--muted); border-top: 1px solid var(--border); background: var(--card); }

/* ===== Session viewer ===== */
.topbar { background: var(--card); border-bottom: 1px solid var(--border); padding: 14px 0; position: sticky; top: 0; z-index: 10; box-shadow: 0 2px 8px rgba(0,0,0,0.03); }
.topbar .container { display: flex; align-items: center; justify-content: space-between; gap: 16px; }
.topbar .back { font-size: 14px; color: var(--muted); }
.topbar .title { font-size: 16px; font-weight: 700; color: var(--brand); }
.topbar .download { font-size: 13px; padding: 6px 12px; background: var(--brand); color: #fff; border-radius: 6px; font-weight: 600; }
.topbar .download:hover { background: var(--brand2); text-decoration: none; }

.layout { display: grid; grid-template-columns: 240px 1fr; gap: 32px; padding: 32px 0; }
@media (max-width: 900px) { .layout { grid-template-columns: 1fr; } }

.sidebar { position: sticky; top: 80px; align-self: start; background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 20px; box-shadow: var(--shadow); max-height: calc(100vh - 100px); overflow-y: auto; }
.sidebar h4 { font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: var(--muted); margin-bottom: 12px; }
.sidebar ul { list-style: none; }
.sidebar li a { display: block; padding: 8px 10px; font-size: 13.5px; color: var(--text); border-radius: 6px; border-left: 3px solid transparent; }
.sidebar li a:hover { background: #F0F4FA; color: var(--brand); text-decoration: none; border-left-color: var(--brand2); }

.main { min-width: 0; }
.part-section { margin-bottom: 48px; scroll-margin-top: 80px; }
.part-section h2 { font-size: 22px; font-weight: 700; color: var(--brand); margin-bottom: 4px; }
.part-section .part-desc { font-size: 13px; color: var(--muted); margin-bottom: 20px; }

.slide-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 18px; }
@media (max-width: 700px) { .slide-grid { grid-template-columns: 1fr; } }
.slide-card { background: var(--card); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; box-shadow: var(--shadow); cursor: zoom-in; transition: transform .15s, box-shadow .15s; }
.slide-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-lg); }
.slide-card .thumb { display: block; width: 100%; aspect-ratio: 16/9; object-fit: cover; background: #f0f0f0; }
.slide-card .meta { padding: 10px 14px; font-size: 13px; color: var(--muted); display: flex; justify-content: space-between; align-items: center; gap: 8px; }
.slide-card .meta b { color: var(--text); font-weight: 600; font-size: 13.5px; flex: 1; }
.slide-card .meta .n { font-size: 11px; background: #F0F4FA; color: var(--brand); padding: 2px 8px; border-radius: 999px; font-weight: 600; }

/* Lightbox */
.lb { position: fixed; inset: 0; background: rgba(15,25,40,0.92); display: none; align-items: center; justify-content: center; z-index: 100; padding: 24px; }
.lb.open { display: flex; }
.lb img { max-width: 100%; max-height: 90vh; box-shadow: 0 30px 80px rgba(0,0,0,0.5); border-radius: 6px; }
.lb .close, .lb .nav { position: absolute; color: #fff; background: rgba(255,255,255,0.12); border: none; cursor: pointer; font-size: 22px; width: 48px; height: 48px; border-radius: 999px; display: flex; align-items: center; justify-content: center; }
.lb .close:hover, .lb .nav:hover { background: rgba(255,255,255,0.22); }
.lb .close { top: 20px; right: 20px; }
.lb .prev { left: 20px; top: 50%; transform: translateY(-50%); }
.lb .next { right: 20px; top: 50%; transform: translateY(-50%); }
.lb .counter { position: absolute; bottom: 20px; left: 50%; transform: translateX(-50%); color: #fff; font-size: 13px; opacity: 0.7; }
"""

LIGHTBOX_JS = """
(function(){
  const cards = Array.from(document.querySelectorAll('.slide-card'));
  if (!cards.length) return;
  const lb = document.createElement('div');
  lb.className = 'lb';
  lb.innerHTML = `
    <button class="close" aria-label="닫기">×</button>
    <button class="nav prev" aria-label="이전">‹</button>
    <img alt="" />
    <button class="nav next" aria-label="다음">›</button>
    <div class="counter"></div>
  `;
  document.body.appendChild(lb);
  const img = lb.querySelector('img');
  const counter = lb.querySelector('.counter');
  let idx = 0;
  function show(i){
    idx = (i + cards.length) % cards.length;
    const src = cards[idx].querySelector('img').getAttribute('data-full') || cards[idx].querySelector('img').src;
    img.src = src;
    counter.textContent = (idx+1) + ' / ' + cards.length;
  }
  function open(i){ show(i); lb.classList.add('open'); document.body.style.overflow = 'hidden'; }
  function close(){ lb.classList.remove('open'); document.body.style.overflow = ''; img.src=''; }
  cards.forEach((c, i) => c.addEventListener('click', () => open(i)));
  lb.querySelector('.close').onclick = close;
  lb.querySelector('.prev').onclick = (e) => { e.stopPropagation(); show(idx-1); };
  lb.querySelector('.next').onclick = (e) => { e.stopPropagation(); show(idx+1); };
  lb.addEventListener('click', (e) => { if (e.target === lb) close(); });
  document.addEventListener('keydown', (e) => {
    if (!lb.classList.contains('open')) return;
    if (e.key === 'Escape') close();
    if (e.key === 'ArrowLeft') show(idx-1);
    if (e.key === 'ArrowRight') show(idx+1);
  });
})();
"""


def html_index(s1_count: int, s2_count: int, updated: str) -> str:
    return f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>오가렌 매장 매니저 교육 · 세일즈 가이드</title>
<meta name="description" content="오가렌 오프라인 쇼룸 매니저 교육과 세일즈 가이드. 우리는 그래서 잘 잔다(세션1) + 그래서 잘 판다(세션2)." />
<style>{CSS}</style>
</head>
<body>
  <section class="hero">
    <div class="container">
      <div>
        <h1>오가렌 매장 매니저<br/><span class="accent">교육 · 세일즈 가이드</span></h1>
        <p class="sub">오프라인 쇼룸 운영을 위한 통합 교육 자료</p>
      </div>
      <div class="meta">
        <div><b>Edition</b> v2.1</div>
        <div><b>Updated</b> {updated}</div>
      </div>
    </div>
  </section>

  <section class="sessions">
    <div class="container">
      <h2>두 세션으로 구성</h2>
      <p class="lead">세션 1에서 제품과 안전성을 익히고, 세션 2에서 실제 응대·세일즈·운영을 배웁니다.</p>
      <div class="session-grid">
        <a class="session-card" href="session1.html">
          <div class="num">Session 01</div>
          <h3>우리는 그래서 잘 잔다</h3>
          <span class="tag">제품 · 인증 · 안전성</span>
          <ul class="parts">
            <li>Part 1 · 오가렌과 3대 브랜드</li>
            <li>Part 2 · 매트리스 마스터하기</li>
            <li>Part 3 · 침대 프레임 마스터하기</li>
            <li>Part 4 · 인증 & 안전성</li>
          </ul>
          <div class="footer"><span class="count">{s1_count}개 슬라이드</span><span class="go">바로 보기 →</span></div>
        </a>
        <a class="session-card" href="session2.html">
          <div class="num">Session 02</div>
          <h3>그래서 잘 판다</h3>
          <span class="tag">응대 · 세일즈 · 운영</span>
          <ul class="parts">
            <li>Part 1 · 고객 응대 프로세스</li>
            <li>Part 2 · 세일즈 스크립트 & 거절 처리</li>
            <li>Part 3 · TOP 매니저 케이스 (실명)</li>
            <li>Part 4 · 운영 · KPI · 부록</li>
            <li>Part 5 · 배송 실무 가이드</li>
          </ul>
          <div class="footer"><span class="count">{s2_count}개 슬라이드</span><span class="go">바로 보기 →</span></div>
        </a>
      </div>
    </div>
  </section>

  <section class="downloads">
    <div class="container">
      <div class="panel">
        <h3>원본 PPT 다운로드</h3>
        <ul>
          <li><a href="assets/download/오가렌_세일즈가이드_세션1.pptx">세션 1 PPT (우리는 그래서 잘 잔다)</a></li>
          <li><a href="assets/download/오가렌_세일즈가이드_세션2.pptx">세션 2 PPT (그래서 잘 판다)</a></li>
          <li><a href="assets/download/오가렌_매니저_교육_세일즈_가이드_v2.1.pptx">통합본 (v2.1, 전체 70슬라이드)</a></li>
        </ul>
      </div>
    </div>
  </section>

  <footer class="footer">
    <div class="container">© 오가렌 DX팀 · 매니저 교육 · v2.1 · last updated {updated}</div>
  </footer>
</body>
</html>
"""


def html_session(session_name: str, session_label: str, subtitle: str, toc: list, titles: list[str], asset_dir: str, ppt_name: str) -> str:
    # 섹션별 슬라이드 정리
    parts_html: list[str] = []
    sidebar_html: list[str] = []
    for sec in toc:
        slides_in_section = []
        for i in range(sec["from"], min(sec["to"], len(titles)) + 1):
            t = titles[i - 1] if i - 1 < len(titles) else ""
            slides_in_section.append((i, t))
        if not slides_in_section:
            continue
        cards = []
        for i, t in slides_in_section:
            t_display = t if t else f"슬라이드 {i}"
            cards.append(
                f'<div class="slide-card">'
                f'<img class="thumb" loading="lazy" '
                f'src="{asset_dir}/slide_{i:02d}.png" '
                f'data-full="{asset_dir}/slide_{i:02d}.png" '
                f'alt="slide {i}" />'
                f'<div class="meta"><b>{t_display}</b><span class="n">#{i:02d}</span></div>'
                f'</div>'
            )
        parts_html.append(
            f'<section class="part-section" id="{sec["id"]}">'
            f'<h2>{sec["label"]}</h2>'
            f'<div class="part-desc">슬라이드 {sec["from"]:02d}–{min(sec["to"], len(titles)):02d}</div>'
            f'<div class="slide-grid">{"".join(cards)}</div>'
            f'</section>'
        )
        sidebar_html.append(f'<li><a href="#{sec["id"]}">{sec["label"]}</a></li>')

    return f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{session_label} · {subtitle} | 오가렌 매니저 교육</title>
<meta name="description" content="오가렌 매니저 교육 · {session_label} · {subtitle}" />
<style>{CSS}</style>
</head>
<body>
  <div class="topbar">
    <div class="container">
      <a href="index.html" class="back">← 메인으로</a>
      <div class="title">{session_label} · {subtitle}</div>
      <a href="assets/download/{ppt_name}" class="download" download>PPT 다운로드</a>
    </div>
  </div>

  <div class="container">
    <div class="layout">
      <aside class="sidebar">
        <h4>목차</h4>
        <ul>{"".join(sidebar_html)}</ul>
      </aside>
      <main class="main">
        {"".join(parts_html)}
      </main>
    </div>
  </div>

  <footer class="footer">
    <div class="container">© 오가렌 DX팀 · 매니저 교육 · v2.1</div>
  </footer>
  <script>{LIGHTBOX_JS}</script>
</body>
</html>
"""


def main():
    SITE.mkdir(exist_ok=True)
    ASSETS_S1.mkdir(parents=True, exist_ok=True)
    ASSETS_S2.mkdir(parents=True, exist_ok=True)
    ASSETS_DL.mkdir(parents=True, exist_ok=True)

    print("[1/4] 세션 1 PPT → PNG …")
    s1_count = ppt_export_pngs(S1_PPT, ASSETS_S1)
    print(f"    {s1_count} slides")

    print("[2/4] 세션 2 PPT → PNG …")
    s2_count = ppt_export_pngs(S2_PPT, ASSETS_S2)
    print(f"    {s2_count} slides")

    print("[3/4] 슬라이드 제목 추출 …")
    s1_titles = extract_titles(S1_PPT)
    s2_titles = extract_titles(S2_PPT)

    print("[4/4] HTML 빌드 + PPT 사본 복사 …")
    for src in [S1_PPT, S2_PPT, FULL_PPT]:
        if src.exists():
            shutil.copy2(src, ASSETS_DL / src.name)

    updated = time.strftime("%Y-%m-%d")

    (SITE / "index.html").write_text(
        html_index(s1_count, s2_count, updated), encoding="utf-8"
    )
    (SITE / "session1.html").write_text(
        html_session("session1", "세션 1", "우리는 그래서 잘 잔다",
                     SESSION1_TOC, s1_titles, "assets/session1",
                     S1_PPT.name),
        encoding="utf-8",
    )
    (SITE / "session2.html").write_text(
        html_session("session2", "세션 2", "그래서 잘 판다",
                     SESSION2_TOC, s2_titles, "assets/session2",
                     S2_PPT.name),
        encoding="utf-8",
    )

    # 메타 JSON (디버그용)
    (SITE / "_meta.json").write_text(
        json.dumps({
            "updated": updated,
            "session1": {"count": s1_count, "titles": s1_titles},
            "session2": {"count": s2_count, "titles": s2_titles},
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"\nDONE -> ./index.html")
    print(f"  session1: {s1_count} slides")
    print(f"  session2: {s2_count} slides")


if __name__ == "__main__":
    main()
