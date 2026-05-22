# -*- coding: utf-8 -*-
"""세션 3 HTML 빌더. session1/2와 동일한 슬라이드 그리드 + 라이트박스.

입력: _session3_slides.json (toc + 슬라이드 메타)
출력: session3.html
"""
from __future__ import annotations
import json
from pathlib import Path

ROOT = Path(__file__).parent
data = json.loads((ROOT / "_session3_slides.json").read_text(encoding="utf-8"))
toc = data["toc"]
total = data["total"]

# 사이드바
sidebar_items = []
for i, part in enumerate(toc):
    href = f"#{part['id']}"
    label = part["label"]
    # "Part 01 · 시그니처 매트리스" → pnum + label
    if "·" in label and label.startswith("Part"):
        pnum, rest = label.split("·", 1)
        sidebar_items.append(
            f'<li><a href="{href}"><span class="pnum">{pnum.strip().replace("Part ", "")}</span>{rest.strip()}</a></li>'
        )
    else:
        sidebar_items.append(f'<li><a href="{href}">{label}</a></li>')
sidebar = "".join(sidebar_items)

# 본문 섹션
sections = []
for part in toc:
    pid = part["id"]
    label = part["label"]
    fr, to = part["from"], part["to"]
    cards = []
    for sl in part["slides"]:
        n = sl["num"]
        title = sl["title"]
        cards.append(
            f'<div class="slide-card">'
            f'<div class="thumb-wrap">'
            f'<img class="thumb" loading="lazy" src="assets/session3/slide_{n:03d}.png" '
            f'data-full="assets/session3/slide_{n:03d}.png" alt="slide {n}" />'
            f'<span class="num-badge">{n:02d}</span>'
            f"</div>"
            f'<div class="meta"><b>{title}</b><span class="view">View</span></div>'
            f"</div>"
        )
    # part head: pnum 분리
    if "·" in label and label.startswith("Part"):
        pnum, rest = label.split("·", 1)
        head = (
            f"<div class=\"part-head\"><span class='pnum'>{pnum.strip()}</span>"
            f'<h2>{rest.strip()}</h2><div class="range">{fr:02d} — {to:02d}</div></div>'
        )
    else:
        head = f'<div class="part-head"><h2>{label}</h2><div class="range">{fr:02d} — {to:02d}</div></div>'
    sections.append(
        f'<section class="part-section reveal" id="{pid}">{head}<div class="slide-grid">{"".join(cards)}</div></section>'
    )

main_html = "".join(sections)

# 같은 CSS를 그대로 가져옴 (session2의 <style>...</style> 본문 동일)
CSS = """@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700;800;900&family=Cormorant+Garamond:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap');

* { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg: #0A0A0B;
  --bg-2: #111114;
  --surface: #15151A;
  --surface-2: #1C1C22;
  --border: rgba(255,255,255,0.08);
  --border-strong: rgba(255,255,255,0.14);
  --text: #F4F2EE;
  --text-2: #B8B5AE;
  --text-3: #76736D;
  --accent: #E8DCC4;
  --accent-warm: #C9B894;
  --serif: 'Cormorant Garamond', 'Noto Serif KR', 'Apple SD Gothic Neo', serif;
  --sans: 'Pretendard','Apple SD Gothic Neo','Malgun Gothic',system-ui,sans-serif;
}
html, body { font-family: var(--sans); color: var(--text); background: var(--bg); line-height: 1.6; -webkit-font-smoothing: antialiased; font-feature-settings: 'ss01','ss02'; }
body { background:
  radial-gradient(1200px 600px at 80% -10%, rgba(232,220,196,0.06), transparent 60%),
  radial-gradient(900px 500px at 0% 10%, rgba(201,184,148,0.04), transparent 60%),
  var(--bg);
  min-height: 100vh;
}
a { color: inherit; text-decoration: none; }
img { display: block; }
.container { max-width: 1320px; margin: 0 auto; padding: 0 32px; }
@media (max-width: 640px) { .container { padding: 0 20px; } }

.topbar { background: rgba(10,10,11,0.85); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px); border-bottom: 1px solid var(--border); padding: 18px 0; position: sticky; top: 0; z-index: 20; }
.topbar .container { display: flex; align-items: center; justify-content: space-between; gap: 16px; }
.topbar .back { font-size: 12px; letter-spacing: 1.5px; text-transform: uppercase; color: var(--text-2); display: inline-flex; align-items: center; gap: 10px; transition: color .25s ease; }
.topbar .back:hover { color: var(--accent); }
.topbar .title { font-family: var(--serif); font-size: 20px; font-weight: 500; letter-spacing: -0.3px; }
.topbar .title em { font-style: normal; color: var(--accent); }
.topbar .download { font-size: 11px; letter-spacing: 2px; text-transform: uppercase; padding: 10px 18px; border: 1px solid var(--border-strong); border-radius: 2px; transition: border-color .25s ease, color .25s ease, background .25s ease; color: var(--text); }
.topbar .download:hover { border-color: var(--accent-warm); color: var(--accent); background: var(--surface); }

.session-hero { padding: 80px 0 40px; }
.session-hero .kicker { font-size: 11px; letter-spacing: 3px; text-transform: uppercase; color: var(--accent-warm); margin-bottom: 24px; }
.session-hero h1 { font-family: var(--serif); font-size: clamp(40px, 6vw, 72px); font-weight: 500; line-height: 1.05; letter-spacing: -1px; }
.session-hero h1 em { font-style: normal; color: var(--accent); font-weight: 400; }

.layout { display: grid; grid-template-columns: 240px 1fr; gap: 56px; padding: 32px 0 80px; }
@media (max-width: 980px) { .layout { grid-template-columns: 1fr; gap: 24px; } }

.sidebar { position: sticky; top: 90px; align-self: start; max-height: calc(100vh - 110px); overflow-y: auto; padding-right: 12px; }
.sidebar::-webkit-scrollbar { width: 4px; }
.sidebar::-webkit-scrollbar-thumb { background: var(--border-strong); border-radius: 2px; }
.sidebar h4 { font-size: 10px; letter-spacing: 3px; text-transform: uppercase; color: var(--text-3); margin-bottom: 18px; }
.sidebar ul { list-style: none; }
.sidebar li a { display: block; padding: 12px 0; font-size: 13px; color: var(--text-2); border-bottom: 1px solid var(--border); transition: color .25s ease, padding-left .25s ease; line-height: 1.4; }
.sidebar li a:hover { color: var(--accent); padding-left: 8px; }
.sidebar li a .pnum { font-family: var(--serif); font-style: normal; color: var(--accent-warm); margin-right: 10px; font-size: 12px; letter-spacing: 1px; }

.main { min-width: 0; }
.part-section { margin-bottom: 80px; scroll-margin-top: 100px; }
.part-section .part-head { margin-bottom: 32px; display: flex; align-items: baseline; gap: 24px; flex-wrap: wrap; padding-bottom: 20px; border-bottom: 1px solid var(--border); }
.part-section .part-head .pnum { font-family: var(--serif); font-style: normal; font-size: 18px; color: var(--accent-warm); letter-spacing: 1px; }
.part-section .part-head h2 { font-family: var(--serif); font-size: 36px; font-weight: 500; letter-spacing: -0.5px; line-height: 1.1; flex: 1; min-width: 0; }
.part-section .part-head .range { font-size: 11px; letter-spacing: 2px; text-transform: uppercase; color: var(--text-3); }

.slide-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; }
@media (max-width: 700px) { .slide-grid { grid-template-columns: 1fr; } }
.slide-card { position: relative; background: var(--surface); border: 1px solid var(--border); border-radius: 4px; overflow: hidden; cursor: zoom-in; transition: border-color .35s ease, transform .35s ease; }
.slide-card:hover { border-color: var(--accent-warm); transform: translateY(-3px); }
.slide-card .thumb-wrap { position: relative; overflow: hidden; }
.slide-card .thumb { width: 100%; aspect-ratio: 16/9; object-fit: cover; background: #1a1a1f; transition: transform .6s cubic-bezier(.2,.6,.2,1); }
.slide-card:hover .thumb { transform: scale(1.04); }
.slide-card .num-badge { position: absolute; top: 14px; left: 14px; font-family: var(--serif); font-style: normal; font-size: 13px; color: var(--accent); background: rgba(10,10,11,0.7); backdrop-filter: blur(8px); padding: 4px 10px; border-radius: 2px; letter-spacing: 1px; }
.slide-card .meta { padding: 16px 18px; display: flex; justify-content: space-between; align-items: center; gap: 10px; border-top: 1px solid var(--border); }
.slide-card .meta b { color: var(--text); font-size: 13px; font-weight: 500; flex: 1; letter-spacing: 0.3px; }
.slide-card .meta .view { font-size: 10px; letter-spacing: 2px; text-transform: uppercase; color: var(--text-3); }

.lb { position: fixed; inset: 0; background: rgba(8,8,10,0.96); backdrop-filter: blur(8px); display: none; align-items: center; justify-content: center; z-index: 100; padding: 32px; }
.lb.open { display: flex; }
.lb img { max-width: 100%; max-height: 88vh; box-shadow: 0 40px 100px rgba(0,0,0,0.7); border: 1px solid var(--border); }
.lb .close, .lb .nav { position: absolute; color: var(--text); background: rgba(255,255,255,0.06); border: 1px solid var(--border-strong); cursor: pointer; font-size: 18px; width: 48px; height: 48px; border-radius: 999px; display: flex; align-items: center; justify-content: center; transition: background .25s ease, border-color .25s ease; backdrop-filter: blur(10px); }
.lb .close:hover, .lb .nav:hover { background: rgba(232,220,196,0.1); border-color: var(--accent-warm); color: var(--accent); }
.lb .close { top: 28px; right: 28px; }
.lb .prev { left: 28px; top: 50%; transform: translateY(-50%); }
.lb .next { right: 28px; top: 50%; transform: translateY(-50%); }
.lb .counter { position: absolute; bottom: 32px; left: 50%; transform: translateX(-50%); font-family: var(--serif); font-style: normal; font-size: 14px; color: var(--text-2); letter-spacing: 2px; }
.lb .counter b { color: var(--accent); font-weight: 500; font-style: normal; }

.footer { padding: 40px 0; border-top: 1px solid var(--border); margin-top: 60px; }
.footer .container { display: flex; justify-content: space-between; align-items: center; gap: 16px; flex-wrap: wrap; font-size: 12px; letter-spacing: 1px; color: var(--text-3); }
.footer .brand-mark { font-family: var(--serif); font-size: 14px; letter-spacing: 1px; color: var(--text-2); }

.reveal { opacity: 0; transform: translateY(20px); transition: opacity .8s ease, transform .8s ease; }
.reveal.in { opacity: 1; transform: translateY(0); }"""

JS = """(function(){
  const cards = Array.from(document.querySelectorAll('.slide-card'));
  if (cards.length) {
    const lb = document.createElement('div');
    lb.className = 'lb';
    lb.innerHTML = `
      <button class="close" aria-label="닫기">×</button>
      <button class="nav prev" aria-label="이전">‹</button>
      <img alt="" />
      <button class="nav next" aria-label="다음">›</button>
      <div class="counter"><b></b> <span class="sep">/</span> <span class="total"></span></div>
    `;
    document.body.appendChild(lb);
    const img = lb.querySelector('img');
    const cur = lb.querySelector('.counter b');
    const tot = lb.querySelector('.counter .total');
    let idx = 0;
    function show(i){
      idx = (i + cards.length) % cards.length;
      const im = cards[idx].querySelector('img');
      const src = im.getAttribute('data-full') || im.src;
      img.src = src;
      cur.textContent = String(idx+1).padStart(2,'0');
      tot.textContent = String(cards.length).padStart(2,'0');
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
  }
  const obs = new IntersectionObserver((entries) => {
    entries.forEach(e => { if (e.isIntersecting) { e.target.classList.add('in'); obs.unobserve(e.target); } });
  }, { rootMargin: '-40px' });
  document.querySelectorAll('.reveal').forEach(el => obs.observe(el));
})();"""

html = f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>세션 3 · 쇼룸·백화점 교육자료 | 오가렌 매니저 교육</title>
<meta name="description" content="오가렌 매니저 교육 · 세션 3 · 쇼룸·백화점 교육자료 (제품 라인업 + 경쟁사 비교 + 브랜드·안전성)" />
<style>
{CSS}
</style>
</head>
<body>
  <div class="topbar">
    <div class="container">
      <a href="index.html" class="back">← Back to Index</a>
      <div class="title"><em>Session</em> 03 · 쇼룸·백화점 교육자료</div>
      <a href="https://drive.google.com/drive/folders/" target="_blank" rel="noopener" class="download">↗ Drive 원본</a>
    </div>
  </div>

  <section class="session-hero">
    <div class="container reveal">
      <div class="kicker">Session 03 · Product Library · Lineup · Reference</div>
      <h1>쇼룸·백화점<br/><em>교육자료.</em></h1>
    </div>
  </section>

  <div class="container">
    <div class="layout">
      <aside class="sidebar">
        <h4>— Contents</h4>
        <ul>{sidebar}</ul>
      </aside>
      <main class="main">
        {main_html}
      </main>
    </div>
  </div>

  <footer class="footer">
    <div class="container">
      <div class="brand-mark">OGAREN <em>Manager Education</em></div>
      <div>v2.3 · 세션 3 · © DX팀</div>
    </div>
  </footer>
  <script>
{JS}
</script>
</body>
</html>
"""

(ROOT / "session3.html").write_text(html, encoding="utf-8")
print(f"OK session3.html  ({total} slides, {len(toc)} parts)")
