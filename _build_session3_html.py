# -*- coding: utf-8 -*-
"""manifest → session3.html 빌드"""
import json
from pathlib import Path

ROOT = Path(__file__).parent
manifest = json.loads((ROOT / "_session3_manifest.json").read_text(encoding="utf-8"))

# 카테고리 그룹
GROUPS = [
    ("lineup",  "Part 01",  "신규 라인업 3.3.3",       "2025-12 출시 4종 라인업 — 시그니처 · 하이브리드 · 스테이블 · 밸런스드"),
    ("compare", "Part 02",  "경쟁사 비교 자료",        "라인업·스펙·가격 — 타사 대비 포지셔닝 응대 자료"),
    ("brand",   "Part 03",  "브랜드 · 분류 자료",      "누어·슬립퍼 두 브랜드의 라인업·소재·포지셔닝 차이"),
    ("safety",  "Part 04",  "안전성",                 "라돈·인증 — 매장 응대에서 가장 자주 받는 질문"),
    ("pillow",  "Part 05",  "베개",                   "경추베개 응대 자료"),
    ("bed",     "Part 06",  "침대 프레임",            "하디 침대 — 구조·소재·구성"),
]


def card_html(item: dict) -> str:
    thumb = item.get("thumb")
    if thumb:
        thumb_html = f'<img class="thumb" src="{thumb}" alt="{item["title"]} 미리보기" loading="lazy" />'
    else:
        ext = item["ext"].upper()
        thumb_html = (
            f'<div class="thumb thumb-fallback"><span class="ext">{ext}</span></div>'
        )

    pages = item.get("pages")
    pages_html = f'<span class="meta-i">{pages}p</span>' if pages else ""

    file_path = item["file"]
    ext_label = item["ext"].upper()

    # PPTX는 새창 미리보기 안되므로 다운로드만
    if item["ext"] == "pdf":
        preview_btn = (
            f'<a class="btn" href="{file_path}" target="_blank" rel="noopener">'
            f'<span>Preview</span><span class="arr">↗</span></a>'
        )
    else:
        preview_btn = ""

    download_btn = (
        f'<a class="btn btn-primary" href="{file_path}" download>'
        f'<span>Download</span><span class="arr">↓</span></a>'
    )

    return f"""
        <article class="doc-card reveal" data-cat="{item['category']}">
          <div class="thumb-wrap">
            {thumb_html}
            <span class="ext-badge">{ext_label}</span>
          </div>
          <div class="body">
            <h3>{item['title']}</h3>
            <p class="desc">{item['desc']}</p>
            <div class="meta-row">
              <span class="meta-i">{item['size']}</span>
              {pages_html}
            </div>
            <div class="actions">
              {preview_btn}
              {download_btn}
            </div>
          </div>
        </article>
"""


# 그룹별 섹션
sections = []
for cat, pnum, title, sub in GROUPS:
    items = [m for m in manifest if m["category"] == cat]
    if not items:
        continue
    cards = "\n".join(card_html(it) for it in items)
    sections.append(
        f"""
      <section class="part-section" id="{cat}">
        <div class="part-head">
          <span class="pnum">{pnum}</span>
          <h2>{title}</h2>
          <span class="range">{len(items)} {'docs' if len(items)>1 else 'doc'}</span>
        </div>
        <p class="part-sub">{sub}</p>
        <div class="doc-grid">{cards}
        </div>
      </section>
"""
    )

# 사이드바
sidebar_items = []
for cat, pnum, title, sub in GROUPS:
    items = [m for m in manifest if m["category"] == cat]
    if not items:
        continue
    sidebar_items.append(
        f'<li><a href="#{cat}"><em class="pnum">{pnum}</em>{title}</a></li>'
    )
sidebar = "\n        ".join(sidebar_items)

total = len(manifest)
sections_joined = "\n".join(sections)

html = f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>세션 3 · 쇼룸·백화점 교육자료 | 오가렌 매니저 교육</title>
<meta name="description" content="오가렌 쇼룸·백화점 교육자료 — 신규 라인업 3.3.3, 경쟁사 비교, 브랜드·안전성 응대 자료" />
<style>
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700;800;900&family=Cormorant+Garamond:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap');

* {{ box-sizing: border-box; margin: 0; padding: 0; }}
:root {{
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
}}
html, body {{ font-family: var(--sans); color: var(--text); background: var(--bg); line-height: 1.6; -webkit-font-smoothing: antialiased; font-feature-settings: 'ss01','ss02'; }}
body {{ background:
  radial-gradient(1200px 600px at 80% -10%, rgba(232,220,196,0.06), transparent 60%),
  radial-gradient(900px 500px at 0% 10%, rgba(201,184,148,0.04), transparent 60%),
  var(--bg);
  min-height: 100vh;
}}
a {{ color: inherit; text-decoration: none; }}
img {{ display: block; }}
.container {{ max-width: 1320px; margin: 0 auto; padding: 0 32px; }}
@media (max-width: 640px) {{ .container {{ padding: 0 20px; }} }}

/* ===== Topbar (carry from existing sessions) ===== */
.topbar {{ background: rgba(10,10,11,0.85); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px); border-bottom: 1px solid var(--border); padding: 18px 0; position: sticky; top: 0; z-index: 20; }}
.topbar .container {{ display: flex; align-items: center; justify-content: space-between; gap: 16px; }}
.topbar .back {{ font-size: 12px; letter-spacing: 1.5px; text-transform: uppercase; color: var(--text-2); display: inline-flex; align-items: center; gap: 10px; transition: color .25s ease; }}
.topbar .back:hover {{ color: var(--accent); }}
.topbar .title {{ font-family: var(--serif); font-size: 20px; font-weight: 500; letter-spacing: -0.3px; }}
.topbar .title em {{ font-style: normal; color: var(--accent); }}
.topbar .download {{ font-size: 11px; letter-spacing: 2px; text-transform: uppercase; padding: 10px 18px; border: 1px solid var(--border-strong); border-radius: 2px; transition: border-color .25s ease, color .25s ease, background .25s ease; color: var(--text); }}
.topbar .download:hover {{ border-color: var(--accent-warm); color: var(--accent); background: var(--surface); }}

.session-hero {{ padding: 80px 0 40px; }}
.session-hero .kicker {{ font-size: 11px; letter-spacing: 3px; text-transform: uppercase; color: var(--accent-warm); margin-bottom: 24px; }}
.session-hero h1 {{ font-family: var(--serif); font-size: clamp(40px, 6vw, 72px); font-weight: 500; line-height: 1.05; letter-spacing: -1px; }}
.session-hero h1 em {{ font-style: normal; color: var(--accent); font-weight: 400; }}
.session-hero .lead {{ font-size: 16px; color: var(--text-2); max-width: 640px; margin-top: 28px; line-height: 1.7; font-weight: 300; }}
.session-hero .meta {{ margin-top: 28px; display: flex; gap: 36px; flex-wrap: wrap; font-size: 12px; letter-spacing: 1.5px; text-transform: uppercase; color: var(--text-3); }}
.session-hero .meta b {{ color: var(--text); font-weight: 500; }}

.layout {{ display: grid; grid-template-columns: 240px 1fr; gap: 56px; padding: 32px 0 80px; }}
@media (max-width: 980px) {{ .layout {{ grid-template-columns: 1fr; gap: 24px; }} }}

.sidebar {{ position: sticky; top: 90px; align-self: start; max-height: calc(100vh - 110px); overflow-y: auto; padding-right: 12px; }}
.sidebar::-webkit-scrollbar {{ width: 4px; }}
.sidebar::-webkit-scrollbar-thumb {{ background: var(--border-strong); border-radius: 2px; }}
.sidebar h4 {{ font-size: 10px; letter-spacing: 3px; text-transform: uppercase; color: var(--text-3); margin-bottom: 18px; }}
.sidebar ul {{ list-style: none; }}
.sidebar li a {{ display: block; padding: 12px 0; font-size: 13px; color: var(--text-2); border-bottom: 1px solid var(--border); transition: color .25s ease, padding-left .25s ease; line-height: 1.4; }}
.sidebar li a:hover {{ color: var(--accent); padding-left: 8px; }}
.sidebar li a .pnum {{ font-family: var(--serif); font-style: normal; color: var(--accent-warm); margin-right: 10px; font-size: 12px; letter-spacing: 1px; }}

.main {{ min-width: 0; }}
.part-section {{ margin-bottom: 72px; scroll-margin-top: 100px; }}
.part-section .part-head {{ margin-bottom: 14px; display: flex; align-items: baseline; gap: 24px; flex-wrap: wrap; padding-bottom: 20px; border-bottom: 1px solid var(--border); }}
.part-section .part-head .pnum {{ font-family: var(--serif); font-style: normal; font-size: 18px; color: var(--accent-warm); letter-spacing: 1px; }}
.part-section .part-head h2 {{ font-family: var(--serif); font-size: 36px; font-weight: 500; letter-spacing: -0.5px; line-height: 1.1; flex: 1; min-width: 0; }}
.part-section .part-head .range {{ font-size: 11px; letter-spacing: 2px; text-transform: uppercase; color: var(--text-3); }}
.part-section .part-sub {{ font-size: 14px; color: var(--text-3); margin-bottom: 28px; font-weight: 300; }}

/* ===== Doc card grid ===== */
.doc-grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 24px; }}
@media (max-width: 760px) {{ .doc-grid {{ grid-template-columns: 1fr; }} }}

.doc-card {{
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 4px;
  overflow: hidden;
  transition: border-color .35s ease, transform .35s ease;
  display: flex; flex-direction: column;
}}
.doc-card:hover {{ border-color: var(--accent-warm); transform: translateY(-3px); }}

.doc-card .thumb-wrap {{ position: relative; aspect-ratio: 16/9; background: #1a1a1f; overflow: hidden; border-bottom: 1px solid var(--border); }}
.doc-card .thumb {{ width: 100%; height: 100%; object-fit: cover; transition: transform .6s cubic-bezier(.2,.6,.2,1); }}
.doc-card:hover .thumb {{ transform: scale(1.04); }}
.doc-card .thumb-fallback {{ width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #1c1c22, #15151a); }}
.doc-card .thumb-fallback .ext {{ font-family: var(--serif); font-size: 56px; color: var(--accent-warm); letter-spacing: 4px; }}
.doc-card .ext-badge {{ position: absolute; top: 14px; left: 14px; font-family: var(--serif); font-style: normal; font-size: 12px; color: var(--accent); background: rgba(10,10,11,0.7); backdrop-filter: blur(8px); padding: 4px 10px; border-radius: 2px; letter-spacing: 2px; }}

.doc-card .body {{ padding: 24px 26px 26px; display: flex; flex-direction: column; flex: 1; }}
.doc-card h3 {{ font-family: var(--serif); font-size: 24px; font-weight: 500; line-height: 1.2; letter-spacing: -0.3px; margin-bottom: 10px; }}
.doc-card .desc {{ font-size: 13px; color: var(--text-3); line-height: 1.55; font-weight: 300; min-height: 40px; }}
.doc-card .meta-row {{ display: flex; gap: 14px; margin: 18px 0 20px; padding-top: 16px; border-top: 1px solid var(--border); flex-wrap: wrap; }}
.doc-card .meta-row .meta-i {{ font-size: 11px; letter-spacing: 1.5px; text-transform: uppercase; color: var(--text-3); }}
.doc-card .meta-row .meta-i + .meta-i {{ padding-left: 14px; border-left: 1px solid var(--border); }}
.doc-card .actions {{ display: flex; gap: 10px; margin-top: auto; }}

.btn {{
  flex: 1;
  display: inline-flex; align-items: center; justify-content: center; gap: 8px;
  font-size: 12px; letter-spacing: 1.5px; text-transform: uppercase;
  padding: 13px 16px; border: 1px solid var(--border-strong); border-radius: 2px;
  color: var(--text-2); background: transparent;
  transition: border-color .25s ease, color .25s ease, background .25s ease;
}}
.btn:hover {{ border-color: var(--accent-warm); color: var(--accent); background: var(--surface-2); }}
.btn .arr {{ color: var(--accent); transition: transform .25s ease; }}
.btn:hover .arr {{ transform: translateX(2px); }}
.btn-primary {{ background: var(--accent); color: #0A0A0B; border-color: var(--accent); }}
.btn-primary .arr {{ color: #0A0A0B; }}
.btn-primary:hover {{ background: var(--accent-warm); color: #0A0A0B; border-color: var(--accent-warm); }}

/* ===== Footer ===== */
.footer {{ padding: 40px 0; border-top: 1px solid var(--border); margin-top: 20px; }}
.footer .container {{ display: flex; justify-content: space-between; align-items: center; gap: 16px; flex-wrap: wrap; font-size: 12px; letter-spacing: 1px; color: var(--text-3); }}
.footer .brand-mark {{ font-family: var(--serif); font-size: 14px; letter-spacing: 1px; color: var(--text-2); }}
.footer em {{ font-style: normal; }}

.reveal {{ opacity: 0; transform: translateY(20px); transition: opacity .8s ease, transform .8s ease; }}
.reveal.in {{ opacity: 1; transform: translateY(0); }}
</style>
</head>
<body>
  <header class="topbar">
    <div class="container">
      <a class="back" href="index.html">← All Sessions</a>
      <div class="title">Session <em>03</em> · 쇼룸·백화점 교육자료</div>
      <a class="download" href="https://drive.google.com/drive/folders/" target="_blank" rel="noopener">↗ Drive 원본</a>
    </div>
  </header>

  <section class="session-hero">
    <div class="container reveal">
      <div class="kicker">Session 03 · Product Library</div>
      <h1>쇼룸·백화점<br/><em>교육자료.</em></h1>
      <p class="lead">신규 라인업 3.3.3과 경쟁사 비교, 브랜드·안전성 응대 자료까지 — 쇼룸과 백화점 매장에서 그대로 꺼내 쓸 수 있는 PDF 라이브러리입니다.</p>
      <div class="meta">
        <div><b>{total}</b> Docs</div>
        <div><b>{len([m for m in manifest if m['category']=='lineup'])}</b> 신규 라인업</div>
        <div><b>2025-12</b> 기준</div>
      </div>
    </div>
  </section>

  <div class="container">
    <div class="layout">
      <aside class="sidebar reveal">
        <h4>Index</h4>
        <ul>
        {sidebar}
        </ul>
      </aside>
      <main class="main">
        {sections_joined}
      </main>
    </div>
  </div>

  <footer class="footer">
    <div class="container">
      <div class="brand-mark">OGAREN <em>Manager Education</em> · Session 03</div>
      <div>v2.2 · last updated 2026-05-22 · © DX팀</div>
    </div>
  </footer>

  <script>
(function(){{
  const obs = new IntersectionObserver((entries) => {{
    entries.forEach(e => {{ if (e.isIntersecting) {{ e.target.classList.add('in'); obs.unobserve(e.target); }} }});
  }}, {{ rootMargin: '-40px' }});
  document.querySelectorAll('.reveal').forEach(el => obs.observe(el));
}})();
</script>
</body>
</html>
"""

(ROOT / "session3.html").write_text(html, encoding="utf-8")
print("OK session3.html written")
