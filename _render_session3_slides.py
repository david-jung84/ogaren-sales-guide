# -*- coding: utf-8 -*-
"""
세션 3 PDF/PPTX → PNG 슬라이드 렌더링.
세션 1/2와 동일한 구조: assets/session3/slide_NNN.png

산출:
  - assets/session3/slide_001.png ... slide_NNN.png
  - _session3_slides.json (TOC + 슬라이드 메타)
"""
from __future__ import annotations
import io
import json
import os
import shutil
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image

ROOT = Path(__file__).parent
DL = ROOT / "assets" / "download" / "session3"
OUT = ROOT / "assets" / "session3"
OUT.mkdir(parents=True, exist_ok=True)
# 기존 slide_*.png 청소
for f in OUT.glob("slide_*.png"):
    f.unlink()

W, H = 1600, 900  # 16:9 PNG 크기

# Part 정의: (id, label, source_pdf, slide_titles)
# slide_titles는 길이가 source PDF 페이지 수와 같아야 함.
PARTS = [
    # ─ 0 표지 (자동 생성된 1슬라이드, 세션 표제)
    # 매트리스 라인업 4종
    {
        "id": "signature",
        "label": "Part 01 · 시그니처 매트리스",
        "src": "01_signature_lineup_251216.pdf",
        "titles": [
            "SIGNATURE · 표지", "라인업 소개", "포지셔닝", "시그니처 1 (포근)",
            "시그니처 1 · 디테일", "시그니처 1 · 구성", "시그니처 2 (균형)",
            "시그니처 2 · 디테일", "시그니처 2 · 구성", "시그니처 3 (탄탄)",
            "시그니처 3 · 디테일", "시그니처 3 · 구성", "3종 비교 1",
            "3종 비교 2", "프레임 호환", "사이즈·가격", "케어", "FAQ",
            "체험 가이드", "CLOSING",
        ],
    },
    {
        "id": "hybrid",
        "label": "Part 02 · 하이브리드 매트리스",
        "src": "02_hybrid_lineup_251216.pdf",
        "titles": [
            "HYBRID · 표지", "라인업 소개", "포지셔닝", "구조 (포켓+메모리폼)",
            "소재", "라인업", "사이즈·가격", "케어", "FAQ", "체험 가이드", "CLOSING",
        ],
    },
    {
        "id": "stable",
        "label": "Part 03 · 스테이블 매트리스",
        "src": "03_stable_lineup_251216.pdf",
        "titles": [
            "STABLE · 표지", "라인업 소개", "포지셔닝", "구조 · 지지력",
            "소재", "라인업", "사이즈·가격", "케어", "FAQ", "체험 가이드", "CLOSING",
        ],
    },
    {
        "id": "balanced",
        "label": "Part 04 · 밸런스드 매트리스",
        "src": "04_balanced_lineup_251216.pdf",
        "titles": [
            "BALANCED · 표지", "라인업 소개", "포지셔닝", "구조 · 균형감",
            "소재", "라인업", "사이즈·가격", "케어", "FAQ", "체험 가이드", "CLOSING",
        ],
    },
    # 경쟁사 비교
    {
        "id": "compete",
        "label": "Part 05 · 경쟁사 비교 (3.3.3)",
        "src": "05_competitor_compare_251212.pdf",
        "titles": [
            "경쟁사 비교 · 표지",
            "12종 시리즈 특장 + 경도",
            "시그니처 vs 시몬스 · 씰리",
            "하이브리드 vs 코웨이 · 씰리 · 베스트슬립",
            "스테이블 vs 씰리 · 코웨이 · 에이스 · 한셀 · 일룸",
            "밸런스드 vs 에이스 · 코웨이 · 일룸",
        ],
    },
    # 브랜드 · 분류
    {
        "id": "nooer",
        "label": "Part 06 · 누어 교육자료",
        "src": "06_nooer_education_240923_lite.pdf",
        "titles": ["NOOER · 포근히, 탄탄히 — 브랜드 1장"],
    },
    {
        "id": "snc",
        "label": "Part 07 · 슬립퍼 vs 누어",
        "src": "07_sleeper_nooer_compare_240220_lite.pdf",
        "titles": ["SLEEPER & NOOER 비교 — 브랜드 포지셔닝"],
    },
    {
        "id": "classify",
        "label": "Part 08 · 슬립퍼 매트리스 분류표",
        "src": "09_mattress_classification_240513_lite.pdf",
        "titles": [
            "분류표 · 표지",
            "포지션 매트릭스 (포근 ~ 탄탄)",
            "라인업 한눈에",
            "소재 분류",
            "사이즈·가격 분류",
            "가격대별",
            "추천 가이드",
        ],
    },
    # 안전성
    {
        "id": "radon",
        "label": "Part 09 · 라돈 안전성",
        "src": "08_radon_safety_240220_lite.pdf",
        "titles": ["라돈 검사 데이터 + 인증 1장"],
    },
    # 베개
    {
        "id": "pillow",
        "label": "Part 10 · 경추베개",
        "src": "10_cervical_pillow_240214.pdf",
        "titles": ["경추베개 라인업 — 매장 응대용 1장"],
    },
]

# 신규 PDF 3종 (2026-05-22 추가)
# - 제품편 v1: p2-26 라인업 + p27-31 타사비교는 세션 3 lineup/compete와 중복 → 제외
#             p32-49 프레임 부분만 채택 (18p)
# - 서비스편 v1: 전체 54p
# - 운영편 v1: 전체 88p
EXT_PARTS = [
    {
        "id": "product_frame",
        "label": "Part 11 · 제품편 — 프레임",
        "src_path": r"C:\Users\USER\Downloads\세일즈매뉴얼_제품part v1_20260410.pdf",
        "page_range": (32, 49),
        "titles_prefix": "제품편",
    },
    {
        "id": "service_manual",
        "label": "Part 12 · 서비스 응대 매뉴얼 v1",
        "src_path": r"C:\Users\USER\Downloads\세일즈매뉴얼_서비스 part v1_20260410.pdf",
        "page_range": (1, None),  # 전체
        "titles_prefix": "서비스편",
    },
    {
        "id": "store_ops_manual",
        "label": "Part 13 · 매장 운영 매뉴얼 v1",
        "src_path": r"C:\Users\USER\Downloads\매장 운영 매뉴얼_part v1_20260410.pdf",
        "page_range": (1, None),  # 전체
        "titles_prefix": "운영편",
    },
]


def render_page_to_png(page: fitz.Page, out_path: Path):
    """16:9 PNG로 렌더링. 종횡비 다르면 검은 패딩."""
    r = page.rect
    aspect = r.width / r.height
    want = W / H
    # zoom to fit
    if aspect > want:
        zoom = W / r.width
    else:
        zoom = H / r.height
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    # 16:9 캔버스 위에 가운데 정렬
    canvas = Image.new("RGB", (W, H), (15, 15, 18))
    x = (W - img.width) // 2
    y = (H - img.height) // 2
    canvas.paste(img, (x, y))
    canvas.save(out_path, "PNG", optimize=True)


# 슬라이드 1: 세션 표지 (자동 생성)
COVER_TXT_LARGE = "Session 03"
COVER_TXT_TITLE = "쇼룸 · 백화점 교육자료"
COVER_SUB = "Product Library · Lineup · Reference"

try:
    from PIL import ImageDraw, ImageFont
    cover = Image.new("RGB", (W, H), (10, 10, 11))
    draw = ImageDraw.Draw(cover)
    # 폰트 fallback
    def _font(size, bold=False):
        for name in ([
            "C:/Windows/Fonts/malgunbd.ttf" if bold else "C:/Windows/Fonts/malgun.ttf",
        ]):
            try:
                return ImageFont.truetype(name, size)
            except Exception:
                pass
        return ImageFont.load_default()
    f_small = _font(22)
    f_title = _font(96, bold=True)
    f_sub = _font(20)
    draw.text((80, 240), COVER_TXT_LARGE.upper(), fill=(201, 184, 148), font=f_small)
    draw.text((80, 290), COVER_TXT_TITLE, fill=(232, 220, 196), font=f_title)
    draw.text((80, 480), COVER_SUB, fill=(118, 115, 109), font=f_sub)
    # 가장자리 선
    draw.line([(80, 590), (W - 80, 590)], fill=(40, 40, 46), width=1)
    draw.text((80, 614), "OGAREN · Manager Education", fill=(118, 115, 109), font=f_sub)
    cover.save(OUT / "slide_001.png", "PNG", optimize=True)
    cover_ok = True
except Exception as e:
    print(f"cover 생성 실패: {e}")
    cover_ok = False

slide_idx = 2 if cover_ok else 1
toc = []
if cover_ok:
    toc.append({
        "id": "cover", "label": "표지",
        "from": 1, "to": 1,
        "slides": [{"num": 1, "title": "SESSION 03 · COVER"}],
    })

for part in PARTS:
    src = DL / part["src"]
    if not src.exists():
        print(f"[SKIP] {src.name} 없음")
        continue
    doc = fitz.open(src)
    n = doc.page_count
    titles = part["titles"]
    if len(titles) < n:
        titles = titles + [f"{titles[0] if titles else part['label']} · {i+1}" for i in range(len(titles), n)]
    elif len(titles) > n:
        titles = titles[:n]

    start = slide_idx
    slide_list = []
    for i in range(n):
        page = doc.load_page(i)
        out_p = OUT / f"slide_{slide_idx:03d}.png"
        render_page_to_png(page, out_p)
        slide_list.append({"num": slide_idx, "title": titles[i]})
        slide_idx += 1
    doc.close()
    toc.append({
        "id": part["id"],
        "label": part["label"],
        "from": start,
        "to": slide_idx - 1,
        "slides": slide_list,
    })
    print(f"  {part['id']}: slides {start} - {slide_idx - 1}  ({n}p)")

# 하디는 단종 → 미렌더링 (2026-05-22)
# 신규 PDF 3종 처리 (제품편 프레임 부분 + 서비스편 + 운영편)
for ext in EXT_PARTS:
    src = Path(ext["src_path"])
    if not src.exists():
        print(f"[SKIP] {src} 없음"); continue
    doc = fitz.open(src)
    n = doc.page_count
    p0, p1 = ext["page_range"]
    if p1 is None:
        p1 = n
    p0 = max(1, p0); p1 = min(n, p1)
    prefix = ext["titles_prefix"]
    start = slide_idx
    slide_list = []
    for src_page in range(p0, p1 + 1):
        page = doc.load_page(src_page - 1)
        out_p = OUT / f"slide_{slide_idx:03d}.png"
        render_page_to_png(page, out_p)
        # 라벨: "운영편 · p23" 식
        title = f"{prefix} · p{src_page}"
        slide_list.append({"num": slide_idx, "title": title, "src_page": src_page})
        slide_idx += 1
    doc.close()
    toc.append({
        "id": ext["id"],
        "label": ext["label"],
        "from": start,
        "to": slide_idx - 1,
        "slides": slide_list,
    })
    print(f"  {ext['id']}: slides {start} - {slide_idx - 1}  ({p1 - p0 + 1}p {prefix} p{p0}-p{p1})")

total = slide_idx - 1
print(f"\nTOTAL: {total} slides")

(ROOT / "_session3_slides.json").write_text(
    json.dumps({"total": total, "toc": toc}, ensure_ascii=False, indent=2),
    encoding="utf-8",
)
print(f"OK toc -> _session3_slides.json")
