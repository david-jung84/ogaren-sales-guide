# 오가렌 매장 매니저 교육 · 세일즈 가이드

오프라인 쇼룸 매니저 교육과 세일즈 가이드의 정적 웹 버전.
원본 PPT를 슬라이드 이미지로 변환하여 누구나 링크로 열어볼 수 있게 한다.

**공개 URL**: https://david-jung84.github.io/ogaren-sales-guide/ (배포 후)

## 구조

```
세션 1 — 우리는 그래서 잘 잔다 (제품 · 인증 · 안전성)
  Part 1  오가렌과 3대 브랜드
  Part 2  매트리스 마스터하기
  Part 3  침대 프레임 마스터하기
  Part 4  인증 & 안전성

세션 2 — 그래서 잘 판다 (응대 · 세일즈 · 운영)
  Part 1  고객 응대 프로세스
  Part 2  세일즈 스크립트 & 거절 처리
  Part 3  TOP 매니저 케이스 (실명)
  Part 4  운영 · KPI · 부록
  Part 5  배송 실무 가이드
```

총 76 슬라이드 (세션1 38 + 세션2 38).

## 파일 구성

| 경로 | 역할 |
|------|------|
| `index.html`, `session1.html`, `session2.html` | 정적 사이트 (배포 대상) |
| `assets/session1/*.png`, `assets/session2/*.png` | 슬라이드 이미지 |
| `assets/download/*.pptx` | PPT 원본 다운로드 |
| `build_deck_v2.js` | PPT 원본 빌더 (v2.1 통합본) |
| `part6_revised.js` | 매월 갱신되는 Part 6 케이스 (TOP 매니저, 매장 벤치마크, 월간 업데이트) |
| `split_and_build.js` | 통합본을 세션 1·2로 분리 + 빌드 |
| `build_site.py` | PPT → PNG + HTML 생성 (정적 사이트 빌드) |
| `images/` | PPT 빌더가 사용하는 제품 사진 원본 |

## 월간 업데이트 워크플로우

본사 데일리 SV 세일즈톡이 발행되면:

```powershell
# 1. part6_revised.js 5분 수정 (TOP 케이스/월간 업데이트 슬라이드)
# 2. 두 PPT 자동 갱신
node split_and_build.js

# 3. 정적 사이트 재빌드 (PNG export + HTML)
python build_site.py

# 4. 커밋 + 푸시 → GitHub Pages 자동 갱신 (수 분 내 반영)
git add .
git commit -m "monthly update: <yyyy-mm>"
git push
```

## 로컬 미리보기

```powershell
python -m http.server 8765
# http://localhost:8765 에서 확인
```

## 요구 환경

- Python 3.10+ (`pip install pywin32 python-pptx`)
- Node.js 18+ (PPT 빌더용, `npm i pptxgenjs`)
- Windows + Microsoft PowerPoint (PPT → PNG 변환에 COM 사용)
