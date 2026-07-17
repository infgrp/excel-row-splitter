# 엑셀 행 분리기 (excel-row-splitter)

엑셀 셀 안의 줄바꿈으로 나열된 자료를 **한 항목당 한 행**으로 분리해 주는 브라우저 도구.

Next.js 16 App Router + React 19 + TypeScript + Tailwind CSS 4. 파일 처리는 모두 브라우저에서 이루어져 서버로 업로드되지 않는다.

## 로컬 실행

```bash
npm install
npm run dev
```

`http://localhost:3000` 접속.

## Vercel 배포

1. https://vercel.com/new 에서 **infgrp/excel-row-splitter** 리포지토리 import
2. Framework Preset: **Next.js** 자동 감지
3. Build/Output 설정은 그대로 두고 **Deploy**

Private 리포지토리라면 Vercel-GitHub 통합 권한이 이 리포에도 허용되어 있는지 확인.

## 스크립트

- `npm run dev` — 개발 서버
- `npm run build` — 프로덕션 빌드
- `npm run start` — 빌드된 앱 실행
- `npm run lint` — ESLint

## 구조

- `app/page.tsx` — 메인 UI
- `app/excel-transform.ts` — 엑셀 분리 로직 (ExcelJS)
- `app/layout.tsx`, `app/globals.css` — 레이아웃·전역 스타일
- `tests/excel-transform.test.mts` — `node --test tests/excel-transform.test.mts`
