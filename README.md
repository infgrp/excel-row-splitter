# 과학전람회 엑셀 도구

엑셀 파일을 서버에 전송하지 않고 브라우저 안에서 처리하는 웹앱이다.

## 기능

### 1. Alt+Enter 행 분리

- 기본 분리 열: D, E
- 사용자가 분리 기준 열을 추가하거나 삭제할 수 있음
- 모든 시트 또는 선택 시트에 적용
- 분리하지 않는 다른 셀의 값은 그대로 복제

### 2. 상장용 명단 생성

포상 대상자 명단에서 처리할 시트를 선택하면 다음 규칙으로 메일머지용
엑셀 파일을 만든다.

- A열 호수 → 출력 A열
- E열 공적내용의 마지막 괄호 → 출력 B열
- J열 상종 및 등급 → 출력 C열
- F열 학생별 소속 → 반복되는 각 행의 D열
- H열 학생별 학년 → `s1`, `s2`, `s3` …
- I열 학생별 이름 → `n1`, `n2`, `n3` …
- I열 학생 수만큼 결과 행 반복
- F열 소속이 하나이면 모든 학생에게 반복
- F열 소속이 여러 개이면 I열 학생 순서대로 대응

## 로컬 실행

```bash
npm install
npm run dev
```

브라우저에서 `http://localhost:3000`을 연다.

## GitHub Pages 배포

1. 이 폴더의 모든 파일을 `infgrp/excel-row-splitter` 저장소의 `main`
   브랜치에 커밋하고 푸시한다.
2. 저장소의 `Settings → Pages`로 이동한다.
3. `Build and deployment → Source`를 `GitHub Actions`로 선택한다.
4. `Actions` 탭에서 `Deploy to GitHub Pages` 작업이 성공할 때까지 기다린다.
5. 배포 주소는 다음과 같다.

   `https://infgrp.github.io/excel-row-splitter/`

저장소 이름을 바꾸면 `next.config.ts`의 `repositoryName`도 같은 이름으로
수정해야 한다.

## 권장 명령

```bash
npm ci
npm run check
```

`npm run check`는 변환 로직 테스트, 코드 검사, 정적 사이트 빌드를 차례로
수행한다.
