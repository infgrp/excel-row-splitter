# 과학전람회 엑셀 도구

엑셀 파일을 외부 서버로 전송하지 않고 브라우저 안에서 처리하는 웹앱이다.

## 기능

1. **Alt+Enter 행 분리**
   - 기본 대상 열 D, E
   - 대상 열 추가·삭제
   - 전체 시트 또는 선택 시트 처리

2. **상장 명단 만들기**
   - 포상 대상자 명단을 메일머지용 상장 명단으로 변환
   - 학생별 이름·학년·소속 대응
   - 여러 학생과 여러 소속의 Alt+Enter 자료 처리

3. **학교 주소 추출**
   - 상장 명단의 `sc1` 열에서 수상학교 추출
   - 학교명·우편번호·주소·수신자 이름·주의사항 엑셀 생성
   - 중복 학교 자동 제거
   - 기본 학교현황 796개교 포함

4. **주소 라벨 인쇄**
   - 아이라벨 CL425 A4 10칸(2열×5행) 규격
   - 주의사항·우편번호·주소·수신자 자동 배치
   - 긴 주소의 단어 단위 줄바꿈
   - 10개 초과 시 다음 A4 페이지 자동 생성

## GitHub Pages 배포

### 1. 저장소 만들기

GitHub에서 새 저장소를 만든다. 저장소 이름은 자유롭게 정할 수 있다.

### 2. 파일 올리기

이 ZIP 파일의 압축을 푼 뒤 `excel-row-splitter` 폴더 **안의 모든 파일과
폴더**를 저장소의 `main` 브랜치에 올린다. `.github` 폴더도 반드시
포함해야 한다.

### 3. Pages 설정

저장소에서 다음 순서로 설정한다.

1. `Settings`
2. `Pages`
3. `Build and deployment`
4. `Source`를 **GitHub Actions**로 선택

`Actions` 탭의 `Deploy to GitHub Pages` 작업이 끝나면 Pages 주소가
생성된다. 저장소 이름은 빌드할 때 자동으로 반영되므로 코드를 수정할
필요가 없다.

## 로컬 실행

Node.js 22 이상이 필요하다.

```bash
npm ci
npm run dev
```

브라우저에서 `http://localhost:3000`을 연다.

## 학교 기본현황 업데이트

- 원본: `data/school-master.xlsx`
- 웹앱용 데이터: `public/data/schools.json`

새 학교 기본현황 엑셀을 `data/school-master.xlsx`로 교체한 뒤 실행한다.

```bash
npm run update:schools
```

변경된 두 데이터 파일을 GitHub에 함께 올리면 된다.

## 검증

```bash
npm run check
```

변환 로직 테스트, 코드 검사, GitHub Pages용 정적 빌드를 차례로 수행한다.
