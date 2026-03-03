# TitleMaker

강의자료/필기족보 파일명을 규칙에 맞춰 자동 생성하고 출력 폴더에 저장하는 데스크톱 앱입니다.

강의자료 모드 형식:

`[수업명][순서]-[수업일]-[교시]-[교수님성함]-[강의제목].pdf`

필족 모드 형식:

`[수업명][순서]-[수업일]-[교시]-[교수님성함]-[강의제목]-[작성자명].docx`

## 핵심 기능

- 상단 버튼으로 `강의자료 모드` / `필족 모드` 전환
- 교시가 `시작 교시 + 시수`로 자동 확장
  - 예: 시작 교시 `5`, 시수 `3` -> `5,6,7교시`
- 현재 학기 시간표(`assets/default-timetable.xlsx`)를 앱에 내장
- 엑셀 의존 완화: `관리자 창`에서 강의 계획 직접 편집/저장
  - 저장 파일: 사용자 폴더의 `lecture-plan.custom.json` (`app.getPath('userData')`)
- 필족 모드에서 `.docx` 입력 시 제목 규칙으로 저장

## 입력 흐름 (공통)

1. 파일 드래그앤드롭
2. 과목명 선택
3. 순서 선택
4. 캘린더에서 수업일 선택
5. 시작 교시 선택
6. 시수 선택
7. 교수님 성함 입력
8. 강의 제목 선택

## 필족 모드 추가 입력

1. `필족 모드`로 전환
2. `필족 작성자 이름` 입력
3. `.docx` 파일 입력
4. 저장 실행

## 실행

```bash
npm install
npm start
```

## 경량 웹 버전 (설치 없음)

Electron 앱을 유지한 채, `lite-web/`에 순수 `html/css/js` 버전을 추가했습니다.

- 파일명 생성 + 복사/CSV 내보내기 전용
- 실제 파일 이름 변경/저장은 하지 않음
- `lite-web/index.html` 더블클릭으로 바로 실행 가능
- 파일 업로드 없이 행을 직접 추가해 제목만 생성 가능

## 빌드

```bash
npm run dist:mac
npm run dist:win
```

빌드가 끝나면 `dist/` 폴더에 실행 파일이 생성됩니다.

- macOS: `TitleMaker-*.dmg` 또는 `TitleMaker-*-mac.zip`
- Windows(x64): `TitleMaker-*-win.zip` (압축 해제 후 `TitleMaker.exe` 실행)

Windows ARM 기기가 필요하면:

```bash
npm run dist:win:arm64
```

최종 사용자(동기/조교)는 `npm` 설치 없이 앱 파일만 받아서 실행하면 됩니다.

## 테스트

```bash
npm test
```
