# 졸업이수요건 점검

https://curriculum-check.netlify.app/ 의 소스를 JSX로 되돌린 Vite 프로젝트.

- `original/index.html` — 배포본에서 받아 둔 원본 (참고용)
- `src/App.jsx` — 앱 전체 코드
- `src/main.jsx` — 진입점 (SheetJS 코드페이지 등록, 마운트)

```bash
npm install
npm run dev      # 개발 서버 (http://localhost:5173)
npm run build    # dist/index.html 한 파일로 빌드 (인터넷 없이 동작)
```

## 학생 기초조사 (Apps Script)

- `src/survey/appsScript.js` — 학생 휴대폰 화면, 구글 시트용 서버 코드, 설정 붙여넣기 창
- `npm run test:survey` — 서버 코드를 가짜 구글 환경에서 확인
- `npm run build:apps-script` — 템플릿 시트용 `apps-script/Code.gs`·`appsscript.json` 생성 ([docs/survey-template.md](docs/survey-template.md))
- `src/config.js`의 `SURVEY_TEMPLATE_COPY_URL`에 템플릿 시트 주소를 넣으면 ‘템플릿 사본 만들기’ 설치 방법이 켜집니다
