// 템플릿 시트에 넣을 Apps Script 코드를 파일로 만듭니다 → apps-script/Code.gs
// 실행: npm run build:apps-script  (학생 화면이나 서버 코드를 고친 뒤에는 다시 만들어 템플릿 시트에 붙여넣어야 합니다)
import fs from "node:fs";
import { buildTemplateAppsScriptCode } from "../src/survey/appsScript.js";

fs.mkdirSync(new URL("../apps-script/", import.meta.url), { recursive: true });
const out = new URL("../apps-script/Code.gs", import.meta.url);
fs.writeFileSync(out, buildTemplateAppsScriptCode() + "\n");
console.log("만들었습니다:", out.pathname);
