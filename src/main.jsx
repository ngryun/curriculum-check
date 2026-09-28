import ReactDOM from "react-dom/client";
import * as XLSX from "xlsx";
import * as cptable from "xlsx/dist/cpexcel.full.mjs";
import GraduationChecker from "./App.jsx";

// 원본은 SheetJS full 빌드(코드페이지 포함)를 썼으므로, 옛 .xls(한글 CP949 등)를 읽을 수 있게 코드페이지를 등록
XLSX.set_cptable(cptable);

function mount() {
  try {
    const el = document.getElementById("root");
    ReactDOM.createRoot(el).render(<GraduationChecker />);
  } catch (e) {
    console.error(e);
    document.getElementById("root").innerHTML =
      '<pre style="white-space:pre-wrap;color:#A2452C;padding:24px;font-family:monospace;font-size:13px;">앱을 불러오는 중 오류가 발생했습니다.\n\n' +
      (e && e.stack ? e.stack : String(e)) +
      "</pre>";
  }
}

mount();
