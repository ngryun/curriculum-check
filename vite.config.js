import { execSync } from "node:child_process";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteSingleFile } from "vite-plugin-singlefile";

// 화면 맨 아래 '프로그램 업데이트' 표시용: 빌드한 때(한국 시간)와 커밋 번호
function buildInfo() {
  let commit = "";
  try {
    commit = execSync("git rev-parse --short HEAD", { stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
  } catch (e) {
    // git이 없거나 저장소가 아니면 커밋 번호 없이
  }
  const builtAt = new Date().toLocaleString("sv-SE", { timeZone: "Asia/Seoul" }).slice(0, 16); // 2026-09-28 16:40
  return { builtAt, commit };
}

// 빌드 결과를 인터넷 없이 동작하는 index.html 한 파일로 묶음 (원본과 같은 배포 방식)
export default defineConfig({
  plugins: [react(), viteSingleFile()],
  define: { __BUILD_INFO__: JSON.stringify(buildInfo()) },
  build: { chunkSizeWarningLimit: 4000 },
});
