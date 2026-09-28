import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteSingleFile } from "vite-plugin-singlefile";

// 빌드 결과를 인터넷 없이 동작하는 index.html 한 파일로 묶음 (원본과 같은 배포 방식)
export default defineConfig({
  plugins: [react(), viteSingleFile()],
  build: { chunkSizeWarningLimit: 4000 },
});
