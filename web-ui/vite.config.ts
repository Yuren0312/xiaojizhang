import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

export default defineConfig({
  plugins: [react(), tailwindcss()],
  // 相对路径：桌面版通过 file:// 加载，绝对路径会 404
  base: './',
  build: {
    // 部分环境（安全软件/回收站不可用）无法清空输出目录，避免构建被阻断
    emptyOutDir: false,
  },
  server: {
    host: '0.0.0.0',
    port: 5173,
    proxy: {
      '/api': {
        target: 'http://localhost:8100',
        changeOrigin: true,
      },
    },
  },
})
