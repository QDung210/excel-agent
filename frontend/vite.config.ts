import path from 'node:path'
import tailwindcss from '@tailwindcss/vite'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: { '@': path.resolve(import.meta.dirname, './src') },
  },
  server: {
    port: 5173,
    proxy: {
      '/api': { target: 'http://127.0.0.1:8000', changeOrigin: true },
    },
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    // FortuneSheet ships its formula engine as one ~2.7 MB chunk; it is lazy-loaded with the workspace route.
    chunkSizeWarningLimit: 2800,
    rolldownOptions: {
      output: {
        codeSplitting: {
          groups: [
            { name: 'fortune-sheet', test: /node_modules[\/]@fortune-sheet/ },
            { name: 'assistant-ui', test: /node_modules[\/](@assistant-ui|assistant-stream|react-markdown|remark-|micromark|mdast-|hast-|unified)/ },
          ],
        },
      },
    },
  },
})
