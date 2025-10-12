import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/api': {
        target: 'https://c67feb255965.ngrok-free.app',
        changeOrigin: true,
        secure: false,
      },
      '/wopi': {
        target: 'https://c67feb255965.ngrok-free.app',
        changeOrigin: true,
        secure: false,
      },
      '/files': {
        target: 'https://c67feb255965.ngrok-free.app',
        changeOrigin: true,
        secure: false,
      }
    }
  }
})