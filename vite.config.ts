import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, '.', '')
  const apiKey = env.CNTAT_API_KEY || env.REACT_APP_API_KEY || ''

  return {
    plugins: [react()],
    server: {
      host: 'localhost',
      port: 3000,
      strictPort: true,
      proxy: {
        '/cntat-api': {
          target: 'https://cntatapi.azurewebsites.net',
          changeOrigin: true,
          secure: true,
          rewrite: (path: string) => path.replace(/^\/cntat-api/, '/api'),
          headers: apiKey ? { 'x-api-key': apiKey } : undefined,
        },
      },
    },
  }
})
