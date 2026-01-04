import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
    plugins: [react()],
    server: {
        proxy: {
            '/upload': 'http://127.0.0.1:5000',
            '/download': 'http://127.0.0.1:5000'
        }
    }
})
