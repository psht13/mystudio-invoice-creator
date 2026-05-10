import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  build: {
    // ExcelJS is isolated behind a lazy import, so this async workbook chunk is expected.
    chunkSizeWarningLimit: 1000,
  },
  plugins: [react()],
})
