import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    environment: 'jsdom',
    setupFiles: ['./tests/setup.ts'],
    coverage: {
      provider: 'v8',
      reporter: ['text'],
      include: [
        'src/domain/invoice/**/*.ts',
        'src/application/services/**/*.ts',
        'src/infrastructure/browser/**/*.ts',
        'src/infrastructure/excel/**/*.ts',
      ],
      exclude: [
        'tests/**',
        'dist/**',
        'dist-ssr/**',
        'coverage/**',
        '*.config.*',
        'vite.config.ts',
        'vitest.config.ts',
        'eslint.config.js',
        'src/main.tsx',
        'src/**/*.css',
        'src/assets/**',
      ],
      thresholds: {
        lines: 90,
        branches: 90,
        functions: 90,
        statements: 90,
      },
    },
  },
})
