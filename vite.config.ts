/// <reference types="vitest/config" />
import { readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, type Plugin } from 'vite'

/**
 * public/sw.js는 Vite 번들링을 거치지 않고 그대로 복사되므로(define이 안 먹음),
 * 빌드마다 CACHE_VERSION이 자동으로 바뀌게 하려면 빌드 산출물을 직접 치환해야
 * 한다 — 수동으로 버전을 올리는 건 깜빡하기 쉽다(구현 계획의 "모바일 PWA화" 참고).
 */
function injectSwCacheVersion(): Plugin {
  return {
    name: 'inject-sw-cache-version',
    writeBundle() {
      const swPath = resolve(import.meta.dirname, 'dist/sw.js')
      const version = Date.now().toString(36)
      const content = readFileSync(swPath, 'utf-8').replace('__SW_CACHE_VERSION__', version)
      writeFileSync(swPath, content)
    },
  }
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), injectSwCacheVersion()],
  test: {
    environment: 'jsdom',
  },
})
