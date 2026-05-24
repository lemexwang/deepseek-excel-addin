import fs from 'node:fs'
import os from 'node:os'
import { fileURLToPath, URL } from 'node:url'

import tailwindcss from '@tailwindcss/vite'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'
import json5Plugin from 'vite-plugin-json5'

const certDir = `${os.homedir()}/.office-addin-dev-certs`
const httpsConfig = fs.existsSync(`${certDir}/localhost.crt`)
  ? { key: fs.readFileSync(`${certDir}/localhost.key`), cert: fs.readFileSync(`${certDir}/localhost.crt`) }
  : true

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [tailwindcss(), vue(), json5Plugin()],
  server: {
    https: httpsConfig,
    port: 3000,
  },
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
      async_hook: fileURLToPath(new URL('./async_hook.js', import.meta.url)),
      'node:async_hooks': fileURLToPath(new URL('./async_hook.js', import.meta.url)),
    },
  },
})
