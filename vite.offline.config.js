import { readFileSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { viteSingleFile } from 'vite-plugin-singlefile'

// Builds offline.html into one self-contained HTML file (scripts, styles and images inlined) that
// partners open from disk on a computer without internet. Run through scripts/build-offline.mjs.
const root = dirname(fileURLToPath(import.meta.url))

function readVersion() {
  try {
    return JSON.parse(readFileSync(resolve(root, 'src/version.json'), 'utf8')).version || '1.0.0'
  } catch {
    return '1.0.0'
  }
}

// Blocks every network request from the built file, so it can't leak data even if opened online
const CSP = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: blob:; font-src data:; connect-src 'none'; form-action 'none'; base-uri 'none'"

const offlineCsp = () => ({
  name: 'offline-csp',
  apply: 'build',
  transformIndexHtml: (html) => html.replace('<meta charset="UTF-8" />', `<meta charset="UTF-8" />\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`),
})

export default defineConfig({
  plugins: [react(), viteSingleFile({ removeViteModuleLoader: true }), offlineCsp()],
  define: {
    __APP_VERSION__: JSON.stringify(readVersion()),
  },
  publicDir: false,
  build: {
    outDir: '.offline-build',
    emptyOutDir: true,
    rollupOptions: { input: resolve(root, 'offline.html') },
  },
})
