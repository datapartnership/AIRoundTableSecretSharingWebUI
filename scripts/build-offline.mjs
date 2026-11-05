// Builds the offline masking tool and places it where the hosted UI serves it:
// public/offline-masker.html for the dev server, and dist/ after a production build.
import { copyFileSync, existsSync, rmSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { build } from 'vite'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const NAME = 'offline-masker.html'

await build({ root, configFile: resolve(root, 'vite.offline.config.js'), logLevel: 'warn' })

const built = resolve(root, '.offline-build/offline.html')
copyFileSync(built, resolve(root, 'public', NAME))
if (existsSync(resolve(root, 'dist'))) copyFileSync(built, resolve(root, 'dist', NAME))
rmSync(resolve(root, '.offline-build'), { recursive: true, force: true })
console.log(`offline tool built: public/${NAME}`)
