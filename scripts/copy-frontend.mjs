// Copies the static frontend into dist/ for Tauri to bundle. The app has no
// build step — this is just a file copy so the bundle excludes repo internals.
import { cpSync, rmSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const dist = join(root, 'dist');

rmSync(dist, { recursive: true, force: true });
mkdirSync(dist);
for (const entry of ['index.html', 'css', 'js']) {
  cpSync(join(root, entry), join(dist, entry), {
    recursive: true,
    filter: (src) => !/\.DS_Store$/.test(src),
  });
}
console.log('frontend copied to dist/');
