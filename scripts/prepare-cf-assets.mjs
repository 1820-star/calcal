import { cp, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const targetDir = path.join(rootDir, 'cf-assets');

const filesToSync = [
  'index.html',
  'styles.css',
  'app.js',
  'manifest.webmanifest',
  'sw.js',
  'favicon.ico',
];

async function main() {
  await mkdir(targetDir, { recursive: true });

  await Promise.all(
    filesToSync.map(async (fileName) => {
      const sourcePath = path.join(rootDir, fileName);
      const targetPath = path.join(targetDir, fileName);
      await cp(sourcePath, targetPath, { force: true });
    })
  );

  process.stdout.write(`Cloudflare assets synced to ${targetDir}\n`);
}

main().catch((error) => {
  process.stderr.write(`Failed to prepare Cloudflare assets: ${error.message}\n`);
  process.exit(1);
});
