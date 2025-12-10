import { serve } from '@hono/node-server';
import { readFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import { parse } from 'jsonc-parser';

// Load Wrangler vars as process.env for Node.js runtime
const __dirname = dirname(fileURLToPath(import.meta.url));
const wranglerPath = join(__dirname, '../wrangler.jsonc');

try {
  const wranglerContent = readFileSync(wranglerPath, 'utf-8');
  const config = parse(wranglerContent);

  if (config.vars) {
    for (const [key, value] of Object.entries(config.vars)) {
      if (process.env[key] === undefined) {
        process.env[key] = value;
      }
    }
    console.log('Loaded Wrangler vars into process.env');
  }
} catch (error) {
  console.warn('Failed to load wrangler.jsonc vars:', error.message);
}

const app = await import('./hono.js');
const port = process.env.PORT || 3000;

serve({
  fetch: app.default.fetch,
  port
});

console.log(`Server running at http://localhost:${port}`);
