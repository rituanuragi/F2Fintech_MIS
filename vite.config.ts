import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { readFileSync } from 'fs';

const pkg = JSON.parse(readFileSync('./package.json', 'utf-8'));

export default defineConfig({
  plugins: [react(), tailwindcss()],
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  // strictPort: never silently move to another port (that is how old copies kept showing up)
  server: {
    port: Number(process.env.WEB_PORT || 5173),
    strictPort: true,
    open: !!process.env.WEB_PORT, // opens the browser on the right address
    proxy: { '/api': `http://localhost:${process.env.API_PORT || process.env.PORT || 3005}` },
  },
});
