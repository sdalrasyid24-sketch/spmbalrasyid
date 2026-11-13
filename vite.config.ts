import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import { transform as lightningTransform } from 'lightningcss';
import path from 'path';
import { fileURLToPath } from 'url';
import { defineConfig, type Plugin } from 'vite';
import { VitePWA } from 'vite-plugin-pwa';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const LEGACY_TARGETS = {
  chrome: 88 << 16,
  edge: 88 << 16,
  firefox: 88 << 16,
  safari: 14 << 16,
};

function unwrapCssLayers(css: string): string {
  let out = '';
  let i = 0;
  while (i < css.length) {
    if (css.startsWith('@layer ', i)) {
      const braceIdx = css.indexOf('{', i);
      const semiIdx = css.indexOf(';', i);
      if (semiIdx !== -1 && (braceIdx === -1 || semiIdx < braceIdx)) {
        i = semiIdx + 1;
        continue;
      }
      if (braceIdx !== -1) {
        let depth = 1;
        let j = braceIdx + 1;
        while (j < css.length && depth > 0) {
          if (css[j] === '{') depth++;
          else if (css[j] === '}') depth--;
          j++;
        }
        out += css.slice(braceIdx + 1, j - 1);
        i = j;
        continue;
      }
    }
    out += css[i];
    i++;
  }
  return out;
}

function downlevelCssForAllBrowsers(css: string, filename: string, minify: boolean): string {
  const unwrapped = unwrapCssLayers(css);
  const res = lightningTransform({
    filename,
    code: Buffer.from(unwrapped),
    targets: LEGACY_TARGETS,
    minify,
  });
  return res.code.toString();
}

function crossBrowserCssPlugin(): Plugin {
  return {
    name: 'cross-browser-css-compat',
    enforce: 'post',
    transform(code, id) {
      if (id.split('?')[0].endsWith('.css')) {
        try {
          return {
            code: downlevelCssForAllBrowsers(code, id, false),
            map: null,
          };
        } catch {
          return null;
        }
      }
      return null;
    },
    generateBundle(_options, bundle) {
      for (const [fileName, asset] of Object.entries(bundle)) {
        if (asset.type === 'asset' && fileName.endsWith('.css') && typeof asset.source === 'string') {
          try {
            asset.source = downlevelCssForAllBrowsers(asset.source, fileName, true);
          } catch {
            // Keep original CSS if transform fails
          }
        }
      }
    },
  };
}

export default defineConfig(() => {
  return {
    plugins: [
      react(),
      tailwindcss(),
      crossBrowserCssPlugin(),
      VitePWA({
        registerType: 'autoUpdate',
        includeAssets: [
          'apple-touch-icon.png',
          'pwa-192x192.png',
          'pwa-512x512.png',
          'pwa-maskable-512x512.png',
        ],
        manifest: {
          id: '/',
          name: 'SPMB SD Al-Rasyid Pekanbaru',
          short_name: 'SPMB Rasyid',
          description: 'Portal administrasi Sistem Penerimaan Murid Baru SD Al-Rasyid Pekanbaru.',
          theme_color: '#0f172a',
          background_color: '#f8fafc',
          display: 'standalone',
          start_url: '/',
          scope: '/',
          lang: 'id',
          icons: [
            {
              src: '/pwa-192x192.png',
              sizes: '192x192',
              type: 'image/png',
              purpose: 'any',
            },
            {
              src: '/pwa-512x512.png',
              sizes: '512x512',
              type: 'image/png',
              purpose: 'any',
            },
            {
              src: '/pwa-maskable-512x512.png',
              sizes: '512x512',
              type: 'image/png',
              purpose: 'maskable',
            },
          ],
        },
        workbox: {
          globPatterns: ['**/*.{js,css,html,ico,png,svg,woff,woff2,webmanifest}'],
          globIgnores: ['**/assets/aistudio/**'],
          cleanupOutdatedCaches: true,
          clientsClaim: true,
          skipWaiting: true,
          navigateFallback: '/index.html',
          runtimeCaching: [
            {
              urlPattern: /^https:\/\/iili\.io\/.*/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'school-logo-cache',
                expiration: {
                  maxEntries: 5,
                  maxAgeSeconds: 60 * 60 * 24 * 30,
                },
                cacheableResponse: {
                  statuses: [0, 200],
                },
              },
            },
          ],
        },
        devOptions: {
          enabled: true,
          type: 'module',
        },
      }),
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    build: {
      target: ['es2020', 'chrome88', 'edge88', 'firefox88', 'safari14'],
      cssTarget: ['chrome88', 'edge88', 'firefox88', 'safari14'],
      sourcemap: false,
      rolldownOptions: {
        output: {
          codeSplitting: {
            groups: [
              {
                name: 'vendor-react',
                test: /node_modules[\\/](react|react-dom|scheduler)/,
                priority: 20,
              },
              {
                name: 'vendor-firebase',
                test: /node_modules[\\/](firebase|@firebase)/,
                priority: 15,
              },
              {
                name: 'vendor-icons',
                test: /node_modules[\\/]lucide-react/,
                priority: 10,
              },
            ],
          },
        },
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify—file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
