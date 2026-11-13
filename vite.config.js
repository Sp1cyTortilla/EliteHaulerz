import { defineConfig } from 'vite'
import { resolve } from 'node:path'

export default defineConfig({
    server: {
        allowedHosts: 'all'
    },
    build: {
        // Top-level await is used in the auth page scripts.
        target: 'es2022',
        rollupOptions: {
            input: {
                main: resolve(__dirname, 'index.html'),
                login: resolve(__dirname, 'login.html'),
                setup: resolve(__dirname, 'setup.html'),
                dashboard: resolve(__dirname, 'dashboard.html'),
            }
        }
    }
})
