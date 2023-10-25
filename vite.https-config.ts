import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { nodePolyfills } from "vite-plugin-node-polyfills";

// https://vitejs.dev/config/
export default defineConfig({
    plugins: [nodePolyfills(), react()],
    resolve: {
        preserveSymlinks: true,
        alias: {
            "node-fetch": "isomorphic-fetch",
        },
    },
    server: {
        hmr: {
            // Needed to make ngrok work with Vite
            clientPort: 443,
        },
        port: 3000,
        open: true,
    },
    optimizeDeps: {
        force: true,
    },
});
