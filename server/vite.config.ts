//import { defineConfig } from 'vite'
//import react from '@vitejs/plugin-react'

// https://vite.dev/config/
//export default defineConfig({
//  plugins: [react()],
//  server: {
//    port: 4000, // 👈 Set your desired port here
//  },
//})

import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import path from 'path'

// Path to your Office Add-in dev cert
const CERT_DIR = path.join(
  process.env.HOME || process.env.USERPROFILE || '',
  '.office-addin-dev-certs'
)

const httpsOptions = {
  key: fs.readFileSync(path.join(CERT_DIR, 'localhost.key')),
  cert: fs.readFileSync(path.join(CERT_DIR, 'localhost.crt')),
}

export default defineConfig({
  plugins: [react()],
  server: {
    port: 4000,
    https: httpsOptions,
  },
})
