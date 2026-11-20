// Mini-service: Eden FastAPI Backend
// Spawns Python uvicorn as a child process, managed by the system

const { spawn } = require('child_process');
const path = require('path');

const FASTAPI_DIR = '/home/z/my-project/study_repo/eden/fastapi_app';
const PORT = 8001;

// Environment variables
const env = {
  ...process.env,
  PYTHONPATH: FASTAPI_DIR,
  DATABASE_URL: 'sqlite+aiosqlite:///./eden_dev.db',
  REDIS_URL: 'redis://localhost:6379/0',
  EDEN_ENV: 'development',
  EDEN_DEBUG: 'true',
  JWT_SECRET_KEY: 'dev-secret-key-at-least-32-chars-long',
  EDEN_ENCRYPTION_KEY: process.env.EDEN_ENCRYPTION_KEY || 'dGVzdC1rZXktMTIzNDU2Nzg5MDEyMzQ1Njc4OTAxMjM0NTY3ODkwMTIzNDU=',
};

console.log(`[eden-backend] Starting FastAPI on port ${PORT}...`);

const child = spawn('python3', ['-m', 'uvicorn', 'main:app', '--host', '127.0.0.1', '--port', String(PORT)], {
  cwd: FASTAPI_DIR,
  env,
  stdio: ['ignore', 'pipe', 'pipe'],
});

child.stdout.on('data', (data) => {
  console.log(`[eden-backend] ${data.toString().trim()}`);
});

child.stderr.on('data', (data) => {
  console.error(`[eden-backend] ${data.toString().trim()}`);
});

child.on('exit', (code) => {
  console.log(`[eden-backend] Process exited with code ${code}`);
  if (code !== 0) {
    console.log('[eden-backend] Restarting in 3 seconds...');
    setTimeout(() => {
      process.exit(1); // Exit so the system can restart us
    }, 3000);
  }
});

// Keep the process alive
process.on('SIGTERM', () => {
  console.log('[eden-backend] Received SIGTERM, killing FastAPI...');
  child.kill('SIGTERM');
  setTimeout(() => process.exit(0), 1000);
});

console.log(`[eden-backend] FastAPI backend started (PID: ${child.pid})`);
