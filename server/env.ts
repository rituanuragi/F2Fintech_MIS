// Loads the .env file in the project folder (APPS_SCRIPT_URL, APPS_SCRIPT_TOKEN, ...).
// Must be the FIRST import of server.ts so the values exist before other files read them.
import * as fs from 'fs';
try {
  if (fs.existsSync('.env') && typeof (process as any).loadEnvFile === 'function') {
    (process as any).loadEnvFile('.env');
  }
} catch (err) {
  console.warn('Could not read .env file:', err);
}
