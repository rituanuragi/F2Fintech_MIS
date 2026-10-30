// Run:  npm run test:apps-script
// Checks your Apps Script link + token and lists every tab of the MASTER MIS sheet.
import fs from 'fs';
try { if (fs.existsSync('.env')) process.loadEnvFile('.env'); } catch {}

const url = (process.env.APPS_SCRIPT_URL || '').trim();
const token = (process.env.APPS_SCRIPT_TOKEN || '').trim();
const sheetId = (process.env.MIS_SPREADSHEET_ID || '15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs').trim();
if (!url || !token) { console.error('Put APPS_SCRIPT_URL and APPS_SCRIPT_TOKEN in the .env file first.'); process.exit(1); }

async function call(action, payload = {}) {
  const res = await fetch(url, {
    method: 'POST', redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ token, action, ...payload }),
  });
  const text = await res.text();
  let json; try { json = JSON.parse(text); } catch { throw new Error(`${action}: not JSON (HTTP ${res.status}): ${text.slice(0, 200)}`); }
  if (json && (json.ok === false || json.error)) throw new Error(`${action}: ${json.error || json.message}`);
  return json;
}

try {
  console.log(`1) listTabs for ${sheetId} ...`);
  const tabs = await call('listTabs', { spreadsheetId: sheetId });
  const list = Array.isArray(tabs) ? tabs : (tabs.tabs || tabs.data || []);
  console.log(`   ${list.length} tab(s):`);
  for (const t of list) console.log('   -', t.name ?? t.title, `(${t.rows ?? t.rowCount ?? '?'} rows)`);
  if (list.length) {
    const first = list[0].name ?? list[0].title;
    console.log(`2) readTab "${first}" (5 rows) ...`);
    const r = await call('readTab', { spreadsheetId: sheetId, tab: first, startRow: 1, numRows: 5 });
    const v = r.values || [];
    console.log('   headers:', (v[0] || []).slice(0, 8).join(' | '), '| rows read:', Math.max(0, v.length - 1), '| total rows in tab:', r.totalRows);
  }
  console.log('\nAll good. Apps Script can read all tabs.');
} catch (e) {
  console.error('\nFAILED:', e.message);
  console.error('Check: the web app is deployed with access "Anyone", the token matches the script, and the script has the actions listTabs and readTab.');
  process.exit(1);
}
