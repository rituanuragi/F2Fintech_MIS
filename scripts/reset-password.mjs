// Usage (stop the server first!):  npm run reset-password -- admin MyNewPass123
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import initSqlJs from 'sql.js';

const [username, newPassword] = process.argv.slice(2);
if (!username || !newPassword || newPassword.length < 8) {
  console.log('Usage: npm run reset-password -- <username> <new password (8+ characters)>');
  process.exit(1);
}
const file = path.resolve('./data/db/fintech.sqlite');
if (!fs.existsSync(file)) { console.log('Database not found:', file); process.exit(1); }

const SQL = await initSqlJs();
const db = new SQL.Database(fs.readFileSync(file));
const salt = crypto.randomBytes(16).toString('hex');
const hash = `scrypt$${salt}$${crypto.scryptSync(newPassword, salt, 64).toString('hex')}`;
try { db.run('ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0;'); } catch {}
const found = db.exec('SELECT id FROM users WHERE lower(username) = lower(?);', [username]);
if (!found.length || !found[0].values.length) {
  const all = db.exec('SELECT username FROM users;')[0]?.values.map((r) => r[0]).join(', ');
  console.log(`No user named "${username}". Users: ${all}`);
  process.exit(1);
}
db.run("UPDATE users SET password = ?, must_change_password = 1, status = 'active' WHERE lower(username) = lower(?);", [hash, username]);
const tmp = `${file}.reset.tmp`;
fs.writeFileSync(tmp, Buffer.from(db.export()));
fs.renameSync(tmp, file);
console.log(`Done. "${username}" can now log in with the new password and will be asked to choose a new one.`);
